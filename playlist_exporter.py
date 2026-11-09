"""Playlist Exporter - turn a YouTube playlist (or your own audio files) into MP4 songs,
PNG album covers and an Excel sheet.

Just run it and answer the questions:

    python playlist_exporter.py

Everything ends up in the "Music" folder next to this script:

    Music/
      Batch 001/   up to 20 songs (.mp4), their album covers (.png) and metadata.xlsx
      Batch 002/
      ...
      metadata.xlsx   every song of every batch

Run the same command again after a stop or crash: finished songs are skipped.

The program is split into three files that must stay in the same folder:
    playlist_exporter.py   this file: settings, the menu and the export run
    youtube.py             reading YouTube playlists and downloading songs
    media.py               tags, album covers and converting audio with FFmpeg
"""

from __future__ import annotations

import sys

if sys.version_info < (3, 10):
    sys.exit("Python 3.10 or newer is required. You are running " + sys.version.split()[0])

import argparse
import hashlib
import importlib
import io
import logging
import os
import shutil
import signal
import subprocess
import tempfile
import time
from dataclasses import asdict
from pathlib import Path
from typing import Any, Optional

# =============================================================================
# SETTINGS - the only part you may want to change
# =============================================================================

SONGS_PER_BATCH = 20            # songs per "Batch 001", "Batch 002", ... folder

# YouTube sometimes asks "Sign in to confirm you're not a bot". Put the browser
# you are logged in to YouTube with here ("chrome", "edge", "firefox", "brave")
# or run the script with:  --cookies chrome
COOKIES_FROM_BROWSER = os.environ.get("COOKIES_FROM_BROWSER", "")
DOWNLOAD_PAUSE_SECONDS = 1      # pause between YouTube downloads (avoids bot checks)

# Optional: a YouTube Data API key gives better song info (album, year).
YOUTUBE_API_KEY = os.environ.get("YOUTUBE_API_KEY", "")

# =============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
MUSIC_DIR = SCRIPT_DIR / "Music"
# Resume progress, logs and caches live outside the project so "Music" stays clean.
APP_DIR = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "PlaylistExporter"
LOG_FILE = APP_DIR / "last_run.log"

# Song titles often contain characters the Windows console code page cannot show.
for _stream in (sys.stdout, sys.stderr):
    try:
        if _stream.isatty():
            _stream.reconfigure(errors="replace")
        else:
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

# ----------------------------------------------------------------- dependencies
PACKAGES = {  # import name -> pip name
    "requests": "requests", "mutagen": "mutagen", "openpyxl": "openpyxl",
    "PIL": "Pillow", "yt_dlp": "yt-dlp", "ytmusicapi": "ytmusicapi",
}


def ensure_packages() -> None:
    missing = []
    for module, package in PACKAGES.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(package)
    if not missing:
        return
    print(f"Some Python packages are missing: {', '.join(missing)}")
    answer = "y"
    if sys.stdin.isatty():
        answer = input("Install them now? [Y/n]: ").strip().lower() or "y"
    if answer not in ("y", "yes"):
        sys.exit("Install them with:  python -m pip install " + " ".join(missing))
    if subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", *missing]).returncode != 0:
        sys.exit("Installing failed. Try running this yourself:  python -m pip install " + " ".join(missing))
    importlib.invalidate_caches()


ensure_packages()

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

import media  # noqa: E402
import youtube  # noqa: E402
from media import (DUPLICATE, EXCEL_COLUMNS, EXPORTED, FAILED, METADATA_FIELDS, PENDING,  # noqa: E402
                   UNAVAILABLE, AudioMetadata, FilenameRegistry, PlaylistTrack, TrackResult, normalize,
                   normalize_artist, parse_filename, sanitize_filename, say)

log = logging.getLogger("exporter")


# =============================================================================
# Resume state
# =============================================================================

class StateStore:
    """Progress for one output folder, saved after every song."""

    VERSION = 1

    def __init__(self, path: Path):
        self.path = path
        data = media.read_json(path, {})
        if data.get("version") != self.VERSION:
            data = {}
        self.tracks = {uid: TrackResult.from_dict(t) for uid, t in data.get("tracks", {}).items()}
        self.covers: dict[str, dict[str, Any]] = data.get("covers", {})
        self.audio_names = FilenameRegistry(data.get("audio_names", {}))
        self.cover_names = FilenameRegistry(data.get("cover_names", {}))

    def save(self) -> None:
        media.write_json(self.path, {
            "version": self.VERSION,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "tracks": {uid: asdict(r) for uid, r in self.tracks.items()},
            "audio_names": self.audio_names.to_dict(),
            "covers": self.covers,
            "cover_names": self.cover_names.to_dict(),
        })


# =============================================================================
# Excel
# =============================================================================

HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
STRIPE_FILL = PatternFill("solid", start_color="EEF2F8", end_color="EEF2F8")
NUMBER_FORMATS = {"year": "0", "track": "0", "disc": "0", "duration_seconds": "0.00", "bitrate": "0"}
REPORT_COLUMNS = ["position", "batch", "status", "reason", "title", "artist", "original_filename",
                  "exported_mp4_filename", "cover_filename", "artwork_status", "conversion",
                  "missing_metadata", "metadata_sources", "warnings"]
ALBUM_COLUMNS = ["album", "album_artist", "songs_in_playlist", "album_track_count", "track_numbers_present",
                 "missing_track_numbers", "complete", "batches"]
STATUS_COLORS = {"exported": "D9F2D9", "failed": "F8D7DA", "duplicate": "E2E3E5", "unavailable": "E2E3E5",
                 "pending": "FFF3CD", "yes": "D9F2D9", "no": "FFF3CD", "unknown": "E2E3E5"}


def _add_sheet(workbook: Workbook, title: str, columns: list[str], rows: list[list[Any]],
               color_column: Optional[str] = None) -> None:
    sheet = workbook.active if workbook.active.title == "Sheet" else workbook.create_sheet()
    sheet.title = title
    sheet.append(columns)
    for row in rows:
        sheet.append([None if v == "" else v for v in row])
    for col, name in enumerate(columns, start=1):
        head = sheet.cell(row=1, column=col)
        head.fill, head.font = HEADER_FILL, Font(bold=True, color="FFFFFF")
        head.alignment = Alignment(horizontal="center", vertical="center")
        longest = len(name)
        for r in range(2, len(rows) + 2):
            cell = sheet.cell(row=r, column=col)
            cell.border = Border(bottom=Side(style="thin", color="C9D1DE"))
            if r % 2 == 1:
                cell.fill = STRIPE_FILL
            if name in NUMBER_FORMATS and isinstance(cell.value, (int, float)):
                cell.number_format = NUMBER_FORMATS[name]
            if name == color_column and str(cell.value) in STATUS_COLORS:
                color = STATUS_COLORS[str(cell.value)]
                cell.fill = PatternFill("solid", start_color=color, end_color=color)
            if cell.value is not None:
                longest = max(longest, len(f"{cell.value:.2f}" if isinstance(cell.value, float) else str(cell.value)))
        sheet.column_dimensions[get_column_letter(col)].width = min(max(longest + 3, 8), 60)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(rows) + 1}"


def album_rows(results: list[TrackResult]) -> list[list[Any]]:
    """One row per album: which track numbers you have and which are missing."""
    groups: dict[tuple[str, str], list[TrackResult]] = {}
    for r in results:
        if r.status == EXPORTED and r.row.get("album"):
            key = (str(r.row["album"]).casefold().strip(),
                   str(r.row.get("album_artist") or r.row.get("artist") or "").casefold().strip())
            groups.setdefault(key, []).append(r)
    rows = []
    for members in groups.values():
        first = members[0].row
        numbers = sorted({int(m.row["track"]) for m in members if m.row.get("track")})
        total = max((m.album_track_count for m in members if m.album_track_count), default=None)
        missing = [n for n in range(1, total + 1) if n not in numbers] if total else []
        complete = ("no" if missing else "yes") if total and len(numbers) == len(members) else "unknown"
        rows.append([first.get("album"), first.get("album_artist") or first.get("artist"), len(members), total,
                     ", ".join(map(str, numbers)) or None, ", ".join(map(str, missing)) or None, complete,
                     ", ".join(sorted({m.batch for m in members if m.batch})) or None])
    rows.sort(key=lambda r: (r[6] != "yes", str(r[1]).casefold(), str(r[0]).casefold()))
    return rows


def write_workbook(path: Path, results: list[TrackResult], with_report: bool) -> Path:
    """Write metadata.xlsx. If it is open in Excel, a timestamped copy is written instead."""
    workbook = Workbook()
    exported = [r for r in results if r.status == EXPORTED]
    _add_sheet(workbook, "Metadata", EXCEL_COLUMNS, [[r.row.get(c) for c in EXCEL_COLUMNS] for r in exported])
    if with_report:
        report = [[r.position, r.batch, r.status, r.reason, r.row.get("title") or r.display_name,
                   r.row.get("artist", ""), r.original_filename, r.exported_filename, r.cover_filename,
                   r.cover_status, r.conversion, ", ".join(r.missing_fields),
                   ", ".join(f"{k}={v}" for k, v in r.metadata_sources.items() if v), " | ".join(r.warnings)]
                  for r in results]
        _add_sheet(workbook, "Processing Report", REPORT_COLUMNS, report, color_column="status")
        _add_sheet(workbook, "Albums", ALBUM_COLUMNS, album_rows(results), color_column="complete")
    buffer = io.BytesIO()
    workbook.save(buffer)
    try:
        media.atomic_write(path, buffer.getvalue())
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_{time.strftime('%Y%m%d_%H%M%S')}{path.suffix}")
        say(f"  [WARN] {path.name} is open in another program; saved as {fallback.name} instead.")
        media.atomic_write(fallback, buffer.getvalue())
        return fallback


# =============================================================================
# The export run
# =============================================================================

def find_duplicates(tracks: list[PlaylistTrack]) -> dict[str, int]:
    """uid -> position of the first occurrence, for songs that are in the playlist twice."""
    first_seen: dict[str, int] = {}
    duplicates: dict[str, int] = {}
    for track in tracks:
        if track.unavailable_reason:
            continue
        keys = [f"meta:{normalize_artist(track.artist)}|{normalize(track.title)}|{normalize(track.album)}"]
        if track.source_id:
            keys.append(f"id:{track.source_id}")
        if track.local_path:
            keys = [f"path:{track.local_path.casefold()}"]
        original = next((first_seen[k] for k in keys if k in first_seen), None)
        if original is not None:
            duplicates[track.uid] = original
        else:
            for key in keys:
                first_seen[key] = track.position
    return duplicates


def folder_tracks(library: list[AudioMetadata], folder: Path) -> list[PlaylistTrack]:
    """'My own files' mode: every audio file in the folder is one song."""
    tracks = []
    for position, audio in enumerate(library, start=1):
        path = Path(audio.path)
        artist, title = parse_filename(path.stem)
        try:
            relative = path.relative_to(folder).as_posix()
        except ValueError:
            relative = path.as_posix()
        artist = audio.artist or artist
        tracks.append(PlaylistTrack(position, f"local:{relative.casefold()}", "local",
                                    title=audio.title or title or path.stem, artists=[artist] if artist else [],
                                    album=audio.album, duration_seconds=audio.duration_seconds,
                                    local_path=str(path)))
    return tracks


class Exporter:
    def __init__(self, output_dir: Path):
        self.output_dir = output_dir.resolve()
        self.http = media.HttpClient()
        self.download_dir: Optional[Path] = None
        self.stop_requested = False
        self._albums: Optional[youtube.YouTubeMusicAlbums] = None
        self._albums_failed = False
        folder_id = hashlib.sha1(str(self.output_dir).casefold().encode("utf-8")).hexdigest()[:12]
        self.state = StateStore(APP_DIR / "state" / f"{sanitize_filename(self.output_dir.name, 'output', 40)}-{folder_id}.json")
        # Covers are kept once here and copied into each batch folder that needs them.
        self.covers_dir = APP_DIR / "covers" / folder_id
        self.work_dir = APP_DIR / "work"

    def dir_for(self, result: TrackResult) -> Path:
        return self.output_dir / result.batch if result.batch else self.output_dir

    def _next_batch(self, uid: str) -> str:
        counts: dict[str, int] = {}
        for other, r in self.state.tracks.items():
            if other != uid and r.status == EXPORTED and r.batch:
                counts[r.batch] = counts.get(r.batch, 0) + 1
        number = 1
        while counts.get(f"Batch {number:03d}", 0) >= SONGS_PER_BATCH:
            number += 1
        return f"Batch {number:03d}"

    def _already_done(self, previous: Optional[TrackResult]) -> bool:
        if not previous or previous.status != EXPORTED or not previous.exported_filename:
            return False
        out = self.dir_for(previous) / previous.exported_filename
        cover_ok = not previous.cover_filename or (self.dir_for(previous) / previous.cover_filename).is_file()
        return out.is_file() and out.stat().st_size == previous.exported_size and cover_ok

    def _skip(self, previous: TrackResult) -> TrackResult:
        say(f"    [SKIP] Already exported: {previous.batch}\\{previous.exported_filename}")
        previous.conversion = previous.conversion.replace(" (resumed)", "") + " (resumed)"
        return previous

    def _forget_deleted_files(self) -> None:
        """Songs whose MP4 was deleted from Music are exported again (and free their batch slot)."""
        for uid, r in list(self.state.tracks.items()):
            if r.status == EXPORTED and not (self.dir_for(r) / r.exported_filename).is_file():
                del self.state.tracks[uid]
                self.state.audio_names.release(uid)

    # ------------------------------------------------------------------ run
    def run(self, url: str, audio_dir: Optional[Path], fresh: bool, limit: int) -> dict[str, Any]:
        library: dict[str, AudioMetadata] = {}
        if audio_dir:
            say(f"\n[Step 1/4] Reading your audio files in {audio_dir} ...")
            files = media.scan_folder(audio_dir)
            library = {a.path: a for a in files}
            name, tracks = audio_dir.name, folder_tracks(files, audio_dir)
        else:
            say("\n[Step 1/4] Reading the YouTube playlist ...")
            name, tracks = youtube.get_playlist(self.http, url, YOUTUBE_API_KEY, COOKIES_FROM_BROWSER)
        if limit and len(tracks) > limit:
            say(f"  [INFO] --limit {limit}: only the first {limit} of {len(tracks)} songs are processed.")
            tracks = tracks[:limit]
        if not tracks:
            raise RuntimeError("The playlist is empty (or the folder contains no audio files).")
        say(f"  {name}: {len(tracks)} songs")

        for directory in (self.output_dir, self.work_dir, self.covers_dir):
            directory.mkdir(parents=True, exist_ok=True)
        media.cleanup_partial_files(self.work_dir)
        if fresh:
            self.state.tracks.clear()
            self.state.covers.clear()
            say("  [INFO] --fresh: everything is processed again (existing files are replaced).")
        self._forget_deleted_files()
        done_before = sum(1 for r in self.state.tracks.values() if r.status == EXPORTED)
        if done_before:
            say(f"  [INFO] {done_before} song(s) are already in {self.output_dir.name} and will be skipped.")
        duplicates = find_duplicates(tracks)
        if duplicates:
            say(f"  {len(duplicates)} song(s) are in the playlist twice (exported once).")

        download_tmp = None
        if not audio_dir:
            download_tmp = tempfile.TemporaryDirectory(prefix="yt_", dir=APP_DIR, ignore_cleanup_errors=True)
            self.download_dir = Path(download_tmp.name)

        say("\n[Step 2/4] Exporting songs ...")
        say("  (Ctrl+C once = stop after this song, twice = stop now. Run again later to continue.)")
        artwork = media.ArtworkManager(self.http, self.covers_dir, APP_DIR / "artwork_cache", self.state)
        results: list[TrackResult] = []
        interrupted = False
        self.stop_requested = False

        def on_ctrl_c(signum, frame):
            if self.stop_requested:
                raise KeyboardInterrupt
            self.stop_requested = True
            say("\n  [STOP] Finishing this song, then stopping. Press Ctrl+C again to stop now.")

        previous_handler = signal.signal(signal.SIGINT, on_ctrl_c)
        try:
            for index, track in enumerate(tracks, start=1):
                if self.stop_requested:
                    raise KeyboardInterrupt
                say(f"[{index}/{len(tracks)}] {track.display_name}")
                try:
                    result = self._process(track, library.get(track.local_path), duplicates, artwork)
                except KeyboardInterrupt:
                    raise
                except Exception as exc:  # one bad song never stops the run
                    log.exception("Unexpected error processing %s", track.display_name)
                    result = TrackResult(track.uid, track.position, track.display_name, FAILED,
                                         f"Unexpected error: {type(exc).__name__}: {exc}")
                    say(f"    [FAIL] {result.reason}")
                result.position = track.position
                results.append(result)
                self.state.tracks[track.uid] = result
                self.state.save()
        except KeyboardInterrupt:
            interrupted = True
            say("\n[STOPPED] Progress saved. Run the same command again to continue.")
            self.state.save()
            done = {r.uid for r in results}
            results += [TrackResult(t.uid, t.position, t.display_name, PENDING, "Not processed yet (run was stopped)")
                        for t in tracks if t.uid not in done]
        finally:
            signal.signal(signal.SIGINT, previous_handler)
            if download_tmp is not None:
                download_tmp.cleanup()
            self.download_dir = None

        say("\n[Step 3/4] Writing the Excel files ...")
        # Music/metadata.xlsx lists every song ever exported into this folder (all playlists/runs).
        current = {r.uid for r in results}
        everything = results + [r for uid, r in self.state.tracks.items() if uid not in current and r.status == EXPORTED]
        workbook = write_workbook(self.output_dir / "metadata.xlsx", everything, with_report=True)
        batches: dict[str, list[TrackResult]] = {}
        for r in everything:
            if r.status == EXPORTED and r.batch:
                batches.setdefault(r.batch, []).append(r)
        for batch, members in sorted(batches.items()):
            write_workbook(self.output_dir / batch / "metadata.xlsx", members, with_report=False)
        say(f"  [OK] {workbook} and one metadata.xlsx in each batch folder")
        return {"name": name, "results": results, "workbook": workbook, "interrupted": interrupted}

    # ------------------------------------------------------------ per song
    def _process(self, track: PlaylistTrack, local: Optional[AudioMetadata], duplicates: dict[str, int],
                 artwork: media.ArtworkManager) -> TrackResult:
        result = TrackResult(track.uid, track.position, track.display_name)
        if track.unavailable_reason:
            result.status, result.reason = UNAVAILABLE, track.unavailable_reason
            say(f"    [SKIP] {track.unavailable_reason}")
            return result
        if track.uid in duplicates:
            result.status = DUPLICATE
            result.reason = f"Same song as #{duplicates[track.uid]} (exported once)"
            say(f"    [SKIP] {result.reason}")
            return result
        if local is not None:
            return self._export(track, local, "local", result, artwork)

        previous = self.state.tracks.get(track.uid)
        if self._already_done(previous):
            return self._skip(previous)
        say("    [..] Downloading from YouTube ...")
        downloaded: Optional[Path] = None
        music: dict[str, Any] = {}
        while downloaded is None:
            try:
                downloaded, music = youtube.download(track.source_id, self.download_dir,
                                                     COOKIES_FROM_BROWSER, DOWNLOAD_PAUSE_SECONDS)
            except youtube.NetworkDownError as exc:
                log.warning("Network problem while downloading %s: %s", track.display_name, exc)
                if not self._wait_for_internet():
                    self.stop_requested = True  # not the song's fault: leave it for the next run
                    result.status, result.reason = PENDING, "No internet connection (run again to continue)"
                    say(f"    [STOP] {result.reason}")
                    return result
                say("    [..] Connection is back - retrying ...")
            except youtube.DownloadError as exc:
                result.status, result.reason = FAILED, f"Download failed: {exc}"
                say(f"    [FAIL] {result.reason}")
                return result
        # YouTube Music's own song info beats a title parsed from the video name.
        if music.get("title") and music.get("artists") and not track.album:
            track.title, track.artists = music["title"], music["artists"]
        if music.get("album") and not track.album:
            track.album = music["album"]
        track.year = track.year or music.get("year")
        track.duration_seconds = track.duration_seconds or music.get("duration")
        if music.get("album") and track.source_id and not track.artwork_url:
            track.artwork_url = f"https://i.ytimg.com/vi/{track.source_id}/maxresdefault.jpg"
        try:
            return self._export(track, media.read_audio_metadata(downloaded), "download", result, artwork)
        finally:
            downloaded.unlink(missing_ok=True)

    def _wait_for_internet(self, max_wait: int = 600, interval: int = 15) -> bool:
        say(f"    [WAIT] No internet - waiting up to {max_wait // 60} min for it to come back (Ctrl+C to stop) ...")
        for _ in range(max_wait // interval):
            if self.stop_requested:
                return False
            if youtube.internet_available():
                return True
            time.sleep(interval)
        return youtube.internet_available()

    def _export(self, track: PlaylistTrack, audio: AudioMetadata, method: str, result: TrackResult,
                artwork: media.ArtworkManager) -> TrackResult:
        source_path = Path(audio.path)
        result.original_path, result.original_filename = str(source_path), source_path.name
        result.source_fingerprint, result.match_method = audio.fingerprint, method
        if audio.error:
            result.status, result.reason = FAILED, audio.error
            say(f"    [FAIL] {audio.error}")
            return result

        previous = self.state.tracks.get(track.uid)
        if (method == "local" and previous and previous.original_path == str(source_path)
                and previous.source_fingerprint == audio.fingerprint and self._already_done(previous)):
            return self._skip(previous)

        values, sources = media.merge_metadata(audio, track)
        self._add_track_number(track, audio, values, sources, result)
        result.metadata_sources = sources
        result.missing_fields = [f for f in METADATA_FIELDS if values.get(f) in (None, "")]
        if result.missing_fields:
            say(f"    [INFO] Missing info: {', '.join(result.missing_fields)}")

        try:
            info = media.probe(source_path)
        except media.ConversionError as exc:
            result.status, result.reason = FAILED, f"Corrupted or unreadable audio: {exc}"
            say(f"    [FAIL] {result.reason}")
            return result
        source_duration = media.probe_duration(info) or audio.duration_seconds

        art = artwork.get(values, source_path, track, crop_embedded=method == "download")
        result.cover_filename, result.cover_status = art.cover_filename, art.status
        if not art.cover_filename:
            say(f"    [INFO] No album cover found ({art.status.replace('missing: ', '')})")

        result.batch = previous.batch if previous and previous.batch else self._next_batch(track.uid)
        target_dir = self.dir_for(result)
        target_dir.mkdir(parents=True, exist_ok=True)
        stem = sanitize_filename(values.get("title") or source_path.stem, sanitize_filename(source_path.stem))
        filename = self.state.audio_names.claim(track.uid, stem, media.OUTPUT_EXTENSION, target_dir)
        destination = target_dir / filename

        try:
            tmp, description = media.convert(source_path, info, self.work_dir)
        except media.ConversionError as exc:
            self.state.audio_names.release(track.uid)
            result.status, result.reason = FAILED, f"Conversion failed: {exc}"
            say(f"    [FAIL] {result.reason}")
            return result
        try:
            try:
                media.write_tags(tmp, values, art.embed)
            except Exception as exc:
                result.warnings.append(f"Could not write tags/artwork into the file: {exc}")
                say(f"    [WARN] Tags could not be written: {exc}")
            problems, out_info = media.check_export(tmp, source_duration)
            if problems:
                self.state.audio_names.release(track.uid)
                result.status, result.reason = FAILED, f"Exported file is broken: {'; '.join(problems)}"
                say(f"    [FAIL] {result.reason}")
                return result
            media.replace_file(tmp, destination)
        finally:
            tmp.unlink(missing_ok=True)

        if art.cover_filename and not (target_dir / art.cover_filename).exists():
            try:
                shutil.copy2(self.covers_dir / art.cover_filename, target_dir / art.cover_filename)
            except OSError as exc:
                result.warnings.append(f"Could not copy cover into {result.batch}: {exc}")
                result.cover_filename = ""

        result.status = EXPORTED
        result.exported_filename = filename
        result.exported_size = destination.stat().st_size
        result.conversion = description
        result.row = {
            "original_filename": source_path.name, "exported_mp4_filename": filename,
            **{k: values[k] for k in METADATA_FIELDS},
            "duration_seconds": round(media.probe_duration(out_info) or source_duration or 0, 2) or None,
            "bitrate": media.probe_bitrate(out_info), "cover_filename": result.cover_filename,
        }
        say(f"    [OK] {result.batch}\\{filename}" + (f"  (cover: {result.cover_filename})" if result.cover_filename else ""))
        return result

    def _add_track_number(self, track: PlaylistTrack, audio: AudioMetadata, values: dict[str, Any],
                          sources: dict[str, str], result: TrackResult) -> None:
        """YouTube has no track numbers; look the song up on its YouTube Music album page."""
        if track.source != "youtube" or values.get("track") or self._albums_failed:
            return
        if self._albums is None:
            try:
                self._albums = youtube.YouTubeMusicAlbums()
            except Exception as exc:
                log.warning("YouTube Music is unavailable: %s", exc)
                self._albums_failed = True
                return
        if sources.get("artist") == "file" and values.get("artist"):
            artists = [str(values["artist"]).split(", ")[0]]
        else:
            artists = track.artists or ([values["artist"]] if values.get("artist") else [])
        info = self._albums.find(track.source_id, values.get("title") or track.title, artists,
                                 values.get("album") or "", audio.duration_seconds)
        if info is None:
            return
        if not values.get("album"):
            values["album"], sources["album"] = info.album, "youtube-music"
        if not youtube.same_album(info.album, values["album"]):
            return
        values["track"], sources["track"] = info.track, "youtube-music"
        values["track_total"] = result.album_track_count = info.track_count
        if not values.get("album_artist") and info.album_artist:
            values["album_artist"], sources["album_artist"] = info.album_artist, "youtube-music"
        if not values.get("year") and info.year:
            values["year"], sources["year"] = info.year, "youtube-music"


def print_summary(summary: dict[str, Any], output_dir: Path) -> None:
    results: list[TrackResult] = summary["results"]
    exported = [r for r in results if r.status == EXPORTED]
    count = lambda status: sum(1 for r in results if r.status == status)  # noqa: E731
    resumed = sum(1 for r in exported if "resumed" in r.conversion)
    size = 0
    for r in exported:
        try:
            size += (output_dir / r.batch / r.exported_filename).stat().st_size
        except OSError:
            pass
    line = "=" * 64
    say(f"\n[Step 4/4] Summary\n{line}")
    say(f"  Playlist:              {summary['name']}")
    say(f"  Songs:                 {len(results)}")
    say(f"  Exported:              {len(exported)}" + (f"  ({resumed} were already done)" if resumed else ""))
    say(f"  Failed:                {count(FAILED)}")
    if count(DUPLICATE) or count(UNAVAILABLE):
        say(f"  Skipped:               {count(DUPLICATE)} duplicate, {count(UNAVAILABLE)} unavailable")
    say(f"  Without album cover:   {sum(1 for r in exported if not r.cover_filename)}")
    say(f"  Size:                  {media.human_size(size)}")
    say(f"  Output folder:         {output_dir}")
    say(f"  Excel:                 {summary['workbook']}")
    say(f"  Log file:              {LOG_FILE}")
    say(line)
    failed = [r for r in results if r.status == FAILED]
    if failed:
        say("  Songs that were not exported:")
        for r in failed[:25]:
            say(f"   #{r.position:<4} {r.display_name}: {r.reason}")
        if len(failed) > 25:
            say(f"   ... and {len(failed) - 25} more - see the 'Processing Report' sheet in metadata.xlsx.")
    if summary["interrupted"] or any(r.status == PENDING for r in results):
        say("\n  Not finished. Run the same command again to continue where it stopped.")


# =============================================================================
# Menu / command line
# =============================================================================

def ask(prompt: str) -> str:
    try:
        return input(f"{prompt}: ").strip().strip('"').strip("'")
    except EOFError:
        return ""


def ask_questions(args: argparse.Namespace) -> None:
    say("\n=== Playlist Exporter ===\n")
    if not args.url and not args.audio_dir:
        say("What do you want to do?")
        say("  1) Download a YouTube / YouTube Music playlist")
        say("  2) Convert audio files I already have (MP3, FLAC, M4A, WAV, ...)")
        choice = ""
        while choice not in ("1", "2"):
            choice = ask("Enter 1 or 2") or "1"
        if choice == "1":
            while not args.url:
                args.url = ask("Paste the playlist link")
        else:
            while not args.audio_dir:
                candidate = ask("Folder that contains your audio files")
                if candidate and Path(candidate).expanduser().is_dir():
                    args.audio_dir = candidate
                elif candidate:
                    say(f"  Folder not found: {candidate}")


def ensure_ffmpeg(interactive: bool) -> bool:
    if media.find_ffmpeg():
        return True
    say("[!] FFmpeg is not installed (it is needed to convert the audio).")
    if os.name == "nt" and shutil.which("winget") and interactive:
        if (input("    Install it now with winget? [Y/n]: ").strip().lower() or "y") in ("y", "yes"):
            subprocess.run(["winget", "install", "--id", "Gyan.FFmpeg", "-e",
                            "--accept-source-agreements", "--accept-package-agreements"])
            if media.find_ffmpeg():
                return True
    say("    Install it (Windows: open a terminal and run  winget install Gyan.FFmpeg ),")
    say("    then close and reopen the terminal and run this script again.")
    return False


def setup_logging() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    root.handlers[:] = [handler]
    for noisy in ("urllib3", "PIL"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def main() -> int:
    global COOKIES_FROM_BROWSER
    parser = argparse.ArgumentParser(description="Turn a YouTube playlist, or a folder of your own audio files, "
                                                 "into MP4 songs, PNG album covers and an Excel sheet, "
                                                 "in batch folders inside 'Music'.")
    parser.add_argument("url", nargs="?", help="YouTube / YouTube Music playlist link")
    parser.add_argument("--audio-dir", help="convert this folder of your own audio files instead")
    parser.add_argument("--cookies", metavar="BROWSER", help="use this browser's YouTube login (chrome, edge, firefox)")
    parser.add_argument("--limit", type=int, default=0, help="only the first N songs (for testing)")
    parser.add_argument("--fresh", action="store_true", help="ignore saved progress and redo everything")
    parser.add_argument("--output", help=argparse.SUPPRESS)  # for testing: another folder instead of Music
    args = parser.parse_args()

    if args.cookies:
        COOKIES_FROM_BROWSER = args.cookies.lower()
    if args.url and args.audio_dir:
        say("[ERROR] Give either a playlist link or --audio-dir, not both.")
        return 2
    interactive = sys.stdin.isatty()
    if not ensure_ffmpeg(interactive):
        return 1
    if interactive:
        ask_questions(args)
    if not args.url and not args.audio_dir:
        say("[ERROR] Give a playlist link (or --audio-dir), or run without arguments to be asked.")
        return 2
    audio_dir = Path(args.audio_dir).expanduser().resolve() if args.audio_dir else None
    if audio_dir and not audio_dir.is_dir():
        say(f"[ERROR] Folder not found: {audio_dir}")
        return 2

    setup_logging()
    output_dir = Path(args.output).expanduser() if args.output else MUSIC_DIR
    log.info("Run: url=%s audio_dir=%s output=%s", args.url, audio_dir, output_dir)
    exporter = Exporter(output_dir)
    try:
        summary = exporter.run(args.url or "", audio_dir, args.fresh, args.limit)
    except (youtube.PlaylistError, media.NetworkError) as exc:
        say(f"\n[ERROR] {exc}")
        return 1
    except KeyboardInterrupt:
        say("\nStopped. Run the same command again to continue.")
        return 130
    except PermissionError as exc:
        say(f"\n[ERROR] Permission denied: {exc}\n  Close any program that has the files open and try again.")
        return 1
    except Exception as exc:
        log.exception("Fatal error")
        say(f"\n[ERROR] {type(exc).__name__}: {exc}\n  Details are in {LOG_FILE}")
        return 1
    print_summary(summary, exporter.output_dir)
    return 130 if summary["interrupted"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
