"""Helpers for playlist_exporter.py: song data, filenames, tags, album covers and FFmpeg.

You do not run this file yourself; playlist_exporter.py uses it.
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import logging
import os
import random
import re
import shutil
import struct
import subprocess
import tempfile
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

import mutagen
import requests
from mutagen.flac import Picture
from mutagen.id3 import ID3
from mutagen.mp4 import MP4, MP4Cover
from PIL import Image, UnidentifiedImageError

log = logging.getLogger("exporter")

OUTPUT_EXTENSION = ".mp4"
LOSSLESS_AAC_BITRATE = 256
FFMPEG_TIMEOUT = 1800
HTTP_TIMEOUT = 30
HTTP_RETRIES = 5
MAX_RATE_LIMIT_WAIT = 600
MAX_FILENAME_LENGTH = 150
USER_AGENT = "PlaylistExporter/1.0 (personal music library tool)"
AUDIO_EXTENSIONS = {".mp3", ".m4a", ".mp4", ".aac", ".flac", ".wav", ".aiff", ".aif",
                    ".ogg", ".oga", ".opus", ".wma", ".alac", ".ape", ".wv"}

FFMPEG = "ffmpeg"    # replaced with the full path by find_ffmpeg()
FFPROBE = "ffprobe"


def say(message: str = "", end: str = "\n") -> None:
    print(message, end=end, flush=True)


# =============================================================================
# Data
# =============================================================================

EXCEL_COLUMNS = [
    "original_filename", "exported_mp4_filename", "title", "artist", "album", "album_artist",
    "year", "track", "disc", "genre", "duration_seconds", "bitrate", "cover_filename",
]
METADATA_FIELDS = ["title", "artist", "album", "album_artist", "year", "track", "disc", "genre"]

EXPORTED, FAILED, DUPLICATE, UNAVAILABLE, PENDING = "exported", "failed", "duplicate", "unavailable", "pending"


@dataclass
class PlaylistTrack:
    """One song to export: a YouTube video or one of your own audio files."""
    position: int
    uid: str
    source: str  # youtube | local
    title: str = ""
    artists: list[str] = field(default_factory=list)
    album: str = ""
    year: Optional[int] = None
    duration_seconds: Optional[float] = None
    artwork_url: str = ""
    source_id: str = ""  # YouTube video id
    unavailable_reason: str = ""
    local_path: str = ""

    @property
    def artist(self) -> str:
        return ", ".join(a for a in self.artists if a)

    @property
    def display_name(self) -> str:
        if self.artist and self.title:
            return f"{self.artist} - {self.title}"
        return self.title or self.local_path or self.uid


@dataclass
class AudioMetadata:
    """Tags and technical info read from an audio file."""
    path: str
    size: int
    mtime: float
    title: str = ""
    artist: str = ""
    album: str = ""
    album_artist: str = ""
    year: Optional[int] = None
    track: Optional[int] = None
    disc: Optional[int] = None
    genre: str = ""
    duration_seconds: Optional[float] = None
    error: str = ""

    @property
    def fingerprint(self) -> str:
        return f"{self.size}:{int(self.mtime)}"


@dataclass
class TrackResult:
    """The outcome for one song. Saved in the resume file."""
    uid: str
    position: int
    display_name: str
    status: str = PENDING
    reason: str = ""
    original_path: str = ""
    original_filename: str = ""
    source_fingerprint: str = ""
    exported_filename: str = ""
    exported_size: int = 0
    cover_filename: str = ""
    cover_status: str = ""
    conversion: str = ""
    match_method: str = ""  # download | local
    row: dict[str, Any] = field(default_factory=dict)
    metadata_sources: dict[str, str] = field(default_factory=dict)
    missing_fields: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    batch: str = ""
    album_track_count: Optional[int] = None

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TrackResult":
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})


# =============================================================================
# Text helpers
# =============================================================================

_NOISE_WORDS = (
    r"official\s*(music\s*)?(video|audio|lyric\s*video|visuali[sz]er)?|lyrics?(\s*video)?|"
    r"(video|audio)\s*oficial|video\s*lyrics?|lyric\s*video|letra|performance\s*video|"
    r"audio|video|hd|hq|4k|mv|visuali[sz]er|explicit|clean|"
    r"(\d{4}\s*)?(digital(ly)?\s*)?remaster(ed)?(\s*\d{4})?(\s*version)?|"
    r"(\d{4}\s*)?mono|stereo(\s*(mix|version))?|album\s*version|original\s*mix"
)
_BRACKET_NOISE = re.compile(r"[\(\[\{]\s*(" + _NOISE_WORDS + r")\s*[\)\]\}]", re.IGNORECASE)
_DASH_NOISE = re.compile(r"\s+[-–—]\s+(" + _NOISE_WORDS + r")\s*$", re.IGNORECASE)
_FEATURING = re.compile(
    r"[\(\[]\s*(feat\.?|ft\.?|featuring|with)\s+[^\)\]]*[\)\]]|\s+(feat\.?|ft\.?|featuring)\s+.*$",
    re.IGNORECASE,
)
_NON_WORD = re.compile(r"[^\w\s]", re.UNICODE)
_SPACES = re.compile(r"\s+")
_LEADING_TRACK_NO = re.compile(r"^\s*(?:cd\s*\d+\s*[-_.]?\s*)?\d{1,3}\s*[-_.)]\s*|^\s*\d{1,3}\s+(?=\D)", re.I)


def clean_title(title: str) -> str:
    """Remove release noise like '(Official Video)' or '- 2011 Remaster'."""
    value = title or ""
    for _ in range(3):
        new = _DASH_NOISE.sub("", _BRACKET_NOISE.sub(" ", value))
        if new == value:
            break
        value = new
    return _SPACES.sub(" ", value).strip(" -–—")


def normalize(value: str) -> str:
    """Lower-case, accent-free, punctuation-free form used for comparisons."""
    if not value:
        return ""
    text = _FEATURING.sub(" ", clean_title(value))
    text = "".join(c for c in unicodedata.normalize("NFKD", text.casefold()) if not unicodedata.combining(c))
    text = text.replace("&", " and ").replace("+", " and ")
    text = _NON_WORD.sub(" ", text).replace("_", " ")
    return _SPACES.sub(" ", text).strip()


def normalize_artist(value: str) -> str:
    text = normalize(value)
    return text[4:] if text.startswith("the ") else text


def parse_int(value) -> Optional[int]:
    """'3', '3/12', ' 03 ', 3.0 -> 3. None when not a positive number."""
    if isinstance(value, (tuple, list)):
        value = value[0] if value else None
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if int(value) > 0 else None
    match = re.match(r"\s*(\d+)", str(value))
    return int(match.group(1)) if match and int(match.group(1)) > 0 else None


def parse_year(value) -> Optional[int]:
    if value is None:
        return None
    match = re.search(r"(1[89]\d{2}|20\d{2}|21\d{2})", str(value))
    return int(match.group(1)) if match else None


def split_artist_title(text: str) -> tuple[str, str]:
    """'Artist - Title' -> ('Artist', 'Title'). ('', text) without a separator."""
    for separator in (" - ", " – ", " — ", " -- ", " | "):
        if separator in text:
            left, right = text.split(separator, 1)
            if left.strip() and right.strip():
                return left.strip(), right.strip()
    return "", text.strip()


def parse_filename(stem: str) -> tuple[str, str]:
    """Best-effort (artist, title) from a filename like '03 - Metallica - Battery'."""
    text = _LEADING_TRACK_NO.sub("", stem.replace("_", " "), count=1)
    artist, title = split_artist_title(text)
    return artist, clean_title(title)


# =============================================================================
# Files
# =============================================================================

_INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                   *(f"LPT{i}" for i in range(1, 10))}


def sanitize_filename(name: str, fallback: str = "Unknown", max_length: int = MAX_FILENAME_LENGTH) -> str:
    """A Windows-safe filename stem (no extension)."""
    text = _INVALID_CHARS.sub("_", unicodedata.normalize("NFC", str(name or "")))
    text = re.sub(r"\s+", " ", text).strip().rstrip(". ")
    if not text or set(text) <= {"_", " ", "."}:
        text = fallback
    if text.split(".")[0].upper() in _RESERVED_NAMES:
        text = "_" + text
    return text[:max_length].rstrip(". ") if len(text) > max_length else text


class FilenameRegistry:
    """Unique filenames per owner (song / album), stable across resumed runs.

    Compared case-insensitively because Windows is.
    """

    def __init__(self, existing: Optional[dict[str, str]] = None):
        self._by_name: dict[str, str] = {}
        self._by_owner: dict[str, str] = {}
        for name, owner in (existing or {}).items():
            self._by_name[name.casefold()] = owner
            self._by_owner[owner] = name

    def get(self, owner: str) -> Optional[str]:
        return self._by_owner.get(owner)

    def claim(self, owner: str, stem: str, extension: str, folder: Optional[Path] = None) -> str:
        """A unique name for ``owner``; files already on disk in ``folder`` are never overwritten."""
        current = self._by_owner.get(owner)
        if current and re.fullmatch(re.escape(stem) + r"( \(\d+\))?" + re.escape(extension), current,
                                    flags=re.IGNORECASE):
            return current
        if current:
            self.release(owner)

        def taken(name: str) -> bool:
            holder = self._by_name.get(name.casefold())
            if holder is not None:
                return holder != owner
            return folder is not None and (folder / name).exists()

        candidate, counter = f"{stem}{extension}", 2
        while taken(candidate):
            candidate = f"{stem} ({counter}){extension}"
            counter += 1
        self._by_name[candidate.casefold()] = owner
        self._by_owner[owner] = candidate
        return candidate

    def release(self, owner: str) -> None:
        name = self._by_owner.pop(owner, None)
        if name:
            self._by_name.pop(name.casefold(), None)

    def to_dict(self) -> dict[str, str]:
        return {name: owner for owner, name in self._by_owner.items()}


def replace_file(src: Path, dst: Path) -> None:
    """os.replace with retries: antivirus / indexers on Windows hold files for a moment."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(5):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == 4:
                raise
            time.sleep(0.3 * (attempt + 1))


def atomic_write(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix=".tmp_", suffix=path.suffix)
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
        replace_file(Path(tmp_name), path)
    except BaseException:
        Path(tmp_name).unlink(missing_ok=True)
        raise


def write_json(path: Path, data: Any) -> None:
    atomic_write(path, json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"))


def read_json(path: Path, default: Any) -> Any:
    if not path.is_file():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        try:  # a corrupt file must never block a run
            path.replace(path.with_suffix(path.suffix + f".corrupt-{int(time.time())}"))
        except OSError:
            pass
        return default


def cleanup_partial_files(folder: Path) -> None:
    """Remove leftovers (.part / .tmp_) of an interrupted run."""
    if folder.is_dir():
        for path in folder.iterdir():
            if path.is_file() and (path.name.startswith(".tmp_") or ".part" in path.suffixes):
                path.unlink(missing_ok=True)


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{int(num)} B" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} GB"


# =============================================================================
# HTTP with retries
# =============================================================================

class NetworkError(Exception):
    pass


class HttpClient:
    """requests.Session that retries network errors, 5xx and 429 (rate limit)."""

    def __init__(self, min_interval: float = 0.0):
        self.session = requests.Session()
        self.session.headers["User-Agent"] = USER_AGENT
        self.min_interval = min_interval
        self._last = 0.0

    def get(self, url: str, **kwargs: Any) -> requests.Response:
        kwargs.setdefault("timeout", HTTP_TIMEOUT)
        last_error = ""
        for attempt in range(HTTP_RETRIES + 1):
            if self.min_interval:
                time.sleep(max(0.0, self._last + self.min_interval - time.monotonic()))
                self._last = time.monotonic()
            backoff = min(60.0, 2 ** attempt + random.uniform(0, 1))
            try:
                response = self.session.get(url, **kwargs)
            except (requests.ConnectionError, requests.Timeout) as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                log.warning("Network error on %s (%s)", url, last_error)
                time.sleep(backoff)
                continue
            if response.status_code == 429:
                try:
                    delay = float(response.headers.get("Retry-After", "")) + 1
                except ValueError:
                    delay = backoff
                if delay > MAX_RATE_LIMIT_WAIT:
                    raise NetworkError(f"Rate limited by {url.split('/')[2]} for {int(delay)}s. "
                                       "Wait a while and run again; finished songs are kept.")
                last_error = "HTTP 429 Too Many Requests"
                say(f"    [WAIT] Rate limit reached, waiting {int(delay)}s ...")
                time.sleep(delay)
                continue
            if response.status_code in (500, 502, 503, 504):
                last_error = f"HTTP {response.status_code}"
                time.sleep(backoff)
                continue
            return response
        raise NetworkError(f"Request to {url} failed after {HTTP_RETRIES + 1} attempts ({last_error})")

    def download(self, url: str, max_bytes: int = 25 * 1024 * 1024) -> Optional[bytes]:
        response = self.get(url, stream=True)
        if response.status_code != 200:
            return None
        chunks, total = [], 0
        for chunk in response.iter_content(65536):
            total += len(chunk)
            if total > max_bytes:
                raise NetworkError(f"Download from {url} is larger than {max_bytes} bytes")
            chunks.append(chunk)
        return b"".join(chunks)


# =============================================================================
# Reading tags from audio files
# =============================================================================

def _first(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (list, tuple)):
        return ", ".join(p for p in dict.fromkeys(_first(v) for v in value) if p)
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace").strip("\x00 ").strip()
    if hasattr(value, "value"):  # ASF attributes
        return _first(value.value)
    return str(value).strip("\x00 ").strip()


def _read_tags(audio: Any) -> dict[str, Any]:
    tags = audio.tags
    file_type = type(audio).__name__
    if isinstance(tags, ID3):  # MP3, AIFF, WAV
        def id3(frame_id: str) -> str:
            frames = tags.getall(frame_id)
            if not frames:
                return ""
            if frame_id == "TCON" and hasattr(frames[0], "genres"):
                return ", ".join(g for g in frames[0].genres if g)
            return _first(list(getattr(frames[0], "text", [])))
        return {"title": id3("TIT2"), "artist": id3("TPE1"), "album": id3("TALB"), "album_artist": id3("TPE2"),
                "year": parse_year(id3("TDRC") or id3("TYER") or id3("TORY")), "track": parse_int(id3("TRCK")),
                "disc": parse_int(id3("TPOS")), "genre": id3("TCON")}
    if tags is None:
        return {}
    if file_type in ("MP4", "EasyMP4"):
        return {"title": _first(tags.get("©nam")), "artist": _first(tags.get("©ART")),
                "album": _first(tags.get("©alb")), "album_artist": _first(tags.get("aART")),
                "year": parse_year(_first(tags.get("©day"))), "track": parse_int((tags.get("trkn") or [None])[0]),
                "disc": parse_int((tags.get("disk") or [None])[0]), "genre": _first(tags.get("©gen"))}
    if file_type == "ASF":  # WMA
        return {"title": _first(tags.get("Title")), "artist": _first(tags.get("Author") or tags.get("WM/AlbumArtist")),
                "album": _first(tags.get("WM/AlbumTitle")), "album_artist": _first(tags.get("WM/AlbumArtist")),
                "year": parse_year(_first(tags.get("WM/Year"))), "track": parse_int(_first(tags.get("WM/TrackNumber"))),
                "disc": parse_int(_first(tags.get("WM/PartOfSet"))), "genre": _first(tags.get("WM/Genre"))}
    if hasattr(tags, "get"):  # Vorbis comments: FLAC, Ogg, Opus

        def get(*keys: str) -> str:
            for key in keys:
                try:
                    value = tags.get(key)
                except (KeyError, ValueError):
                    value = None
                if value:
                    return _first(value)
            return ""
        return {"title": get("title"), "artist": get("artist"), "album": get("album"),
                "album_artist": get("albumartist", "album artist", "album_artist"),
                "year": parse_year(get("date", "year", "originaldate")), "track": parse_int(get("tracknumber", "track")),
                "disc": parse_int(get("discnumber", "disc")), "genre": get("genre")}
    return {}


def read_audio_metadata(path: Path) -> AudioMetadata:
    """Never raises: problems end up in ``error``."""
    try:
        stat = path.stat()
    except OSError as exc:
        return AudioMetadata(str(path), 0, 0, error=f"Cannot access file: {exc}")
    meta = AudioMetadata(str(path), stat.st_size, stat.st_mtime)
    if stat.st_size == 0:
        meta.error = "File is empty (0 bytes)"
        return meta
    try:
        audio = mutagen.File(path)
    except Exception as exc:
        meta.error = f"Corrupted or unreadable audio file ({type(exc).__name__}: {exc})"
        return meta
    if audio is None:
        meta.error = "Unrecognised or unsupported audio format"
        return meta
    length = getattr(getattr(audio, "info", None), "length", 0) or 0
    if length <= 0:
        meta.error = "Audio stream has zero length (file is probably corrupted)"
        return meta
    meta.duration_seconds = round(float(length), 3)
    try:
        values = _read_tags(audio)
    except Exception as exc:
        log.warning("Tag parsing failed for %s: %s", path, exc)
        values = {}
    for key, value in values.items():
        if value not in (None, ""):
            setattr(meta, key, value)
    return meta


def scan_folder(folder: Path) -> list[AudioMetadata]:
    """Read every audio file in a folder (and its sub-folders)."""
    files = sorted((p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() in AUDIO_EXTENSIONS
                    and not p.name.startswith((".", "~$"))), key=lambda p: str(p).casefold())
    library = []
    for index, path in enumerate(files, start=1):
        if index % 25 == 0 or index == len(files):
            say(f"  Reading tags: {index}/{len(files)}", end="\r")
        library.append(read_audio_metadata(path))
    if files:
        say()
    return library


def merge_metadata(local: AudioMetadata, track: PlaylistTrack) -> tuple[dict[str, Any], dict[str, str]]:
    """File tags win; YouTube info only fills gaps; nothing is invented."""
    playlist_values = {"title": track.title, "artist": track.artist, "album": track.album, "year": track.year}
    label = "filename" if track.source == "local" else track.source
    values: dict[str, Any] = {}
    sources: dict[str, str] = {}
    for name in METADATA_FIELDS:
        file_value = getattr(local, name, None)
        if file_value not in (None, ""):
            values[name], sources[name] = file_value, "file"
        elif playlist_values.get(name) not in (None, ""):
            values[name], sources[name] = playlist_values[name], label
        else:
            values[name], sources[name] = (None if name in ("year", "track", "disc") else ""), ""
    if not values["title"] or not values["artist"]:
        file_artist, file_title = parse_filename(Path(local.path).stem)
        if not values["artist"] and file_artist:
            values["artist"], sources["artist"] = file_artist, "filename"
        if not values["title"] and file_title:
            values["title"], sources["title"] = file_title, "filename"
    return values, sources


# =============================================================================
# Album covers: embedded art > YouTube thumbnail > MusicBrainz / Cover Art Archive
# =============================================================================

def extract_artwork(path: Path) -> Optional[bytes]:
    """Bytes of the front cover embedded in an audio file, or None."""
    try:
        audio = mutagen.File(path)
    except Exception:
        return None
    if audio is None:
        return None
    tags = audio.tags
    try:
        if isinstance(tags, ID3):
            frames = tags.getall("APIC")
            if frames:
                return next((f for f in frames if f.type == 3), frames[0]).data
        elif type(audio).__name__ == "MP4" and tags and tags.get("covr"):
            return bytes(tags["covr"][0])
        elif getattr(audio, "pictures", None):
            return next((p for p in audio.pictures if p.type == 3), audio.pictures[0]).data
        elif tags is not None and hasattr(tags, "get") and tags.get("metadata_block_picture"):
            for encoded in tags["metadata_block_picture"]:
                picture = Picture(base64.b64decode(encoded))
                if picture.data:
                    return picture.data
        elif type(audio).__name__ == "ASF" and tags and tags.get("WM/Picture"):
            data = tags["WM/Picture"][0].value  # type(1) size(4) mime\0 description\0 data (UTF-16)
            size, pos = struct.unpack_from("<I", data, 1)[0], 5
            for _ in range(2):
                while data[pos:pos + 2] != b"\x00\x00":
                    pos += 2
                pos += 2
            return data[pos:pos + size]
    except Exception as exc:
        log.warning("Could not extract artwork from %s: %s", path, exc)
    return None


@dataclass
class Artwork:
    cover_filename: str                 # "" when missing
    status: str
    embed: Optional[tuple[bytes, str]]  # (bytes, mime) to put inside the audio file


def to_png(data: bytes, square_crop: bool = False) -> tuple[bytes, str]:
    Image.MAX_IMAGE_PIXELS = 6000 * 6000
    with Image.open(io.BytesIO(data)) as image:
        mime = Image.MIME.get(image.format or "", "image/jpeg")
        image.load()
        if square_crop and image.width != image.height:  # 16:9 YouTube thumbnail -> square cover
            side = min(image.width, image.height)
            left, top = (image.width - side) // 2, (image.height - side) // 2
            image = image.crop((left, top, left + side, top + side))
        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA" if "A" in image.getbands() or image.mode == "P" else "RGB")
        out = io.BytesIO()
        image.save(out, format="PNG", optimize=True)
        return out.getvalue(), mime


def png_to_jpeg(png: bytes) -> bytes:
    with Image.open(io.BytesIO(png)) as image:
        out = io.BytesIO()
        image.convert("RGB").save(out, format="JPEG", quality=92)
        return out.getvalue()


class ArtworkManager:
    """Finds each album's cover once, saves it as "<Album>.png" and reuses it.

    ``state`` is the resume state: it remembers covers (``state.covers``) and their names
    (``state.cover_names``) between runs.
    """

    def __init__(self, http: HttpClient, covers_dir: Path, cache_dir: Path, state: Any):
        self.http = http
        self.covers_dir = covers_dir
        self.cache_dir = cache_dir
        self.state = state
        self.musicbrainz = HttpClient(min_interval=1.1)  # MusicBrainz allows ~1 request/second
        self._missing: set[str] = set()

    def _cache(self, key: str) -> Path:
        return self.cache_dir / (hashlib.sha1(key.encode("utf-8")).hexdigest() + ".bin")

    def _cover_art_archive(self, album: str, artist: str) -> Optional[bytes]:
        def esc(text: str) -> str:
            return re.sub(r'([+\-&|!(){}\[\]^"~*?:\\/])', r"\\\1", text)
        response = self.musicbrainz.get("https://musicbrainz.org/ws/2/release-group/", params={
            "query": f'releasegroup:"{esc(album)}" AND artist:"{esc(artist)}"', "fmt": "json", "limit": 5})
        if response.status_code != 200:
            return None
        for group in response.json().get("release-groups", []):
            if int(group.get("score", 0)) < 90 or normalize(group.get("title", "")) != normalize(album):
                continue
            credited = normalize_artist(" ".join(c.get("name", "") for c in group.get("artist-credit", [])
                                                 if isinstance(c, dict)))
            if normalize_artist(artist) not in credited and credited not in normalize_artist(artist):
                continue
            data = self.musicbrainz.download(f"https://coverartarchive.org/release-group/{group['id']}/front-500")
            if data:
                return data
        return None

    def get(self, values: dict[str, Any], audio_path: Path, track: PlaylistTrack, crop_embedded: bool) -> Artwork:
        album = (values.get("album") or "").strip()
        album_artist = (values.get("album_artist") or values.get("artist") or "").strip()
        if album:
            key, base_name = f"{normalize_artist(album_artist)}|{normalize(album)}", album
        else:
            key = f"track|{normalize_artist(values.get('artist') or '')}|{normalize(values.get('title') or '')}"
            base_name = " - ".join(p for p in (values.get("artist"), values.get("title")) if p) or "Unknown"

        saved = self.state.covers.get(key)
        if saved and saved.get("filename") and (self.covers_dir / saved["filename"]).is_file():
            embed = None
            try:
                cached = self._cache(key)
                if cached.is_file():
                    data = cached.read_bytes()
                    embed = (data, "image/png" if data[:8] == b"\x89PNG\r\n\x1a\n" else "image/jpeg")
                else:
                    embed = (png_to_jpeg((self.covers_dir / saved["filename"]).read_bytes()), "image/jpeg")
            except (OSError, UnidentifiedImageError) as exc:
                log.warning("Could not load artwork for embedding: %s", exc)
            return Artwork(saved["filename"], f"reused ({saved.get('source', 'cached')})", embed)

        attempts: list[str] = []
        found: Optional[tuple[bytes, str, bool]] = None  # (bytes, source, crop)
        embedded = extract_artwork(audio_path)
        if embedded:
            found = (embedded, "embedded", crop_embedded)
        else:
            attempts.append("no embedded art")
        if found is None and track.artwork_url:
            try:
                data = self.http.download(track.artwork_url)
                if not data and "/maxresdefault." in track.artwork_url:
                    data = self.http.download(track.artwork_url.replace("/maxresdefault.", "/hqdefault."))
                if data:
                    found = (data, track.source, True)
                else:
                    attempts.append("YouTube image unavailable")
            except NetworkError as exc:
                attempts.append(f"YouTube image download failed ({exc})")
        if found is None and album and album_artist and key not in self._missing:
            try:
                data = self._cover_art_archive(album, album_artist)
                if data:
                    found = (data, "coverartarchive", False)
                else:
                    attempts.append("not on Cover Art Archive")
            except NetworkError as exc:
                attempts.append(f"MusicBrainz lookup failed ({exc})")
        if found is None:
            self._missing.add(key)
            return Artwork("", "missing: " + ("; ".join(attempts) or "no source available"), None)

        data, source, crop = found
        try:
            png, mime = to_png(data, square_crop=crop)
        except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as exc:
            return Artwork("", f"missing: {source} image is not a valid picture ({exc})", None)

        stem = sanitize_filename(base_name, "Unknown Album")
        had_name = self.state.cover_names.get(key)
        filename = self.state.cover_names.claim(key, stem, ".png", self.covers_dir)
        if filename != f"{stem}.png" and album and album_artist and not had_name:
            # Another album has this name: "Album (Artist).png" beats "Album (2).png".
            self.state.cover_names.release(key)
            filename = self.state.cover_names.claim(key, sanitize_filename(f"{album} ({album_artist})"),
                                                    ".png", self.covers_dir)
        atomic_write(self.covers_dir / filename, png)
        embed_data, embed_mime = (data, mime) if not crop else (png_to_jpeg(png), "image/jpeg")
        if embed_mime not in ("image/jpeg", "image/png"):
            embed_data, embed_mime = png_to_jpeg(png), "image/jpeg"
        atomic_write(self._cache(key), embed_data)
        self.state.covers[key] = {"filename": filename, "source": source, "album": album}
        return Artwork(filename, source, (embed_data, embed_mime))


# =============================================================================
# FFmpeg: probe, convert to a real MP4 (AAC), tag, check
# =============================================================================

LOSSLESS_CODECS = {"flac", "alac", "wavpack", "ape", "tta", "mlp", "truehd", "shorten", "tak"}


class ConversionError(Exception):
    pass


def no_ctrl_c() -> dict[str, Any]:
    """Run FFmpeg in its own process group so Ctrl+C does not kill a conversion half-way."""
    if os.name == "nt":
        return {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP}
    return {"start_new_session": True}


def find_ffmpeg() -> bool:
    """Locate ffmpeg/ffprobe on PATH, or where 'winget install Gyan.FFmpeg' puts them."""
    global FFMPEG, FFPROBE
    if shutil.which("ffmpeg") and shutil.which("ffprobe"):
        FFMPEG, FFPROBE = shutil.which("ffmpeg"), shutil.which("ffprobe")
        return True
    winget = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "WinGet"
    program_files = Path(os.environ.get("ProgramFiles", "C:/Program Files"))
    for root, pattern in ((winget, "Links/ffmpeg.exe"), (winget, "Packages/*/*/bin/ffmpeg.exe"),
                          (program_files, "ffmpeg*/bin/ffmpeg.exe"), (Path("C:/"), "ffmpeg*/bin/ffmpeg.exe"),
                          (Path("C:/"), "ffmpeg*/*/bin/ffmpeg.exe")):
        if not root.is_dir():
            continue
        for exe in root.glob(pattern):
            probe_exe = exe.with_name("ffprobe.exe")
            if probe_exe.is_file():
                FFMPEG, FFPROBE = str(exe), str(probe_exe)
                return True
    return False


def probe(path: Path) -> dict[str, Any]:
    try:
        result = subprocess.run([FFPROBE, "-v", "error", "-print_format", "json", "-show_format",
                                 "-show_streams", str(path)], capture_output=True, text=True,
                                encoding="utf-8", errors="replace", timeout=120, **no_ctrl_c())
    except (subprocess.TimeoutExpired, OSError) as exc:
        raise ConversionError(f"ffprobe failed: {exc}") from None
    if result.returncode != 0:
        lines = (result.stderr or "").strip().splitlines()
        raise ConversionError(f"ffprobe cannot read the file: {lines[-1] if lines else 'unknown error'}")
    try:
        data = json.loads(result.stdout or "{}")
    except json.JSONDecodeError as exc:
        raise ConversionError(f"ffprobe returned invalid output: {exc}") from exc
    audio = [s for s in data.get("streams", []) if s.get("codec_type") == "audio"]
    if not audio:
        raise ConversionError("File contains no audio stream")
    data["audio_stream"] = audio[0]
    return data


def probe_duration(info: dict[str, Any]) -> Optional[float]:
    for value in (info.get("format", {}).get("duration"), info.get("audio_stream", {}).get("duration")):
        try:
            if value is not None and float(value) > 0:
                return float(value)
        except (TypeError, ValueError):
            pass
    return None


def probe_bitrate(info: dict[str, Any]) -> Optional[int]:
    for value in (info.get("audio_stream", {}).get("bit_rate"), info.get("format", {}).get("bit_rate")):
        try:
            if value is not None and int(value) > 0:
                return int(round(int(value) / 1000))
        except (TypeError, ValueError):
            pass
    return None


def convert(source: Path, info: dict[str, Any], work_dir: Path) -> tuple[Path, str]:
    """Convert into a temporary .part file. AAC is copied untouched; everything else becomes AAC.

    Returns (temporary file, description of what was done).
    """
    stream = info["audio_stream"]
    codec = (stream.get("codec_name") or "").lower()
    source_kbps = probe_bitrate(info)
    if codec == "aac":
        container = (info.get("format", {}).get("format_name") or "").lower()
        args = ["-c:a", "copy"] + (["-bsf:a", "aac_adtstoasc"] if container in ("aac", "adts") else [])
        description = "stream copy (AAC, no re-encode)"
    elif codec == "alac":
        args, description = ["-c:a", "copy"], "stream copy (ALAC, no re-encode)"
    else:
        lossless = codec in LOSSLESS_CODECS or codec.startswith("pcm_")
        # Lossy sources (MP3, Opus...): match the source bitrate, 192-320 kbps, to limit quality loss.
        bitrate = LOSSLESS_AAC_BITRATE if lossless or not source_kbps else \
            max(192, min(320, ((source_kbps + 31) // 32) * 32))
        args = ["-c:a", "aac", "-b:a", f"{bitrate}k"]
        if int(stream.get("sample_rate") or 0) > 48000:
            args += ["-ar", "48000"]
        label = codec.upper() + (f" {source_kbps} kbps" if source_kbps and not lossless else "")
        description = f"{label} -> AAC {bitrate} kbps"

    work_dir.mkdir(parents=True, exist_ok=True)
    tmp = work_dir / (source.stem[:80] + ".part" + OUTPUT_EXTENSION)
    tmp.unlink(missing_ok=True)
    cmd = [FFMPEG, "-hide_banner", "-nostdin", "-v", "error", "-y", "-i", str(source),
           "-map", "0:a:0", "-vn", "-sn", "-dn", "-map_metadata", "-1", "-map_chapters", "-1",
           *args, "-movflags", "+faststart", "-f", "mp4", str(tmp)]
    log.debug("FFmpeg: %s", " ".join(cmd))
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                                timeout=FFMPEG_TIMEOUT, **no_ctrl_c())
    except (subprocess.TimeoutExpired, OSError) as exc:
        tmp.unlink(missing_ok=True)
        raise ConversionError(f"FFmpeg failed: {exc}") from None
    if result.returncode != 0 or not tmp.is_file() or tmp.stat().st_size == 0:
        tmp.unlink(missing_ok=True)
        lines = (result.stderr or "").strip().splitlines()
        raise ConversionError(f"FFmpeg failed: {lines[-1] if lines else 'no output produced'}")
    return tmp, description


def write_tags(path: Path, values: dict[str, Any], artwork: Optional[tuple[bytes, str]]) -> None:
    audio = MP4(path)
    if audio.tags is None:
        audio.add_tags()
    tags = audio.tags
    for name, atom in {"title": "©nam", "artist": "©ART", "album": "©alb", "album_artist": "aART",
                       "genre": "©gen"}.items():
        if values.get(name):
            tags[atom] = [str(values[name])]
    if values.get("year"):
        tags["©day"] = [str(values["year"])]
    if values.get("track"):
        tags["trkn"] = [(int(values["track"]), int(values.get("track_total") or 0))]
    if values.get("disc"):
        tags["disk"] = [(int(values["disc"]), 0)]
    if artwork:
        fmt = MP4Cover.FORMAT_PNG if "png" in artwork[1] else MP4Cover.FORMAT_JPEG
        tags["covr"] = [MP4Cover(artwork[0], imageformat=fmt)]
    audio.save()


def check_export(path: Path, expected_duration: Optional[float]) -> tuple[list[str], dict[str, Any]]:
    """Problems with an exported file (empty list = fine), plus its ffprobe info."""
    try:
        info = probe(path)
    except ConversionError as exc:
        return [str(exc)], {}
    problems = []
    container = info.get("format", {}).get("format_name", "")
    if not any(c in container for c in ("mp4", "m4a", "mov")):
        problems.append(f"container is '{container}', not MP4")
    if info["audio_stream"].get("codec_name") not in ("aac", "alac"):
        problems.append(f"audio codec is '{info['audio_stream'].get('codec_name')}', expected AAC")
    duration = probe_duration(info)
    if not duration:
        problems.append("audio duration is zero or unknown")
    elif expected_duration and abs(duration - expected_duration) > max(2.0, expected_duration * 0.02):
        problems.append(f"duration {duration:.1f}s differs from source {expected_duration:.1f}s")
    return problems, info
