"""YouTube for playlist_exporter.py: reading playlists, downloading songs, finding track numbers.

You do not run this file yourself; playlist_exporter.py uses it.
"""

from __future__ import annotations

import logging
import re
import socket
import time
import urllib.parse
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Optional

from mutagen.mp4 import MP4

import media
from media import (PlaylistTrack, clean_title, normalize, normalize_artist, parse_year, say,
                   split_artist_title)

log = logging.getLogger("exporter")


class PlaylistError(Exception):
    pass


# =============================================================================
# Reading the playlist
# =============================================================================

YT_API = "https://www.googleapis.com/youtube/v3"
UNAVAILABLE_TITLES = {"deleted video", "private video", "[deleted video]", "[private video]"}


def playlist_id_from_url(url: str) -> str:
    text = (url or "").strip()
    if re.fullmatch(r"(PL|OL|UU|LL|FL|RD|OLAK5uy_)[A-Za-z0-9_-]{10,}", text):
        return text
    parsed = urllib.parse.urlparse(text)
    if parsed.netloc and "youtu" in parsed.netloc:
        query = urllib.parse.parse_qs(parsed.query)
        if query.get("list"):
            return query["list"][0]
    raise PlaylistError("That does not look like a YouTube playlist link. Expected something like "
                        "https://www.youtube.com/playlist?list=PL... or https://music.youtube.com/playlist?list=...")


def parse_auto_generated_description(description: str) -> dict[str, Any]:
    """Parse YouTube's 'Provided to YouTube by <Label>' descriptions of label-delivered songs.

    Provided to YouTube by <Label> / <Title> · <Artist> · <Artist2> / <Album> / ℗ 1986 / Released on: 1986-03-03
    """
    if not description or "Provided to YouTube by" not in description:
        return {}
    lines = [line.strip() for line in description.splitlines() if line.strip()]
    start = next((i for i, line in enumerate(lines) if line.startswith("Provided to YouTube by")), None)
    if start is None:
        return {}
    result: dict[str, Any] = {}
    if start + 1 < len(lines) and " · " in lines[start + 1]:
        parts = [p.strip() for p in lines[start + 1].split(" · ") if p.strip()]
        result["title"], result["artists"] = parts[0], parts[1:]
        if start + 2 < len(lines) and not lines[start + 2].startswith(("℗", "©", "Released on:", "Auto-generated")):
            result["album"] = lines[start + 2]
    for line in lines:
        if line.startswith("Released on:") and parse_year(line):
            result["year"] = parse_year(line)
        elif "year" not in result and line.startswith("℗") and parse_year(line):
            result["year"] = parse_year(line)
    return result


def derive_metadata(title: str, channel: str, description: str = "") -> dict[str, Any]:
    """Best guess of title/artists/album/year: description > '- Topic' channel > 'Artist - Title'."""
    parsed = parse_auto_generated_description(description)
    if parsed.get("title"):
        return {**parsed, "auto_generated": True}
    channel = (channel or "").strip()
    if channel.endswith(" - Topic"):
        return {"title": clean_title(title), "artists": [channel[:-len(" - Topic")].strip()], "auto_generated": True}
    artist, song = split_artist_title(title or "")
    if artist:
        return {"title": clean_title(song), "artists": [clean_title(artist)], "auto_generated": False}
    result: dict[str, Any] = {"title": clean_title(title or ""), "auto_generated": False}
    if re.search(r"vevo$", channel, re.IGNORECASE):
        result["artists"] = [re.sub(r"\s*vevo$", "", channel, flags=re.IGNORECASE)]
    return result


def best_thumbnail(thumbnails: Any, video_id: str) -> str:
    if isinstance(thumbnails, dict):
        for key in ("maxres", "standard", "high", "medium", "default"):
            if thumbnails.get(key, {}).get("url"):
                return thumbnails[key]["url"]
    if isinstance(thumbnails, list) and thumbnails:
        best = max(thumbnails, key=lambda t: (t.get("width") or 0) * (t.get("height") or 0))
        if best.get("url"):
            return best["url"]
    return f"https://i.ytimg.com/vi/{video_id}/hqdefault.jpg" if video_id else ""


def _iso_duration(value: str) -> Optional[float]:
    match = re.fullmatch(r"P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?", value or "")
    if not match or not any(match.groups()):
        return None
    d, h, m, s = (float(g) if g else 0.0 for g in match.groups())
    return d * 86400 + h * 3600 + m * 60 + s


def _entries_via_api(http: media.HttpClient, playlist_id: str, api_key: str) -> tuple[str, list[dict[str, Any]]]:
    """YouTube Data API: gives full descriptions, so album and year can be read."""
    def api(endpoint: str, params: dict[str, Any]) -> dict[str, Any]:
        response = http.get(f"{YT_API}/{endpoint}", params={**params, "key": api_key})
        if response.status_code == 200:
            return response.json()
        try:
            error = response.json().get("error", {})
            reason, message = (error.get("errors") or [{}])[0].get("reason", ""), error.get("message", "")
        except ValueError:
            reason, message = "", response.text[:200]
        if reason in ("quotaExceeded", "dailyLimitExceeded"):
            raise PlaylistError("YouTube API daily quota used up. Try again tomorrow; finished songs are kept.")
        if reason == "playlistNotFound" or response.status_code == 404:
            raise PlaylistError("YouTube playlist not found. Is it private? Make it Public or Unlisted.")
        raise PlaylistError(f"YouTube API refused the request ({reason or response.status_code}): {message}")

    info = api("playlists", {"part": "snippet", "id": playlist_id})
    if not info.get("items"):
        raise PlaylistError("YouTube playlist not found or private. Make it Public or Unlisted.")
    name = info["items"][0]["snippet"].get("title") or f"YouTube playlist {playlist_id}"
    items: list[dict[str, Any]] = []
    page_token = None
    while True:
        params = {"part": "snippet,contentDetails", "playlistId": playlist_id, "maxResults": 50}
        if page_token:
            params["pageToken"] = page_token
        page = api("playlistItems", params)
        items.extend(page.get("items") or [])
        say(f"  Fetched {len(items)}/{(page.get('pageInfo') or {}).get('totalResults', '?')} entries", end="\r")
        page_token = page.get("nextPageToken")
        if not page_token:
            break
    say()
    ids = [v for v in dict.fromkeys((i.get("contentDetails") or {}).get("videoId") for i in items) if v]
    details: dict[str, dict[str, Any]] = {}
    for start in range(0, len(ids), 50):
        for video in api("videos", {"part": "snippet,contentDetails", "id": ",".join(ids[start:start + 50])}).get("items") or []:
            details[video["id"]] = video
    entries = []
    for item in items:
        snippet = item.get("snippet") or {}
        video_id = (item.get("contentDetails") or {}).get("videoId") or ""
        video = details.get(video_id, {})
        vsnip = video.get("snippet") or {}
        entries.append({
            "id": video_id, "title": vsnip.get("title") or snippet.get("title") or "",
            "channel": vsnip.get("channelTitle") or snippet.get("videoOwnerChannelTitle") or "",
            "description": vsnip.get("description") or snippet.get("description") or "",
            "duration": _iso_duration((video.get("contentDetails") or {}).get("duration", "")),
            "thumbnail": best_thumbnail(vsnip.get("thumbnails") or snippet.get("thumbnails"), video_id),
            "available": bool(video),
        })
    return name, entries


def _entries_via_ytdlp(playlist_id: str, cookies: str) -> tuple[str, list[dict[str, Any]]]:
    from yt_dlp import YoutubeDL
    from yt_dlp.utils import DownloadError as YtDlpError
    options: dict[str, Any] = {"extract_flat": "in_playlist", "skip_download": True, "quiet": True,
                               "no_warnings": True, "ignoreerrors": True, "logger": _YtDlpLogger()}
    if cookies:
        options["cookiesfrombrowser"] = (cookies,)
    info, last_error = None, None
    for attempt in range(3):
        try:
            with YoutubeDL(options) as ydl:
                info = ydl.extract_info(f"https://www.youtube.com/playlist?list={playlist_id}", download=False)
            if info:
                break
        except YtDlpError as exc:
            last_error = exc
            log.warning("yt-dlp attempt %d failed: %s", attempt + 1, exc)
    if not info:
        raise PlaylistError(f"Could not read the playlist: {last_error or 'no data returned'}. "
                            "Check the link and make the playlist Public or Unlisted.")
    entries = []
    for entry in info.get("entries") or []:
        if not entry:
            entries.append({"id": "", "title": "", "available": False})
            continue
        entries.append({
            "id": entry.get("id") or "", "title": entry.get("title") or "",
            "channel": entry.get("channel") or entry.get("uploader") or "",
            "description": entry.get("description") or "", "duration": entry.get("duration"),
            "thumbnail": best_thumbnail(entry.get("thumbnails"), entry.get("id") or ""), "available": True,
        })
    return info.get("title") or f"YouTube playlist {playlist_id}", entries


def get_playlist(http: media.HttpClient, url: str, api_key: str = "", cookies: str = ""
                 ) -> tuple[str, list[PlaylistTrack]]:
    """(playlist name, one PlaylistTrack per video)."""
    playlist_id = playlist_id_from_url(url)
    if api_key:
        say("  Using the YouTube Data API for song info.")
        name, entries = _entries_via_api(http, playlist_id, api_key)
    else:
        name, entries = _entries_via_ytdlp(playlist_id, cookies)
    tracks: list[PlaylistTrack] = []
    seen: dict[str, int] = {}
    for position, entry in enumerate(entries, start=1):
        video_id, title = entry.get("id") or "", entry.get("title") or ""
        base = f"youtube:{video_id}" if video_id else f"youtube-missing:{position}"
        seen[base] = seen.get(base, 0) + 1
        uid = base if seen[base] == 1 else f"{base}#{seen[base]}"
        if not entry.get("available") or title.strip().lower() in UNAVAILABLE_TITLES:
            tracks.append(PlaylistTrack(position, uid, "youtube", title=title or f"(entry #{position})",
                                        source_id=video_id,
                                        unavailable_reason="Video is deleted, private or unavailable"))
            continue
        meta = derive_metadata(title, entry.get("channel") or "", entry.get("description", ""))
        tracks.append(PlaylistTrack(
            position, uid, "youtube", title=meta.get("title", ""), artists=meta.get("artists", []),
            album=meta.get("album", ""), year=meta.get("year"), duration_seconds=entry.get("duration"),
            # Thumbnails are album covers only for auto-generated music uploads.
            artwork_url=entry.get("thumbnail", "") if meta.get("auto_generated") else "",
            source_id=video_id,
        ))
    return name, tracks


# =============================================================================
# Downloading (yt-dlp)
# =============================================================================

_ANSI = re.compile(r"\x1b\[[0-9;]*m")
_NETWORK_HINTS = (
    "getaddrinfo", "failed to resolve", "name or service not known", "nodename nor servname",
    "temporary failure in name resolution", "network is unreachable", "no route to host",
    "connection refused", "connection reset", "connection aborted", "timed out", "transporterror",
    "remote end closed connection", "unable to connect", "winerror 10051", "winerror 10060", "winerror 10065",
)


class DownloadError(Exception):
    pass


class NetworkDownError(DownloadError):
    """The internet connection is down - not a problem with the song."""


def is_network_error(message: str) -> bool:
    return any(hint in message.lower() for hint in _NETWORK_HINTS)


def internet_available(host: str = "www.youtube.com") -> bool:
    try:
        with socket.create_connection((host, 443), timeout=5):
            return True
    except OSError:
        return False


class _YtDlpLogger:
    """Send yt-dlp's chatter to the log file instead of the screen."""

    def debug(self, msg: str) -> None:
        if not msg.startswith("[debug] "):
            log.debug("yt-dlp: %s", msg)

    def info(self, msg: str) -> None:
        log.debug("yt-dlp: %s", msg)

    def warning(self, msg: str) -> None:
        log.warning("yt-dlp: %s", msg)

    def error(self, msg: str) -> None:
        log.error("yt-dlp: %s", _ANSI.sub("", msg))


def _explain(message: str) -> str:
    message = _ANSI.sub("", message).replace("ERROR: ", "").strip()
    lower = message.lower()
    if "sign in to confirm" in lower or "not a bot" in lower:
        return (f"{message} -> YouTube wants a bot check. Run again with  --cookies chrome  "
                "(or edge / firefox: a browser where you are logged in to YouTube).")
    if "confirm your age" in lower or ("age" in lower and "restrict" in lower):
        return f"{message} -> age-restricted; run again with  --cookies chrome  (a logged-in browser)."
    if "private video" in lower or "video unavailable" in lower or "has been removed" in lower:
        return f"Video is unavailable: {message}"
    if "not available in your country" in lower or "geo" in lower:
        return f"Blocked in your country: {message}"
    return message


def download(video_id: str, folder: Path, cookies: str = "", pause: float = 1.0) -> tuple[Path, dict[str, Any]]:
    """Download one video's audio (preferring the AAC stream, so it is copied without re-encoding).

    Returns (file, song info that YouTube Music attached to the video).
    """
    from yt_dlp import YoutubeDL
    from yt_dlp.utils import DownloadError as YtDlpError
    if not video_id:
        raise DownloadError("No YouTube video id for this entry")
    folder.mkdir(parents=True, exist_ok=True)
    options: dict[str, Any] = {
        "format": "bestaudio[ext=m4a]/bestaudio[acodec^=mp4a]/bestaudio/best",
        "outtmpl": str(folder / "%(title).120B [%(id)s].%(ext)s"),
        "windowsfilenames": True, "noplaylist": True, "quiet": True, "no_warnings": True, "noprogress": True,
        "logger": _YtDlpLogger(), "retries": 10, "fragment_retries": 10, "extractor_retries": 3,
        "socket_timeout": media.HTTP_TIMEOUT, "overwrites": True, "continuedl": False,
        "writethumbnail": True,  # the thumbnail and tags are written INTO the file
        "postprocessors": [{"key": "FFmpegMetadata", "add_metadata": True, "add_chapters": False},
                           {"key": "EmbedThumbnail", "already_have_thumbnail": False}],
    }
    if Path(media.FFMPEG).is_file():
        options["ffmpeg_location"] = str(Path(media.FFMPEG).parent)
    if pause > 0:
        options["sleep_interval"], options["max_sleep_interval"] = pause, pause * 2
    if cookies:
        options["cookiesfrombrowser"] = (cookies,)
    try:
        with YoutubeDL(options) as ydl:
            info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=True)
    except YtDlpError as exc:
        if is_network_error(str(exc)):
            raise NetworkDownError(_ANSI.sub("", str(exc))) from None
        raise DownloadError(_explain(str(exc))) from None
    except Exception as exc:
        if is_network_error(str(exc)):
            raise NetworkDownError(str(exc)) from None
        raise DownloadError(f"yt-dlp failed: {type(exc).__name__}: {exc}") from None
    if not info:
        raise DownloadError("yt-dlp returned no information for this video")
    path = next((Path(p) for item in info.get("requested_downloads") or []
                 for p in [item.get("filepath") or item.get("_filename")] if p and Path(p).is_file()), None)
    if path is None:
        found = [p for p in folder.glob(f"*[[]{video_id}[]].*") if not p.name.endswith(".part")]
        path = found[0] if found else None
    if path is None or path.stat().st_size == 0:
        raise DownloadError("Download finished but no audio file was produced")

    artists = info.get("artists") or ([info["artist"]] if info.get("artist") else [])
    if len(artists) == 1 and ", " in artists[0]:
        artists = [a.strip() for a in artists[0].split(", ") if a.strip()]
    release_year = info.get("release_year") or parse_year(info.get("release_date"))
    try:  # yt-dlp writes the UPLOAD date as year and guesses; fix the tags it wrote
        audio = MP4(path)
        if audio.tags is None:
            audio.add_tags()
        if release_year:
            audio.tags["\xa9day"] = [str(release_year)]
        else:
            audio.tags.pop("\xa9day", None)
        if info.get("album") and artists and "aART" not in audio.tags:
            audio.tags["aART"] = [artists[0]]
        for key in ("trkn", "disk", "\xa9gen"):
            audio.tags.pop(key, None)
        audio.save()
    except Exception:
        pass  # not an M4A download (e.g. Opus); the playlist info is used instead
    return path, {"title": info.get("track") or "", "artists": [a for a in artists if a],
                  "album": info.get("album") or "", "year": release_year, "duration": info.get("duration")}


# =============================================================================
# Track numbers from YouTube Music album pages
# =============================================================================

_EDITION = re.compile(r"\s*[\(\[][^\)\]]*(deluxe|edition|remaster|expanded|anniversary|version|bonus)[^\)\]]*[\)\]]\s*",
                      re.IGNORECASE)


def same_album(a: str, b: str) -> bool:
    ka, kb = normalize(_EDITION.sub(" ", a or "")), normalize(_EDITION.sub(" ", b or ""))
    return bool(ka and kb) and (ka == kb or SequenceMatcher(None, ka, kb).ratio() >= 0.92)


def _same_title(a: str, b: str) -> bool:
    na, nb = normalize(a), normalize(b)
    return bool(na and nb) and (na == nb or SequenceMatcher(None, na, nb).ratio() >= 0.9)


@dataclass
class AlbumTrackInfo:
    album: str
    album_artist: str
    track: int
    track_count: Optional[int]
    year: Optional[int]


class YouTubeMusicAlbums:
    """Which album a video belongs to, and its position on that album (no account needed)."""

    def __init__(self):
        from ytmusicapi import YTMusic
        self.yt = YTMusic()
        self._albums: dict[str, Optional[dict[str, Any]]] = {}

    @staticmethod
    def _call(func, *args, **kwargs):
        for attempt in range(4):
            try:
                return func(*args, **kwargs)
            except Exception as exc:
                if attempt == 3 or not is_network_error(f"{type(exc).__name__} {exc}"):
                    raise
                time.sleep(5 * (attempt + 1))

    def _album(self, album_id: str) -> Optional[dict[str, Any]]:
        if album_id not in self._albums:
            try:
                self._albums[album_id] = self._call(self.yt.get_album, album_id)
            except Exception as exc:
                log.warning("YouTube Music album %s could not be read: %s", album_id, exc)
                self._albums[album_id] = None
        return self._albums[album_id]

    def _ref_for_video(self, video_id: str) -> Optional[tuple[str, str]]:
        try:
            watch = self._call(self.yt.get_watch_playlist, videoId=video_id, limit=1)
        except Exception as exc:
            log.warning("YouTube Music watch info failed for %s: %s", video_id, exc)
            return None
        first = (watch.get("tracks") or [{}])[0]
        album = first.get("album") or {}
        return (album["id"], video_id) if first.get("videoId") == video_id and album.get("id") else None

    def _ref_by_search(self, title: str, artists: list[str], duration: Optional[float]) -> Optional[tuple[str, str]]:
        """For music videos: find the album version of the same recording."""
        try:
            results = self._call(self.yt.search, f"{artists[0]} {title}", filter="songs", limit=10)
        except Exception as exc:
            log.warning("YouTube Music search failed for %s - %s: %s", artists[0], title, exc)
            return None
        wanted = [normalize_artist(a) for a in artists]
        for item in results:
            album = item.get("album") or {}
            names = [normalize_artist(a.get("name", "")) for a in item.get("artists") or []]
            artist_ok = any(n and w and (n == w or n in w or w in n) for n in names for w in wanted)
            if not (album.get("id") and _same_title(item.get("title", ""), title) and artist_ok):
                continue
            seconds = item.get("duration_seconds")
            if duration and seconds and abs(seconds - duration) > 5:
                continue  # live / extended / radio edit
            return album["id"], item.get("videoId") or ""
        return None

    def find(self, video_id: str, title: str, artists: list[str], album_name: str,
             duration: Optional[float]) -> Optional[AlbumTrackInfo]:
        ref = self._ref_for_video(video_id) if video_id else None
        if ref is None and title and artists:
            ref = self._ref_by_search(title, artists, duration)
        if ref is None:
            return None
        album = self._album(ref[0])
        if not album or (album_name and not same_album(album.get("title", ""), album_name)):
            return None
        tracks = album.get("tracks") or []
        number = next((item.get("trackNumber") or i for i, item in enumerate(tracks, start=1)
                       if ref[1] and item.get("videoId") == ref[1]), None)
        if number is None:
            number = next((item.get("trackNumber") or i for i, item in enumerate(tracks, start=1)
                           if _same_title(item.get("title", ""), title)), None)
        if number is None:
            return None
        album_artists = [a.get("name", "") for a in album.get("artists") or [] if a.get("name")]
        return AlbumTrackInfo(album.get("title", ""), album_artists[0] if album_artists else "", int(number),
                              album.get("trackCount") or (len(tracks) or None), parse_year(album.get("year")))
