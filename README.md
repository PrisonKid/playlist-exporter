# Playlist Exporter

Turns a YouTube / YouTube Music playlist, or a folder of audio files you already have, into a `Music` folder of songs (MP4 audio), with album covers (PNG) and an Excel sheet of all the song info.

```text
Music/
├── Batch 001/          20 songs + their album covers + metadata.xlsx
├── Batch 002/
├── ...
└── metadata.xlsx       every song, plus a report of what failed and why
```


## Setup (once)

1. **Install Python 3.10 or newer** from <https://www.python.org/downloads/>.
   In the installer, tick **"Add python.exe to PATH"**.
2. **Download this project**: on the GitHub page click **Code → Download ZIP**, then unzip it
   (or `git clone` it if you use git). Keep the three `.py` files together in the same folder.

That's it. The first time you run the script it offers to install the Python packages it needs
and, on Windows, FFmpeg (the audio converter) as well. Just answer `Y`.

## Run it

Open the project folder, click the address bar, type `cmd` and press Enter. Then:

```cmd
python playlist_exporter.py
```

It asks what you want to do:

1. **Download a YouTube / YouTube Music playlist**: paste the playlist link.
2. **Convert audio files I already have** (MP3, FLAC, M4A, WAV, ...): give the folder they are in.
   Sub-folders are included, and the tags already in the files are used.

You can also put the link straight on the command line, in quotes:

```cmd
python playlist_exporter.py "https://music.youtube.com/playlist?list=..."
```

The playlist must be **Public** or **Unlisted** (not Private).

**Stopping and continuing:** press `Ctrl+C` to stop after the current song. Run the same command again
and it carries on where it stopped. Songs that are already done are skipped. If you delete a song's
`.mp4` from `Music`, the next run downloads it again.

## When something goes wrong

| Problem | Fix |
|---|---|
| `python` is not recognized | Reinstall Python and tick **Add python.exe to PATH**, then open a new terminal |
| "Sign in to confirm you're not a bot" | Run with your browser's YouTube login: `python playlist_exporter.py --cookies chrome` (or `edge` / `firefox`) |
| Lots of downloads suddenly fail | Update the downloader: `python -m pip install -U yt-dlp` |
| FFmpeg install did not work | Run `winget install Gyan.FFmpeg`, close the terminal, open a new one and try again |
| `metadata.xlsx` is open in Excel | The script then saves a copy with the time in its name. Close Excel next time |

Each run writes a log to `%LOCALAPPDATA%\PlaylistExporter\last_run.log`. Send that file if you ask
someone for help. Progress and cached covers are kept in the same folder.

## Other options

```text
python playlist_exporter.py --limit 5         only the first 5 songs (try it out)
python playlist_exporter.py --fresh           redo everything, ignoring saved progress
python playlist_exporter.py --audio-dir "D:\My Songs"   option 2 without the questions
```

Settings such as the number of songs per batch (`SONGS_PER_BATCH = 20`) are at the top of
`playlist_exporter.py`.

### Spotify playlists

Spotify does not give out audio. Copy the playlist to YouTube first: use <https://www.tunemymusic.com>
(Spotify → YouTube Music), make the new playlist Public or Unlisted, and use that link with option 1.

### The files

| File | What it does |
|---|---|
| `playlist_exporter.py` | The one you run: settings, the menu and the export itself |
| `youtube.py` | Reads YouTube playlists and downloads the songs |
| `media.py` | Song tags, album covers and converting audio with FFmpeg |
