#!/usr/bin/env bash
# get_audio.sh: download the audio of the chosen videos.
# Step 1 (videos). Version 1.1, 23 October 2026.
#
# Downloads the audio of the chosen videos (the To transcribe sheet of harrison_v7_VIDEO_SELECTION.xlsx) with
# yt-dlp, one file per video, named by its YouTube ID (transcribe_videos.py matches the files to the V IDs).
# It reads a list written by select_videos_v2.py, which must be in the same folder as this script:
#   video_audio_list_pilot.txt   one act's videos, to try the download first
#   video_audio_list_all.txt     every chosen video, priority 1 first
#
# Run it on a local computer, in a terminal (YouTube refuses downloads from Colab and other cloud machines):
#   bash get_audio.sh pilot
#   bash get_audio.sh all
# (zsh get_audio.sh ... works too.)
# It needs yt-dlp and Deno, the JavaScript runtime yt-dlp uses for YouTube; see the README for the install steps.
# By default they are at ~/yt-dlp and ~/.deno/bin/deno; set YTDLP and DENO to use other paths.
#
# Which audio track: some videos carry YouTube AI-dubbed tracks in other languages beside the original. The track
# marked "original" is taken first, then an English track, and only then the default track. The track taken for
# each video is recorded in video_audio/tracks.txt.
# Safe to run again: videos already downloaded are skipped (video_audio/downloaded.txt).
# On a Mac the computer is kept awake while it runs (caffeinate); leave the terminal window open.

cd "$(dirname "$0")" || exit 1
WHICH="${1:-pilot}"
LIST="video_audio_list_${WHICH}.txt"
YTDLP="${YTDLP:-$HOME/yt-dlp}"
DENO="${DENO:-$HOME/.deno/bin/deno}"
FORMAT='ba[format_note*=original][abr<=?96]/ba[format_note*=original]/ba[language^=en][abr<=?96]/ba[language^=en]/ba[abr<=?96]/ba'

if [ ! -f "$LIST" ]; then echo "There is no list called $LIST in $(pwd). Run select_videos_v2.py first."; exit 1; fi
if [ ! -x "$YTDLP" ]; then echo "yt-dlp is not at $YTDLP. Install it, or set YTDLP to where it is."; exit 1; fi
if [ ! -x "$DENO" ]; then echo "Deno is not at $DENO. Install it, or set DENO to where it is."; exit 1; fi

if command -v caffeinate >/dev/null 2>&1; then KEEP_AWAKE=(caffeinate -i); else KEEP_AWAKE=(); fi

mkdir -p video_audio
echo "Downloading the videos in $LIST into $(pwd)/video_audio"
echo "yt-dlp $("$YTDLP" --version), $("$DENO" --version | head -1)"
echo "Started $(date)" >> video_audio/download_log.txt

"${KEEP_AWAKE[@]}" "$YTDLP" \
  --js-runtimes "deno:$DENO" \
  --batch-file "$LIST" \
  --format "$FORMAT" \
  --output "video_audio/%(id)s.%(ext)s" \
  --download-archive video_audio/downloaded.txt \
  --print-to-file "after_move:%(id)s | format %(format_id)s | language %(language)s | %(format_note)s" video_audio/tracks.txt \
  --no-playlist --ignore-errors --no-overwrites \
  --retries 10 --fragment-retries 10 \
  --sleep-requests 1 --sleep-interval 2 --max-sleep-interval 6 \
  2>&1 | tee -a video_audio/download_log.txt

N_FILES=$(ls video_audio | grep -v -e '\.txt$' -e '\.part$' -e '\.ytdl$' | wc -l | tr -d ' ')
N_ERRORS=$(grep -c '^ERROR' video_audio/download_log.txt)
echo
echo "Finished $(date). Audio files in video_audio: $N_FILES. Error lines in the log so far: $N_ERRORS."
echo "Finished $(date): $N_FILES files" >> video_audio/download_log.txt
