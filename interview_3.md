# Interview 3 — recording not received

The task brief lists three WhatsApp interview videos (≈2.8 min, ≈2.5 min and ≈4.4 min, recorded 27 Sept 2026) inside `drive-download-20260929T164016Z-1-001.zip`.

Only two files arrived in the upload folder for this session:

| File | Duration | Transcribed as |
|---|---|---|
| `WhatsApp Video 2026-09-27 at 16.41.50.mp4` | 2:48 | `interview_1.md` |
| `WhatsApp Video 2026-09-27 at 16.42.14.mp4` | 2:27 | `interview_2.md` |

The zip archive itself and the third (~4.4 min) video were **not** present, so there is no transcript for interview 3 and nothing from it appears on the slides.

## What to do when the third recording is available

1. Extract audio: `ffmpeg -i <video>.mp4 -ac 1 -ar 16000 -vn interview_3.wav`
2. Transcribe with the same pipeline (`scratchpad/transcribe_sherpa.py <model> en 15 interview_3.wav`) and write this file in the same table format as interviews 1–2.
3. Update `evidence.md`, the Method slide (count, total minutes, descriptor), the Quotes slide (if a better quote exists), and the placeholders listed in `TODO.md` (winter difference, home type, water-heating setup, tenure, morning routine).
