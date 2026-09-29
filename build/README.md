# Build scripts for IEN_301_Project_1_v2.pptx

Reproducible pipeline used to rebuild the deck (run from a scratch directory containing the original `orig.pptx`, its unpacked copy in `orig_unpacked/`, and the interview audio in `audio/`):

| Script | Purpose |
|---|---|
| `transcribe_sherpa.py <small|medium|large-v3> <lang|""> <window_s> file.wav…` | Whisper transcription via sherpa-onnx (models from the sherpa-onnx GitHub releases) with silero VAD windows; writes timestamped JSON. |
| `deck_lib.py` | Theme helpers reproducing slides 2–10 (fonts, colours, positions), plus `<p:timing>` animation and fade-transition injection. |
| `build_deck.py` | Builds `IEN_301_Project_1_v2.pptx` from `orig.pptx`: keeps slides 2–5, rebuilds the rest natively, writes fresh speaker notes, prints per-slide note word counts and computed numbers. |
| `verify_deck.py IEN_301_Project_1_v2.pptx` | Checks: no image-only slides, no banned words (solution / heater / device / fix / convenient), no leftover icon names, no old notes, every slide animated with a fade transition, XML parses, animation targets exist. |

Render / round-trip checks were done with LibreOffice (`soffice --headless --convert-to pdf|pptx`) and `pdftoppm`.
