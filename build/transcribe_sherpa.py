import sys, json, time, wave, os
import numpy as np
import sherpa_onnx
from faster_whisper.vad import get_speech_timestamps, VadOptions

size = sys.argv[1]            # small | medium
lang = sys.argv[2]            # en | "" (auto)
maxwin = float(sys.argv[3])   # seconds per window
files = sys.argv[4:]
mdir = f"models/sherpa-onnx-whisper-{size}"
t0 = time.time()

def load(lang, task):
    return sherpa_onnx.OfflineRecognizer.from_whisper(
        encoder=f"{mdir}/{size}-encoder.int8.onnx",
        decoder=f"{mdir}/{size}-decoder.int8.onnx",
        tokens=f"{mdir}/{size}-tokens.txt",
        language=lang, task=task, num_threads=4, decoding_method="greedy_search", tail_paddings=-1)

rec = load(lang, "transcribe")
print("model loaded", size, round(time.time()-t0,1), "s", flush=True)

def read_wav(path):
    with wave.open(path) as w:
        assert w.getframerate()==16000 and w.getnchannels()==1
        data = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16).astype(np.float32)/32768.0
    return data

for f in files:
    audio = read_wav(f)
    sr = 16000
    # VAD -> speech chunks, then merge into windows <= 28 s
    ts = get_speech_timestamps(audio, VadOptions(threshold=0.4, min_speech_duration_ms=200, min_silence_duration_ms=500, speech_pad_ms=300, max_speech_duration_s=maxwin))
    windows = []
    cur = None
    for s in ts:
        st, en = s["start"], s["end"]
        if cur and (en - cur[0]) <= maxwin*sr and (st - cur[1]) < 1.5*sr:
            cur[1] = en
        else:
            if cur: windows.append(tuple(cur))
            cur = [st, en]
    if cur: windows.append(tuple(cur))
    print(f"== {f}: {len(audio)/sr:.1f}s, {len(ts)} speech spans -> {len(windows)} windows", flush=True)
    out = {"file": f, "model": size, "duration": len(audio)/sr, "segments": []}
    for (st, en) in windows:
        chunk = audio[st:en]
        stream = rec.create_stream()
        stream.accept_waveform(sr, chunk)
        rec.decode_stream(stream)
        r = stream.result
        seg = {"start": st/sr, "end": en/sr, "text": r.text.strip(), "lang": getattr(r, "lang", "")}
        out["segments"].append(seg)
        print(f"[{st/sr:7.2f} -> {en/sr:7.2f}] ({seg['lang']}) {seg['text']}", flush=True)
    outname = f.replace(".wav", f".sherpa-{size}-{lang or 'auto'}-{int(maxwin)}s.json")
    json.dump(out, open(outname, "w"), indent=1, ensure_ascii=False)
    print("WROTE", outname, flush=True)
print("DONE", size, round(time.time()-t0,1), "s")
