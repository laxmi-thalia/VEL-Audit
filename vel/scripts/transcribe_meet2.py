"""Local transcription v2 for Hinglish meetings: faster-whisper medium/int8, task=translate (Hindi/Hinglish -> English),
language auto, lighter VAD so short remarks survive. Usage: python transcribe_meet2.py <mp4> <tag>  -> <tag>_segs2.json + <tag>_transcript.md"""
import sys, os, json, subprocess, time
from faster_whisper import WhisperModel
mp4, tag = sys.argv[1], sys.argv[2]; SP = os.path.dirname(os.path.abspath(__file__)); wav = os.path.join(SP, tag + ".wav")
if not os.path.exists(wav): subprocess.run(["ffmpeg", "-y", "-v", "error", "-i", mp4, "-vn", "-ac", "1", "-ar", "16000", wav], check=True)
t0 = time.time(); model = WhisperModel("medium", device="cpu", compute_type="int8", cpu_threads=6)
segs, info = model.transcribe(wav, task="translate", beam_size=5, vad_filter=True, condition_on_previous_text=False, temperature=[0.0, 0.2, 0.4],
                              vad_parameters={"min_silence_duration_ms": 300, "threshold": 0.35})
out = [{"start": round(s.start, 1), "end": round(s.end, 1), "text": s.text.strip()} for s in segs]
json.dump(out, open(os.path.join(SP, tag + "_segs2.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
md = ["# %s — transcript (English, machine translation of Hinglish; local faster-whisper medium)" % os.path.basename(mp4), "", "Detected language: %s (p=%.2f), %.1f min" % (info.language, info.language_probability, info.duration / 60), ""]
for s in out: md.append("- **[%02d:%02d]** %s" % (s["start"] // 60, s["start"] % 60, s["text"]))
open(os.path.join(SP, tag + "_transcript.md"), "w", encoding="utf-8").write("\n".join(md))
print("DONE %s: %d segments, lang %s, %.0fs" % (tag, len(out), info.language, time.time() - t0), flush=True)
