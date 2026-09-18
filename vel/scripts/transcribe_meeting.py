"""Rashid meeting transcription v2 - anti-loop settings:
condition_on_previous_text=False (the repetition-loop fix), language hint, temperature fallback."""
import json, time
from faster_whisper import WhisperModel
t0=time.time()
model = WhisperModel("medium", device="cpu", compute_type="int8", cpu_threads=8)
print("model loaded %.0fs" % (time.time()-t0), flush=True)
segments, info = model.transcribe("meeting_audio.wav", task="translate", language="hi",
    vad_filter=True, beam_size=5,
    condition_on_previous_text=False,
    temperature=[0.0,0.2,0.4,0.6],
    compression_ratio_threshold=2.2, no_speech_threshold=0.5)
out=open("meeting_segments.jsonl","w",encoding="utf-8")
n=0
for seg in segments:
    out.write(json.dumps({"start":seg.start,"end":seg.end,"text":seg.text})+"\n"); out.flush()
    n+=1
    if n%50==0:
        print("progress: %5.1f min | %d segs | wall %.0f min" % (seg.end/60,n,(time.time()-t0)/60), flush=True)
out.close()
print("DONE: %d segments | wall %.1f min" % (n,(time.time()-t0)/60), flush=True)
