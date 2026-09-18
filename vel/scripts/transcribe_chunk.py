import json, time, sys
from faster_whisper import WhisperModel
wav, out_path, offset = sys.argv[1], sys.argv[2], float(sys.argv[3])
t0=time.time()
model = WhisperModel("medium", device="cpu", compute_type="int8", cpu_threads=4)
segments, info = model.transcribe(wav, task="translate", language="hi",
    vad_filter=True, beam_size=5, condition_on_previous_text=False,
    temperature=[0.0,0.2,0.4,0.6], compression_ratio_threshold=2.2, no_speech_threshold=0.5)
out=open(out_path,"w",encoding="utf-8"); n=0
for seg in segments:
    out.write(json.dumps({"start":seg.start+offset,"end":seg.end+offset,"text":seg.text})+"\n"); out.flush()
    n+=1
    if n%50==0: print("%s: %5.1f min | wall %.0f min" % (out_path, seg.end/60, (time.time()-t0)/60), flush=True)
out.close()
print("%s DONE %d segs %.1f min wall" % (out_path, n, (time.time()-t0)/60), flush=True)
