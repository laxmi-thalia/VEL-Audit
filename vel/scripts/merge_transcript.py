import json, sys
segs=[]
for p in ("part0.jsonl","part1.jsonl","part2.jsonl"):
    try:
        for l in open(p,encoding="utf-8"):
            s=json.loads(l)
            if s["text"].strip(): segs.append(s)
    except FileNotFoundError: pass
segs.sort(key=lambda s:s["start"])
# drop overlap dupes: same-ish start (<1.5s) or contained interval with same text head
out=[]; 
for s in segs:
    if out and s["start"] < out[-1]["end"] - 0.5:
        continue   # overlap zone: keep first source's version
    out.append(s)
def ts(x): return "%02d:%02d" % (x//60, x%60)
paras=[]; cur=None
for s in out:
    t=s["text"].strip()
    if cur and s["start"]-cur["end"]<=2.5 and len(cur["text"])<500:
        cur["text"]+=" "+t; cur["end"]=s["end"]
    else:
        if cur: paras.append(cur)
        cur={"start":s["start"],"end":s["end"],"text":t}
if cur: paras.append(cur)
md=["# Meeting with Rashid Faisal — 27-Aug-2026, 83 min","",
"English translation of the Hinglish/Gujarati conversation. Transcribed locally",
"(faster-whisper medium, translate task) — the audio never left this machine.",
"Speakers are NOT labelled (single-model pass); timestamps are into the recording.",
"Machine translation of code-switched speech — wording is approximate; verify anything",
"load-bearing against the recording itself.","","---",""]
for p in paras:
    md.append("**(%s)** %s" % (ts(p["start"]), p["text"])); md.append("")
open(r"C:\Users\pawar\Downloads\rashid-meeting-transcript.md","w",encoding="utf-8").write("\n".join(md))
print("segments merged:",len(out),"| paragraphs:",len(paras),"| words:",sum(len(p['text'].split()) for p in paras),"| coverage to %.1f min"%out[-1]["end"]/1 if False else "| coverage to %.1f min"%(out[-1]["end"]/60))
