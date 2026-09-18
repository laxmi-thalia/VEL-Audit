"""Build the GST knowledge base from 'gst manual_01072026.pdf' (Garg & Garg, 7th half-yearly
edition, 01-Jul-2026). Fully deterministic: split by the PDF's own bookmark tree, one .md per
leaf (section / rule / schedule), per-part index + master index. No LLM involved.
Output: .claude/skills/gst-audit/references/gst-manual/   (GITIGNORED - book is for personal use)"""
import os, re, fitz
PDF = r"C:\Users\pawar\Downloads\gst manual_01072026.pdf"
OUT = r"c:\PROJECTS\accountic\.claude\skills\gst-audit\references\gst-manual"
EDITION = "Garg & Garg, GST Manual, Seventh Half-Yearly Edition, updated to 01 July 2026"
def slug(t):
    s = re.sub(r"[^A-Za-z0-9]+", "-", t).strip("-").lower()
    return s[:70] or "untitled"
doc = fitz.open(PDF)
toc = doc.get_toc()
pages = {}  # lazy page text cache
def ptext(p):
    if p not in pages: pages[p] = doc[p - 1].get_text()
    return pages[p]
# flatten with spans: each entry runs to the next entry's page (inclusive of boundary page)
entries = []
for i, (lvl, title, page) in enumerate(toc):
    nxt = toc[i + 1][2] if i + 1 < len(toc) else doc.page_count + 1
    entries.append([lvl, title.strip(), page, nxt])
# determine current part (level-1) and leaf-ness (no deeper entry starts within my heading position)
PART = None; part_slug = None
counts = {}; master = []
os.makedirs(OUT, exist_ok=True)
for i, (lvl, title, page, nxt) in enumerate(entries):
    if lvl == 1:
        PART = title; part_slug = slug(title)
        if PART not in ("Title Pages", "Last Page"):
            os.makedirs(os.path.join(OUT, part_slug), exist_ok=True)
            counts[PART] = 0; master.append((PART, part_slug, page))
        continue
    if PART in ("Title Pages", "Last Page", None): continue
    is_parent = i + 1 < len(entries) and entries[i + 1][0] > lvl
    if is_parent: continue                      # chapters render via their children
    if title.startswith("Index_"): continue     # the book's own index pages
    end = min(nxt, entries[i + 1][2] if i + 1 < len(entries) else doc.page_count)
    end = max(page, end)
    text = "\n".join(ptext(p) for p in range(page, min(end + 1, doc.page_count + 1)))
    counts[PART] += 1
    fn = "%03d-%s.md" % (counts[PART], slug(title))
    body = ("# %s\n\n_Part: %s | PDF pages %d-%d | Source: %s._\n"
            "_Boundary pages may carry adjacent text; the section number in the heading is authoritative._\n\n%s\n"
            % (title, PART, page, end, EDITION, text.strip()))
    with open(os.path.join(OUT, part_slug, fn), "w", encoding="utf-8") as f:
        f.write(body)
# per-part indexes + master index
lines = ["# GST Manual knowledge base\n",
         "_%s._\n" % EDITION,
         "_Split by the PDF's own bookmarks; one file per section/rule/schedule. **Personal-use copy - "
         "gitignored, never commit or redistribute.**_\n",
         "| Part | Folder | Files |", "|---|---|---|"]
for PART, ps, pg in master:
    n = counts.get(PART, 0)
    lines.append("| %s | %s/ | %d |" % (PART, ps, n))
    idx = ["# %s - index\n" % PART]
    for fn in sorted(os.listdir(os.path.join(OUT, ps))):
        if fn.endswith(".md") and fn != "index.md":
            idx.append("- [%s](%s)" % (fn[4:-3].replace("-", " "), fn))
    with open(os.path.join(OUT, ps, "index.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(idx) + "\n")
lines += ["", "Usage: grep the part folder for a section number or phrase, e.g.",
          '`grep -l "tax invoice" cgst-act-2017/*.md` or open `cgst-act-2017/index.md`.', ""]
with open(os.path.join(OUT, "INDEX.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print("parts:", {k: v for k, v in counts.items()})
print("total files:", sum(counts.values()) + len(master) + 1)
