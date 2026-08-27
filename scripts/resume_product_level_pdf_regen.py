from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[0]))
import regen_product_level_pdfs as R
import time

def stale_jobs():
    out=[]
    for folder in sorted(p for p in R.PRODUCT_SETS.iterdir() if p.is_dir()):
        for stem in R.STEMS:
            docx=folder/f"{stem}.docx"
            pdf=folder/f"{stem}.pdf"
            if not docx.exists():
                continue
            if (not pdf.exists()) or pdf.stat().st_size==0 or pdf.stat().st_mtime < docx.stat().st_mtime - 1:
                out.append((docx,pdf))
    return out

def main():
    jobs=stale_jobs()
    print(f"resume_total={len(jobs)}", flush=True)
    ok=fail=0
    R.kill()
    for i,(docx,pdf) in enumerate(jobs,1):
        try:
            if R.convert_one(docx,pdf):
                ok+=1
            else:
                fail+=1
        except Exception as exc:
            fail+=1
            R.LOG.open("a",encoding="utf-8").write(f"FAIL {docx}: {exc}\n")
            R.kill(); time.sleep(3)
        if i%16==0:
            print(f"progress {i}/{len(jobs)} ok={ok} fail={fail}", flush=True)
            R.kill(); time.sleep(2)
    R.kill()
    print(f"DONE ok={ok} fail={fail}", flush=True)

if __name__=="__main__":
    main()
