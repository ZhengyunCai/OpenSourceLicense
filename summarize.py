"""Summarise a ScanCode JSON scan result for the report."""
import json
import sys
from collections import Counter, defaultdict

src = sys.argv[1] if len(sys.argv) > 1 else "scancode-result.json"
with open(src, encoding="utf-8") as f:
    data = json.load(f)

files = data.get("files", [])
print("scancode version:", data.get("scancode_version"), "| headers:", data.get("headers", [{}])[0])
print("total entries:", len(files))
print("scan duration (s):", (data.get("timing") or {}).get("scan_end", 0) - (data.get("timing") or {}).get("scan_start", 0))

scanned = [f for f in files if not f.get("is_filtered")]
print("scanned files:", len(scanned))

with_lic = [f for f in scanned if f.get("license_detections")]
without = [f for f in scanned if not f.get("license_detections")]
print("files with license detection:", len(with_lic))
print("files without license detection:", len(without))

# expression level
expr = Counter()
for f in scanned:
    keys = [d.get("license_expression") for d in f.get("license_detections", [])]
    expr[" AND ".join(sorted(keys)) if keys else "(none)"] += 1
print("\n== expression distribution ==")
for k, v in expr.most_common():
    print(f"{v:5d}  {k}")

# per license key
key_counter = Counter()
key_files = defaultdict(list)
for f in scanned:
    seen = set()
    for d in f.get("license_detections", []):
        for m in d.get("matches", []):
            k = m.get("license_expression") or m.get("license")
            if k and k not in seen:
                seen.add(k)
                key_counter[k] += 1
                key_files[k].append(f["path"])
print("\n== license key file counts ==")
for k, v in key_counter.most_common():
    print(f"{v:5d}  {k}")

print("\n== non apache-2.0 files (sample) ==")
for k, v in key_counter.most_common():
    if k != "apache-2.0":
        print(f"-- {k} ({v})")
        for p in key_files[k][:6]:
            print("    ", p)

print("\n== files without license detection ==")
for f in without:
    print("   ", f["path"], "|", f.get("file_type"))

# copyrights
holders = Counter()
stmts = Counter()
for f in scanned:
    for c in f.get("copyrights", []):
        stmts[c.get("copyright", "").strip()] += 1
        if c.get("copyright_holder"):
            holders[c["copyright_holder"].strip()] += 1
print("\n== top copyright statements ==")
for s, n in stmts.most_common(12):
    print(f"{n:5d}  {s}")
print("\n== copyright holders ==")
for s, n in holders.most_common(12):
    print(f"{n:5d}  {s}")

# packages
print("\n== packages ==")
for f in scanned:
    for p in f.get("package_data", []):
        print("   ", p.get("type"), p.get("name"), p.get("version"), p.get("purl"), "|", f["path"])

# emails/urls
mails = Counter()
for f in scanned:
    for e in f.get("emails", []):
        mails[e.get("email")] += 1
print("\n== emails ==", mails.most_common(8))

urls = Counter()
for f in scanned:
    for u in f.get("urls", []):
        urls[u.get("url")] += 1
print("== urls (top) ==")
for u, n in urls.most_common(10):
    print(f"{n:5d}  {u}")
