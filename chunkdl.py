import os, sys, time, threading, urllib.request, json

URL = "https://mirrors.aliyun.com/pypi/packages/02/bb/abfa88f637b2711e0a7b3bf1bb7bb0ffc49941d4f8fe418f7be1bb47c479/scancode_toolkit-32.5.0-cp313-none-any.whl"
OUT = sys.argv[1] if len(sys.argv) > 1 else "scancode_toolkit-32.5.0-cp313-none-any.whl"
TOTAL = 126229828
CHUNK = 1 << 20  # 1 MB
PARTS = "parts"

os.makedirs(PARTS, exist_ok=True)
nchunks = (TOTAL + CHUNK - 1) // CHUNK
lock = threading.Lock()
done = [0]


def fetch(i):
    path = os.path.join(PARTS, f"{i:04d}")
    start = i * CHUNK
    end = min(start + CHUNK, TOTAL) - 1
    want = end - start + 1
    if os.path.exists(path) and os.path.getsize(path) == want:
        with lock:
            done[0] += 1
        return True
    for attempt in range(40):
        try:
            req = urllib.request.Request(URL, headers={"Range": f"bytes={start}-{end}"})
            with urllib.request.urlopen(req, timeout=45) as r:
                data = r.read()
            if len(data) == want:
                with open(path, "wb") as f:
                    f.write(data)
                with lock:
                    done[0] += 1
                    print(f"{done[0]}/{nchunks} chunk {i}", flush=True)
                return True
        except Exception as e:
            time.sleep(1.5)
    print(f"FAILED chunk {i}", flush=True)
    return False


def worker(q):
    while True:
        with lock:
            if not q:
                return
            i = q.pop(0)
        fetch(i)


q = list(range(nchunks))
threads = [threading.Thread(target=worker, args=(q,), daemon=True) for _ in range(6)]
t0 = time.time()
for t in threads:
    t.start()
for t in threads:
    t.join()

missing = [i for i in range(nchunks) if not (os.path.exists(os.path.join(PARTS, f"{i:04d}")) and os.path.getsize(os.path.join(PARTS, f"{i:04d}")) == min(CHUNK, TOTAL - i * CHUNK))]
if missing:
    print("MISSING:", missing[:20], len(missing))
    sys.exit(1)

with open(OUT, "wb") as out:
    for i in range(nchunks):
        with open(os.path.join(PARTS, f"{i:04d}"), "rb") as f:
            out.write(f.read())
print("DONE", os.path.getsize(OUT), f"{time.time()-t0:.0f}s")
