# 04_02_BatchRunner.py
# 批量并行计算 -- 13 个任务覆盖全部 8 种模式，多进程并发执行
# LAST UPDATE BY LIFANGU IN 20260618

import subprocess, sys, os, time
from concurrent.futures import ProcessPoolExecutor, as_completed

SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "04_03_Calculation.py")
MAX_WORKERS = 4

TASKS = [
    # dict(mode=0, duringday=3,  nametag="Day1-3"),
    # dict(mode=0, duringday=5,  nametag="Day1-5"),
    # dict(mode=0, duringday=7,  nametag="Day1-7"),
    # dict(mode=0, duringday=14, nametag="Day1-14"),
    # dict(mode=1, interimday=5,  duringday=11, nametag="Day6-11"),
    # dict(mode=1, interimday=6,  duringday=14, nametag="Day7-14"),
    # dict(mode=1, interimday=8,  duringday=14, nametag="Day9-14"),
    # dict(mode=2, duringday=0, nametag="AllDays"),
    # dict(mode=3, duringday=3, nametag="3DaysBeforeDischarge"),
    # dict(mode=4, duringday=0, nametag="DuringHospitalization"),
    # dict(mode=5, duringday=7, nametag="7DaysAfterDischarge"),
    # dict(mode=6, duringday=0, nametag="AllDaysAfterDischarge"),
    dict(mode=7, duringday=0, nametag="InsulinPump"),
]

def fmt_sec(s):
    if s < 60: return str(int(s)) + "s"
    return str(int(s // 60)) + "m" + str(int(s % 60)) + "s"

def run_one(task):
    tag = task["nametag"]
    cmd = [sys.executable, SCRIPT, "--mode", str(task["mode"]),
           "--during", str(task["duringday"]), "--nametag", tag]
    if task.get("interimday"):
        cmd += ["--interim", str(task["interimday"])]
    t0 = time.time()
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=os.path.dirname(SCRIPT))
    elapsed = time.time() - t0
    lines = [l for l in r.stdout.strip().split(chr(10)) if l.strip()]
    summary = lines[-1] if lines else "(no output)"
    error = r.stderr.strip()[-200:] if r.stderr.strip() else None
    return dict(tag=tag, mode=task["mode"], ok=r.returncode == 0,
                elapsed=elapsed, summary=summary, error=error)

def main():
    n = len(TASKS)
    print("=" * 60)
    print("  CGM Batch Runner")
    print("  脚本: " + os.path.basename(SCRIPT))
    print("  任务: " + str(n) + " | 并行: " + str(MAX_WORKERS) + " 进程")
    print("  开始: " + time.strftime("%Y-%m-%d %H:%M:%S"))
    print("=" * 60)
    print("")

    t_total = time.time()
    done, fail = 0, 0
    pending = dict(enumerate(TASKS))

    with ProcessPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {}
        # 提交时打印
        for i, t in enumerate(TASKS):
            f = pool.submit(run_one, t)
            futures[f] = (i, t)
            detail = ""
            if t.get("interimday"):
                detail = " int=" + str(t["interimday"])
            print("  [" + str(i + 1).rjust(2) + "/" + str(n) + "] submitted  "
                  + t["nametag"].ljust(24) + " M" + str(t["mode"]) + " dur=" + str(t["duringday"]) + detail, flush=True)

        print("")
        completed_order = 0
        for f in as_completed(futures):
            idx, task = futures[f]
            r = f.result()
            completed_order += 1

            st = "OK" if r["ok"] else "FAIL"
            elapsed_str = fmt_sec(r["elapsed"])
            remaining = len(futures) - completed_order
            if completed_order > 0:
                avg = (time.time() - t_total) / completed_order
                eta = fmt_sec(avg * remaining) if remaining > 0 else "0s"

            tag_padded = r["tag"].ljust(24)
            progress = "[" + str(completed_order) + "/" + str(n) + "]"
            msg = "  " + progress + " [" + st + "] " + tag_padded + " M" + str(r["mode"]) + "  " + elapsed_str
            if remaining > 0:
                msg += "  ETA " + eta

            # 打印子进程的关键输出行（患者/文件处理摘要）
            key_lines = [l for l in r["summary"].split(chr(10)) if l.strip() and ("处理" in l or "完成" in l or "保存" in l)]
            if key_lines:
                msg += "  |  " + key_lines[-1]

            print(msg, flush=True)

            if r["ok"]:
                done += 1
            else:
                fail += 1
                if r["error"]:
                    for el in r["error"].split(chr(10))[-3:]:
                        if el.strip():
                            print("         [ERR] " + el.strip())

    total = time.time() - t_total
    print("")
    print("=" * 60)
    print("  完成: " + str(done) + "/" + str(n) + " OK, " + str(fail) + " fail")
    print("  总耗时: " + fmt_sec(total))
    print("  结束: " + time.strftime("%Y-%m-%d %H:%M:%S"))
    print("=" * 60)

if __name__ == "__main__":
    main()
