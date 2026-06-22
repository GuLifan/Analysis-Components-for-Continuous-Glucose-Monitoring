# 04_08_BatchRunner2.py
# 批量运行 04_04 ~ 04_07 专项分析
# 通过行级字符串替换覆盖 config，不改变原始文件结构和类型
# LAST UPDATE BY LIFANGU IN 20260618

import subprocess, sys, os, time, shutil

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(ROOT, "config.yaml")
CONFIG_BAK = os.path.join(ROOT, "_config_backup.yaml")
PYTHON = sys.executable

TASKS = [
    # ---- 04_04: 每日详细输出 (Mode 0), CLI 参数直接覆盖 config ----
    ("04_04_DailyCalculation.py",     [("duringday", "3"),  ("nametag", "Daily_D1-3")],  "--duringday 3 --nametag Daily_D1-3"),
    ("04_04_DailyCalculation.py",     [("duringday", "14"), ("nametag", "Daily_D1-14")], "--duringday 14 --nametag Daily_D1-14"),
    # ---- 04_05~04_07: 专项分析, 仅覆盖 nametag ----
    ("04_05_FBSExtractor.py",         [("nametag", "FBS")]),
    ("04_06_CGMDaysExtractor.py",     [("nametag", "Days")]),
    ("04_07_FBSExtractor_24H.py",     [("nametag", "FBS_24H")]),
]

def fmt_sec(s):
    if s < 60: return str(int(s)) + "s"
    return str(int(s // 60)) + "m" + str(int(s % 60)) + "s"

def write_config(overrides):
    """行级替换 -- 只修改匹配 key: 开头的行"""
    with open(CONFIG_BAK, "r", encoding="utf-8") as f:
        orig = f.readlines()
    for key, val in overrides:
        for i in range(len(orig)):
            s = orig[i].lstrip()
            if s.startswith(key + ":") or s.startswith(key + " "):
                ind = len(orig[i]) - len(orig[i].lstrip())
                orig[i] = " " * ind + key + ": " + val + chr(10)
                break
    with open(CONFIG, "w", encoding="utf-8") as f:
        f.writelines(orig)

def main():
    shutil.copy(CONFIG, CONFIG_BAK)
    n = len(TASKS)
    print("=" * 60)
    print("  Batch Runner 2 -- 专项分析 (04_04 ~ 04_07)")
    print("  任务: " + str(n) + " | 开始: " + time.strftime("%Y-%m-%d %H:%M:%S"))
    print("=" * 60)
    print("")
    t_total = time.time()
    done, fail = 0, 0
    for i, item in enumerate(TASKS):
        script     = item[0]
        overrides  = item[1]
        extra_args = item[2].split() if len(item) > 2 else []  # 可选的 CLI 参数
        tag = [v for k, v in overrides if k == "nametag"]
        label = script.replace(".py", "") + " " + (tag[0] if tag else "")
        print("  [" + str(i + 1).rjust(2) + "/" + str(n) + "] " + label.ljust(44), end=" ", flush=True)
        write_config(overrides)
        t0 = time.time()
        cmd = [PYTHON, os.path.join(ROOT, script)] + extra_args
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=ROOT)
        shutil.copy(CONFIG_BAK, CONFIG)
        elapsed = time.time() - t0
        st = "OK" if r.returncode == 0 else "FAIL"
        print("[" + st + "] " + fmt_sec(elapsed), flush=True)
        if r.returncode == 0: done += 1
        else: fail += 1
        for line in (r.stdout + r.stderr).strip().split(chr(10))[-3:]:
            if line.strip():
                print("         " + line.strip())
    if os.path.exists(CONFIG_BAK):
        shutil.copy(CONFIG_BAK, CONFIG)
        os.remove(CONFIG_BAK)
    print("")
    print("=" * 60)
    print("  完成: " + str(done) + "/" + str(n) + " OK, " + str(fail) + " fail")
    print("  总耗时: " + fmt_sec(time.time() - t_total))
    print("=" * 60)

if __name__ == "__main__":
    main()
