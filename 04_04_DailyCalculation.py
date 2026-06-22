# 04_04_DailyCalculation.py
# Mode 0 每日详细输出 -- 每个 Sheet 为一天，每行一个患者
# LAST UPDATE BY LIFANGU IN 20260618

import os, sys, math, yaml, argparse
import pandas as pd, numpy as np
from datetime import datetime, time, timedelta

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

def load_config():
    with open(os.path.join(ROOT_DIR, "config.yaml"), "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

config = load_config()
MATCH_BY        = config.get("match_by", "sensor_id")
DATA_FOLDER     = config.get("data_folder", "")
OUTPUT_FOLDER   = config.get("output_folder", "")
PATIENT_LIST    = config.get("patient_list_file", "")
DATETAG         = config.get("datetag", "")
DURING_DAY      = config.get("duringday", 14)
CALC_GROUPS     = config.get("calc_groups", {i: True for i in range(1, 9)})
MIN_LEN_PER_DAY = 12 * 24

CSV_COLUMN_MAP = {
    "住院号": "hospital_id",
    "入院日期": "admission_time",
    "出院日期": "discharge_time",
    "一次性探头编号": "sensor_id",
    "胰岛素泵开始时间": "pump_start_time",
    "胰岛素泵停止时间": "pump_end_time",
    "德谷开始时间": "degu_start_time",
    "甘精开立时间": "ganjing_start_time",
}

def read_patient_csv(csv_path):
    df = pd.read_csv(csv_path, dtype=object, encoding="utf-8-sig")
    rn = {k: v for k, v in CSV_COLUMN_MAP.items() if k in df.columns}
    df = df.rename(columns=rn)
    for c in ["hospital_id", "sensor_id", "admission_time", "discharge_time",
              "pump_start_time", "pump_end_time", "degu_start_time", "ganjing_start_time"]:
        if c not in df.columns: df[c] = np.nan
    df = df[df["hospital_id"].notna() & (df["hospital_id"] != "")].copy()
    if MATCH_BY == "sensor_id":
        df = df[df["sensor_id"].notna() & (df["sensor_id"] != "")].copy()
    return df

def find_patient_file(device_id, folder_path):
    if pd.isna(device_id): return None
    s = str(int(device_id)) if isinstance(device_id, float) else str(device_id)
    for fn in os.listdir(folder_path):
        if s in fn and fn.endswith((".xls", ".xlsx")):
            return os.path.join(folder_path, fn)
    return None

def load_and_sample(file_path):
    for eng in ("calamine", "openpyxl"):
        try:
            df = pd.read_excel(file_path, header=None, engine=eng)
            break
        except Exception:
            if eng == "openpyxl": raise
    df = df.drop(0).reset_index(drop=True)
    df.columns = ["timestamp", "glucose"]
    df["timestamp"] = pd.to_datetime(df["timestamp"], format="%Y-%m-%d %H:%M:%S")
    df["glucose"] = pd.to_numeric(df["glucose"], errors="coerce")
    nr = (df["glucose"].dropna() <= 0.5).sum() / max(len(df["glucose"].dropna()), 1)
    if nr > 0.7:
        df = df[df["glucose"] > 0.5].copy()
    else:
        df = df.iloc[::5].copy()
    df["glucose"] = df["glucose"].clip(1.8, 33.3)
    return df

# ====================================================================
def calc_basic_stats(g):
    if g.empty: return {}
    m = g.mean(); s = g.std()
    cv = round(s/m, 4) if m != 0 and not np.isnan(m) else None
    gmi = round(3.31 + 0.02392 * 18 * m, 4) if not np.isnan(m) else None
    return {"MEAN": round(m,4) if not np.isnan(m) else None,
            "SD": round(s,4) if not np.isnan(s) else None, "CV": cv, "GMI": gmi}

def calc_lbgi_hbgi_adrr(df):
    if df.empty: return None, None, None
    df = df[df["glucose"] >= 1.0].copy()
    if df.empty: return None, None, None
    f = 1.794 * (np.log(df["glucose"]) ** 1.026 - 1.861)
    r = 10 * (f ** 2); rl = np.where(f < 0, r, 0); rh = np.where(f > 0, r, 0)
    return round(float(rl.mean()),4), round(float(rh.mean()),4), round(float(rl.max()+rh.max()),4)

def calc_daily_modd(curr, prev):
    if prev is None or curr.empty or prev.empty: return None
    common = set(pd.unique(curr["timestamp"].dt.time)) & set(pd.unique(prev["timestamp"].dt.time))
    diffs = []
    for ct in common:
        cv = curr[curr["timestamp"].dt.time == ct]["glucose"].values
        pv = prev[prev["timestamp"].dt.time == ct]["glucose"].values
        if len(cv)>0 and len(pv)>0: diffs.append(abs(cv[0]-pv[0]))
    return round(pd.Series(diffs).mean(),4) if diffs else None

def calc_mage_daily(g):
    d = g.dropna().values
    if len(d)<3: return None
    sd = np.std(d, ddof=1)
    if sd==0: return 0.0
    pk, nd = [], []
    for i in range(1, len(d)-1):
        if d[i]>d[i-1] and d[i]>d[i+1]: pk.append((i,d[i]))
        elif d[i]<d[i-1] and d[i]<d[i+1]: nd.append((i,d[i]))
    if not pk or not nd: return None
    tp = sorted(pk+nd, key=lambda x: x[0])
    fd, s, c = None, 0, 0
    for i in range(1, len(tp)):
        diff = tp[i][1]-tp[i-1][1]; amp = abs(diff)
        if amp > sd:
            dr = 1 if diff>0 else -1
            if fd is None: fd=dr; s+=amp; c+=1
            elif dr==fd: s+=amp; c+=1
    return s/c if c>0 else None

def calc_range_stats(ddx, p=""):
    if ddx.empty: return {}
    t = len(ddx); g = ddx["glucose"].values
    def cr(c): return round(c/t, 4) if t>0 else 0
    r = {}
    r["TIR"+p] = cr(np.sum((g>=3.9)&(g<=10.0)))
    r["TAR"+p] = cr(np.sum(g>10.0))
    r["TBR"+p] = cr(np.sum(g<3.9))
    r["TAR1"+p] = cr(np.sum((g>10.0)&(g<=13.9)))
    r["TAR2"+p] = cr(np.sum(g>13.9))
    r["TBR1"+p] = cr(np.sum((g>=3.0)&(g<3.9)))
    r["TBR2"+p] = cr(np.sum(g<3.0))
    r["TITR"+p] = cr(np.sum((g>=3.9)&(g<=7.8)))
    r["GRI"+p] = round(3.0*r["TBR2"+p]+2.4*r["TBR1"+p]+1.6*r["TAR2"+p]+0.8*r["TAR1"+p],4)
    r["TIR-TITR"+p] = round(r["TIR"+p]-r["TITR"+p],4)
    return r

def find_simple_events(df, th, cf, min_d=15):
    if df.empty: return []
    ts = df["timestamp"].tolist(); vs = df["glucose"].tolist()
    ev, ie, st, cet = [], False, None, []
    for i in range(len(vs)):
        if cf(vs[i], th):
            if not ie: ie=True; st=ts[i]; cet=[ts[i]]
            else:
                if (ts[i]-cet[-1]).total_seconds()>15*60:
                    dur=(cet[-1]-st).total_seconds()/60
                    if dur>=min_d: ev.append((st,cet[-1]))
                    st=ts[i]; cet=[ts[i]]
                else: cet.append(ts[i])
        else:
            if ie:
                dur=(cet[-1]-st).total_seconds()/60
                if dur>=min_d: ev.append((st,cet[-1]))
                ie=False; cet=[]
    if ie:
        dur=(cet[-1]-st).total_seconds()/60
        if dur>=min_d: ev.append((st,cet[-1]))
    return ev

def calc_event_stats(ddx):
    if ddx.empty: return {}
    ddx = ddx.sort_values("timestamp").copy()
    def fe(el):
        if not el: return 0, None
        return len(el), ",".join(s.strftime("%Y-%m-%d %H:%M:%S")+"~"+e.strftime("%Y-%m-%d %H:%M:%S") for s,e in el)
    he = find_simple_events(ddx, 3.9, lambda x,t: x<t, 15)
    s = {}
    s["HYPO"], s["Time-HYPO"] = fe(he)
    s["HYPO 0TO6AM"], s["Time-HYPO 0TO6AM"] = fe([e for e in he if 0<=e[0].hour<6])
    def ehr(vs, ts, idx):
        if idx>=len(vs): return False,0
        st=ts[idx]; j=idx
        while j<len(vs):
            if vs[j]<3.9: return False,0
            if (ts[j]-st).total_seconds()/60>=15: return True,0
            j+=1
        return False,0
    exh = []
    if not ddx.empty:
        ts=ddx["timestamp"].tolist(); vs=ddx["glucose"].tolist()
        i=0; n=len(vs)
        while i<n:
            if vs[i]<3.9:
                si=i; st=ts[i]; ei=-1; j=si+1
                while j<n:
                    t,_=ehr(vs,ts,j)
                    if t: ei=j; break
                    if j>si and (ts[j]-ts[j-1]).total_seconds()>30*60: ei=j; break
                    j+=1
                if ei==-1: ei=n
                ee=ts[ei] if ei<n else ts[n-1]
                dur=(ee-st).total_seconds()/60
                if dur>=120: exh.append((st,ee))
                i=ei
            else: i+=1
    s["EX HYPO"], s["Time-EX HYPO"] = fe(exh)
    s["EX HYPO 0TO6AM"], s["Time-EX HYPO 0TO6AM"] = fe([e for e in exh if 0<=e[0].hour<6])
    return s

def time_period_stats(ddx, sh, eh):
    st = datetime.strptime(f"{sh:02d}:00:00","%H:%M:%S").time()
    et = datetime.strptime(f"{eh:02d}:59:59","%H:%M:%S").time()
    pdf = ddx[(ddx["timestamp"].dt.time>=st)&(ddx["timestamp"].dt.time<=et)].copy()
    if pdf.empty:
        return {"mean":None,"std":None,"cv":None,"vv_list":None,"vv_time_list":None}
    g = pdf["glucose"].dropna()
    if g.empty:
        return {"mean":None,"std":None,"cv":None,"vv_list":None,"vv_time_list":None}
    m = g.mean(); s = g.std(); cv = s/m if m!=0 and not np.isnan(m) else None
    pdf["date"] = pdf["timestamp"].dt.date
    dms=[]; dmt=[]
    for _, grp in pdf.groupby("date"):
        vg = grp.dropna(subset=["glucose"])
        if vg.empty: continue
        mv = vg["glucose"].min(); dms.append(round(mv,4))
        dmt.append(vg[vg["glucose"]==mv]["timestamp"].iloc[0].strftime("%H:%M:%S"))
    return {"mean":round(m,4) if not np.isnan(m) else None,
            "std":round(s,4) if not np.isnan(s) else None,
            "cv":round(cv,4) if cv and not np.isnan(cv) else None,
            "vv_list":",".join(map(str,dms)) if dms else None,
            "vv_time_list":",".join(dmt) if dmt else None}

# ====================================================================
def process_one(file_path, admission_time=None, discharge_time=None, during_day=None):
    """处理单个患者文件, during_day 优先于模块级 DURING_DAY"""
    if during_day is None:
        during_day = DURING_DAY  # 兼容直接调用的场景
    try:
        df = load_and_sample(file_path)
    except Exception as e:
        print("  load error: " + str(e))
        return None
    if df.empty:
        return None
    data_start = df["timestamp"].min()
    start_date = data_start.normalize()
    start_time = start_date
    end_time = start_time + timedelta(days=during_day)
    df = df[(df["timestamp"]>=start_time)&(df["timestamp"]<end_time)]
    if df.empty: return None
    global_min = df["glucose"].min()
    daily = []; prev = None
    for d in range(during_day):
        dd = df[(df["timestamp"]>=start_time+timedelta(days=d))&
                (df["timestamp"]<start_time+timedelta(days=d+1))].copy()
        if dd.empty or len(dd)<MIN_LEN_PER_DAY//2:
            daily.append(None); prev=dd; continue
        sts = {}
        if CALC_GROUPS.get(1,True): sts.update(calc_basic_stats(dd["glucose"]))
        if CALC_GROUPS.get(2,True):
            l,h,a=calc_lbgi_hbgi_adrr(dd); sts["LBGI"]=l; sts["HBGI"]=h; sts["ADRR"]=a
            sts["MODD"]=calc_daily_modd(dd,prev)
        if CALC_GROUPS.get(3,True):
            gs=dd["glucose"].dropna()
            if len(gs)>=144:
                sts["LAGE"]=round(gs.max()-gs.min(),4)
                m=calc_mage_daily(gs); sts["MAGE"]=round(m,4) if m else None
            else: sts["LAGE"]=None; sts["MAGE"]=None
        if CALC_GROUPS.get(4,True):
            sts.update(calc_range_stats(dd,""))
            sts.update(calc_range_stats(dd[dd["timestamp"].dt.hour<6],"-0TO6AM"))
            sts.update(calc_range_stats(dd[dd["timestamp"].dt.hour>=6],"-6AMTO0"))
        if CALC_GROUPS.get(5,True):
            p1=time_period_stats(dd,0,5); p2=time_period_stats(dd,6,23)
            sts["MEAN-0TO6AM"]=p1["mean"]; sts["SD-0TO6AM"]=p1["std"]
            sts["CV-0TO6AM"]=p1["cv"]; sts["VV-0TO6AM"]=p1["vv_list"]
            sts["VVtime-0TO6AM"]=p1["vv_time_list"]
            sts["MEAN-6AMTO0"]=p2["mean"]; sts["SD-6AMTO0"]=p2["std"]
            sts["CV-6AMTO0"]=p2["cv"]
        if CALC_GROUPS.get(6,True): sts.update(calc_event_stats(dd))
        if CALC_GROUPS.get(7,True):
            l2=find_simple_events(dd,3.0,lambda x,t:x<t,15)
            def fe(el):
                if not el: return 0,None
                return len(el),",".join(a.strftime("%Y-%m-%d %H:%M:%S")+"~"+b.strftime("%Y-%m-%d %H:%M:%S") for a,b in el)
            sts["LV2 HYPO"],sts["Time-LV2 HYPO"]=fe(l2)
            sts["LV2 HYPO 0TO6AM"],sts["Time-LV2 HYPO 0TO6AM"]=fe([e for e in l2 if 0<=e[0].hour<6])
        if CALC_GROUPS.get(8,True):
            l2=find_simple_events(dd,3.0,lambda x,t:x<t,15)
            for cond, ns, nn in [(3.0,False,""),(3.0,True," 0TO6AM"),(3.5,False,""),(3.5,True," 0TO6AM")]:
                if global_min is not None and global_min>=cond:
                    sts["HYPO_COND_"+str(cond)+nn]="#N/A"
                    sts["Time-HYPO_COND_"+str(cond)+nn]="#N/A"
                else:
                    ev=[e for e in l2 if 0<=e[0].hour<6] if ns else l2; c,t=fe(ev)
                    sts["HYPO_COND_"+str(cond)+nn]=c; sts["Time-HYPO_COND_"+str(cond)+nn]=t
        daily.append(sts); prev=dd
    return {"info":{"patient_id":os.path.splitext(os.path.basename(file_path))[0],
            "start_time":start_time,"end_time":end_time,
            "admission_time":admission_time,"discharge_time":discharge_time},"days":daily}

def main():
    # ---- CLI 参数解析, 覆盖 config.yaml 中的对应值 ----
    parser = argparse.ArgumentParser(description="04_04 Mode 0 Daily Calculation")
    parser.add_argument("--duringday", type=int, default=None,
                        help="覆盖 config 中的 duringday (处理天数)")
    parser.add_argument("--nametag", type=str, default=None,
                        help="覆盖 config 中的 nametag (输出文件标识)")
    cli = parser.parse_args()

    during_day = cli.duringday if cli.duringday is not None else DURING_DAY
    nametag    = cli.nametag    if cli.nametag    is not None else config.get("nametag", "Daily_D1-"+str(during_day))

    print("04_04 DailyCalculation -- Mode 0, " + str(during_day) + " days", flush=True)
    patient_df = read_patient_csv(PATIENT_LIST)
    for c in ["admission_time","discharge_time"]:
        if c in patient_df.columns: patient_df[c] = pd.to_datetime(patient_df[c], errors="coerce")
    all_data = []
    for _, row in patient_df.iterrows():
        hid = row.get("hospital_id"); adm = row.get("admission_time"); dis = row.get("discharge_time")
        mk = row.get(MATCH_BY) if MATCH_BY in patient_df.columns else hid
        if pd.isna(mk): continue
        pf = find_patient_file(mk, DATA_FOLDER)
        if pf:
            print("  " + str(hid) + " -> " + os.path.basename(pf), flush=True)
            r = process_one(pf, adm, dis, during_day)
            if r: r["info"]["hospital_id"] = hid; all_data.append(r)
        else:
            print("  " + str(hid) + ": file not found", flush=True)
    if not all_data:
        print("No results.")
        return
    out = os.path.join(OUTPUT_FOLDER, "CGM_"+DATETAG+"_"+nametag+"_Mode0_Results.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as wr:
        for di in range(during_day):
            rows = []
            for p in all_data:
                rd = p["info"].copy()
                for k in ["start_time","end_time","admission_time","discharge_time"]:
                    if k in rd and isinstance(rd[k], (datetime, pd.Timestamp)) and pd.notna(rd[k]):
                        rd[k] = rd[k].strftime("%Y-%m-%d %H:%M:%S")
                if di < len(p["days"]) and p["days"][di]: rd.update(p["days"][di])
                rows.append(rd)
            if rows:
                dfd = pd.DataFrame(rows)
                pr = ["hospital_id","patient_id","admission_time","discharge_time","start_time","end_time"]
                fc = [c for c in pr if c in dfd.columns] + [c for c in dfd.columns if c not in pr]
                dfd[fc].to_excel(wr, sheet_name="Day "+str(di+1), index=False)
    print("Done: " + out)

if __name__ == "__main__":
    main()
