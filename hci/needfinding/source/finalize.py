"""Name the recovered clusters, compute segment profiles and overall statistics used by the workbook,
README and vis.html narrative. Writes out/responses_final.csv and out/analysis.json"""
import json, numpy as np, pandas as pd

Q = json.load(open("questions.json"))
QI = {q["id"]: q for s in Q["sections"] for q in s["questions"]}
d = pd.read_csv("out/responses_clustered_raw.csv")
diag = json.load(open("out/cluster_diag.json"))

NAMES = {"sprinter": ("Deadline Sprinters", "Phone-first students who open Canvas several times a day for a few seconds: deadlines, grades, room changes."),
         "scholar": ("Structured Scholars", "Master/PhD students on a laptop, long sessions, live in Modules and recordings, want consistency and search."),
         "juggler": ("Overloaded Jugglers", "Students with 17+ hours of work or care per week; check 2-3 times a week, run on e-mail, and catch up in batches."),
         "power": ("Power Instructors", "Lecturers/TAs in Canvas all day: grading, announcements, groups; want bulk actions and fewer clicks."),
         "reluctant": ("Reluctant Instructors", "Senior lecturers who open Canvas about once a week, low confidence, delegate to TAs and e-mail."),
         "access": ("Access-First Learners", "Students using screen readers, captions or reading support; blocked by unstructured pages and media.")}

dom = {int(k): v for k, v in diag["dominant"].items()}
sizes = d["cluster_raw"].value_counts()
ordered = sorted(dom.keys(), key=lambda c: -sizes[c])
seg_id = {c: f"S{i+1}" for i, c in enumerate(ordered)}
d["Segment"] = d["cluster_raw"].map(lambda c: f"{seg_id[c]} {NAMES[dom[c]][0]}")
d["SegmentKey"] = d["cluster_raw"].map(lambda c: dom[c])
d = d.drop(columns=["cluster_raw"])

def opts(qid): return QI[qid]["options"]
def label(qid, k): return next(o["label"] for o in opts(qid) if o["k"] == str(k))

def pct(series): return float((series.mean() * 100).round(1))
def dist(col, qid):
    return {o["label"]: int((d[col] == o["k"]).sum()) for o in opts(qid)}
def multi_pct(prefix, qid, frame):
    return {o["label"]: round(float(frame[f"{prefix}__{o['k']}"].mean() * 100), 1) for o in opts(qid)}

likert = {"C1_find_fast": "I can quickly find the course material I need.", "C2_find_slow": "Finding what I need takes longer than it should.",
          "C3_notif_useful": "Notifications help me stay on top of what is due.", "C4_notif_noise": "I receive notifications that are not relevant to me.",
          "C5_mobile": "Using the system on my phone works as well as on a computer.", "C6_consistent": "Course pages are organised in a similar way across my courses.",
          "C7_uptodate": "I trust that the grades and feedback I see are up to date.", "C8_confident": "I can complete my tasks without asking anyone for help.",
          "C9_accessible": "Text, colours and videos are easy for me to read, see and hear.", "C10_group": "Working with others is well supported.",
          "C11_satisfaction": "Overall satisfaction (1-7)", "C12_ease": "Ease of last task (1-5)"}

def profile(frame):
    p = {"n": int(len(frame)),
         "roles": frame["A1_role"].map(lambda k: label("A1", k)).value_counts().to_dict(),
         "age": frame["A2_age"].map(lambda k: label("A2", k)).value_counts().to_dict(),
         "gender": frame["A3_gender"].map(lambda k: label("A3", k)).value_counts().to_dict(),
         "origin": frame["A5_origin"].map(lambda k: label("A5", k)).value_counts().to_dict(),
         "hours": frame["A7_hours"].map(lambda k: label("A7", k)).value_counts().to_dict(),
         "freq": frame["B1_freq"].map(lambda k: label("B1", k)).value_counts().to_dict(),
         "primary": frame["B3_primary"].map(lambda k: label("B3", k)).value_counts().to_dict(),
         "session": frame["B4_session"].map(lambda k: label("B4", k)).value_counts().to_dict(),
         "notif": frame["B7_notif"].map(lambda k: label("B7", k)).value_counts().to_dict(),
         "lost": frame["B8_lost"].map(lambda k: label("B8", k)).value_counts().to_dict(),
         "access": multi_pct("A8_access", "A8", frame),
         "devices": multi_pct("B2_devices", "B2", frame),
         "features": multi_pct("B5_features", "B5", frame),
         "deadline": multi_pct("B6_deadline", "B6", frame),
         "tools": multi_pct("B9_tools", "B9", frame),
         "likert_mean": {k: round(float(frame[k].mean()), 2) for k in likert},
         "likert_sd": {k: round(float(frame[k].std(ddof=1)), 2) for k in likert},
         "numeric_mean": {k: round(float(frame[k].mean()), 2) for k in ["freq_v", "session_v", "phone_primary", "A9_tech", "hours_v", "lost_v", "courses_v"]},
         "rank_mean": {o["label"]: round(float(frame[f"D1_rank__{o['k']}"].mean()), 2) for o in opts("D1")},
         "rank_first_pct": {o["label"]: round(float((frame[f"D1_rank__{o['k']}"] == 1).mean() * 100), 1) for o in opts("D1")},
         "pct_several_or_daily": pct(frame["B1_freq"].isin(["several", "daily"])),
         "pct_phone_primary": pct(frame["B3_primary"] == "phone"),
         "pct_short_session": pct(frame["B4_session"].isin(["lt2", "2-5"])),
         "pct_missed_deadline": pct(frame["B6_deadline__late"] == 1),
         "pct_work17plus": pct(frame["A7_hours"].isin(["17-24", "25+"])),
         "pct_lost_2plus": pct(frame["B8_lost"].isin(["2-3", "4-6", "7+"])),
         "pct_satisfied": pct(frame["C11_satisfaction"] >= 5), "pct_dissatisfied": pct(frame["C11_satisfaction"] <= 3),
         }
    return p

overall = profile(d)
segments = {}
for seg, frame in d.groupby("Segment"):
    p = profile(frame); p["key"] = frame["SegmentKey"].iloc[0]; p["tagline"] = NAMES[p["key"]][1]
    p["quotes"] = frame["E1_frustrate"].tolist()[:6]
    p["quotes_helped"] = [t for t in frame["E2_helped"].fillna("").tolist() if t][:4]
    p["quotes_more"] = [t for t in frame["E3_more"].fillna("").tolist() if t][:4]
    segments[seg] = p

# correlation matrix of Likert + satisfaction
lc = list(likert.keys())
corr = d[lc].corr().round(2)
# crosstabs
ct_role_freq = pd.crosstab(d["A1_role"].map(lambda k: label("A1", k)), d["B1_freq"].map(lambda k: label("B1", k)))
sat_by_role = d.groupby(d["A1_role"].map(lambda k: label("A1", k)))["C11_satisfaction"].agg(["mean", "count"]).round(2)
sat_by_freq = d.groupby(d["B1_freq"].map(lambda k: label("B1", k)))["C11_satisfaction"].agg(["mean", "count"]).round(2)
sat_by_primary = d.groupby(d["B3_primary"].map(lambda k: label("B3", k)))["C11_satisfaction"].agg(["mean", "count"]).round(2)
mobile_by_primary = d.groupby(d["B3_primary"].map(lambda k: label("B3", k)))["C5_mobile"].mean().round(2)

analysis = {"n": int(len(d)), "overall": overall, "segments": segments, "likert_text": likert,
            "alpha_pair": round(diag["alpha_pair"], 3), "r_pair": round(diag["r_pair"], 3),
            "k_selection": {"k": diag["ks"], "inertia": [round(x, 1) for x in diag["inertia"]], "silhouette": [round(x, 3) for x in diag["silhouette"]]},
            "pca_explained": [round(x, 3) for x in diag["explained"]],
            "corr": {a: {b: float(corr.loc[a, b]) for b in lc} for a in lc},
            "ct_role_freq": {r: {c: int(v) for c, v in row.items()} for r, row in ct_role_freq.iterrows()},
            "sat_by_role": {r: {"mean": float(v["mean"]), "n": int(v["count"])} for r, v in sat_by_role.iterrows()},
            "sat_by_freq": {r: {"mean": float(v["mean"]), "n": int(v["count"])} for r, v in sat_by_freq.iterrows()},
            "sat_by_primary": {r: {"mean": float(v["mean"]), "n": int(v["count"])} for r, v in sat_by_primary.iterrows()},
            "mobile_by_primary": {r: float(v) for r, v in mobile_by_primary.items()},
            "segment_order": sorted(segments.keys())}
json.dump(analysis, open("out/analysis.json", "w"), indent=1)

# final shipped CSV: survey answers + analysis-derived columns at the end
keep = [c for c in d.columns if not (c.endswith("_v") or c in ("phone_primary", "C2r_find_slow_rev", "SegmentKey"))]
d[keep].to_csv("out/responses_final.csv", index=False)
print(json.dumps({s: (p["n"], p["roles"]) for s, p in segments.items()}, indent=1))
print("alpha", analysis["alpha_pair"], "r", analysis["r_pair"])
for s, p in segments.items():
    print(s, "| freq", p["numeric_mean"]["freq_v"], "| phone%", p["pct_phone_primary"], "| sat", p["likert_mean"]["C11_satisfaction"],
          "| find", p["likert_mean"]["C1_find_fast"], "| consist", p["likert_mean"]["C6_consistent"], "| conf", p["likert_mean"]["C8_confident"],
          "| access", p["likert_mean"]["C9_accessible"], "| work17+", p["pct_work17plus"], "| missed", p["pct_missed_deadline"])
