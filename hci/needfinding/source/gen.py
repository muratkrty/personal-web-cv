"""Simulate iCanvas needfinding survey respondents from latent behavioural segments.

Outputs (in ./out):
  responses.csv   wide table, one row per respondent (survey answers only)
  latent.csv      hidden generating segment per respondent (for the author's own checks, not shipped)
"""
import json, random, csv, os
import numpy as np
from texts import FRUSTRATE, HELPED, MORE, COURSES, TOPICS

SEED = 2026
rng = np.random.default_rng(SEED)
random.seed(SEED)

Q = json.load(open("questions.json"))
QI = {q["id"]: q for s in Q["sections"] for q in s["questions"]}

def pick(d):
    ks = list(d.keys()); ps = np.array([d[k] for k in ks], float); ps /= ps.sum()
    return ks[rng.choice(len(ks), p=ps)]

def likert(mean, sd=1.0, lo=1, hi=7):
    return int(np.clip(np.round(rng.normal(mean, sd)), lo, hi))

def multi(probs, at_least_one=True, exclusive_none=None):
    chosen = [k for k, p in probs.items() if rng.random() < p]
    if exclusive_none and exclusive_none in chosen and len(chosen) > 1:
        chosen.remove(exclusive_none)
    if at_least_one and not chosen:
        ks = list(probs.keys()); ps = np.array(list(probs.values()), float); ps /= ps.sum()
        chosen = [ks[rng.choice(len(ks), p=ps)]]
    return chosen

# ---------------------------------------------------------------------------
# Latent segments. Numbers are means (Likert) or categorical probabilities.
# ---------------------------------------------------------------------------
SEG = {
 "sprinter": dict(n=34,
   role={"bsc": .78, "msc": .22}, age={"17-20": .45, "21-24": .48, "25-29": .05, "na": .02},
   origin={"nl": .6, "eu": .22, "noneu": .15, "na": .03}, hours={"0": .2, "1-8": .35, "9-16": .35, "17-24": .1},
   tech=5.4, freq={"daily": .3, "several": .62, "4-6": .08},
   devices={"laptop": .85, "desktop": .05, "phone_browser": .35, "phone_app": .85, "tablet": .12},
   primary={"phone": .72, "laptop": .27, "tablet": .01}, session={"lt2": .42, "2-5": .43, "5-15": .15},
   features={"announcements": .9, "assignments": .85, "grades": .85, "modules": .6, "calendar": .7, "inbox": .25,
             "discussions": .15, "quizzes": .35, "recordings": .35, "groups": .2, "syllabus": .1},
   deadline={"canvas_cal": .45, "email": .2, "push": .6, "in_class": .35, "peer": .55, "syllabus": .1, "own_cal": .15, "late": .12},
   notif={"push": .68, "email": .17, "inapp": .1, "off": .05}, lost={"0": .25, "1": .35, "2-3": .3, "4-6": .1},
   tools={"paper": .15, "calendar": .55, "notes": .2, "chat": .9, "email": .2, "none": .05},
   L=dict(find_fast=4.2, notif_useful=5.0, notif_noise=5.4, mobile=3.0, consistent=3.4, uptodate=4.0, confident=5.6, accessible=5.6, group=4.0),
   sat=4.3, ease=3.6, rank=["deadlines", "grades", "materials", "communication", "recordings", "groupwork"]),
 "scholar": dict(n=28,
   role={"msc": .68, "phd": .22, "bsc": .10}, age={"21-24": .5, "25-29": .4, "30-39": .08, "na": .02},
   origin={"nl": .4, "eu": .3, "noneu": .27, "na": .03}, hours={"0": .3, "1-8": .4, "9-16": .25, "17-24": .05},
   tech=6.0, freq={"daily": .5, "4-6": .35, "several": .15},
   devices={"laptop": .97, "desktop": .2, "phone_browser": .3, "phone_app": .35, "tablet": .25},
   primary={"laptop": .9, "desktop": .07, "tablet": .03}, session={"5-15": .3, "15-30": .5, "gt30": .2},
   features={"announcements": .8, "assignments": .75, "grades": .55, "modules": .95, "calendar": .45, "inbox": .3,
             "discussions": .45, "quizzes": .3, "recordings": .75, "groups": .35, "syllabus": .45},
   deadline={"canvas_cal": .5, "email": .3, "push": .1, "in_class": .3, "peer": .2, "syllabus": .45, "own_cal": .55, "late": .05},
   notif={"email": .7, "push": .1, "inapp": .17, "off": .03}, lost={"0": .2, "1": .3, "2-3": .35, "4-6": .12, "7+": .03},
   tools={"paper": .2, "calendar": .8, "notes": .7, "chat": .5, "email": .3, "none": .03},
   L=dict(find_fast=3.4, notif_useful=4.0, notif_noise=4.5, mobile=3.6, consistent=2.4, uptodate=4.6, confident=6.0, accessible=5.6, group=3.8),
   sat=4.0, ease=3.7, rank=["materials", "recordings", "deadlines", "grades", "communication", "groupwork"]),
 "juggler": dict(n=20,
   role={"bsc": .55, "msc": .45}, age={"21-24": .35, "25-29": .35, "30-39": .25, "40-49": .05},
   origin={"nl": .65, "eu": .15, "noneu": .15, "na": .05}, hours={"17-24": .45, "25+": .45, "9-16": .1},
   tech=4.6, freq={"1": .15, "2-3": .5, "4-6": .3, "daily": .05},
   devices={"laptop": .85, "desktop": .1, "phone_browser": .6, "phone_app": .45, "tablet": .1},
   primary={"laptop": .55, "phone": .43, "desktop": .02}, session={"2-5": .2, "5-15": .5, "15-30": .3},
   features={"announcements": .85, "assignments": .8, "grades": .5, "modules": .7, "calendar": .35, "inbox": .3,
             "discussions": .15, "quizzes": .3, "recordings": .8, "groups": .3, "syllabus": .2},
   deadline={"canvas_cal": .25, "email": .6, "push": .2, "in_class": .1, "peer": .4, "syllabus": .2, "own_cal": .35, "late": .35},
   notif={"email": .75, "push": .12, "inapp": .1, "dontknow": .03}, lost={"0": .1, "1": .2, "2-3": .4, "4-6": .22, "7+": .08},
   tools={"paper": .4, "calendar": .65, "notes": .2, "chat": .6, "email": .6, "none": .02},
   L=dict(find_fast=3.4, notif_useful=3.3, notif_noise=5.6, mobile=3.4, consistent=3.0, uptodate=4.0, confident=4.4, accessible=5.0, group=3.3),
   sat=3.4, ease=3.2, rank=["deadlines", "recordings", "materials", "grades", "communication", "groupwork"]),
 "power": dict(n=14,
   role={"lecturer": .7, "ta": .3}, age={"25-29": .25, "30-39": .4, "40-49": .25, "50+": .1},
   origin={"nl": .55, "eu": .3, "noneu": .15}, hours={"0": .3, "1-8": .35, "9-16": .25, "17-24": .1},
   tech=5.6, freq={"several": .75, "daily": .25},
   devices={"laptop": .9, "desktop": .55, "phone_browser": .3, "phone_app": .3, "tablet": .15},
   primary={"laptop": .65, "desktop": .35}, session={"15-30": .35, "gt30": .6, "5-15": .05},
   features={"announcements": .9, "assignments": .95, "grades": .9, "modules": .85, "calendar": .45, "inbox": .8,
             "discussions": .35, "quizzes": .5, "recordings": .45, "groups": .6, "syllabus": .5},
   deadline={"canvas_cal": .5, "email": .3, "push": .05, "in_class": .1, "peer": .15, "syllabus": .3, "own_cal": .75, "late": .03},
   notif={"email": .8, "inapp": .1, "off": .1}, lost={"0": .3, "1": .35, "2-3": .3, "4-6": .05},
   tools={"paper": .3, "calendar": .9, "notes": .35, "chat": .2, "email": .8, "none": .02},
   L=dict(find_fast=4.6, notif_useful=3.8, notif_noise=5.5, mobile=2.6, consistent=4.0, uptodate=5.0, confident=6.0, accessible=5.5, group=3.0),
   sat=4.3, ease=3.0, rank=["grades", "materials", "communication", "deadlines", "groupwork", "recordings"]),
 "reluctant": dict(n=12,
   role={"lecturer": .95, "ta": .05}, age={"40-49": .45, "50+": .45, "30-39": .1},
   origin={"nl": .7, "eu": .2, "noneu": .1}, hours={"0": .4, "1-8": .35, "9-16": .25},
   tech=3.0, freq={"lt1": .15, "1": .45, "2-3": .35, "4-6": .05},
   devices={"laptop": .6, "desktop": .7, "phone_browser": .1, "phone_app": .03, "tablet": .15},
   primary={"desktop": .6, "laptop": .38, "tablet": .02}, session={"2-5": .15, "5-15": .5, "15-30": .35},
   features={"announcements": .7, "assignments": .35, "grades": .35, "modules": .65, "calendar": .1, "inbox": .3,
             "discussions": .05, "quizzes": .1, "recordings": .2, "groups": .15, "syllabus": .3},
   deadline={"canvas_cal": .1, "email": .35, "push": .0, "in_class": .1, "peer": .25, "syllabus": .5, "own_cal": .7, "late": .1},
   notif={"email": .6, "inapp": .1, "off": .05, "dontknow": .25}, lost={"0": .15, "1": .3, "2-3": .4, "4-6": .15},
   tools={"paper": .6, "calendar": .7, "notes": .05, "chat": .1, "email": .9, "none": .05},
   L=dict(find_fast=3.0, notif_useful=3.0, notif_noise=4.8, mobile=3.8, consistent=3.2, uptodate=4.0, confident=2.4, accessible=5.0, group=3.0),
   sat=3.2, ease=2.1, rank=["materials", "communication", "deadlines", "grades", "recordings", "groupwork"]),
 "access": dict(n=10,
   role={"bsc": .5, "msc": .4, "phd": .1}, age={"17-20": .2, "21-24": .45, "25-29": .25, "30-39": .1},
   origin={"nl": .6, "eu": .25, "noneu": .15}, hours={"0": .3, "1-8": .35, "9-16": .25, "17-24": .1},
   tech=5.2, freq={"daily": .5, "4-6": .3, "several": .2},
   devices={"laptop": .95, "desktop": .3, "phone_browser": .2, "phone_app": .4, "tablet": .25},
   primary={"laptop": .85, "desktop": .1, "tablet": .05}, session={"5-15": .2, "15-30": .5, "gt30": .3},
   features={"announcements": .85, "assignments": .8, "grades": .6, "modules": .9, "calendar": .5, "inbox": .3,
             "discussions": .25, "quizzes": .3, "recordings": .7, "groups": .25, "syllabus": .35},
   deadline={"canvas_cal": .5, "email": .5, "push": .2, "in_class": .3, "peer": .3, "syllabus": .3, "own_cal": .5, "late": .15},
   notif={"email": .6, "push": .15, "inapp": .25}, lost={"0": .05, "1": .15, "2-3": .35, "4-6": .3, "7+": .15},
   tools={"paper": .15, "calendar": .7, "notes": .6, "chat": .5, "email": .4, "none": .02},
   L=dict(find_fast=2.6, notif_useful=4.0, notif_noise=4.4, mobile=2.6, consistent=2.6, uptodate=4.4, confident=3.6, accessible=2.0, group=3.4),
   sat=2.9, ease=2.4, rank=["materials", "recordings", "deadlines", "grades", "communication", "groupwork"]),
}
# support staff are generated from two behavioural profiles (half power-like, half reluctant-like)
STAFF = [("power", 3), ("reluctant", 3)]

SCHOOL = {"tshd": .44, "tisem": .2, "tsb": .2, "tls": .09, "tst": .02, "other": .05}
GENDER = {"female": .5, "male": .43, "nonbinary": .04, "na": .03}

def courses_for(role):
    return pick({"bsc": {"3": .2, "4": .45, "5": .25, "6+": .1}, "msc": {"2": .3, "3": .45, "4": .2, "5": .05},
                 "phd": {"1": .6, "2": .35, "3": .05}, "lecturer": {"1": .3, "2": .4, "3": .25, "4": .05},
                 "ta": {"1": .55, "2": .4, "3": .05}, "staff": {"1": .1, "2": .1, "4": .15, "5": .15, "6+": .5}}[role])

def access_for(seg):
    if seg == "access":
        primary = pick({"screenreader": .3, "captions": .25, "dyslexia": .35, "motor": .1})
        extra = [k for k in ["contrast", "captions", "dyslexia"] if k != primary and rng.random() < .3]
        return [primary] + extra
    r = rng.random()
    if r < .86: return ["none"]
    if r < .9: return ["na"]
    return [pick({"captions": .4, "contrast": .35, "dyslexia": .25})]

def ranking(order, swaps_lambda=1.4):
    r = list(order)
    for _ in range(rng.poisson(swaps_lambda)):
        i = rng.integers(0, 5); r[i], r[i+1] = r[i+1], r[i]
    if rng.random() < .1:
        i, j = rng.choice(6, 2, replace=False); r[i], r[j] = r[j], r[i]
    return {k: r.index(k) + 1 for k in order}

def fill(t):
    return t.format(course=random.choice(COURSES), n=int(rng.integers(6, 15)), topic=random.choice(TOPICS))

def stylise(t):
    if not t: return t
    r = rng.random()
    if r < .12: t = t[0].lower() + t[1:]
    if rng.random() < .1 and t.endswith("."): t = t[:-1]
    return t

pools_used = {k: {"f": [], "h": [], "m": []} for k in list(FRUSTRATE.keys())}
def draw(pool_dict, seg, kind, p_generic=0.08):
    key = seg if rng.random() > p_generic else "generic"
    used = pools_used[key][kind]
    avail = [t for t in pool_dict[key] if t not in used]
    if not avail:
        used.clear(); avail = list(pool_dict[key])
    t = random.choice(avail); used.append(t)
    return fill(t)

rows, latent = [], []
plan = [(s, SEG[s]["n"], False) for s in SEG] + [(s, n, True) for s, n in STAFF]
for seg, n, is_staff in plan:
    P = SEG[seg]
    for _ in range(n):
        role = "staff" if is_staff else pick(P["role"])
        age = pick({"25-29": .2, "30-39": .4, "40-49": .25, "50+": .15}) if is_staff else pick(P["age"])
        L = {k: likert(v, 1.0) for k, v in P["L"].items()}
        L["find_slow"] = likert(8 - L["find_fast"] + rng.normal(0, 0.9), 0.4)
        sat = likert(P["sat"] + 0.25 * (L["find_fast"] - 4) - 0.15 * (L["notif_noise"] - 4) + 0.15 * (L["confident"] - 4), 0.9)
        ease = likert(P["ease"] + 0.15 * (L["confident"] - 4), 0.8, 1, 5)
        tech = likert(P["tech"], 1.0)
        access = access_for(seg)
        tpool = "staff" if is_staff else seg
        row = {
            "A1_role": role, "A2_age": age, "A3_gender": pick(GENDER), "A4_school": pick(SCHOOL),
            "A5_origin": pick(P["origin"]), "A6_courses": courses_for(role), "A7_hours": pick(P["hours"]),
            "A8_access": access, "A9_tech": tech,
            "B1_freq": pick(P["freq"]), "B2_devices": multi(P["devices"]), "B3_primary": pick(P["primary"]),
            "B4_session": pick(P["session"]), "B5_features": multi(P["features"]),
            "B6_deadline": multi(P["deadline"]), "B7_notif": pick(P["notif"]), "B8_lost": pick(P["lost"]),
            "B9_tools": multi(P["tools"], exclusive_none="none"),
            "C1_find_fast": L["find_fast"], "C2_find_slow": L["find_slow"], "C3_notif_useful": L["notif_useful"],
            "C4_notif_noise": L["notif_noise"], "C5_mobile": L["mobile"], "C6_consistent": L["consistent"],
            "C7_uptodate": L["uptodate"], "C8_confident": L["confident"], "C9_accessible": L["accessible"],
            "C10_group": L["group"], "C11_satisfaction": sat, "C12_ease": ease,
            "D1_rank": ranking(P["rank"]),
            "E1_frustrate": stylise(draw(FRUSTRATE, tpool, "f")),
            "E2_helped": stylise(draw(HELPED, tpool, "h", p_generic=.15)) if rng.random() < .88 else "",
            "E3_more": stylise(draw(MORE, tpool, "m", p_generic=.05)) if rng.random() < .6 else "",
        }
        # a handful of realistic data-quality quirks: one straight-liner, one who skipped optional demographics
        rows.append(row); latent.append(seg if not is_staff else f"staff/{seg}")

# quirks
rows[7]["A2_age"] = "na"; rows[7]["A3_gender"] = "na"; rows[7]["A5_origin"] = "na"
for k in ["C1_find_fast", "C2_find_slow", "C3_notif_useful", "C4_notif_noise", "C5_mobile", "C6_consistent",
          "C7_uptodate", "C8_confident", "C9_accessible", "C10_group"]:
    rows[41][k] = 4  # straight-liner
rows[41]["E1_frustrate"] = "nothing really"; rows[41]["E2_helped"] = ""; rows[41]["E3_more"] = ""

# shuffle respondent order (as if collected over a week) and assign ids
order = rng.permutation(len(rows))
rows = [rows[i] for i in order]; latent = [latent[i] for i in order]

# ---------------------------------------------------------------------------
# flatten to wide CSV
# ---------------------------------------------------------------------------
def opts(qid): return [o["k"] for o in QI[qid]["options"]]
cols = ["id", "submitted"]
for s in Q["sections"]:
    for q in s["questions"]:
        col = f'{q["id"]}_{q["key"]}'
        if q["type"] == "multi":
            cols += [f"{col}__{k}" for k in opts(q["id"])]
        elif q["type"] == "rank":
            cols += [f"{col}__{k}" for k in opts(q["id"])]
        else:
            cols.append(col)

os.makedirs("out", exist_ok=True)
start = np.datetime64("2026-09-08T10:15")
stamps = sorted(start + np.array(sorted(rng.integers(0, 6 * 24 * 60, len(rows))), "timedelta64[m]"))
with open("out/responses.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=cols); w.writeheader()
    for i, r in enumerate(rows):
        flat = {"id": f"R{i+1:03d}", "submitted": str(stamps[i]).replace("T", " ")}
        for k, v in r.items():
            if isinstance(v, list):
                for o in opts(k.split("_")[0]): flat[f"{k}__{o}"] = 1 if o in v else 0
            elif isinstance(v, dict):
                for o, rank in v.items(): flat[f"{k}__{o}"] = rank
            else:
                flat[k] = v
        w.writerow(flat)
with open("out/latent.csv", "w", newline="") as f:
    w = csv.writer(f); w.writerow(["id", "latent"])
    for i, l in enumerate(latent): w.writerow([f"R{i+1:03d}", l])
print(len(rows), "respondents written to out/responses.csv")
