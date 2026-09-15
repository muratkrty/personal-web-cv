"""Step 2 + 3 of the persona process: find patterns, cluster segments. Writes out/analysis.json
and out/responses_clustered.csv (adds Segment, PC1, PC2)."""
import json, numpy as np, pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score, adjusted_rand_score
from sklearn.decomposition import PCA

Q = json.load(open("questions.json"))
QI = {q["id"]: q for s in Q["sections"] for q in s["questions"]}
d = pd.read_csv("out/responses.csv")
lat = pd.read_csv("out/latent.csv").set_index("id")["latent"]

def vmap(qid):
    return {o["k"]: o.get("v") for o in QI[qid]["options"]}

# numeric recodes (Muller et al. Table 55.2 style numerical equivalences)
d["freq_v"] = d["B1_freq"].map(vmap("B1"))
d["session_v"] = d["B4_session"].map(vmap("B4"))
d["hours_v"] = d["A7_hours"].map(vmap("A7"))
d["lost_v"] = d["B8_lost"].map(vmap("B8"))
d["courses_v"] = d["A6_courses"].map(vmap("A6"))
d["phone_primary"] = (d["B3_primary"] == "phone").astype(int)
d["C2r_find_slow_rev"] = 8 - d["C2_find_slow"]

likert_cols = ["C1_find_fast", "C2_find_slow", "C3_notif_useful", "C4_notif_noise", "C5_mobile", "C6_consistent",
               "C7_uptodate", "C8_confident", "C9_accessible", "C10_group"]
rank_cols = [c for c in d.columns if c.startswith("D1_rank__")]

# --- reliability of the inverted pair (C1, C2 reversed): 2-item Cronbach alpha = 2r/(1+r)
r = np.corrcoef(d["C1_find_fast"], d["C2r_find_slow_rev"])[0, 1]
alpha = 2 * r / (1 + r)

# --- clustering features
feat = ["freq_v", "session_v", "phone_primary", "A9_tech", "hours_v", "lost_v"] + likert_cols + \
       ["C11_satisfaction", "C12_ease"] + rank_cols
X = StandardScaler().fit_transform(d[feat].values)

ks = list(range(2, 9)); inertia, sil = [], []
for k in ks:
    km = KMeans(n_clusters=k, n_init=25, random_state=2026).fit(X)
    inertia.append(float(km.inertia_)); sil.append(float(silhouette_score(X, km.labels_)))
    print(f"k={k}: inertia={km.inertia_:.1f} silhouette={sil[-1]:.3f} ARI vs latent={adjusted_rand_score(lat.loc[d['id']].str.replace('staff/', ''), km.labels_):.3f}")

K = 6
km = KMeans(n_clusters=K, n_init=50, random_state=2026).fit(X)
d["cluster_raw"] = km.labels_
ct = pd.crosstab(d["cluster_raw"], lat.loc[d["id"]].values)
print(ct)

pca = PCA(n_components=2, random_state=0).fit(X)
pcs = pca.transform(X)
d["PC1"], d["PC2"] = pcs[:, 0].round(3), pcs[:, 1].round(3)
print("explained variance", pca.explained_variance_ratio_)

# name clusters by their dominant latent profile (author check), then relabel 1..K ordered by size
dom = ct.idxmax(axis=1).to_dict()
print("dominant latent per raw cluster:", dom)
d.to_csv("out/responses_clustered_raw.csv", index=False)
json.dump({"ks": ks, "inertia": inertia, "silhouette": sil, "alpha_pair": alpha, "r_pair": r,
           "explained": pca.explained_variance_ratio_.tolist(), "dominant": {int(k): v for k, v in dom.items()},
           "sizes": d["cluster_raw"].value_counts().to_dict()}, open("out/cluster_diag.json", "w"), indent=1, default=str)
