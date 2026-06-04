# Viés demográfico de falso positivo (FMR) no FairFace.
# Sem identidades repetidas: todo par de imagens distintas é um par impostor;
# mede-se a FMR por grupo sob um limiar global único.
import sys, numpy as np, json, cv2
sys.path.insert(0, "auditoria_vies")
from embedder import embed_batch
import pyarrow.parquet as pq
from pathlib import Path

OUT = Path("auditoria_vies/resultados"); OUT.mkdir(parents=True, exist_ok=True)
RACES  = ['East Asian','Indian','Black','White','Middle Eastern','Latino_Hispanic','Southeast Asian']
GENDER = ['Male','Female']
rng = np.random.default_rng(42)

EMB_CACHE = OUT / "fairface_emb.npz"
if EMB_CACHE.exists():
    d = np.load(EMB_CACHE)
    emb, race, gender, age = d["emb"], d["race"], d["gender"], d["age"]
    print("embeddings carregados do cache:", emb.shape)
else:
    df = pq.ParquetFile("data/fairface_val.parquet").read().to_pandas()
    print("imagens:", len(df))
    imgs, race, gender, age = [], [], [], []
    for _, r in df.iterrows():
        buf = np.frombuffer(r["image"]["bytes"], np.uint8)
        im = cv2.imdecode(buf, cv2.IMREAD_COLOR)      # BGR
        imgs.append(cv2.cvtColor(im, cv2.COLOR_BGR2RGB))
        race.append(r["race"]); gender.append(r["gender"]); age.append(r["age"])
    race = np.array(race); gender = np.array(gender); age = np.array(age)
    embs = []
    for i in range(0, len(imgs), 64):
        embs.append(embed_batch(imgs[i:i+64]))
        if i % 640 == 0: print(f"  {i}/{len(imgs)}")
    emb = np.concatenate(embs, 0)
    np.savez(EMB_CACHE, emb=emb, race=race, gender=gender, age=age)
    print("embeddings salvos:", emb.shape)

def sample_impostor_sims(idx, n_pairs=40000):
    """Amostra n_pairs pares (i!=j) dentre os indices dados e retorna cossenos."""
    if len(idx) < 2: return np.array([])
    a = rng.choice(idx, n_pairs); b = rng.choice(idx, n_pairs)
    keep = a != b; a, b = a[keep], b[keep]
    return np.sum(emb[a] * emb[b], axis=1)

# 1) Limiar global unico calibrado p/ FMR alvo no conjunto inteiro (pooled)
pooled = sample_impostor_sims(np.arange(len(emb)), n_pairs=300000)
TARGET_FMR = 1e-3
thr = float(np.quantile(pooled, 1 - TARGET_FMR))
print(f"limiar global @FMR={TARGET_FMR}: {thr:.4f}  (pooled mean sim={pooled.mean():.4f})")

# Tambem reporta com limiar do LFW, se existir
lfw_thr = None
lfw_json = OUT / "lfw_metrics.json"
if lfw_json.exists():
    lfw_thr = json.load(open(lfw_json))["threshold@FMR=1e-3"]

def boot_ci(vals_bool, n=1000):
    if len(vals_bool)==0: return (float('nan'),float('nan'))
    m = [vals_bool[rng.integers(0,len(vals_bool),len(vals_bool))].mean() for _ in range(n)]
    return float(np.quantile(m,.025)), float(np.quantile(m,.975))

# 2) FMR por grupo (raca) e por raca x genero
rows = []
for ri, rname in enumerate(RACES):
    idx = np.where(race == ri)[0]
    s = sample_impostor_sims(idx, 60000)
    acc = (s >= thr)
    ci = boot_ci(acc)
    row = {"grupo": rname, "n_imgs": int(len(idx)), "n_pairs": int(len(s)),
           "mean_impostor_sim": float(s.mean()), "FMR@global": float(acc.mean()),
           "FMR_CI_low": ci[0], "FMR_CI_high": ci[1]}
    if lfw_thr is not None:
        row["FMR@LFWthr"] = float((s >= lfw_thr).mean())
    rows.append(row)

import pandas as pd
tab = pd.DataFrame(rows).sort_values("FMR@global", ascending=False)
print(tab.to_string(index=False))
tab.to_csv(OUT / "fairface_fmr_por_raca.csv", index=False)

# disparidade
fmrs = tab["FMR@global"].values
disp = float(fmrs.max() / max(fmrs.min(), 1e-9))
print(f"\nDisparidade FMR (max/min) entre racas: {disp:.2f}x")

# 3) interseccional raca x genero
rows2 = []
for ri, rname in enumerate(RACES):
    for gi, gname in enumerate(GENDER):
        idx = np.where((race == ri) & (gender == gi))[0]
        s = sample_impostor_sims(idx, 40000)
        if len(s)==0: continue
        rows2.append({"raca": rname, "genero": gname, "n_imgs": int(len(idx)),
                      "mean_impostor_sim": float(s.mean()), "FMR@global": float((s>=thr).mean())})
tab2 = pd.DataFrame(rows2).sort_values("FMR@global", ascending=False)
tab2.to_csv(OUT / "fairface_fmr_interseccional.csv", index=False)
print("\nTop interseccional (FMR):")
print(tab2.head(6).to_string(index=False))

# 4) Teste qui-quadrado: erro (falso positivo) x raca, em amostra balanceada por grupo
from scipy.stats import chi2_contingency
per_group = []
for ri,rname in enumerate(RACES):
    idx = np.where(race==ri)[0]
    s = sample_impostor_sims(idx, 40000)
    fp = (s>=thr).astype(int)
    per_group.append((rname, int(fp.sum()), int(len(fp)-fp.sum())))
ct = np.array([[a,b] for _,a,b in per_group])
chi2,p,dof,_ = chi2_contingency(ct)
print(f"\nqui-quadrado erro x raca: chi2={chi2:.1f} dof={dof} p={p:.3e}")

summary = {"global_threshold": thr, "target_fmr": TARGET_FMR,
           "disparity_max_min": disp, "chi2": float(chi2), "chi2_p": float(p),
           "lfw_threshold": lfw_thr}
json.dump(summary, open(OUT / "fairface_summary.json","w"), indent=2)
print("\nsalvo em", OUT)
