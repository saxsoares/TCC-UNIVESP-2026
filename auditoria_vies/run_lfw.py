# Desempenho de verificação no LFW (6000 pares): ROC, AUC, EER, acurácia.
import sys, numpy as np, json
sys.path.insert(0, "auditoria_vies")
from embedder import embed_batch
from sklearn.metrics import roc_curve, auc
from pathlib import Path

OUT = Path("auditoria_vies/resultados"); OUT.mkdir(parents=True, exist_ok=True)

pairs = np.load("data/lfw_pairs.npy")      # (6000,2,125,94,3) uint8 RGB
target = np.load("data/lfw_target.npy")    # 1=genuino, 0=impostor
N = len(pairs)
print("pares:", N, "genuinos:", int(target.sum()))

# Embeddings em lote
def embed_all(imgs, bs=64):
    out = []
    for i in range(0, len(imgs), bs):
        out.append(embed_batch(list(imgs[i:i+bs])))
    return np.concatenate(out, 0)

print("extraindo embeddings das imagens A...")
embA = embed_all(pairs[:, 0])
print("extraindo embeddings das imagens B...")
embB = embed_all(pairs[:, 1])
sims = np.sum(embA * embB, axis=1)         # cosseno (embeddings ja L2-normalizados)

# ROC / AUC
fpr, tpr, thr = roc_curve(target, sims)
roc_auc = auc(fpr, tpr)
fnr = 1 - tpr
eer_idx = np.nanargmin(np.abs(fnr - fpr))
eer = (fpr[eer_idx] + fnr[eer_idx]) / 2
eer_thr = thr[eer_idx]

# Limiar global @ FMR alvo (1e-3) sobre impostores
imp = sims[target == 0]; gen = sims[target == 1]
thr_1e3 = float(np.quantile(imp, 1 - 1e-3))
fmr_at = float(np.mean(imp >= thr_1e3))
fnmr_at = float(np.mean(gen < thr_1e3))
# Acuracia no melhor limiar (balanceado)
best_acc, best_t = 0, 0
for t in np.linspace(sims.min(), sims.max(), 400):
    acc = np.mean((sims >= t) == (target == 1))
    if acc > best_acc: best_acc, best_t = acc, t

res = {
    "n_pairs": int(N), "AUC": float(roc_auc), "EER": float(eer),
    "EER_threshold": float(eer_thr), "best_accuracy": float(best_acc),
    "best_threshold": float(best_t),
    "threshold@FMR=1e-3": thr_1e3, "FMR@thr": fmr_at, "FNMR@thr": fnmr_at,
    "mean_sim_genuine": float(gen.mean()), "std_sim_genuine": float(gen.std()),
    "mean_sim_impostor": float(imp.mean()), "std_sim_impostor": float(imp.std()),
}
print(json.dumps(res, indent=2))
json.dump(res, open(OUT / "lfw_metrics.json", "w"), indent=2)
np.savez(OUT / "lfw_scores.npz", sims=sims, target=target, fpr=fpr, tpr=tpr)
print("salvo em", OUT)
