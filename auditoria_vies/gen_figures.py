"""Gera as figuras do TCC a partir dos resultados reais dos experimentos."""
import numpy as np, json, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from pathlib import Path
from sklearn.metrics import auc

OUT = Path("auditoria_vies/resultados")
plt.rcParams.update({"font.size": 11, "font.family": "DejaVu Sans", "figure.dpi": 200})

# ---- Fig 1: distribuicoes de similaridade + ROC (LFW) ----
z = np.load(OUT / "lfw_scores.npz")
sims, target, fpr, tpr = z["sims"], z["target"], z["fpr"], z["tpr"]
m = json.load(open(OUT / "lfw_metrics.json"))

fig, ax = plt.subplots(1, 2, figsize=(10, 4))
ax[0].hist(sims[target == 1], bins=50, alpha=.6, label="Genuínos (mesma pessoa)", color="#2a7")
ax[0].hist(sims[target == 0], bins=50, alpha=.6, label="Impostores", color="#c33")
ax[0].axvline(m["EER_threshold"], ls="--", c="k", lw=1, label="Limiar (EER)")
ax[0].set_xlabel("Similaridade de cosseno"); ax[0].set_ylabel("Frequência")
ax[0].set_title("(a) Distribuição de similaridade — LFW"); ax[0].legend(fontsize=8)

ax[1].plot(fpr, tpr, c="#06c", lw=2, label=f"ArcFace (AUC = {m['AUC']:.4f})")
ax[1].plot([0, 1], [0, 1], "k--", lw=.8)
ax[1].scatter([m["EER"]], [1 - m["EER"]], c="r", zorder=5, label=f"EER = {m['EER']*100:.2f}%")
ax[1].set_xlabel("Taxa de falsa correspondência (FMR)")
ax[1].set_ylabel("Taxa de correspondência verdadeira (1−FNMR)")
ax[1].set_title("(b) Curva ROC — LFW"); ax[1].legend(fontsize=8)
fig.tight_layout(); fig.savefig(OUT / "fig1_lfw_roc.png"); plt.close(fig)
print("fig1 ok")

# ---- Fig 2: FMR por grupo demografico (FairFace) ----
tab = pd.read_csv(OUT / "fairface_fmr_por_raca.csv")
tab = tab.sort_values("FMR@global", ascending=True)
fig, ax = plt.subplots(figsize=(8, 4.5))
yerr = [tab["FMR@global"] - tab["FMR_CI_low"], tab["FMR_CI_high"] - tab["FMR@global"]]
bars = ax.barh(tab["grupo"], tab["FMR@global"], xerr=yerr, color="#3a6ea5", capsize=3)
ax.set_xlabel("FMR sob limiar global único (calibrado a FMR médio = 1e-3)")
ax.set_title("Taxa de falsa correspondência por grupo demográfico — FairFace")
for b, v in zip(bars, tab["FMR@global"]):
    ax.text(v, b.get_y()+b.get_height()/2, f" {v*100:.3f}%", va="center", fontsize=8)
fig.tight_layout(); fig.savefig(OUT / "fig2_fairface_fmr.png"); plt.close(fig)
print("fig2 ok")

# ---- Fig 3: similaridade media de impostores por grupo (proxy do vies) ----
tab2 = tab.sort_values("mean_impostor_sim", ascending=True)
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.barh(tab2["grupo"], tab2["mean_impostor_sim"], color="#a5683a")
ax.set_xlabel("Similaridade média de pares impostores (intra-grupo)")
ax.set_title("Propensão a falsa correspondência por grupo — FairFace")
fig.tight_layout(); fig.savefig(OUT / "fig3_fairface_simmedia.png"); plt.close(fig)
print("fig3 ok")
print("figuras salvas em", OUT)
