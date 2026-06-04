# Auditoria empírica de viés — material computacional do TCC

Código da auditoria de viés demográfico em reconhecimento facial que acompanha o
TCC *"Análise de Viabilidade Técnica e Jurídica no Uso da Inteligência Artificial
para Apoio à Identificação de Pessoas Desaparecidas"* (UNIVESP, 2026).

A auditoria avalia o modelo ArcFace pré-treinado (sem treino do zero) sobre bases
públicas de pesquisa, medindo desempenho de verificação e disparidade de erro
entre grupos demográficos. Não usa imagens de pessoas reais desaparecidas.

## Ambiente

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Dados e modelo (não versionados)

Não ficam no repositório por causa do tamanho. Baixar antes de rodar:

- **Modelo ArcFace** (`models/arcface_w600k_r50.onnx`): pacote `buffalo_l`
  do InsightFace (espelho `immich-app/buffalo_l` no Hugging Face).
- **LFW**: baixado automaticamente por `sklearn.datasets.fetch_lfw_pairs`.
- **FairFace** (`data/fairface_val.parquet`): validação 0.25 do repositório
  `HuggingFaceM4/FairFace` no Hugging Face.

## Execução

```
python run_lfw.py        # desempenho de verificação no LFW -> ROC, AUC, EER
python run_fairface.py   # FMR por grupo demográfico no FairFace
python gen_figures.py    # figuras 1 e 2
```

Os scripts `edit_docx.py` e `edit_pptx.py` inserem os resultados no TCC e na
apresentação (executar uma vez cada).

## Saídas (`resultados/`)

- `lfw_metrics.json`, `fairface_*.csv`, `fairface_summary.json` — métricas.
- `fig1_lfw_roc.png`, `fig2_fairface_fmr.png` — figuras usadas no trabalho.

## Arquivos

| Arquivo | Função |
|---|---|
| `embedder.py` | Extração de embeddings ArcFace via ONNX Runtime |
| `run_lfw.py` | Experimento de verificação (LFW) |
| `run_fairface.py` | Experimento de viés demográfico (FairFace) |
| `gen_figures.py` | Geração das figuras |
| `edit_docx.py` / `edit_pptx.py` | Inserção dos resultados nos documentos |
