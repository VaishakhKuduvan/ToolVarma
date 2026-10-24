# Audit Suite

One browser-based tool for Indian audit practitioners, built from four separate scripts:

| Page | Based on | What it does |
|---|---|---|
| 🎯 1. Risk Assessment (SA 315) | `Audit Sampling using Python/app.py` | Weighted questionnaire across 8 domains → risk score, level and audit approach |
| 💰 2. Materiality (SA 320) | same | Benchmark recommendation from 3 years of financials → OM, PM, SM, CTT |
| 📋 3. Sampling (SA 530) | same | Key items, red flags and value-weighted samples from a Trial Balance and a Day Book |
| 📊 Benford's Law Analysis | `BENFORD_MATERIALITY_NEWCHARTS_ANAMOLIES_NOREPETITION.py` | Nigrini 5-test per materiality tier, Excel workpaper of unique anomalies |
| 🤖 AI Ledger Mapping | `ledger_ai_studio.py` | Gemini maps ledgers to Schedule III, scores risk 1–5, writes findings |
| ✉️ External Confirmations | `externalconfirmation.py` | Mail-merge letters from a .docx / fillable .pdf template → ZIP download |

## Run it

**Windows:** double-click `run.bat`. The first run sets up a local Python environment (needs Python 3.10+).

**Any OS:**

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens at http://localhost:8501.

## Share it online (Streamlit Community Cloud)

1. Push this folder to a GitHub repository (`app.py` at the repo root).
2. Sign in at https://share.streamlit.io with that GitHub account, then **Create app** and pick the repo, branch `main` and file `app.py`.
3. Under **Advanced settings**, choose Python 3.12, then click **Deploy**. The first build takes a few minutes.
4. To limit who can open it, keep the repo private and add viewers' emails under the app's **Share** settings.

Each push to `main` redeploys the app. `packages.txt` installs Chromium so the Benford Excel charts render on the server.
Don't add a shared Gemini key to the app's secrets: each user pastes their own key on the AI page.
Files people upload are processed on Streamlit's servers, so check this fits your firm's data policy.

## How the pages connect

- The **risk level** from page 1 sets the PM % on page 2 (High 55 / Medium 70 / Low 80) and the sample sizes on page 3.
- **Materiality** from page 2 (₹ Lakhs) is converted to ₹ and offered as the Benford thresholds.
- A **Day Book** uploaded on Sampling or Benford can be reused on the other page without uploading again.
- The **client name / FY** in the sidebar prefixes every downloaded file name.
- Answers stay in place when you switch pages. They are kept only while the browser tab is open, so download the Excel workpapers to keep a record.

## Notes

- **AI Ledger Mapping** needs a Google Gemini API key (free at https://aistudio.google.com/apikey). Paste it on the page, or set the `GEMINI_API_KEY` environment variable before starting. The ledger data is sent to Google.
- **Benford charts in Excel** are rendered with Kaleido, which needs Google Chrome installed. Without it, the workbook is still produced but without chart images.
- Sample input files are in the original folders under `../extracted/`.

## Troubleshooting

**"DLL load failed ... An Application Control policy has blocked this file"** means Windows Smart App Control
is blocking a package version it has not seen before. On the machine this was built on, these versions load fine:

```bash
.venv\Scripts\python -m pip install pandas==3.0.5 numpy==2.5.2 pyarrow==25.0.1
```

SciPy is deliberately not used. Its newest DLLs were blocked the same way, so the Benford chi-square test is computed in plain Python.

## Layout

```
app.py              entry point and navigation
core/               engines (no UI): sampling, benford, ledger_ai, confirmations, shared state
views/              one Streamlit page per module
```
