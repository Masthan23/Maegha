# 🎬 Video QC Tool — Local

Processes 70+ videos (including 1 GB+ files) from a local folder with parallel workers and live results.

## Install & run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501

## Usage

1. Paste your folder path (e.g. `C:\Videos\Project` or `/home/you/videos`)
2. Set Expected Width × Height
3. Adjust Frame Samples (10 is good; raise to 20 for thorough border checks)
4. Set Parallel Workers (4 is safe; raise to 6-8 if you have 16+ GB RAM)
5. Click **▶ Run QC** — results appear live as each video finishes
6. Use **⏹ Stop** to cancel mid-run
7. Download CSV or Excel report

## Memory note

Videos are never loaded fully into RAM.
`cv2.VideoCapture` streams the file and seeks to sample frames only.
A 1 GB video uses under 10 MB of RAM during processing.

## Parallel Workers guide

| RAM   | Recommended workers |
|-------|---------------------|
| 8 GB  | 2–3                 |
| 16 GB | 4–5                 |
| 32 GB | 6–8                 |
