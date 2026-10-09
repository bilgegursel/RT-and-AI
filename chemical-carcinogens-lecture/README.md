# Chemical Carcinogens — 45-Minute Medical Lecture

English teaching pack for a ~45-minute lecture on chemical carcinogens.

## Files

| File | Purpose |
|------|---------|
| `Chemical_Carcinogens_Lecture_Script.docx` | Full speaking script with timing, slide cues, appendices |
| `Chemical_Carcinogens_Lecture_Slides.pptx` | Widescreen slides with **speaker notes** under each slide |
| `build_lecture.py` | Regenerates both documents |

## How to present

1. Open the PowerPoint in **Presenter View** — the notes pane under each slide contains what to say.
2. Use the Word script for rehearsal or as a printable handout for yourself.
3. Suggested pace: ~110–125 words/minute with brief pauses; spoken notes are ~5,000+ words (~40 minutes of continuous speech plus case/Q&A time ≈ 45 minutes).

## Topics covered

- Definitions and multistage carcinogenesis (initiation / promotion / progression)
- Genotoxic vs non-genotoxic agents; IARC classification
- Mechanisms and metabolic activation
- Major classes: PAHs, aromatic amines, N-nitroso compounds, aflatoxins, metals/asbestos, alkylators, hormones
- Exposure settings, dose/latency, clinical history-taking, prevention
- Mini-case, summary, exam-style questions

## Regenerate

```bash
pip install python-docx python-pptx
python3 build_lecture.py
```
