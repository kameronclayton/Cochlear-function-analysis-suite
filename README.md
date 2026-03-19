# CoFAST: COchlear Function Analysis SuiTe

**Developed by the Functional Testing Core at Eaton-Peabody Laboratories (EPL).**
Integrates with the EPL CFTS acquisition system.

A point-and-click GUI application for analysing auditory brainstem response (ABR) and distortion product otoacoustic emission (DPOAE) data.

---

## Quick start

1. **Double-click `Run_ABR_Tool.bat`** (Windows) — finds Python automatically, installs missing packages, and launches the app.
2. Click **➕ Add Data** to load a folder or individual files.
3. Navigate the tabs to extract, plot, and export your data.

> **No coding required.** The only prerequisite is Python 3.9+ (Anaconda works perfectly).
> Or download a standalone executable from the [Releases](../../releases) page — no Python needed at all.

---

## Features

| Tab | What it does |
|-----|--------------|
| **Overview** | User guide, keyboard reference, and a live summary of every loaded session |
| **ABR Thresholds** | Table of thresholds (dB SPL) — rows: animals, columns: frequencies. Plot + Excel export. |
| **Wave Growth** | Peak-to-trough amplitude (P1–P5 minus N1–N5, µV) vs. level for a selected frequency and wave. Plot + Excel export. |
| **Latencies** | P and N peak latencies (ms) vs. level for a selected frequency and wave. Plot + Excel export. |
| **Plot Traces** | Mean ± SEM waveform traces by group. Four layout modes: Stacked, Overlay, By Group, Mean + Individuals. |
| **DPOAE Thresholds** | DPOAE threshold table and audiogram. Dual criterion: absolute DP level + 5 dB SNR above noise floor. |

---

## Expected folder structure

```
Root folder/
├── GroupA DATA/
│   ├── Mouse001/
│   │   ├── ABR-1-1              ← raw CFTS waveform file
│   │   ├── ABR-1-1-analyzed.txt ← peak-labelled file (EPL Peak Analysis)
│   │   ├── DP-1-1               ← DPOAE data file
│   │   └── ...
│   └── Mouse002/
│       └── ...
└── GroupB DATA/
    └── ...
```

- The **group** is the name of the parent folder (e.g. `GroupA DATA`). Assign `All` or `N/A` to exclude an animal from plots.
- The **animal ID** is the name of the innermost folder (e.g. `Mouse001`).

---

## Peak Analysis (keyboard shortcuts)

Open the peak analysis window by double-clicking an unanalyzed row.

| Key | Action |
|-----|--------|
| ↑ / ↓ | Change level |
| P | Invert polarity |
| N | Toggle normalised view |
| + / − | Scale up / down |
| 1–5 | Select P1–P5 |
| Shift + 1–5 | Select N1–N5 |
| ← / → | Snap peak to nearest extremum |
| Shift + ← / → | Fine-adjust ±1 sample |
| I | Auto-detect all peaks |
| U | Propagate peak to lower levels |
| T | Auto-estimate threshold (Suthakar & Liberman, 2019) |
| Enter | Set threshold to current level |
| S | Save & advance to next file |
| X | Clear analysis & re-detect |
| R | Restore last saved analysis |
| Ctrl + Z | Undo |

---

## Installation (manual / development)

```bash
pip install -r requirements.txt
python CoFAST.py
```

### Requirements
- Python 3.9+
- numpy, pandas, openpyxl, matplotlib, scipy

---

## Acknowledgements

The peak detection algorithm and waveform analysis approach were inspired by the **EPL ABR Peak Analysis** software developed by **Brad Buran** and **Ken Hancock** at the Eaton-Peabody Laboratories, Massachusetts Eye and Ear ([source](https://github.com/EPL-Engineering/abr-peak-analysis)).

The correlation-based automatic threshold estimation (T key) implements the adjacent-level cross-covariance method described in **Suthakar & Liberman (2019)**.
