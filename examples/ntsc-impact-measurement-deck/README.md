# NTSC — Participation Impact Measurement Model

A redesigned, corporate-grade presentation for the **National Transport Safety
Center** (المركز الوطني لسلامة النقل · NTSC), Communications Department:
the *Participation Impact Measurement Model* (نموذج قياس أثر المشاركات في
الفعاليات والمؤتمرات).

Built with the `ui-ux-pro-max` / `slides` design skill as a worked example of a
bilingual (Arabic-first, RTL) executive deck under a strict government brand
system.

## Deliverables

| File | Purpose |
|------|---------|
| `NTSC-Impact-Measurement-Model.pdf` | Print/share-ready deck (pixel-verified) |
| `NTSC-Impact-Measurement-Model.pptx` | Editable Microsoft PowerPoint (native shapes & text) |
| `ntsc-impact-model.html` | Self-contained HTML source (1280×720 slides, IBM Plex Sans Arabic) |
| `preview/slide-01..09.png` | Retina PNG of each slide |
| `assets/` | Official NTSC logo + palm/road emblem (SVG/PNG, recolored variants) |
| `render.py` | Renders the HTML to PNG + PDF via headless Chromium |
| `make_pptx.py` | Generates the editable PPTX via python-pptx |
| `NTSC-SMART-Goal-Template.docx` | Ready-to-use, fillable SMART goal template for event participation (RTL Arabic) |
| `make_smart_template.py` | Generates the SMART goal template via python-docx |

## SMART Goal Template

A fillable Word document for the **planning (before-event) stage**, tailored to
NTSC event participation. Sections: event data · overall objective · the SMART
framework (Specific/Measurable/Achievable/Relevant/Time-bound) · a repeatable
goal worksheet · a pre-filled worked example · a KPI-target reference linked to
the deck's measures (5–10% visitors, 2–5 min dwell, 2 partnerships, 1 media
coverage) · preparation & approval sign-off. Swap the font to brand *TheSans*
where installed.

```bash
python3 make_smart_template.py   # -> NTSC-SMART-Goal-Template.docx
```

## Brand system (official NTSC guidelines)

| Token | Hex | Role |
|-------|-----|------|
| Evening Sea | `#004E43` | Primary |
| Gable Green | `#17332F` | Deep secondary |
| Highland | `#71946A` | Sage support |
| Swirl / Paper | `#D5D0C3` / `#F6F4EE` | Neutral grounds |
| Anzac Gold | `#DEB83B` | Accent |

Type: IBM Plex Sans Arabic + IBM Plex Sans (HTML) / Segoe UI (PPTX, swap to the
brand *TheSans* where installed).

## Deck structure (9 slides)

1. Cover
2. The Mechanism — definition + 3 objectives
3. 5-Phase Methodology — RTL journey diagram
4. Data Collection — before vs during/after
5. Outputs & KPIs — scorecard + target gauge
6. Impact & Optimization
7. Knowledge Management — archive flow + cadence
8. Post-Participation Report — the deliverable template
9. Closing

## Rebuild

```bash
pip install playwright pillow pymupdf python-pptx cairosvg
python3 render.py      # HTML -> preview/*.png + PDF
python3 make_pptx.py   # -> editable .pptx
```
