# Wholesale Travel Content System — South Africa

A 30-day, compliance-first content system for a wholesale-travel distribution business
selling to the South African market. Built on one rule: **every public claim must rest on
proof the operator actually holds** (real booking screenshots), and earnings claims are
excluded entirely because no typical-earnings data exists to substantiate them.

## What's in the repo

| File | Purpose |
|---|---|
| `30-Day_Wholesale_Travel_Calendar.xlsx` | The core deliverable. 8 tabs: strategy, 30-day calendar (filterable, status dropdowns), DM scripts (RATES + BLUEPRINT flows, objection handling), hooks bank, claim rules, proof checklist, rewrites of the original three assets, weekly tracker. |
| `Week1_Captions_and_Hashtags.md` | Copy-paste captions, Story frames, alt text and first comments for Days 1–7 (+ Day 9), the hashtag rotation system, the never-use tag list, and the caption keyword bank. |
| `carousel_day9/slide1–7.png` | Upload-ready 1080×1350 carousel artwork: "What this role is. What it is not." |
| `carousel_day9/slide1–7.svg` | Editable vector sources for the same slides. |
| `carousel_day9/carousel_preview.html` | All seven slides in one page with the pre-post placeholder checklist. |
| `build_calendar.py` | Regenerates the .xlsx workbook end-to-end (edit content, re-run). |
| `build_carousel.py` | Regenerates the carousel SVGs, PNGs and preview page. |

## Regenerating

```bash
python3 build_calendar.py     # rebuilds 30-Day_Wholesale_Travel_Calendar.xlsx
python3 build_carousel.py     # rebuilds carousel_day9/ (needs ImageMagick for PNG export)
```

Requires: Python 3 with `openpyxl`; ImageMagick (`convert`) for the PNG export step.

## Before anything is posted

1. Replace placeholders: `@yourhandle`, `R{joining}`, `R{monthly}`, `{registered name}`,
   `{number}`, and every example figure (R21,900 / R13,400 / R8,500) with the operator's own
   documented booking.
2. Obtain **written confirmation from the company** that a member may join purely as a
   customer with no recruitment obligation. Day 10 of the calendar must not run without it.
3. Collect customer permissions in writing before using any name, photo or voice note.

## Compliance posture (summary)

- No earnings figures, no "passive/residual income", no "financial freedom" in public content.
- No comment-keyword gating ("comment BLUEPRINT"); all CTAs route to DM.
- Savings claims always carry the supporting screenshot; "up to X%" only next to a shown example.
- Business-side content states that commission is earned only on real travel bookings and that
  results vary.
- South African context: s43(2) and s43(4) of the Consumer Protection Act 68 of 2008 prohibit
  directly or indirectly promoting pyramid schemes; the test is whether compensation derives
  primarily from recruitment rather than from sales of goods or services. This repo's copy is
  written to keep every claim on the product/savings side of that line.
- Marketing guidance, not legal advice. Confirm against the company's own compliance policy.

## Status

Week 1 captions and the Day-9 carousel are complete. Weeks 2–4 caption treatment, brand-colour
carousel variants, and the Day-2/Day-16 carousels are queued as next iterations.
