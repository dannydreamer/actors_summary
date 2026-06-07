# SimCenter — Actor Payment Report Tool
## Full Specification for Claude Code

---

## 1. Overview

A small internal web application for the SimCenter admin. Given a monthly workshop schedule (Excel file), it produces a two-sheet Excel report summarizing:
- All workshops that took place that month
- All actors who worked, how many workshops, and a checkbox column for payment order tracking

One admin user. No login required. No backend. Runs entirely in the browser (or as a simple local Node/Python app — CC to decide simplest path given the constraints below).

---

## 2. Input

The user uploads a `.xlsx` file (the monthly workshop schedule).

- App reads it client-side using **SheetJS (xlsx)**
- No file size concerns — these files are small (~50 sheets, minimal data)

---

## 3. Sheet Structure

Each tab in the workbook = one workshop. The structure of every tab is identical:

```
Row 1: headers — [blank], שם לקוח, מתחקרים.ות, שעת אימון שחקנים, כמות שחקנים, כמות שחקניות, סוג הסימולציה, השתלמות?
Row 2: data    — date (col A), client name (col B), facilitator (col C), training time (col D), ...
Row 3: [סוג קבוצה row] — col E = male actors (confirmed), col F = female actors (confirmed)
Row 4+: scenario rows, notes, etc. — IGNORED
```

**Key fields to extract per sheet:**
| Field | Location |
|-------|----------|
| Date | Row 2, col A (e.g. `1.2`, `17.6`) |
| Client name | Row 2, col B |
| Male actors | Row 3, col E |
| Female actors | Row 3, col F |

**Sheet name** = used only to detect skip conditions (see §4).

---

## 4. Parsing Rules

### 4.1 Skip conditions
Skip a sheet entirely (do not include in output) if the sheet **name** contains:
- `נדחה` (postponed)
- `מבוטל` (cancelled)

Case-insensitive, anywhere in the name.

### 4.2 Actor name parsing

Actor names appear in row 3, cols E and F. The **canonical format** is:

> Full names separated by commas: `קובי ירחי, ישי מאיר, אליאב שאול`

The parser must handle **all of the following separators** (legacy/human error):
- Comma: `קובי ירחי, עדי זינגר`
- Hebrew ו as conjunction with space before it: `קובי ירחי ועדי זינגר` → split on ` ו`
- Leading ו glued to a name: `ויעלה` → strip leading ו → `יעלה`

**Algorithm:**
1. Strip leading ו if glued to first character (e.g. `ויעלה` → `יעלה`)
2. Split on `,` (comma)
3. For each chunk, split on ` ו` (space + ו)
4. Trim whitespace and `*` from each result
5. Discard empty strings and strings shorter than 2 characters

### 4.3 Name deduplication
Within a single month's report, the same actor may appear as a full name (`קובי ירחי`) in one sheet and first name only (`קובי`) in another.

**Rule:** If a full name (2+ words) is seen anywhere in the sheet, map any single-word match of the first name to that full name.

Example: `קובי ירחי` seen → `קובי` elsewhere → canonical = `קובי ירחי`

If two full names share a first name (e.g. `דניאל כהן` and `דניאל לוי`), do NOT deduplicate — keep both as-is.

---

## 5. Validation (runs before any output is generated)

After parsing, before showing any results, validate every extracted actor name:

**Rule:** Any name that is a **single word** (no space) is considered incomplete and must be flagged.

If any violations found → **block output entirely** and show an error screen:

```
⚠ נמצאו שמות חלקיים — יש לתקן בגיליון לפני המשך

• יובל — 18.2 משרד החוץ
• קרן — 17.2 ישיבת בני עקיבא

נא לוודא שכל שם שחקן/ית כולל שם פרטי ושם משפחה, מופרדים בפסיק.
```

User fixes the source sheet, re-pastes the URL or re-uploads the file, tries again.

**Only if zero violations** → proceed to generate output.

**Exception:** Empty actor cells are valid (workshop had no actors listed) — not an error.

---

## 6. Output — Excel File

Generated client-side using **SheetJS** and downloaded automatically.

Filename: `actor_report_{month}_{year}.xlsx` where month/year are inferred from the dates in the data (e.g. `actor_report_02_2025.xlsx`).

The workbook has **two sheets**:

---

### Sheet 1: `סיכום סדנאות` (Workshop Summary)

One row per workshop (skipped sheets excluded).

| Column | Content |
|--------|---------|
| תאריך | Date from row 2 col A |
| שם לקוח | Client name from row 2 col B |
| שחקנים | Comma-separated canonical actor names; `—` if none |

- Sorted by date (ascending)
- RTL sheet direction
- Alternating row shading

---

### Sheet 2: `סיכום שחקנים` (Actor Summary)

One row per unique actor (after deduplication).

| Column | Content |
|--------|---------|
| שם שחקן/ית | Canonical full name |
| מספר סדנאות | Count of workshops this actor appeared in |
| שמות סדנאות | Comma-separated list of client names |
| קיבלנו הזמנת תשלום | Empty — admin fills in ✓ manually |

- Sorted alphabetically by actor name
- Total row at bottom: `סה"כ שחקנים` + count
- RTL sheet direction
- Alternating row shading
- Payment column has a distinct light green fill to draw the eye

---

## 7. UI / UX

Single-page app. Three states:

### State 1 — Input
- Single input: **העלאת קובץ Excel** — file upload input (`.xlsx` only)
- Hebrew UI, RTL layout

### State 2 — Validation Error
- Clear error panel listing every problematic name and the workshop it came from
- "חזור" button to reset and try again
- No partial output shown

### State 3 — Success
- Brief summary: `X סדנאות | Y שחקנים`
- "הורד דוח Excel" download button
- Option to start over

---

## 8. Tech Stack Recommendation

Since this is a single admin user with no backend needs:

- **Pure HTML + JS** (single file, no build step) — simplest to deploy and maintain
- **SheetJS** (xlsx) for both reading uploaded Excel AND writing output Excel
- No framework required — vanilla JS is fine given the scope

Alternatively a minimal React app if CC prefers, but keep it dependency-light.

---

## 9. Edge Cases

| Case | Handling |
|------|---------|
| Sheet with no actors in E/F | Include in workshop summary with `—`, skip in actor summary |
| Same client appears multiple times in a month | Each appearance = separate row in workshop summary |
| Actor name with `*` or punctuation | Strip `*`, trim whitespace |
| Leading ו glued to name (e.g. `ויעלה`) | Strip the ו |
| Sheet name contains נדחה or מבוטל | Skip entirely |
| Empty sheet (no data) | Skip silently |
| Two actors with same first name, different last names | Keep as separate actors, no dedup |

---

## 10. What This Tool Is NOT

- Not multi-user
- Not a database — no persistence beyond the downloaded Excel
- Not connected to accounting software
- Not responsible for sending payment orders — only tracking receipt of them (the ✓ column)
