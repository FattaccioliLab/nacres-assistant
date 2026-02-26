# 🧪 NACRES Code Assistant — Google Sheets + Claude AI

A Google Apps Script that automatically assigns **NACRES codes** to laboratory purchase items, from a supplier quotation PDF or a free-text list. Built for French academic research labs.

> **NACRES** (Nomenclature Achats Recherche Enseignement Supérieur) is the standard purchasing classification used by French public research institutions (CNRS, INSERM, universities…). The script embeds the full official **2026 Amue nomenclature** (1621 active codes).

---

## Features

- 📄 **PDF input** — upload a supplier quote to Google Drive, the script extracts all line items automatically
- 📋 **Text list input** — paste a free-text list of items, one per row
- 🤖 **AI-powered classification** — uses Claude (Anthropic) to suggest the most appropriate NACRES code for each item, constrained to the official 2026 nomenclature
- 🟢 **Confidence indicator** — color-coded High / Medium / Low for each suggestion
- ✏️ **Inline validation** — review and correct codes directly in the sheet before sending to admin
- 📚 **Editable reference sheet** — full 1621-code nomenclature visible and editable in a dedicated tab, no script change needed for yearly updates
- 💶 **Near-zero cost** — uses Claude Haiku (~$0.001 per quotation)

---

## Prerequisites

- A Google account (to use Google Sheets + Google Drive)
- An [Anthropic API key](https://console.anthropic.com) (free to create, pay-as-you-go)

---

## Installation

### 1. Create a new Google Sheet

Open [Google Sheets](https://sheets.google.com) and create a blank spreadsheet.

### 2. Open the Apps Script editor

**Extensions > Apps Script**

### 3. Paste the script

Select all (Ctrl+A), delete the default content, paste the entire contents of [`NACRES_script.gs`](./NACRES_script.gs), and save (Ctrl+S or Cmd+S).

### 4. Reload the spreadsheet

Close the script editor tab and **refresh** your Google Sheet. A **🧪 NACRES** menu will appear in the menu bar.

### 5. Set your API key

**🧪 NACRES > ⚙️ Set API Key**

Paste your Anthropic API key (starts with `sk-ant-…`). It is stored securely in Script Properties — never in any cell, never visible to other users.

### 6. Load the reference sheet (first time only)

**🧪 NACRES > 📚 Create / reset reference sheet**

This writes all 1621 NACRES 2026 codes into a `NACRES_ref` tab. The script reads this sheet at runtime — Claude will only assign codes that exist in it. This step is optional (the codes are also embedded in the script as a fallback), but recommended so the nomenclature is visible and editable by the team.

---

## Usage

### Option A — Process a PDF quotation

1. Upload the supplier PDF to **Google Drive**
2. In the sheet: **🧪 NACRES > 📄 Process PDF (Drive)**
3. Paste the Google Drive file URL or file ID when prompted
4. Wait ~15–30 seconds — results appear in the **Résultats** sheet

### Option B — Process a text list

1. **🧪 NACRES > 📋 Process text list** (first run creates an **Input** sheet)
2. Paste your item list in column A of the **Input** sheet (one item per row)
3. Run **📋 Process text list** again
4. Results appear in the **Résultats** sheet

### Output sheet columns

| Column | Content |
|--------|---------|
| A | Supplier reference / catalog number |
| B | Item description |
| C | Quantity |
| D | Unit price |
| E | NACRES code suggested by Claude |
| F | Confidence (🟢 High / 🟡 Medium / 🔴 Low) |
| **G** | **✏️ Validated NACRES** ← edit this column if needed |
| H | Justification (Claude's reasoning) |

Column G is pre-filled with Claude's suggestion. Correct it directly in the cell if needed. **This is the column to send to your admin office.**

---

## Keeping the nomenclature up to date

The NACRES nomenclature is updated yearly by Amue (typically in January). When a new version is released:

1. Download the new Excel file from [amue.fr > Dossier NACRES > Documentation](https://www.amue.fr/finances/metier/dossier-nacres/documentation/)
2. Open the `NACRES_ref` sheet in your Google Sheet
3. Replace the content with the new codes (two columns: code + label)
4. That's it — no script change needed. The next API call will use the updated list.

For a full reset to the embedded 2026 defaults, use **🧪 NACRES > 📚 Create / reset reference sheet** (this will overwrite any manual edits).

> The `NACRES_ref` sheet also serves as a human-readable reference for all lab members — they can consult it at any time to look up a code manually without running the AI.

A standalone CSV of the 2026 nomenclature is also included in this repository: [`nacres_2026_reference.csv`](./nacres_2026_reference.csv).

---

## NACRES codes — key families for a research lab

The official nomenclature is maintained by the **Amue** and updated yearly:

🔗 [Nomenclature NACRES — Amue (documentation officielle)](https://www.amue.fr/finances/metier/dossier-nacres/documentation/)

The full 1621-code list is in [`nacres_2026_reference.csv`](./nacres_2026_reference.csv).

---

## Cost estimate

The script uses **Claude Haiku**, Anthropic's lightest model, which is entirely sufficient for classification tasks.

| Scenario | Approximate cost |
|----------|-----------------|
| 1 quotation (~20 items) | < $0.001 |
| 100 quotations/year | ~$0.05–0.10 |

You only pay for actual API usage. There is no subscription required beyond your Anthropic account.

---

## Sharing with your lab

Simply **share the Google Sheet** with your lab members (Editor access). The API key is stored at the script level — all users benefit from it without any additional setup. Only one person (typically the PI or lab manager) needs to set the key once.

> ⚠️ All API calls are billed to the account associated with the API key. For a lab, a shared key managed by the PI or lab manager is the simplest approach.

---

## Troubleshooting

**"API key not set"** — Run 🧪 NACRES > ⚙️ Set API Key and paste your key.

**"Could not access file"** — Make sure the PDF is in Google Drive and accessible to the Google account running the script.

**"Claude API error 401"** — Your API key is invalid or expired. Generate a new one at [console.anthropic.com](https://console.anthropic.com).

**PDF extraction incomplete** — Very large or scanned (image-only) PDFs may yield incomplete results. For scanned PDFs, run OCR first (Google Drive does this automatically when you open a PDF with Google Docs, then re-export as PDF).

**Wrong or unexpected NACRES codes** — Check that the `NACRES_ref` sheet exists and is populated. If it is empty, run **📚 Create / reset reference sheet**.

---

## Customisation

To use a more capable model (higher accuracy, slightly higher cost), edit the top of the script:

```javascript
const CLAUDE_MODEL = "claude-haiku-4-5-20251001";  // change to claude-sonnet-4-6 for higher accuracy
```

To adapt the nomenclature to your institution's specific subset, edit the `NACRES_ref` sheet directly — no script modification needed.

---

## Repository contents

| File | Description |
|------|-------------|
| `NACRES_script.gs` | Google Apps Script — paste into your sheet's script editor |
| `nacres_2026_reference.csv` | Full NACRES 2026 nomenclature (1621 codes, source: Amue) |
| `README.md` | This file |

---

## License

GPL 

