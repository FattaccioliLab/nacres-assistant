# 🧪 NACRES Code Assistant — Google Sheets + Claude AI

A Google Apps Script that automatically assigns **NACRES codes** to laboratory purchase items, from a supplier quotation PDF or a free-text list. Built for French academic research labs.

> **NACRES** (Nomenclature Analytique des Coûts de la Recherche pour les Établissements Scientifiques) is the standard purchasing classification used by French public research institutions (CNRS, INSERM, universities…).

---

## Features

- 📄 **PDF input** — upload a supplier quote to Google Drive, the script extracts all line items automatically
- 📋 **Text list input** — paste a free-text list of items, one per row
- 🤖 **AI-powered classification** — uses Claude (Anthropic) to suggest the most appropriate NACRES code for each item
- 🟢 **Confidence indicator** — color-coded High / Medium / Low for each suggestion
- ✏️ **Inline validation** — review and correct codes directly in the sheet before sending to admin
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

Delete the default `myFunction()` content, paste the entire contents of [`NACRES_script.gs`](./NACRES_script.gs), and save (Ctrl+S or Cmd+S).

### 4. Reload the spreadsheet

Close the script editor tab and **refresh** your Google Sheet. A **🧪 NACRES** menu will appear in the menu bar.

### 5. Set your API key

**🧪 NACRES > ⚙️ Set API Key**

Paste your Anthropic API key (starts with `sk-ant-…`). It is stored securely in Script Properties — never in any cell, never visible to other users.

---

## Usage

### Option A — Process a PDF quotation

1. Upload the supplier PDF to **Google Drive**
2. In the sheet: **🧪 NACRES > 📄 Process PDF (Drive)**
3. Paste the Google Drive file URL or file ID when prompted
4. Wait ~15 seconds — results appear in the **Résultats** sheet

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

Column G is pre-filled with Claude's suggestion. Correct it directly in the cell if needed. This is the column to send to your admin office.

---

## NACRES codes reference

| Code | Category |
|------|----------|
| NA.10 | Verrerie et consommables plastique |
| NA.20 | Réactifs et produits chimiques courants |
| NA.21 | Réactifs biologiques (anticorps, enzymes, kits…) |
| NA.22 | Milieux de culture et suppléments |
| NA.23 | Acides nucléiques, oligonucléotides, vecteurs |
| NA.30 | Gaz et cryogénie |
| NA.40 | Petits équipements et instruments (≤ 5 000 €) |
| NA.41 | Équipements scientifiques (> 5 000 €) |
| NA.50 | Informatique et électronique |
| NA.60 | Animalerie et expérimentation animale |
| NA.70 | Prestations de service scientifique |
| NA.80 | Maintenance et réparation d'équipements |
| NA.90 | Fournitures de bureau et divers |

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

Simply **share the Google Sheet** with your lab members (Editor access). The API key is stored at the script level — all users benefit from it without any additional setup. Only one person (typically the lab manager) needs to set the key once.

> ⚠️ Each user's API calls are billed to the account associated with the API key. For a lab, a shared key managed by the PI or lab manager is the simplest approach.

---

## Troubleshooting

**"API key not set"** — Run 🧪 NACRES > ⚙️ Set API Key and paste your key.

**"Could not access file"** — Make sure the PDF is in Google Drive and accessible to your Google account. The script runs under the account of whoever opened the sheet.

**"Claude API error 401"** — Your API key is invalid or expired. Generate a new one at [console.anthropic.com](https://console.anthropic.com).

**PDF extraction incomplete** — Very large or scanned (image-only) PDFs may yield incomplete results. For scanned PDFs, use OCR first (Google Drive does this automatically when you open a PDF with Google Docs).

---

## Customisation

To adjust which model is used, edit the top of the script:

```javascript
const CLAUDE_MODEL = "claude-haiku-4-5-20251001";  // change to claude-sonnet-4-6 for higher accuracy
```

To add or refine NACRES codes in the prompt, edit the `buildSystemPrompt()` function.

---

## License

MIT — free to use, adapt, and share.

---

*Developed at [Institut Pierre-Gilles de Gennes](https://www.ipgg.espci.fr), Paris. Contributions welcome.*
