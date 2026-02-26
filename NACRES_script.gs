// ============================================================
//  NACRES CODE ASSISTANT — Google Apps Script
//  Paste this entire file into Tools > Script editor
//  Then: Run > onOpen once to create the menu
// ============================================================

// ---- CONFIGURATION -----------------------------------------
const CLAUDE_MODEL = "claude-haiku-4-5-20251001";   // cheap & fast, sufficient for this task
const MAX_TOKENS   = 4096;

// Sheet names
const SHEET_INPUT  = "Input";
const SHEET_OUTPUT = "Résultats";

// Column layout in Output sheet
const COL_REF         = 1;  // A - Supplier ref / catalog number
const COL_DESCRIPTION = 2;  // B - Item description
const COL_QTY         = 3;  // C - Quantity
const COL_UNIT_PRICE  = 4;  // D - Unit price
const COL_NACRES_SUGG = 5;  // E - Suggested NACRES (by Claude)
const COL_NACRES_CONF = 6;  // F - Confidence (High / Medium / Low)
const COL_NACRES_VALID= 7;  // G - ✏️ Validated NACRES (editable by user)
const COL_NOTES       = 8;  // H - Claude's justification


// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧪 NACRES")
    .addItem("⚙️  Set API Key",          "setApiKey")
    .addSeparator()
    .addItem("📄 Process PDF (Drive)",   "processPdfFromDrive")
    .addItem("📋 Process text list",     "processTextList")
    .addSeparator()
    .addItem("🗑️  Clear results",        "clearResults")
    .addToUi();
}


// ============================================================
//  API KEY MANAGEMENT  (stored in Script Properties — never in cells)
// ============================================================
function setApiKey() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "Claude API Key",
    "Paste your Anthropic API key (starts with sk-ant-).\n" +
    "It will be stored securely in Script Properties.",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key.startsWith("sk-")) {
    ui.alert("❌ This doesn't look like a valid API key.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("CLAUDE_API_KEY", key);
  ui.alert("✅ API key saved.");
}

function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!key) throw new Error("API key not set. Use menu NACRES > Set API Key.");
  return key;
}


// ============================================================
//  PDF INPUT — pick a file from Google Drive
// ============================================================
function processPdfFromDrive() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "PDF from Google Drive",
    "Paste the Google Drive file URL or file ID of the supplier quotation PDF:",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const input = res.getResponseText().trim();
  if (!input) return;

  // Extract file ID from URL or use directly
  let fileId = input;
  const match = input.match(/[-\w]{25,}/);
  if (match) fileId = match[0];

  let file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch(e) {
    ui.alert("❌ Could not access file. Check the ID/URL and sharing permissions.\n" + e.message);
    return;
  }

  if (file.getMimeType() !== "application/pdf") {
    ui.alert("❌ File does not appear to be a PDF.");
    return;
  }

  const pdfBytes = file.getBlob().getBytes();
  const pdfBase64 = Utilities.base64Encode(pdfBytes);

  ui.alert("⏳ Sending PDF to Claude… This may take 10–20 seconds. Click OK and wait.");

  try {
    const items = extractItemsFromPdf(pdfBase64);
    writeResultsToSheet(items);
    SpreadsheetApp.getActive().getSheetByName(SHEET_OUTPUT).activate();
    ui.alert("✅ Done! " + items.length + " items found. Check the '" + SHEET_OUTPUT + "' sheet.\n\nPlease review and correct column G (Validated NACRES) if needed.");
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}


// ============================================================
//  TEXT LIST INPUT
// ============================================================
function processTextList() {
  const ui  = SpreadsheetApp.getUi();

  // Check if Input sheet has data
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let inSheet = ss.getSheetByName(SHEET_INPUT);

  if (!inSheet) {
    // Create it with instructions
    inSheet = ss.insertSheet(SHEET_INPUT);
    inSheet.getRange("A1").setValue("Paste your item list below (one item per row, or free text in A1):");
    inSheet.getRange("A1").setFontWeight("bold");
    ui.alert("A '" + SHEET_INPUT + "' sheet has been created.\n\nPaste your items there (one per row in column A), then run this again.");
    return;
  }

  const data = inSheet.getDataRange().getValues();
  const lines = data.flat().filter(v => v.toString().trim() !== "").join("\n");

  if (!lines) {
    ui.alert("The '" + SHEET_INPUT + "' sheet is empty. Paste your items there first.");
    return;
  }

  ui.alert("⏳ Sending list to Claude… Click OK and wait ~10 seconds.");

  try {
    const items = extractItemsFromText(lines);
    writeResultsToSheet(items);
    SpreadsheetApp.getActive().getSheetByName(SHEET_OUTPUT).activate();
    ui.alert("✅ Done! " + items.length + " items found. Check the '" + SHEET_OUTPUT + "' sheet.");
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}


// ============================================================
//  CLAUDE API CALLS
// ============================================================

function extractItemsFromPdf(pdfBase64) {
  const key = getApiKey();

  const systemPrompt = buildSystemPrompt();

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: MAX_TOKENS,
    system: systemPrompt,
    messages: [
      {
        role: "user",
        content: [
          {
            type: "document",
            source: {
              type: "base64",
              media_type: "application/pdf",
              data: pdfBase64
            }
          },
          {
            type: "text",
            text: "Voici un devis fournisseur. Extrais tous les articles et assigne un code NACRES à chacun. Réponds uniquement en JSON valide, sans texte avant ou après."
          }
        ]
      }
    ]
  };

  return callClaudeAndParse(key, payload);
}

function extractItemsFromText(text) {
  const key = getApiKey();

  const systemPrompt = buildSystemPrompt();

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: MAX_TOKENS,
    system: systemPrompt,
    messages: [
      {
        role: "user",
        content: "Voici une liste d'articles de laboratoire :\n\n" + text +
                 "\n\nAssigne un code NACRES à chacun. Réponds uniquement en JSON valide, sans texte avant ou après."
      }
    ]
  };

  return callClaudeAndParse(key, payload);
}

function buildSystemPrompt() {
  return `Tu es un expert en achats de laboratoire académique français. 
Tu dois identifier le code NACRES (système de classification des achats de la recherche publique française) le plus approprié pour chaque article.

Les codes NACRES les plus courants en laboratoire :
- NA.10 : Verrerie et consommables plastique (tubes, pipettes, boîtes de Petri, flacons…)
- NA.20 : Réactifs et produits chimiques courants
- NA.21 : Réactifs biologiques (anticorps, enzymes, kits ELISA…)
- NA.22 : Milieux de culture et suppléments
- NA.23 : Acides nucléiques, oligonucléotides, vecteurs
- NA.30 : Gaz et cryogénie
- NA.40 : Petits équipements et instruments (<= 5000 €)
- NA.41 : Équipements scientifiques (> 5000 €)
- NA.50 : Informatique et électronique
- NA.60 : Animalerie et expérimentation animale
- NA.70 : Prestations de service scientifique
- NA.80 : Maintenance et réparation d'équipements
- NA.90 : Fournitures de bureau et divers

Pour chaque article, retourne un objet JSON avec exactement ces champs :
{
  "ref": "référence catalogue ou chaîne vide",
  "description": "description de l'article",
  "qty": "quantité ou chaîne vide",
  "unit_price": "prix unitaire ou chaîne vide",
  "nacres_code": "NA.XX",
  "confidence": "High | Medium | Low",
  "notes": "justification courte en français (1 ligne)"
}

Retourne un tableau JSON de ces objets. Rien d'autre.`;
}

function callClaudeAndParse(apiKey, payload) {
  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code !== 200) {
    throw new Error("Claude API error " + code + ": " + body);
  }

  const json = JSON.parse(body);
  const text = json.content[0].text.trim();

  // Strip markdown fences if present
  const clean = text.replace(/^```json\s*/i, "").replace(/```\s*$/i, "").trim();

  let items;
  try {
    items = JSON.parse(clean);
  } catch(e) {
    throw new Error("Could not parse Claude's response as JSON.\n\nRaw response:\n" + text.substring(0, 500));
  }

  if (!Array.isArray(items)) {
    throw new Error("Expected a JSON array from Claude, got: " + typeof items);
  }

  return items;
}


// ============================================================
//  WRITE RESULTS TO OUTPUT SHEET
// ============================================================
function writeResultsToSheet(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_OUTPUT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OUTPUT);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // Header row
  const headers = [
    "Référence", "Description", "Qté", "Prix unitaire",
    "NACRES suggéré", "Confiance", "✏️ NACRES validé", "Justification"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a5276");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);

  if (items.length === 0) return;

  // Data rows
  const rows = items.map(item => [
    item.ref         || "",
    item.description || "",
    item.qty         || "",
    item.unit_price  || "",
    item.nacres_code || "",
    item.confidence  || "",
    item.nacres_code || "",   // pre-fill validated = suggested, user corrects if needed
    item.notes       || ""
  ]);

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Color-code confidence
  for (let i = 0; i < rows.length; i++) {
    const conf  = rows[i][5];  // confidence column
    const cell  = sheet.getRange(i + 2, COL_NACRES_CONF);
    if      (conf === "High")   { cell.setBackground("#d4efdf"); cell.setFontColor("#1e8449"); }
    else if (conf === "Medium") { cell.setBackground("#fef9e7"); cell.setFontColor("#d4ac0d"); }
    else if (conf === "Low")    { cell.setBackground("#fadbd8"); cell.setFontColor("#c0392b"); }
  }

  // Highlight editable column G
  sheet.getRange(2, COL_NACRES_VALID, rows.length, 1)
    .setBackground("#eaf4fb")
    .setFontWeight("bold");

  // Column widths
  sheet.setColumnWidth(COL_REF, 120);
  sheet.setColumnWidth(COL_DESCRIPTION, 280);
  sheet.setColumnWidth(COL_QTY, 60);
  sheet.setColumnWidth(COL_UNIT_PRICE, 100);
  sheet.setColumnWidth(COL_NACRES_SUGG, 110);
  sheet.setColumnWidth(COL_NACRES_CONF, 90);
  sheet.setColumnWidth(COL_NACRES_VALID, 120);
  sheet.setColumnWidth(COL_NOTES, 320);

  sheet.setFrozenRows(1);

  // Protect suggested column (informational only)
  // (optional — skip if you want to keep it simple)
}


// ============================================================
//  CLEAR RESULTS
// ============================================================
function clearResults() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_OUTPUT);
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
    SpreadsheetApp.getUi().alert("✅ Results sheet cleared.");
  } else {
    SpreadsheetApp.getUi().alert("No results sheet found.");
  }
}
