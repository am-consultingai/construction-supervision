/************************************************************
 * Sheet Utilities – Lists & Duplication (with protections)
 * OPTIMIZED VERSION with faster caching
 ************************************************************/

/** ====== Config (edit as needed) ====== **/
const BUDGET_SHEET = 'בקרת תקציב';
const BUDGET_HEADER_ROW = 4;
const TARGET_SHEET_NAME = 'בקרת תקציב';
const TARGET_START_ROW  = 5;
const TARGET_COL        = 1;
const TEMPLATE_NAME   = 'תבנית קבלן';
const CONTRACTORS_TENDER   = 'מכרזי קבלנים';
const EXCLUDED_SHEETS = new Set([
  'הוראות שימוש',
  'צרו קשר',
  'קבועים',
  'בקרת תקציב',
  CONTRACTORS_TENDER,
  TEMPLATE_NAME
]);
const debug = 0;

// ===== CACHE CONFIGURATION =====
const SUMMARY_CACHE_KEY = 'summary_payload_v1';
const SUMMARY_CACHE_TTL = 300; // 5 minutes

/** ====== Dialog Functions ====== **/
function openSummaryDialog() {

    const t = HtmlService.createTemplateFromFile('summary');
  t.payload = getBudgetSummaryAndContractorsCached_(); // or your data getter
  const html = t.evaluate().setWidth(520).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, ' '); // <- blank title

  /**const t = HtmlService.createTemplateFromFile('summary');
  t.payload = getBudgetSummaryAndContractorsCached_();
  const html = t.evaluate().setWidth(520).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'סיכום');**/
}

function openSummarySidebar() {
  const html = HtmlService.createHtmlOutputFromFile('summary')
    .setTitle('Summary');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Activate the sheet whose name equals `name` (handles RTL marks). */
function activateSheetByName(name) {
  const ss = SpreadsheetApp.getActive();
  const sanitize = s => String(s || '').replace(/[\u200E\u200F\u202A-\u202E]/g, '').trim();
  const targetName = sanitize(name);

  const sh = ss.getSheets().find(s => sanitize(s.getName()) === targetName);
  if (sh) {
    ss.setActiveSheet(sh);
    return { ok: true };
  } else {
    ss.toast('לא נמצא גליון בשם: ' + name);
    return { ok: false };
  }
}

/** ====== onOpen Menu ====== **/
function onOpen() {
  console.log('onOpen triggered');
  
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🧰 Utilities')
    .addItem('📑 צור קבלן חדש', 'duplicateFromTemplate')
    .addItem('🎯 הצג סיכום', 'openSummaryDialog');      
    /**.addItem('פתח בר', 'openSummarySidebar');**/

  if (debug) {
    menu.addSeparator()
      .addItem('📄 שכפל גליון נוכחי', 'duplicateActiveSheet')
      .addItem('🔄 עדכן בקרת תקציב', 'updateSheetList_')
      .addItem('⚙️ אפשר עדכון אוטומטי', 'installOnChangeTrigger')
      .addItem('🔥 Prewarm Cache', 'prewarmSummaryCache_')
      .addItem('❌ Clear Cache', 'invalidateSummaryCache_')
      .addItem('🔍 Debug Triggers', 'debugTriggers_')
      .addItem('▶️ Test Auto-Open Setup', 'checkAndPromptAutoOpen_');
  }
  menu.addToUi();

  // Prewarm cache in background (non-blocking)
  prewarmSummaryCache_();
  
  // Refresh sheet list
  updateSheetList_();
  
  // Show a toast to confirm onOpen is running
  SpreadsheetApp.getActive().toast('מערכת הופעלה', '🧰 Utilities', 2);
  
  // Check and prompt for auto-open trigger if needed
  try {
    checkAndPromptAutoOpen_();
  } catch (e) {
    //console.log('checkAndPromptAutoOpen_ failed:', e);
    //SpreadsheetApp.getActive().toast('שגיאה בטעינת אוטומטית: ' + e.message, '⚠️', 3);
  }
}

/** ====== Data Fetching (OPTIMIZED) ====== **/
function getBudgetSummaryAndContractors() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BUDGET_SHEET);
  if (!sh) throw new Error(`Missing sheet "${BUDGET_SHEET}"`);

  // Read summary cells in ONE batch call
  const ranges = sh.getRangeList(['A2', 'D1', 'B1', 'B3', 'F3']).getRanges();
  
  // Calculate completion percentage: B3 / B2 * 100
  const estCostValue = ranges[2].getValue(); // B2
  const actualCostValue = ranges[3].getValue(); // B3
  let completionPct = '';
  if (estCostValue && actualCostValue && estCostValue > 0) {
    const pct = Math.round((actualCostValue / estCostValue) * 100);
    completionPct = pct + '%';
  }
  
  const summary = {
    status:         ranges[0].getDisplayValue(), // A1
    totalBudget:    ranges[1].getDisplayValue(), // D1
    estCost:        ranges[2].getDisplayValue(), // B2
    completionPct:  completionPct,               // B3/B2 as percentage
    expectedEnd:    ranges[4].getDisplayValue()  // F3
  };

  // Read contractor list
  const startRow = BUDGET_HEADER_ROW + 1;
  const last = sh.getLastRow();
  const rows = [];

  if (last >= startRow) {
    const disp = sh.getRange(startRow, 1, last - startRow + 1, 2).getDisplayValues();

    let end = disp.length - 1;
    while (end >= 0 && disp[end][0] === '' && disp[end][1] === '') end--;

    for (let i = 0; i <= end; i++) {
      const name = (disp[i][0] || '').trim();
      const cost = disp[i][1] || '';
      if (!name) continue;
      rows.push({ name, value: cost });
    }
  }

  return { summary, rows };
}

/** ====== Cache Functions (IMPROVED) ====== **/
function getBudgetSummaryAndContractorsCached_() {
  const cache = CacheService.getDocumentCache();
  
  try {
    const hit = cache.get(SUMMARY_CACHE_KEY);
    if (hit) {
      console.log('Cache HIT');
      return JSON.parse(hit);
    }
  } catch (e) {
    console.warn('Cache read failed:', e);
  }
  
  console.log('Cache MISS - computing fresh');
  const fresh = getBudgetSummaryAndContractors();
  
  try {
    cache.put(SUMMARY_CACHE_KEY, JSON.stringify(fresh), SUMMARY_CACHE_TTL);
  } catch (e) {
    console.warn('Cache write failed:', e);
  }
  
  return fresh;
}

function invalidateSummaryCache_() {
  try {
    CacheService.getDocumentCache().remove(SUMMARY_CACHE_KEY);
    console.log('Cache invalidated');
  } catch (e) {
    console.warn('Cache invalidation failed:', e);
  }
}

function prewarmSummaryCache_() {
  try {
    const cache = CacheService.getDocumentCache();
    if (!cache.get(SUMMARY_CACHE_KEY)) {
      console.log('Prewarming cache...');
      const fresh = getBudgetSummaryAndContractors();
      cache.put(SUMMARY_CACHE_KEY, JSON.stringify(fresh), SUMMARY_CACHE_TTL);
    }
  } catch (e) {
    console.log('Prewarm failed', e);
  }
}

function onEdit(e) {
  const sh = e && e.range && e.range.getSheet();
  if (!sh) return;
  
  if (sh.getName() === BUDGET_SHEET) {
    const row = e.range.getRow();
    const col = e.range.getColumn();
    
    const isRelevant = (
      (row === 1 && (col === 1 || col === 4)) ||  // A1 or D1
      (row === 2 && col === 2) ||                  // B2
      (row === 3 && (col === 2 || col === 6)) ||   // B3 or F3
      (row >= BUDGET_HEADER_ROW + 1 && col <= 2)   // A5+ or B5+
    );
    
    if (isRelevant) {
      invalidateSummaryCache_();
    }
  }
}

/** ====== Sheet List Management ====== **/
/** ====== Sheet List Management (with clickable links) ====== **/
function updateSheetList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!target) throw new Error('Target sheet not found: ' + TARGET_SHEET_NAME);

  // Collect tab names + gids, excluding configured tabs
  const sheets = ss.getSheets()
    .filter(sh => !EXCLUDED_SHEETS.has(sh.getName()))
    .map(sh => ({ name: sh.getName(), gid: sh.getSheetId() }));

  // Ensure the target has enough rows to write
  const needed = Math.max(1, sheets.length);
  const firstWriteRow = TARGET_START_ROW;
  const lastWriteRow = firstWriteRow + needed - 1;
  if (target.getMaxRows() < lastWriteRow) {
    target.insertRowsAfter(target.getMaxRows(), lastWriteRow - target.getMaxRows());
  }

  // Clear old list (contents only) from A5 down
  const clearRows = Math.max(1, target.getMaxRows() - TARGET_START_ROW + 1);
  target.getRange(TARGET_START_ROW, TARGET_COL, clearRows, 1).clearContent();

  if (!sheets.length) {
    console.log('Updated sheet list (0 items)');
    return;
  }

  // Build RichText hyperlinks to each tab (use full URL for reliability)
  const baseUrl = ss.getUrl().split('#')[0];
  const rich = sheets.map(({ name, gid }) => {
    const url = `${baseUrl}#gid=${gid}`;
    return [SpreadsheetApp.newRichTextValue()
      .setText(name)
      .setLinkUrl(url)
      .build()];
  });

  // Write clickable names
  target.getRange(TARGET_START_ROW, TARGET_COL, rich.length, 1).setRichTextValues(rich);

  console.log(`Updated sheet list (${sheets.length} items) with hyperlinks`);
}


function installOnChangeTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'updateSheetList_' &&
    t.getEventType() === ScriptApp.EventType.ON_CHANGE &&
    t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
    (typeof t.getTriggerSourceId === 'function' ? t.getTriggerSourceId() === ssId : true)
  );

  if (exists) {
    if(debug) {
      SpreadsheetApp.getUi().alert('On-change trigger already installed');
    }
    return;
  }

  ScriptApp.newTrigger('updateSheetList_')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert('On-change trigger installed');
}

/** Check if auto-open trigger exists, prompt user if not */
function checkAndPromptAutoOpen_() {
  console.log('Checking auto-open trigger...');
  
  // Check if trigger already exists
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'autoOpenSummary_' &&
    t.getEventType() === ScriptApp.EventType.ON_OPEN
  );
  
  console.log('Trigger exists:', exists);
  
  if (exists) {
    console.log('Auto-open trigger already installed');
    return;
  }
  
  // Check if user has been prompted before
  const props = PropertiesService.getDocumentProperties();
  const prompted = props.getProperty('autoOpenPrompted');
  
  console.log('Previously prompted:', prompted);
  
  if (prompted === 'yes') {
    console.log('User already responded to prompt');
    return;
  }
  
  // Prompt user
  console.log('Prompting user...');
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'פתיחה אוטומטית של סיכום',
    'האם תרצה שחלון הסיכום יפתח אוטומטית בכל פעם שתפתח את הקובץ?\n\n(ניתן לפתוח אותו תמיד דרך התפריט)',
    ui.ButtonSet.YES_NO
  );
  
  // Mark as prompted regardless of answer
  props.setProperty('autoOpenPrompted', 'yes');
  console.log('User responded:', response);
  
  if (response === ui.Button.YES) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ScriptApp.newTrigger('autoOpenSummary_')
        .forSpreadsheet(ss)
        .onOpen()
        .create();
      console.log('Trigger created successfully');
      ui.alert('✅ הופעל בהצלחה!\n\nהסיכום יפתח אוטומטית מהפעם הבאה.');
    } catch (e) {
      console.error('Failed to install trigger:', e);
      ui.alert('⚠️ לא הצלחנו להפעיל.\nאפשר לפתוח את הסיכום תמיד דרך התפריט.');
    }
  } else {
    console.log('User declined auto-open');
  }
}

/** Install trigger to auto-open summary on spreadsheet open (silent, once only) */
function installAutoOpenTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  
  // Check if already installed
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'autoOpenSummary_' &&
    t.getEventType() === ScriptApp.EventType.ON_OPEN &&
    (typeof t.getTriggerSourceId === 'function' ? t.getTriggerSourceId() === ssId : true)
  );
  
  if (exists) {
    return; // Already installed, do nothing
  }
  
  // Create installable ON_OPEN trigger
  try {
    ScriptApp.newTrigger('autoOpenSummary_')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
    console.log('Auto-open trigger installed successfully');
  } catch (e) {
    console.log('Could not install auto-open trigger:', e);
  }
}

/** Trigger function to auto-open summary (installable trigger) */
function autoOpenSummary_() {
  try {
    openSummaryDialog();
  } catch (e) {
    console.log('Auto-open failed:', e);
  }
}

/** ====== Duplication Functions ====== **/
function duplicateActiveSheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getActiveSheet();

  const suggested = src.getName() + ' Copy';
  const inputName = promptName_(suggested);
  if (!inputName) return;

  const newName = uniqueName_(ss, inputName);
  const dst = src.copyTo(ss).setName(newName);
  ss.setActiveSheet(dst);

  SpreadsheetApp.flush();

  copyRangeProtections_(src, dst);
  //copySheetProtection_(src, dst);
  updateSheetList_();
  console.log(`Duplicate complete: "${src.getName()}" → "${dst.getName()}"`);
}

function duplicateFromTemplate() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(TEMPLATE_NAME);
  if (!src) throw new Error('Template sheet not found: ' + TEMPLATE_NAME);

  const inputName = promptName_('קבלן חדש');
  if (!inputName) return;

  const newName = uniqueName_(ss, inputName);
  const dst = src.copyTo(ss).setName(newName);
  ss.setActiveSheet(dst);

  const K = 11;
  dst.getRange(1, K).setValue(newName);
  dst.getRange(2, K).setValue(1);
  dst.hideColumns(K);

  SpreadsheetApp.flush();
  installOnChangeTrigger();
  copyRangeProtections_(src, dst);
  //copySheetProtection_(src, dst);
  updateSheetList_();
  
  showCompletionModal_(newName);

  console.log(`Template duplicated: "${src.getName()}" → "${dst.getName()}"`);
}


/** Show blocking modal; on close, activate CONTRACTORS_TENDER. */
/** Show blocking modal; button explicitly redirects to CONTRACTORS_TENDER and then closes. */
function showCompletionModal_(newName) {
  const targetNameJs = JSON.stringify(CONTRACTORS_TENDER);

  const html = HtmlService.createHtmlOutput(
    `<div style="
        font:14px system-ui; line-height:1.5;
        direction:rtl; text-align:right; padding:12px 20px 14px 10px;">
       <div style="font-weight:600; margin-bottom:6px">הפעולה הושלמה</div>
       ✅ הגליון "<b>${newName}</b>" נוצר.
       <div style="margin-top:12px; color:#666; font-size:13px">
       עבור אל הגליון ${CONTRACTORS_TENDER} כדי לקבוע עלות
       </div>
       <div style="margin-top:14px">
         <button id="go" style="
           padding:8px 14px; border:0; border-radius:6px;
           background:#0b57d0; color:#fff; cursor:pointer; font-weight:600;">
           עבור לגליון ${CONTRACTORS_TENDER}
         </button>
       </div>
       <script>
         document.getElementById('go').addEventListener('click', function () {
           google.script.run
             .withSuccessHandler(function(){ google.script.host.close(); })
             .activateSheetByName(${targetNameJs});
         });
       </script>
     </div>`
  ).setWidth(420).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(html, 'קבלן חדש נוצר');
}






/** ====== Protection Copy Helpers ====== **/
function copyRangeProtections_(src, dst) {
  const protections = src.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  console.log(`Copying ${protections.length} range protections`);

  protections.forEach((p, i) => {
    const a1 = p.getRange().getA1Notation();
    try {
      const newProt = dst.getRange(a1).protect();
      newProt.setDescription(p.getDescription() || '');
      newProt.setWarningOnly(p.isWarningOnly());

      if (!p.isWarningOnly()) {
        try { newProt.setDomainEdit(p.canDomainEdit()); }
        catch (e) { console.warn(`[${a1}] setDomainEdit skipped`); }

        const emails = p.getEditors().map(u => u.getEmail()).filter(Boolean);
        if (emails.length) {
          try { newProt.addEditors(emails); }
          catch (e) { console.warn(`[${a1}] addEditors failed`); }
        }

        try {
          const me = Session.getEffectiveUser().getEmail();
          const current = newProt.getEditors().map(u => u.getEmail());
          const toRemove = current.filter(e => !emails.includes(e) && e !== me);
          if (toRemove.length) newProt.removeEditors(toRemove);
        } catch (e) {
          console.warn(`[${a1}] cleanup skipped`);
        }
      }
    } catch (err) {
      console.error(`Failed to copy ${a1}: ${err}`);
    }
  });
}

function copySheetProtection_(src, dst) {
  const sProt = src.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (!sProt) return;

  try {
    const p2 = dst.protect()
      .setDescription(sProt.getDescription() || '')
      .setWarningOnly(sProt.isWarningOnly());

    if (!sProt.isWarningOnly()) {
      try { p2.setDomainEdit(sProt.canDomainEdit()); }
      catch (e) { console.warn('setDomainEdit skipped'); }

      const unprot = sProt.getUnprotectedRanges().map(r => dst.getRange(r.getA1Notation()));
      if (unprot.length) p2.setUnprotectedRanges(unprot);

      const emails = sProt.getEditors().map(u => u.getEmail()).filter(Boolean);
      if (emails.length) {
        try { p2.addEditors(emails); }
        catch (e) { console.warn('addEditors failed'); }
      }

      try {
        const me = Session.getEffectiveUser().getEmail();
        const current = p2.getEditors().map(u => u.getEmail());
        const toRemove = current.filter(e => !emails.includes(e) && e !== me);
        if (toRemove.length) p2.removeEditors(toRemove);
      } catch (e) {
        console.warn('cleanup skipped');
      }
    }
  } catch (err) {
    console.error(`Failed to copy sheet protection: ${err}`);
  }
}

/** ====== Utility Helpers ====== **/
function promptName_(suggestion) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('תן שם לקבלן החדש', suggestion, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  return (res.getResponseText() || '').trim();
}

function uniqueName_(ss, desired) {
  let name = desired, i = 2;
  while (ss.getSheetByName(name)) name = `${desired} (${i++})`;
  return name;
}

function lastNonEmptyRow_(colValues) {
  for (let i = colValues.length - 1; i >= 0; i--) {
    if (String(colValues[i][0]).trim() !== '') return i + 1;
  }
  return 0;
}

function sanitizeRtl_(s) {
  return String(s || '').replace(/[\u200E\u200F\u202A-\u202E]/g, '').trim();
}

/** Debug helper: show all triggers */
function debugTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  let info = `Total triggers: ${triggers.length}\n\n`;
  
  triggers.forEach((t, i) => {
    info += `#${i+1}: ${t.getHandlerFunction()}\n`;
    info += `  Type: ${t.getEventType()}\n`;
    info += `  Source: ${t.getTriggerSource()}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert('Triggers Debug', info, SpreadsheetApp.getUi().ButtonSet.OK);
  console.log(info);
}