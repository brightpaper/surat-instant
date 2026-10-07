/**
 * Bright Paper · Dashboard "New Client" backend (Google Apps Script)
 * Form ki entry "Enquiry Sheet" me nayi row ke roop me add karta hai.
 * FMS khud =index('Enquiry Sheet'!A2:G) se data utha leta hai — FMS me kabhi seedha mat likho.
 *
 * Isi project me aapka Fill Form wala doGet() hai, isliye yaha doGet NAHI hai,
 * aur saare naam "nc" se shuru hain taaki kisi aur file se takraaye nahi.
 *
 * Code badalne ke baad: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy
 */

var NC_SHEET = 'Enquiry Sheet';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return ncJson_(data.action === 'update' ? ncUpdateCell_(data) : ncAddClient_(data));
  } catch (err) {
    return ncJson_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function ncAddClient_(data) {
  var company = ncClean_(data.company), owner = ncClean_(data.owner), mobile = String(data.mobile || '').replace(/\D/g, '');
  var location = ncClean_(data.location), sales = ncClean_(data.sales);
  if (!company) throw new Error('Company Name bharna zaroori hai.');
  if (!owner) throw new Error('Owner Name bharna zaroori hai.');
  if (!/^\d{10}$/.test(mobile)) throw new Error('Mobile Number 10 digit ka hona chahiye.');

  var sh = SpreadsheetApp.getActive().getSheetByName(NC_SHEET);
  if (!sh) throw new Error('Sheet "' + NC_SHEET + '" nahi mili.');

  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(ncNorm_);
  function col(names) { for (var i = 0; i < head.length; i++) if (names.indexOf(head[i]) !== -1) return i; return -1; }
  var c = {
    ts: col(['timestamp']),
    company: col(['companyname']),
    owner: col(['ownername']),
    mobile: col(['mobilenumber']),
    location: col(['location']),
    sales: col(['salesperson']),
    platform: col(['platform'])
  };
  if (c.ts !== 0 || c.company === -1 || c.owner === -1 || c.mobile === -1) throw new Error('Enquiry Sheet ke headers match nahi hue — kuch nahi likha.');

  // last row jisme Timestamp hai (Lead No column ARRAYFORMULA hai, use nahi chhedna)
  var ts = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  var last = 1;
  for (var i = ts.length - 1; i >= 1; i--) if (ts[i][0] !== '' && ts[i][0] !== null) { last = i + 1; break; }
  var newRow = last + 1;
  if (newRow > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);

  var dupe = sh.getRange(2, c.mobile + 1, Math.max(last - 1, 1), 1).getDisplayValues()
    .some(function (r) { return String(r[0]).replace(/\D/g, '').slice(-10) === mobile; });

  // sirf format copy (formulas nahi)
  if (last >= 2) sh.getRange(last, 1, 1, lastCol).copyTo(sh.getRange(newRow, 1, 1, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  sh.getRange(newRow, c.ts + 1).setValue(new Date());
  sh.getRange(newRow, c.company + 1).setValue(company);
  sh.getRange(newRow, c.owner + 1).setValue(owner);
  sh.getRange(newRow, c.mobile + 1).setNumberFormat('@').setValue(mobile);
  if (c.location >= 0) sh.getRange(newRow, c.location + 1).setValue(location);
  if (c.sales >= 0) sh.getRange(newRow, c.sales + 1).setValue(sales);
  if (c.platform >= 0) sh.getRange(newRow, c.platform + 1).setValue('Manual');
  SpreadsheetApp.flush();

  var leadNo = sh.getRange(newRow, 2).getDisplayValue();
  try { ncLinkNewLead_(leadNo); } catch (err) { Logger.log('ncLinkNewLead_: ' + err.message); }
  return { ok: true, row: newRow, leadNo: leadNo, duplicate: dupe };
}

/**
 * Nayi lead ka FMS Links cell turant bharo — wahi jo ProductFormCode.gs ka refreshProductFormLinks() is row ke liye karta
 * (Fill Form link, ya META_ONLY ho to "Manual"). Pura refresh (saari rows) bahut slow tha; baaki kaam 5 min trigger kar lega.
 */
function ncLinkNewLead_(ld) {
  if (!ld || typeof pf_linkFormula_ !== 'function') return;
  pf_resolveCols_();
  var C = PF_CFG, fms = pf_getFmsSheet_(), r = pf_findFmsRowByLd_(fms, ld), url = pf_webAppUrl_();
  if (!r || !url) return;
  var cell = fms.getRange(r, C.FMS_COL_LINK);
  if (C.META_ONLY) cell.setValue(C.MANUAL_TEXT);
  else cell.setFormula(pf_linkFormula_(url, ld, C.LINK_TEXT));
}

/**
 * Dashboard ki row ka dropdown save karta hai (FMS ya MKT).
 * Sirf us cell me likhta hai jisme SHEET me dropdown (data validation list) laga hai, aur sirf usi list ki value.
 * Isliye columns aage-peeche karo to bhi sahi chalta hai — jo sheet me dropdown hai wahi dashboard me.
 * Row Lead No se dhundhi jaati hai (FMS: Lead No column, MKT: Links ke "?ld=" se).
 * Likhne ke baad ProductFormCode.gs ka onEdit trigger (pfOnEditFms) chalta hai — bilkul sheet me haath se badalne jaisa.
 */
var NC_TABS = { fms: 'FMS', mkt: 'MKT' };

function ncUpdateCell_(data) {
  var name = NC_TABS[data.tab];
  if (!name) throw new Error('Galat tab.');
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" sheet nahi mili.');
  var value = ncClean_(data.value), ld = ncClean_(data.ld), remark = ncClean_(data.remark), col = Number(data.col);
  if (!ld) throw new Error('Lead No nahi mila.');
  if (!col) throw new Error('Column nahi mila.');
  if (data.tab === 'fms' && col <= 7) throw new Error('FMS ke A:G Enquiry Sheet ke formula se aate hain — edit nahi ho sakte.');

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  var top = sh.getRange(1, 1, Math.min(lastRow, 30), lastCol).getDisplayValues(), hr = -1;
  for (var r = 0; r < top.length; r++) {
    var n = top[r].map(ncNorm_);
    if (n.indexOf('companyname') !== -1 && (n.indexOf('links') !== -1 || n.indexOf('link') !== -1)) { hr = r; break; }
  }
  if (hr === -1) throw new Error(name + ' ki header row nahi mili.');
  var head = top[hr].map(ncNorm_);
  var row = ncFindRow_(sh, head, hr + 2, lastRow, ld);
  if (!row) throw new Error(ld + ' ' + name + ' me nahi mila.');

  var rng = sh.getRange(row, col), dv = rng.getDataValidation();
  if (!dv || dv.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    throw new Error(ld + ': is cell me sheet me dropdown nahi hai — edit nahi ho sakta.');
  }
  var allowed = (dv.getCriteriaValues()[0] || []).map(String);
  if (value && allowed.indexOf(value) === -1) throw new Error('Galat value: ' + value);

  // Decision "No" → Remark zaroori (Decision ke right ka pehla Remark, agle step ke Planned se pehle)
  var remarkCol = 0;
  if (head[col - 1] === 'decision') {
    for (var x = col; x < head.length && head[x] !== 'planned'; x++) if (head[x] === 'remark') { remarkCol = x + 1; break; }
    if (value === 'No' && !remark) throw new Error('Decision "No" ke liye Remark likhna zaroori hai.');
  }
  if (remarkCol && remark) sh.getRange(row, remarkCol).setValue(remark);

  var old = rng.getDisplayValue();
  rng.setValue(value);
  SpreadsheetApp.flush();
  // script ke likhne se onEdit trigger nahi chalta → khud chalao (Status/Actual, agle step ke dropdown, MKT me bhejna…)
  if (typeof pfOnEditFms === 'function') {
    pfOnEditFms({ range: rng, value: value, oldValue: old, source: SpreadsheetApp.getActive() });
  }
  return { ok: true, row: row };
}

// LD wali row: "Lead No" column ho to usse, warna kisi bhi formula ke "?ld=LD-xxx" se (MKT ke Links)
function ncFindRow_(sh, head, firstRow, lastRow, ld) {
  var n = lastRow - firstRow + 1;
  if (n < 1) return 0;
  var target = ld.toUpperCase(), cLead = head.indexOf('leadno');
  if (cLead !== -1) {
    var v = sh.getRange(firstRow, cLead + 1, n, 1).getDisplayValues();
    for (var i = 0; i < n; i++) if (ncClean_(v[i][0]).toUpperCase() === target) return firstRow + i;
    return 0;
  }
  var f = sh.getRange(firstRow, 1, n, head.length).getFormulas();
  var re = new RegExp('[?&]ld=' + target.replace(/[^A-Z0-9-]/g, '') + '(?![0-9])', 'i');
  for (var r = 0; r < n; r++) for (var c = 0; c < f[r].length; c++) if (f[r][c] && re.test(f[r][c])) return firstRow + r;
  return 0;
}

function ncClean_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function ncNorm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function ncJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
