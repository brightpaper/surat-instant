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
 * Dashboard ki row se Verify / Sales Person (STEP2) dropdown save karta hai — sirf FMS ke in 2 columns me.
 * Row Lead No se dhundhi jaati hai. A:G kabhi nahi likhta (wo Enquiry Sheet ke formula se aata hai).
 */
var NC_EDIT = {
  verify: { head: 'verify', values: ['Yes', 'No'] },
  sales:  { head: 'salesperson', values: [] }   // list ProductFormCode.gs ke PF_CFG.SALES_PERSONS se aati hai
};

function ncUpdateCell_(data) {
  if (data.tab === 'mkt') return ncUpdateMkt_(data);
  var f = NC_EDIT[data.field];
  if (data.field === 'sales') f.values = (typeof PF_CFG !== 'undefined' && PF_CFG.SALES_PERSONS) || [];
  if (!f) throw new Error('Ye column edit nahi ho sakta.');
  var value = ncClean_(data.value);
  if (value && f.values.indexOf(value) === -1) throw new Error('Galat value: ' + value);
  var ld = ncClean_(data.ld);
  if (!ld) throw new Error('Lead No nahi mila.');

  var sh = SpreadsheetApp.getActive().getSheetByName('FMS');
  // sirf upar ke 30 rows (header) + Lead No column padho — pura sheet nahi (fast)
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  var top = sh.getRange(1, 1, Math.min(lastRow, 30), lastCol).getDisplayValues();
  var hr = -1;
  for (var r = 0; r < top.length; r++) {
    var n = top[r].map(ncNorm_);
    if (n.indexOf('links') !== -1 && n.indexOf('leadno') !== -1) { hr = r; break; }
  }
  if (hr === -1) throw new Error('FMS header row nahi mili.');
  var head = top[hr].map(ncNorm_), cLead = head.indexOf('leadno'), cCol = head.lastIndexOf(f.head);
  if (cCol < 7) throw new Error('"' + f.head + '" column nahi mila.');   // A:G = formula area, wahan nahi likhna

  var lds = sh.getRange(hr + 2, cLead + 1, Math.max(lastRow - hr - 1, 1), 1).getDisplayValues();
  for (var k = 0; k < lds.length; k++) {
    var i = hr + 1 + k;
    if (ncClean_(lds[k][0]) === ld) {
      sh.getRange(i + 1, cCol + 1).setValue(value);
      // ProductFormCode.gs wala flow turant chalao: Verify Yes → Sales Person dropdown, Yes + Sales Person → MKT
      // (sheet me haath se badalne pe ye pfOnEditFms karta hai; script ke likhne pe onEdit nahi chalta)
      if (typeof pf_handleAssign_ === 'function') {
        pf_resolveCols_();
        pf_handleAssign_(sh, i + 1, i + 1);
      }
      return { ok: true, row: i + 1 };
    }
  }
  throw new Error(ld + ' FMS me nahi mila.');
}

/**
 * MKT sheet edit. Row LD se (MKT me Lead No column nahi, Links ke "ld=" se — ProductFormCode.gs ka pf_mktLdMap_).
 *  field 'verify'           : Verify Yes / No  (sirf jab form wale step ka Actual aa gaya ho)
 *  field 'cell' + col (1-based): aage ke steps ka "Status" (Done) ya "Decision" (Yes / No) — sirf jab us step ka Planned aa gaya ho
 * Likhne ke baad ProductFormCode.gs ka pf_onEditMkt_ chalta hai — bilkul sheet me haath se badalne jaisa:
 *  Status Done → us step ka Actual, Decision → Status Done + Actual, khaali → dono khaali, agle step ke dropdowns.
 */
function ncUpdateMkt_(data) {
  if (typeof pf_mktCols_ !== 'function' || typeof pf_onEditMkt_ !== 'function') throw new Error('ProductFormCode.gs nahi mila.');
  var value = ncClean_(data.value), ld = ncClean_(data.ld), remark = ncClean_(data.remark);
  if (!ld) throw new Error('Lead No nahi mila.');

  var mk = pf_mktCols_(), h = mk.hdr;   // MKT ke liye FMS/Enquiry headers padhne ki zaroorat nahi (fast)
  var row = pf_findMktRow_(mk, ld);
  if (!row) throw new Error(ld + ' MKT me nahi mila.');

  var col, allowed, remarkCol = 0;
  if (data.field === 'verify') {
    if (!mk.verify) throw new Error('MKT me "Verify" column nahi mila.');
    if (mk.actual && !String(mk.sh.getRange(row, mk.actual).getDisplayValue()).trim()) {
      throw new Error(ld + ': pehle MKT form bharo (Actual khaali hai).');
    }
    col = mk.verify; allowed = ['Yes', 'No'];
  } else if (data.field === 'cell') {
    col = Number(data.col);
    var name = h[col - 1];
    if (!col || col === mk.status || (name !== 'status' && name !== 'decision')) throw new Error('Ye column edit nahi ho sakta.');
    var planned = 0;
    for (var j = col - 2; j >= 0; j--) if (h[j] === 'planned') { planned = j + 1; break; }
    if (!planned || planned < mk.status) throw new Error('Ye column edit nahi ho sakta.');
    if (!String(mk.sh.getRange(row, planned).getDisplayValue()).trim()) throw new Error(ld + ': is step ka Planned abhi nahi aaya.');
    allowed = name === 'decision' ? ['Yes', 'No'] : [PF_CFG.DONE_TEXT];
    if (name === 'decision') {
      // Decision ke right ka pehla "Remark" (agle step ke Planned se pehle)
      for (var x = col; x < h.length && h[x] !== 'planned'; x++) if (h[x] === 'remark') { remarkCol = x + 1; break; }
      if (value === 'No' && !remark) throw new Error('Decision "No" ke liye Remark likhna zaroori hai.');
      if (remark && !remarkCol) throw new Error('Decision ke baad "Remark" column nahi mila.');
    }
  } else {
    throw new Error('Ye column edit nahi ho sakta.');
  }
  if (value && allowed.indexOf(value) === -1) throw new Error('Galat value: ' + value);

  if (remarkCol && remark) mk.sh.getRange(row, remarkCol).setValue(remark);
  var rng = mk.sh.getRange(row, col);
  rng.setValue(value);
  pf_onEditMkt_(rng);          // backend ka onEdit wala hi kaam (wo khud flush karta hai)
  return { ok: true, row: row };
}

function ncClean_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function ncNorm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function ncJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
