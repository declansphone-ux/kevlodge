// ============================================================
// Masonic Lodge - Attendance Web App
// Google Apps Script Backend - v24 (multi-lodge, Set Up driven)
// Column map (Meeting Dates sheet) - emails now live in Set Up:
//  A=0 Date, B=1 DateName, C=2 Month, D=3 MaxCapacity,
//  E=4 MeetingType, F=5 WhoFor, G=6 SummonsLink,
//  H=7 Starter, I=8 Main, J=9 Dessert, K=10 Price, L=11 TyleTime
// Set Up sheet: column A = label, column B = value (read by keyword)
// ============================================================

const MEETINGS_SHEET  = 'Meeting Dates';
const RESPONSES_SHEET = 'Responses';
const CONTACT_SHEET   = 'Contact';
const SETUP_SHEET     = 'Set Up';

// The spreadsheet this script is bound to (each copy reads its own data)
var _ssCache = null;
function getSS() { if (!_ssCache) _ssCache = SpreadsheetApp.getActiveSpreadsheet(); return _ssCache; }

// ------------------------------------------------------------
// LODGE CONFIG - reads the "Set Up" tab (label in col A, value in col B)
// Matched by keyword so small label-wording differences still work.
// Cached per execution.
// ------------------------------------------------------------
var _lodgeCache = null;
// Convert a Google Drive "share" link into a direct-image URL.
// Handles /file/d/ID/view, ?id=ID, open?id=ID. Leaves other URLs untouched.
function toDirectImage(u) {
  u = String(u || '').trim();
  if (!u) return u;
  const m = u.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || u.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return 'https://lh3.googleusercontent.com/d/' + m[1];
  return u;
}
function getLodge() {
  if (_lodgeCache) return _lodgeCache;
  const out = {
    name: 'The Lodge', number: '', address: '', crest: '',
    craftChapter: 'Craft', secretaryEmail: '', stewardEmail: '', dcEmail: '',
    bankName: '', sortCode: '', accountNumber: '', almsLink: '',
    website: '', province: '', provinceUrl: '', adminPassword: '',
    worshipfulMaster: '', seniorWarden: '', juniorWarden: '', chaplain: '', directorOfCeremonies: '',
    tableAdminPassword: '', almsQrLink: '', charityLogo: ''
  };
  try {
    const sheet = getSS().getSheetByName(SETUP_SHEET);
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      for (let i = 0; i < rows.length; i++) {
        const label = String(rows[i][0] || '').toLowerCase().trim();
        const value = String(rows[i][1] == null ? '' : rows[i][1]).trim();
        if (!label) continue;
        if (label.indexOf('provincial') > -1 || label.indexOf('province') > -1) out.provinceUrl = value;  // before generic 'website'
        else if (label.indexOf('worshipful master') > -1) out.worshipfulMaster = value;
        else if (label.indexOf('senior warden') > -1) out.seniorWarden = value;
        else if (label.indexOf('junior warden') > -1) out.juniorWarden = value;
        else if (label.indexOf('chaplain') > -1) out.chaplain = value;
        else if (label.indexOf('director of ceremonies') > -1 && !/email|mail|@/.test(label)) out.directorOfCeremonies = value;
        else if (label.indexOf('table admin') > -1) out.tableAdminPassword = value;
        else if (label.indexOf('alms') > -1 && (label.indexOf('qr') > -1 || label.indexOf('code') > -1)) out.almsQrLink = value;
        else if (label.indexOf('charity') > -1 && label.indexOf('logo') > -1) out.charityLogo = value;
        else if (label.indexOf('host') > -1) out.website = value;
        else if (label.indexOf('name') > -1 && label.indexOf('account') < 0 && label.indexOf('bank') < 0) out.name = value || out.name;
        else if (label.indexOf('number') > -1 && label.indexOf('account') < 0) out.number = value;
        else if (label.indexOf('address') > -1) out.address = value;
        else if (label.indexOf('chapter') > -1 || label.indexOf('craft') > -1) out.craftChapter = /chapter/i.test(value) ? 'Chapter' : 'Craft';
        else if (label.indexOf('crest') > -1 || label.indexOf('logo') > -1) out.crest = value;
        else if (label.indexOf('secretary') > -1) out.secretaryEmail = value;
        else if (label.indexOf('steward') > -1) out.stewardEmail = value;
        else if (label.indexOf('ceremonies') > -1 || label.indexOf('dc') > -1) out.dcEmail = value;
        else if ((label.indexOf('account') > -1 && label.indexOf('name') > -1) || label.indexOf('bank') > -1) out.bankName = value;
        else if (label.indexOf('password') > -1 || label.indexOf('passcode') > -1 || label.indexOf('admin') > -1) out.adminPassword = value;
        else if (label.indexOf('sort') > -1) out.sortCode = value;
        else if (label.indexOf('account') > -1 && (label.indexOf('number') > -1 || label.indexOf('no') > -1)) out.accountNumber = value;
      }
    }
  } catch(e) {}

  // Number: strip everything except digits (removes a stray C/L or "No.")
  out.number = String(out.number).replace(/[^0-9]/g, '');

  // Name: ensure it ends with "Lodge" or "Chapter" (per Craft/Chapter)
  const suffix = (out.craftChapter === 'Chapter') ? 'Chapter' : 'Lodge';
  let nm = String(out.name || '').trim();
  if (nm && !new RegExp(suffix + '$', 'i').test(nm) && !/lodge$|chapter$/i.test(nm)) nm = nm + ' ' + suffix;
  out.name = nm;

  // Crest: convert Drive share links to direct-image URLs
  out.crest = toDirectImage(out.crest);

  // Alms QR destination: parse any Drive link to a direct URL
  out.almsQrLink = toDirectImage(out.almsQrLink);
  out.charityLogo = toDirectImage(out.charityLogo);

  // Alms link: Craft = L, Chapter = C, + number
  const prefix = (out.craftChapter === 'Chapter') ? 'C' : 'L';
  out.almsLink = 'https://gtap.uk/' + prefix + out.number;

  // The Freemasons' Charity / alms collection image used on emails and table cards.
  out.almsImage = 'https://lh3.googleusercontent.com/d/199nzhpdRIq7rXvSUTDeS-eWKjPIG1XSt';

  // Full name with "No." prefix on the number if not already present
  const numLabel = out.number ? (/^no\.?\s/i.test(out.number) ? out.number : ('No. ' + out.number)) : '';
  out.numberLabel = numLabel;
  out.fullName = out.name + (numLabel ? ' ' + numLabel : '');

  // Website host (pretty cancel URL); normalise protocol + trailing slash
  let site = String(out.website || '').trim();
  if (site) {
    if (!/^https?:\/\//i.test(site)) site = 'https://' + site;
    if (!/\/$/.test(site)) site += '/';
    out.website = site;
  }
  out.cancelBase = out.website || (function(){ try { return ScriptApp.getService().getUrl(); } catch(e){ return ''; } })();

  // Province URL normalise (optional)
  let prov = String(out.provinceUrl || '').trim();
  if (prov && !/^https?:\/\//i.test(prov)) prov = 'https://' + prov;
  out.provinceUrl = prov;

  _lodgeCache = out;
  return out;
}


function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'book';
  const code = (e && e.parameter && e.parameter.code) || '';
  const template = HtmlService.createTemplateFromFile('Index');
  template.initMode = mode;
  template.initCode = code;
  const lodge = getLodge();
  return template.evaluate()
    .setTitle(lodge.fullName || 'Lodge Booking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------
// RANK & SALUTATION DATA - returns the correct set for Craft or Chapter.
// Chapter lists use standard UGLE Royal Arch (Supreme Grand Chapter) terms.
// Ranks never appear in both - the Set Up "Chapter or Craft" cell decides.
// ------------------------------------------------------------
function getRankData(craftChapter) {
  const isChapter = (craftChapter === 'Chapter');

  if (!isChapter) {
    return {
      mode: 'Craft',
      salutations: ['Bro','W Bro','VW Bro','RW Bro','MW Bro'],
      provLabel: 'Provincial Rank',
      grandLabel: 'Grand Rank',
      provRanks: [
        { label: 'Active Provincial', opts: ['Provincial Grand Master','Deputy Provincial Grand Master','Assistant Provincial Grand Master','Provincial Senior Grand Warden','Provincial Junior Grand Warden','Provincial Grand Chaplain','Provincial Grand Treasurer','Provincial Grand Registrar','Provincial Grand Secretary','Provincial Grand Director of Ceremonies','Provincial Grand Sword Bearer','Provincial Grand Superintendent of Works','Provincial Deputy Grand Chaplain','Provincial Deputy Grand Registrar','Provincial Deputy Grand Secretary','Provincial Deputy Grand Director of Ceremonies','Provincial Deputy Grand Sword Bearer','Provincial Deputy Grand Superintendent of Works','Provincial Grand Almoner','Provincial Grand Charity Steward','Provincial Grand Mentor','Provincial Grand Learning and Development Officer','Provincial Grand Membership Officer','Provincial Grand Communications Officer','Provincial Senior Grand Deacon','Provincial Junior Grand Deacon','Provincial Assistant Grand Chaplain','Provincial Assistant Grand Registrar','Provincial Assistant Grand Secretary','Provincial Assistant Grand Director of Ceremonies','Provincial Assistant Grand Sword Bearer','Provincial Assistant Grand Superintendent of Works','Provincial Grand Organist','Provincial Grand Standard Bearer','Provincial Assistant Grand Standard Bearer','Provincial Deputy Grand Organist','Provincial Grand Pursuivant','Provincial Assistant Grand Pursuivant','Provincial Grand Steward','Provincial Grand Tyler','Other Active Provincial Rank'] },
        { label: 'Past Provincial', opts: ['Past Provincial Grand Master','Past Deputy Provincial Grand Master','Past Assistant Provincial Grand Master','Past Provincial Senior Grand Warden','Past Provincial Junior Grand Warden','Past Provincial Grand Chaplain','Past Provincial Grand Treasurer','Past Provincial Grand Registrar','Past Provincial Grand Secretary','Past Provincial Grand Director of Ceremonies','Past Provincial Grand Sword Bearer','Past Provincial Grand Superintendent of Works','Past Provincial Deputy Grand Chaplain','Past Provincial Deputy Grand Registrar','Past Provincial Deputy Grand Secretary','Past Provincial Deputy Grand Director of Ceremonies','Past Provincial Deputy Grand Sword Bearer','Past Provincial Deputy Grand Superintendent of Works','Past Provincial Grand Almoner','Past Provincial Grand Charity Steward','Past Provincial Grand Mentor','Past Provincial Grand Learning and Development Officer','Past Provincial Grand Membership Officer','Past Provincial Grand Communications Officer','Past Provincial Senior Grand Deacon','Past Provincial Junior Grand Deacon','Past Provincial Assistant Grand Chaplain','Past Provincial Assistant Grand Registrar','Past Provincial Assistant Grand Secretary','Past Provincial Assistant Grand Director of Ceremonies','Past Provincial Assistant Grand Sword Bearer','Past Provincial Assistant Grand Superintendent of Works','Past Provincial Grand Organist','Past Provincial Grand Standard Bearer','Past Provincial Assistant Grand Standard Bearer','Past Provincial Deputy Grand Organist','Past Provincial Grand Pursuivant','Past Provincial Assistant Grand Pursuivant','Past Provincial Grand Steward','Past Provincial Grand Tyler','Other Past Provincial Rank'] },
        { label: 'Uncollared Ranks', opts: ['Assistant Provincial Grand Almoner','Assistant Provincial Grand Charity Steward'] }
      ],
      grandRanks: ['Grand Master','Pro Grand Master','Deputy Grand Master','Assistant Grand Master','Metropolitan, Provincial, and District Grand Masters','Senior Grand Warden','Junior Grand Warden','Grand Chaplain','President of the Board of General Purposes','Grand Chancellor','Grand Registrar','Grand Secretary','President of the Masonic Charitable Foundation','Grand Director of Ceremonies','Grand Sword Bearer','Grand Superintendent of Works','Grand Inspector','Deputy Grand Chaplain','Deputy President of the Board of General Purposes','Deputy Grand Chancellor','Deputy Grand Registrar','Deputy Grand Secretary','Deputy Grand Director of Ceremonies','Deputy Grand Sword Bearer','Deputy Grand Superintendent of Works','Senior Grand Deacon','Junior Grand Deacon','Assistant Grand Chaplain','Assistant Grand Chancellor','Assistant Grand Registrar','Assistant Grand Secretary','Assistant Grand Director of Ceremonies','Assistant Grand Sword Bearer','Assistant Grand Superintendent of Works','Grand Organist','Grand Standard Bearer','Assistant Grand Standard Bearer','Deputy Grand Organist','Grand Pursuivant','Assistant Grand Pursuivant','Past Grand Officer']
    };
  }

  // ---- Royal Arch (Chapter) - standard UGLE Supreme Grand Chapter terms ----
  return {
    mode: 'Chapter',
    salutations: ['Comp','E Comp','VE Comp','RE Comp','ME Comp'],
    provLabel: 'Provincial Grand Chapter Rank',
    grandLabel: 'Supreme Grand Chapter Rank',
    provRanks: [
      { label: 'Active Provincial Grand Chapter', opts: ['Grand Superintendent','Second Provincial Grand Principal','Third Provincial Grand Principal','Provincial Grand Scribe Ezra','Provincial Grand Scribe Nehemiah','Provincial Grand Treasurer','Provincial Grand Registrar','Provincial Grand Principal Sojourner','Provincial Grand Director of Ceremonies','Provincial Grand Sword Bearer','Provincial Grand Standard Bearer','Provincial Grand Almoner','Provincial Grand Charity Steward','Provincial Deputy Grand Director of Ceremonies','Provincial Assistant Grand Director of Ceremonies','Provincial Grand Organist','Provincial Grand Assistant Sojourner','Provincial Grand Steward','Provincial Grand Janitor','Other Active Provincial Grand Chapter Rank'] },
      { label: 'Past Provincial Grand Chapter', opts: ['Past Provincial Second Grand Principal','Past Provincial Third Grand Principal','Past Provincial Grand Scribe Ezra','Past Provincial Grand Scribe Nehemiah','Past Provincial Grand Treasurer','Past Provincial Grand Registrar','Past Provincial Grand Principal Sojourner','Past Provincial Grand Director of Ceremonies','Past Provincial Grand Sword Bearer','Past Provincial Grand Standard Bearer','Past Provincial Grand Almoner','Past Provincial Grand Charity Steward','Past Provincial Deputy Grand Director of Ceremonies','Past Provincial Assistant Grand Director of Ceremonies','Past Provincial Grand Organist','Past Provincial Grand Assistant Sojourner','Past Provincial Grand Steward','Other Past Provincial Grand Chapter Rank'] }
    ],
    grandRanks: ['First Grand Principal','Second Grand Principal','Third Grand Principal','Grand Scribe Ezra','Grand Scribe Nehemiah','Grand Treasurer','Grand Registrar','Grand Principal Sojourner','Grand Director of Ceremonies','Grand Sword Bearer','Grand Standard Bearer','Deputy Grand Director of Ceremonies','Assistant Grand Director of Ceremonies','Grand Almoner','Grand Charity Steward','Grand Organist','Grand Assistant Sojourner','Grand Steward','Past Grand Officer (Royal Arch)']
  };
}

// ============================================================
// CONFIG - reads upcoming meeting, falls back to last row
// ============================================================
// ------------------------------------------------------------
// ADMIN VIEW - returns current meeting bookings ONLY if the password matches.
// Data never leaves the server unless the password is correct.
// ------------------------------------------------------------
function getAdminBookings(password) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real) return { error: 'no_password', message: 'No admin password is set in the Set Up tab.' };
  if (String(password || '').trim() !== real) return { error: 'bad_password', message: 'Incorrect password.' };

  const config = getConfig();
  if (!config || !config.success) return { error: 'no_meeting', message: 'No current meeting found.' };
  const meetingName = config.meeting.dateName;
  const price = (function(){ const m = String(config.meeting.price||'').replace(/,/g,'').match(/\d+(\.\d+)?/); return m ? parseFloat(m[0]) : 0; })();

  const sheet = getSS().getSheetByName(RESPONSES_SHEET);
  if (sheet.getFilter()) sheet.getFilter().remove();
  const data = sheet.getDataRange().getValues();

  // Classify a person into a rank group from their grand/prov rank text.
  function rankGroup(grandRank, provRank) {
    const g = String(grandRank || '').trim();
    const p = String(provRank || '').trim();
    if (g && g.toLowerCase() !== 'none') return 'grand';
    if (p && p.toLowerCase() !== 'none') return (/past/i.test(p) ? 'pastProv' : 'activeProv');
    return 'brethren';
  }

  const rows = [];
  let diningCount = 0, meetingOnlyCount = 0, apologyCount = 0, total = 0;
  const groupCounts = { grand: 0, activeProv: 0, pastProv: 0, brethren: 0, guest: 0 };

  // Reduce any value (a real Date, a date string, or the formatted name like
  // "Thursday, 8 October 2026") to a canonical YYYY-MM-DD so they compare
  // reliably. The sheet often stores column D as an actual date, while the
  // config name is formatted text - this bridges both.
  const MONTHS = {january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
  function toYMD(val) {
    if (val instanceof Date && !isNaN(val.getTime())) {
      return val.getFullYear() + '-' + (val.getMonth()+1) + '-' + val.getDate();
    }
    const s = String(val || '').trim();
    if (!s) return '';
    // Try "8 October 2026" / "Thursday, 8 October 2026"
    const m = s.toLowerCase().match(/(\d{1,2})\s+([a-z]+)\s+(\d{4})/);
    if (m && MONTHS[m[2]] != null) return parseInt(m[3],10) + '-' + (MONTHS[m[2]]+1) + '-' + parseInt(m[1],10);
    // Fall back to Date parsing
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.getFullYear() + '-' + (d.getMonth()+1) + '-' + d.getDate();
    return '';
  }
  // Text-key fallback (letters+digits only) for odd formats
  function dateKey(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
  const wantYMD = toYMD(config.meeting.date) || toYMD(meetingName);
  const wantKey = dateKey(meetingName);

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const haveYMD = toYMD(r[3]);
    const haveKey = dateKey(r[3]);
    const dateMatch = (wantYMD && haveYMD === wantYMD) ||
                      (haveKey && (haveKey === wantKey || haveKey.indexOf(wantKey) >= 0 || wantKey.indexOf(haveKey) >= 0));
    if (!dateMatch) continue;
    // Exclude only rows explicitly cancelled (blank status counts as active).
    if (/cancel/i.test(String(r[19]))) continue;
    let role = String(r[5]).trim();
    const type = String(r[4]).trim();
    const isGuestRow = /guest/i.test(role);
    const sal  = String(r[6] || '').trim();
    const name = [r[7], r[8]].filter(Boolean).join(' ').trim();
    const diet = String(r[16] || '').trim();
    const provRank  = String(r[12] || '').trim();
    const grandRank = String(r[14] || '').trim();
    const cost = parseFloat(String(r[18]).replace(/[^0-9.]/g,'')) || 0;
    // Skip totally empty rows (no name and no type)
    if (!name && !type) continue;
    total += cost;
    if (type === 'dining') diningCount++;

    if (!isGuestRow) {
      if (type === 'apology') apologyCount++;
      else if (type === 'meeting') meetingOnlyCount++;
      const grp = rankGroup(grandRank, provRank);
      const rankDisplay = grandRank || provRank || '';
      if (type !== 'apology') groupCounts[grp]++;
      rows.push({
        sal: sal, name: name, surname: String(r[8] || '').trim(),
        fullName: (sal ? sal + ' ' : '') + name,
        type: type, group: grp, rank: rankDisplay,
        diet: diet, isGuest: false, cost: cost,
        email: String(r[9] || '').trim(), payMethod: String(r[22] || '').trim(),
        rowIndex: i + 1, paid: String(r[26] || '').trim(), arrived: String(r[27] || '').trim(), notes: String(r[28] || '').trim()
      });
    } else {
      const ggrp = rankGroup(grandRank, provRank);
      if (type !== 'apology') groupCounts.guest++;
      rows.push({
        sal: sal, name: name, surname: String(r[8] || '').trim(),
        fullName: (sal ? sal + ' ' : '') + name,
        type: 'guest', group: ggrp, rank: (grandRank || provRank || ''),
        host: String(r[20] || '').trim(), diet: diet, isGuest: true, cost: cost,
        email: String(r[9] || '').trim(), payMethod: String(r[22] || '').trim(),
        rowIndex: i + 1, paid: String(r[26] || '').trim(), arrived: String(r[27] || '').trim(), notes: String(r[28] || '').trim()
      });
    }
  }

  return {
    success: true,
    lodge: { name: lodge.fullName, craftChapter: lodge.craftChapter, crest: lodge.crest, almsQrLink: lodge.almsQrLink, almsLink: lodge.almsLink, charityLogo: lodge.charityLogo, almsImage: lodge.almsImage },
    officers: { wm: lodge.worshipfulMaster, sw: lodge.seniorWarden, jw: lodge.juniorWarden, dc: lodge.directorOfCeremonies, chaplain: lodge.chaplain },
    contacts: { secretaryEmail: lodge.secretaryEmail, stewardEmail: lodge.stewardEmail, dcEmail: lodge.dcEmail },
    meeting: { dateName: meetingName, date: config.meeting.date, tyleTime: config.meeting.tyleTime, price: price },
    rows: rows,
    summary: {
      diningCount: diningCount,
      meetingOnlyCount: meetingOnlyCount,
      apologyCount: apologyCount,
      festiveBoardTotal: '£' + total.toFixed(2),
      headcount: rows.filter(function(x){ return x.type !== 'apology'; }).length,
      groups: groupCounts
    }
  };
}

// Email the attendance list to a typed address (officers only - password re-checked).
function emailAdminList(password, toEmail) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real || String(password||'').trim() !== real) return { error: 'bad_password', message: 'Incorrect password.' };
  if (!toEmail || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(toEmail).trim())) return { error: 'bad_email', message: 'Please enter a valid email address.' };
  const d = getAdminBookings(password);
  if (!d || !d.success) return d || { error: 'fail' };

  const groupLabels = { grand: 'Grand Officers', activeProv: 'Active Provincial Officers', pastProv: 'Past Provincial Officers', brethren: (lodge.craftChapter==='Chapter'?'Companions':'Brethren'), guest: 'Guests' };
  const order = ['grand','activeProv','pastProv','brethren','guest'];
  let sections = '';
  order.forEach(function(g){
    const people = d.rows.filter(function(x){ return x.group===g && x.type!=='apology'; });
    if (!people.length) return;
    let lis = people.map(function(p){
      const tag = p.isGuest ? (' <span style="color:#7A5E12">(guest'+(p.host?' of '+p.host:'')+')</span>') : (p.type==='meeting'?' <span style="color:#7A6A5A">(meeting only)</span>':'');
      const diet = p.diet ? ' <span style="color:#8B1A1A;font-size:12px">[' + p.diet + ']</span>' : '';
      return '<tr><td style="padding:5px 0;border-bottom:1px solid #EEE6D2;">'+p.fullName+(p.rank?' <span style="color:#7A5E12;font-size:12px">'+p.rank+'</span>':'')+tag+diet+'</td></tr>';
    }).join('');
    sections += '<h3 style="font-family:Georgia,serif;color:#8B1A1A;font-size:14px;letter-spacing:1px;text-transform:uppercase;margin:18px 0 6px;border-bottom:2px solid rgba(212,175,55,.4);padding-bottom:4px;">'+groupLabels[g]+' ('+people.length+')</h3><table style="width:100%;border-collapse:collapse;font-family:Georgia,serif;font-size:14px;color:#1A0F0A;">'+lis+'</table>';
  });
  const apologies = d.rows.filter(function(x){ return x.type==='apology'; });
  if (apologies.length) {
    sections += '<h3 style="font-family:Georgia,serif;color:#7A6A5A;font-size:14px;letter-spacing:1px;text-transform:uppercase;margin:18px 0 6px;border-bottom:2px solid #DDD;padding-bottom:4px;">Apologies ('+apologies.length+')</h3><table style="width:100%;border-collapse:collapse;font-family:Georgia,serif;font-size:14px;color:#5A4736;">'+apologies.map(function(p){return '<tr><td style="padding:5px 0;border-bottom:1px solid #EEE;">'+p.fullName+'</td></tr>';}).join('')+'</table>';
  }
  const s = d.summary;
  const html = `<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:#FBF4E2;border:1px solid rgba(212,175,55,.4);">
    ${emailHeader('ATTENDANCE LIST')}
    <div style="padding:24px;">
      <p style="font-size:16px;color:#1A0F0A;margin:0 0 4px;font-weight:bold;">${d.meeting.dateName}</p>
      <p style="font-size:13px;color:#7A6A5A;margin:0 0 16px;">${d.meeting.tyleTime?('Tyles '+d.meeting.tyleTime):''}</p>
      <table style="width:100%;border-collapse:collapse;margin-bottom:8px;">
        <tr>
          <td style="text-align:center;padding:10px;background:#fff;border:1px solid rgba(212,175,55,.3);"><div style="font-size:22px;color:#8B1A1A;font-weight:bold;">${s.headcount}</div><div style="font-size:11px;color:#7A6A5A;text-transform:uppercase;letter-spacing:1px;">Attending</div></td>
          <td style="text-align:center;padding:10px;background:#fff;border:1px solid rgba(212,175,55,.3);"><div style="font-size:22px;color:#8B1A1A;font-weight:bold;">${s.diningCount}</div><div style="font-size:11px;color:#7A6A5A;text-transform:uppercase;letter-spacing:1px;">Dining</div></td>
          <td style="text-align:center;padding:10px;background:#fff;border:1px solid rgba(212,175,55,.3);"><div style="font-size:22px;color:#8B1A1A;font-weight:bold;">${s.apologyCount}</div><div style="font-size:11px;color:#7A6A5A;text-transform:uppercase;letter-spacing:1px;">Apologies</div></td>
          <td style="text-align:center;padding:10px;background:#fff;border:1px solid rgba(212,175,55,.3);"><div style="font-size:18px;color:#8B1A1A;font-weight:bold;">${s.festiveBoardTotal}</div><div style="font-size:11px;color:#7A6A5A;text-transform:uppercase;letter-spacing:1px;">Festive Board</div></td>
        </tr>
      </table>
      ${sections}
      ${emailFooter()}
    </div>
  </div>`;
  try {
    GmailApp.sendEmail(String(toEmail).trim(), `${lodge.name} - Attendance List - ${d.meeting.dateName}`, '', { htmlBody: html, name: lodge.name + ' - Attendance' });
    return { success: true };
  } catch(e) { return { error: 'send_fail', message: e.toString() }; }
}

function getConfig() {
  try {
    const ss            = getSS();
    const lodge         = getLodge();
    const sheet         = ss.getSheetByName(MEETINGS_SHEET);
    const values        = sheet.getDataRange().getValues();
    const displayValues = sheet.getDataRange().getDisplayValues();
    const tz            = Session.getScriptTimeZone();
    const todayStr      = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const tp            = todayStr.split('-');
    const todayLocal    = new Date(parseInt(tp[0]), parseInt(tp[1])-1, parseInt(tp[2]), 0, 0, 0, 0);
    const todayMs       = todayLocal.getTime();

    let lastValidRow = null;

    for (let i = 1; i < values.length; i++) {
      const row     = values[i];
      const dispRow = displayValues[i];
      const raw     = row[0];
      if (!raw) continue;

      let meetingDateObj = (raw instanceof Date) ? raw : parseUKDate(raw);
      if (!meetingDateObj || isNaN(meetingDateObj.getTime())) continue;

      const meetingStr   = Utilities.formatDate(meetingDateObj, tz, 'yyyy-MM-dd');
      const mp           = meetingStr.split('-');
      const meetingLocal = new Date(parseInt(mp[0]), parseInt(mp[1])-1, parseInt(mp[2]), 0, 0, 0, 0);
      const meetingMs    = meetingLocal.getTime();
      const cleanName    = dispRow[1].trim();
      const daysUntil    = Math.round((meetingMs - todayMs) / 864e5);
      const booked       = countBookingsForMeeting(ss, cleanName);
      const capCell      = row[3];

      let capacity, spotsLeft;
      if (capCell === '' || capCell === null) { capacity = -1; spotsLeft = 999; }
      else { capacity = parseInt(capCell); spotsLeft = capacity <= 0 ? 0 : Math.max(0, capacity - booked); }

      const meetingObj = {
        date:           meetingStr,
        dateName:       cleanName,
        month:          String(row[2]).trim(),
        capacity:       capacity,
        spotsLeft:      spotsLeft,
        meetingType:    String(row[4]  || '').trim(),
        whoFor:         String(row[5]  || '').trim(),
        summonsUrl:     String(row[6]  || '').trim(),   // G
        starter:        String(row[7]  || '').trim(),   // H
        main:           String(row[8]  || '').trim(),   // I
        dessert:        String(row[9]  || '').trim(),   // J
        price:          formatCurrency(row[10]),         // K
        tyleTime:       formatTime(row[11]),             // L
        secretaryEmail: lodge.secretaryEmail,
        stewardEmail:   lodge.stewardEmail,
        dcEmail:        lodge.dcEmail,
        daysUntil:      daysUntil,
        bookingOpen:    daysUntil >= 7
      };

      lastValidRow = meetingObj;
      if (meetingMs >= todayMs) return { success: true, meeting: meetingObj, lodge: lodge, ranks: getRankData(lodge.craftChapter) };
    }

    if (lastValidRow) {
      lastValidRow.bookingOpen = false;
      lastValidRow.spotsLeft   = 0;
      return { success: true, meeting: lastValidRow, noUpcoming: true, lodge: lodge, ranks: getRankData(lodge.craftChapter) };
    }

    return { error: 'no_meeting', message: 'No meetings found.', lodge: lodge };
  } catch(e) {
    return { error: 'sheet_error', message: e.toString() };
  }
}

function countBookingsForMeeting(ss, dateName) {
  try {
    const sheet = ss.getSheetByName(RESPONSES_SHEET);
    const data  = sheet.getDataRange().getDisplayValues();
    let count   = 0;
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (r[3].trim() === dateName.trim() && r[4].trim() === 'dining' && r[19].trim() === 'Confirmed') count++;
    }
    return count;
  } catch(e) { return 0; }
}

// ============================================================
// LOOKUP PREVIOUS BOOKER - for email pre-fill
// ============================================================
function lookupPreviousBooker(email) {
  try {
    const target = String(email || '').toLowerCase().trim();
    if (!target) return { success: false, reason: 'no-email' };

    const ss    = getSS();
    const sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (!sheet) return { success: false, reason: 'no-sheet' };
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, reason: 'empty-sheet' };

    // Normalise an email cell for comparison (handle stray spaces / case)
    const norm = v => String(v == null ? '' : v).toLowerCase().trim();

    let leadMatch = null;   // preferred: a row flagged 'Lead Booker'
    let anyMatch  = null;   // fallback: most recent row with this email anywhere

    for (let i = data.length - 1; i >= 1; i--) {
      const r = data[i];
      // The email normally lives in column 9, but scan the row so a shifted
      // column or a guest-only row still resolves the person.
      const emailInRow = norm(r[9]) === target ||
        r.some(cell => norm(cell) === target);
      if (!emailInRow) continue;

      if (!anyMatch) anyMatch = r;
      if (String(r[5]).trim().toLowerCase() === 'lead booker') { leadMatch = r; break; }
    }

    const r = leadMatch || anyMatch;
    if (!r) return { success: false, reason: 'not-found' };

    return {
      success:    true,
      salutation: String(r[6]  || '').trim(),
      firstName:  String(r[7]  || '').trim(),
      surname:    String(r[8]  || '').trim(),
      provRank:   String(r[12] || '').trim(),
      grandRank:  String(r[14] || '').trim(),
      dietary:    String(r[16] || '').trim(),
      fallback:   !leadMatch    // true if matched via a non-Lead row
    };
  } catch(e) {
    return { success: false, reason: 'error', message: String(e) };
  }
}

// ============================================================
// SUBMIT BOOKING
// ============================================================
// Check whether an active (non-cancelled) booking already exists for this
// full name at the SAME meeting. Returns the most recent match so the front
// end can show it and ask the member whether to supersede it.
function checkDuplicateBooking(meetingDate, firstName, surname) {
  try {
    const sheet = getSS().getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    const data = sheet.getDataRange().getValues();
    const wantName = (String(firstName||'').trim() + ' ' + String(surname||'').trim()).toLowerCase().replace(/\s+/g,' ').trim();
    const wantMeet = String(meetingDate||'').toLowerCase().replace(/[^a-z0-9]/g,'');
    let found = null;
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (/cancel/i.test(String(r[19]))) continue;                 // skip cancelled
      const haveMeet = String(r[3]||'').toLowerCase().replace(/[^a-z0-9]/g,'');
      if (haveMeet !== wantMeet) continue;                          // same meeting only
      const haveName = (String(r[7]||'').trim() + ' ' + String(r[8]||'').trim()).toLowerCase().replace(/\s+/g,' ').trim();
      if (haveName !== wantName) continue;
      // record this match (keep the latest by timestamp / row order)
      found = {
        rowIndex: i + 1,
        type: String(r[4]||''),
        salutation: String(r[6]||''),
        name: haveName,
        when: (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'd MMM yyyy, HH:mm') : String(r[0]||''),
        ref: String(r[2]||''),
        attendanceLabel: /apolog/i.test(String(r[4])) ? 'an apology' : (/meeting/i.test(String(r[4])) ? 'a meeting-only booking' : 'a dining booking')
      };
    }
    return { success: true, duplicate: !!found, existing: found };
  } catch (e) {
    return { success: false, duplicate: false, message: e.toString() };
  }
}

// Cancel any active rows for the same name + meeting (lead booker and their guests
// share the cancel via ref; we cancel rows matching the lead name for that meeting).
function supersedeExisting(sheet, meetingDate, firstName, surname) {
  const data = sheet.getDataRange().getValues();
  const wantName = (String(firstName||'').trim() + ' ' + String(surname||'').trim()).toLowerCase().replace(/\s+/g,' ').trim();
  const wantMeet = String(meetingDate||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const refsToCancel = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (/cancel/i.test(String(r[19]))) continue;
    const haveMeet = String(r[3]||'').toLowerCase().replace(/[^a-z0-9]/g,'');
    if (haveMeet !== wantMeet) continue;
    const haveName = (String(r[7]||'').trim() + ' ' + String(r[8]||'').trim()).toLowerCase().replace(/\s+/g,' ').trim();
    if (haveName === wantName) refsToCancel[String(r[2]||'')] = true;
  }
  // Cancel every row sharing those refs (so guests of a superseded booking go too)
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (refsToCancel[String(r[2]||'')] && !/cancel/i.test(String(r[19]))) {
      sheet.getRange(i + 1, 20).setValue('Cancelled (superseded)');   // col T
    }
  }
}

function submitBooking(formData) {
  try {
    const ss    = getSS();
    const sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();

    // If the user confirmed they want to supersede an existing booking for the
    // same name + meeting, cancel the previous active row(s) first.
    if (formData.supersede) {
      supersedeExisting(sheet, formData.meetingDate, formData.memberFirstName, formData.memberSurname);
    }

    const ts       = new Date();
    const ref        = generateRef(formData.memberSurname, formData.meetingDate);
    const cancelCode = generateCancelCode(sheet);
    // Parse the per-head price defensively - take the FIRST number only, so a
    // value like "£25.00 / £50.00" or any stray second number can never double it.
    function parseFirstPrice(v) {
      const m = String(v == null ? '' : v).replace(/,/g,'').match(/\d+(\.\d+)?/);
      const n = m ? parseFloat(m[0]) : NaN;
      return (isNaN(n) || n <= 0) ? 24.50 : n;
    }
    const priceNum = parseFirstPrice(formData.price);
    const isDining = formData.attendanceType === 'dining';
    const hostName = (formData.memberFirstName + ' ' + formData.memberSurname).trim();
    const mCat     = formData.memberDiningCategory || 'Standard Dining';
    const mCost    = (isDining && mCat === 'Standard Dining') ? priceNum : 0;
    const pMethod  = isDining ? (formData.paymentMethod || 'BACS') : '';

    // Use config passed from front end - avoids duplicate getConfig() call
    const meetingM  = formData.meetingConfig || null;
    const tyleTime  = formData.tyleTime || (meetingM ? meetingM.tyleTime : '18:15');
    const secEmail  = meetingM ? meetingM.secretaryEmail : '';
    const stewEmail = meetingM ? meetingM.stewardEmail   : '';
    const dcEmail   = meetingM ? meetingM.dcEmail        : '';

    sheet.appendRow([
      ts, formData.device || 'Unknown', ref, formData.meetingDate || '', formData.attendanceType,
      'Lead Booker', formData.memberSalutation || '', formData.memberFirstName || '', formData.memberSurname || '',
      formData.memberEmail || '', '', '',
      formData.memberProvRank || '', getProvAbbrev(formData.memberProvRank),
      formData.memberGrandRank || '', getGrandAbbrev(formData.memberGrandRank),
      formData.memberDietary || '', formData.apologyReason || '',
      '£' + mCost.toFixed(2), 'Confirmed', '', mCat, pMethod,
      (formData.isOfficer === 'Yes') ? 'Yes' : '', formData.officerWork || '', cancelCode
    ]);

    let runningTotal = mCost;
    if (isDining && formData.guests && formData.guests.length > 0) {
      formData.guests.forEach(g => {
        const gCat  = g.diningCategory || 'Standard Dining';
        const gCost = (gCat === 'Standard Dining') ? priceNum : 0;
        runningTotal += gCost;
        sheet.appendRow([
          ts, formData.device || 'Unknown', ref, formData.meetingDate || '', 'dining', 'Guest',
          g.salutation || '', g.firstName || '', g.surname || '', formData.memberEmail || '',
          '', '', g.provRank || '', getProvAbbrev(g.provRank),
          g.grandRank || '', getGrandAbbrev(g.grandRank), g.dietary || '', '',
          '£' + gCost.toFixed(2), 'Confirmed', hostName, gCat, pMethod, '', '', cancelCode
        ]);
      });
    }

    sendBookingConfirmationEmail(formData, ref, runningTotal, pMethod, tyleTime, meetingM, stewEmail, cancelCode);
    if (formData.attendanceType === 'apology') {
      sendApologyNotifications(formData, ref, secEmail, stewEmail, dcEmail);
    }

    return { success: true, ref: ref, cancelCode: cancelCode, tyleTime: tyleTime, meetingDate: formData.meetingDate };
  } catch(e) {
    return { error: 'submit_error', message: e.toString() };
  }
}

// ============================================================
// LOOKUP BOOKING (for cancellation)
// ============================================================
function lookupBooking(code) {
  try {
    const ss     = getSS();
    const sheet  = ss.getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    const data   = sheet.getDataRange().getValues();

    const target = String(code||'').toUpperCase().trim();
    if (!target) return { success: false, message: 'Please enter your cancellation code.' };

    let allRefs        = [];
    let memberRowMatch = null;

    // Find the Lead Booker row whose cancellation code matches
    for (let i = data.length - 1; i >= 1; i--) {
      const r = data[i];
      if (
        String(r[25]).toUpperCase().trim() === target                        &&
        String(r[5]).trim() === 'Lead Booker'                                &&
        String(r[19]).trim() === 'Confirmed'
      ) {
        allRefs.push(String(r[2]));
        if (!memberRowMatch) memberRowMatch = r;
      }
    }

    if (!memberRowMatch) {
      let wasCancelled = false;
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        if (
          String(r[25]).toUpperCase().trim() === target &&
          String(r[5]).trim() === 'Lead Booker' &&
          String(r[19]).trim() === 'Cancelled'
        ) { wasCancelled = true; break; }
      }
      return {
        success: false, alreadyCancelled: wasCancelled,
        message: wasCancelled
          ? 'It looks like this booking has already been cancelled. No further action is needed.'
          : 'No booking found for that code. Please check it and try again, or use the option below if you do not have your code.'
      };
    }

    let totalCost  = 0;
    let guestNames = [];

    // Sum cost and gather guests by the UNIQUE cancellation code - not the
    // shared SURNAME-MONYY reference, which can collide between different people.
    for (let j = 1; j < data.length; j++) {
      const r = data[j];
      if (String(r[25]).toUpperCase().trim() === target && String(r[19]).trim() === 'Confirmed') {
        totalCost += parseFloat(String(r[18]).replace(/[^0-9.]/g,'')) || 0;
        if (String(r[5]).trim() === 'Guest') {
          const gn = [String(r[6]||''), String(r[7]||''), String(r[8]||'')].filter(Boolean).join(' ');
          if (gn.trim()) guestNames.push(gn.trim());
        }
      }
    }

    return {
      success:        true,
      ref:            allRefs[0],
      refs:           allRefs,
      bookingCount:   allRefs.length,
      attendanceType: String(memberRowMatch[4]),
      meetingDate:    cleanSheetDate(memberRowMatch[3]),
      bookerName:     [memberRowMatch[6], memberRowMatch[7], memberRowMatch[8]].filter(Boolean).join(' '),
      guests:         guestNames,
      totalCost:      '£' + totalCost.toFixed(2)
    };
  } catch(e) { return { error: 'lookup_error', message: e.toString() }; }
}

// ============================================================
// CANCEL BOOKING
// ============================================================
// ============================================================
// ADD GUEST(S) TO AN EXISTING BOOKING (by cancellation/booking code)
// No password: the member's unique code authorises it, exactly like cancel.
// Each added guest is charged the same per-head meeting price and appended
// under the same reference + code so they cancel together.
// ============================================================
function addGuestsToBooking(code, guests) {
  try {
    if (!guests || !guests.length) return { success: false, message: 'No guests to add.' };
    const ss = getSS();
    const sheet = ss.getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    const data = sheet.getDataRange().getValues();
    const target = String(code || '').toUpperCase().trim();
    if (!target) return { success: false, message: 'Please enter your booking code.' };

    // Find the Lead Booker row for this code.
    let lead = null;
    for (let i = data.length - 1; i >= 1; i--) {
      const r = data[i];
      if (String(r[25]).toUpperCase().trim() === target &&
          String(r[5]).trim() === 'Lead Booker' &&
          String(r[19]).trim() === 'Confirmed') { lead = r; break; }
    }
    if (!lead) return { success: false, message: 'No confirmed booking found for that code. Please check the code from your confirmation email.' };
    if (/apolog/i.test(String(lead[4]))) return { success: false, message: 'This is an apology, not a dining booking, so guests cannot be added.' };

    const ref        = String(lead[2]);
    const meetingDate= String(lead[3]);
    const hostName   = [lead[6], lead[7], lead[8]].filter(Boolean).join(' ');
    const memberEmail= String(lead[9] || '');
    const pMethod    = String(lead[22] || '');

    // Per-head price from the meeting config (same price as the booking).
    const config = getConfig();
    const priceNum = config && config.success ? (parseFloat(String(config.meeting.price).replace(/[^0-9.]/g,'')) || 0) : 0;

    const ts = new Date();
    let added = 0, addedCost = 0;
    guests.forEach(function (g) {
      if (!g || (!g.firstName && !g.surname)) return;
      const gCat = g.diningCategory || 'Standard Dining';
      const gCost = priceNum;                       // same per-head price
      addedCost += gCost; added++;
      sheet.appendRow([
        ts, 'Amend', ref, meetingDate, 'dining', 'Guest',
        g.salutation || '', g.firstName || '', g.surname || '', memberEmail,
        '', '', g.provRank || '', getProvAbbrev(g.provRank),
        g.grandRank || '', getGrandAbbrev(g.grandRank), g.dietary || '', '',
        '£' + gCost.toFixed(2), 'Confirmed', hostName, gCat, pMethod, '', '', target
      ]);
    });
    if (!added) return { success: false, message: 'No valid guests to add.' };

    // Recalculate the full confirmed total for this code.
    const data2 = sheet.getDataRange().getValues();
    let total = 0, guestNames = [];
    for (let j = 1; j < data2.length; j++) {
      const r = data2[j];
      if (String(r[25]).toUpperCase().trim() === target && String(r[19]).trim() === 'Confirmed') {
        total += parseFloat(String(r[18]).replace(/[^0-9.]/g,'')) || 0;
        if (String(r[5]).trim() === 'Guest') {
          const gn = [r[6], r[7], r[8]].filter(Boolean).join(' ');
          if (gn.trim()) guestNames.push(gn.trim());
        }
      }
    }

    // Notify the member and stewards of the amendment.
    try { sendAmendmentEmail(lead, guests, added, addedCost, total, ref, target); } catch (e) {}

    return { success: true, added: added, addedCost: '£' + addedCost.toFixed(2), total: '£' + total.toFixed(2), guests: guestNames, ref: ref };
  } catch (e) {
    return { error: 'amend_error', message: e.toString() };
  }
}

// Email confirmation of an amendment (added guests) to the member + stewards.
function sendAmendmentEmail(lead, guests, added, addedCost, total, ref, code) {
  const lodge = getLodge();
  const memberEmail = String(lead[9] || '');
  const bookerName = [lead[6], lead[7], lead[8]].filter(Boolean).join(' ');
  const meetingDate = String(lead[3] || '');
  const names = guests.filter(function(g){return g && (g.firstName||g.surname);})
    .map(function(g){return esc_([g.salutation,g.firstName,g.surname].filter(Boolean).join(' '));});
  const list = names.map(function(n){return '<li>'+n+'</li>';}).join('');
  const html = '<div style="font-family:Georgia,serif;color:#1A2744;max-width:520px;margin:0 auto">'
    + '<h2 style="color:#1A2744">Booking Amended</h2>'
    + '<p>Dear ' + esc_(bookerName) + ',</p>'
    + '<p>The following ' + (added===1?'guest has':'guests have') + ' been added to your booking for ' + esc_(meetingDate) + ':</p>'
    + '<ul>' + list + '</ul>'
    + '<p>Added cost: <b>' + esc_(addedCost) + '</b><br>New booking total: <b>' + esc_(total) + '</b></p>'
    + '<p>Your booking reference and code remain: <b>' + esc_(code) + '</b></p>'
    + '<p style="color:#5A6066;font-size:12px">Sent from the ' + esc_(lodge.name) + ' booking system.</p></div>';
  const to = [memberEmail, lodge.stewardEmail].filter(Boolean).join(',');
  if (to) GmailApp.sendEmail(to, lodge.name + ' - Booking Amended (' + ref + ')', '', { htmlBody: html, name: lodge.name + ' - Booking' });
}

function cancelBooking(code, isOfficer, officerWork) {
  try {
    const ss     = getSS();
    const sheet  = ss.getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    const data   = sheet.getDataRange().getValues();
    const config = getConfig();
    const secEmail    = config.success ? config.meeting.secretaryEmail : '';
    const stewEmail   = config.success ? config.meeting.stewardEmail   : '';
    const dcEmail     = config.success ? config.meeting.dcEmail        : '';

    const target = String(code||'').toUpperCase().trim();
    if (!target) return { success: false, message: 'Please enter your cancellation code.' };

    let memberRowMatch = null;
    let memberRowIndex = -1;

    // Find the Lead Booker row for this code, then cancel every row sharing the code (lead + guests)
    for (let i = data.length - 1; i >= 1; i--) {
      const r = data[i];
      if (
        String(r[25]).toUpperCase().trim() === target &&
        String(r[5]).trim() === 'Lead Booker'         &&
        String(r[19]).trim() === 'Confirmed'
      ) { memberRowMatch = r; memberRowIndex = i + 1; break; }
    }

    if (!memberRowMatch) return { success: false, message: 'No confirmed booking found for that code.' };

    let cancelledRefs = [];
    for (let k = 1; k < data.length; k++) {
      if (String(data[k][25]).toUpperCase().trim() === target && String(data[k][19]).trim() === 'Confirmed') {
        sheet.getRange(k + 1, 20).setValue('Cancelled');
        cancelledRefs.push(String(data[k][2]));
      }
    }

    if (memberRowIndex > 0) {
      sheet.getRange(memberRowIndex, 24).setValue(isOfficer === 'Yes' ? 'Yes' : '');
      sheet.getRange(memberRowIndex, 25).setValue(officerWork || '');
    }

    sendCancellationNotificationToStaff(memberRowMatch, isOfficer, officerWork, secEmail, stewEmail, dcEmail, cancelledRefs.length);
    sendCancellationEmail(memberRowMatch, cancelledRefs);

    return { success: true, ref: cancelledRefs[0], refs: cancelledRefs, cancelledCount: cancelledRefs.length };
  } catch(e) { return { error: 'cancel_error', message: e.toString() }; }
}

// ============================================================
// CONTACT FORM
// ============================================================
function submitContact(formData) {
  try {
    const ss   = getSS();
    let sheet  = ss.getSheetByName(CONTACT_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONTACT_SHEET);
      sheet.appendRow(['Timestamp','Name','Phone','Email','Subject','Message','Recipients','Device']);
      sheet.getRange(1,1,1,8).setFontWeight('bold');
    }

    const lodge     = getLodge();
    const secEmail  = lodge.secretaryEmail;
    const stewEmail = lodge.stewardEmail;
    const dcEmail   = lodge.dcEmail;

    const recipientMap    = { secretary: secEmail, steward: stewEmail, dc: dcEmail };
    const recipientLabels = { secretary: 'Secretary', steward: 'Senior Steward', dc: 'Director of Ceremonies' };
    let recipientEmails   = [];

    (formData.recipients || []).forEach(r => {
      if (recipientMap[r] && recipientMap[r].includes('@')) {
        recipientEmails.push({ label: recipientLabels[r], email: recipientMap[r] });
      }
    });

    if (!recipientEmails.length) return { success: false, message: 'No valid recipient email addresses found.' };

    const recipientListStr = recipientEmails.map(r => r.label).join(', ');

    sheet.appendRow([
      new Date(), formData.name||'', formData.phone||'', formData.senderEmail||'',
      formData.subject||'', formData.message||'', recipientListStr, formData.device||'Unknown'
    ]);

    recipientEmails.forEach(rec => {
      const others     = recipientEmails.filter(r => r.email !== rec.email).map(r => r.label).join(' and ');
      const othersNote = others ? `<p style="font-size:13px;color:#7A6A5A;margin-top:12px;border-top:1px dashed rgba(0,0,0,.1);padding-top:10px;">This message has also been sent to: <strong>${others}</strong>.</p>` : '';
      const replyTo    = formData.senderEmail && formData.senderEmail.includes('@') ? formData.senderEmail : null;
      const subject = `${lodge.name} - Contact Enquiry: ${formData.subject||'General'} (from ${formData.name||'Unknown'})`;
      const htmlBody = buildContactEmailHtml(formData, othersNote, rec.label);
      const opts = { htmlBody: htmlBody, name: lodge.name + ' - Contact' };
      if (replyTo) opts.replyTo = replyTo;
      GmailApp.sendEmail(rec.email, subject, '', opts);
    });

    if (formData.senderEmail && formData.senderEmail.includes('@')) {
      GmailApp.sendEmail(
        formData.senderEmail,
        `${lodge.name} - Your Enquiry Has Been Sent`,
        '',
        { htmlBody: buildContactCopyHtml(formData, recipientListStr), name: lodge.name + ' - Contact' }
      );
    }

    return { success: true };
  } catch(e) { return { error: 'contact_error', message: e.toString() }; }
}

// ============================================================
// EMAIL HELPERS
// ============================================================
function emailHeader(subtitle) {
  const lodge = getLodge();
  return `<div style="background:#1A2744;padding:22px 24px;text-align:center;">
    <h1 style="color:#D4AF37;margin:0;font-size:19px;letter-spacing:3px;font-family:Georgia,serif;">${(lodge.name||'').toUpperCase()}${lodge.number?(' No. '+lodge.number):''}</h1>
    <p style="color:rgba(212,175,55,0.7);margin:5px 0 0;font-size:11px;letter-spacing:2px;">${subtitle}</p>
  </div>`;
}

function emailFooter() {
  const lodge = getLodge();
  return `<p style="margin-top:22px;font-size:15px;color:#1A0F0A;">Yours fraternally,<br>${lodge.fullName}</p>`;
}

function almsEmailBlock(isApology) {
  const lodge = getLodge();
  const body = isApology
    ? `Although you are unable to join us on this occasion, you remain very much in our thoughts. If you wish, you can still support the Lodge Alms Collection online. Your generosity makes a real difference.`
    : `As part of our move towards a digital Alms Collection, you are warmly invited to make your contribution online ahead of the meeting.`;
  const envelopeNote = isApology
    ? ''
    : `<p style="font-size:12px;color:#7A6A5A;margin-top:12px;">Traditional collection envelopes will also be available on the night for those who prefer them.</p>`;
  return `<div style="background:#ffffff;border:1px solid rgba(212,175,55,0.35);border-radius:4px;padding:18px;margin:20px 0;text-align:center;">
    <img src="https://lh3.googleusercontent.com/d/199nzhpdRIq7rXvSUTDeS-eWKjPIG1XSt" alt="Lodge Alms Collection" width="120" style="max-height:48px;width:auto;margin:0 auto 12px;display:block;">
    <p style="font-family:Georgia,serif;font-size:13px;letter-spacing:2px;text-transform:uppercase;color:#8B1A1A;margin:0 0 10px;font-weight:bold;">Lodge Alms Collection</p>
    <p style="font-size:14px;line-height:1.65;color:#1A0F0A;margin-bottom:14px;">${body}</p>
    <a href="${lodge.almsLink}" target="_blank" style="background-color:#1A2744;color:#ffffff;padding:14px 28px;text-decoration:none;font-size:13px;font-family:Georgia,serif;font-weight:bold;border-radius:4px;display:inline-block;border:2px solid #D4AF37;letter-spacing:1px;mso-padding-alt:0;">Contribute to the Alms Collection</a>
    ${envelopeNote}
  </div>`;
}

function sendBookingConfirmationEmail(formData, ref, totalCost, pMethod, tyleTime, meetingM, stewEmail, cancelCode) {
  if (!formData.memberEmail) return;
  const lodge = getLodge();
  const _pm = String(formData.price == null ? '' : formData.price).replace(/,/g,'').match(/\d+(\.\d+)?/);
  const perHead = '£' + (_pm ? parseFloat(_pm[0]).toFixed(2) : '24.50') + ' per person';
  const isApology  = formData.attendanceType === 'apology';
  const isDining   = formData.attendanceType === 'dining';
  const typeLabel  = {apology:'Apology',meeting:'Meeting only',dining:'Meeting & Festive Board'}[formData.attendanceType] || formData.attendanceType;
  const subject    = isApology
    ? `${lodge.fullName} - Apology Received (${ref})`
    : `${lodge.fullName} - Booking Confirmed (${ref})`;
  const opening    = isApology
    ? `Your apology for the upcoming meeting has been duly recorded. We are sorry you will not be with us on this occasion.`
    : `Thank you for registering your attendance. Your booking has been confirmed - we look forward to seeing you.`;

  let guestRows = '';
  if (isDining && formData.guests && formData.guests.length > 0) {
    guestRows = '<h3 style="color:#8B1A1A;margin:16px 0 8px;font-family:Georgia,serif;font-size:15px;">Guests</h3>';
    formData.guests.forEach((g,i) => {
      const hl = g.homeLodge ? ` - ${g.homeLodge}` : '';
      guestRows += `<p style="margin:4px 0;font-size:14px;color:#1A0F0A;"><strong>Guest ${i+1}:</strong> ${[g.firstName,g.surname].filter(Boolean).join(' ')}${hl} (${g.diningCategory||'Standard Dining'})</p>`;
    });
  }

  // Menu - only for dining
  let menuHtml = '';
  if (isDining && meetingM && (meetingM.starter || meetingM.main || meetingM.dessert)) {
    const hasRealMenu = [meetingM.starter,meetingM.main,meetingM.dessert].some(x => x && x !== 'To be confirmed' && x !== 'To be confirmed');
    if (hasRealMenu) {
      menuHtml = `<h3 style="color:#8B1A1A;margin:20px 0 8px;font-family:Georgia,serif;font-size:14px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(212,175,55,0.25);padding-bottom:4px;">Festive Board Menu</h3>
      <table style="width:100%;font-size:14px;margin-bottom:16px;line-height:1.6;border-collapse:collapse;">
        <tr><td style="padding:5px 0;color:#7A6A5A;width:30%;font-weight:bold;">Starter</td><td style="color:#1A0F0A;">${meetingM.starter||'To be confirmed'}</td></tr>
        <tr><td style="padding:5px 0;color:#7A6A5A;font-weight:bold;">Main Course</td><td style="color:#1A0F0A;">${meetingM.main||'To be confirmed'}</td></tr>
        <tr><td style="padding:5px 0;color:#7A6A5A;font-weight:bold;">Dessert</td><td style="color:#1A0F0A;">${meetingM.dessert||'To be confirmed'}</td></tr>
      </table>`;
    }
  }

  let costHtml = '';
  if (isDining) {
    const payInstr = pMethod === 'BACS'
      ? `Please settle via <strong>BACS Bank Transfer</strong>:<br><br><strong>Account Name:</strong> ${lodge.bankName}<br><strong>Sort Code:</strong> ${lodge.sortCode}<br><strong>Account Number:</strong> ${lodge.accountNumber}<br><strong>Payment Reference:</strong> Please use your name`
      : `You have chosen to pay <strong>by Card on the Night</strong>. Please speak to the Festive Board Steward upon arrival at ${lodge.address||'the venue'} to settle before dining.`;
    costHtml = `<div style="background:#1A2744;color:#fff;padding:18px;margin:16px 0;border-left:4px solid #D4AF37;border-radius:2px;">
      <p style="color:#D4AF37;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 10px;font-family:Georgia,serif;">Payment Details</p>
      <p style="font-size:14px;line-height:1.7;margin:0;color:#fff;">
        <strong style="color:#E8CC68;">Festive Board:</strong> ${perHead}<br>
        <strong style="color:#E8CC68;">Total Due:</strong> £${totalCost.toFixed(2)}<br>
        <strong style="color:#E8CC68;">Category:</strong> ${formData.memberDiningCategory||'Standard Dining'}<br>
        <strong style="color:#E8CC68;">Payment Method:</strong> ${pMethod==='BACS'?'BACS Bank Transfer':'Card on the Night'}
      </p>
      <p style="font-size:14px;line-height:1.7;margin:12px 0 0;border-top:1px dashed rgba(255,255,255,0.2);padding-top:12px;color:#fff;">${payInstr}</p>
    </div>`;
  }

  let apologyHtml = '';
  if (isApology && formData.apologyReason) {
    apologyHtml = `<div style="background:#f9f5ec;border-left:3px solid #8B1A1A;padding:12px 16px;margin:14px 0;font-size:14px;color:#1A0F0A;line-height:1.6;">
      <strong style="display:block;margin-bottom:4px;color:#8B1A1A;">Reason Given:</strong>${formData.apologyReason}
    </div>`;
  }

  let notifiedHtml = '';
  if (isApology) {
    const who = formData.isOfficer === 'Yes'
      ? 'the Director of Ceremonies, Secretary and Senior Steward have been notified.'
      : 'the Secretary and Senior Steward have been notified.';
    notifiedHtml = `<p style="font-size:14px;color:#1A0F0A;line-height:1.6;margin-top:12px;">For your records, ${who}</p>`;
    if (formData.isOfficer === 'Yes' && formData.officerWork) {
      notifiedHtml += `<div style="background:#FCE4D6;border-left:3px solid #8B1A1A;padding:12px 14px;margin:12px 0;font-size:14px;line-height:1.6;color:#1A0F0A;"><strong>Your office / ritual note has been flagged to the Director of Ceremonies:</strong><br>${formData.officerWork}</div>`;
    }
  }

  const cancelLink = !isApology ? `<div style="margin-top:26px;padding:18px;background:#ffffff;border:1px solid rgba(139,26,26,0.3);border-radius:4px;text-align:center;">
    <p style="font-size:12px;letter-spacing:2px;text-transform:uppercase;color:#8B1A1A;margin:0 0 10px;font-family:Georgia,serif;font-weight:bold;">Your Cancellation Code</p>
    <div style="background-color:#1A2744;border:2px solid #D4AF37;border-radius:5px;padding:12px 8px;margin:0 0 12px;"><span style="font-size:28px;letter-spacing:5px;font-family:monospace;font-weight:bold;color:#ffffff;">${cancelCode||''}</span></div>
    <p style="font-size:13px;color:#5A4736;line-height:1.6;margin:0 0 14px;">Keep this code safe. You will need it - and nothing else - to cancel this booking.</p>
    <a href="${lodge.cancelBase}?mode=cancel&code=${encodeURIComponent(cancelCode||'')}" target="_blank" style="background-color:#8B1A1A;color:#ffffff;padding:14px 28px;text-decoration:none;font-size:13px;font-family:Georgia,serif;font-weight:bold;border-radius:4px;display:inline-block;border:2px solid #D4AF37;letter-spacing:1px;mso-padding-alt:0;">Cancel a Booking</a>
    <div style="margin-top:12px;"><a href="${lodge.cancelBase}?mode=amend&code=${encodeURIComponent(cancelCode||'')}" target="_blank" style="background-color:#1A2744;color:#ffffff;padding:12px 26px;text-decoration:none;font-size:13px;font-family:Georgia,serif;font-weight:bold;border-radius:4px;display:inline-block;border:2px solid #D4AF37;letter-spacing:1px;mso-padding-alt:0;">Add a Guest to this Booking</a></div>
    <p style="font-size:12px;color:#7A6A5A;margin:10px 0 0;line-height:1.5;">Use the same code above to add a guest to your booking.</p>
  </div>` : '';

  const cancNote = !isApology ? `<p style="font-size:13px;color:#7A6A5A;margin-top:14px;line-height:1.6;">Please note: cancellations within 4 days of the meeting will incur the full Festive Board charge.</p>` : '';
  const tyleRow  = !isApology ? `<tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Tyling Time</td><td style="padding:7px 0;color:#1A0F0A;font-weight:bold;">${tyleTime}</td></tr>` : '';

  const html = `<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;color:#1A0F0A;">
    ${emailHeader(isApology?'APOLOGY RECEIVED':'BOOKING CONFIRMATION')}
    <div style="padding:24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
      <p style="font-size:15px;">Dear ${formData.memberSalutation?formData.memberSalutation+' ':''}${formData.memberFirstName},</p>
      <p style="font-size:15px;line-height:1.65;">${opening}</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin:14px 0;">
        <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;width:42%;font-weight:bold;">${isApology?'Apology Reference':'Booking Reference'}</td><td style="padding:7px 0;color:#1A0F0A;font-family:monospace;font-size:15px;font-weight:bold;">${ref}</td></tr>
        <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Meeting Date</td><td style="padding:7px 0;color:#1A0F0A;">${formData.meetingDate}</td></tr>
        <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Venue</td><td style="padding:7px 0;color:#1A0F0A;">${lodge.address||''}</td></tr>
        ${tyleRow}
        <tr><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">${isApology?'Status':'Attendance'}</td><td style="padding:7px 0;color:#1A0F0A;">${isApology?'Apology Recorded':typeLabel}</td></tr>
      </table>
      ${apologyHtml}${notifiedHtml}${menuHtml}${guestRows}${costHtml}
      ${almsEmailBlock(isApology)}
      ${cancNote}${cancelLink}
      ${emailFooter()}
    </div>
  </div>`;

  const replyTo = isApology
    ? (formData.isOfficer==='Yes' ? dcEmail||stewEmail : stewEmail)
    : stewEmail;
  const opts = { htmlBody: html, name: isApology ? (lodge.name + ' - Apology') : (lodge.name + ' - Booking') };
  if (replyTo && replyTo.includes('@')) opts.replyTo = replyTo;
  if (!isApology) {
    try {
      const ics = buildICS({
        meetingDate: formData.meetingDate, tyleTime: tyleTime, attendanceType: formData.attendanceType,
        bookerName: ((formData.memberSalutation?formData.memberSalutation+' ':'') + (formData.memberFirstName||'') + ' ' + (formData.memberSurname||'')).trim(),
        ref: ref, cancelCode: cancelCode, totalCost: (isDining ? ('£'+Number(totalCost).toFixed(2)) : '')
      });
      if (ics) opts.attachments = [Utilities.newBlob(ics, 'text/calendar', 'meeting.ics')];
    } catch(e) {}
  }
  GmailApp.sendEmail(formData.memberEmail, subject, '', opts);
}

function sendApologyNotifications(formData, ref, secEmail, stewEmail, dcEmail) {
  const lodge = getLodge();
  const isOfficer  = formData.isOfficer === 'Yes';
  const senderName = `${formData.memberSalutation||''} ${formData.memberFirstName} ${formData.memberSurname}`.trim();
  const allRecipients = isOfficer
    ? [
        { email:dcEmail,   label:'Director of Ceremonies', others:'Secretary and Senior Steward' },
        { email:secEmail,  label:'Secretary',               others:'Director of Ceremonies and Senior Steward' },
        { email:stewEmail, label:'Senior Steward',          others:'Director of Ceremonies and Secretary' }
      ]
    : [
        { email:secEmail,  label:'Secretary',      others:'Senior Steward' },
        { email:stewEmail, label:'Senior Steward', others:'Secretary' }
      ];

  const roleHtml = isOfficer ? `<div style="background:#FCE4D6;border-left:4px solid #8B1A1A;padding:12px 14px;margin:12px 0;font-size:14px;line-height:1.6;color:#1A0F0A;font-family:Georgia,serif;">
    <strong>Action Required - Office / Ritual Role Impacted:</strong><br>${formData.officerWork||'Not specified'}
  </div>` : '';

  const rankLine = [formData.memberProvRank, formData.memberGrandRank].filter(Boolean).join(', ');
  const subjectPfx = isOfficer ? '[ACTION REQUIRED] ' : '';

  allRecipients.forEach(rec => {
    if (!rec.email || !rec.email.includes('@')) return;
    const othersNote = `<p style="font-size:13px;color:#7A6A5A;margin-top:12px;border-top:1px dashed rgba(0,0,0,.1);padding-top:10px;">This notification has also been sent to: <strong>${rec.others}</strong>.</p>`;
    const html = `<div style="font-family:Georgia,serif;max-width:600px;color:#1A0F0A;">
      ${emailHeader('APOLOGY NOTIFICATION')}
      <div style="padding:20px 24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
        <p style="font-size:15px;">Dear ${rec.label},</p>
        <p style="font-size:15px;"><strong>${senderName}</strong> has submitted an apology for the meeting on <strong>${formData.meetingDate}</strong>.</p>
        <table style="width:100%;border-collapse:collapse;font-size:14px;margin:10px 0;">
          <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:6px 0;color:#7A6A5A;width:38%;font-weight:bold;">Reference</td><td style="padding:6px 0;font-family:monospace;">${ref}</td></tr>
          <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:6px 0;color:#7A6A5A;font-weight:bold;">Reason</td><td style="padding:6px 0;">${formData.apologyReason||'None provided'}</td></tr>
          ${rankLine?`<tr><td style="padding:6px 0;color:#7A6A5A;font-weight:bold;">Rank</td><td style="padding:6px 0;">${rankLine}</td></tr>`:''}
        </table>
        ${roleHtml}${othersNote}
      </div>
    </div>`;
    const subject = `${subjectPfx}${lodge.name} - Apology: ${senderName} (${ref})`;
    const opts = { htmlBody: html, name: lodge.name + ' - Apology' };
    if (formData.memberEmail && formData.memberEmail.includes('@')) opts.replyTo = formData.memberEmail;
    GmailApp.sendEmail(rec.email, subject, '', opts);
  });
  // Note: the member receives a single "Apology Received" confirmation from
  // sendBookingConfirmationEmail (which already lists who was notified and the
  // reason given), so no separate member email is sent here.
}

function sendCancellationNotificationToStaff(memberRow, isOfficer, officerWork, secEmail, stewEmail, dcEmail, bookingCount) {
  const lodge = getLodge();
  const name       = `${memberRow[7]} ${memberRow[8]}`.trim();
  const meeting    = cleanSheetDate(memberRow[3]);
  const ref        = memberRow[2];
  const multiNote  = bookingCount > 1 ? `<p style="font-size:14px;color:#8B1A1A;font-weight:bold;">Note: ${bookingCount} bookings were found and all have been cancelled.</p>` : '';
  const roleHtml   = isOfficer==='Yes' ? `<div style="background:#FCE4D6;border-left:4px solid #8B1A1A;padding:12px 14px;margin:12px 0;font-size:14px;line-height:1.6;color:#1A0F0A;font-family:Georgia,serif;"><strong>Action Required - Office / Ritual Role Cancelled:</strong><br>${officerWork||'Not specified'}</div>` : '';
  const subjectPfx = isOfficer==='Yes' ? '[ACTION REQUIRED] ' : '';
  const subject    = `${subjectPfx}${lodge.name} - Cancellation: ${name} (${ref})`;

  const recs = [
    { email:secEmail,  label:'Secretary',      others:'Senior Steward' },
    { email:stewEmail, label:'Senior Steward', others:'Secretary' },
    ...(isOfficer==='Yes' ? [{ email:dcEmail, label:'Director of Ceremonies', others:'Secretary and Senior Steward' }] : [])
  ];

  recs.forEach(rec => {
    if (!rec.email || !rec.email.includes('@')) return;
    const othersNote = `<p style="font-size:13px;color:#7A6A5A;margin-top:12px;border-top:1px dashed rgba(0,0,0,.1);padding-top:10px;">This notification has also been sent to: <strong>${rec.others}</strong>.</p>`;
    const html = `<div style="font-family:Georgia,serif;max-width:600px;color:#1A0F0A;">
      ${emailHeader('BOOKING CANCELLATION')}
      <div style="padding:20px 24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
        <p style="font-size:15px;">Dear ${rec.label},</p>
        <p style="font-size:15px;"><strong>${name}</strong> has cancelled their booking for the meeting on <strong>${meeting}</strong>.</p>
        <p style="font-size:14px;">Reference: <strong>${ref}</strong></p>
        ${multiNote}${roleHtml}${othersNote}
      </div>
    </div>`;
    GmailApp.sendEmail(rec.email, subject, '', { htmlBody:html, name:lodge.name + ' - Booking' });
  });
}

function sendCancellationEmail(row, allRefs) {
  const lodge = getLodge();
  if (!row) return;
  const email = String(row[9]);
  if (!email.includes('@')) return;
  const multiNote = allRefs.length > 1 ? `<p style="font-size:14px;color:#8B1A1A;line-height:1.6;">We found <strong>${allRefs.length} bookings</strong> under your name for this meeting - all have been cancelled.</p>` : '';
  const html = `<div style="font-family:Georgia,serif;max-width:600px;color:#1A0F0A;">
    ${emailHeader('BOOKING CANCELLED')}
    <div style="padding:24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
      <p style="font-size:15px;">Dear ${row[7]},</p>
      <p style="font-size:15px;line-height:1.65;">Your booking for the meeting on <strong>${cleanSheetDate(row[3])}</strong> has been successfully cancelled.</p>
      ${multiNote}
      <div style="background:#FCE4D6;border-left:3px solid #8B1A1A;padding:12px 14px;margin:14px 0;font-size:14px;line-height:1.6;color:#1A0F0A;">If you added this meeting to your calendar, please remember to <strong>delete the calendar event</strong>, as it will not be removed automatically.</div>
      <p style="font-size:14px;color:#7A6A5A;line-height:1.6;">Please note: if this cancellation was made within 4 days of the meeting, the Festive Board charge will still be applied.</p>
      ${almsEmailBlock(false)}
      ${emailFooter()}
    </div>
  </div>`;
  GmailApp.sendEmail(email, `${lodge.fullName} - Booking Cancelled (${row[2]})`, '', { htmlBody:html, name:lodge.name + ' - Booking' });
}

function buildContactEmailHtml(formData, othersNote, recipientLabel) {
  return `<div style="font-family:Georgia,serif;max-width:600px;color:#1A0F0A;">
    ${emailHeader('CONTACT ENQUIRY')}
    <div style="padding:20px 24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
      <p style="font-size:15px;">Dear ${recipientLabel},</p>
      <p style="font-size:15px;">A new enquiry has been submitted via the lodge booking system.</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin:14px 0;">
        <tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;width:35%;font-weight:bold;">From</td><td style="padding:7px 0;">${formData.name||'Not provided'}</td></tr>
        ${formData.phone?`<tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Phone</td><td style="padding:7px 0;">${formData.phone}</td></tr>`:''}
        ${formData.senderEmail?`<tr style="border-bottom:1px solid rgba(212,175,55,0.15);"><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Email</td><td style="padding:7px 0;">${formData.senderEmail}</td></tr>`:''}
        <tr><td style="padding:7px 0;color:#7A6A5A;font-weight:bold;">Subject</td><td style="padding:7px 0;">${formData.subject||'General Enquiry'}</td></tr>
      </table>
      <div style="background:#f0ebe0;border-left:3px solid #D4AF37;padding:14px 16px;margin:14px 0;font-size:14px;line-height:1.7;color:#1A0F0A;">
        <strong style="display:block;margin-bottom:6px;color:#8B1A1A;">Message:</strong>${formData.message||''}
      </div>
      ${othersNote}
    </div>
  </div>`;
}

function buildContactCopyHtml(formData, recipientListStr) {
  const lodge = getLodge();
  return `<div style="font-family:Georgia,serif;max-width:600px;color:#1A0F0A;">
    ${emailHeader('ENQUIRY RECEIVED')}
    <div style="padding:24px;background:#FDF8F0;border:1px solid rgba(212,175,55,0.3);">
      <p style="font-size:15px;">Dear ${formData.name||'there'},</p>
      <p style="font-size:15px;line-height:1.65;">Thank you for getting in touch with ${lodge.name}. Your message has been sent to: <strong>${recipientListStr}</strong>. We will be in touch shortly.</p>
      <div style="background:#f0ebe0;border-left:3px solid #D4AF37;padding:14px 16px;margin:14px 0;font-size:14px;line-height:1.7;color:#1A0F0A;">
        <strong style="display:block;margin-bottom:6px;color:#8B1A1A;">Your message:</strong>${formData.message||''}
      </div>
      ${emailFooter()}
    </div>
  </div>`;
}

// ============================================================
// TABLE PLAN - save / load seating for a meeting
// Stored in a "Table Plan" tab, keyed by meeting + seat id.
// Saves the PERSON NAME against each seat so a plan survives
// row shifts in the Responses sheet; the front end re-matches
// names to people on load.
// ============================================================
function saveTablePlan(password, meetingName, payload) {
  const lodge = getLodge();
  const real = String(lodge.tableAdminPassword || lodge.adminPassword || '').trim();
  if (!real) return { error: 'no_password', message: 'No Table Admin password is set in the Set Up tab.' };
  if (String(password || '').trim() !== real) return { error: 'bad_password', message: 'Incorrect Table Admin password.' };
  try {
    const ss = getSS();
    let sheet = ss.getSheetByName('Table Plan');
    if (!sheet) {
      sheet = ss.insertSheet('Table Plan');
      sheet.appendRow(['Meeting', 'Top', 'Legs', 'SeatId', 'PersonName', 'SavedAt']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(meetingName).trim()) sheet.deleteRow(i + 1);
    }
    const place = (payload && payload.placement) || {};
    const now = new Date();
    const rows = [];
    for (const seat in place) {
      rows.push([meetingName, (payload && payload.top) || '', (payload && payload.legs) || '', seat, place[seat], now]);
    }
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return { success: true, saved: rows.length };
  } catch (e) {
    return { error: 'save_error', message: e.toString() };
  }
}

// Verify the Table Admin password (used to unlock plan editing).
function checkTableAdmin(password) {
  const lodge = getLodge();
  const real = String(lodge.tableAdminPassword || lodge.adminPassword || '').trim();
  return { ok: !!real && String(password || '').trim() === real };
}

// Live attendance write-back. field = 'paid' | 'arrived' | 'notes'.
// Writes to Responses columns AA (27) Paid, AB (28) Arrived, AC (29) Notes.
function setAttendance(password, rowIndex, field, value) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real || String(password || '').trim() !== real) return { error: 'bad_password', message: 'Not authorised.' };
  try {
    const sheet = getSS().getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    // ensure headers exist on AA/AB/AC
    const head = sheet.getRange(1, 27, 1, 3).getValues()[0];
    if (!head[0]) sheet.getRange(1, 27).setValue('Paid');
    if (!head[1]) sheet.getRange(1, 28).setValue('Arrived');
    if (!head[2]) sheet.getRange(1, 29).setValue('Notes');
    const col = field === 'paid' ? 27 : (field === 'arrived' ? 28 : 29);
    const r = parseInt(rowIndex, 10);
    if (!r || r < 2) return { error: 'bad_row', message: 'Invalid row.' };
    sheet.getRange(r, col).setValue(value);
    return { success: true };
  } catch (e) {
    return { error: 'attend_error', message: e.toString() };
  }
}

function getTablePlan(meetingName) {
  try {
    const sheet = getSS().getSheetByName('Table Plan');
    if (!sheet) return { success: true, placement: {}, top: 5, legs: 2 };
    const data = sheet.getDataRange().getValues();
    const place = {};
    let top = 5, legs = 2;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(meetingName).trim()) {
        top = data[i][1] || top;
        legs = data[i][2] || legs;
        place[String(data[i][3])] = String(data[i][4]);   // seatId -> personName
      }
    }
    return { success: true, placement: place, top: Number(top) || 5, legs: Number(legs) || 2 };
  } catch (e) {
    return { success: false, placement: {}, top: 5, legs: 2 };
  }
}

// ============================================================
// EMAIL REPORTS - send chosen attachments to chosen recipients.
// attachments: [{name, mimeType, dataBase64}] built on the client
//   (CSV/PDF generated in the browser, plus any user uploads).
// recipients: { dc:bool, steward:bool, secretary:bool, extra:'a@b.com, c@d.com' }
// ============================================================
function emailReports(password, meetingName, recipients, attachments, note) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real || String(password || '').trim() !== real) return { error: 'bad_password', message: 'Not authorised.' };
  try {
    const to = [];
    if (recipients && recipients.dc && lodge.dcEmail) to.push(lodge.dcEmail);
    if (recipients && recipients.steward && lodge.stewardEmail) to.push(lodge.stewardEmail);
    if (recipients && recipients.secretary && lodge.secretaryEmail) to.push(lodge.secretaryEmail);
    if (recipients && recipients.extra) {
      String(recipients.extra).split(/[,;]+/).forEach(function (e) {
        e = e.trim(); if (e && /@/.test(e)) to.push(e);
      });
    }
    if (!to.length) return { error: 'no_recipients', message: 'No valid recipients selected.' };

    const blobs = (attachments || []).map(function (a) {
      const bytes = Utilities.base64Decode(a.dataBase64);
      return Utilities.newBlob(bytes, a.mimeType || 'application/octet-stream', a.name || 'attachment');
    });

    const subject = lodge.name + ' - Reports - ' + (meetingName || '');
    const body = 'Please find attached the requested reports for ' + (meetingName || 'the meeting') + '.' +
      (note ? '\n\n' + note : '') + '\n\nSent from the ' + lodge.name + ' booking system.';
    GmailApp.sendEmail(to.join(','), subject, body, { attachments: blobs, name: lodge.name });
    return { success: true, sent: to.length };
  } catch (e) {
    return { error: 'email_error', message: e.toString() };
  }
}

// ============================================================
// PURGE OLD ENTRIES - delete Responses rows older than the most
// recent N meetings (default 4). Never touches future meetings,
// never any other tab. Gated by the Table Admin password.
// ============================================================
function purgeOldMeetings(password, keepCount) {
  const lodge = getLodge();
  const real = String(lodge.tableAdminPassword || lodge.adminPassword || '').trim();
  if (!real || String(password || '').trim() !== real) return { error: 'bad_password', message: 'Incorrect Table Admin password.' };
  try {
    const keep = Math.max(1, parseInt(keepCount, 10) || 4);
    const sheet = getSS().getSheetByName(RESPONSES_SHEET);
    if (sheet.getFilter()) sheet.getFilter().remove();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, deleted: 0, keptMeetings: 0 };
    const MONTHS = {january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
    function toDate(val) {
      if (val instanceof Date && !isNaN(val.getTime())) return val;
      const s = String(val||'').trim(); if (!s) return null;
      const m = s.toLowerCase().match(/(\d{1,2})\s+([a-z]+)\s+(\d{4})/);
      if (m && MONTHS[m[2]] != null) return new Date(parseInt(m[3],10), MONTHS[m[2]], parseInt(m[1],10));
      const d = new Date(s); return isNaN(d.getTime()) ? null : d;
    }
    const meetingDates = {};
    for (let i = 1; i < data.length; i++) {
      const raw = data[i][3];
      const key = String(raw||'').toLowerCase().replace(/[^a-z0-9]/g,'');
      if (!key) continue;
      if (!meetingDates[key]) meetingDates[key] = { date: toDate(raw), key: key };
    }
    const now = new Date(); now.setHours(0,0,0,0);
    const distinct = Object.keys(meetingDates).map(function(k){return meetingDates[k];})
      .filter(function(m){return m.date;}).sort(function(a,b){return b.date - a.date;});
    const keepKeys = {};
    distinct.forEach(function(m){ if (m.date >= now) keepKeys[m.key] = true; });
    distinct.filter(function(m){return m.date < now;}).slice(0, keep).forEach(function(m){ keepKeys[m.key] = true; });
    let deleted = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      const key = String(data[i][3]||'').toLowerCase().replace(/[^a-z0-9]/g,'');
      if (key && !keepKeys[key]) { sheet.deleteRow(i + 1); deleted++; }
    }
    return { success: true, deleted: deleted, keptMeetings: Object.keys(keepKeys).length };
  } catch (e) {
    return { error: 'purge_error', message: e.toString() };
  }
}

// Shared recipient resolver for report emails.
function resolveRecipients(lodge, recipients) {
  const to = [];
  if (recipients && recipients.dc && lodge.dcEmail) to.push(lodge.dcEmail);
  if (recipients && recipients.steward && lodge.stewardEmail) to.push(lodge.stewardEmail);
  if (recipients && recipients.secretary && lodge.secretaryEmail) to.push(lodge.secretaryEmail);
  if (recipients && recipients.extra) {
    String(recipients.extra).split(/[,;]+/).forEach(function (e) { e = e.trim(); if (e && /@/.test(e)) to.push(e); });
  }
  return to;
}

// Render a table-plan HTML (built on the client) to a PDF and email it.
function emailTablePlanPdf(password, meetingName, planHtml, recipients, note) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real || String(password || '').trim() !== real) return { error: 'bad_password', message: 'Not authorised.' };
  try {
    const to = resolveRecipients(lodge, recipients);
    if (!to.length) return { error: 'no_recipients', message: 'No valid recipients selected.' };
    const html = '<html><head><meta charset="utf-8"><style>body{font-family:Helvetica,Arial,sans-serif;color:#15202B;margin:0;padding:14px}</style></head><body>' + (planHtml || '<p>No plan.</p>') + '</body></html>';
    const blob = htmlToPdfBlob(html, 'Table_Plan_' + String(meetingName).replace(/[^a-z0-9]+/gi,'_'));
    const subject = lodge.name + ' - Table Plan - ' + (meetingName || '');
    const body = 'Please find attached the table plan for ' + (meetingName || 'the meeting') + '.' + (note ? '\n\n' + note : '') + '\n\nSent from the ' + lodge.name + ' booking system.';
    GmailApp.sendEmail(to.join(','), subject, body, { attachments: [blob], name: lodge.name });
    return { success: true, sent: to.length };
  } catch (e) {
    return { error: 'pdf_error', message: e.toString() };
  }
}

// Final attendance & payment report (paid & attended) as a PDF.
// Reliable HTML -> PDF with no advanced services required.
// DriveApp renders the HTML file through Drive's converter when getAs('application/pdf')
// is called, unlike Utilities.newBlob(...).getAs(...) which often yields a blank page.
function htmlToPdfBlob(html, filename) {
  var file = null;
  try {
    var htmlBlob = Utilities.newBlob(html, 'text/html', (filename || 'doc') + '.html');
    file = DriveApp.createFile(htmlBlob);
    var pdf = file.getAs('application/pdf').setName((filename || 'document') + '.pdf');
    // copy bytes so we can delete the temp file safely
    var bytes = pdf.getBytes();
    return Utilities.newBlob(bytes, 'application/pdf', (filename || 'document') + '.pdf');
  } finally {
    if (file) { try { file.setTrashed(true); } catch (e) {} }
  }
}

function emailFinalReport(password, meetingName, recipients, note, liveRows) {
  const lodge = getLodge();
  const real = String(lodge.adminPassword || '').trim();
  if (!real || String(password || '').trim() !== real) return { error: 'bad_password', message: 'Not authorised.' };
  try {
    const to = resolveRecipients(lodge, recipients);
    if (!to.length) return { error: 'no_recipients', message: 'No valid recipients selected.' };

    let rows = [];
    if (liveRows && liveRows.length) {
      // Use live data passed directly from the UI (reflects current checked state)
      rows = liveRows.map(function(r) {
        var nameParts = String(r.name||'').trim().split(' ');
        return { name: r.name, surname: nameParts[nameParts.length-1]||'', paid: !!r.paid, arrived: !!r.arrived, pay: r.bacs ? 'BACS' : (r.paid ? 'Paid' : '') };
      });
    } else {
      // Fallback: read from spreadsheet
      const sheet = getSS().getSheetByName(RESPONSES_SHEET);
      if (sheet.getFilter()) sheet.getFilter().remove();
      const data = sheet.getDataRange().getValues();
      const wantMeet = String(meetingName||'').toLowerCase().replace(/[^a-z0-9]/g,'');
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        if (/cancel/i.test(String(r[19]))) continue;
        if (/apolog/i.test(String(r[4]))) continue;
        const haveMeet = String(r[3]||'').toLowerCase().replace(/[^a-z0-9]/g,'');
        if (haveMeet !== wantMeet) continue;
        const isBacs = /bacs|bank|transfer/i.test(String(r[22]||''));
        const paid = isBacs || /^y/i.test(String(r[26]||''));
        const arrived = /^y/i.test(String(r[27]||''));
        rows.push({ name: ((r[6]?r[6]+' ':'') + (r[7]||'') + ' ' + (r[8]||'')).trim(), surname: String(r[8]||''), paid: paid, arrived: arrived, pay: isBacs ? 'BACS' : String(r[22]||'') });
      }
    }
    rows.sort(function(a,b){return a.surname.localeCompare(b.surname);});
    const paidAttended = rows.filter(function(r){return r.paid && r.arrived;});
    const outstanding = rows.filter(function(r){return !r.paid || !r.arrived;});
    function tableRows(list){
      return list.map(function(r){
        return '<tr><td>'+esc_(r.name)+'</td><td style="text-align:center">'+(r.paid?'&#10003;':'&mdash;')+'</td><td style="text-align:center">'+(r.arrived?'&#10003;':'&mdash;')+'</td><td>'+esc_(r.pay)+'</td></tr>';
      }).join('');
    }
    const html = '<html><head><meta charset="utf-8"><style>'
      + 'body{font-family:Helvetica,Arial,sans-serif;color:#15202B;padding:18px}'
      + 'h1{font-size:18px;margin:0 0 2px}h2{font-size:13px;color:#5A6066;margin:16px 0 6px}'
      + '.sub{color:#5A6066;font-size:12px;margin-bottom:10px}'
      + 'table{width:100%;border-collapse:collapse;font-size:12px}'
      + 'th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}'
      + 'th{background:#1F3C6E;color:#fff}'
      + '.tot{margin-top:8px;font-size:12px;color:#5A6066}'
      + '</style></head><body>'
      + '<h1>' + esc_(lodge.name) + '</h1>'
      + '<div class="sub">Festive Board Attendance &amp; Payment Report &mdash; ' + esc_(meetingName) + '</div>'
      + '<div class="tot"><b>' + paidAttended.length + '</b> paid &amp; attended &middot; <b>' + outstanding.length + '</b> outstanding &middot; <b>' + rows.length + '</b> total</div>'
      + '<h2>Paid &amp; Attended (' + paidAttended.length + ')</h2><table><tr><th>Name</th><th>Paid</th><th>Here</th><th>Payment</th></tr>' + (tableRows(paidAttended)||'<tr><td colspan="4">None yet</td></tr>') + '</table>'
      + '<h2>Outstanding (' + outstanding.length + ')</h2><table><tr><th>Name</th><th>Paid</th><th>Here</th><th>Payment</th></tr>' + (tableRows(outstanding)||'<tr><td colspan="4">None</td></tr>') + '</table>'
      + (note ? '<p style="margin-top:14px;font-size:12px">'+esc_(note)+'</p>' : '')
      + '</body></html>';
    const blob = htmlToPdfBlob(html, 'Attendance_Report_' + String(meetingName).replace(/[^a-z0-9]+/gi,'_'));
    const subject = lodge.name + ' - Attendance Report - ' + (meetingName || '');
    const body = 'Attendance and payment report for ' + (meetingName || 'the meeting') + '.\n\nPaid & attended: ' + paidAttended.length + ' of ' + rows.length + '.' + (note ? '\n\n' + note : '') + '\n\nSent from the ' + lodge.name + ' booking system.';
    GmailApp.sendEmail(to.join(','), subject, body, { attachments: [blob], name: lodge.name });
    return { success: true, sent: to.length, paidAttended: paidAttended.length, total: rows.length };
  } catch (e) {
    return { error: 'report_error', message: e.toString() };
  }
}

function esc_(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }


// ============================================================
// UTILITIES
// ============================================================
// ------------------------------------------------------------
// CALENDAR (.ics) - builds a universal calendar event.
// Event starts 30 min before tyling, ends 5 hours after tyling (but never
// past 23:59 same day). Title = Lodge Name + tyling time. Notes = details.
// meetingDate is the sheet's date; tyleTime like "17:15". Returns ICS string.
// ------------------------------------------------------------
function buildICS(opts) {
  const lodge = getLodge();
  const dateStr = String(opts.meetingDate || '');
  const tyle = String(opts.tyleTime || '').trim();
  // Parse the meeting date into Y/M/D
  let dt = new Date(dateStr);
  if (isNaN(dt.getTime())) {
    const m = dateStr.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
    if (m) dt = new Date(m[2] + ' ' + m[1] + ', ' + m[3]);
  }
  if (isNaN(dt.getTime())) return '';
  // Tyling hour/min
  let th = 18, tm = 0;
  const tm2 = tyle.match(/(\d{1,2})[:.](\d{2})/);
  if (tm2) { th = parseInt(tm2[1],10); tm = parseInt(tm2[2],10); }
  const y = dt.getFullYear(), mo = dt.getMonth(), d = dt.getDate();
  // Start = 30 min before tyling
  let start = new Date(y, mo, d, th, tm, 0); start.setMinutes(start.getMinutes() - 30);
  // End = 5 hours after tyling, capped at 23:59 same day
  let end = new Date(y, mo, d, th, tm, 0); end.setHours(end.getHours() + 5);
  const midnight = new Date(y, mo, d, 23, 59, 0);
  if (end > midnight) end = midnight;
  function fmt(x){
    function p(n){return (n<10?'0':'')+n;}
    return x.getFullYear()+p(x.getMonth()+1)+p(x.getDate())+'T'+p(x.getHours())+p(x.getMinutes())+'00';
  }
  function esc(s){return String(s||'').replace(/\\/g,'\\\\').replace(/;/g,'\\;').replace(/,/g,'\\,').replace(/\n/g,'\\n');}
  const title = lodge.fullName + (tyle ? (' - Tyles ' + tyle) : '');
  const descLines = [];
  if (opts.attendanceType) descLines.push('Booking: ' + ({apology:'Apology',meeting:'Meeting only',dining:'Meeting & Festive Board'}[opts.attendanceType]||opts.attendanceType));
  if (opts.bookerName) descLines.push('Name: ' + opts.bookerName);
  if (tyle) descLines.push('Tyling time: ' + tyle);
  if (opts.ref) descLines.push('Reference: ' + opts.ref);
  if (opts.cancelCode) descLines.push('Cancellation code: ' + opts.cancelCode);
  if (opts.guests) descLines.push('Guests: ' + opts.guests);
  if (opts.totalCost) descLines.push('Festive Board balance: ' + opts.totalCost);
  descLines.push('');
  descLines.push('Please arrive in good time before tyling.');
  const uid = (opts.cancelCode || opts.ref || String(Date.now())) + '@' + (lodge.number || 'lodge');
  return [
    'BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//Lodge Booking//EN','CALSCALE:GREGORIAN','METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + fmt(new Date()),
    'DTSTART:' + fmt(start),
    'DTEND:' + fmt(end),
    'SUMMARY:' + esc(title),
    'LOCATION:' + esc(lodge.address),
    'DESCRIPTION:' + esc(descLines.join('\n')),
    'END:VEVENT','END:VCALENDAR'
  ].join('\r\n');
}

function getICSForBooking(opts) {
  try { return { success: true, ics: buildICS(opts || {}) }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

function generateRef(last, meetingDate) {
  const surname = String(last||'GUEST').toUpperCase().replace(/[^A-Z]/g,'') || 'GUEST';
  const s = String(meetingDate||'');
  const monthNames = [['JANUARY','JAN'],['FEBRUARY','FEB'],['MARCH','MAR'],['APRIL','APR'],['MAY','MAY'],['JUNE','JUN'],['JULY','JUL'],['AUGUST','AUG'],['SEPTEMBER','SEP'],['OCTOBER','OCT'],['NOVEMBER','NOV'],['DECEMBER','DEC']];
  let mon = '', yr = '';
  const up = s.toUpperCase();
  for (let i = 0; i < monthNames.length; i++) {
    if (up.indexOf(monthNames[i][0]) > -1) { mon = monthNames[i][1]; break; }
  }
  const yrMatch = s.match(/\b(20\d{2})\b/);   // find a 4-digit year
  if (yrMatch) yr = yrMatch[1].slice(-2);
  return (mon && yr) ? surname + '-' + mon + yr : surname;
}

// Unique cancellation code - short, easy to type, avoids ambiguous chars (no O/0/I/1).
// Checked against the sheet to guarantee uniqueness.
function generateCancelCode(sheet) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const data = (sheet ? sheet.getDataRange().getValues() : []);
  const existing = {};
  for (let i = 1; i < data.length; i++) { if (data[i][25]) existing[String(data[i][25]).trim().toUpperCase()] = true; }
  for (let attempt = 0; attempt < 50; attempt++) {
    let code = '';
    for (let k = 0; k < 6; k++) code += chars.charAt(Math.floor(Math.random() * chars.length));
    if (!existing[code]) return code;
  }
  return 'C' + Date.now().toString(36).toUpperCase();   // fallback, still unique
}


function getProvAbbrev(rank) {
  if (!rank||!rank.trim()||rank.trim()==='None') return '';
  let str=rank.trim(), isPast=false;
  if(str.startsWith('Past ')){ isPast=true; str=str.substring(5); }
  const map={'Provincial Grand Master':'ProvGM','Deputy Provincial Grand Master':'DepProvGM','Assistant Provincial Grand Master':'AsstProvGM','Provincial Senior Grand Warden':'ProvSGW','Provincial Junior Grand Warden':'ProvJGW','Provincial Grand Chaplain':'ProvGCh','Provincial Grand Treasurer':'ProvGT','Provincial Grand Registrar':'ProvGReg','Provincial Grand Secretary':'ProvGSec','Provincial Grand Director of Ceremonies':'ProvGDC','Provincial Grand Sword Bearer':'ProvGSB','Provincial Grand Superintendent of Works':'ProvGSuptWks','Provincial Grand Almoner':'ProvGAlm','Provincial Grand Charity Steward':'ProvGChStwd','Provincial Grand Mentor':'ProvGMentor','Provincial Grand Learning and Development Officer':'ProvGLDO','Provincial Grand Membership Officer':'ProvGMemO','Provincial Grand Communications Officer':'ProvGCommsO','Provincial Senior Grand Deacon':'ProvSGD','Provincial Junior Grand Deacon':'ProvJGD','Provincial Grand Organist':'ProvGOrg','Provincial Grand Standard Bearer':'ProvGStB','Provincial Grand Tyler':'ProvGTyler','Other Active Provincial Rank':'ProvRank','Other Past Provincial Rank':'PProvRank','Assistant Provincial Grand Almoner':'AsstProvGAlm','Assistant Provincial Grand Charity Steward':'AsstProvGChStwd'};
  return isPast?'P'+(map[str]||str):(map[str]||str);
}

function getGrandAbbrev(rank) {
  if (!rank||!rank.trim()||rank.trim()==='None') return '';
  const map={'Grand Master':'GM','Pro Grand Master':'ProGM','Deputy Grand Master':'DepGM','Assistant Grand Master':'AsstGM','Senior Grand Warden':'SGW','Junior Grand Warden':'JGW','Grand Chaplain':'GCh','Grand Registrar':'GReg','Grand Secretary':'GSec','Grand Director of Ceremonies':'GDC','Grand Sword Bearer':'GSB','Grand Superintendent of Works':'GSuptWks','Grand Inspector':'GInsp','Senior Grand Deacon':'SGD','Junior Grand Deacon':'JGD','Grand Organist':'GOrg','Grand Standard Bearer':'GStB','Grand Pursuivant':'GPurs','Past Grand Officer':'PGOff'};
  return map[rank.trim()]||rank;
}

function formatTime(val) {
  if(!val) return '18:15';
  if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  return String(val).trim().substring(0,5);
}

function formatCurrency(val) {
  if(!val) return '£24.50';
  const n=parseFloat(String(val).replace(/[^0-9.]/g,''));
  return isNaN(n)?String(val):'£'+n.toFixed(2);
}

function cleanSheetDate(val) {
  if(!val) return '';
  if(val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'EEEE, d MMMM yyyy');
  return String(val).trim();
}

function parseUKDate(str) {
  if(!str) return null;
  const parts=String(str).split('/');
  if(parts.length!==3) return new Date(str);
  return new Date(parseInt(parts[2]),parseInt(parts[1])-1,parseInt(parts[0]),0,0,0,0);
}
