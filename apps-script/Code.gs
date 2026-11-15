/**
 * Day & Night — Matcha & Coffee
 * الخادم: Google Apps Script مربوط بملف قوقل شيت.
 * التركيب: الصق هذا الملف في (Extensions > Apps Script)، شغّل الدالة setup مرة وحدة،
 * بعدها Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * تطوير: نرجس العمري
 */

const SH = {
  MENU: 'المنيو',
  ADD: 'الإضافات',
  MILK: 'الحليب',
  SET: 'الإعدادات',
  ORD: 'الطلبات',
  STAR: 'النجوم'
};
const MILK_MODES = ['اختيار بسعر', 'اختيار مجاني', 'أوريجينال فقط', 'بدون'];

/* =============== التركيب الأول =============== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, head, rows) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(head);
    if (sh.getLastRow() <= 1 && rows && rows.length) {
      const fixed = rows.map(r => { const x = r.slice(0, head.length); while (x.length < head.length) x.push(''); return x; });
      sh.getRange(2, 1, fixed.length, head.length).setValues(fixed);
    }
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    return sh;
  };
  make(SH.MENU, ['المعرف', 'الصنف', 'السعر', 'متوفر', 'الوصف', 'شارة', 'الصورة', 'الحليب', 'الإضافات'], [
    ['dn', 'D&N Matcha', 24, 'نعم', 'ماتشا فاخرة بكريمة ونجوم فضية', 'Best Seller', 'dn.webp', 'اختيار مجاني', 'extra, syrup, sweet, blue, pink, party'],
    ['classic', 'Classic Matcha', 20, 'نعم', 'ماتشا كلاسيك بالثلج', '', 'classic.webp', 'اختيار بسعر', 'extra, syrup, sweet, blue, pink, party'],
    ['icecream', 'Ice Cream Matcha', 24, 'نعم', 'ماتشا مع آيس كريم ماتشا', 'جديد', 'icecream.webp', 'أوريجينال فقط', 'syrup'],
    ['bday', 'Celebration Matcha', 35, 'نعم', 'كوب ميلاد بصورتك وزينتك', 'للمناسبات', 'bday.webp', 'بدون', ''],
    ['drip', 'Ice Drip', 20, 'نعم', 'قهوة مقطّرة باردة', '', 'drip.webp', 'بدون', ''],
    ['v60', 'Hot V60', 18, 'نعم', 'قهوة مختصة حارة', '', 'v60.webp', 'بدون', '']
  ]);
  make(SH.ADD, ['المعرف', 'الإضافة', 'السعر', 'متوفر'], [
    ['extra', 'Extra Matcha', 3, 'نعم'],
    ['syrup', 'Syrup', 2, 'نعم'],
    ['sweet', 'Sweet Foam', 2, 'نعم'],
    ['blue', 'Blue Foam', 2, 'نعم'],
    ['pink', 'Pink Foam', 2, 'نعم'],
    ['party', 'Party Pick', 1, 'نعم']
  ]);
  make(SH.MILK, ['المعرف', 'الحليب', 'السعر', 'متوفر'], [
    ['original', 'Original Milk', 0, 'نعم'],
    ['vanilla', 'Vanilla Milk', 2, 'نعم'],
    ['coconut', 'Coconut Milk', 2, 'نعم'],
    ['oat', 'Oat Milk', 2, 'نعم']
  ]);
  make(SH.SET, ['الإعداد', 'القيمة', 'ملاحظة'], [
    ['اسم المحل', 'Day & Night', ''],
    ['يفتح الساعة', 15, 'بنظام 24 ساعة (15 = 3 العصر)'],
    ['يسكر الساعة', 22, 'بنظام 24 ساعة (22 = 10 الليل)'],
    ['يوم الإجازة', '', 'اكتبي اسم اليوم، أو اتركيه فاضي'],
    ['استقبال الطلبات', 'نعم', 'اكتبي «لا» لإيقاف الطلبات مؤقتًا'],
    ['رسالة البانر', '', 'رسالة تظهر فوق المنيو، اتركيها فاضية لإخفائها'],
    ['لما يخلص صنف', 'يختفي', 'يختفي / يظهر باهت'],
    ['نجوم المشروب المجاني', 3, 'عدد النجوم، والمشروب اللي بعدها علينا'],
    ['مكان الاستلام', 'حي الملك فهد، المدينة المنورة', 'يظهر للعميل'],
    ['رابط اللوكيشن', '', 'رابط قوقل ماب، يظهر للعميل بعد ما يطلب'],
    ['بيانات التحويل', '', 'اسم البنك واسم صاحب الحساب والآيبان، تظهر لما العميل يختار تحويل'],
    ['رقم الواتساب', '966538897274', 'بالصيغة الدولية بدون + (للتواصل)'],
    ['كلمة مرور اللوحة', 'dn' + Math.floor(1000 + Math.random() * 9000), 'غيّريها لكلمة تعرفونها أنتم بس']
  ]);
  make(SH.ORD, ['رقم الطلب', 'التاريخ', 'وقت الاستلام', 'الاسم', 'الجوال', 'الطلب', 'المبلغ', 'الدفع', 'الحالة', 'الأكواب', 'ملاحظات', 'صورة الميلاد', 'مكافأة النجوم']);
  make(SH.STAR, ['الجوال', 'الاسم', 'النجوم', 'آخر طلب']);
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('الورقة1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  photoFolder_();
  tidy();
}

/* =============== ترتيب الشيت وتنسيقه =============== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('✦ Day & Night').addItem('ترتيب الشيت', 'tidy').addToUi();
}

function tidy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const INK = '#2f4255', BAND = '#eef3f7';
  const list = (vals) => SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build();
  const style = (name, widths, notes) => {
    const sh = ss.getSheetByName(name); if (!sh) return null;
    const cols = widths.length, rows = Math.max(sh.getMaxRows(), 200);
    if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
    widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setRightToLeft(true); sh.setFrozenRows(1); sh.setHiddenGridlines(true);
    const all = sh.getRange(1, 1, rows, cols);
    all.setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('right').setFontSize(11);
    sh.setRowHeight(1, 38);
    sh.getBandings().forEach(b => b.remove());
    all.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false).setHeaderRowColor(INK).setFirstRowColor('#ffffff').setSecondRowColor(BAND);
    sh.getRange(1, 1, 1, cols).setFontWeight('bold').setFontColor('#ffffff').setFontSize(12);
    if (notes) notes.forEach((n, i) => { if (n) sh.getRange(1, i + 1).setNote(n); });
    return { sh, rows };
  };
  const cf = (range, pairs) => pairs.map(([txt, bg, fg]) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build());
  const yesNo = (sh, col, n) => {
    const r = sh.getRange(2, col, n);
    r.setDataValidation(list(['نعم', 'لا'])).setHorizontalAlignment('center');
    return cf(r, [['نعم', '#dcefe2', '#1f6b3a'], ['لا', '#fbe2e4', '#8b1c24']]);
  };

  let t = style(SH.MENU, [100, 170, 80, 80, 260, 110, 160, 130, 260], [
    'حروف إنجليزية بدون مسافات وما تتكرر (bday = كوب الميلاد بالزينة والصورة)', 'اسم المشروب كما يظهر', 'بالريال',
    'نعم = ظاهر، لا = مخفي (خلص)', 'وصف قصير', 'كلمة صغيرة فوق المشروب (اختياري)', 'اسم صورة من مجلد images، أو رابط قوقل درايف',
    'اختيار بسعر = العميل يختار الحليب بسعره من تبويب الحليب\nاختيار مجاني = يختار بدون زيادة\nأوريجينال فقط = بدون اختيار\nبدون = ما فيه حليب',
    'معرفات الإضافات المسموحة لهذا المشروب من تبويب الإضافات، بينها فاصلة. فاضي = بدون إضافات']);
  if (t) {
    const n = t.rows - 1;
    t.sh.getRange(2, 3, n).setNumberFormat('0 "ر.س"').setHorizontalAlignment('center');
    t.sh.getRange(2, 8, n).setDataValidation(list(MILK_MODES));
    t.sh.getRange(2, 1, n).setFontColor('#888888');
    t.sh.setConditionalFormatRules(yesNo(t.sh, 4, n));
  }
  [SH.ADD, SH.MILK].forEach(name => {
    const x = style(name, [100, 170, 80, 80]);
    if (x) { const n = x.rows - 1; x.sh.getRange(2, 3, n).setNumberFormat('0 "ر.س"').setHorizontalAlignment('center'); x.sh.getRange(2, 1, n).setFontColor('#888888'); x.sh.setConditionalFormatRules(yesNo(x.sh, 4, n)); }
  });
  t = style(SH.SET, [200, 300, 380]);
  if (t) {
    t.sh.getRange(2, 1, t.rows - 1).setFontWeight('bold');
    t.sh.getRange(2, 3, t.rows - 1).setFontColor('#888888').setFontStyle('italic').setFontSize(10);
    const keys = t.sh.getRange(2, 1, Math.max(1, t.sh.getLastRow() - 1)).getValues().map(r => String(r[0]).trim());
    const at = k => keys.indexOf(k) + 2;
    if (at('استقبال الطلبات') > 1) t.sh.getRange(at('استقبال الطلبات'), 2).setDataValidation(list(['نعم', 'لا']));
    if (at('لما يخلص صنف') > 1) t.sh.getRange(at('لما يخلص صنف'), 2).setDataValidation(list(['يختفي', 'يظهر باهت']));
    if (at('يوم الإجازة') > 1) t.sh.getRange(at('يوم الإجازة'), 2).setDataValidation(list(['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']));
  }
  t = style(SH.ORD, [95, 140, 110, 120, 120, 330, 90, 110, 100, 70, 200, 200, 110]);
  if (t) {
    const n = t.rows - 1;
    t.sh.getRange(2, 1, n).setFontWeight('bold');
    t.sh.getRange(2, 7, n).setNumberFormat('0 "ر.س"');
    const st = t.sh.getRange(2, 9, n);
    st.setDataValidation(list(['جديد', 'يتجهز', 'جاهز', 'استلم', 'ملغي'])).setHorizontalAlignment('center');
    t.sh.setConditionalFormatRules(cf(st, [['جديد', '#fbe2e4', '#8b1c24'], ['يتجهز', '#fff3c4', '#7a5a00'], ['جاهز', '#dcefe2', '#1f6b3a'], ['استلم', '#eeeeee', '#555555'], ['ملغي', '#eeeeee', '#999999']]));
  }
  t = style(SH.STAR, [140, 150, 90, 130]);
  if (t) t.sh.getRange(2, 3, t.rows - 1).setHorizontalAlignment('center').setFontWeight('bold');
  ss.getSheets().forEach(sh => { if (sh.getLastRow() > 1) sh.autoResizeRows(2, sh.getLastRow() - 1); });
}

/* =============== أدوات =============== */
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function rows_(n) {
  const sh = sheet_(n); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function settings_() { const o = {}; rows_(SH.SET).forEach(r => { if (r[0] !== '') o[String(r[0]).trim()] = r[1]; }); return o; }
function yes_(v) { return String(v).trim() === 'نعم'; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function phone_(p) {
  p = String(p || '').replace(/[^\d]/g, '');
  if (p.indexOf('966') === 0) p = '0' + p.slice(3);
  if (p.length === 9 && p[0] === '5') p = '0' + p;
  return p;
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd'); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[<>]/g, '').slice(0, n || 200); }
function ids_(v) { return String(v || '').split(/[,،\s]+/).map(x => x.trim()).filter(String); }
function milkMode_(v) {
  const t = String(v || '').trim();
  if (t.indexOf('مجاني') >= 0) return 'free';
  if (t.indexOf('أوريجينال') >= 0 || t.indexOf('اوريجينال') >= 0) return 'original';
  if (t.indexOf('اختيار') >= 0) return 'paid';
  return 'none';
}

function publicData_() {
  const s = settings_();
  const menu = rows_(SH.MENU).filter(r => r[0] !== '' && r[1] !== '').map(r => ({
    id: String(r[0]).trim(), name: String(r[1]), price: Number(r[2]) || 0, on: yes_(r[3]), desc: String(r[4] || ''),
    badge: String(r[5] || ''), img: String(r[6] || '').trim(), milk: milkMode_(r[7]), adds: ids_(r[8])
  }));
  const opt = n => rows_(n).filter(r => r[0] !== '' && r[1] !== '').map(r => ({ id: String(r[0]).trim(), name: String(r[1]), price: Number(r[2]) || 0, on: yes_(r[3]) }));
  return {
    menu: menu, adds: opt(SH.ADD), milks: opt(SH.MILK),
    settings: {
      name: s['اسم المحل'] || 'Day & Night',
      open: Number(s['يفتح الساعة']) || 15, close: Number(s['يسكر الساعة']) || 22,
      off: String(s['يوم الإجازة'] || ''), accepting: s['استقبال الطلبات'] === '' ? true : yes_(s['استقبال الطلبات']),
      banner: String(s['رسالة البانر'] || ''),
      hideOut: String(s['لما يخلص صنف'] || 'يختفي').trim() !== 'يظهر باهت',
      freeAt: Number(s['نجوم المشروب المجاني']) || 3,
      place: String(s['مكان الاستلام'] || ''),
      whatsapp: String(s['رقم الواتساب'] || '').replace(/[^\d]/g, ''),
      transfer: String(s['بيانات التحويل'] || '')
    }
  };
}

/* سعر المشروب يتحسب هنا في الخادم من الشيت، مو من المتصفح */
function priceItem_(m, it, data) {
  let unit = m.price; const parts = [];
  if (m.milk === 'paid' || m.milk === 'free') {
    const mk = data.milks.find(x => x.id === String(it.milk) && x.on) || data.milks.find(x => x.on);
    if (mk) { parts.push(mk.name); if (m.milk === 'paid') unit += mk.price; }
  } else if (m.milk === 'original') parts.push('Original Milk');
  (Array.isArray(it.adds) ? it.adds : []).slice(0, 10).forEach(id => {
    if (m.adds.indexOf(String(id)) < 0) return;
    const a = data.adds.find(x => x.id === String(id) && x.on); if (!a) return;
    unit += a.price; parts.push(a.name);
  });
  if (m.id === 'bday') {
    const tops = (Array.isArray(it.tops) ? it.tops : []).map(x => clean_(x, 20)).slice(0, 8);
    if (tops.length) parts.push('الزينة: ' + tops.join('، '));
    if (it.bname) parts.push('الاسم: ' + clean_(it.bname, 20));
  }
  return { unit: unit, text: parts.join(' · ') };
}

/* =============== GET =============== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.a === 'data') return json_(publicData_());
    if (p.a === 'track') {
      const r = rows_(SH.ORD).find(x => String(x[0]) === String(p.no));
      return json_(r ? { ok: true, status: String(r[8]) } : { ok: false });
    }
    if (p.a === 'stars') {
      const ph = phone_(p.phone);
      const r = rows_(SH.STAR).find(x => phone_(x[0]) === ph);
      return json_({ ok: true, stars: r ? Number(r[2]) || 0 : 0, freeAt: Number(settings_()['نجوم المشروب المجاني']) || 3 });
    }
    return json_({ ok: true, name: 'Day & Night API' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

/* =============== POST =============== */
function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (x) { return json_({ ok: false, error: 'bad json' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (b.a === 'order') return json_(placeOrder_(b));
    if (String(b.pass || '') !== String(settings_()['كلمة مرور اللوحة'])) return json_({ ok: false, error: 'auth' });
    if (b.a === 'admin') return json_(adminData_());
    if (b.a === 'status') return json_(setStatus_(b.no, b.status));
    if (b.a === 'stock') return json_(setOn_(b.kind === 'add' ? SH.ADD : b.kind === 'milk' ? SH.MILK : SH.MENU, b.id, b.on));
    if (b.a === 'accepting') return json_(setSetting_('استقبال الطلبات', b.on ? 'نعم' : 'لا'));
    if (b.a === 'redeem') return json_(redeem_(b.phone));
    return json_({ ok: false, error: 'unknown' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}

function placeOrder_(b) {
  const data = publicData_(), s = data.settings;
  if (!s.accepting) return { ok: false, msg: 'متوقفين عن استقبال الطلبات حاليًا' };
  const ph = phone_(b.phone);
  if (!/^05\d{8}$/.test(ph)) return { ok: false, msg: 'اكتبي رقم جوال صحيح' };
  let total = 0, cups = 0, cheapest = Infinity; const lines = []; let hasBday = false;
  (b.items || []).slice(0, 30).forEach(it => {
    const m = data.menu.find(x => x.id === String(it.id));
    const q = Math.max(1, Math.min(20, Number(it.q) || 1));
    if (!m || !m.on) return;
    const pr = priceItem_(m, it, data);
    total += pr.unit * q; cups += q; cheapest = Math.min(cheapest, pr.unit);
    if (m.id === 'bday') hasBday = true;
    lines.push(q + '× ' + m.name + (pr.text ? ' (' + pr.text + ')' : '') + ' = ' + (pr.unit * q));
  });
  if (!lines.length) return { ok: false, msg: 'المشروبات المطلوبة خلصت' };

  // مكافأة النجوم: أرخص مشروب علينا
  let reward = 0;
  const stSh = sheet_(SH.STAR), stAll = rows_(SH.STAR);
  let si = stAll.findIndex(x => phone_(x[0]) === ph);
  let stars = si >= 0 ? Number(stAll[si][2]) || 0 : 0;
  if (b.reward && stars >= s.freeAt) { reward = cheapest; total -= reward; stars -= s.freeAt; }
  stars += cups;
  if (si >= 0) stSh.getRange(si + 2, 2, 1, 3).setValues([[clean_(b.name, 40) || stAll[si][1], stars, today_()]]);
  else stSh.appendRow(["'" + ph, clean_(b.name, 40), stars, today_()]);

  let photo = '';
  if (hasBday && b.photo) photo = savePhoto_(b.photo, ph);

  const sh = sheet_(SH.ORD);
  const no = 'DN-' + (100 + sh.getLastRow());
  sh.appendRow([no, now_(), clean_(b.time, 30), clean_(b.name, 40) || 'ضيف', "'" + ph, lines.join('\n'), total,
    b.pay === 'تحويل' ? 'تحويل بنكي' : 'كاش', 'جديد', cups, clean_(b.note, 300), photo, reward ? ('-' + reward + ' ر.س') : '']);
  return { ok: true, no: no, total: total, reward: reward, stars: stars, freeAt: s.freeAt,
    place: s.place, map: String(settings_()['رابط اللوكيشن'] || ''), transfer: s.transfer };
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Day & Night — صور أكواب الميلاد');
  props.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
function savePhoto_(dataUrl, ph) {
  try {
    const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
    if (!m || m[2].length > 2800000) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'bday-' + ph + '-' + Date.now() + '.jpg');
    const file = photoFolder_().createFile(blob);
    return file.getUrl();
  } catch (e) { return ''; }
}

function adminData_() {
  const t = today_();
  const ord = rows_(SH.ORD).map(r => ({ no: String(r[0]), at: String(r[1]), time: String(r[2]), name: String(r[3]),
    phone: String(r[4]), items: String(r[5]), total: Number(r[6]) || 0, pay: String(r[7]), status: String(r[8]), cups: Number(r[9]) || 0,
    note: String(r[10] || ''), photo: String(r[11] || ''), reward: String(r[12] || '') }));
  const todayOrd = ord.filter(o => o.at.indexOf(t) === 0 && o.status !== 'ملغي');
  const active = ord.filter(o => o.status !== 'استلم' && o.status !== 'ملغي').slice(-60);
  const stars = {}; rows_(SH.STAR).forEach(r => stars[phone_(r[0])] = Number(r[2]) || 0);
  const pub = publicData_();
  return { ok: true, today: { orders: todayOrd.length, cups: todayOrd.reduce((a, o) => a + o.cups, 0), total: todayOrd.reduce((a, o) => a + o.total, 0) },
    orders: active.reverse(), menu: pub.menu, adds: pub.adds, milks: pub.milks, settings: pub.settings, stars: stars };
}
function setStatus_(no, status) {
  if (['جديد', 'يتجهز', 'جاهز', 'استلم', 'ملغي'].indexOf(status) < 0) return { ok: false };
  const all = rows_(SH.ORD); const i = all.findIndex(r => String(r[0]) === String(no));
  if (i < 0) return { ok: false };
  sheet_(SH.ORD).getRange(i + 2, 9).setValue(status);
  return { ok: true };
}
function setOn_(name, id, on) {
  const col = name === SH.MENU ? 4 : 4;
  const all = rows_(name); const i = all.findIndex(r => String(r[0]).trim() === String(id));
  if (i < 0) return { ok: false };
  sheet_(name).getRange(i + 2, col).setValue(on ? 'نعم' : 'لا');
  return { ok: true };
}
function setSetting_(key, val) {
  const all = rows_(SH.SET); const i = all.findIndex(r => String(r[0]).trim() === key);
  if (i < 0) return { ok: false };
  sheet_(SH.SET).getRange(i + 2, 2).setValue(val);
  return { ok: true };
}
function redeem_(phone) {
  const ph = phone_(phone), need = Number(settings_()['نجوم المشروب المجاني']) || 3;
  const all = rows_(SH.STAR); const i = all.findIndex(r => phone_(r[0]) === ph);
  if (i < 0 || (Number(all[i][2]) || 0) < need) return { ok: false, msg: 'ما عندها نجوم كافية' };
  const left = Number(all[i][2]) - need;
  sheet_(SH.STAR).getRange(i + 2, 3, 1, 2).setValues([[left, today_()]]);
  return { ok: true, stars: left };
}
