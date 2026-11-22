/**
 * CRM CMA | SKT — máy chủ dữ liệu chạy trên Google Apps Script, lưu vào Google Sheets này.
 * Website (GitHub Pages, crmtkt.com) gọi tới địa chỉ Web App của script.
 * Đổi mật khẩu: tạo ADMIN_SALT/ADMIN_HASH mới (xem README trên GitHub) rồi Triển khai lại.
 */
const ADMIN_USER = 'crmtkt';
const ADMIN_SALT = '__SALT__';
const ADMIN_HASH = '__HASH__';
const STAFF_DEFAULT = ['Kim Anh', 'Mỹ Linh', 'Oanh'];

const COLS = {
  forms: ['id', 'slug', 'ten', 'khoa', 'nam', 'mau', 'moTa', 'mo', 'createdAt'],
  records: ['id', 'formId', 'hoTen', 'sdt', 'email', 'trangThai', 'phuTrach', 'createdAt', 'updatedAt', 'data', 'anh', 'anh2', 'anh3', 'anh4'],
  settings: ['k', 'v'],
};
const ANSWER_KEYS = ['hoTen','namSinh','sdt','email','thanhPho','congViec','vaiTro','khoaDaHoc','lyDo','vanDe','mongMuon','ketQua','ngaiNgai','sanPhamRieng','thoiGian','aiLevel','aiTools','aiDaLam','aiMuon','dongY'];
const REQUIRED = {
  hocvien: ['hoTen','namSinh','sdt','email','congViec','khoaDaHoc','lyDo','vanDe','mongMuon','aiLevel','dongY'],
  nhanh: ['hoTen','sdt','dongY'],
};
const LABELS = {hoTen:'họ tên',namSinh:'năm sinh',sdt:'số điện thoại',email:'email',congViec:'công việc',khoaDaHoc:'các khóa đã và đang học',lyDo:'lý do đăng ký',vanDe:'vấn đề cần hỗ trợ',mongMuon:'mong muốn',aiLevel:'mức độ AI',dongY:'ô đồng ý'};
const STATUSES = ['Mới đăng ký','Đã liên hệ','Đã vào nhóm Zalo','Đang học','Cần chăm sóc gấp','Tạm dừng','Hoàn thành'];
const PAY_ST = ['Đã đóng đủ', 'Đặt cọc', 'Chưa đóng', 'Bảo lưu / chưa chốt'];
// Ảnh chân dung rõ nét: tối đa ~190.000 ký tự, chia vào 4 ô (mỗi ô Google Sheet giới hạn 50.000 ký tự).
const PHOTO_MAX = 190000, PHOTO_CH = 48000, PHOTO_COLS = ['anh', 'anh2', 'anh3', 'anh4'];
function photoOk_(a) { return typeof a === 'string' && a.length <= PHOTO_MAX && /^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/.test(a); }
function setPhoto_(r, a) { PHOTO_COLS.forEach((k, i) => r[k] = a.slice(i * PHOTO_CH, (i + 1) * PHOTO_CH)); }
function getPhoto_(r) { return PHOTO_COLS.map(k => r[k] || '').join(''); }
function money_(v) { const n = Math.round(Number(v) || 0); return n > 0 && n < 1e11 ? n : 0; }
const COURSE = {Diamond: 'CMA Diamond', Trainer: 'CMA Trainer'};

/* ---------- tiện ích bảng tính ---------- */
function book_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  let s = book_().getSheetByName(name);
  if (!s) {
    s = book_().insertSheet(name);
    s.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    s.getRange('A:Z').setNumberFormat('@'); // giữ số 0 đầu số điện thoại
    s.setFrozenRows(1);
  }
  return s;
}
function rows_(name) {
  const s = sheet_(name), n = s.getLastRow(), cols = COLS[name];
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, cols.length).getValues().map((r, i) => {
    const o = {_row: i + 2}; cols.forEach((c, j) => o[c] = String(r[j])); return o;
  });
}
function put_(name, o) {
  const cols = COLS[name], vals = cols.map(c => o[c] === undefined || o[c] === null ? '' : String(o[c]));
  const s = sheet_(name);
  if (s.getLastColumn() < cols.length) s.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  if (o._row) s.getRange(o._row, 1, 1, cols.length).setValues([vals]);
  else { s.appendRow(vals); }
}
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function now_() { return new Date().toISOString(); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'); }
function phone_(s) { return String(s || '').replace(/\D+/g, '').slice(0, 20); }
function str_(v, max) { return (v === null || v === undefined || typeof v === 'object') ? '' : String(v).trim().slice(0, max || 2000); }
function list_(v, n, max) { if (!Array.isArray(v)) return []; const o = []; v.slice(0, n || 30).forEach(x => { const s = str_(x, max || 200); if (s && o.indexOf(s) < 0) o.push(s); }); return o; }
function hex_(bytes) { return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join(''); }
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function fail_(msg, code) { return out_({ok: false, error: msg, code: code || 400}); }

function seed_() {
  if (rows_('forms').length) return;
  const t = now_();
  put_('forms', {id: uid_(), slug: 'cma-diamond-2026', ten: 'Phiếu thông tin học viên CMA Diamond', khoa: 'Diamond', nam: '2026', mau: 'hocvien', moTa: 'Phiếu chính cho học viên CMA Diamond 2026.', mo: '1', createdAt: t});
  put_('forms', {id: uid_(), slug: 'cma-trainer-2026', ten: 'Phiếu thông tin học viên CMA Trainer', khoa: 'Trainer', nam: '2026', mau: 'hocvien', moTa: 'Phiếu cho học viên CMA Trainer 2026.', mo: '1', createdAt: t});
}

/* ---------- đăng nhập bằng mã ký (không dùng cookie) ---------- */
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function sign_(payload) { return hex_(Utilities.computeHmacSha256Signature(payload, secret_())); }
function makeToken_() { const p = ADMIN_USER + '|' + (Date.now() + 12 * 3600 * 1000); return Utilities.base64EncodeWebSafe(p) + '.' + sign_(p); }
function checkToken_(t) {
  if (typeof t !== 'string' || t.indexOf('.') < 0) return false;
  const parts = t.split('.'); let p;
  try { p = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { return false; }
  if (sign_(p) !== parts[1]) return false;
  return Number(p.split('|')[1]) > Date.now();
}
function passOk_(pw) {
  let h = sha_(ADMIN_SALT + ':' + pw);
  for (let i = 0; i < 1000; i++) h = sha_(h + ADMIN_SALT);
  return h === ADMIN_HASH;
}

/* ---------- định dạng trả về ---------- */
function formOut_(f, count) { return {id: f.id, slug: f.slug, ten: f.ten, khoa: f.khoa, nam: f.nam, mau: f.mau, moTa: f.moTa, mo: f.mo === '1', createdAt: f.createdAt.slice(0, 10), soPhieu: count || 0}; }
function recOut_(r) {
  let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {}
  d.id = r.id; d.formId = r.formId; d.createdAt = r.createdAt.slice(0, 10); d.updatedAt = r.updatedAt; d.hasPhoto = !!r.anh;
  return d;
}
function staff_() {
  const r = rows_('settings').filter(x => x.k === 'staff')[0];
  try { const l = JSON.parse(r.v); if (l.length) return l; } catch (e) {}
  return STAFF_DEFAULT;
}

/* ---------- làm sạch dữ liệu ---------- */
function answers_(a) {
  const o = {};
  ANSWER_KEYS.forEach(k => {
    if (!(k in a)) return;
    if (k === 'aiTools') o[k] = list_(a[k]);
    else if (k === 'aiLevel') o[k] = Math.max(0, Math.min(5, parseInt(a[k], 10) || 0));
    else if (k === 'dongY') o[k] = a[k] === true;
    else o[k] = str_(a[k], ['hoTen','sdt','email','namSinh','thanhPho','congViec','vaiTro','thoiGian'].indexOf(k) >= 0 ? 200 : 2000);
  });
  return o;
}
function entries_(v, keys) {
  if (!Array.isArray(v)) return [];
  return v.slice(-300).map(e => { const r = {}; keys.forEach(k => r[k] = str_(e && e[k], 1000)); return r; }).filter(r => r.noiDung);
}
function crm_(p) {
  const o = {};
  if ('trangThai' in p) o.trangThai = STATUSES.indexOf(p.trangThai) >= 0 ? p.trangThai : 'Mới đăng ký';
  if ('phuTrach' in p) o.phuTrach = list_(p.phuTrach, 20, 80);
  ['khoa', 'nam', 'nhom', 'ghiChu'].forEach(k => { if (k in p) o[k] = str_(p[k], 2000); });
  if ('zaloGroup' in p) { const z = str_(p.zaloGroup, 500); o.zaloGroup = /^https:\/\/[^\s"'<>]+$/.test(z) ? z : ''; }
  if ('hanhTrinh' in p) o.hanhTrinh = entries_(p.hanhTrinh, ['ngay', 'noiDung']);
  if ('nhatKy' in p) o.nhatKy = entries_(p.nhatKy, ['ngay', 'nguoi', 'noiDung']);
  ['chuongTrinh', 'dongHanh', 'lichHen', 'canhBao', 'cccd', 'diaChi', 'ghiChuTT'].forEach(k => { if (k in p) o[k] = str_(p[k], 2000); });
  if ('phieuOffline' in p) o.phieuOffline = str_(p.phieuOffline, 12000);
  if ('ttTT' in p) o.ttTT = PAY_ST.indexOf(p.ttTT) >= 0 ? p.ttTT : '';
  if ('hanTT' in p) o.hanTT = /^\d{4}-\d{2}-\d{2}$/.test(String(p.hanTT)) ? String(p.hanTT) : '';
  ['hocPhi', 'khauTru', 'daDong'].forEach(k => { if (k in p) o[k] = money_(p[k]); });
  if ('lichSuTT' in p) o.lichSuTT = entries_(p.lichSuTT, ['ngay', 'soTien', 'noiDung']);
  if ('tienDo' in p && p.tienDo && typeof p.tienDo === 'object') { o.tienDo = {}; Object.keys(p.tienDo).forEach(m => { if (/^\d{4}-\d{2}$/.test(m)) { const s = str_(p.tienDo[m], 1000); if (s) o.tienDo[m] = s; } }); }
  return o;
}
function slug_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 70);
}

/* ---------- điểm vào ---------- */
function doGet(e) {
  const a = (e.parameter || {}).a;
  if (a === 'form') {
    seed_();
    const f = rows_('forms').filter(x => x.slug === String(e.parameter.s || ''))[0];
    if (!f) return fail_('Không tìm thấy form.', 404);
    return out_({ok: true, form: {s: f.slug, ten: f.ten, khoa: f.khoa, khoaTen: COURSE[f.khoa] || f.khoa, mau: f.mau, mo: f.mo === '1'}});
  }
  return out_({ok: true, app: 'CRM CMA | SKT'});
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return fail_('Dữ liệu không hợp lệ.'); }
  try { return route_(b); } catch (err) { console.error(err); return fail_('Lỗi máy chủ, vui lòng thử lại.', 500); }
}

function route_(b) {
  const a = b.a;
  if (a === 'submit') return submit_(b);
  if (a === 'login') {
    const c = CacheService.getScriptCache(), n = Number(c.get('fail') || 0);
    if (n >= 10) return fail_('Sai quá nhiều lần. Vui lòng thử lại sau 15 phút.', 429);
    if (str_(b.u, 100) !== ADMIN_USER || !passOk_(String(b.p || ''))) { c.put('fail', String(n + 1), 900); return fail_('Sai tài khoản hoặc mật khẩu.', 401); }
    return out_({ok: true, tok: makeToken_(), user: ADMIN_USER});
  }
  if (!checkToken_(b.tok)) return fail_('Chưa đăng nhập.', 401);
  seed_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    switch (a) {
      case 'me': return out_({ok: true, user: ADMIN_USER});
      case 'bootstrap': {
        const recs = rows_('records'), counts = {};
        recs.forEach(r => counts[r.formId] = (counts[r.formId] || 0) + 1);
        return out_({ok: true, forms: rows_('forms').reverse().map(f => formOut_(f, counts[f.id])), records: recs.reverse().map(recOut_), staff: staff_()});
      }
      case 'photos': {
        const ids = list_(b.ids, 40, 40), map = {};
        rows_('records').forEach(r => { if (ids.indexOf(r.id) >= 0 && r.anh) map[r.id] = getPhoto_(r); });
        return out_({ok: true, photos: map});
      }
      case 'form_save': {
        const ten = str_(b.ten, 200); if (!ten) return fail_('Vui lòng đặt tên form.');
        const slug = slug_(str_(b.slug, 80) || ten); if (!slug) return fail_('Đường dẫn chưa hợp lệ.');
        const all = rows_('forms'), id = str_(b.id, 32);
        if (all.some(f => f.slug === slug && f.id !== id)) return fail_('Đường dẫn này đã có form khác dùng, vui lòng đổi.');
        const cur = all.filter(f => f.id === id)[0] || {id: uid_(), createdAt: now_()};
        Object.assign(cur, {slug: slug, ten: ten, khoa: str_(b.khoa, 80) || 'Diamond', nam: /^\d{4}$/.test(b.nam) ? b.nam : '2026', mau: b.mau === 'nhanh' ? 'nhanh' : 'hocvien', moTa: str_(b.moTa, 1000), mo: b.mo ? '1' : '0'});
        put_('forms', cur); return out_({ok: true});
      }
      case 'form_delete': {
        const id = str_(b.id, 32);
        if (rows_('records').some(r => r.formId === id)) return fail_('Form đã có phiếu gửi về nên không xóa được. Anh/chị có thể bấm "Đóng form".');
        const f = rows_('forms').filter(x => x.id === id)[0]; if (f) sheet_('forms').deleteRow(f._row);
        return out_({ok: true});
      }
      case 'record_save': {
        const r = rows_('records').filter(x => x.id === str_(b.id, 32))[0]; if (!r) return fail_('Không tìm thấy hồ sơ.', 404);
        let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (err) {}
        const P = b.patch || {}, ap = {};
        ANSWER_KEYS.forEach(k => { if (k in P && k !== 'dongY') ap[k] = P[k]; });
        const A = answers_(ap);
        if ('hoTen' in A && !A.hoTen) return fail_('Họ tên không được trống.');
        if ('sdt' in A && !/^\+?\d{9,12}$/.test(String(A.sdt).replace(/[\s.\-]/g, ''))) return fail_('Số điện thoại chưa đúng.');
        if (A.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(A.email)) return fail_('Email chưa đúng.');
        Object.assign(d, A, crm_(P));
        r.data = JSON.stringify(d); r.hoTen = d.hoTen || ''; r.sdt = phone_(d.sdt); r.email = d.email || ''; r.trangThai = d.trangThai || ''; r.phuTrach = (d.phuTrach || []).join(', '); r.updatedAt = now_();
        put_('records', r); return out_({ok: true, record: recOut_(r)});
      }
      case 'record_create': {
        // Admin thêm hoặc cập nhật học viên (nhập từ báo cáo). Trùng SĐT/tên + khóa + nhóm + chương trình thì cập nhật.
        const P = b.patch || {}, A = answers_(P);
        if (!A.hoTen) return fail_('Họ tên không được trống.');
        const sd = String(A.sdt || '').replace(/[\s.\-]/g, '');
        if (sd && !/^\+?\d{9,12}$/.test(sd)) return fail_('Số điện thoại chưa đúng.');
        if (A.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(A.email)) return fail_('Email chưa đúng.');
        const anh = typeof b.anh === 'string' ? b.anh : '';
        if (anh && !photoOk_(anh)) return fail_('Ảnh chưa hợp lệ.');
        const C = crm_(P), key = d => [phone_(d.sdt) || String(d.hoTen || '').toLowerCase(), d.khoa || '', d.nhom || '', d.chuongTrinh || ''].join('|');
        const nd = Object.assign({khoa: 'Diamond', nam: '2026', nhom: '', trangThai: 'Mới đăng ký', phuTrach: [], hanhTrinh: [], nhatKy: [], tienDo: {}}, A, C);
        nd.sdt = sd;
        const fid = str_(b.formId, 32);
        let r = rows_('records').filter(x => { let d = {}; try { d = JSON.parse(x.data || '{}'); } catch (e) {} return key(d) === key(nd); })[0];
        if (r) { let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {} nd.hanhTrinh = C.hanhTrinh || d.hanhTrinh || []; r.data = JSON.stringify(Object.assign(d, nd)); if (anh) setPhoto_(r, anh); }
        else { r = {id: uid_(), formId: fid, createdAt: now_(), data: JSON.stringify(nd)}; setPhoto_(r, anh); }
        const d = JSON.parse(r.data);
        Object.assign(r, {hoTen: d.hoTen, sdt: phone_(d.sdt), email: d.email || '', trangThai: d.trangThai || '', phuTrach: (d.phuTrach || []).join(', '), updatedAt: now_()});
        put_('records', r); return out_({ok: true, record: recOut_(r)});
      }
      case 'photo_save': {
        const r = rows_('records').filter(x => x.id === str_(b.id, 32))[0]; if (!r) return fail_('Không tìm thấy hồ sơ.', 404);
        if (!photoOk_(b.anh)) return fail_('Ảnh chưa hợp lệ hoặc quá lớn.');
        setPhoto_(r, b.anh); r.updatedAt = now_(); put_('records', r);
        return out_({ok: true});
      }
      case 'record_delete': {
        const r = rows_('records').filter(x => x.id === str_(b.id, 32))[0]; if (r) sheet_('records').deleteRow(r._row);
        return out_({ok: true});
      }
      case 'staff_save': {
        const l = list_(b.list, 50, 80); if (!l.length) return fail_('Danh sách nhân sự không được trống.');
        const r = rows_('settings').filter(x => x.k === 'staff')[0] || {k: 'staff'}; r.v = JSON.stringify(l); put_('settings', r);
        return out_({ok: true});
      }
    }
    return fail_('Không rõ yêu cầu.', 404);
  } finally { lock.releaseLock(); }
}

function submit_(b) {
  if (str_(b.web, 100)) return out_({ok: true}); // bẫy bot
  seed_();
  const f = rows_('forms').filter(x => x.slug === str_(b.s, 80))[0];
  if (!f) return fail_('Không tìm thấy form.', 404);
  if (f.mo !== '1') return fail_('Form này đã đóng, vui lòng liên hệ Ban tổ chức.', 403);
  const ans = answers_(b.data && typeof b.data === 'object' ? b.data : {});
  const anh = typeof b.anh === 'string' ? b.anh : '';
  for (const k of (REQUIRED[f.mau] || REQUIRED.nhanh)) {
    const v = ans[k]; if (v === undefined || v === '' || v === false || v === 0) return fail_('Vui lòng điền ' + (LABELS[k] || k) + '.');
  }
  if (f.mau === 'hocvien' && !anh) return fail_('Vui lòng chọn ảnh chân dung.');
  if (anh && !photoOk_(anh)) return fail_('Ảnh chưa hợp lệ, vui lòng chọn ảnh khác.');
  if (!/^\+?\d{9,12}$/.test(String(ans.sdt || '').replace(/[\s.\-]/g, ''))) return fail_('Số điện thoại chưa đúng.');
  if (ans.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(ans.email)) return fail_('Email chưa đúng.');
  ans.guiLuc = now_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sdt = phone_(ans.sdt);
    const old = rows_('records').filter(r => r.formId === f.id && phone_(r.sdt) === sdt)[0];
    let d, r;
    if (old) {
      try { d = JSON.parse(old.data || '{}'); } catch (e) { d = {}; }
      Object.assign(d, ans);
      d.hanhTrinh = (d.hanhTrinh || []).concat([{ngay: today_(), noiDung: 'Cập nhật lại phiếu: ' + f.ten}]);
      r = old; if (anh) setPhoto_(r, anh);
    } else {
      d = Object.assign(ans, {khoa: f.khoa, nam: f.nam, nhom: '', trangThai: 'Mới đăng ký', phuTrach: [], hanhTrinh: [{ngay: today_(), noiDung: 'Điền ' + f.ten}], nhatKy: [], tienDo: {}});
      r = {id: uid_(), formId: f.id, createdAt: now_()}; setPhoto_(r, anh);
    }
    Object.assign(r, {hoTen: d.hoTen, sdt: sdt, email: d.email || '', trangThai: d.trangThai, phuTrach: (d.phuTrach || []).join(', '), updatedAt: now_(), data: JSON.stringify(d)});
    put_('records', r);
  } finally { lock.releaseLock(); }
  return out_({ok: true});
}

/** Chạy hàm này một lần trong trình soạn thảo để tạo các trang tính và cấp quyền. */
function caiDat() { seed_(); sheet_('records'); sheet_('settings'); secret_(); }
