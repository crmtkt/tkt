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
  records: ['id', 'formId', 'hoTen', 'sdt', 'email', 'trangThai', 'phuTrach', 'createdAt', 'updatedAt', 'data', 'anh', 'anh2', 'anh3', 'anh4', 'coAnh'],
  settings: ['k', 'v'],
  camps: ['id', 'ma', 'ten', 'data', 'duyet', 'updatedAt'],
  users: ['id', 'hoTen', 'sdt', 'email', 'phong', 'salt', 'hash', 'active', 'createdAt', 'updatedAt', 'lastLogin'],
  txns: ['id', 'key', 'ngay', 'soTien', 'noiDung', 'nguoiCK', 'taiKhoan', 'trangThai', 'alloc', 'ghiChu', 'createdAt', 'xacNhanLuc', 'lo', 'hd'],
  uploads: ['id', 'ten', 'taiKhoan', 'soDong', 'them', 'trung', 'tongTien', 'tuNgay', 'denNgay', 'nguoi', 'luc'],
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
function setPhoto_(r, a) { PHOTO_COLS.forEach((k, i) => r[k] = a.slice(i * PHOTO_CH, (i + 1) * PHOTO_CH)); r.coAnh = a ? 'v' + Date.now() : '-'; }
const LITE = 10; // số cột đầu (id → data) đọc nhanh, không kéo theo ảnh
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
function rows_(name, full) {
  const s = sheet_(name), n = s.getLastRow(), cols = COLS[name];
  if (n < 2) return [];
  if (name === 'records' && !full) {
    // Đọc nhanh: bỏ qua 4 cột ảnh rất nặng; cột coAnh cho biết có ảnh hay không (và phiên bản ảnh).
    const ci = cols.indexOf('coAnh') + 1, base = s.getRange(2, 1, n - 1, LITE).getValues(), flag = s.getRange(2, ci, n - 1, 1).getValues();
    if (flag.some(f => String(f[0]) === '')) {
      const k = s.getRange(2, LITE + 1, n - 1, 1).getValues();
      flag.forEach((f, i) => { if (String(f[0]) === '') f[0] = String(k[i][0]) ? 'v1' : '-'; });
      s.getRange(2, ci, n - 1, 1).setValues(flag);
      s.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    }
    return base.map((r, i) => { const o = {_row: i + 2, _lite: true}; cols.slice(0, LITE).forEach((c, j) => o[c] = String(r[j])); o.coAnh = String(flag[i][0]); return o; });
  }
  return s.getRange(2, 1, n - 1, cols.length).getValues().map((r, i) => {
    const o = {_row: i + 2}; cols.forEach((c, j) => o[c] = String(r[j])); return o;
  });
}
function put_(name, o) {
  const cols = COLS[name], vals = cols.map(c => o[c] === undefined || o[c] === null ? '' : String(o[c]));
  const s = sheet_(name);
  if (s.getLastColumn() < cols.length) s.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  if (o._row && o._lite && o.anh === undefined) { s.getRange(o._row, 1, 1, LITE).setValues([vals.slice(0, LITE)]); return; } // không ghi đè ảnh
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
function makeToken_(uid) { const p = (uid || ADMIN_USER) + '|' + (Date.now() + 12 * 3600 * 1000); return Utilities.base64EncodeWebSafe(p) + '.' + sign_(p); }
function checkToken_(t) {
  if (typeof t !== 'string' || t.indexOf('.') < 0) return false;
  const parts = t.split('.'); let p;
  try { p = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { return false; }
  if (sign_(p) !== parts[1]) return false;
  return Number(p.split('|')[1]) > Date.now() ? p.split('|')[0] : false;
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
  d.id = r.id; d.formId = r.formId; d.createdAt = r.createdAt.slice(0, 10); d.updatedAt = r.updatedAt; d.hasPhoto = !!r.coAnh && r.coAnh !== '-'; d.phV = r.coAnh || '';
  return d;
}
function staff_() {
  let l = STAFF_DEFAULT; const r = rows_('settings').filter(x => x.k === 'staff')[0];
  try { const v = JSON.parse(r.v); if (v.length) l = v; } catch (e) {}
  const u = rows_('users').filter(x => x.active === '1').map(x => x.hoTen).filter(n => n && l.indexOf(n) < 0);
  return l.concat(u);
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
  if ('lichSuTT' in p) o.lichSuTT = entries_(p.lichSuTT, ['ngay', 'soTien', 'noiDung', 'gd']);
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
  if ((e.parameter || {}).mcp) return mcpOut_({ok: true, server: 'CRM CMA SKT – MCP (chỉ đọc)', huongDan: 'Dán nguyên link này vào ChatGPT hoặc Claude (Add custom MCP server / connector).'});
  if (e.pathInfo || a === 'gpt') { try { return gpt_(e); } catch (err) { console.error(err); return fail_('Lỗi máy chủ, vui lòng thử lại.', 500); } }
  if (a === 'form') {
    seed_();
    const f = rows_('forms').filter(x => x.slug === String(e.parameter.s || ''))[0];
    if (!f) return fail_('Không tìm thấy form.', 404);
    return out_({ok: true, form: {s: f.slug, ten: f.ten, khoa: f.khoa, khoaTen: COURSE[f.khoa] || f.khoa, mau: f.mau, mo: f.mo === '1'}});
  }
  return out_({ok: true, app: 'CRM CMA | SKT'});
}

function doPost(e) {
  if ((e.parameter || {}).zalo) { let zb = {}; try { zb = JSON.parse(e.postData.contents); } catch (err) { return fail_('Dữ liệu không hợp lệ.'); } try { return zlPush_(zb); } catch (err) { console.error(err); return fail_('Lỗi máy chủ, thử lại sau.', 500); } }
  if ((e.parameter || {}).mcp) { try { return mcp_(e); } catch (err) { console.error(err); return mcpOut_({jsonrpc: '2.0', id: null, error: {code: -32603, message: 'Lỗi máy chủ CRM'}}); } }
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return fail_('Dữ liệu không hợp lệ.'); }
  try { return route_(b); } catch (err) { console.error(err); return fail_('Lỗi máy chủ, vui lòng thử lại.', 500); }
}

function route_(b) {
  const a = b.a;
  if (a === 'submit') return submit_(b);
  if (a === 'lookup') return lookup_(b);
  if (a === 'login') return login_(b);
  const uid = checkToken_(b.tok); if (!uid) return fail_('Chưa đăng nhập.', 401);
  const ME = who_(uid); if (!ME) return fail_('Tài khoản đã bị khóa hoặc không còn tồn tại.', 401);
  if (PERM_[a] && PERM_[a].indexOf(ME.role) < 0) return fail_('Tài khoản của anh/chị (' + ROLES_[ME.role] + ') không có quyền làm việc này.', 403);
  if (ME.role === 'marketing' && b.patch && (a === 'record_save' || a === 'record_create')) MONEY_KEYS_.forEach(k => delete b.patch[k]);
  b.__me = ME;
  seed_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    switch (a) {
      case 'me': return out_({ok: true, user: ME.ten, me: ME});
      case 'bootstrap': {
        const recs = rows_('records'), counts = {};
        recs.forEach(r => counts[r.formId] = (counts[r.formId] || 0) + 1);
        return out_({ok: true, forms: rows_('forms').reverse().map(f => formOut_(f, counts[f.id])), records: recs.reverse().map(recOut_), zaloOk: zAny_(ME), staff: staff_(), camps: camps_(), txns: txns_(), uploads: ['quantri', 'ketoan'].indexOf(ME.role) >= 0 ? uploads_() : [], me: ME, users: ME.role === 'quantri' ? users_() : []});
      }
      case 'photos': {
        const ids = list_(b.ids, 40, 40), map = {};
        rows_('records', true).forEach(r => { if (ids.indexOf(r.id) >= 0 && r.anh) map[r.id] = getPhoto_(r); });
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
    if (KT_[a]) return KT_[a](b);
    if (US_[a]) return US_[a](b, ME);
    if (GP_[a]) return GP_[a](b, ME);
    if (ZL_[a]) return ZL_[a](b, ME);
    return fail_('Không rõ yêu cầu.', 404);
  } finally { lock.releaseLock(); }
}

/* ---------- tra cứu học viên cũ: form chỉ hỏi những câu còn thiếu ---------- */
const FORM_KEYS = {
  hocvien: ['anh', 'namSinh', 'email', 'thanhPho', 'congViec', 'vaiTro', 'khoaDaHoc', 'lyDo', 'vanDe', 'mongMuon', 'ketQua', 'ngaiNgai', 'sanPhamRieng', 'thoiGian', 'aiLevel', 'aiTools', 'aiDaLam', 'aiMuon'],
  nhanh: ['anh', 'email', 'congViec', 'mongMuon'],
};
function phoneN_(s) { let p = phone_(s); if (/^84\d{9,10}$/.test(p)) p = '0' + p.slice(2); if (/^[1-9]\d{8}$/.test(p)) p = '0' + p; return p; } // Sheet có thể làm mất số 0 đầu
function nameKey_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/[^a-z0-9 ]+/g, ' ').trim().split(/\s+/).filter(String); }
function sameName_(a, b) { const x = nameKey_(a), y = nameKey_(b); return x.length > 0 && y.length > 0 && x[x.length - 1] === y[y.length - 1]; }
function matchRecs_(sdt, hoTen) { return rows_('records').filter(r => phoneN_(r.sdt) === sdt && sameName_(r.hoTen, hoTen)); }
function filled_(v) { return Array.isArray(v) ? v.length > 0 : typeof v === 'number' ? v > 0 : typeof v === 'boolean' ? v : String(v === null || v === undefined ? '' : v).trim() !== ''; }
function have_(recs, keys) {
  const h = {};
  recs.forEach(r => { let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {} keys.forEach(k => { if (k === 'anh' ? (!!r.coAnh && r.coAnh !== '-') : filled_(d[k])) h[k] = 1; }); });
  return Object.keys(h);
}
function lookup_(b) {
  // Chỉ trả về TÊN các câu đã có (không trả nội dung) để form ẩn đi; phải khớp cả SĐT lẫn tên.
  const c = CacheService.getScriptCache(), n = Number(c.get('lk') || 0);
  if (n > 400) return fail_('Hệ thống đang bận, anh/chị thử lại sau ít phút.', 429);
  c.put('lk', String(n + 1), 600);
  seed_();
  const f = rows_('forms').filter(x => x.slug === str_(b.s, 80))[0];
  if (!f) return fail_('Không tìm thấy form.', 404);
  const sdt = phoneN_(b.sdt), ten = str_(b.hoTen, 200);
  if (sdt.length < 9 || !ten) return fail_('Vui lòng nhập họ tên và số điện thoại.');
  const recs = matchRecs_(sdt, ten);
  return out_({ok: true, found: recs.length > 0, have: recs.length ? have_(recs, FORM_KEYS[f.mau] || FORM_KEYS.nhanh) : []});
}

function submit_(b) {
  if (str_(b.web, 100)) return out_({ok: true}); // bẫy bot
  seed_();
  const f = rows_('forms').filter(x => x.slug === str_(b.s, 80))[0];
  if (!f) return fail_('Không tìm thấy form.', 404);
  if (f.mo !== '1') return fail_('Form này đã đóng, vui lòng liên hệ Ban tổ chức.', 403);
  const ans = answers_(b.data && typeof b.data === 'object' ? b.data : {});
  const anh = typeof b.anh === 'string' ? b.anh : '';
  if (!ans.hoTen) return fail_('Vui lòng điền họ tên.');
  if (!/^\+?\d{9,12}$/.test(String(ans.sdt || '').replace(/[\s.\-]/g, ''))) return fail_('Số điện thoại chưa đúng.');
  const sdt = phoneN_(ans.sdt);
  const known = matchRecs_(sdt, ans.hoTen), have = known.length ? have_(known, FORM_KEYS[f.mau] || FORM_KEYS.nhanh) : [];
  for (const k of (REQUIRED[f.mau] || REQUIRED.nhanh)) {
    if (have.indexOf(k) >= 0) continue;
    const v = ans[k]; if (v === undefined || v === '' || v === false || v === 0) return fail_('Vui lòng điền ' + (LABELS[k] || k) + '.');
  }
  if (f.mau === 'hocvien' && !anh && have.indexOf('anh') < 0) return fail_('Vui lòng chọn ảnh chân dung.');
  if (anh && !photoOk_(anh)) return fail_('Ảnh chưa hợp lệ, vui lòng chọn ảnh khác.');
  if (ans.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(ans.email)) return fail_('Email chưa đúng.');
  ans.guiLuc = now_();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (known.length) {
      // Học viên đã có: chỉ bổ sung câu trả lời mới, không xóa thông tin cũ; cập nhật mọi lớp của người này.
      const add = {};
      Object.keys(ans).forEach(k => { if (k !== 'hoTen' && k !== 'sdt' && filled_(ans[k])) add[k] = ans[k]; });
      known.forEach(r => {
        let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {}
        Object.assign(d, add);
        d.hanhTrinh = (d.hanhTrinh || []).concat([{ngay: today_(), noiDung: 'Bổ sung thông tin qua phiếu: ' + f.ten}]);
        if (anh) setPhoto_(r, anh);
        Object.assign(r, {email: d.email || '', updatedAt: now_(), data: JSON.stringify(d)});
        put_('records', r);
      });
      return out_({ok: true, updated: known.length});
    }
    const old = rows_('records').filter(r => r.formId === f.id && phoneN_(r.sdt) === sdt)[0];
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
    Object.assign(r, {hoTen: d.hoTen, sdt: phone_(d.sdt), email: d.email || '', trangThai: d.trangThai, phuTrach: (d.phuTrach || []).join(', '), updatedAt: now_(), data: JSON.stringify(d)});
    put_('records', r);
  } finally { lock.releaseLock(); }
  return out_({ok: true});
}

/** Chạy hàm này một lần trong trình soạn thảo để tạo các trang tính và cấp quyền. */
function caiDat() { seed_(); sheet_('records'); sheet_('settings'); sheet_('camps'); sheet_('txns'); sheet_('users'); secret_(); }

/* ---------- Chiến dịch (bảng giá, cú pháp chuyển khoản) + Kế toán đối soát sao kê ---------- */
function code_(s, n) { return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/Đ/g, 'D').replace(/[^A-Z0-9]/g, '').slice(0, n || 10); }
function day_(s) { s = String(s || ''); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function campIn_(c) {
  c = c && typeof c === 'object' ? c : {};
  const tk = c.taiKhoan && typeof c.taiKhoan === 'object' ? c.taiKhoan : {};
  const sp = (Array.isArray(c.sp) ? c.sp : []).slice(0, 30).map(p => ({
    ma: code_(p && p.ma, 8), ten: str_(p && p.ten, 120), khoa: str_(p && p.khoa, 60) || 'Diamond', nhom: str_(p && p.nhom, 80),
    hocPhi: money_(p && p.hocPhi), uuDai: money_(p && p.uuDai), coc: money_(p && p.coc),
    dot: (Array.isArray(p && p.dot) ? p.dot : []).slice(0, 12).map(d => ({soTien: money_(d && d.soTien), han: day_(d && d.han)})).filter(d => d.soTien),
  })).filter(p => p.ma && p.ten);
  return {ma: code_(c.ma, 10), mau: str_(c.mau, 60), ten: str_(c.ten, 160), tu: day_(c.tu), den: day_(c.den), phuTrach: list_(c.phuTrach, 20, 80), ghiChu: str_(c.ghiChu, 2000),
    taiKhoan: {bin: /^\d{6}$/.test(String(tk.bin || '')) ? String(tk.bin) : '', nganHang: str_(tk.nganHang, 80), soTK: String(tk.soTK || '').replace(/[^0-9A-Za-z]/g, '').slice(0, 30), chuTK: code_(tk.chuTK, 60) ? str_(tk.chuTK, 80).toUpperCase() : ''},
    sp: sp};
}
function campOut_(r) { let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {} let v = null; try { v = r.duyet ? JSON.parse(r.duyet) : null; } catch (e) {} return Object.assign(d, {id: r.id, ma: r.ma, ten: r.ten, duyet: v, updatedAt: r.updatedAt}); }
function camps_() { return rows_('camps').map(campOut_); }
function txnOut_(r) { let a = []; try { a = JSON.parse(r.alloc || '[]'); } catch (e) {} return {id: r.id, ngay: r.ngay, soTien: Number(r.soTien) || 0, noiDung: r.noiDung, nguoiCK: r.nguoiCK, taiKhoan: r.taiKhoan, trangThai: r.trangThai, alloc: a, ghiChu: r.ghiChu, createdAt: r.createdAt, xacNhanLuc: r.xacNhanLuc, lo: r.lo || '', hd: (() => { try { return r.hd ? JSON.parse(r.hd) : null; } catch (e) { return null; } })()}; }
function uploads_() { return rows_('uploads').slice(-1000).map(u => ({id: u.id, ten: u.ten, taiKhoan: u.taiKhoan, soDong: Number(u.soDong) || 0, them: Number(u.them) || 0, trung: Number(u.trung) || 0, tongTien: Number(u.tongTien) || 0, tuNgay: u.tuNgay, denNgay: u.denNgay, nguoi: u.nguoi, luc: u.luc})); }
function txns_() { return rows_('txns').slice(-3000).map(txnOut_); }
function recData_(r) { let d = {}; try { d = JSON.parse(r.data || '{}'); } catch (e) {} return d; }
function recSet_(r, d) { r.data = JSON.stringify(d); r.updatedAt = now_(); put_('records', r); }
function payState_(d) { if (d.ttTT === 'Bảo lưu / chưa chốt') return d.ttTT; const phai = Math.max(0, (Number(d.hocPhi) || 0) - (Number(d.khauTru) || 0)); if (!phai) return d.ttTT || ''; return (Number(d.daDong) || 0) >= phai ? 'Đã đóng đủ' : (Number(d.daDong) || 0) > 0 ? 'Đặt cọc' : 'Chưa đóng'; }
const KT_ = {
  camp_save: b => {
    const c = campIn_(b.camp), id = str_(b.camp && b.camp.id, 32);
    if (!c.ten) return fail_('Vui lòng đặt tên chiến dịch.');
    if (!c.ma) return fail_('Vui lòng đặt mã chiến dịch (chữ không dấu, số), ví dụ TR1211.');
    const all = rows_('camps');
    if (all.some(x => x.ma === c.ma && x.id !== id)) return fail_('Mã chiến dịch này đã có, vui lòng đặt mã khác.');
    if (all.some(x => x.ten === c.ten && x.id !== id)) return fail_('Tên chiến dịch này đã có.');
    const codes = c.sp.map(p => p.ma); if (codes.some((m, i) => codes.indexOf(m) !== i)) return fail_('Mã khóa trong bảng giá bị trùng.');
    const cur = all.filter(x => x.id === id)[0] || {id: uid_()};
    const old = cur.data ? recData_(cur) : null;
    if (cur.duyet && old && JSON.stringify([old.sp, old.taiKhoan]) !== JSON.stringify([c.sp, c.taiKhoan])) return fail_('Bảng giá đã được kế toán duyệt và đang khóa. Bấm "Mở khóa để sửa" trước.');
    if (cur.ten && cur.ten !== c.ten) { // đổi tên: cập nhật chương trình trong hồ sơ học viên
      rows_('records').forEach(r => { const d = recData_(r); if (d.chuongTrinh === cur.ten) { d.chuongTrinh = c.ten; recSet_(r, d); } });
    }
    Object.assign(cur, {ma: c.ma, ten: c.ten, data: JSON.stringify(c), updatedAt: now_()});
    put_('camps', cur); return out_({ok: true, camps: camps_()});
  },
  camp_approve: b => {
    const r = rows_('camps').filter(x => x.id === str_(b.id, 32))[0]; if (!r) return fail_('Không tìm thấy chiến dịch.', 404);
    r.duyet = b.on ? JSON.stringify({nguoi: str_(b.nguoi, 80) || 'Kế toán', luc: now_()}) : '';
    r.updatedAt = now_(); put_('camps', r); return out_({ok: true, camps: camps_()});
  },
  camp_delete: b => {
    const r = rows_('camps').filter(x => x.id === str_(b.id, 32))[0]; if (!r) return out_({ok: true, camps: camps_()});
    if (r.duyet) return fail_('Chiến dịch đã duyệt, mở khóa trước khi xóa.');
    if (rows_('records').some(x => recData_(x).chuongTrinh === r.ten)) return fail_('Chiến dịch đã có học viên nên không xóa được.');
    sheet_('camps').deleteRow(r._row); return out_({ok: true, camps: camps_()});
  },
  txn_import: b => {
    const list = Array.isArray(b.rows) ? b.rows.slice(0, 3000) : [], have = {};
    rows_('txns').forEach(r => have[r.key] = 1);
    let add = 0, dup = 0, sum = 0, d1 = '', d2 = ''; const out = [], t = now_(), tk = str_(b.taiKhoan, 80), lo = uid_();
    list.forEach(x => {
      const ngay = day_(x && x.ngay), so = money_(x && x.soTien), nd = str_(x && x.noiDung, 500), ma = str_(x && x.maGD, 80);
      if (!ngay || !so) return;
      const key = ma ? 'm:' + tk + ':' + ma : 'h:' + sha_([tk, ngay, so, nd].join('|')).slice(0, 24);
      if (have[key]) { dup++; return; } have[key] = 1;
      out.push([uid_(), key, ngay, String(so), nd, str_(x.nguoiCK, 120), tk, 'cho', '[]', '', t, '', lo, '']); add++; sum += so;
      if (!d1 || ngay < d1) d1 = ngay; if (!d2 || ngay > d2) d2 = ngay;
    });
    if (out.length) { const s = sheet_('txns'); if (s.getLastColumn() < COLS.txns.length) s.getRange(1, 1, 1, COLS.txns.length).setValues([COLS.txns]).setFontWeight('bold'); s.getRange(s.getLastRow() + 1, 1, out.length, COLS.txns.length).setValues(out); }
    put_('uploads', {id: lo, ten: str_(b.file, 160) || 'Dữ liệu dán', taiKhoan: tk, soDong: String(list.length), them: String(add), trung: String(dup), tongTien: String(sum), tuNgay: d1, denNgay: d2, nguoi: (b.__me && b.__me.ten) || '', luc: t});
    return out_({ok: true, added: add, dup: dup, txns: txns_(), uploads: uploads_()});
  },
  txn_confirm: b => {
    const tx = rows_('txns').filter(x => x.id === str_(b.id, 32))[0]; if (!tx) return fail_('Không tìm thấy giao dịch.', 404);
    if (tx.trangThai === 'da') return fail_('Giao dịch này đã được xác nhận rồi.');
    const al = (Array.isArray(b.alloc) ? b.alloc : []).slice(0, 10).map(a => ({recId: str_(a && a.recId, 32), soTien: money_(a && a.soTien), dot: str_(a && a.dot, 60)})).filter(a => a.recId && a.soTien);
    if (!al.length) return fail_('Chưa chọn học viên.');
    if (al.reduce((s, a) => s + a.soTien, 0) !== Number(tx.soTien)) return fail_('Tổng tiền chia cho các khóa phải bằng đúng số tiền giao dịch.');
    const recs = rows_('records'), changed = [];
    for (const a of al) if (!recs.some(r => r.id === a.recId)) return fail_('Không tìm thấy hồ sơ học viên.', 404);
    al.forEach(a => {
      const r = recs.filter(x => x.id === a.recId)[0], d = recData_(r);
      d.daDong = (Number(d.daDong) || 0) + a.soTien;
      d.lichSuTT = (d.lichSuTT || []).concat([{ngay: tx.ngay, soTien: String(a.soTien), noiDung: (a.dot ? a.dot + ' · ' : '') + 'Chuyển khoản: ' + (tx.noiDung || '').slice(0, 300) + ' (kế toán đối soát)', gd: tx.id}]);
      d.ttTT = payState_(d);
      recSet_(r, d); a.hoTen = d.hoTen || ''; changed.push(recOut_(r));
    });
    Object.assign(tx, {trangThai: 'da', alloc: JSON.stringify(al), ghiChu: str_(b.ghiChu, 500), xacNhanLuc: now_()}); put_('txns', tx);
    return out_({ok: true, txn: txnOut_(tx), records: changed});
  },
  txn_undo: b => {
    const tx = rows_('txns').filter(x => x.id === str_(b.id, 32))[0]; if (!tx) return fail_('Không tìm thấy giao dịch.', 404);
    const changed = [], dels = [], removed = [];
    if (tx.trangThai === 'da') {
      let al = []; try { al = JSON.parse(tx.alloc || '[]'); } catch (e) {}
      const recs = rows_('records');
      al.forEach(a => {
        const r = recs.filter(x => x.id === a.recId)[0]; if (!r) return; const d = recData_(r);
        const keep = (d.lichSuTT || []).filter(p => p.gd !== tx.id), gone = (d.lichSuTT || []).length - keep.length;
        if (gone && a.moi && !keep.length) { dels.push(r); return; } // hồ sơ tạo nhầm từ chính giao dịch này: xóa luôn
        if (gone) { d.lichSuTT = keep; d.daDong = Math.max(0, (Number(d.daDong) || 0) - a.soTien); d.ttTT = payState_(d); recSet_(r, d); changed.push(recOut_(r)); }
      });
      dels.sort((x, y) => y._row - x._row).forEach(r => { sheet_('records').deleteRow(r._row); removed.push(r.id); });
    }
    Object.assign(tx, {trangThai: 'cho', alloc: '[]', ghiChu: '', xacNhanLuc: '', hd: ''}); put_('txns', tx);
    return out_({ok: true, txn: txnOut_(tx), records: changed, removed: removed});
  },
  txn_skip: b => {
    const tx = rows_('txns').filter(x => x.id === str_(b.id, 32))[0]; if (!tx) return fail_('Không tìm thấy giao dịch.', 404);
    if (tx.trangThai === 'da') return fail_('Giao dịch đã xác nhận, hãy hoàn tác trước.');
    Object.assign(tx, {trangThai: 'bo', ghiChu: str_(b.ghiChu, 500) || 'Không phải học phí', xacNhanLuc: now_()}); put_('txns', tx);
    return out_({ok: true, txn: txnOut_(tx)});
  },
  txn_newrecs: b => {
    // Tạo hồ sơ học viên mới từ giao dịch sao kê (chưa có trong CRM) rồi ghi tiền luôn. Cùng SĐT + khóa + nhóm + chương trình thì dùng chung hồ sơ.
    const items = (Array.isArray(b.items) ? b.items : []).slice(0, 300), recs = rows_('records'), txs = rows_('txns'), changed = {}, outTx = [];
    const keyOf = d => [phone_(d.sdt) || String(d.hoTen || '').toLowerCase(), d.khoa || '', d.nhom || '', d.chuongTrinh || ''].join('|');
    let made = 0, bad = 0;
    items.forEach(it => {
      const tx = txs.filter(x => x.id === str_(it && it.id, 32))[0], R = (it && it.rec) || {};
      const hoTen = str_(R.hoTen, 120), sd = String(R.sdt || '').replace(/[\s.\-]/g, '');
      if (!tx || tx.trangThai !== 'cho' || !hoTen || (sd && !/^\+?\d{9,12}$/.test(sd))) { bad++; return; }
      const so = Number(tx.soTien) || 0, dot = str_(it.dot, 60) || 'Đợt 1';
      const pay = {ngay: tx.ngay, soTien: String(so), noiDung: dot + ' · Chuyển khoản: ' + (tx.noiDung || '').slice(0, 300) + ' (kế toán đối soát)', gd: tx.id};
      const nd = {hoTen: hoTen, sdt: sd, khoa: str_(R.khoa, 60) || 'Diamond', nhom: str_(R.nhom, 80), chuongTrinh: str_(R.chuongTrinh, 160), hocPhi: money_(R.hocPhi), khauTru: 0, daDong: 0, nam: '2026', trangThai: 'Mới đăng ký', phuTrach: list_(R.phuTrach, 10, 80), hanhTrinh: [{ngay: tx.ngay, noiDung: 'Tạo hồ sơ từ sao kê ngân hàng (kế toán đối soát)'}], nhatKy: [], tienDo: {}, lichSuTT: []};
      let r = recs.filter(x => keyOf(recData_(x)) === keyOf(nd))[0];
      let isNew = 0;
      if (r) { const d = recData_(r); d.daDong = (Number(d.daDong) || 0) + so; d.lichSuTT = (d.lichSuTT || []).concat([pay]); d.ttTT = payState_(d); recSet_(r, d); }
      else { isNew = 1;
        nd.daDong = so; nd.lichSuTT = [pay]; nd.ttTT = payState_(nd);
        r = {id: uid_(), formId: '', createdAt: now_(), data: JSON.stringify(nd), hoTen: hoTen, sdt: phone_(sd), email: '', trangThai: 'Mới đăng ký', phuTrach: nd.phuTrach.join(', '), updatedAt: now_()};
        setPhoto_(r, ''); put_('records', r); r._row = sheet_('records').getLastRow(); recs.push(r); made++;
      }
      changed[r.id] = r;
      Object.assign(tx, {trangThai: 'da', alloc: JSON.stringify([{recId: r.id, soTien: so, dot: dot, hoTen: hoTen, moi: isNew}]), ghiChu: isNew ? 'Tạo hồ sơ mới từ sao kê' : 'Ghi vào hồ sơ đang có', xacNhanLuc: now_()}); put_('txns', tx); outTx.push(txnOut_(tx));
    });
    return out_({ok: true, made: made, bad: bad, txns: outTx, records: Object.keys(changed).map(k => recOut_(changed[k]))});
  },
  txn_invoice: b => {
    // Hóa đơn: cập nhật trạng thái / số hóa đơn / thông tin người mua cho 1 hoặc nhiều giao dịch đã đối soát; lưu lịch sử ai sửa lúc nào.
    const ids = (Array.isArray(b.ids) ? b.ids : []).slice(0, 500).map(x => str_(x, 32)), P = b.patch || {}, me = (b.__me && b.__me.ten) || '';
    const TT = ['cho', 'thieu', 'da', 'khong'];
    if (P.tt && TT.indexOf(P.tt) < 0) return fail_('Trạng thái hóa đơn chưa đúng.');
    const all = rows_('txns'), out = [];
    ids.forEach(id => {
      const tx = all.filter(x => x.id === id)[0]; if (!tx || tx.trangThai !== 'da') return;
      let h = {}; try { h = tx.hd ? JSON.parse(tx.hd) : {}; } catch (e) {}
      const ch = [];
      if (P.loai && ['canhan', 'congty'].indexOf(P.loai) < 0) delete P.loai;
      ['tt', 'so', 'loai', 'ngay', 'ten', 'mst', 'dc', 'email', 'ghiChu'].forEach(k => { if (k in P) { const v = k === 'ngay' ? day_(P[k]) : str_(P[k], k === 'dc' || k === 'ghiChu' ? 500 : 160); if (v !== (h[k] || '')) { ch.push(k); h[k] = v; } } });
      if (!ch.length) return;
      h.nguoi = me; h.capNhat = now_(); h.lichSu = (h.lichSu || []).concat([{luc: h.capNhat, nguoi: me, tt: h.tt || 'cho', so: h.so || '', thay: ch.join(',')}]).slice(-30);
      tx.hd = JSON.stringify(h); put_('txns', tx); out.push(txnOut_(tx));
    });
    return out_({ok: true, txns: out});
  },
  txn_delete: b => {
    const tx = rows_('txns').filter(x => x.id === str_(b.id, 32))[0]; if (!tx) return out_({ok: true});
    if (tx.trangThai === 'da') return fail_('Giao dịch đã xác nhận, hãy hoàn tác trước.');
    sheet_('txns').deleteRow(tx._row); return out_({ok: true});
  },
};

/* ---------- Tài khoản nhân sự & phân quyền theo phòng ban ---------- */
const ROLES_ = {quantri: 'Toàn quyền', cskh: 'Chăm sóc khách hàng', marketing: 'Marketing', ketoan: 'Kế toán'};
const MONEY_KEYS_ = ['hocPhi', 'khauTru', 'daDong', 'lichSuTT', 'ttTT', 'hanTT', 'canhBao', 'ghiChuTT', 'dongHanh'];
const PERM_ = {
  record_save: ['quantri', 'cskh', 'ketoan', 'marketing'], record_create: ['quantri', 'cskh', 'ketoan', 'marketing'], photo_save: ['quantri', 'cskh', 'marketing'],
  record_delete: ['quantri'], form_save: ['quantri', 'marketing'], form_delete: ['quantri'], staff_save: ['quantri'],
  camp_save: ['quantri', 'marketing'], camp_approve: ['quantri', 'ketoan'], camp_delete: ['quantri', 'marketing'],
  txn_import: ['quantri', 'ketoan'], txn_confirm: ['quantri', 'ketoan'], txn_undo: ['quantri', 'ketoan'], txn_skip: ['quantri', 'ketoan'], txn_delete: ['quantri', 'ketoan'], txn_newrecs: ['quantri', 'ketoan'], txn_invoice: ['quantri', 'ketoan'],
  user_save: ['quantri'], user_delete: ['quantri'], gpt_list: ['quantri'], gpt_create: ['quantri'], gpt_revoke: ['quantri'],
  zl_dev_list: ['quantri', 'cskh', 'marketing', 'ketoan'], zl_dev_create: ['quantri', 'cskh', 'marketing', 'ketoan'], zl_dev_revoke: ['quantri'], zl_perm_set: ['quantri'], zl_reset: ['quantri'], zl_threads: ['quantri', 'cskh', 'marketing', 'ketoan'], zl_msgs: ['quantri', 'cskh', 'marketing', 'ketoan'],
};
function pwHash_(salt, pw) { let h = sha_(salt + ':' + pw); for (let i = 0; i < 1000; i++) h = sha_(h + salt); return h; }
function loginKey_(s) { s = String(s || '').trim().toLowerCase(); if (s.indexOf('@') > 0) return s; const p = phone_(s); return p || s; }
function userOut_(u) { return {id: u.id, hoTen: u.hoTen, sdt: u.sdt, email: u.email, phong: u.phong, active: u.active === '1', coMK: !!u.hash, lastLogin: u.lastLogin || '', createdAt: String(u.createdAt || '').slice(0, 10)}; }
function users_() { return rows_('users').map(userOut_); }
function who_(uid) {
  if (uid === ADMIN_USER) return {uid: uid, role: 'quantri', ten: 'Quản trị (crmtkt)', phong: 'quantri'};
  const u = rows_('users').filter(x => x.id === uid)[0];
  if (!u || u.active !== '1' || !ROLES_[u.phong]) return null;
  return {uid: u.id, role: u.phong, ten: u.hoTen, phong: u.phong};
}
function login_(b) {
  const c = CacheService.getScriptCache(), id = loginKey_(str_(b.u, 120)), ck = 'fail:' + sha_(id).slice(0, 16), n = Number(c.get(ck) || 0);
  if (n >= 8) return fail_('Sai quá nhiều lần. Vui lòng thử lại sau 15 phút.', 429);
  const pw = String(b.p || '');
  if (id === ADMIN_USER) {
    if (!passOk_(pw)) { c.put(ck, String(n + 1), 900); return fail_('Sai tài khoản hoặc mật khẩu.', 401); }
    return out_({ok: true, tok: makeToken_(ADMIN_USER), user: 'Quản trị (crmtkt)', me: who_(ADMIN_USER)});
  }
  const u = rows_('users').filter(x => x.active === '1' && x.hash && (String(x.email).toLowerCase() === id || phone_(x.sdt) === id))[0];
  if (!u || pwHash_(u.salt, pw) !== u.hash) { c.put(ck, String(n + 1), 900); return fail_('Sai tài khoản hoặc mật khẩu.', 401); }
  u.lastLogin = now_(); put_('users', u);
  return out_({ok: true, tok: makeToken_(u.id), user: u.hoTen, me: who_(u.id)});
}
const US_ = {
  user_save: b => {
    const U = b.user || {}, id = str_(U.id, 32), all = rows_('users');
    const hoTen = str_(U.hoTen, 80), sdt = String(U.sdt || '').replace(/[\s.\-]/g, ''), email = str_(U.email, 120).toLowerCase(), phong = str_(U.phong, 20), pw = String(U.pw || '');
    if (!hoTen) return fail_('Vui lòng nhập họ tên.');
    if (!ROLES_[phong]) return fail_('Vui lòng chọn phòng ban.');
    if (sdt && !/^\+?\d{9,12}$/.test(sdt)) return fail_('Số điện thoại chưa đúng.');
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail_('Email chưa đúng.');
    if (!sdt && !email) return fail_('Cần có số điện thoại hoặc email để nhân sự đăng nhập.');
    if (pw && pw.length < 8) return fail_('Mật khẩu cần ít nhất 8 ký tự.');
    if (sdt && all.some(x => x.id !== id && phone_(x.sdt) === phone_(sdt))) return fail_('Số điện thoại này đã dùng cho tài khoản khác.');
    if (email && all.some(x => x.id !== id && String(x.email).toLowerCase() === email)) return fail_('Email này đã dùng cho tài khoản khác.');
    const u = all.filter(x => x.id === id)[0] || {id: uid_(), createdAt: now_(), salt: '', hash: '', lastLogin: ''};
    Object.assign(u, {hoTen: hoTen, sdt: sdt ? phone_(sdt) : '', email: email, phong: phong, active: U.active === false ? '0' : '1', updatedAt: now_()});
    if (pw) { u.salt = Utilities.getUuid().replace(/-/g, ''); u.hash = pwHash_(u.salt, pw); }
    put_('users', u); return out_({ok: true, users: users_(), staff: staff_()});
  },
  user_delete: b => {
    const u = rows_('users').filter(x => x.id === str_(b.id, 32))[0]; if (u) sheet_('users').deleteRow(u._row);
    return out_({ok: true, users: users_(), staff: staff_()});
  },
  pw_change: (b, ME) => {
    if (ME.uid === ADMIN_USER) return fail_('Tài khoản quản trị crmtkt đổi mật khẩu trong mã máy chủ (xem hướng dẫn).');
    const u = rows_('users').filter(x => x.id === ME.uid)[0]; if (!u) return fail_('Không tìm thấy tài khoản.', 404);
    if (pwHash_(u.salt, String(b.old || '')) !== u.hash) return fail_('Mật khẩu hiện tại chưa đúng.');
    const pw = String(b.pw || ''); if (pw.length < 8) return fail_('Mật khẩu mới cần ít nhất 8 ký tự.');
    u.salt = Utilities.getUuid().replace(/-/g, ''); u.hash = pwHash_(u.salt, pw); u.updatedAt = now_(); put_('users', u);
    return out_({ok: true});
  },
};

/* ---------- Kết nối ChatGPT (chỉ đọc) ----------
   Mỗi mã kết nối gắn với 1 tài khoản CRM, nên ChatGPT chỉ thấy đúng phần tài khoản đó được thấy.
   Mã lưu dạng băm trong Script Properties; quản trị tạo / thu hồi ở ô Nhân sự. */
function gkLoad_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('GPTKEYS') || '[]'); } catch (e) { return []; } }
function gkSave_(l) { PropertiesService.getScriptProperties().setProperty('GPTKEYS', JSON.stringify(l.slice(-20))); }
function gkOut_(k) { const u = who_(k.uid); return {id: k.id, ten: k.ten, uid: k.uid, nguoi: u ? u.ten : '(tài khoản không còn)', quyen: u ? ROLES_[u.role] : '', active: !!k.active && !!u, createdAt: k.createdAt, lastUsed: k.lastUsed || '', dem: k.dem || 0, duoi: k.duoi, by: k.by || ''}; }
const GP_ = {
  gpt_list: () => out_({ok: true, keys: gkLoad_().map(gkOut_)}),
  gpt_create: (b, ME) => {
    const pw = String(b.pw || '');
    if (ME.uid === ADMIN_USER ? !passOk_(pw) : !(rows_('users').filter(x => x.id === ME.uid && x.hash && pwHash_(x.salt, pw) === x.hash)[0])) return fail_('Mật khẩu xác nhận chưa đúng, chưa tạo mã.', 403);
    const uid = str_(b.uid, 40) || ADMIN_USER; if (!who_(uid)) return fail_('Tài khoản này đã khóa hoặc không còn.');
    const key = 'skt_' + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    const l = gkLoad_(), k = {id: uid_(), h: sha_(key), uid: uid, ten: str_(b.ten, 80) || 'ChatGPT', active: true, createdAt: now_(), by: ME.ten, duoi: key.slice(-4), dem: 0};
    l.push(k); gkSave_(l);
    return out_({ok: true, key: key, item: gkOut_(k)});
  },
  gpt_revoke: b => { const l = gkLoad_(), k = l.filter(x => x.id === str_(b.id, 32))[0]; if (!k) return fail_('Không thấy mã.'); k.active = false; gkSave_(l); return out_({ok: true}); },
};
const GPT_HIDE_ = ['cccd', 'diaChi', 'phieuOffline', 'anh', 'phV', 'hasPhoto', 'formId', 'dongY', 'zaloGroup'];
function gNum_(v) { return Number(v) || 0; }
function gRec_(d, role, full) {
  const phai = Math.max(0, gNum_(d.hocPhi) - gNum_(d.khauTru)), o = {
    id: d.id, hoTen: d.hoTen, sdt: d.sdt, email: d.email, khoa: COURSE[d.khoa] || d.khoa || '', nhom: d.nhom || '', chuongTrinh: d.chuongTrinh || '',
    trangThai: d.trangThai || 'Mới đăng ký', phuTrach: (d.phuTrach || []).join(', '), ngayDangKy: d.createdAt,
    lanChamSocCuoi: (d.nhatKy || []).map(e => e.ngay).filter(Boolean).sort().pop() || ''};
  if (role !== 'marketing') Object.assign(o, {hocPhi: gNum_(d.hocPhi), khauTru: gNum_(d.khauTru), phaiThu: phai, daThu: gNum_(d.daDong), congNo: Math.max(0, phai - gNum_(d.daDong)), thanhToan: d.ttTT || payState_(d), hanDongNot: d.hanTT || ''});
  if (full) {
    Object.keys(d).forEach(k => { if (!(k in o) && GPT_HIDE_.indexOf(k) < 0 && (role !== 'marketing' || MONEY_KEYS_.indexOf(k) < 0)) o[k] = d[k]; });
    o.nhatKy = (d.nhatKy || []).slice(-40); o.hanhTrinh = (d.hanhTrinh || []).slice(-20);
  }
  return o;
}
function gIn_(s, a, b) { s = String(s || '').slice(0, 10); return (!a || s >= a) && (!b || s <= b); }
function gNorm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd'); }
function gpt_(e) {
  const P = e.parameter || {}, c = CacheService.getScriptCache();
  const key = String(P.key || ''); if (!/^skt_[0-9a-f]{40}$/.test(key)) return fail_('Thiếu hoặc sai mã kết nối (key).', 401);
  const l = gkLoad_(), k = l.filter(x => x.h === sha_(key))[0]; if (!k || !k.active) return fail_('Mã kết nối không hợp lệ hoặc đã bị thu hồi.', 401);
  const ME = who_(k.uid); if (!ME) return fail_('Tài khoản gắn với mã này đã bị khóa.', 401);
  const rk = 'gpt:' + k.id, n = Number(c.get(rk) || 0); if (n > 300) return fail_('Quá nhiều yêu cầu, thử lại sau ít phút.', 429); c.put(rk, String(n + 1), 600);
  if (!c.get('gptu:' + k.id)) { c.put('gptu:' + k.id, '1', 300); k.lastUsed = now_(); k.dem = (k.dem || 0) + 1; gkSave_(l); }
  const op = String(e.pathInfo || P.op || 'tongquan').replace(/^\/+|\/+$/g, ''), role = ME.role;
  const tu = day_(P.tuNgay), den = day_(P.denNgay), today = today_();
  const recs = rows_('records').map(recOut_);
  const base = {ok: true, nguoiDung: ME.ten, quyen: ROLES_[role], homNay: today, donViTien: 'VND'};
  if (op === 'tongquan') {
    const sel = recs.filter(d => gIn_(d.createdAt, tu, den)), by = (f) => { const m = {}; sel.forEach(d => { const ks = f(d); (Array.isArray(ks) ? (ks.length ? ks : ['(chưa phân công)']) : [ks || '(chưa có)']).forEach(x => { const g = m[x] || (m[x] = {so: 0, phaiThu: 0, daThu: 0, congNo: 0}); const r = gRec_(d, 'quantri'); g.so++; g.phaiThu += r.phaiThu; g.daThu += r.daThu; g.congNo += r.congNo; }); }); return Object.keys(m).map(x => Object.assign({ten: x}, role === 'marketing' ? {so: m[x].so} : m[x])); };
    const nguoi = {}; sel.forEach(d => nguoi[phoneN_(d.sdt) || d.id] = 1);
    const R = sel.map(d => gRec_(d, role)), old = new Date(Date.now() - 14 * 864e5).toISOString().slice(0, 10), soon = new Date(Date.now() + 7 * 864e5).toISOString().slice(0, 10);
    const careBy = {}; recs.forEach(d => (d.nhatKy || []).forEach(x => { if (gIn_(x.ngay, tu || old, den || today)) careBy[x.nguoi || '(không rõ)'] = (careBy[x.nguoi || '(không rõ)'] || 0) + 1; }));
    const o = Object.assign(base, {pham_vi: tu || den ? ('đăng ký từ ' + (tu || 'đầu') + ' đến ' + (den || today)) : 'toàn bộ', soLuotDangKy: sel.length, soHocVien: Object.keys(nguoi).length,
      theoTrangThai: by(d => d.trangThai || 'Mới đăng ký').map(x => ({ten: x.ten, so: x.so})), theoChuongTrinh: by(d => d.chuongTrinh), theoKhoa: by(d => (COURSE[d.khoa] || d.khoa || '') + (d.nhom ? ' · ' + d.nhom : '')), theoNhanSuPhuTrach: by(d => d.phuTrach || []),
      soLanChamSocTheoNguoi: {pham_vi: (tu || old) + ' → ' + (den || today), so: careBy},
      canChuY: {
        canChamSocGap: R.filter(r => r.trangThai === 'Cần chăm sóc gấp').map(r => ({id: r.id, hoTen: r.hoTen, phuTrach: r.phuTrach})),
        chuaPhanCong: R.filter(r => !r.phuTrach).map(r => ({id: r.id, hoTen: r.hoTen, chuongTrinh: r.chuongTrinh})),
        chuaChamSoc14Ngay: R.filter(r => !r.lanChamSocCuoi || r.lanChamSocCuoi < old).map(r => ({id: r.id, hoTen: r.hoTen, phuTrach: r.phuTrach, lanChamSocCuoi: r.lanChamSocCuoi || 'chưa có'})).slice(0, 60),
        quaHanDongTien: role === 'marketing' ? undefined : R.filter(r => r.congNo > 0 && r.hanDongNot && r.hanDongNot < today).map(r => ({id: r.id, hoTen: r.hoTen, congNo: r.congNo, han: r.hanDongNot, phuTrach: r.phuTrach})),
        sapDenHanDong7Ngay: role === 'marketing' ? undefined : R.filter(r => r.congNo > 0 && r.hanDongNot >= today && r.hanDongNot <= soon).map(r => ({id: r.id, hoTen: r.hoTen, congNo: r.congNo, han: r.hanDongNot}))
      }});
    if (role !== 'marketing') { o.tongPhaiThu = R.reduce((a, r) => a + r.phaiThu, 0); o.tongDaThu = R.reduce((a, r) => a + r.daThu, 0); o.tongCongNo = R.reduce((a, r) => a + r.congNo, 0); }
    if (role === 'quantri' || role === 'ketoan') { const t = txns_(); o.keToan = {choXacNhan: {so: t.filter(x => x.trangThai === 'cho').length, tien: t.filter(x => x.trangThai === 'cho').reduce((a, x) => a + x.soTien, 0)}, hoaDonChuaXuat: t.filter(x => x.trangThai === 'da' && (!x.hd || !x.hd.tt || x.hd.tt === 'cho' || x.hd.tt === 'thieu')).length}; }
    return out_(o);
  }
  if (op === 'hocvien') {
    const q = gNorm_(P.q), lim = Math.min(200, Math.max(1, parseInt(P.limit, 10) || 50));
    let L = recs.filter(d => gIn_(d.createdAt, tu, den)).map(d => gRec_(d, role));
    if (q) L = L.filter(r => gNorm_([r.hoTen, r.sdt, r.email].join(' ')).indexOf(q) >= 0 || String(r.sdt).replace(/\D/g, '').indexOf(String(P.q).replace(/\D/g, '') || '~') >= 0);
    if (P.chuongTrinh) L = L.filter(r => gNorm_(r.chuongTrinh).indexOf(gNorm_(P.chuongTrinh)) >= 0);
    if (P.khoa) L = L.filter(r => gNorm_(r.khoa + ' ' + r.nhom).indexOf(gNorm_(P.khoa)) >= 0);
    if (P.phuTrach) L = L.filter(r => gNorm_(r.phuTrach).indexOf(gNorm_(P.phuTrach)) >= 0);
    if (P.trangThai) L = L.filter(r => gNorm_(r.trangThai) === gNorm_(P.trangThai));
    if (String(P.conNo) === 'true' && role !== 'marketing') L = L.filter(r => r.congNo > 0);
    return out_(Object.assign(base, {tong: L.length, hienThi: Math.min(lim, L.length), hocVien: L.slice(0, lim)}));
  }
  if (op === 'chitiet') {
    const id = str_(P.id, 40), sdt = phoneN_(P.sdt);
    const L = recs.filter(d => (id && d.id === id) || (sdt && phoneN_(d.sdt) === sdt));
    if (!L.length) return fail_('Không tìm thấy học viên.', 404);
    return out_(Object.assign(base, {soDangKy: L.length, dangKy: L.map(d => gRec_(d, role, true))}));
  }
  if (op === 'nhatky') {
    const a = tu || new Date(Date.now() - 7 * 864e5).toISOString().slice(0, 10), b2 = den || today, ng = gNorm_(P.nguoi), out = [];
    recs.forEach(d => (d.nhatKy || []).forEach(x => { if (gIn_(x.ngay, a, b2) && (!ng || gNorm_(x.nguoi).indexOf(ng) >= 0)) out.push({ngay: x.ngay, nguoi: x.nguoi, hocVien: d.hoTen, id: d.id, chuongTrinh: d.chuongTrinh || '', noiDung: x.noiDung}); }));
    out.sort((x, y) => x.ngay < y.ngay ? 1 : -1);
    const dem = {}; out.forEach(x => dem[x.nguoi || '(không rõ)'] = (dem[x.nguoi || '(không rõ)'] || 0) + 1);
    return out_(Object.assign(base, {tuNgay: a, denNgay: b2, tongSoLan: out.length, theoNguoi: dem, nhatKy: out.slice(0, 300)}));
  }
  if (op === 'doisoat') {
    if (role !== 'quantri' && role !== 'ketoan') return fail_('Tài khoản ' + ROLES_[role] + ' không xem được phần kế toán.', 403);
    const t = txns_().filter(x => gIn_(x.ngay, tu, den)), day = {};
    t.forEach(x => { if (x.trangThai === 'bo') return; const g = day[x.ngay] || (day[x.ngay] = {ngay: x.ngay, so: 0, tien: 0, daXacNhan: 0, choXacNhan: 0}); g.so++; g.tien += x.soTien; if (x.trangThai === 'da') g.daXacNhan += x.soTien; else g.choXacNhan += x.soTien; });
    const hd = {}; t.filter(x => x.trangThai === 'da').forEach(x => { const s = (x.hd && x.hd.tt) || 'cho'; const lb = {cho: 'Chưa xuất', thieu: 'Thiếu thông tin', da: 'Đã xuất', khong: 'Không xuất'}[s] || s; hd[lb] = (hd[lb] || 0) + 1; });
    return out_(Object.assign(base, {pham_vi: (tu || 'đầu') + ' → ' + (den || today), tienVeTheoNgay: Object.keys(day).sort().map(k => day[k]),
      choXacNhan: t.filter(x => x.trangThai === 'cho').slice(0, 100).map(x => ({ngay: x.ngay, soTien: x.soTien, noiDung: x.noiDung, nguoiCK: x.nguoiCK})), hoaDon: hd}));
  }
  return fail_('Không có mục "' + op + '". Dùng: tongquan, hocvien, chitiet, nhatky, doisoat.', 404);
}

/* ---------- Máy chủ MCP (chỉ đọc) cho ChatGPT / Claude ----------
   Link kết nối: {Web App}/exec?mcp=1&key=skt_...  (mã tạo ở mục "Kết nối AI", chỉ quản trị).
   Dùng lại toàn bộ kiểm tra quyền của gpt_(): mã gắn với 1 tài khoản, đọc đúng phần tài khoản đó được thấy. */
const MCP_TOOLS_ = [
  {name: 'tong_quan', op: 'tongquan', title: 'Tổng quan CRM', description: 'Tổng quan CRM CMA | SKT: số học viên theo trạng thái / chương trình / khóa / nhân sự phụ trách, doanh thu (đã thu), công nợ, số lần chăm sóc của từng nhân sự, việc cần chú ý (cần chăm sóc gấp, chưa phân công, quá 14 ngày chưa chăm sóc, quá hạn / sắp đến hạn đóng tiền), kế toán chờ xác nhận. Gọi đầu tiên cho mọi câu hỏi chung hoặc báo cáo.',
    props: {tuNgay: 'Lọc theo ngày đăng ký, từ ngày YYYY-MM-DD', denNgay: 'Đến ngày YYYY-MM-DD'}},
  {name: 'danh_sach_hoc_vien', op: 'hocvien', title: 'Danh sách học viên', description: 'Danh sách học viên (mỗi lượt đăng ký 1 dòng) có lọc: tên/SĐT/email, chương trình, khóa, nhân sự phụ trách, trạng thái, chỉ người còn công nợ.',
    props: {q: 'Tìm theo tên, SĐT hoặc email', chuongTrinh: 'Tên chương trình / chiến dịch', khoa: 'Khóa hoặc nhóm: Diamond, Trainer, viết sách…', phuTrach: 'Tên nhân sự phụ trách', trangThai: 'Trạng thái chăm sóc', conNo: '"true" để chỉ lấy học viên còn công nợ', limit: 'Số dòng tối đa (mặc định 50, tối đa 200)', tuNgay: 'Từ ngày đăng ký YYYY-MM-DD', denNgay: 'Đến ngày YYYY-MM-DD'}},
  {name: 'chi_tiet_hoc_vien', op: 'chitiet', title: 'Hồ sơ học viên', description: 'Hồ sơ đầy đủ 1 học viên: mọi khóa đăng ký, lịch sử đóng tiền, nhật ký chăm sóc, hành trình, câu trả lời form. Cần id hoặc sdt.',
    props: {id: 'Mã hồ sơ (lấy từ danh sách)', sdt: 'Số điện thoại'}},
  {name: 'nhat_ky_cham_soc', op: 'nhatky', title: 'Nhật ký chăm sóc', description: 'Nhân sự đã làm gì với học viên nào trong khoảng ngày (mặc định 7 ngày gần nhất). Dùng để báo cáo tình hình công việc của nhân sự.',
    props: {tuNgay: 'Từ ngày YYYY-MM-DD', denNgay: 'Đến ngày YYYY-MM-DD', nguoi: 'Lọc theo tên nhân sự'}},
  {name: 'doi_soat_ke_toan', op: 'doisoat', title: 'Đối soát kế toán', description: 'Tiền về theo ngày, khoản chờ xác nhận, tình trạng hóa đơn. Chỉ tài khoản Toàn quyền / Kế toán xem được.',
    props: {tuNgay: 'Từ ngày YYYY-MM-DD', denNgay: 'Đến ngày YYYY-MM-DD'}},
  {name: 'search', op: 'hocvien', title: 'Tìm học viên', description: 'Tìm học viên theo tên, số điện thoại hoặc email. Trả về danh sách kết quả (id, tên).', props: {query: 'Từ khóa tìm kiếm'}, req: ['query']},
  {name: 'fetch', op: 'chitiet', title: 'Mở hồ sơ', description: 'Lấy toàn bộ hồ sơ của 1 học viên theo id (lấy từ kết quả search).', props: {id: 'Mã hồ sơ'}, req: ['id']},
];
const MCP_INFO_ = 'CRM CMA | SKT (công ty đào tạo của thầy Trần Khánh Tư). Kết nối CHỈ ĐỌC. Luôn gọi công cụ để lấy số liệu mới nhất, không đoán. Trả lời tiếng Việt; tiền VND (65.000.000 đ hoặc 65 triệu); luôn viết "công nợ". Báo cáo: số chính trước, bảng 4 cột [Nhóm] · Số lượng · Doanh thu (đã thu) · Công nợ cần thu có dòng Tổng cộng, cuối cùng 3–5 việc cần làm ngay. "Tuần này" tính từ thứ Hai; ngày hôm nay xem trường homNay.';
function mcpOut_(o) { return ContentService.createTextOutput(o == null ? '' : JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function mcp_(e) {
  let body; try { body = JSON.parse(e.postData.contents); } catch (err) { return mcpOut_({jsonrpc: '2.0', id: null, error: {code: -32700, message: 'Parse error'}}); }
  const key = String((e.parameter || {}).key || '');
  const kk = /^skt_[0-9a-f]{40}$/.test(key) && gkLoad_().filter(x => x.h === sha_(key) && x.active)[0];
  if (!kk || !who_(kk.uid)) { const ids = (Array.isArray(body) ? body : [body]).filter(x => x && x.id != null); return mcpOut_(ids.length ? ids.map(x => ({jsonrpc: '2.0', id: x.id, error: {code: -32001, message: 'Link kết nối CRM không hợp lệ hoặc đã bị thu hồi. Vào crmtkt.com → Kết nối AI để tạo link mới.'}})) [0] : null); }
  const one = m => {
    const id = m.id, ok = r => ({jsonrpc: '2.0', id: id, result: r}), er = (c, s) => ({jsonrpc: '2.0', id: id, error: {code: c, message: s}});
    if (id === undefined || id === null) return null; // notification
    if (m.method === 'initialize') return ok({protocolVersion: (m.params && m.params.protocolVersion) || '2025-06-18', capabilities: {tools: {listChanged: false}}, serverInfo: {name: 'crm-cma-skt', title: 'CRM CMA SKT', version: '1.0.0'}, instructions: MCP_INFO_});
    if (m.method === 'ping') return ok({});
    if (m.method === 'tools/list') return ok({tools: MCP_TOOLS_.map(t => ({name: t.name, title: t.title, description: t.description, annotations: {title: t.title, readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false},
      inputSchema: {type: 'object', properties: Object.keys(t.props).reduce((a, k) => { a[k] = {type: 'string', description: t.props[k]}; return a; }, {}), required: t.req || []}}))});
    if (m.method === 'resources/list') return ok({resources: []});
    if (m.method === 'prompts/list') return ok({prompts: []});
    if (m.method === 'tools/call') {
      const p = m.params || {}, t = MCP_TOOLS_.filter(x => x.name === p.name)[0]; if (!t) return er(-32602, 'Không có công cụ ' + p.name);
      const a = Object.assign({}, p.arguments || {}); Object.keys(a).forEach(k => a[k] = String(a[k]));
      if (t.name === 'search') { a.q = a.query; a.limit = '20'; }
      const txt = gpt_({parameter: Object.assign(a, {key: key}), pathInfo: t.op}).getContent();
      let o = {}; try { o = JSON.parse(txt); } catch (err) {}
      if (o.ok === false) return ok({content: [{type: 'text', text: o.error || 'Lỗi'}], isError: true});
      if (t.name === 'search') { const r = {results: (o.hocVien || []).map(h => ({id: h.id, title: h.hoTen + ' · ' + (h.khoa || '') + (h.nhom ? ' ' + h.nhom : '') + ' · ' + (h.chuongTrinh || ''), url: 'https://crmtkt.com/#khach-hang'}))}; return ok({content: [{type: 'text', text: JSON.stringify(r)}], structuredContent: r}); }
      if (t.name === 'fetch') { const d = (o.dangKy || [])[0] || {}; const r = {id: d.id || a.id, title: d.hoTen || '', text: JSON.stringify(o), url: 'https://crmtkt.com/#khach-hang', metadata: {soDangKy: String(o.soDangKy || 0)}}; return ok({content: [{type: 'text', text: JSON.stringify(r)}], structuredContent: r}); }
      return ok({content: [{type: 'text', text: txt}], structuredContent: o});
    }
    return er(-32601, 'Method not found: ' + m.method);
  };
  if (Array.isArray(body)) { const r = body.map(one).filter(x => x); return mcpOut_(r.length ? r : null); }
  return mcpOut_(one(body));
}

/* ---------- Zalo cá nhân: tiện ích Chrome đọc Zalo Web → CRM (chỉ xem) ----------
   Mỗi nick có 1 "mã kết nối Zalo" (zl_...) do quản trị tạo; tiện ích dùng mã này để gửi tin nhắn về.
   Xem tin: quản trị thấy tất cả; nhân sự chỉ thấy nick được quản trị cấp quyền. */
COLS.zacc = ['id', 'ten', 'anh', 'devId', 'chuNick', 'lastSync', 'createdAt'];
COLS.zthr = ['key', 'nick', 'thread', 'ten', 'anh', 'nhom', 'lastTs', 'lastText', 'updatedAt'];
COLS.zmsg = ['key', 'nick', 'thread', 'msgId', 'ts', 'me', 'nguoiGui', 'loai', 'noiDung', 'kieu', 'url', 'tep', 'luc'];
function zdLoad_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('ZALODEV') || '[]'); } catch (e) { return []; } }
function zdSave_(l) { PropertiesService.getScriptProperties().setProperty('ZALODEV', JSON.stringify(l.slice(-20))); }
function zPerm_() { const r = rows_('settings').filter(x => x.k === 'zaloPerm')[0]; try { return r ? JSON.parse(r.v) : {}; } catch (e) { return {}; } }
function zCan_(ME, nick) { if (ME.role === 'quantri') return true; const p = zPerm_()[ME.uid] || []; return p.indexOf('*') >= 0 || p.indexOf(String(nick)) >= 0; }
function zAny_(ME) { if (ME.role === 'quantri') return true; return (zPerm_()[ME.uid] || []).length > 0; }
function zDevOut_(d) { return {id: d.id, ten: d.ten, chuNick: d.chuNick, nick: d.nick || '', nickTen: d.nickTen || '', active: !!d.active, createdAt: d.createdAt, lastUsed: d.lastUsed || '', duoi: d.duoi, by: d.by || ''}; }
const ZL_ = {
  zl_dev_list: (b, ME) => ME.role === 'quantri' ? out_({ok: true, devs: zdLoad_().map(zDevOut_), perm: zPerm_(), accs: rows_('zacc')}) : out_({ok: true, devs: zdLoad_().filter(d => d.chuNick === ME.ten).map(zDevOut_)}),
  zl_dev_create: (b, ME) => {
    if (ME.role !== 'quantri' && str_(b.chuNick, 80) !== ME.ten) return fail_('Anh/chị chỉ kết nối được Zalo của chính mình.', 403);
    const code = 'zl_' + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
    const l = zdLoad_(); if (str_(b.chuNick, 80)) l.forEach(x => { if (x.chuNick === str_(b.chuNick, 80)) x.active = false; });
    const d = {id: uid_(), h: sha_(code), ten: str_(b.ten, 80) || 'Zalo', chuNick: str_(b.chuNick, 80), active: true, createdAt: now_(), by: ME.ten, duoi: code.slice(-4)};
    l.push(d); zdSave_(l); return out_({ok: true, code: code, dev: zDevOut_(d)});
  },
  zl_reset: b => {
    const nick = str_(b.nick, 40); if (!nick) return fail_('Thiếu nick.');
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try { let bo = 0; ['zmsg', 'zthr'].forEach(nm => { const sh = sheet_(nm), n = sh.getLastRow(); if (n < 2) return; const cols = COLS[nm], v = sh.getRange(2, 1, n - 1, cols.length).getValues(), keep = v.filter(r => String(r[1]) !== nick); bo += v.length - keep.length;
        sh.getRange(2, 1, n - 1, cols.length).clearContent(); if (keep.length) sh.getRange(2, 1, keep.length, cols.length).setNumberFormat('@').setValues(keep.map(r => r.map(String))); });
      return out_({ok: true, xoa: bo}); } finally { lock.releaseLock(); }
  },
  zl_dev_revoke: b => { const l = zdLoad_(), d = l.filter(x => x.id === str_(b.id, 32))[0]; if (!d) return fail_('Không thấy mã.'); d.active = false; zdSave_(l); return out_({ok: true}); },
  zl_perm_set: b => {
    const P = {}; Object.keys(b.perm || {}).slice(0, 60).forEach(u => { P[str_(u, 40)] = list_(b.perm[u], 20, 40); });
    const r = rows_('settings').filter(x => x.k === 'zaloPerm')[0] || {k: 'zaloPerm'}; r.v = JSON.stringify(P); put_('settings', r);
    return out_({ok: true, perm: P});
  },
  zl_threads: (b, ME) => {
    if (!zAny_(ME)) return fail_('Anh/chị chưa được cấp quyền xem Zalo.', 403);
    const accs = rows_('zacc').filter(a => zCan_(ME, a.id)), ok = {}; accs.forEach(a => ok[a.id] = 1);
    const th = rows_('zthr').filter(t => ok[t.nick]).sort((x, y) => Number(y.lastTs) - Number(x.lastTs)).slice(0, 600)
      .map(t => ({key: t.key, nick: t.nick, thread: t.thread, ten: t.ten, anh: t.anh, nhom: t.nhom === '1', lastTs: Number(t.lastTs) || 0, lastText: t.lastText}));
    return out_({ok: true, accs: accs.map(a => ({id: a.id, ten: a.ten, anh: a.anh, chuNick: a.chuNick, lastSync: a.lastSync})), threads: th});
  },
  zl_msgs: (b, ME) => {
    const nick = str_(b.nick, 40), thread = str_(b.thread, 60); if (!zCan_(ME, nick)) return fail_('Không có quyền xem nick này.', 403);
    const s = sheet_('zmsg'), n = s.getLastRow(); if (n < 2) return out_({ok: true, msgs: []});
    const pre = nick + '|' + thread + '|', cols = COLS.zmsg, lo = Math.max(2, n - 40000), out = [];
    s.getRange(lo, 1, n - lo + 1, cols.length).getValues().forEach(v => { if (String(v[0]).indexOf(pre) !== 0) return; const o = {}; cols.forEach((c, i) => o[c] = String(v[i])); out.push(o); });
    out.splice(0, Math.max(0, out.length - 400));
    out.sort((x, y) => Number(x.ts) - Number(y.ts));
    return out_({ok: true, msgs: out.map(m => ({id: m.msgId, ts: Number(m.ts) || 0, me: m.me === '1', nguoiGui: m.nguoiGui, loai: m.loai, noiDung: m.noiDung, kieu: m.kieu, url: m.url, tep: m.tep}))});
  },
};
function zlPush_(b) {
  const code = String(b.code || ''); if (!/^zl_[0-9a-f]{40}$/.test(code)) return fail_('Sai mã kết nối Zalo.', 401);
  const l = zdLoad_(), d = l.filter(x => x.h === sha_(code))[0]; if (!d || !d.active) return fail_('Mã kết nối Zalo không hợp lệ hoặc đã bị thu hồi.', 401);
  const acc = b.acc || {}, nick = str_(acc.uid, 40); if (!nick) return fail_('Thiếu thông tin nick Zalo.');
  if (d.nick && d.nick !== nick) return fail_('Mã này đã gắn với nick Zalo khác (' + (d.nickTen || d.nick) + '). Tạo mã mới cho nick này.', 409);
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    if (!d.nick) { d.nick = nick; d.nickTen = str_(acc.name, 80); }
    d.lastUsed = now_(); zdSave_(l);
    const A = rows_('zacc').filter(a => a.id === nick)[0] || {id: nick, createdAt: now_()};
    A.ten = str_(acc.name, 80) || A.ten; A.anh = str_(acc.avatar, 500) || A.anh; A.devId = d.id; A.chuNick = d.chuNick; A.lastSync = now_(); put_('zacc', A);
    // hội thoại
    const T = {}, tAll = rows_('zthr'); tAll.forEach(t => T[t.key] = t);
    (Array.isArray(b.threads) ? b.threads : []).slice(0, 300).forEach(t => {
      const th = str_(t.id, 60); if (!th) return; const k = nick + '|' + th; if (!T[k] && !t.lastTs) return; const o = T[k] || {key: k, nick: nick, thread: th};
      o.ten = str_(t.name, 120) || o.ten; o.anh = str_(t.avatar, 500) || o.anh; o.nhom = t.group ? '1' : (o.nhom || '0');
      if (t.lastTs && Number(t.lastTs) >= (Number(o.lastTs) || 0)) { o.lastTs = String(Number(t.lastTs)); o.lastText = str_(t.lastText, 200); }
      o.updatedAt = now_(); o._dirty = 1; T[k] = o;
    });
    // tin nhắn (bỏ trùng theo key)
    const s = sheet_('zmsg'), n = s.getLastRow(), have = {};
    const msgs = (Array.isArray(b.msgs) ? b.msgs : []).slice(0, 500);
    if (msgs.length && n > 1) { const lo = Math.max(2, n - 20000); s.getRange(lo, 1, n - lo + 1, 1).getValues().forEach(v => have[v[0]] = 1); }
    const add = [];
    msgs.forEach(m => {
      const th = str_(m.thread, 60), id = str_(m.id, 80); if (!th || !id) return; const k = nick + '|' + th + '|' + id; if (have[k]) return; have[k] = 1;
      const ts = Number(m.ts) || Date.now();
      add.push([k, nick, th, id, String(ts), m.me ? '1' : '0', str_(m.fromName, 120), str_(m.type, 20) || 'text', str_(m.text, 5000), str_(typeof m.styles === 'string' ? m.styles : JSON.stringify(m.styles || ''), 2000), str_(m.url, 1000), str_(m.file, 200), now_()]);
      const tk = nick + '|' + th, o = T[tk] || (T[tk] = {key: tk, nick: nick, thread: th, ten: str_(m.threadName, 120), nhom: '0'});
      if (ts >= (Number(o.lastTs) || 0)) { o.lastTs = String(ts); o.lastText = (m.me ? 'Bạn: ' : '') + (str_(m.text, 150) || '[' + ({photo: 'Hình ảnh', file: 'Tệp', sticker: 'Sticker', voice: 'Tin nhắn thoại', video: 'Video', link: 'Liên kết'}[str_(m.type, 20)] || 'Tệp') + ']'); o._dirty = 1; o.updatedAt = now_(); }
    });
    if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, COLS.zmsg.length).setNumberFormat('@').setValues(add);
    if (Object.keys(T).some(k => T[k]._dirty)) { const sh = sheet_('zthr'), cols = COLS.zthr, all = Object.keys(T).map(k => cols.map(c => T[k][c] === undefined ? '' : String(T[k][c])));
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).clearContent();
      if (all.length) sh.getRange(2, 1, all.length, cols.length).setNumberFormat('@').setValues(all); }
    return out_({ok: true, them: add.length, nick: nick});
  } finally { lock.releaseLock(); }
}
