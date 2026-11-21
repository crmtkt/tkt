<?php
declare(strict_types=1);
require __DIR__ . '/lib.php';
security_headers();

const ANSWER_KEYS = ['hoTen','namSinh','sdt','email','thanhPho','congViec','vaiTro','khoaDaHoc','lyDo','vanDe','mongMuon','ketQua','ngaiNgai','sanPhamRieng','thoiGian','aiLevel','aiTools','aiDaLam','aiMuon','dongY'];
const CRM_KEYS = ['trangThai','phuTrach','khoa','nam','nhom','zaloGroup','ghiChu','hanhTrinh','tienDo','nhatKy'];
const REQUIRED = [
  'hocvien' => ['hoTen','namSinh','sdt','email','congViec','khoaDaHoc','lyDo','vanDe','mongMuon','aiLevel','dongY'],
  'nhanh'   => ['hoTen','sdt','dongY'],
];
const STATUSES = ['Mới đăng ký','Đã liên hệ','Đã vào nhóm Zalo','Đang học','Cần chăm sóc gấp','Tạm dừng','Hoàn thành'];

function clean_str($v, int $max = 3000): string {
  $s = is_scalar($v) ? trim((string)$v) : '';
  return mb_substr($s, 0, $max);
}
function clean_list($v, int $n = 30, int $max = 200): array {
  if (!is_array($v)) return [];
  $out = [];
  foreach (array_slice($v, 0, $n) as $x) { $s = clean_str($x, $max); if ($s !== '') $out[] = $s; }
  return array_values(array_unique($out));
}
function clean_answers(array $in): array {
  $o = [];
  foreach (ANSWER_KEYS as $k) {
    if (!array_key_exists($k, $in)) continue;
    if ($k === 'aiTools') $o[$k] = clean_list($in[$k]);
    elseif ($k === 'aiLevel') $o[$k] = max(0, min(5, (int)$in[$k]));
    elseif ($k === 'dongY') $o[$k] = (bool)$in[$k];
    else $o[$k] = clean_str($in[$k], in_array($k, ['hoTen','sdt','email','namSinh','thanhPho','congViec','vaiTro','thoiGian'], true) ? 200 : 4000);
  }
  return $o;
}
function clean_entries($v, array $keys): array {
  if (!is_array($v)) return [];
  $out = [];
  foreach (array_slice($v, -300) as $e) {
    if (!is_array($e)) continue;
    $r = [];
    foreach ($keys as $k) $r[$k] = clean_str($e[$k] ?? '', 2000);
    if (($r['noiDung'] ?? '') !== '') $out[] = $r;
  }
  return $out;
}
function clean_crm(array $in): array {
  $o = [];
  foreach (CRM_KEYS as $k) {
    if (!array_key_exists($k, $in)) continue;
    $v = $in[$k];
    switch ($k) {
      case 'trangThai': $o[$k] = in_array($v, STATUSES, true) ? $v : 'Mới đăng ký'; break;
      case 'phuTrach': $o[$k] = clean_list($v, 20, 80); break;
      case 'zaloGroup': $s = clean_str($v, 500); $o[$k] = preg_match('~^https://[^\s"\'<>]+$~', $s) ? $s : ''; break;
      case 'hanhTrinh': $o[$k] = clean_entries($v, ['ngay','noiDung']); break;
      case 'nhatKy': $o[$k] = clean_entries($v, ['ngay','nguoi','noiDung']); break;
      case 'tienDo':
        $t = [];
        if (is_array($v)) foreach ($v as $m => $txt) if (preg_match('/^\d{4}-\d{2}$/', (string)$m)) { $s = clean_str($txt, 2000); if ($s !== '') $t[$m] = $s; }
        $o[$k] = (object)$t; break;
      default: $o[$k] = clean_str($v, 2000);
    }
  }
  return $o;
}
function validate_answers(array $a, string $mau, bool $hasPhoto): ?string {
  foreach (REQUIRED[$mau] ?? REQUIRED['nhanh'] as $k) {
    $v = $a[$k] ?? null;
    $names = ['hoTen'=>'họ tên','namSinh'=>'năm sinh','sdt'=>'số điện thoại','email'=>'email','congViec'=>'công việc','khoaDaHoc'=>'các khóa đã và đang học','lyDo'=>'lý do đăng ký','vanDe'=>'vấn đề cần hỗ trợ','mongMuon'=>'mong muốn','aiLevel'=>'mức độ AI','dongY'=>'ô đồng ý'];
    if ($v === null || $v === '' || $v === false || $v === 0 || $v === []) return 'Vui lòng điền ' . ($names[$k] ?? $k) . '.';
  }
  if ($mau === 'hocvien' && !$hasPhoto) return 'Thiếu ảnh chân dung.';
  if (!preg_match('/^\+?\d{9,12}$/', preg_replace('/[\s.\-]/', '', $a['sdt'] ?? ''))) return 'Số điện thoại chưa đúng.';
  if (($a['email'] ?? '') !== '' && !filter_var($a['email'], FILTER_VALIDATE_EMAIL)) return 'Email chưa đúng.';
  if (($a['namSinh'] ?? '') !== '' && !preg_match('/^(19|20)\d\d$/', $a['namSinh'])) return 'Năm sinh chưa đúng.';
  return null;
}
/** Lưu ảnh: chỉ nhận JPEG/PNG/WebP, vẽ lại qua GD để loại bỏ mọi thứ không phải ảnh. */
function save_photo(string $dataUrl, string $id): string {
  if (!preg_match('~^data:image/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$~', $dataUrl, $m)) return '';
  $bin = base64_decode($m[2], true);
  if ($bin === false || strlen($bin) > 3_000_000) return '';
  $info = @getimagesizefromstring($bin);
  if (!$info || !in_array($info[2], [IMAGETYPE_JPEG, IMAGETYPE_PNG, IMAGETYPE_WEBP], true)) return '';
  $file = $id . '.jpg';
  if (function_exists('imagecreatefromstring')) {
    $im = @imagecreatefromstring($bin);
    if (!$im) return '';
    $w = imagesx($im); $h = imagesy($im); $k = min(1, 600 / max($w, $h));
    $out = imagecreatetruecolor(max(1, (int)round($w * $k)), max(1, (int)round($h * $k)));
    imagefill($out, 0, 0, imagecolorallocate($out, 255, 255, 255));
    imagecopyresampled($out, $im, 0, 0, 0, 0, imagesx($out), imagesy($out), $w, $h);
    imagejpeg($out, UPLOAD_DIR . '/' . $file, 84);
  } else {
    if ($info[2] !== IMAGETYPE_JPEG) return '';
    file_put_contents(UPLOAD_DIR . '/' . $file, $bin);
  }
  return $file;
}
function record_out(array $r): array {
  $d = json_decode($r['data'], true) ?: [];
  return $d + ['id' => $r['id'], 'formId' => $r['form_id'], 'createdAt' => substr($r['created_at'], 0, 10), 'updatedAt' => $r['updated_at'], 'hasPhoto' => $r['photo'] !== ''];
}
function form_out(array $f, int $count = 0): array {
  return ['id' => $f['id'], 'slug' => $f['slug'], 'ten' => $f['ten'], 'khoa' => $f['khoa'], 'nam' => $f['nam'], 'mau' => $f['mau'], 'moTa' => (string)$f['mo_ta'], 'mo' => (bool)$f['mo'], 'createdAt' => substr($f['created_at'], 0, 10), 'soPhieu' => $count];
}
function staff_list(): array {
  $st = db()->prepare('SELECT v FROM settings WHERE k=?'); $st->execute(['staff']);
  $v = $st->fetchColumn();
  $l = $v ? json_decode($v, true) : null;
  return is_array($l) && $l ? $l : cfg('staff_default');
}
function slugify(string $s): string {
  $s = mb_strtolower(trim($s));
  $map = ['à'=>'a','á'=>'a','ả'=>'a','ã'=>'a','ạ'=>'a','ă'=>'a','ằ'=>'a','ắ'=>'a','ẳ'=>'a','ẵ'=>'a','ặ'=>'a','â'=>'a','ầ'=>'a','ấ'=>'a','ẩ'=>'a','ẫ'=>'a','ậ'=>'a','đ'=>'d','è'=>'e','é'=>'e','ẻ'=>'e','ẽ'=>'e','ẹ'=>'e','ê'=>'e','ề'=>'e','ế'=>'e','ể'=>'e','ễ'=>'e','ệ'=>'e','ì'=>'i','í'=>'i','ỉ'=>'i','ĩ'=>'i','ị'=>'i','ò'=>'o','ó'=>'o','ỏ'=>'o','õ'=>'o','ọ'=>'o','ô'=>'o','ồ'=>'o','ố'=>'o','ổ'=>'o','ỗ'=>'o','ộ'=>'o','ơ'=>'o','ờ'=>'o','ớ'=>'o','ở'=>'o','ỡ'=>'o','ợ'=>'o','ù'=>'u','ú'=>'u','ủ'=>'u','ũ'=>'u','ụ'=>'u','ư'=>'u','ừ'=>'u','ứ'=>'u','ử'=>'u','ữ'=>'u','ự'=>'u','ỳ'=>'y','ý'=>'y','ỷ'=>'y','ỹ'=>'y','ỵ'=>'y'];
  $s = strtr($s, $map);
  $s = preg_replace('/[^a-z0-9]+/', '-', $s) ?? '';
  return trim(substr($s, 0, 70), '-');
}

$a = (string)($_GET['a'] ?? '');
try {
  switch ($a) {

  /* ---------- công khai ---------- */
  case 'form': {
    $f = form_by_slug((string)($_GET['s'] ?? ''));
    if (!$f) fail('Không tìm thấy form.', 404);
    json_out(['ok' => true, 'form' => ['ten' => $f['ten'], 'khoa' => $f['khoa'], 'khoaTen' => course_name($f['khoa']), 'mau' => $f['mau'], 'mo' => (bool)$f['mo']]]);
  }
  case 'submit': {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') fail('Sai phương thức.', 405);
    if (too_many('submit', 8, 600)) fail('Anh/chị gửi quá nhiều lần, vui lòng thử lại sau ít phút.', 429);
    $b = body();
    $f = form_by_slug(clean_str($b['s'] ?? '', 80));
    if (!$f) fail('Không tìm thấy form.', 404);
    if (!$f['mo']) fail('Form này đã đóng, vui lòng liên hệ Ban tổ chức.', 403);
    if (clean_str($b['web'] ?? '') !== '') json_out(['ok' => true]); // bẫy bot
    $ans = clean_answers(is_array($b['data'] ?? null) ? $b['data'] : []);
    $photoIn = is_string($b['anh'] ?? null) ? $b['anh'] : '';
    if ($e = validate_answers($ans, $f['mau'], $photoIn !== '')) fail($e);
    hit('submit');
    $pdo = db();
    $sdt = phone_norm($ans['sdt']);
    $st = $pdo->prepare('SELECT * FROM records WHERE form_id=? AND sdt=?');
    $st->execute([$f['id'], $sdt]);
    $old = $st->fetch();
    $id = $old['id'] ?? uid();
    $photo = $photoIn !== '' ? save_photo($photoIn, $id) : '';
    if ($photoIn !== '' && $photo === '') fail('Ảnh chưa hợp lệ, vui lòng chọn ảnh JPG hoặc PNG khác.');
    $ans['guiLuc'] = now();
    if ($old) {
      // Học viên gửi lại: cập nhật câu trả lời, giữ nguyên phần chăm sóc của CRM.
      $d = json_decode($old['data'], true) ?: [];
      $d = array_merge($d, $ans);
      $d['hanhTrinh'][] = ['ngay' => date('Y-m-d'), 'noiDung' => 'Cập nhật lại phiếu: ' . $f['ten']];
      $pdo->prepare('UPDATE records SET data=?, photo=?, updated_at=? WHERE id=?')
          ->execute([json_encode($d, JSON_UNESCAPED_UNICODE), $photo ?: $old['photo'], now(), $id]);
    } else {
      $d = $ans + ['khoa' => $f['khoa'], 'nam' => $f['nam'], 'nhom' => '', 'trangThai' => 'Mới đăng ký', 'phuTrach' => [],
        'hanhTrinh' => [['ngay' => date('Y-m-d'), 'noiDung' => 'Điền ' . $f['ten']]], 'nhatKy' => [], 'tienDo' => (object)[]];
      $pdo->prepare('INSERT INTO records (id,form_id,sdt,data,photo,created_at,updated_at) VALUES (?,?,?,?,?,?,?)')
          ->execute([$id, $f['id'], $sdt, json_encode($d, JSON_UNESCAPED_UNICODE), $photo, now(), now()]);
    }
    json_out(['ok' => true]);
  }

  /* ---------- đăng nhập ---------- */
  case 'login': {
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') fail('Sai phương thức.', 405);
    if (too_many('login', 5, 900)) fail('Sai quá 5 lần. Vui lòng thử lại sau 15 phút.', 429);
    $b = body();
    $u = clean_str($b['u'] ?? '', 100); $p = (string)($b['p'] ?? '');
    $ok = hash_equals((string)cfg('admin_user'), $u) & password_verify($p, (string)cfg('admin_hash'));
    if (!$ok) { hit('login'); fail('Sai tài khoản hoặc mật khẩu.', 401); }
    start_session(); session_regenerate_id(true);
    $_SESSION['admin'] = $u; $_SESSION['exp'] = time() + 8 * 3600; $_SESSION['tok'] = bin2hex(random_bytes(16));
    json_out(['ok' => true, 'tok' => $_SESSION['tok']]);
  }
  case 'logout': { start_session(); $_SESSION = []; session_destroy(); json_out(['ok' => true]); }
  case 'me': { if (!is_admin()) fail('Chưa đăng nhập.', 401); json_out(['ok' => true, 'user' => $_SESSION['admin'], 'tok' => $_SESSION['tok'], 'site' => cfg('site_url')]); }

  /* ---------- quản trị ---------- */
  case 'bootstrap': {
    require_admin();
    $pdo = db();
    $counts = [];
    foreach ($pdo->query('SELECT form_id, COUNT(*) c FROM records GROUP BY form_id') as $r) $counts[$r['form_id']] = (int)$r['c'];
    $forms = array_map(fn($f) => form_out($f, $counts[$f['id']] ?? 0), $pdo->query('SELECT * FROM forms ORDER BY created_at DESC')->fetchAll());
    $recs = array_map('record_out', $pdo->query('SELECT * FROM records ORDER BY created_at DESC')->fetchAll());
    json_out(['ok' => true, 'forms' => $forms, 'records' => $recs, 'staff' => staff_list()]);
  }
  case 'form_save': {
    require_admin(); $b = body(); $pdo = db();
    $ten = clean_str($b['ten'] ?? '', 200); if ($ten === '') fail('Vui lòng đặt tên form.');
    $khoa = clean_str($b['khoa'] ?? '', 80) ?: 'Diamond';
    $nam = preg_match('/^\d{4}$/', (string)($b['nam'] ?? '')) ? (string)$b['nam'] : date('Y');
    $mau = in_array($b['mau'] ?? '', ['hocvien', 'nhanh'], true) ? $b['mau'] : 'hocvien';
    $slug = slugify(clean_str($b['slug'] ?? '', 80) ?: $ten); if ($slug === '') fail('Đường dẫn chưa hợp lệ.');
    $moTa = clean_str($b['moTa'] ?? '', 1000); $mo = !empty($b['mo']) ? 1 : 0;
    $id = clean_str($b['id'] ?? '', 32);
    $dup = $pdo->prepare('SELECT id FROM forms WHERE slug=?'); $dup->execute([$slug]); $d = $dup->fetchColumn();
    if ($d && $d !== $id) fail('Đường dẫn này đã có form khác dùng, vui lòng đổi tên đường dẫn.');
    if ($id) $pdo->prepare('UPDATE forms SET slug=?,ten=?,khoa=?,nam=?,mau=?,mo_ta=?,mo=? WHERE id=?')->execute([$slug, $ten, $khoa, $nam, $mau, $moTa, $mo, $id]);
    else $pdo->prepare('INSERT INTO forms (id,slug,ten,khoa,nam,mau,mo_ta,mo,created_at) VALUES (?,?,?,?,?,?,?,?,?)')->execute([uid(), $slug, $ten, $khoa, $nam, $mau, $moTa, $mo, now()]);
    json_out(['ok' => true]);
  }
  case 'form_delete': {
    require_admin(); $b = body(); $pdo = db(); $id = clean_str($b['id'] ?? '', 32);
    $c = $pdo->prepare('SELECT COUNT(*) FROM records WHERE form_id=?'); $c->execute([$id]);
    if ((int)$c->fetchColumn() > 0) fail('Form đã có phiếu gửi về nên không xóa được. Anh/chị có thể bấm "Đóng form".');
    $pdo->prepare('DELETE FROM forms WHERE id=?')->execute([$id]);
    json_out(['ok' => true]);
  }
  case 'record_save': {
    require_admin(); $b = body(); $pdo = db(); $id = clean_str($b['id'] ?? '', 32);
    $st = $pdo->prepare('SELECT * FROM records WHERE id=?'); $st->execute([$id]); $r = $st->fetch();
    if (!$r) fail('Không tìm thấy hồ sơ.', 404);
    $d = array_merge(json_decode($r['data'], true) ?: [], clean_crm(is_array($b['patch'] ?? null) ? $b['patch'] : []));
    $pdo->prepare('UPDATE records SET data=?, updated_at=? WHERE id=?')->execute([json_encode($d, JSON_UNESCAPED_UNICODE), now(), $id]);
    json_out(['ok' => true, 'record' => record_out(['data' => json_encode($d)] + $r)]);
  }
  case 'record_delete': {
    require_admin(); $b = body(); $pdo = db(); $id = clean_str($b['id'] ?? '', 32);
    $st = $pdo->prepare('SELECT photo FROM records WHERE id=?'); $st->execute([$id]); $p = $st->fetchColumn();
    if ($p && preg_match('/^[a-f0-9]{16}\.jpg$/', $p)) @unlink(UPLOAD_DIR . '/' . $p);
    $pdo->prepare('DELETE FROM records WHERE id=?')->execute([$id]);
    json_out(['ok' => true]);
  }
  case 'staff_save': {
    require_admin(); $b = body();
    $l = clean_list($b['list'] ?? [], 50, 80); if (!$l) fail('Danh sách nhân sự không được trống.');
    $pdo = db(); $pdo->prepare('DELETE FROM settings WHERE k=?')->execute(['staff']);
    $pdo->prepare('INSERT INTO settings (k,v) VALUES (?,?)')->execute(['staff', json_encode($l, JSON_UNESCAPED_UNICODE)]);
    json_out(['ok' => true]);
  }
  case 'photo': {
    if (!is_admin()) { http_response_code(401); exit; }
    $st = db()->prepare('SELECT photo FROM records WHERE id=?'); $st->execute([clean_str($_GET['id'] ?? '', 32)]);
    $p = (string)$st->fetchColumn();
    if (!preg_match('/^[a-f0-9]{16}\.jpg$/', $p) || !is_file(UPLOAD_DIR . '/' . $p)) { http_response_code(404); exit; }
    header('Content-Type: image/jpeg'); header('Cache-Control: private, max-age=86400');
    readfile(UPLOAD_DIR . '/' . $p); exit;
  }
  case 'export': {
    if (!is_admin() || !hash_equals((string)($_SESSION['tok'] ?? 'x'), (string)($_GET['t'] ?? ''))) { http_response_code(401); exit('Chưa đăng nhập.'); }
    $pdo = db(); $fid = clean_str($_GET['form'] ?? '', 32);
    $f = $pdo->prepare('SELECT * FROM forms WHERE id=?'); $f->execute([$fid]); $form = $f->fetch();
    if (!$form) { http_response_code(404); exit; }
    $st = $pdo->prepare('SELECT * FROM records WHERE form_id=? ORDER BY created_at'); $st->execute([$fid]);
    $cols = ['Ngày gửi'=>'createdAt','Họ tên'=>'hoTen','Năm sinh'=>'namSinh','SĐT'=>'sdt','Email'=>'email','Tỉnh/thành'=>'thanhPho','Công việc'=>'congViec','Vai trò'=>'vaiTro','Khóa đã và đang học'=>'khoaDaHoc','Lý do đăng ký'=>'lyDo','Cần thầy hỗ trợ'=>'vanDe','Mong muốn'=>'mongMuon','Kết quả mong đợi'=>'ketQua','Rào cản'=>'ngaiNgai','Muốn đóng gói'=>'sanPhamRieng','Thời gian học/tuần'=>'thoiGian','Cấp độ AI'=>'aiLevel','Công cụ AI'=>'aiTools','Đã dùng AI để'=>'aiDaLam','Muốn AI giúp'=>'aiMuon','Trạng thái'=>'trangThai','Người phụ trách'=>'phuTrach','Nhóm'=>'nhom','Link Zalo'=>'zaloGroup','Ghi chú'=>'ghiChu'];
    header('Content-Type: text/csv; charset=utf-8');
    header('Content-Disposition: attachment; filename="' . $form['slug'] . '-' . date('Ymd') . '.csv"');
    $out = fopen('php://output', 'w'); fwrite($out, "\xEF\xBB\xBF");
    fputcsv($out, array_keys($cols));
    foreach ($st as $r) {
      $d = record_out($r);
      fputcsv($out, array_map(function ($k) use ($d) {
        $v = $d[$k] ?? ''; $v = is_array($v) ? implode(', ', $v) : (string)$v;
        return preg_match('/^[=+\-@]/', $v) ? "'" . $v : $v; // chặn công thức Excel
      }, array_values($cols)));
    }
    exit;
  }
  default: fail('Không rõ yêu cầu.', 404);
  }
} catch (Throwable $e) {
  error_log('[crmtkt] ' . $e->getMessage());
  fail('Lỗi máy chủ, vui lòng thử lại.', 500);
}
