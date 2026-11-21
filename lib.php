<?php
declare(strict_types=1);

const DATA_DIR = __DIR__ . '/data';
const UPLOAD_DIR = DATA_DIR . '/uploads';

function cfg(?string $k = null) {
  static $c = null;
  if ($c === null) {
    $f = __DIR__ . '/config.php';
    if (!is_file($f)) { http_response_code(500); exit('Thiếu file config.php (xem config.sample.php).'); }
    $c = require $f;
  }
  return $k === null ? $c : ($c[$k] ?? null);
}

function db(): PDO {
  static $pdo = null;
  if ($pdo) return $pdo;
  if (!is_dir(UPLOAD_DIR)) @mkdir(UPLOAD_DIR, 0750, true);
  $pdo = new PDO(cfg('db_dsn'), cfg('db_user'), cfg('db_pass'), [
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
  ]);
  $mysql = $pdo->getAttribute(PDO::ATTR_DRIVER_NAME) === 'mysql';
  $big = $mysql ? 'MEDIUMTEXT' : 'TEXT';
  $eng = $mysql ? ' ENGINE=InnoDB DEFAULT CHARSET=utf8mb4' : '';
  $pdo->exec("CREATE TABLE IF NOT EXISTS forms (id VARCHAR(32) PRIMARY KEY, slug VARCHAR(80) NOT NULL UNIQUE, ten VARCHAR(255) NOT NULL, khoa VARCHAR(80) NOT NULL, nam VARCHAR(10) NOT NULL, mau VARCHAR(20) NOT NULL, mo_ta TEXT, mo INT NOT NULL DEFAULT 1, created_at VARCHAR(30) NOT NULL)$eng");
  $pdo->exec("CREATE TABLE IF NOT EXISTS records (id VARCHAR(32) PRIMARY KEY, form_id VARCHAR(32) NOT NULL, sdt VARCHAR(20) NOT NULL DEFAULT '', data $big NOT NULL, photo VARCHAR(64) NOT NULL DEFAULT '', created_at VARCHAR(30) NOT NULL, updated_at VARCHAR(30) NOT NULL)$eng");
  $pdo->exec("CREATE TABLE IF NOT EXISTS settings (k VARCHAR(64) PRIMARY KEY, v $big NOT NULL)$eng");
  $pdo->exec("CREATE TABLE IF NOT EXISTS hits (ip VARCHAR(64) NOT NULL, kind VARCHAR(16) NOT NULL, t INT NOT NULL)$eng");
  if ((int)$pdo->query('SELECT COUNT(*) FROM forms')->fetchColumn() === 0) seed($pdo);
  return $pdo;
}

function seed(PDO $pdo): void {
  $now = date('c');
  $ins = $pdo->prepare('INSERT INTO forms (id,slug,ten,khoa,nam,mau,mo_ta,mo,created_at) VALUES (?,?,?,?,?,?,?,1,?)');
  $ins->execute([uid(), 'cma-diamond-2026', 'Phiếu thông tin học viên CMA Diamond', 'Diamond', '2026', 'hocvien', 'Phiếu chính cho học viên CMA Diamond 2026.', $now]);
  $ins->execute([uid(), 'cma-trainer-2026', 'Phiếu thông tin học viên CMA Trainer', 'Trainer', '2026', 'hocvien', 'Phiếu cho học viên CMA Trainer 2026.', $now]);
}

function uid(): string { return bin2hex(random_bytes(8)); }
function now(): string { return date('c'); }
function ip(): string { return substr((string)($_SERVER['REMOTE_ADDR'] ?? ''), 0, 64); }
function phone_norm(string $s): string { return substr(preg_replace('/\D+/', '', $s) ?? '', 0, 20); }

function json_out($data, int $code = 200): void {
  http_response_code($code);
  header('Content-Type: application/json; charset=utf-8');
  header('Cache-Control: no-store');
  echo json_encode($data, JSON_UNESCAPED_UNICODE);
  exit;
}
function fail(string $msg, int $code = 400): void { json_out(['ok' => false, 'error' => $msg], $code); }

function body(): array {
  $raw = file_get_contents('php://input') ?: '';
  if (strlen($raw) > 4_000_000) fail('Dữ liệu quá lớn.', 413);
  $j = json_decode($raw, true);
  return is_array($j) ? $j : [];
}

/** Đếm số lần trong khoảng thời gian (chống spam / dò mật khẩu). */
function too_many(string $kind, int $max, int $windowSec): bool {
  $pdo = db();
  $since = time() - $windowSec;
  $pdo->prepare('DELETE FROM hits WHERE t < ?')->execute([time() - 86400]);
  $st = $pdo->prepare('SELECT COUNT(*) FROM hits WHERE ip=? AND kind=? AND t>=?');
  $st->execute([ip(), $kind, $since]);
  return (int)$st->fetchColumn() >= $max;
}
function hit(string $kind): void { db()->prepare('INSERT INTO hits (ip,kind,t) VALUES (?,?,?)')->execute([ip(), $kind, time()]); }

function start_session(): void {
  if (session_status() === PHP_SESSION_ACTIVE) return;
  $https = (($_SERVER['HTTPS'] ?? '') !== '' && $_SERVER['HTTPS'] !== 'off') || (($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? '') === 'https');
  session_name('crmtkt');
  session_set_cookie_params(['lifetime' => 0, 'path' => '/', 'httponly' => true, 'secure' => $https, 'samesite' => 'Strict']);
  session_start();
}
function is_admin(): bool { start_session(); return !empty($_SESSION['admin']) && ($_SESSION['exp'] ?? 0) > time(); }
function require_admin(): void {
  if (!is_admin()) fail('Chưa đăng nhập.', 401);
  // Chặn gọi chéo trang: trang quản trị luôn gửi header này.
  if (($_SERVER['HTTP_X_CRM'] ?? '') !== '1') fail('Yêu cầu không hợp lệ.', 403);
  $_SESSION['exp'] = time() + 8 * 3600;
}

function form_by_slug(string $slug): ?array {
  $st = db()->prepare('SELECT * FROM forms WHERE slug=?');
  $st->execute([$slug]);
  return $st->fetch() ?: null;
}
function course_name(string $k): string {
  return ['Diamond' => 'CMA Diamond', 'Trainer' => 'CMA Trainer'][$k] ?? $k;
}
function base_path(): string {
  $d = rtrim(str_replace('\\', '/', dirname($_SERVER['SCRIPT_NAME'] ?? '/')), '/');
  return $d === '' ? '/' : $d . '/';
}
function h($s): string { return htmlspecialchars((string)$s, ENT_QUOTES, 'UTF-8'); }

function security_headers(): void {
  header('X-Content-Type-Options: nosniff');
  header('Referrer-Policy: same-origin');
  header('X-Frame-Options: SAMEORIGIN');
}
