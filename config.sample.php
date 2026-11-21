<?php
// Cấu hình CRM CMA | SKT. File này KHÔNG đưa lên GitHub (có trong .gitignore).
// Đổi mật khẩu: chạy  php -r 'echo password_hash("MAT_KHAU_MOI", PASSWORD_BCRYPT);'
// rồi dán kết quả vào admin_hash.
return [
  'admin_user' => 'TEN_DANG_NHAP',
  'admin_hash' => 'DAN_MA_HASH_MAT_KHAU_VAO_DAY',

  // Mặc định lưu bằng SQLite trong thư mục data/ (không cần tạo database).
  // Muốn dùng MySQL của hosting Mắt Bão: đổi dsn thành
  //   'mysql:host=localhost;dbname=TEN_DATABASE;charset=utf8mb4'
  // và điền user / pass.
  'db_dsn'  => 'sqlite:' . __DIR__ . '/data/crm.sqlite',
  'db_user' => null,
  'db_pass' => null,

  'site_url' => 'https://crmtkt.com',
  'staff_default' => ['Kim Anh', 'Mỹ Linh', 'Oanh'],
];
