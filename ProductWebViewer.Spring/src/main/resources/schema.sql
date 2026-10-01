DROP TABLE IF EXISTS audit_logs;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS substrates;

CREATE TABLE products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_name TEXT,
    product_name TEXT,
    product_model TEXT,
    product_type TEXT,
    order_number TEXT,
    product_number TEXT,
    oles_number TEXT,
    quantity INTEGER,
    reg_date TEXT,
    comment TEXT,
    created_at TEXT,
    is_deleted INTEGER NOT NULL DEFAULT 0,
    deleted_at TEXT
);

CREATE TABLE substrates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_name TEXT,
    product_name TEXT,
    substrate_name TEXT,
    substrate_model TEXT,
    order_number TEXT,
    substrate_number TEXT,
    increase INTEGER,
    decrease INTEGER,
    defect INTEGER,
    person_info TEXT,
    reg_date TEXT,
    comment TEXT,
    created_at TEXT,
    is_deleted INTEGER NOT NULL DEFAULT 0,
    deleted_at TEXT
);

-- target_type ('PRODUCT' / 'SUBSTRATE') + target_id で、製品・基板どちらの操作ログも同じテーブルにまとめる
CREATE TABLE audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action TEXT NOT NULL,
    target_type TEXT NOT NULL,
    target_id INTEGER NOT NULL,
    detail TEXT,
    created_at TEXT NOT NULL
);
