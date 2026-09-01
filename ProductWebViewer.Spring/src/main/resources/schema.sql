DROP TABLE IF EXISTS audit_logs;
DROP TABLE IF EXISTS products;

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

CREATE TABLE audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action TEXT NOT NULL,
    product_id INTEGER NOT NULL,
    detail TEXT,
    created_at TEXT NOT NULL
);
