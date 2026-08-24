-- 練習用データベースのスキーマ定義。
-- 本番の ProductRegistry.db とは別物（ダミーデータ用）。
-- テーブル構成は ProductWebViewer が依存する範囲（製品/基板マスター、登録実績、シリアル、ビュー）に限定して再現している。

CREATE TABLE M_Person (
    PersonID   INTEGER PRIMARY KEY AUTOINCREMENT,
    PersonName TEXT NOT NULL,
    IsActive   INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE M_ProductDef (
    ProductID     INTEGER PRIMARY KEY AUTOINCREMENT,
    CategoryName  TEXT NOT NULL,
    ProItemNumber TEXT,
    ProductName   TEXT NOT NULL,
    ProductType   TEXT,
    ProductModel  TEXT,
    Initial       TEXT,
    RevisionGroup INTEGER,
    RegType       INTEGER NOT NULL DEFAULT 0,
    SerialType    INTEGER,
    Checkbox      TEXT,
    SerialPrintType INTEGER NOT NULL DEFAULT 0,
    SheetPrintType  INTEGER NOT NULL DEFAULT 0,
    Visible       INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE M_SubstrateDef (
    SubstrateID   INTEGER PRIMARY KEY AUTOINCREMENT,
    CategoryName  TEXT NOT NULL,
    SubItemNumber TEXT,
    ProductName   TEXT NOT NULL,
    SubstrateName TEXT NOT NULL,
    SubstrateModel TEXT,
    RegType       INTEGER NOT NULL DEFAULT 0,
    Checkbox      TEXT,
    SerialPrintType INTEGER NOT NULL DEFAULT 0,
    Visible       INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE T_Product (
    ID              INTEGER PRIMARY KEY AUTOINCREMENT,
    ProductID       INTEGER NOT NULL REFERENCES M_ProductDef(ProductID),
    PersonID        INTEGER REFERENCES M_Person(PersonID),
    OrderNumber     TEXT,
    ProductNumber   TEXT,
    OLesNumber      TEXT,
    Quantity        INTEGER,
    RegDate         TEXT NOT NULL,
    Revision        TEXT,
    RevisionGroup   INTEGER,
    SerialFirst     TEXT,
    SerialLast      TEXT,
    SerialLastNumber INTEGER,
    Comment         TEXT,
    CreatedAt       TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    IsDeleted       INTEGER NOT NULL DEFAULT 0,
    DeletedAt       TEXT
);

CREATE TABLE T_Substrate (
    ID              INTEGER PRIMARY KEY AUTOINCREMENT,
    SubstrateID     INTEGER NOT NULL REFERENCES M_SubstrateDef(SubstrateID),
    PersonID        INTEGER REFERENCES M_Person(PersonID),
    UseID           INTEGER REFERENCES T_Product(ID),
    SubstrateNumber TEXT,
    OrderNumber     TEXT,
    Increase        INTEGER,
    Decrease        INTEGER,
    Defect          INTEGER,
    RegDate         TEXT NOT NULL,
    Comment         TEXT,
    CreatedAt       TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    IsDeleted       INTEGER NOT NULL DEFAULT 0,
    DeletedAt       TEXT
);

CREATE TABLE T_Serial (
    ProductID TEXT NOT NULL REFERENCES M_ProductDef(ProductID),
    UsedID    INTEGER REFERENCES T_Product(ID),
    Serial    TEXT NOT NULL,
    OLesSerial TEXT
);

-- ProductWebViewer/Data/ViewDefinitionVerifier.cs の定義と一致させること
CREATE VIEW V_Product AS
SELECT
    t.ID,
    t.ProductID,
    m.CategoryName,
    m.ProductName,
    m.ProductType,
    m.ProductModel,
    t.OrderNumber,
    t.ProductNumber,
    t.OLesNumber,
    t.Quantity,
    t.SerialFirst,
    t.SerialLast,
    t.Revision,
    t.RevisionGroup,
    t.SerialLastNumber,
    COALESCE(p.PersonName, '') AS PersonInfo,
    t.RegDate,
    t.Comment,
    t.CreatedAt,
    t.IsDeleted,
    t.DeletedAt
FROM T_Product AS t
LEFT JOIN M_ProductDef AS m ON t.ProductID = m.ProductID
LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
WHERE t.IsDeleted = 0;

CREATE VIEW V_Substrate AS
SELECT
    t.ID,
    t.SubstrateID,
    m.CategoryName,
    m.ProductName,
    m.SubstrateName,
    m.SubstrateModel,
    t.OrderNumber,
    t.SubstrateNumber,
    t.Increase,
    t.Decrease,
    t.Defect,
    COALESCE(p.PersonName, '') AS PersonInfo,
    t.RegDate,
    t.Comment,
    t.UseID,
    t.CreatedAt,
    t.IsDeleted,
    t.DeletedAt
FROM T_Substrate AS t
LEFT JOIN M_SubstrateDef AS m ON t.SubstrateID = m.SubstrateID
LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
WHERE t.IsDeleted = 0;

CREATE VIEW V_Serial AS
SELECT
    s.rowid,
    s.Serial,
    s.OLesSerial,
    s.UsedID,
    s.ProductID,
    m.ProductName,
    m.CategoryName
FROM T_Serial AS s
LEFT JOIN M_ProductDef AS m ON s.ProductID = m.ProductID;
