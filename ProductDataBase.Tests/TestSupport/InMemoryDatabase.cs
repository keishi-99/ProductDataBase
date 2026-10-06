using Dapper;
using Microsoft.Data.Sqlite;

namespace ProductDataBase.Tests.TestSupport {
    // テスト用のメモリ上SQLite。接続を開いている間だけDBが存在するので、テストごとに作り直して独立させる
    internal sealed class InMemoryDatabase : IDisposable {

        public SqliteConnection Connection { get; }

        public InMemoryDatabase() {
            Connection = new SqliteConnection("Data Source=:memory:");
            Connection.Open();
            Connection.Execute(SchemaSql);
        }

        public void Dispose() => Connection.Dispose();

        // 登録処理が参照するテーブル・ビューだけを再現したスキーマ（列は本体が使うものに絞っている）
        // 本体のスキーマ変更に追従して手で更新すること（ビューは ProductRepository の定義と同じ内容）
        private const string SchemaSql = """
            CREATE TABLE M_Person (
                PersonID   INTEGER PRIMARY KEY AUTOINCREMENT,
                PersonName TEXT NOT NULL
            );

            CREATE TABLE M_ProductDef (
                ProductID        INTEGER PRIMARY KEY AUTOINCREMENT,
                CategoryName     TEXT NOT NULL DEFAULT '',
                ProductName      TEXT NOT NULL,
                ProductType      TEXT,
                ProductModel     TEXT,
                OLesSerialSuffix TEXT
            );

            CREATE TABLE M_SubstrateDef (
                SubstrateID    INTEGER PRIMARY KEY AUTOINCREMENT,
                CategoryName   TEXT NOT NULL DEFAULT '',
                ProductName    TEXT NOT NULL DEFAULT '',
                SubstrateName  TEXT NOT NULL,
                SubstrateModel TEXT
            );

            CREATE TABLE T_Product (
                ID               INTEGER PRIMARY KEY AUTOINCREMENT,
                ProductID        INTEGER NOT NULL,
                PersonID         INTEGER,
                OrderNumber      TEXT,
                ProductNumber    TEXT,
                OLesNumber       TEXT,
                Quantity         INTEGER,
                RegDate          TEXT NOT NULL,
                Revision         TEXT,
                RevisionGroup    INTEGER,
                SerialFirst      TEXT,
                SerialLast       TEXT,
                SerialLastNumber INTEGER,
                Comment          TEXT,
                CreatedAt        TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
                IsDeleted        INTEGER NOT NULL DEFAULT 0,
                DeletedAt        TEXT
            );

            CREATE TABLE T_Substrate (
                ID              INTEGER PRIMARY KEY AUTOINCREMENT,
                SubstrateID     INTEGER NOT NULL,
                PersonID        INTEGER,
                UseID           INTEGER,
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
                ProductID  INTEGER NOT NULL,
                UsedID     INTEGER,
                Serial     TEXT NOT NULL,
                OLesSerial TEXT
            );

            CREATE VIEW V_Product AS
            SELECT
                t.ID, t.ProductID, m.CategoryName, m.ProductName, m.ProductType, m.ProductModel,
                t.OrderNumber, t.ProductNumber, t.OLesNumber, t.Quantity, t.SerialFirst, t.SerialLast,
                t.Revision, t.RevisionGroup, t.SerialLastNumber,
                COALESCE(p.PersonName, '') AS PersonInfo,
                t.RegDate, t.Comment, t.CreatedAt, t.IsDeleted, t.DeletedAt
            FROM T_Product AS t
            LEFT JOIN M_ProductDef AS m ON t.ProductID = m.ProductID
            LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
            WHERE t.IsDeleted = 0;

            CREATE VIEW V_Substrate AS
            SELECT
                t.ID, t.SubstrateID, m.CategoryName, m.ProductName, m.SubstrateName, m.SubstrateModel,
                t.OrderNumber, t.SubstrateNumber, t.Increase, t.Decrease, t.Defect,
                COALESCE(p.PersonName, '') AS PersonInfo,
                t.RegDate, t.Comment, t.UseID, t.CreatedAt, t.IsDeleted, t.DeletedAt
            FROM T_Substrate AS t
            LEFT JOIN M_SubstrateDef AS m ON t.SubstrateID = m.SubstrateID
            LEFT JOIN M_Person AS p ON t.PersonID = p.PersonID
            WHERE t.IsDeleted = 0;
            """;
    }
}
