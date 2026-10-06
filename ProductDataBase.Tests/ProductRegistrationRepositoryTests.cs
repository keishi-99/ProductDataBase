using Dapper;
using ProductDatabase.Data;
using ProductDatabase.Models;
using ProductDataBase.Tests.TestSupport;

namespace ProductDataBase.Tests {
    public class ProductRegistrationRepositoryTests : IDisposable {

        private readonly InMemoryDatabase _db = new();

        public void Dispose() => _db.Dispose();

        // ===== InsertProductRecord =====

        [Fact]
        public void InsertProductRecord_空白の項目_NULLで保存される() {
            var work = new ProductRegisterWork {
                OrderNumber = "  ",
                ProductNumber = "",
                OLesNumber = " ",
                Revision = "",
                SerialFirst = "",
                SerialLast = "",
                RegDate = "2026-10-01"
            };

            InsertProduct(new ProductMaster { ProductID = 1 }, work, comment: "  ");

            Assert.Null(Scalar<string?>("SELECT OrderNumber FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT ProductNumber FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT OLesNumber FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT Revision FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT SerialFirst FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT SerialLast FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT Comment FROM T_Product"));
        }

        [Fact]
        public void InsertProductRecord_値あり_そのまま保存されROWIDを返す() {
            var work = new ProductRegisterWork {
                OrderNumber = "ORD-1",
                ProductNumber = "12345",
                Quantity = 10,
                RegDate = "2026-10-01",
                Revision = "A",
                SerialFirst = "A001",
                SerialLast = "A010"
            };

            var rowId = InsertProduct(new ProductMaster { ProductID = 7, RevisionGroup = 2 }, work, comment: "メモ");

            Assert.Equal(1, rowId);
            Assert.Equal("ORD-1", Scalar<string>("SELECT OrderNumber FROM T_Product"));
            Assert.Equal("12345", Scalar<string>("SELECT ProductNumber FROM T_Product"));
            Assert.Equal(10, Scalar<int>("SELECT Quantity FROM T_Product"));
            Assert.Equal(7, Scalar<int>("SELECT ProductID FROM T_Product"));
            Assert.Equal(2, Scalar<int>("SELECT RevisionGroup FROM T_Product"));
            Assert.Equal("メモ", Scalar<string>("SELECT Comment FROM T_Product"));
        }

        [Fact]
        public void InsertProductRecord_シリアル生成あり_SerialLastNumberが保存される() {
            var master = new ProductMaster { ProductID = 1, RegType = 1 };

            InsertProduct(master, NewWork(), serialLastNumber: 25);

            Assert.Equal(25, Scalar<int>("SELECT SerialLastNumber FROM T_Product"));
        }

        [Fact]
        public void InsertProductRecord_シリアル生成なし_SerialLastNumberはNULLになる() {
            var master = new ProductMaster { ProductID = 1, RegType = 0 };

            InsertProduct(master, NewWork(), serialLastNumber: 25);

            Assert.Null(Scalar<int?>("SELECT SerialLastNumber FROM T_Product"));
        }

        // ===== InsertRevisionChangeRecord =====

        [Fact]
        public void InsertRevisionChangeRecord_Revision変更_ROWIDを返しコメントの空白はNULLになる() {
            var work = new ProductRegisterWork { Revision = "B", RegDate = "2026-10-01", Comment = " " };
            using var tx = _db.Connection.BeginTransaction();

            var rowId = ProductRegistrationRepository.InsertRevisionChangeRecord(
                _db.Connection, tx, new ProductMaster { ProductID = 3, RevisionGroup = 1 }, work, serialLastNumber: 40);
            tx.Commit();

            Assert.Equal(1, rowId);
            Assert.Equal("B", Scalar<string>("SELECT Revision FROM T_Product"));
            Assert.Equal(40, Scalar<int>("SELECT SerialLastNumber FROM T_Product"));
            Assert.Null(Scalar<string?>("SELECT Comment FROM T_Product"));
        }

        // ===== InsertSerials =====

        [Fact]
        public void InsertSerials_複数件_全て保存されO_LesシリアルはNULLも許す() {
            var serials = new[] {
                new SerialInsertData("A001", "OA001", ProductRowId: 1, ProductId: 5),
                new SerialInsertData("A002", null, ProductRowId: 1, ProductId: 5)
            };
            using var tx = _db.Connection.BeginTransaction();

            ProductRegistrationRepository.InsertSerials(_db.Connection, tx, serials);
            tx.Commit();

            Assert.Equal(2, Scalar<int>("SELECT COUNT(*) FROM T_Serial"));
            Assert.Equal("OA001", Scalar<string>("SELECT OLesSerial FROM T_Serial WHERE Serial = 'A001'"));
            Assert.Null(Scalar<string?>("SELECT OLesSerial FROM T_Serial WHERE Serial = 'A002'"));
        }

        // ===== CheckSerialDuplication =====

        [Fact]
        public void CheckSerialDuplication_既存と同じシリアル_重複として返る() {
            var productId = AddProduct("製品A");
            AddSerial(productId, "A001");

            var duplicated = ProductRegistrationRepository.CheckSerialDuplication(_db.Connection, productId, ["A001", "A002"]);

            Assert.Equal(["A001"], duplicated);
        }

        [Fact]
        public void CheckSerialDuplication_同じ製品名の別ProductID_重複として返る() {
            var productId1 = AddProduct("製品A");
            var productId2 = AddProduct("製品A");
            AddSerial(productId1, "A001");

            var duplicated = ProductRegistrationRepository.CheckSerialDuplication(_db.Connection, productId2, ["A001"]);

            Assert.Equal(["A001"], duplicated);
        }

        [Fact]
        public void CheckSerialDuplication_前後に空白がある_Trimして照合する() {
            var productId = AddProduct("製品A");
            AddSerial(productId, "A001");

            var duplicated = ProductRegistrationRepository.CheckSerialDuplication(_db.Connection, productId, [" A001 "]);

            Assert.Equal(["A001"], duplicated);
        }

        [Fact]
        public void CheckSerialDuplication_別の製品名の同じシリアル_重複とみなさない() {
            var productId1 = AddProduct("製品A");
            var productId2 = AddProduct("製品B");
            AddSerial(productId1, "A001");

            var duplicated = ProductRegistrationRepository.CheckSerialDuplication(_db.Connection, productId2, ["A001"]);

            Assert.Empty(duplicated);
        }

        // ===== CheckOlesSerialDuplication =====

        [Fact]
        public void CheckOlesSerialDuplication_既存と同じO_Lesシリアル_重複として返る() {
            var productId = AddProduct("製品A");
            AddSerial(productId, "A001", olesSerial: "OA001");

            var duplicated = ProductRegistrationRepository.CheckOlesSerialDuplication(_db.Connection, productId, ["OA001", "OA002"]);

            Assert.Equal(["OA001"], duplicated);
        }

        // ===== FindDuplicateProductNumber / FindDuplicateOrderNumber =====

        [Fact]
        public void FindDuplicateProductNumber_同じ製番が登録済み_製品モデルを返す() {
            var productId = AddProduct("製品A", productModel: "MODEL-A");
            AddProductRecord(productId, productNumber: "12345");

            var model = FindDuplicateProductNumber("製品A", "12345");

            Assert.Equal("MODEL-A", model);
        }

        [Fact]
        public void FindDuplicateProductNumber_削除済み_重複とみなさない() {
            var productId = AddProduct("製品A");
            AddProductRecord(productId, productNumber: "12345", isDeleted: true);

            Assert.Null(FindDuplicateProductNumber("製品A", "12345"));
        }

        [Fact]
        public void FindDuplicateProductNumber_別の製品名_重複とみなさない() {
            var productId = AddProduct("製品A");
            AddProductRecord(productId, productNumber: "12345");

            Assert.Null(FindDuplicateProductNumber("製品B", "12345"));
        }

        [Fact]
        public void FindDuplicateOrderNumber_同じ注番が登録済み_製品モデルを返す() {
            var productId = AddProduct("製品A", productModel: "MODEL-A");
            AddProductRecord(productId, orderNumber: "ORD-1");
            using var tx = _db.Connection.BeginTransaction();

            var model = ProductRegistrationRepository.FindDuplicateOrderNumber(_db.Connection, tx, "製品A", "ORD-1");

            Assert.Equal("MODEL-A", model);
        }

        [Fact]
        public void FindDuplicateOrderNumber_削除済み_重複とみなさない() {
            var productId = AddProduct("製品A");
            AddProductRecord(productId, orderNumber: "ORD-1", isDeleted: true);
            using var tx = _db.Connection.BeginTransaction();

            var model = ProductRegistrationRepository.FindDuplicateOrderNumber(_db.Connection, tx, "製品A", "ORD-1");

            Assert.Null(model);
        }

        // ===== GetLatestRevision =====

        [Fact]
        public void GetLatestRevision_複数登録済み_削除済みを除いた最新を返す() {
            var productId = AddProduct("製品A");
            AddProductRecord(productId, revision: "A", revisionGroup: 1);
            AddProductRecord(productId, revision: "B", revisionGroup: 1);
            AddProductRecord(productId, revision: "C", revisionGroup: 1, isDeleted: true);

            var revision = ProductRegistrationRepository.GetLatestRevision(_db.Connection, "製品A", "1");

            Assert.Equal("B", revision);
        }

        [Fact]
        public void GetLatestRevision_該当なし_NULLを返す() {
            Assert.Null(ProductRegistrationRepository.GetLatestRevision(_db.Connection, "製品A", "1"));
        }

        // ===== GetLatestSerialLastNumber =====

        [Fact]
        public void GetLatestSerialLastNumber_NULLの行は飛ばして最新の値を返す() {
            var productId = AddProduct("製品A");
            AddProductRecord(productId, serialLastNumber: 10);
            AddProductRecord(productId, serialLastNumber: 20);
            AddProductRecord(productId, serialLastNumber: null);

            var last = ProductRegistrationRepository.GetLatestSerialLastNumber(_db.Connection, productId);

            Assert.Equal(20, last);
        }

        [Fact]
        public void GetLatestSerialLastNumber_同じ製品名の別ProductIDの登録も対象になる() {
            var productId1 = AddProduct("製品A");
            var productId2 = AddProduct("製品A");
            AddProductRecord(productId1, serialLastNumber: 10);
            AddProductRecord(productId2, serialLastNumber: 30);

            var last = ProductRegistrationRepository.GetLatestSerialLastNumber(_db.Connection, productId1);

            Assert.Equal(30, last);
        }

        [Fact]
        public void GetLatestSerialLastNumber_登録なし_NULLを返す() {
            var productId = AddProduct("製品A");

            Assert.Null(ProductRegistrationRepository.GetLatestSerialLastNumber(_db.Connection, productId));
        }

        // ===== UpdateOLesSerialSuffix =====

        [Fact]
        public void UpdateOLesSerialSuffix_指定製品だけ更新される() {
            var productId1 = AddProduct("製品A");
            var productId2 = AddProduct("製品B");
            using var tx = _db.Connection.BeginTransaction();

            ProductRegistrationRepository.UpdateOLesSerialSuffix(_db.Connection, tx, productId1, "B");
            tx.Commit();

            Assert.Equal("B", Scalar<string>("SELECT OLesSerialSuffix FROM M_ProductDef WHERE ProductID = @productId1", new { productId1 }));
            Assert.Null(Scalar<string?>("SELECT OLesSerialSuffix FROM M_ProductDef WHERE ProductID = @productId2", new { productId2 }));
        }

        // ===== 基板の使用・在庫 =====

        [Fact]
        public void InsertSubstrateUsage_使用数_負数のDecreaseとして保存される() {
            using var tx = _db.Connection.BeginTransaction();

            ProductRegistrationRepository.InsertSubstrateUsage(
                _db.Connection, tx, substrateID: 1, substrateNumber: "N1", orderNumber: " ",
                useValue: 3, useID: 9, personID: null, regDate: "2026-10-01", comment: "");
            tx.Commit();

            Assert.Equal(-3, Scalar<int>("SELECT Decrease FROM T_Substrate"));
            Assert.Equal("N1", Scalar<string>("SELECT SubstrateNumber FROM T_Substrate"));
            Assert.Null(Scalar<string?>("SELECT OrderNumber FROM T_Substrate"));
            Assert.Null(Scalar<string?>("SELECT Comment FROM T_Substrate"));
            Assert.Equal(9, Scalar<int>("SELECT UseID FROM T_Substrate"));
        }

        [Fact]
        public void FetchSubstrateStock_在庫が0を超える明細だけ返す() {
            var substrateId = AddSubstrate("基板A");
            AddSubstrateRecord(substrateId, "N1", increase: 10);
            AddSubstrateRecord(substrateId, "N1", decrease: -10);
            AddSubstrateRecord(substrateId, "N2", increase: 5);
            AddSubstrateRecord(substrateId, "N2", decrease: -2);
            using var tx = _db.Connection.BeginTransaction();

            var stocks = ProductRegistrationRepository.FetchSubstrateStock(_db.Connection, tx, substrateId).ToList();

            var stock = Assert.Single(stocks);
            Assert.Equal("N2", stock.SubstrateNumber);
            Assert.Equal(3, stock.Stock);
        }

        [Fact]
        public void GetSubstrateOrderNumber_該当あり_注番を返す() {
            var substrateId = AddSubstrate("基板A");
            AddSubstrateRecord(substrateId, "N1", increase: 5, orderNumber: "ORD-1");
            using var tx = _db.Connection.BeginTransaction();

            var orderNumber = ProductRegistrationRepository.GetSubstrateOrderNumber(_db.Connection, tx, substrateId, "N1");

            Assert.Equal("ORD-1", orderNumber);
        }

        [Fact]
        public void GetSubstrateOrderNumber_該当なし_空文字を返す() {
            var substrateId = AddSubstrate("基板A");
            using var tx = _db.Connection.BeginTransaction();

            var orderNumber = ProductRegistrationRepository.GetSubstrateOrderNumber(_db.Connection, tx, substrateId, "N1");

            Assert.Equal(string.Empty, orderNumber);
        }

        // ===== CheckRegistrationExists =====

        [Fact]
        public void CheckRegistrationExists_登録済みのID_例外にならない() {
            var productId = AddProduct("製品A");
            var rowId = AddProductRecord(productId);

            var exception = Record.Exception(() => ProductRegistrationRepository.CheckRegistrationExists(_db.Connection, rowId));

            Assert.Null(exception);
        }

        [Fact]
        public void CheckRegistrationExists_存在しないID_例外になる() {
            var exception = Assert.Throws<Exception>(() => ProductRegistrationRepository.CheckRegistrationExists(_db.Connection, 999));

            Assert.Equal("登録に失敗しました。IDが見つかりません。", exception.Message);
        }

        // ===== トランザクション =====

        [Fact]
        public void InsertProductRecord_ロールバック_登録が残らない() {
            using (var tx = _db.Connection.BeginTransaction()) {
                ProductRegistrationRepository.InsertProductRecord(
                    _db.Connection, tx, new ProductMaster { ProductID = 1 }, NewWork(), comment: "", serialLastNumber: 0);
                tx.Rollback();
            }

            Assert.Equal(0, Scalar<int>("SELECT COUNT(*) FROM T_Product"));
        }

        // ===== テスト用の補助 =====

        private static ProductRegisterWork NewWork() => new() { RegDate = "2026-10-01" };

        private T Scalar<T>(string sql, object? param = null) => _db.Connection.ExecuteScalar<T>(sql, param)!;

        // トランザクションを開始してInsertProductRecordを呼び、コミットしてROWIDを返す
        private long InsertProduct(ProductMaster master, ProductRegisterWork work, string comment = "", int serialLastNumber = 0) {
            using var tx = _db.Connection.BeginTransaction();
            var rowId = ProductRegistrationRepository.InsertProductRecord(_db.Connection, tx, master, work, comment, serialLastNumber);
            tx.Commit();
            return rowId;
        }

        private string? FindDuplicateProductNumber(string productName, string productNumber) {
            using var tx = _db.Connection.BeginTransaction();
            return ProductRegistrationRepository.FindDuplicateProductNumber(_db.Connection, tx, productName, productNumber);
        }

        private long AddProduct(string productName, string? productModel = null) =>
            _db.Connection.ExecuteScalar<long>(
                "INSERT INTO M_ProductDef (ProductName, ProductModel) VALUES (@productName, @productModel); SELECT last_insert_rowid();",
                new { productName, productModel });

        private long AddProductRecord(
            long productId, string? orderNumber = null, string? productNumber = null, string? revision = null,
            int? revisionGroup = null, int? serialLastNumber = null, bool isDeleted = false) =>
            _db.Connection.ExecuteScalar<long>(
                """
                INSERT INTO T_Product (ProductID, OrderNumber, ProductNumber, Revision, RevisionGroup, SerialLastNumber, RegDate, IsDeleted)
                VALUES (@productId, @orderNumber, @productNumber, @revision, @revisionGroup, @serialLastNumber, '2026-10-01', @isDeleted);
                SELECT last_insert_rowid();
                """,
                new { productId, orderNumber, productNumber, revision, revisionGroup, serialLastNumber, isDeleted });

        private void AddSerial(long productId, string serial, string? olesSerial = null) =>
            _db.Connection.Execute(
                "INSERT INTO T_Serial (ProductID, UsedID, Serial, OLesSerial) VALUES (@productId, 1, @serial, @olesSerial)",
                new { productId, serial, olesSerial });

        private long AddSubstrate(string substrateName) =>
            _db.Connection.ExecuteScalar<long>(
                "INSERT INTO M_SubstrateDef (SubstrateName) VALUES (@substrateName); SELECT last_insert_rowid();",
                new { substrateName });

        private void AddSubstrateRecord(
            long substrateId, string substrateNumber, int? increase = null, int? decrease = null, string? orderNumber = null) =>
            _db.Connection.Execute(
                """
                INSERT INTO T_Substrate (SubstrateID, SubstrateNumber, OrderNumber, Increase, Decrease, RegDate)
                VALUES (@substrateId, @substrateNumber, @orderNumber, @increase, @decrease, '2026-10-01')
                """,
                new { substrateId, substrateNumber, orderNumber, increase, decrease });
    }
}
