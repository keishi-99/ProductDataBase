using Dapper;
using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Blazor.Data {
    // 練習用のダミーデータを投入する。件数はページネーション・検索フィルターの動作確認ができる程度の規模にしている。
    internal static class SeedDummyData {
        private static readonly string[] _categories = ["制御機器", "センサー機器"];
        private static readonly string[] _productNames = ["コントローラA", "コントローラB", "センサーユニットC", "センサーユニットD"];
        private static readonly string[] _productTypes = ["標準", "防水", "耐熱"];
        private static readonly string[] _substrateNames = ["電源基板", "制御基板", "通信基板", "センサー基板"];
        private static readonly string[] _personNames = ["山田太郎", "佐藤花子", "鈴木一郎", "高橋みどり"];

        public static void Run(SqliteConnection con) {
            var random = new Random(20260824);

            var personIds = _personNames
                .Select(name => con.ExecuteScalar<long>(
                    "INSERT INTO M_Person (PersonName, IsActive) VALUES (@Name, 1) RETURNING PersonID",
                    new { Name = name }))
                .ToList();

            var productDefIds = new List<long>();
            foreach (var category in _categories) {
                foreach (var name in _productNames) {
                    var id = con.ExecuteScalar<long>("""
                        INSERT INTO M_ProductDef (CategoryName, ProductName, ProductType, ProductModel, RegType, SerialType, SerialPrintType, SheetPrintType, Visible)
                        VALUES (@Category, @Name, @Type, @Model, 1, 4, 1, 1, 1)
                        RETURNING ProductID
                        """, new {
                        Category = category,
                        Name = name,
                        Type = _productTypes[random.Next(_productTypes.Length)],
                        Model = $"MDL-{random.Next(1000, 9999)}"
                    });
                    productDefIds.Add(id);
                }
            }

            var substrateDefIds = new List<long>();
            foreach (var category in _categories) {
                foreach (var name in _substrateNames) {
                    var id = con.ExecuteScalar<long>("""
                        INSERT INTO M_SubstrateDef (CategoryName, ProductName, SubstrateName, SubstrateModel, RegType, SerialPrintType, Visible)
                        VALUES (@Category, @ProductName, @Name, @Model, 0, 0, 1)
                        RETURNING SubstrateID
                        """, new {
                        Category = category,
                        ProductName = _productNames[random.Next(_productNames.Length)],
                        Name = name,
                        Model = $"SUB-{random.Next(100, 999)}"
                    });
                    substrateDefIds.Add(id);
                }
            }

            // 基板の入荷実績（在庫のプラス分）
            var substrateStockLots = new List<(long SubstrateDefId, string LotNumber)>();
            for (var i = 0; i < 40; i++) {
                var substrateId = substrateDefIds[random.Next(substrateDefIds.Count)];
                var lotNumber = $"LOT-{random.Next(10000, 99999)}";
                var regDate = RandomDate(random);
                con.Execute("""
                    INSERT INTO T_Substrate (SubstrateID, PersonID, SubstrateNumber, OrderNumber, Increase, RegDate, Comment)
                    VALUES (@SubstrateID, @PersonID, @SubstrateNumber, @OrderNumber, @Increase, @RegDate, @Comment)
                    """, new {
                    SubstrateID = substrateId,
                    PersonID = personIds[random.Next(personIds.Count)],
                    SubstrateNumber = lotNumber,
                    OrderNumber = $"PO-{random.Next(1000, 9999)}",
                    Increase = random.Next(10, 100),
                    RegDate = regDate,
                    Comment = (string?)null
                });
                substrateStockLots.Add((substrateId, lotNumber));
            }

            // 製品登録実績（一部にシリアルと使用基板を紐づける）
            // serialCursor はシリアル番号が重複しないよう quantity 分だけ毎回進める
            var serialCursor = 2600000;
            for (var i = 0; i < 60; i++) {
                var productDefId = productDefIds[random.Next(productDefIds.Count)];
                var regDate = RandomDate(random);
                var quantity = random.Next(1, 20);
                var serialFirst = $"{serialCursor:D7}";
                var serialLast = $"{serialCursor + quantity - 1:D7}";

                var productId = con.ExecuteScalar<long>("""
                    INSERT INTO T_Product (ProductID, PersonID, OrderNumber, ProductNumber, OLesNumber, Quantity, RegDate, Revision, SerialFirst, SerialLast, Comment)
                    VALUES (@ProductDefID, @PersonID, @OrderNumber, @ProductNumber, @OLesNumber, @Quantity, @RegDate, @Revision, @SerialFirst, @SerialLast, @Comment)
                    RETURNING ID
                    """, new {
                    ProductDefID = productDefId,
                    PersonID = personIds[random.Next(personIds.Count)],
                    OrderNumber = $"PO-{random.Next(1000, 9999)}",
                    ProductNumber = $"PN-{random.Next(10000, 99999)}",
                    OLesNumber = $"OL-{random.Next(1000, 9999)}",
                    Quantity = quantity,
                    RegDate = regDate,
                    Revision = "A",
                    SerialFirst = serialFirst,
                    SerialLast = serialLast,
                    Comment = i % 7 == 0 ? "特急対応" : null
                });

                for (var s = 0; s < quantity; s++) {
                    con.Execute("""
                        INSERT INTO T_Serial (ProductID, UsedID, Serial)
                        VALUES (@ProductID, @UsedID, @Serial)
                        """, new {
                        ProductID = productDefId,
                        UsedID = productId,
                        Serial = $"{serialCursor + s:D7}"
                    });
                }
                serialCursor += quantity;

                // おおよそ半数の製品登録に、使用基板の実績（在庫減）を紐づける
                if (i % 2 == 0 && substrateStockLots.Count > 0) {
                    var (substrateId, lotNumber) = substrateStockLots[random.Next(substrateStockLots.Count)];
                    con.Execute("""
                        INSERT INTO T_Substrate (SubstrateID, PersonID, UseID, SubstrateNumber, OrderNumber, Decrease, RegDate, Comment)
                        VALUES (@SubstrateID, @PersonID, @UseID, @SubstrateNumber, @OrderNumber, @Decrease, @RegDate, @Comment)
                        """, new {
                        SubstrateID = substrateId,
                        PersonID = personIds[random.Next(personIds.Count)],
                        UseID = productId,
                        SubstrateNumber = lotNumber,
                        OrderNumber = $"PO-{random.Next(1000, 9999)}",
                        Decrease = -random.Next(1, 5),
                        RegDate = regDate,
                        Comment = (string?)null
                    });
                }
            }
        }

        // yyyy/MM/dd 形式（固定の基準日から180日以内のランダムな日付）
        // DateTime.Now を使わないのは、再生成のたびに日付が変わらないようにするため（再現性のため）
        private static readonly DateTime _seedBaseDate = new(2026, 8, 24);

        private static string RandomDate(Random random) =>
            _seedBaseDate.AddDays(-random.Next(0, 180)).ToString("yyyy/MM/dd");
    }
}
