using Dapper;
using System.Data.Odbc;

namespace ProductDatabase.Services {
    internal class BarcodeService(string dsn, string uid, string pwd) {
        private readonly string _dsn = dsn;
        private readonly string _uid = uid;
        private readonly string _pwd = pwd;

        // バーコードの手配管理番号からODBCで手配情報を取得して返す
        public BarcodeQueryResult Query(string managementNumber) {
            using var con = new OdbcConnection($"DSN={_dsn}; UID={_uid}; PWD={_pwd}");
            con.Open();
            var param = new DynamicParameters();
            param.Add("手配管理番号", managementNumber);
            var result = con.QueryFirstOrDefault<OrderDto>(
                @"SELECT 手配製番, 品目番号, 品目名称, 手配数, 請求先注番
                  FROM V_宮崎手配情報
                  WHERE 手配管理番号 = ?",
                param)
                ?? throw new Exception($"一致する情報がありません。{Environment.NewLine}手配管理番号:{managementNumber}");

            return new BarcodeQueryResult {
                ProductNumber = result.手配製番 ?? string.Empty,
                ProductModel = result.品目番号 ?? string.Empty,
                ProductName = result.品目名称 ?? string.Empty,
                Quantity = result.手配数,
                OrderNumber = result.請求先注番 ?? string.Empty
            };
        }

        private sealed class OrderDto {
            public string 手配製番 { get; set; } = string.Empty;
            public string 品目番号 { get; set; } = string.Empty;
            public string 品目名称 { get; set; } = string.Empty;
            public int 手配数 { get; set; }
            public string 請求先注番 { get; set; } = string.Empty;
        }
    }

    internal sealed class BarcodeQueryResult {
        public string ProductNumber { get; init; } = string.Empty;
        public string ProductModel { get; init; } = string.Empty;
        public string ProductName { get; init; } = string.Empty;
        public int Quantity { get; init; }
        public string OrderNumber { get; init; } = string.Empty;
    }
}
