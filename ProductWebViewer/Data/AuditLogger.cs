using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // WebViewer経由での編集・削除操作を、メインアプリ(ProductDataBase)と共通の
    // ログファイル(db/logs/log_yyyyMM.csv)に記録する。
    // 列構成・エスケープ処理はメインアプリの Logger.AppendLog / HistoryAuditLogger に合わせてある
    // （列見出し: 日時,操作種別,カテゴリ,ID,注文番号,製造番号,OLes番号,製品名,タイプ,型式,数量,シリアル先頭,シリアル末尾,Revision,登録日,担当者,コメント）
    public class AuditLogger {
        private readonly string _logDirectory;
        private static readonly object _lockObject = new();

        public AuditLogger(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"] ?? "db/ProductRegistry.db";
            var dbFullPath = Path.IsPathRooted(dbPath) ? dbPath : Path.Combine(AppContext.BaseDirectory, dbPath);
            // メインアプリの Logger._logDirectory ( {BaseDirectory}/db/logs ) と同じ場所を指す
            _logDirectory = Path.Combine(Path.GetDirectoryName(dbFullPath) ?? AppContext.BaseDirectory, "logs");
        }

        public void LogProductEdit(ProductRecord before, ProductRecord after) {
            AppendLog([
                "[製品履歴編集:前] (Web)",
                $"[{before.CategoryName}]",
                $"ID[{before.Id}]",
                $"注文番号[{before.OrderNumber}]",
                $"製造番号[{before.ProductNumber}]",
                $"OLes番号[{before.OLesNumber}]",
                $"製品名[{before.ProductName}]",
                $"タイプ[{before.ProductType}]",
                $"型式[{before.ProductModel}]",
                $"数量[{before.Quantity}]",
                $"シリアル先頭[{before.SerialFirst}]",
                $"シリアル末尾[{before.SerialLast}]",
                $"Revision[{before.Revision}]",
                $"登録日[{before.RegDate}]",
                $"担当者[{before.PersonInfo}]",
                $"コメント[{before.Comment}]"
            ]);
            AppendLog([
                "[製品履歴編集:後] (Web)",
                $"[{after.CategoryName}]",
                $"ID[{after.Id}]",
                $"注文番号[{after.OrderNumber}]",
                $"製造番号[{after.ProductNumber}]",
                $"OLes番号[{after.OLesNumber}]",
                $"製品名[{after.ProductName}]",
                $"タイプ[{after.ProductType}]",
                $"型式[{after.ProductModel}]",
                $"数量[{after.Quantity}]",
                $"シリアル先頭[{after.SerialFirst}]",
                $"シリアル末尾[{after.SerialLast}]",
                $"Revision[{after.Revision}]",
                $"登録日[{after.RegDate}]",
                $"担当者[{after.PersonInfo}]",
                $"コメント[{after.Comment}]"
            ]);
        }

        public void LogProductDelete(ProductRecord record) {
            AppendLog([
                "[製品履歴削除] (Web)",
                $"[{record.CategoryName}]",
                $"ID[{record.Id}]",
                $"注文番号[{record.OrderNumber}]",
                $"製造番号[{record.ProductNumber}]",
                $"OLes番号[{record.OLesNumber}]",
                $"製品名[{record.ProductName}]",
                $"タイプ[{record.ProductType}]",
                $"型式[{record.ProductModel}]",
                $"数量[{record.Quantity}]",
                $"シリアル先頭[{record.SerialFirst}]",
                $"シリアル末尾[{record.SerialLast}]",
                $"Revision[{record.Revision}]",
                $"登録日[{record.RegDate}]",
                $"担当者[{record.PersonInfo}]",
                $"コメント[{record.Comment}]"
            ]);
        }

        // 製品削除に連動して削除された基板使用履歴1件ごとのログ（categoryNameは親製品のカテゴリを使用）
        public void LogProductSubstrateDelete(IEnumerable<dynamic> substrates, string? categoryName) {
            foreach (var item in substrates) {
                AppendLog([
                    "[製品削除に伴う基板削除] (Web)",
                    $"[{categoryName}]",
                    $"ID[{item.ID}]",
                    $"注文番号[{item.OrderNumber}]",
                    $"製造番号[{item.SubstrateNumber}]",
                    "[]",
                    $"製品名[{item.ProductName}]",
                    $"基板名[{item.SubstrateName}]",
                    $"型式[{item.SubstrateModel}]",
                    $"追加数[{item.Increase}]",
                    $"使用数[{item.Decrease}]",
                    $"減少数[{item.Defect}]",
                    $"登録日[{item.RegDate}]",
                    $"担当者[{item.PersonInfo}]",
                    $"コメント[{item.Comment}]",
                    $"UseID[{item.UseID}]"
                ]);
            }
        }

        // 製品削除に連動して削除されたシリアル1件ごとのログ
        public void LogProductSerialDelete(IEnumerable<dynamic> serials, string? categoryName) {
            foreach (var item in serials) {
                AppendLog([
                    "[製品削除に伴うシリアル削除] (Web)",
                    $"[{categoryName}]",
                    $"ID[{item.rowid}]",
                    $"製品名[{item.ProductName}]",
                    $"Serial[{item.Serial}]",
                    $"UsedID[{item.UsedID}]",
                    "[]", "[]", "[]", "[]",
                    "[]", "[]", "[]", "[]",
                    "[]", "[]"
                ]);
            }
        }

        public void LogSubstrateDelete(SubstrateRecord record) {
            AppendLog([
                "[基板履歴削除] (Web)",
                $"[{record.CategoryName}]",
                $"ID[{record.Id}]",
                $"注文番号[{record.OrderNumber}]",
                $"製造番号[{record.SubstrateNumber}]",
                "[]",
                $"製品名[{record.ProductName}]",
                $"基板名[{record.SubstrateName}]",
                $"型式[{record.SubstrateModel}]",
                $"追加数[{record.Increase}]",
                $"使用数[{record.Decrease}]",
                $"減少数[{record.Defect}]",
                "[]",
                $"登録日[{record.RegDate}]",
                $"担当者[{record.PersonInfo}]",
                $"コメント[{record.Comment}]"
            ]);
        }

        // メインアプリの Logger.AppendLog と同じ列構成・エスケープでCSV1行を追記する
        // lockはWebViewerプロセス内のみの排他のため、メインアプリと同時書き込みが重なった場合に備えて
        // 数回だけ短い間隔でリトライする（両アプリを跨ぐ完全な排他にはならないが、瞬間的な衝突は緩和できる）
        private void AppendLog(string[] message) {
            var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(CsvEscape))}";

            lock (_lockObject) {
                if (!Directory.Exists(_logDirectory)) Directory.CreateDirectory(_logDirectory);
                var logFilePath = Path.Combine(_logDirectory, $"log_{DateTime.Now:yyyyMM}.csv");

                for (var attempt = 1; attempt <= 3; attempt++) {
                    try {
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
                        return;
                    } catch (IOException) when (attempt < 3) {
                        Thread.Sleep(100);
                    }
                }
            }
        }

        // ダブルクォートのエスケープと改行の除去（改行を残すとCSVの行構造が壊れ、偽の行を追記できてしまう）
        private static string CsvEscape(string value) =>
            $"\"{value.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\"";
    }
}
