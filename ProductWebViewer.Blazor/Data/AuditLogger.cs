using ProductWebViewer.Blazor.Models;

namespace ProductWebViewer.Blazor.Data {
    // 編集・削除操作を db/logs/log_web_yyyyMM.csv に記録する（練習用DBの隣に作られる、本番とは別ファイル）。
    // 列構成は本家 ProductWebViewer の AuditLogger と同じにしてある（OperationLog側のパース処理と対応させるため）。
    public class AuditLogger {
        private readonly string _logDirectory;
        private static readonly Lock _lockObject = new();

        public AuditLogger(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"] ?? "db/ProductRegistry.Practice.db";
            var dbFullPath = Path.IsPathRooted(dbPath) ? dbPath : Path.Combine(AppContext.BaseDirectory, dbPath);
            _logDirectory = Path.Combine(Path.GetDirectoryName(dbFullPath) ?? AppContext.BaseDirectory, "logs");
        }

        public void LogProductEdit(ProductRecord before, string? orderNumber, string? productNumber, string? oLesNumber, string? comment) {
            AppendLog([
                BuildProductFields("[製品履歴編集:前] (Web)", before, before.OrderNumber, before.ProductNumber, before.OLesNumber, before.Comment),
                BuildProductFields("[製品履歴編集:後] (Web)", before, orderNumber, productNumber, oLesNumber, comment)
            ]);
        }

        public void LogProductDelete(ProductRecord record) {
            AppendLog(BuildProductFields("[製品履歴削除] (Web)", record, record.OrderNumber, record.ProductNumber, record.OLesNumber, record.Comment));
        }

        private static string[] BuildProductFields(string label, ProductRecord r, string? orderNumber, string? productNumber, string? oLesNumber, string? comment) => [
            label,
            $"[{r.CategoryName}]",
            $"ID[{r.Id}]",
            $"注文番号[{orderNumber}]",
            $"製造番号[{productNumber}]",
            $"OLes番号[{oLesNumber}]",
            $"製品名[{r.ProductName}]",
            $"タイプ[{r.ProductType}]",
            $"型式[{r.ProductModel}]",
            $"数量[{r.Quantity}]",
            $"シリアル先頭[{r.SerialFirst}]",
            $"シリアル末尾[{r.SerialLast}]",
            $"Revision[{r.Revision}]",
            $"登録日[{r.RegDate}]",
            $"担当者[{r.PersonInfo}]",
            $"コメント[{comment}]"
        ];

        // 製品削除に連動して削除された基板使用履歴のログ（categoryNameは親製品のカテゴリを使用）
        public void LogProductSubstrateDelete(IEnumerable<CascadeSubstrateRow> substrates, string? categoryName) {
            AppendLog(substrates.Select(item => new[] {
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
            }).ToArray());
        }

        // 製品削除に連動して削除されたシリアルのログ
        public void LogProductSerialDelete(IEnumerable<CascadeSerialRow> serials, string? categoryName) {
            AppendLog(serials.Select(item => new[] {
                "[製品削除に伴うシリアル削除] (Web)",
                $"[{categoryName}]",
                $"ID[{item.RowId}]",
                $"製品名[{item.ProductName}]",
                $"Serial[{item.Serial}]",
                $"UsedID[{item.UsedID}]",
                "[]", "[]", "[]", "[]",
                "[]", "[]", "[]", "[]",
                "[]", "[]"
            }).ToArray());
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

        private void AppendLog(string[] message) => AppendLog([message]);

        private void AppendLog(IReadOnlyList<string[]> messages) {
            lock (_lockObject) {
                var now = DateTime.Now;
                var logEntries = messages
                    .Select(message => $"\"{now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(CsvEscape))}")
                    .ToList();
                var content = string.Join(Environment.NewLine, logEntries) + Environment.NewLine;

                if (!Directory.Exists(_logDirectory)) Directory.CreateDirectory(_logDirectory);
                var logFilePath = Path.Combine(_logDirectory, $"log_web_{now:yyyyMM}.csv");

                for (var attempt = 1; attempt <= 3; attempt++) {
                    try {
                        File.AppendAllText(logFilePath, content);
                        return;
                    } catch (IOException) when (attempt < 3) {
                        Thread.Sleep(100);
                    }
                }
            }
        }

        private static string CsvEscape(string value) =>
            $"\"{value.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\"";
    }
}
