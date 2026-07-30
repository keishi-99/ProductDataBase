using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // WebViewer経由での編集・削除操作を db/logs/log_web_yyyyMM.csv に記録する。
    // メインアプリの db/logs/log_yyyyMM.csv とはファイルを分けている（複数プロセスからの同時書き込みを避けるため）が、
    // 列構成・エスケープ処理はメインアプリの Logger.AppendLog / HistoryAuditLogger に合わせてあり、
    // 閲覧側（OperationLog・メインアプリのLogViewerWindow）で両ファイルをマージして表示する。
    // 【重要】この2ファイルは列構成・区切り文字・改行エスケープを完全に一致させる必要がある。
    // 片方だけ変更すると閲覧側で列がずれる・色分けが効かなくなる等の不整合が起きるため、
    // 変更する際は必ずペアで更新すること：
    //   - ProductDataBase/Common/Logger.cs（メインアプリの書き込み側）
    //   - ProductDataBase/LogViewer/LogViewerWindow.cs（メインアプリの閲覧側パーサー）
    //   - ProductWebViewer/Data/LogRecordRepository.cs（WebViewerの閲覧側パーサー）
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

        // 編集可能なのは注文番号・製造番号・OLes番号・コメントのみ（メインアプリのHistoryEditDialogに合わせている）。
        // それ以外の項目はbeforeの値をそのまま「後」欄にも使う（変更されないため）
        public void LogProductEdit(ProductRecord before, string? orderNumber, string? productNumber, string? oLesNumber, string? comment) {
            AppendLog([
                BuildProductFields("[製品履歴編集:前] (Web)", before, before.OrderNumber, before.ProductNumber, before.OLesNumber, before.Comment),
                BuildProductFields("[製品履歴編集:後] (Web)", before, orderNumber, productNumber, oLesNumber, comment)
            ]);
        }

        public void LogProductDelete(ProductRecord record) {
            AppendLog(BuildProductFields("[製品履歴削除] (Web)", record, record.OrderNumber, record.ProductNumber, record.OLesNumber, record.Comment));
        }

        // 製品履歴ログ1行分の項目を組み立てる（編集可能な4項目だけを差し替えられるようにしてある）
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

        // 1行だけ追記する場合のオーバーロード
        private void AppendLog(string[] message) => AppendLog([message]);

        // メインアプリのLogger.AppendLogと同じ列構成・エスケープでCSV行を追記する。
        // ファイル名を log_web_*.csv とし、メインアプリの log_*.csv とは別ファイルに分けている
        // （同じファイルに複数プロセスが書き込むとロック競合の懸念があるため）。
        // 複数行はまとめて1回のファイルオープンで書き込む（連動削除のように多数の行を記録する場合の
        // 開閉コストを抑えるため）。
        // アンチウイルス・バックアップソフト等による一時的なファイルロックに備え、IOException時は
        // 数回だけ短い間隔でリトライする。
        private void AppendLog(IReadOnlyList<string[]> messages) {
            lock (_lockObject) {
                // 行の日時とファイル名(年月)がずれないよう、同じ時刻を使い回す
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

        // ダブルクォートのエスケープと改行の除去（改行を残すとCSVの行構造が壊れ、偽の行を追記できてしまう）
        private static string CsvEscape(string value) =>
            $"\"{value.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\"";
    }
}
