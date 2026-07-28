using ProductWebViewer.Models;

namespace ProductWebViewer.Data {
    // WebViewer経由での編集・削除操作を追跡するための監査ログ
    // メインアプリのログ形式とは独立した、Web操作専用のシンプルなテキストログ
    public class AuditLogger {
        private readonly string _logDirectory;
        private static readonly object _lockObject = new();

        public AuditLogger(IConfiguration configuration) {
            var dbPath = configuration["DatabasePath"] ?? "db/ProductRegistry.db";
            var dbFullPath = Path.IsPathRooted(dbPath) ? dbPath : Path.Combine(AppContext.BaseDirectory, dbPath);
            _logDirectory = Path.Combine(Path.GetDirectoryName(dbFullPath) ?? AppContext.BaseDirectory, "auditlog");
        }

        public void LogProductEdit(string operatorName, ProductRecord before, ProductRecord after) {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Append(
                $"[{timestamp}] 編集 操作者[{operatorName}] ID[{after.Id}] 製品名[{after.ProductName}] " +
                $"注文番号[{before.OrderNumber}]→[{after.OrderNumber}] " +
                $"製造番号[{before.ProductNumber}]→[{after.ProductNumber}] " +
                $"O-Les番号[{before.OLesNumber}]→[{after.OLesNumber}] " +
                $"登録日[{before.RegDate}]→[{after.RegDate}] " +
                $"Revision[{before.Revision}]→[{after.Revision}] " +
                $"コメント[{before.Comment}]→[{after.Comment}]");
        }

        public void LogProductDelete(string operatorName, ProductRecord record) {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Append(
                $"[{timestamp}] 削除 操作者[{operatorName}] ID[{record.Id}] 製品名[{record.ProductName}] " +
                $"注文番号[{record.OrderNumber}] 製造番号[{record.ProductNumber}] O-Les番号[{record.OLesNumber}] " +
                $"登録日[{record.RegDate}] Revision[{record.Revision}] コメント[{record.Comment}]");
        }

        public void LogSubstrateDelete(string operatorName, SubstrateRecord record) {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Append(
                $"[{timestamp}] 削除(基板) 操作者[{operatorName}] ID[{record.Id}] 製品名[{record.ProductName}] " +
                $"基板名[{record.SubstrateName}] 基板型式[{record.SubstrateModel}] " +
                $"注文番号[{record.OrderNumber}] 製造番号[{record.SubstrateNumber}] " +
                $"入庫[{record.Increase}] 出庫[{record.Decrease}] 不良[{record.Defect}] " +
                $"登録日[{record.RegDate}] コメント[{record.Comment}]");
        }

        private void Append(string line) {
            lock (_lockObject) {
                if (!Directory.Exists(_logDirectory)) Directory.CreateDirectory(_logDirectory);
                var logFile = Path.Combine(_logDirectory, $"webviewer_{DateTime.Now:yyyyMM}.log");
                File.AppendAllText(logFile, line + Environment.NewLine);
            }
        }
    }
}
