using Microsoft.Data.Sqlite;

namespace ProductWebViewer.Data {
    // SQLiteのロック競合(SQLITE_BUSY/SQLITE_LOCKED)発生時に、ユーザー向けの分かりやすいメッセージを返すヘルパー
    internal static class SqliteBusyErrorHelper {

        // ロック競合（SQLITE_BUSY=5 / SQLITE_LOCKED=6）の場合のみ分かりやすいメッセージに変換し、それ以外は ex.Message をそのまま返す
        public static string GetUserMessage(Exception ex) {
            if (ex is SqliteException sqliteEx && (sqliteEx.SqliteErrorCode == 5 || sqliteEx.SqliteErrorCode == 6)) {
                return "他の操作でデータベースが使用中のため保存できませんでした。時間をおいて再度お試しください。";
            }
            return ex.Message;
        }
    }
}
