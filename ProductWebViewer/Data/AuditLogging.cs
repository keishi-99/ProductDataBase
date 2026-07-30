namespace ProductWebViewer.Data {
    // 監査ログ呼び出しを失敗しても本処理（DB更新）自体を失敗扱いにしないための共通ヘルパー
    internal static class AuditLogging {
        public static void TryLog(ILogger logger, Action logAction) {
            try {
                logAction();
            } catch (Exception ex) {
                logger.LogError(ex, "監査ログの記録に失敗しました。");
            }
        }
    }
}
