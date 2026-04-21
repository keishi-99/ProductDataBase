namespace ProductDatabase.Common {
    internal static class LogOperationTypes {
        internal const string RevChange = "[Rev変更]";
        internal const string RePrint = "[再印刷]";
        internal const string SubstrateRegistration = "[基板登録]";
        internal const string SubstrateChange = "[基板変更]";
        internal const string ProductRegistration = "[製品登録]";
        internal const string SubstrateHistoryEditBefore = "[基板履歴編集:前]";
        internal const string SubstrateHistoryEditAfter = "[基板履歴編集:後]";
        internal const string SubstrateHistoryDelete = "[基板履歴削除]";
        internal const string ProductHistoryEditBefore = "[製品履歴編集:前]";
        internal const string ProductHistoryEditAfter = "[製品履歴編集:後]";
        internal const string ProductHistoryDelete = "[製品履歴削除]";
        internal const string ProductRelatedSubstrateDelete = "[製品削除に伴う基板削除]";
        internal const string ProductRelatedSerialDelete = "[製品削除に伴うシリアル削除]";
        internal const string SerialHistoryDelete = "[シリアル履歴削除]";
    }
}
