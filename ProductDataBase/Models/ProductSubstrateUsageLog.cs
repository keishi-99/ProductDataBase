namespace ProductDatabase.Models {
    // 製品登録に伴う基板引き落とし1件分のログ情報
    public record ProductSubstrateUsageLog(
        string ProductName, string SubstrateName, string SubstrateModel,
        string SubstrateNumber, string OrderNumber, long UseValue,
        string RegDate, string? PersonName, string? Comment);
}
