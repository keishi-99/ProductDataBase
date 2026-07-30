namespace ProductWebViewer.Models {
    // db/logs/log_yyyyMM.csv の1行に対応する（メインアプリのLogViewerWindowと同じ列構成）
    public class LogEntry {
        public string Timestamp { get; set; } = string.Empty;
        public string OperationType { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string OrderNumber { get; set; } = string.Empty;
        public string ProductNumber { get; set; } = string.Empty;
        public string OLesNumber { get; set; } = string.Empty;
        public string ProductName { get; set; } = string.Empty;
        public string ProductType { get; set; } = string.Empty;
        public string ProductModel { get; set; } = string.Empty;
        public string Quantity { get; set; } = string.Empty;
        public string SerialFirst { get; set; } = string.Empty;
        public string SerialLast { get; set; } = string.Empty;
        public string Revision { get; set; } = string.Empty;
        public string RegDate { get; set; } = string.Empty;
        public string PersonInfo { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
    }
}
