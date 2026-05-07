namespace ProductWebViewer.Models {
    public class SerialRecord {
        // T_Serial テーブルの SQLite 物理 rowid（アプリケーション ID とは別）
        public long RowId { get; set; }
        public string? Serial { get; set; }
        public string? OLesSerial { get; set; }
        public string? OrderNumber { get; set; }
        public string? ProductNumber { get; set; }
        public string? ProductName { get; set; }
        public string? ProductType { get; set; }
        public string? ProductModel { get; set; }
        public string? RegDate { get; set; }
    }
}
