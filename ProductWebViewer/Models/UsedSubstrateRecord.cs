namespace ProductWebViewer.Models {
    public class UsedSubstrateRecord {
        public long Id { get; set; }
        public string? SubstrateName { get; set; }
        public string? SubstrateModel { get; set; }
        public string? SubstrateNumber { get; set; }
        // 使用数量（DB に負数で格納されているため表示時は絶対値で扱う）
        public int? Decrease { get; set; }
    }
}
