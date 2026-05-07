namespace ProductWebViewer.Models {
    public class ProductRecord {
        public long Id { get; set; }
        public long ProductID { get; set; }
        public string? CategoryName { get; set; }
        public string? ProductName { get; set; }
        public string? ProductModel { get; set; }
        public string? ProductType { get; set; }
        public string? OrderNumber { get; set; }
        public string? ProductNumber { get; set; }
        public string? OLesNumber { get; set; }
        public long? Quantity { get; set; }
        public string? Person { get; set; }
        public string? RegDate { get; set; }
        public string? Revision { get; set; }
        public string? SerialFirst { get; set; }
        public string? SerialLast { get; set; }
        public string? Comment { get; set; }
        public string? CreatedAt { get; set; }
    }
}
