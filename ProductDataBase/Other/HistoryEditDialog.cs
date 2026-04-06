namespace ProductDatabase.Other {
    public partial class HistoryEditDialog : Form {

        // 外部から読み取る編集後の値（空欄はnullとして扱う）
        public string? OrderNumber { get; private set; }
        public string? ProductNumber { get; private set; }
        public string? OLesNumber { get; private set; }
        public string? Person { get; private set; }
        public string? Comment { get; private set; }

        private readonly DataGridViewRow _sourceRow;

        public HistoryEditDialog(DataGridViewRow sourceRow) {
            InitializeComponent();
            _sourceRow = sourceRow;
        }

        private void HistoryEditDialog_Load(object sender, EventArgs e) {
            PopulateFields();
        }

        // DataGridViewRowの値をコントロールに設定する
        private void PopulateFields() {
            IdValueLabel.Text = GetCellText("ID");
            CategoryNameValueLabel.Text = GetCellText("CategoryName");
            ProductNameValueLabel.Text = GetCellText("ProductName");
            ProductModelValueLabel.Text = GetCellText("ProductModel");

            OrderNumberTextBox.Text = GetCellText("OrderNumber");
            ProductNumberTextBox.Text = GetCellText("ProductNumber");
            OLesNumberTextBox.Text = GetCellText("OLesNumber");
            PersonTextBox.Text = GetCellText("Person");
            RegDateValueLabel.Text = GetCellText("RegDate");
            RevisionValueLabel.Text = GetCellText("Revision");
            RevisionGroupValueLabel.Text = GetCellText("RevisionGroup");
            CommentTextBox.Text = GetCellText("Comment");
        }

        // セル値を文字列で取得しnullの場合は空文字を返す
        private string GetCellText(string columnName) {
            if (_sourceRow.Cells[columnName] is DataGridViewCell cell) {
                return cell.Value?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        // 保存ボタン：コントロールの値をプロパティに格納してダイアログを閉じる（空欄はnull）
        private void SaveButton_Click(object sender, EventArgs e) {
            OrderNumber = string.IsNullOrWhiteSpace(OrderNumberTextBox.Text) ? null : OrderNumberTextBox.Text;
            ProductNumber = string.IsNullOrWhiteSpace(ProductNumberTextBox.Text) ? null : ProductNumberTextBox.Text;
            OLesNumber = string.IsNullOrWhiteSpace(OLesNumberTextBox.Text) ? null : OLesNumberTextBox.Text;
            Person = string.IsNullOrWhiteSpace(PersonTextBox.Text) ? null : PersonTextBox.Text;
            Comment = string.IsNullOrWhiteSpace(CommentTextBox.Text) ? null : CommentTextBox.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // キャンセルボタン
        private void CancelButton_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
