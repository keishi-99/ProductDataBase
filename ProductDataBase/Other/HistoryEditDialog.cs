using System.Data;

namespace ProductDatabase.Other {
    public partial class HistoryEditDialog : Form {

        // 外部から読み取る編集後の値
        public string OrderNumber    { get; private set; } = string.Empty;
        public string ProductNumber  { get; private set; } = string.Empty;
        public string OLesNumber     { get; private set; } = string.Empty;
        public string Person         { get; private set; } = string.Empty;
        public string Comment        { get; private set; } = string.Empty;

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
            IdValueLabel.Text           = GetCellText("ID");
            CategoryNameValueLabel.Text = GetCellText("CategoryName");
            ProductNameValueLabel.Text  = GetCellText("ProductName");
            ProductModelValueLabel.Text = GetCellText("ProductModel");

            OrderNumberTextBox.Text     = GetCellText("OrderNumber");
            ProductNumberTextBox.Text   = GetCellText("ProductNumber");
            OLesNumberTextBox.Text      = GetCellText("OLesNumber");
            PersonTextBox.Text          = GetCellText("Person");
            RegDateValueLabel.Text      = GetCellText("RegDate");
            RevisionValueLabel.Text     = GetCellText("Revision");
            RevisionGroupValueLabel.Text = GetCellText("RevisionGroup");
            CommentTextBox.Text         = GetCellText("Comment");
        }

        // セル値を文字列で取得しnullの場合は空文字を返す
        private string GetCellText(string columnName) {
            if (_sourceRow.Cells[columnName] is DataGridViewCell cell) {
                return cell.Value?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        // 保存ボタン：コントロールの値をプロパティに格納してダイアログを閉じる
        private void SaveButton_Click(object sender, EventArgs e) {
            OrderNumber   = OrderNumberTextBox.Text;
            ProductNumber = ProductNumberTextBox.Text;
            OLesNumber    = OLesNumberTextBox.Text;
            Person        = PersonTextBox.Text;
            Comment       = CommentTextBox.Text;

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
