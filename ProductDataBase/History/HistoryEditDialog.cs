namespace ProductDatabase.History {
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

            // 担当者をコンボボックスにバインド
            var persons = ProductDatabase.Data.PersonRepository.GetAll()
                .Where(p => p.IsActive != 0)
                .ToList();
            PersonComboBox.DataSource = persons;
            PersonComboBox.DisplayMember = "DisplayName";
            PersonComboBox.ValueMember = "PersonID";

            // 現在の担当者名から対応する PersonDef を選択
            var currentPersonInfo = GetCellText("PersonInfo");
            if (!string.IsNullOrEmpty(currentPersonInfo)) {
                var matchingPerson = persons.FirstOrDefault(p => p.PersonName == currentPersonInfo);
                if (matchingPerson != null) {
                    PersonComboBox.SelectedValue = matchingPerson.PersonID;
                }
            } else {
                PersonComboBox.SelectedIndex = -1;
            }

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

            // 担当者をコンボボックスから取得
            if (PersonComboBox.SelectedValue != null && PersonComboBox.SelectedItem is ProductDatabase.Models.PersonDef selectedPerson) {
                Person = selectedPerson.PersonName;
            } else {
                Person = null;
            }

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
