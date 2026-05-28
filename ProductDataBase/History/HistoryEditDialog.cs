namespace ProductDatabase.History {
    public partial class HistoryEditDialog : Form {

        // 外部から読み取る編集後の値（空欄はnullとして扱う）
        public string? OrderNumber { get; private set; }
        public string? ProductNumber { get; private set; }
        public string? OLesNumber { get; private set; }
        public long? PersonID { get; private set; }
        public string? PersonInfo { get; private set; }
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
            var allPersons = ProductDatabase.Data.PersonRepository.GetAll();
            var displayPersons = allPersons.Where(p => p.IsActive != 0).ToList();
            var currentPersonInfo = GetCellText("PersonInfo");

            // PersonInfo は "09.徳留" 形式なので、"徳留" だけを抽出
            var personNameOnly = currentPersonInfo;
            var dotIndex = currentPersonInfo.IndexOf('.');
            if (dotIndex > 0 && dotIndex < currentPersonInfo.Length - 1) {
                personNameOnly = currentPersonInfo.Substring(dotIndex + 1);
            }

            // 現在の担当者が無効（IsActive == 0）またはマスタに存在しない場合、一時的に追加
            if (!string.IsNullOrEmpty(personNameOnly)) {
                var matchingPerson = allPersons.FirstOrDefault(p => p.PersonName == personNameOnly);
                if (matchingPerson != null) {
                    // 削除済みだが過去のデータに使用されている担当者を追加（重複チェック）
                    if (matchingPerson.IsActive == 0 && !displayPersons.Any(p => p.PersonID == matchingPerson.PersonID)) {
                        displayPersons.Add(matchingPerson);
                    }
                }
                else if (!displayPersons.Any(p => p.PersonName.Contains(personNameOnly))) {
                    // マスタに存在しない担当者用のダミーを追加（有効なリストに同じ名前がない場合のみ）
                    displayPersons.Add(new ProductDatabase.Models.PersonDef {
                        PersonID = -1,
                        PersonName = $"(不明) {personNameOnly}",
                        IsActive = 0
                    });
                }
            }

            PersonComboBox.DataSource = displayPersons;
            PersonComboBox.DisplayMember = "DisplayName";
            PersonComboBox.ValueMember = "PersonID";

            // 現在の担当者を選択
            if (!string.IsNullOrEmpty(personNameOnly)) {
                var matchingPerson = allPersons.FirstOrDefault(p => p.PersonName == personNameOnly);
                if (matchingPerson != null) {
                    PersonComboBox.SelectedValue = matchingPerson.PersonID;
                }
                else {
                    // マスタに存在しない場合、ダミーの PersonID -1 を選択
                    PersonComboBox.SelectedValue = (long)-1;
                }
            }
            else {
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
                var selectedId = (long)PersonComboBox.SelectedValue;
                PersonID = selectedId == -1 ? null : (long?)selectedId;
                PersonInfo = selectedPerson.PersonName;
            }
            else {
                PersonID = null;
                PersonInfo = null;
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
