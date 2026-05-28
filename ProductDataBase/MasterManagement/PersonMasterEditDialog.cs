using ProductDatabase.Data;
using ProductDatabase.Models;

namespace ProductDatabase.MasterManagement {
    public partial class PersonMasterEditDialog : Form {

        private readonly bool _isNewRecord;
        private PersonDef _person;

        public PersonMasterEditDialog(PersonDef? source) {
            InitializeComponent();
            _isNewRecord = source == null;
            _person = source != null
                ? new PersonDef {
                    PersonID = source!.PersonID,
                    PersonName = source!.PersonName,
                    IsActive = source!.IsActive
                }
                : new PersonDef { IsActive = 1 };
        }

        private void PersonMasterEditDialog_Load(object sender, EventArgs e) {
            this.Text = _isNewRecord ? "担当者追加" : "担当者編集";

            if (!_isNewRecord) {
                PersonNameTextBox.Text = _person.PersonName;
                IsActiveCheckBox.Checked = _person.IsActive != 0;
            }
            else {
                IsActiveCheckBox.Checked = true;
            }
        }

        // 保存ボタン：バリデーション → DB保存
        private void SaveButton_Click(object sender, EventArgs e) {
            if (!ValidateInput()) { return; }

            try {
                _person.PersonName = PersonNameTextBox.Text.Trim();
                _person.IsActive = IsActiveCheckBox.Checked ? 1 : 0;

                if (_isNewRecord) {
                    PersonRepository.Insert(_person);
                }
                else {
                    PersonRepository.Update(_person);
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // キャンセルボタン
        private void DialogCancelButton_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // 入力値のバリデーションを行い、エラーがあれば false を返す
        private bool ValidateInput() {
            if (string.IsNullOrWhiteSpace(PersonNameTextBox.Text)) {
                MessageBox.Show("名前を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PersonNameTextBox.Focus();
                return false;
            }

            // 名前重複チェック（編集時は自身を除外）
            long excludeId = _isNewRecord ? 0 : _person.PersonID;
            if (PersonRepository.ExistsName(PersonNameTextBox.Text.Trim(), excludeId)) {
                MessageBox.Show(
                    $"名前 [{PersonNameTextBox.Text.Trim()}] は既に登録されています。",
                    "名前重複", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                PersonNameTextBox.Focus();
                return false;
            }

            return true;
        }
    }
}
