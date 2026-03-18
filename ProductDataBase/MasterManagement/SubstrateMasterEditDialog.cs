using ProductDatabase.Data;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.MasterManagement {
    public partial class SubstrateMasterEditDialog : Form {

        private readonly ProductRepository _repository;
        private readonly bool _isNewRecord;
        private readonly SubstrateMaster _substrate;
        private readonly DataRow? _sourceRow;

        public SubstrateMasterEditDialog(ProductRepository repository, DataRow? sourceRow) {
            InitializeComponent();
            _repository = repository;
            _sourceRow = sourceRow;
            _isNewRecord = sourceRow == null;
            _substrate = new SubstrateMaster();
        }

        private void SubstrateMasterEditDialog_Load(object sender, EventArgs e) {
            InitializeComboBoxes();

            if (!_isNewRecord && _sourceRow != null) {
                _substrate.LoadFrom(_sourceRow);
                PopulateFields();
            } else {
                // 新規追加時のデフォルト値
                VisibleCheckBox.Checked = true;
            }

            this.Text = _isNewRecord ? "基板マスター追加" : "基板マスター編集";
        }

        // RegTypeのComboBoxに選択肢を設定する
        private void InitializeComboBoxes() {
            RegTypeComboBox.Items.Clear();
            RegTypeComboBox.Items.Add(new ComboItem(0, "0 - 登録無"));
            RegTypeComboBox.Items.Add(new ComboItem(1, "1 - 登録有"));
            RegTypeComboBox.SelectedIndex = 0;
        }

        // モデルの値をフォームコントロールに反映する
        private void PopulateFields() {
            CategoryNameTextBox.Text = _substrate.CategoryName;
            ProductNameTextBox.Text = _substrate.ProductName;
            SubstrateNameTextBox.Text = _substrate.SubstrateName;
            SubstrateModelTextBox.Text = _substrate.SubstrateModel;
            VisibleCheckBox.Checked = _substrate.Visible;

            SelectComboByValue(RegTypeComboBox, _substrate.RegType);

            // SerialPrintType チェックボックス
            var spf = (SerialPrintTypeFlags)_substrate.SerialPrintType;
            LabelPrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Label);

            // CheckBin チェックボックス（11ビット）
            CheckBinCheckBox0.Checked = (_substrate.CheckBin & (1 << 0)) != 0;
            CheckBinCheckBox1.Checked = (_substrate.CheckBin & (1 << 1)) != 0;
            CheckBinCheckBox2.Checked = (_substrate.CheckBin & (1 << 2)) != 0;
            CheckBinCheckBox3.Checked = (_substrate.CheckBin & (1 << 3)) != 0;
            CheckBinCheckBox4.Checked = (_substrate.CheckBin & (1 << 4)) != 0;
            CheckBinCheckBox5.Checked = (_substrate.CheckBin & (1 << 5)) != 0;
            CheckBinCheckBox6.Checked = (_substrate.CheckBin & (1 << 6)) != 0;
            CheckBinCheckBox7.Checked = (_substrate.CheckBin & (1 << 7)) != 0;
            CheckBinCheckBox8.Checked = (_substrate.CheckBin & (1 << 8)) != 0;
            CheckBinCheckBox9.Checked = (_substrate.CheckBin & (1 << 9)) != 0;
            CheckBinCheckBox10.Checked = (_substrate.CheckBin & (1 << 10)) != 0;
        }

        // ComboBoxから値に一致するアイテムを選択する
        private static void SelectComboByValue(ComboBox comboBox, int value) {
            for (int i = 0; i < comboBox.Items.Count; i++) {
                if (comboBox.Items[i] is ComboItem item && item.Value == value) {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }
            if (comboBox.Items.Count > 0) {
                comboBox.SelectedIndex = 0;
            }
        }

        // 保存ボタン：バリデーション → DB保存
        private void SaveButton_Click(object sender, EventArgs e) {
            if (!ValidateInput()) { return; }

            try {
                CollectFieldValues();

                if (_isNewRecord) {
                    _repository.InsertSubstrate(_substrate);
                }
                else {
                    _repository.UpdateSubstrate(_substrate);
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // キャンセルボタン
        private void CancelButton_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // 入力値のバリデーションを行い、エラーがあれば false を返す
        private bool ValidateInput() {
            if (string.IsNullOrWhiteSpace(CategoryNameTextBox.Text)) {
                MessageBox.Show("カテゴリ名を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                CategoryNameTextBox.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(ProductNameTextBox.Text)) {
                MessageBox.Show("製品名を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ProductNameTextBox.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(SubstrateNameTextBox.Text)) {
                MessageBox.Show("基板名称を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SubstrateNameTextBox.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(SubstrateModelTextBox.Text)) {
                MessageBox.Show("基板型式を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SubstrateModelTextBox.Focus();
                return false;
            }

            // 基板型式の重複チェック
            int excludeId = _isNewRecord ? 0 : _substrate.SubstrateID;
            if (_repository.ExistsSubstrateModel(SubstrateModelTextBox.Text.Trim(), excludeId)) {
                MessageBox.Show(
                    $"基板型式 [{SubstrateModelTextBox.Text.Trim()}] は既に登録されています。",
                    "重複エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SubstrateModelTextBox.Focus();
                return false;
            }

            return true;
        }

        // フォームコントロールの値をSubstrateMasterオブジェクトに収集する
        private void CollectFieldValues() {
            _substrate.CategoryName = CategoryNameTextBox.Text.Trim();
            _substrate.ProductName = ProductNameTextBox.Text.Trim();
            _substrate.SubstrateName = SubstrateNameTextBox.Text.Trim();
            _substrate.SubstrateModel = SubstrateModelTextBox.Text.Trim();

            _substrate.RegType = RegTypeComboBox.SelectedItem is ComboItem regItem ? regItem.Value : 0;
            _substrate.Visible = VisibleCheckBox.Checked;

            // SerialPrintType（ビットフラグ）
            int spf = 0;
            if (LabelPrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Label;
            _substrate.SerialPrintType = spf;

            // CheckBin（11ビット）
            int checkBin = 0;
            if (CheckBinCheckBox0.Checked) checkBin |= (1 << 0);
            if (CheckBinCheckBox1.Checked) checkBin |= (1 << 1);
            if (CheckBinCheckBox2.Checked) checkBin |= (1 << 2);
            if (CheckBinCheckBox3.Checked) checkBin |= (1 << 3);
            if (CheckBinCheckBox4.Checked) checkBin |= (1 << 4);
            if (CheckBinCheckBox5.Checked) checkBin |= (1 << 5);
            if (CheckBinCheckBox6.Checked) checkBin |= (1 << 6);
            if (CheckBinCheckBox7.Checked) checkBin |= (1 << 7);
            if (CheckBinCheckBox8.Checked) checkBin |= (1 << 8);
            if (CheckBinCheckBox9.Checked) checkBin |= (1 << 9);
            if (CheckBinCheckBox10.Checked) checkBin |= (1 << 10);
            _substrate.CheckBin = checkBin;
        }

        // ComboBoxのアイテムを保持する内部クラス
        private sealed class ComboItem(int value, string label) {
            public int Value { get; } = value;
            public string Label { get; } = label;

            public override string ToString() => Label;
        }
    }
}
