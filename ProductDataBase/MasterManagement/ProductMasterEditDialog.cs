using ProductDatabase.Data;
using ProductDatabase.Models;
using System.Data;

namespace ProductDatabase.MasterManagement {
    public partial class ProductMasterEditDialog : Form {

        private readonly ProductRepository _repository;
        private readonly bool _isNewRecord;

        // 編集中の製品マスターデータ
        private readonly ProductMaster _product;

        // 編集対象の元データ行（新規追加時はnull）
        private readonly DataRow? _sourceRow;

        public ProductMasterEditDialog(ProductRepository repository, DataRow? sourceRow) {
            InitializeComponent();
            _repository = repository;
            _sourceRow = sourceRow;
            _isNewRecord = sourceRow == null;
            _product = new ProductMaster();
        }

        private void ProductMasterEditDialog_Load(object sender, EventArgs e) {
            InitializeComboBoxes();

            if (!_isNewRecord && _sourceRow != null) {
                _product.LoadFrom(_sourceRow);
                PopulateFields();
            }
            else {
                // 新規追加時のデフォルト値
                VisibleCheckBox.Checked = true;
            }

            // フィールド反映後に製品名でフィルタしてリストを構築する
            LoadSubstrateCheckedList();

            // 新規追加時：製品名が変わったらリストを再フィルタする
            ProductNameTextBox.TextChanged += (s, e) => LoadSubstrateCheckedList();

            this.Text = _isNewRecord ? "製品マスター追加" : "製品マスター編集";
        }

        // RegTypeとSerialTypeのComboBoxに選択肢を設定する
        private void InitializeComboBoxes() {
            RegTypeComboBox.Items.Clear();
            RegTypeComboBox.Items.Add(new ComboItem(0, "0 - 登録なし"));
            RegTypeComboBox.Items.Add(new ComboItem(1, "1 - シリアル登録のみ"));
            RegTypeComboBox.Items.Add(new ComboItem(2, "2 - 使用基板登録(先入れ先出し)"));
            RegTypeComboBox.Items.Add(new ComboItem(3, "3 - 使用基板登録(注番割当)"));
            RegTypeComboBox.Items.Add(new ComboItem(4, "4 - 使用基板登録(シリアル登録無)"));
            RegTypeComboBox.Items.Add(new ComboItem(9, "9 - サービス向け用"));
            RegTypeComboBox.SelectedIndex = 0;

            SerialTypeComboBox.Items.Clear();
            SerialTypeComboBox.Items.Add(new ComboItem(3, "3  - 3桁 (001〜999)"));
            SerialTypeComboBox.Items.Add(new ComboItem(4, "4  - 4桁 (0001〜9999)"));
            SerialTypeComboBox.Items.Add(new ComboItem(101, "101 - 3桁 (001〜899)"));
            SerialTypeComboBox.Items.Add(new ComboItem(102, "102 - 3桁 (901〜999)"));
            SerialTypeComboBox.SelectedIndex = 0;
        }

        // 製品名が一致する基板マスターをCheckedListBoxに読み込み、既存紐づけをチェック状態にする
        private void LoadSubstrateCheckedList() {
            SubstrateCheckedListBox.Items.Clear();

            var filterProductName = ProductNameTextBox.Text.Trim();

            // 既存の紐づき基板IDを取得（編集時のみ）
            var linkedSubstrateIds = new HashSet<int>();
            if (!_isNewRecord && _sourceRow != null) {
                var productId = Convert.ToInt32(_sourceRow["ProductID"]);
                linkedSubstrateIds = _repository.ProductUseSubstrate.AsEnumerable()
                    .Where(r => r["P_ProductID"] != DBNull.Value
                             && Convert.ToInt32(r["P_ProductID"]) == productId
                             && r["S_SubstrateID"] != DBNull.Value)
                    .Select(r => Convert.ToInt32(r["S_SubstrateID"]))
                    .ToHashSet();
            }

            // 製品名が一致する基板のみ追加し、紐づき状態のものはチェック済みにする
            foreach (DataRow row in _repository.SubstrateDataTable.Rows) {
                var substrateProductName = row["ProductName"]?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(filterProductName) &&
                    !substrateProductName.Equals(filterProductName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                var substrateId = Convert.ToInt32(row["SubstrateID"]);
                var substrateName = row["SubstrateName"]?.ToString() ?? string.Empty;
                var substrateModel = row["SubstrateModel"]?.ToString() ?? string.Empty;
                var item = new ListItem<int> {
                    Id = substrateId,
                    Name = $"{substrateName} [{substrateModel}]"
                };
                var isChecked = linkedSubstrateIds.Contains(substrateId);
                SubstrateCheckedListBox.Items.Add(item, isChecked);
            }
        }

        // モデルの値をフォームコントロールに反映する
        private void PopulateFields() {
            CategoryNameTextBox.Text = _product.CategoryName;
            ProductNameTextBox.Text = _product.ProductName;
            ProductModelTextBox.Text = _product.ProductModel;
            ProductTypeTextBox.Text = _product.ProductType;
            InitialTextBox.Text = _product.Initial;
            VisibleCheckBox.Checked = _product.Visible;

            // RegType ComboBox
            SelectComboByValue(RegTypeComboBox, _product.RegType);

            // SerialType ComboBox
            SelectComboByValue(SerialTypeComboBox, _product.SerialDigitType);

            // RevisionGroup
            RevisionGroupNumericUpDown.Value = Math.Clamp(_product.RevisionGroup, 0, 99);

            // SerialPrintType チェックボックス（ビットフラグ）
            var spf = (SerialPrintTypeFlags)_product.SerialPrintType;
            LabelPrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Label);
            BarcodePrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Barcode);
            NameplatePrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Nameplate);
            UnderlinePrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Underline);
            Last4DigitsPrintCheckBox.Checked = spf.HasFlag(SerialPrintTypeFlags.Last4Digits);

            // SheetPrintType チェックボックス（ビットフラグ）
            var shf = (SheetPrintTypeFlags)_product.SheetPrintType;
            CheckSheetPrintCheckBox.Checked = shf.HasFlag(SheetPrintTypeFlags.CheckSheet);
            ListPrintCheckBox.Checked = shf.HasFlag(SheetPrintTypeFlags.List);

            // CheckBin チェックボックス（11ビット）
            CheckBinCheckBox0.Checked = (_product.CheckBin & (1 << 0)) != 0;  // 注文番号
            CheckBinCheckBox1.Checked = (_product.CheckBin & (1 << 1)) != 0;  // 製造番号
            CheckBinCheckBox2.Checked = (_product.CheckBin & (1 << 2)) != 0;  // 数量
            CheckBinCheckBox3.Checked = (_product.CheckBin & (1 << 3)) != 0;  // 最初のシリアル番号
            CheckBinCheckBox4.Checked = (_product.CheckBin & (1 << 4)) != 0;  // リビジョン
            CheckBinCheckBox5.Checked = (_product.CheckBin & (1 << 5)) != 0;  // 予備1
            CheckBinCheckBox6.Checked = (_product.CheckBin & (1 << 6)) != 0;  // OLes番号
            CheckBinCheckBox7.Checked = (_product.CheckBin & (1 << 7)) != 0;  // 予備3
            CheckBinCheckBox8.Checked = (_product.CheckBin & (1 << 8)) != 0;  // 登録日
            CheckBinCheckBox9.Checked = (_product.CheckBin & (1 << 9)) != 0;  // 担当者
            CheckBinCheckBox10.Checked = (_product.CheckBin & (1 << 10)) != 0;  // コメント
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

        // 保存ボタン：バリデーション → DB保存 → 使用基板更新
        private void SaveButton_Click(object sender, EventArgs e) {
            if (!ValidateInput()) { return; }

            try {
                CollectFieldValues();

                int productId;
                if (_isNewRecord) {
                    productId = _repository.InsertProduct(_product);
                }
                else {
                    productId = _product.ProductID;
                    _repository.UpdateProduct(_product);
                }

                // 使用基板の紐づけ更新
                var selectedSubstrateIds = SubstrateCheckedListBox.CheckedItems
                    .Cast<ListItem<int>>()
                    .Select(item => item.Id)
                    .ToList();
                _repository.UpdateProductUseSubstrates(productId, selectedSubstrateIds);

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
            if (string.IsNullOrWhiteSpace(ProductModelTextBox.Text)) {
                MessageBox.Show("製品型式を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ProductModelTextBox.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(ProductTypeTextBox.Text)) {
                MessageBox.Show("製品タイプを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ProductTypeTextBox.Focus();
                return false;
            }
            // 製品型式の重複チェック
            int excludeId = _isNewRecord ? 0 : _product.ProductID;
            if (_repository.ExistsProductModel(ProductModelTextBox.Text.Trim(), excludeId)) {
                MessageBox.Show(
                    $"製品型式 [{ProductModelTextBox.Text.Trim()}] は既に登録されています。",
                    "重複エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ProductModelTextBox.Focus();
                return false;
            }

            return true;
        }

        // フォームコントロールの値をProductMasterオブジェクトに収集する
        private void CollectFieldValues() {
            _product.CategoryName = CategoryNameTextBox.Text.Trim();
            _product.ProductName = ProductNameTextBox.Text.Trim();
            _product.ProductModel = ProductModelTextBox.Text.Trim();
            _product.ProductType = ProductTypeTextBox.Text.Trim();
            _product.Initial = InitialTextBox.Text.Trim();
            _product.Visible = VisibleCheckBox.Checked;

            // RegType
            _product.RegType = RegTypeComboBox.SelectedItem is ComboItem regItem ? regItem.Value : 0;

            // SerialType（SerialDigitType に格納）
            _product.SerialDigitType = SerialTypeComboBox.SelectedItem is ComboItem serialItem ? serialItem.Value : 3;

            // RevisionGroup
            _product.RevisionGroup = (int)RevisionGroupNumericUpDown.Value;

            // SerialPrintType（ビットフラグ）
            int spf = 0;
            if (LabelPrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Label;
            if (BarcodePrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Barcode;
            if (NameplatePrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Nameplate;
            if (UnderlinePrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Underline;
            if (Last4DigitsPrintCheckBox.Checked) spf |= (int)SerialPrintTypeFlags.Last4Digits;
            _product.SerialPrintType = spf;

            // SheetPrintType（ビットフラグ）
            int shf = 0;
            if (CheckSheetPrintCheckBox.Checked) shf |= (int)SheetPrintTypeFlags.CheckSheet;
            if (ListPrintCheckBox.Checked) shf |= (int)SheetPrintTypeFlags.List;
            _product.SheetPrintType = shf;

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
            _product.CheckBin = checkBin;
        }

        // ComboBoxのアイテムを保持する内部クラス
        private sealed class ComboItem(int value, string label) {
            public int Value { get; } = value;
            public string Label { get; } = label;

            public override string ToString() => Label;
        }
    }
}
