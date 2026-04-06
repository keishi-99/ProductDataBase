using ProductDatabase.Data;
using ProductDatabase.Models;
using ProductDatabase.Print;
using System.Data;
using System.Text.Json;
using static ProductDatabase.Print.PrintManager;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase.MasterManagement {
    public partial class MasterManagementWindow : Form {

        private readonly ProductRepository _repository;
        private readonly AppSettings _appSettings;

        // 製品タブの検索フィルタ文字列
        private string _productSearchText = string.Empty;

        // 基板タブの検索フィルタ文字列
        private string _substrateSearchText = string.Empty;

        // 製品グリッドのソート状態
        private string _productSortColumn = "ProductID";
        private bool _productSortAscending = true;

        // 基板グリッドのソート状態
        private string _substrateSortColumn = "SubstrateID";
        private bool _substrateSortAscending = true;

        // 印刷設定のシリアライズ・デシリアライズで共有するオプション
        private static readonly JsonSerializerOptions s_jsonOptions = new() {
            PropertyNamingPolicy = null,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        // PrintSettingsWindow から参照される印刷設定（読み込み後に保持）
        public DocumentPrintSettings ProductPrintSettings { get; private set; } = new DocumentPrintSettings();

        // PrintSettingsWindow から参照される設定ファイルパス（選択行のDB保存済み値から生成）
        public string PrintSettingPath { get; private set; } = string.Empty;

        public MasterManagementWindow(ProductRepository repository, AppSettings appSettings) {
            InitializeComponent();
            _repository = repository;
            _appSettings = appSettings;
        }

        private void MasterManagementWindow_Load(object sender, EventArgs e) {
            try {
                LoadData();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // DBから全データを再読み込みして両グリッドを更新する
        private void LoadData() {
            _repository.LoadAll();
            BindProductGrid();
            BindSubstrateGrid();
        }

        // 製品グリッドをキャッシュ済みDataTableとフィルタ条件から再描画する
        private void BindProductGrid() {
            var query = _repository.ProductDataTable.AsEnumerable()
                .Where(FilterProduct)
                .Select(r => new {
                    ProductID = r.Field<long>("ProductID"),
                    CategoryName = r["CategoryName"]?.ToString() ?? string.Empty,
                    ProductName = r["ProductName"]?.ToString() ?? string.Empty,
                    ProductType = r["ProductType"]?.ToString() ?? string.Empty,
                    ProductModel = r["ProductModel"]?.ToString() ?? string.Empty,
                    RegType = (int)r.Field<long>("RegType"),
                    SerialPrintType = (int)(r.Field<long?>("SerialPrintType") ?? 0L),
                    SheetPrintType = (int)(r.Field<long?>("SheetPrintType") ?? 0L),
                    SerialType = (int)(r.Field<long?>("SerialType") ?? 0L),
                    Visible = (int)(r.Field<long?>("Visible") ?? 0L)
                });

            var rows = (_productSortColumn, _productSortAscending) switch {
                ("ProductID", true) => query.OrderBy(r => r.ProductID),
                ("ProductID", false) => query.OrderByDescending(r => r.ProductID),
                ("CategoryName", true) => query.OrderBy(r => r.CategoryName),
                ("CategoryName", false) => query.OrderByDescending(r => r.CategoryName),
                ("ProductName", true) => query.OrderBy(r => r.ProductName),
                ("ProductName", false) => query.OrderByDescending(r => r.ProductName),
                ("ProductType", true) => query.OrderBy(r => r.ProductType),
                ("ProductType", false) => query.OrderByDescending(r => r.ProductType),
                ("ProductModel", true) => query.OrderBy(r => r.ProductModel),
                ("ProductModel", false) => query.OrderByDescending(r => r.ProductModel),
                ("RegType", true) => query.OrderBy(r => r.RegType),
                ("RegType", false) => query.OrderByDescending(r => r.RegType),
                ("SerialPrintType", true) => query.OrderBy(r => r.SerialPrintType),
                ("SerialPrintType", false) => query.OrderByDescending(r => r.SerialPrintType),
                ("SheetPrintType", true) => query.OrderBy(r => r.SheetPrintType),
                ("SheetPrintType", false) => query.OrderByDescending(r => r.SheetPrintType),
                ("SerialType", true) => query.OrderBy(r => r.SerialType),
                ("SerialType", false) => query.OrderByDescending(r => r.SerialType),
                ("Visible", true) => query.OrderBy(r => r.Visible),
                ("Visible", false) => query.OrderByDescending(r => r.Visible),
                _ => query.OrderBy(r => r.ProductID)
            };

            ProductDataGridView.DataSource = rows.ToList();

            if (ProductDataGridView.Columns.Count > 0) {
                SetProductGridColumnHeaders();
                ApplySortGlyph(ProductDataGridView, _productSortColumn, _productSortAscending);
            }
        }

        // 基板グリッドをキャッシュ済みDataTableとフィルタ条件から再描画する
        private void BindSubstrateGrid() {
            var query = _repository.SubstrateDataTable.AsEnumerable()
                .Where(FilterSubstrate)
                .Select(r => new {
                    SubstrateID = r.Field<long>("SubstrateID"),
                    CategoryName = r["CategoryName"]?.ToString() ?? string.Empty,
                    ProductName = r["ProductName"]?.ToString() ?? string.Empty,
                    SubstrateName = r["SubstrateName"]?.ToString() ?? string.Empty,
                    SubstrateModel = r["SubstrateModel"]?.ToString() ?? string.Empty,
                    RegType = (int)r.Field<long>("RegType"),
                    Visible = (int)(r.Field<long?>("Visible") ?? 0L)
                });

            var rows = (_substrateSortColumn, _substrateSortAscending) switch {
                ("SubstrateID", true) => query.OrderBy(r => r.SubstrateID),
                ("SubstrateID", false) => query.OrderByDescending(r => r.SubstrateID),
                ("CategoryName", true) => query.OrderBy(r => r.CategoryName),
                ("CategoryName", false) => query.OrderByDescending(r => r.CategoryName),
                ("ProductName", true) => query.OrderBy(r => r.ProductName),
                ("ProductName", false) => query.OrderByDescending(r => r.ProductName),
                ("SubstrateName", true) => query.OrderBy(r => r.SubstrateName),
                ("SubstrateName", false) => query.OrderByDescending(r => r.SubstrateName),
                ("SubstrateModel", true) => query.OrderBy(r => r.SubstrateModel),
                ("SubstrateModel", false) => query.OrderByDescending(r => r.SubstrateModel),
                ("RegType", true) => query.OrderBy(r => r.RegType),
                ("RegType", false) => query.OrderByDescending(r => r.RegType),
                ("Visible", true) => query.OrderBy(r => r.Visible),
                ("Visible", false) => query.OrderByDescending(r => r.Visible),
                _ => query.OrderBy(r => r.SubstrateID)
            };

            SubstrateDataGridView.DataSource = rows.ToList();

            if (SubstrateDataGridView.Columns.Count > 0) {
                SetSubstrateGridColumnHeaders();
                ApplySortGlyph(SubstrateDataGridView, _substrateSortColumn, _substrateSortAscending);
            }
        }

        // 製品グリッドの列ヘッダーを日本語に設定する
        private void SetProductGridColumnHeaders() {
            var headers = new Dictionary<string, string> {
                { "ProductID",    "ID" },
                { "CategoryName", "カテゴリ名" },
                { "ProductName",  "製品名" },
                { "ProductType",  "製品タイプ" },
                { "ProductModel", "製品型式" },
                { "RegType",         "登録タイプ" },
                { "SerialPrintType", "シリアル印字タイプ" },
                { "SheetPrintType",  "シート印字タイプ" },
                { "SerialType",      "シリアル桁数タイプ" },
                { "Visible",         "表示" }
            };
            foreach (DataGridViewColumn col in ProductDataGridView.Columns) {
                if (headers.TryGetValue(col.Name, out var header)) {
                    col.HeaderText = header;
                }
            }
        }

        // 基板グリッドの列ヘッダーを日本語に設定する
        private void SetSubstrateGridColumnHeaders() {
            var headers = new Dictionary<string, string> {
                { "SubstrateID",    "ID" },
                { "CategoryName",   "カテゴリ名" },
                { "ProductName",    "製品名" },
                { "SubstrateName",  "基板名称" },
                { "SubstrateModel", "基板型式" },
                { "RegType",        "登録タイプ" },
                { "Visible",      "表示" }
            };
            foreach (DataGridViewColumn col in SubstrateDataGridView.Columns) {
                if (headers.TryGetValue(col.Name, out var header)) {
                    col.HeaderText = header;
                }
            }
        }

        // 製品行がフィルタ条件（製品名・製品型式）に一致するか判定する
        private bool FilterProduct(DataRow r) {
            if (string.IsNullOrWhiteSpace(_productSearchText)) { return true; }
            var name = r["ProductName"]?.ToString() ?? string.Empty;
            var model = r["ProductModel"]?.ToString() ?? string.Empty;
            return name.Contains(_productSearchText, StringComparison.OrdinalIgnoreCase) ||
                   model.Contains(_productSearchText, StringComparison.OrdinalIgnoreCase);
        }

        // 基板行がフィルタ条件（製品名・基板型式）に一致するか判定する
        private bool FilterSubstrate(DataRow r) {
            if (string.IsNullOrWhiteSpace(_substrateSearchText)) { return true; }
            var name = r["ProductName"]?.ToString() ?? string.Empty;
            var model = r["SubstrateModel"]?.ToString() ?? string.Empty;
            return name.Contains(_substrateSearchText, StringComparison.OrdinalIgnoreCase) ||
                   model.Contains(_substrateSearchText, StringComparison.OrdinalIgnoreCase);
        }

        // 製品を追加する
        private void ProductAddButton_Click(object sender, EventArgs e) {
            using var dialog = new ProductMasterEditDialog(_repository, null, _appSettings);
            if (dialog.ShowDialog(this) == DialogResult.OK) {
                LoadData();
            }
        }

        // 選択中の製品を編集する
        private void ProductEditButton_Click(object sender, EventArgs e) {
            if (ProductDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("編集する製品を選択してください。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var productId = (long)ProductDataGridView.SelectedRows[0].Cells["ProductID"].Value;
            var row = _repository.ProductDataTable.AsEnumerable()
                .FirstOrDefault(r => r.Field<long>("ProductID") == productId);

            if (row == null) {
                MessageBox.Show("製品データが見つかりません。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var dialog = new ProductMasterEditDialog(_repository, row, _appSettings);
            if (dialog.ShowDialog(this) == DialogResult.OK) {
                LoadData();
            }
        }

        // 選択中の製品を削除する
        private void ProductDeleteButton_Click(object sender, EventArgs e) {
            if (ProductDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("削除する製品を選択してください。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var productId = (long)ProductDataGridView.SelectedRows[0].Cells["ProductID"].Value;
            var productName = ProductDataGridView.SelectedRows[0].Cells["ProductName"].Value?.ToString() ?? string.Empty;
            var productModel = ProductDataGridView.SelectedRows[0].Cells["ProductModel"].Value?.ToString() ?? string.Empty;

            var confirm = MessageBox.Show(
                $"以下の製品マスターを削除します。よろしいですか？\n\n製品名: {productName}\n製品型式: {productModel}",
                "削除確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) { return; }

            try {
                ProductRepository.DeleteProduct(productId);
                LoadData();
                MessageBox.Show("削除しました。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (InvalidOperationException ex) {
                MessageBox.Show(ex.Message, "削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 基板を追加する
        private void SubstrateAddButton_Click(object sender, EventArgs e) {
            using var dialog = new SubstrateMasterEditDialog(_repository, null);
            if (dialog.ShowDialog(this) == DialogResult.OK) {
                LoadData();
            }
        }

        // 選択中の基板を編集する
        private void SubstrateEditButton_Click(object sender, EventArgs e) {
            if (SubstrateDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("編集する基板を選択してください。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var substrateId = (long)SubstrateDataGridView.SelectedRows[0].Cells["SubstrateID"].Value;
            var row = _repository.SubstrateDataTable.AsEnumerable()
                .FirstOrDefault(r => r.Field<long>("SubstrateID") == substrateId);

            if (row == null) {
                MessageBox.Show("基板データが見つかりません。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using var dialog = new SubstrateMasterEditDialog(_repository, row);
            if (dialog.ShowDialog(this) == DialogResult.OK) {
                LoadData();
            }
        }

        // 選択中の基板を削除する
        private void SubstrateDeleteButton_Click(object sender, EventArgs e) {
            if (SubstrateDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("削除する基板を選択してください。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var substrateId = (long)SubstrateDataGridView.SelectedRows[0].Cells["SubstrateID"].Value;
            var substrateName = SubstrateDataGridView.SelectedRows[0].Cells["SubstrateName"].Value?.ToString() ?? string.Empty;
            var substrateModel = SubstrateDataGridView.SelectedRows[0].Cells["SubstrateModel"].Value?.ToString() ?? string.Empty;

            var confirm = MessageBox.Show(
                $"以下の基板マスターを削除します。よろしいですか？\n\n基板名称: {substrateName}\n基板型式: {substrateModel}",
                "削除確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) { return; }

            try {
                ProductRepository.DeleteSubstrate(substrateId);
                LoadData();
                MessageBox.Show("削除しました。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (InvalidOperationException ex) {
                MessageBox.Show(ex.Message, "削除できません", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 製品検索テキスト変更時にグリッドを再フィルタする
        private void ProductSearchBox_TextChanged(object sender, EventArgs e) {
            _productSearchText = ProductSearchBox.Text;
            BindProductGrid();
        }

        // 基板検索テキスト変更時にグリッドを再フィルタする
        private void SubstrateSearchBox_TextChanged(object sender, EventArgs e) {
            _substrateSearchText = SubstrateSearchBox.Text;
            BindSubstrateGrid();
        }

        // 製品グリッドの列ヘッダークリック時にソートを切り替える
        private void ProductDataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e) {
            var columnName = ProductDataGridView.Columns[e.ColumnIndex].Name;
            if (_productSortColumn == columnName) {
                _productSortAscending = !_productSortAscending;
            }
            else {
                _productSortColumn = columnName;
                _productSortAscending = true;
            }
            BindProductGrid();
        }

        // 基板グリッドの列ヘッダークリック時にソートを切り替える
        private void SubstrateDataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e) {
            var columnName = SubstrateDataGridView.Columns[e.ColumnIndex].Name;
            if (_substrateSortColumn == columnName) {
                _substrateSortAscending = !_substrateSortAscending;
            }
            else {
                _substrateSortColumn = columnName;
                _substrateSortAscending = true;
            }
            BindSubstrateGrid();
        }

        // 印刷詳細設定ボタン：ContextMenuStrip をボタン直下に表示する
        private void ProductPrintDetailSettingsButton_Click(object sender, EventArgs e)
            => ProductPrintDetailContextMenuStrip.Show(
                ProductPrintDetailSettingsButton,
                new Point(0, ProductPrintDetailSettingsButton.Height));

        private void ProductLabelSettingsMenuItem_Click(object sender, EventArgs e)
            => OpenPrintSettings(SerialType.Label);

        private void ProductBarcodeSettingsMenuItem_Click(object sender, EventArgs e)
            => OpenPrintSettings(SerialType.Barcode);

        private void ProductNameplateSettingsMenuItem_Click(object sender, EventArgs e)
            => OpenPrintSettings(SerialType.Nameplate);

        // 選択中の製品の印刷設定ウィンドウを開く
        private void OpenPrintSettings(SerialType serialType) {
            if (ProductDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("印刷設定を開く製品を選択してください。", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var row = ProductDataGridView.SelectedRows[0];
            var categoryName = row.Cells["CategoryName"].Value?.ToString() ?? string.Empty;
            var productName  = row.Cells["ProductName"].Value?.ToString() ?? string.Empty;
            var productModel = row.Cells["ProductModel"].Value?.ToString() ?? string.Empty;
            var productId    = (long)row.Cells["ProductID"].Value;

            // DB保存済みの値からパスを生成するためUIの未保存変更に影響されない
            PrintSettingPath = Path.Combine(
                Environment.CurrentDirectory, "config", "Product",
                categoryName, productName,
                $"PrintConfig_{productName}_{productId}_{productModel}.json");

            if (!File.Exists(PrintSettingPath)) {
                MessageBox.Show("印刷設定ファイルが見つかりませんでした。",
                    "印刷設定", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try {
                var json = File.ReadAllText(PrintSettingPath);
                ProductPrintSettings = JsonSerializer.Deserialize<DocumentPrintSettings>(json, s_jsonOptions) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show($"印刷設定の読み込みに失敗しました。{Environment.NewLine}{ex.Message}",
                    $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            CurrentSerialType = serialType;
            using PrintSettingsWindow window = new() { AppSettings = _appSettings };
            window.ShowDialog(this);
        }

        // ソートグリフ（↑↓）をグリッドのヘッダーに反映する
        private static void ApplySortGlyph(DataGridView grid, string sortColumn, bool ascending) {
            foreach (DataGridViewColumn col in grid.Columns) {
                col.SortMode = DataGridViewColumnSortMode.Programmatic;
                col.HeaderCell.SortGlyphDirection = col.Name == sortColumn
                    ? (ascending ? SortOrder.Ascending : SortOrder.Descending)
                    : SortOrder.None;
            }
        }
    }
}
