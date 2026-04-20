using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.Excel;
using ProductDatabase.Models;
using ProductDatabase.Other;
using ProductDatabase.Services;
using System.Data;

namespace ProductDatabase.History {
    public partial class HistoryWindow : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly SubstrateMaster _substrateMaster;
        private readonly SubstrateRegisterWork _substrateRegisterWork;
        private readonly AppSettings _appSettings;

        private readonly int _radioButtonNumber = 0;

        private System.Data.DataTable _historyTable = new();

        private readonly List<string> _listColFilter = [];
        private string _tableName = string.Empty;

        // DataGridView の既定の背景色（編集モード終了時に復元する）
        private readonly Color _originalGridBackColor;

        private readonly Dictionary<string, Dictionary<string, string>> _headerTextMap = new() {
                {
                    "Substrate", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "SubstrateID", "基板ID" },
                        { "ProductName", "製品名" },
                        { "SubstrateName", "基板名" },
                        { "SubstrateModel", "基板型式" },
                        { "SubstrateNumber", "製造番号" },
                        { "OrderNumber", "注文番号" },
                        { "Increase", "追加数" },
                        { "Decrease", "使用数" },
                        { "Defect", "減少数" },
                        { "ProductType", "使用製品名" },
                        { "ProductNumber", "使用製番" },
                        { "OrderNumber1", "使用注番" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Comment", "コメント" },
                        { "usedID", "UsedID" },
                        { "CreatedAt", "作成日" },
                    }
                },
                {
                    "SubstrateStock", new Dictionary<string, string>
                    {
                        { "ProductName", "製品名" },
                        { "SubstrateID", "基板ID" },
                        { "SubstrateName", "基板名" },
                        { "SubstrateModel", "基板型式" },
                        { "SubstrateNumber", "製造番号" },
                        { "OrderNumber", "注文番号" },
                        { "Stock", "残数" }
                    }
                },
                {
                    "Product", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "ProductID", "製品ID" },
                        { "CategoryName", "カテゴリ" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "OLesNumber", "OLes番号" },
                        { "ProductName", "製品名" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "Quantity", "数量" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Revision", "Rev" },
                        { "RevisionGroup", "Group" },
                        { "SerialFirst", "シリアル先頭" },
                        { "SerialLast", "シリアル末尾" },
                        { "SerialLastNumber", "末番" },
                        { "Comment", "コメント" },
                        { "CreatedAt", "作成日" },
                    }
                },
                {
                    "Serial", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "Serial", "シリアル" },
                        { "OLesSerial", "O-Lesシリアル" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "ProductName", "製品名" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "RegDate", "登録日" },
                        { "usedID", "UsedID" }
                    }
                },
                {
                    "Reprint", new Dictionary<string, string>
                    {
                        { "ID", "ID" },
                        { "SerialPrintType", "印刷対象" },
                        { "ProductName", "製品名" },
                        { "OrderNumber", "注文番号" },
                        { "ProductNumber", "製造番号" },
                        { "ProductType", "タイプ" },
                        { "ProductModel", "製品型式" },
                        { "Quantity", "数量" },
                        { "Person", "担当者" },
                        { "RegDate", "登録日" },
                        { "Revision", "Rev" },
                        { "SerialFirst", "シリアル先頭" },
                        { "SerialLast", "シリアル末尾" },
                        { "Comment", "コメント" },
                        { "CreatedAt", "作成日" },
                    }
                }
            };

        public HistoryWindow(ProductMaster productMaster, ProductRegisterWork productRegisterWork, SubstrateMaster substrateMaster, SubstrateRegisterWork substrateRegisterWork, int radioButtonNumber, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _substrateMaster = substrateMaster;
            _substrateRegisterWork = substrateRegisterWork;
            _radioButtonNumber = radioButtonNumber;
            _appSettings = appSettings;

            if (Screen.PrimaryScreen is not null) {
                var h = Screen.PrimaryScreen.Bounds.Height;
                var w = Screen.PrimaryScreen.Bounds.Width;
                DataBaseDataGridView.MaximumSize = new Size(w, h);
            }
            DataBaseDataGridView.AllowUserToAddRows = false;
            DataBaseDataGridView.AllowUserToDeleteRows = false;
            DataBaseDataGridView.AllowUserToResizeRows = false;
            DataBaseDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(DataBaseDataGridView.Font, FontStyle.Bold);
            DataBaseDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DataBaseDataGridView.ReadOnly = true;
            DataBaseDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DataBaseDataGridView.RowTemplate.DefaultCellStyle.Padding = new Padding(5);
            DataBaseDataGridView.RowTemplate.Height += 10;
            // 編集モード終了時に復元するために既定の背景色を記録する
            _originalGridBackColor = DataBaseDataGridView.BackgroundColor;
        }

        // ロード時に初期UIを設定しラジオボタンモードに応じた表示制御を行う
        private void LoadEvents() {
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);

            CategoryRadioButton1.Checked = true;
            CategoryComboBox.SelectedIndex = 0;

            // 管理者のみ「編集」メニューを表示する
            編集ToolStripMenuItem.Visible = _appSettings.IsAdministrator;

            switch (_radioButtonNumber) {
                case 1:
                    CategoryRadioButton2.Text = "在庫";
                    CategoryRadioButton3.Visible = false;
                    StockCheckBox.Visible = false;
                    AllSubstrateCheckBox.Visible = true;
                    GroupModelCheckBox.Visible = false;
                    ShowUsedSubstrateButton.Visible = false;
                    GenerateReportButton.Visible = false;
                    GenerateListButton.Visible = _productMaster.IsListPrint;
                    GenerateCheckSheetButton.Visible = _productMaster.IsCheckSheetPrint;
                    if (string.IsNullOrEmpty(_substrateMaster.SubstrateModel)) { AllSubstrateCheckBox.Checked = true; }
                    break;
                case 2:
                    CategoryRadioButton2.Text = "全てのタイプ";
                    CategoryRadioButton3.Text = "シリアル";
                    StockCheckBox.Visible = false;
                    AllSubstrateCheckBox.Visible = false;
                    GroupModelCheckBox.Visible = false;
                    break;
                case 3:
                    CategoryRadioButton1.Visible = false;
                    CategoryRadioButton2.Visible = false;
                    CategoryRadioButton3.Visible = false;
                    StockCheckBox.Visible = false;
                    AllSubstrateCheckBox.Visible = false;
                    GroupModelCheckBox.Visible = false;
                    ShowUsedSubstrateButton.Visible = false;
                    GenerateReportButton.Visible = false;
                    GenerateListButton.Visible = false;
                    GenerateCheckSheetButton.Visible = false;
                    // 再印刷履歴は編集対象外なので「編集」メニューを非表示にする
                    編集ToolStripMenuItem.Visible = false;
                    break;
            }
        }

        // DataTableをDataGridViewにバインドしカラムヘッダーを日本語に変換する
        private void LoadDataAndDisplay(string categoryName, DataTable data) {
            _historyTable = data;

            DataBaseDataGridView.Columns.Clear();
            DataBaseDataGridView.DataSource = _historyTable;

            _listColFilter.Clear();
            _listColFilter.Add("");

            for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                _listColFilter.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty);
            }

            if (_headerTextMap.TryGetValue(categoryName, out var columnHeaders)) {
                foreach (var column in DataBaseDataGridView.Columns.Cast<DataGridViewColumn>()) {
                    if (columnHeaders.TryGetValue(column.Name, out var headerText)) {
                        column.HeaderCell.Value = headerText;
                    }
                }

                CategoryComboBox.Items.Clear();
                CategoryComboBox.Items.Add("");
                for (var i = 0; i < DataBaseDataGridView.ColumnCount; i++) {
                    CategoryComboBox.Items.Add(DataBaseDataGridView.Columns[i].HeaderCell.Value?.ToString() ?? string.Empty);
                }
            }
            DataBaseDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        // ラジオボタン選択に応じて基板登録履歴または在庫サマリーを表示する
        private void ViewSubstrateRegistrationLog() {
            if (CategoryRadioButton1.Checked) {
                _tableName = "Substrate";
                在庫調整ToolStripMenuItem.Enabled = false;
                StockCheckBox.Visible = false;
                AllSubstrateCheckBox.Visible = true;
                GroupModelCheckBox.Visible = false;
                LoadDataAndDisplay("Substrate",
                    HistoryRepository.QuerySubstrateHistory(_substrateMaster, AllSubstrateCheckBox.Checked));
            }
            if (CategoryRadioButton2.Checked) {
                _tableName = string.Empty;
                在庫調整ToolStripMenuItem.Enabled = true;
                StockCheckBox.Visible = true;
                AllSubstrateCheckBox.Visible = true;
                GroupModelCheckBox.Visible = true;
                LoadDataAndDisplay("SubstrateStock",
                    HistoryRepository.QuerySubstrateStock(
                        _substrateMaster,
                        AllSubstrateCheckBox.Checked,
                        StockCheckBox.Checked,
                        GroupModelCheckBox.Checked));
            }
        }

        // フィルター条件に基づいて製品登録履歴を取得しDataGridViewに表示する
        private void ViewProductRegistration() {
            _tableName = "Product";
            GenerateReportButton.Visible = true;
            ShowUsedSubstrateButton.Visible = true;
            if (string.IsNullOrEmpty(_productMaster.ProductModel)) {
                GenerateListButton.Visible = true;
                GenerateCheckSheetButton.Visible = true;
                CategoryRadioButton1.Visible = false;
                CategoryRadioButton2.Checked = true;
            }
            else {
                GenerateListButton.Visible = _productMaster.IsListPrint;
                GenerateCheckSheetButton.Visible = _productMaster.IsCheckSheetPrint;
            }
            LoadDataAndDisplay("Product",
                HistoryRepository.QueryProductHistory(_productMaster, !CategoryRadioButton1.Checked));
        }

        // フィルター条件に基づいてシリアル番号履歴を取得しDataGridViewに表示する
        private void ViewSerialLog() {
            _tableName = "Serial";
            GenerateReportButton.Visible = false;
            ShowUsedSubstrateButton.Visible = false;
            GenerateListButton.Visible = false;
            GenerateCheckSheetButton.Visible = false;
            LoadDataAndDisplay("Serial",
                HistoryRepository.QuerySerialHistory(_productMaster, !CategoryRadioButton1.Checked));
        }

        // 製品型式でフィルタして再印刷履歴をDataGridViewに表示する
        private void ViewReprintLog() {
            LoadDataAndDisplay("Reprint",
                HistoryRepository.QueryReprintHistory(_productMaster.ProductModel));
        }

        // 選択行の製品履歴を編集ダイアログで編集してDBに保存する
        private async Task EditProductRecord() {
            if (DataBaseDataGridView.CurrentCell is null) { return; }
            var rowIndex = DataBaseDataGridView.CurrentCell.RowIndex;
            if (rowIndex < 0) { return; }

            var dgvRow = DataBaseDataGridView.Rows[rowIndex];

            using var dialog = new HistoryEditDialog(dgvRow);
            if (dialog.ShowDialog(this) != DialogResult.OK) { return; }

            // UIスレッドでDataRowの変更を記録（Task.Run前に実施）
            if (dgvRow.DataBoundItem is not DataRowView drv) { return; }
            var row = drv.Row;

            row.BeginEdit();
            row["OrderNumber"] = (object?)dialog.OrderNumber ?? DBNull.Value;
            row["ProductNumber"] = (object?)dialog.ProductNumber ?? DBNull.Value;
            row["OLesNumber"] = (object?)dialog.OLesNumber ?? DBNull.Value;
            row["Person"] = (object?)dialog.Person ?? DBNull.Value;
            row["Comment"] = (object?)dialog.Comment ?? DBNull.Value;
            row.EndEdit();

            List<string[]> pendingLogs = [];
            using var overlay = new LoadingOverlay(this);
            try {
                // DB操作をバックグラウンドスレッドで実行
                await Task.Run(() => {
                    using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                    con.Open();
                    using var tx = con.BeginTransaction();

                    // Original=変更前, Current=変更後
                    HistoryAuditLogger.LogProductEdit(row, pendingLogs, _productMaster.CategoryName);
                    HistoryRepository.UpdateProductRow(con, row, tx);

                    tx.Commit();
                    BackupManager.CreateBackup();

                    foreach (var log in pendingLogs) {
                        Logger.AppendLog(log);
                    }
                });

                RefreshCurrentView();

            } catch (Exception ex) {
                // エラー時はDataRow変更を取り消す（UIスレッドで実施）
                row.RejectChanges();
                MessageBox.Show(ex.Message, "編集エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 選択行をまとめて削除する（Substrate/Product/Serial対応、複数行選択可）
        private async Task DeleteSelectedRows() {
            // 選択行の収集（SelectedRowsが空の場合はCurrentCellの行を使用）
            var rowsToDelete = new List<DataRow>();
            var selectedRows = DataBaseDataGridView.SelectedRows;
            if (selectedRows.Count > 0) {
                foreach (DataGridViewRow dgvRow in selectedRows) {
                    if (dgvRow.DataBoundItem is DataRowView drv) { rowsToDelete.Add(drv.Row); }
                }
            }
            else {
                if (DataBaseDataGridView.CurrentCell is null) { return; }
                var idx = DataBaseDataGridView.CurrentCell.RowIndex;
                if (DataBaseDataGridView.Rows[idx].DataBoundItem is DataRowView drv) {
                    rowsToDelete.Add(drv.Row);
                }
            }
            if (rowsToDelete.Count == 0) { return; }

            var result = MessageBox.Show(
                $"{rowsToDelete.Count}件を削除しますか？\nこの操作は取り消せません。",
                "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes) { return; }

            // Task.Run でキャプチャするため変数をローカルにコピー
            var tableName = _tableName;
            List<string[]> pendingLogs = [];
            using var overlay = new LoadingOverlay(this);
            try {
                // DB操作をバックグラウンドスレッドで実行
                await Task.Run(() => {
                    using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());
                    con.Open();
                    using var tx = con.BeginTransaction();

                    switch (tableName) {
                        case "Substrate":
                            foreach (var row in rowsToDelete) {
                                HistoryAuditLogger.LogSubstrateDelete(row, pendingLogs, _substrateMaster.CategoryName);
                                HistoryRepository.DeleteSubstrateRow(con, row, tx);
                            }
                            break;
                        case "Product":
                            foreach (var row in rowsToDelete) {
                                var id = Convert.ToInt64(row["ID"]);
                                var substrates = HistoryRepository.GetSubstratesByUseId(con, id, tx);
                                var serials = HistoryRepository.GetSerialsByUsedId(con, id, tx);
                                HistoryAuditLogger.LogProductDelete(row, pendingLogs, _productMaster.CategoryName);
                                HistoryAuditLogger.LogProductSubstrateDelete(substrates, pendingLogs, _substrateMaster.CategoryName);
                                HistoryAuditLogger.LogProductSerialDelete(serials, pendingLogs, _productMaster.CategoryName);
                                HistoryRepository.DeleteProductRow(con, row, tx);
                                HistoryRepository.DeleteProductSubstrateRow(con, row, tx);
                                HistoryRepository.DeleteProductSerialRow(con, row, tx);
                            }
                            break;
                        case "Serial":
                            foreach (var row in rowsToDelete) {
                                HistoryAuditLogger.LogSerialDelete(row, pendingLogs, _productMaster.CategoryName);
                                HistoryRepository.DeleteSerialRow(con, row, tx);
                            }
                            break;
                    }

                    tx.Commit();
                    BackupManager.CreateBackup();

                    foreach (var log in pendingLogs) {
                        Logger.AppendLog(log);
                    }
                });

                RefreshCurrentView();

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "削除エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 現在の_tableNameに対応するViewメソッドを呼び出して表示を更新する
        private void RefreshCurrentView() {
            switch (_tableName) {
                case "Substrate":
                    ViewSubstrateRegistrationLog();
                    break;
                case "Product":
                    ViewProductRegistration();
                    break;
                case "Serial":
                    ViewSerialLog();
                    break;
            }
        }

        // ラジオボタンのTagとモード番号の組み合わせで表示する履歴の種類を切り替える
        private void CategorySelect(object sender) {
            var selectedRadioButton = (RadioButton)sender;
            if (!selectedRadioButton.Checked) { return; }
            var tag = selectedRadioButton.Tag?.ToString() ?? string.Empty;

            var actionMap = new Dictionary<(int, string), System.Action>
            {
                { (1, "1"), ViewSubstrateRegistrationLog },
                { (1, "2"), ViewSubstrateRegistrationLog },
                { (2, "1"), ViewProductRegistration },
                { (2, "2"), ViewProductRegistration },
                { (2, "3"), ViewSerialLog },
                { (3, "1"), ViewReprintLog }
            };

            if (actionMap.TryGetValue((_radioButtonNumber, tag), out var action)) {
                action();
            }
            else {
                MessageBox.Show("無効な選択です。", "[CategorySelect]エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 選択カラムとキーワードでDataGridViewのRowFilterを設定して表示を絞り込む
        private void HistoryTableFilter() {
            if (DataBaseDataGridView.DataSource is not System.Data.DataTable historyTable) {
                return;
            }

            if (string.IsNullOrEmpty(FilterStringTextBox.Text) || CategoryComboBox.SelectedIndex == -1) {
                historyTable.DefaultView.RowFilter = null;
            }
            else if (historyTable.Columns[_listColFilter[CategoryComboBox.SelectedIndex]]?.DataType == typeof(long)) {
                historyTable.DefaultView.RowFilter = $"CONVERT({_listColFilter[CategoryComboBox.SelectedIndex]}, 'System.String') LIKE '%{FilterStringTextBox.Text}%'";
            }
            else if (CategoryComboBox.Text != "") {
                historyTable.DefaultView.RowFilter = $"{_listColFilter[CategoryComboBox.SelectedIndex]} LIKE '*{FilterStringTextBox.Text}*'";
            }
        }

        // 選択行の基板情報を読み込み基板登録ウィンドウを開いて在庫を調整する
        private void InventoryAdjustment() {
            var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _substrateMaster.SubstrateID = int.TryParse(DataBaseDataGridView.Rows[i].Cells["SubstrateID"].Value?.ToString(), out var subId) ? subId : 0;
            _substrateRegisterWork.OrderNumber = DataBaseDataGridView.Rows[i].Cells["OrderNumber"].Value?.ToString() ?? string.Empty;
            _substrateRegisterWork.ProductNumber = DataBaseDataGridView.Rows[i].Cells["SubstrateNumber"].Value?.ToString() ?? string.Empty;

            var result = HistoryRepository.QuerySubstrateMasterById(_substrateMaster.SubstrateID);

            if (result != null) {
                _substrateMaster.SubstrateName = result.SubstrateName ?? string.Empty;
                _substrateMaster.SubstrateModel = result.SubstrateModel ?? string.Empty;
                _substrateMaster.ProductName = result.ProductName ?? string.Empty;
                _substrateMaster.RegType = int.TryParse(result.RegType?.ToString(), out int rt) ? rt : 0;
                _substrateMaster.CheckBin = Convert.ToInt32(result.Checkbox?.ToString() ?? "0", 2); // バイナリ文字列変換のため維持
                _substrateMaster.SerialPrintType = int.TryParse(result.SerialPrintType?.ToString(), out int spt) ? spt : 0;
            }

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
            LoadEvents();
        }

        // 選択行の製品IDに紐づく使用基板一覧をサブウィンドウに表示する
        private void ShowUsedSubstrateDetails() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }

            var dataForm = new Form {
                Text = "使用基板",
                ShowIcon = false,
                Width = 800,
                Height = 400,
                MaximizeBox = false,
                MinimizeBox = false,
                StartPosition = FormStartPosition.CenterScreen
            };

            var dataGridView = new DataGridView {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            };

            dataForm.Controls.Add(dataGridView);

            var i = DataBaseDataGridView.SelectedCells[0].RowIndex;
            var id = int.TryParse(DataBaseDataGridView.Rows[i].Cells["ID"].Value?.ToString(), out var idVal) ? idVal : 0;

            dataGridView.DataSource = HistoryRepository.QueryUsedSubstrates(id);
            dataForm.ShowDialog();
        }

        // 選択行の製品情報をセットしてExcel製造報告書を生成する
        private async Task GenerateReport() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.RowID = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value?.ToString(), out var rid1) ? rid1 : 0;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value?.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value?.ToString() ?? string.Empty;

            // [UIスレッド] 前処理: Config読み込み + ファイル選択 + 保存先選択
            (string templateFilePath, string savePath, ReportGeneratorClosedXml.ReportConfigClosedXml config)? prepared;
            try {
                prepared = ReportGeneratorClosedXml.PrepareReport(
                    _productMaster.ProductModel, _productRegisterWork.ProductNumber);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{nameof(GenerateReport)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (prepared is null) return; // キャンセル

            GenerateReportButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ReportGeneratorClosedXml.ExecuteReport(
                            _productMaster, _productRegisterWork,
                            prepared.Value.templateFilePath,
                            prepared.Value.savePath,
                            prepared.Value.config));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateReportButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateReport)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else {
                MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 選択行の製品情報をセットしてExcel製品一覧を生成する
        private async Task GenerateList() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.RowID = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["ID"].Value?.ToString(), out var rid2) ? rid2 : 0;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value?.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.Comment = DataBaseDataGridView.Rows[selectRow].Cells["Comment"].Value?.ToString() ?? string.Empty;

            GenerateListButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ListGeneratorClosedXml.GenerateList(
                            _productMaster, _productRegisterWork));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateListButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateList)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 選択行の製品情報をセットしてExcelチェックシートを生成する
        private async Task GenerateCheckSheet() {
            if (DataBaseDataGridView.CurrentCell is null && DataBaseDataGridView.SelectedCells.Count <= 0) { return; }
            var selectRow = DataBaseDataGridView.SelectedCells[0].RowIndex;
            _productRegisterWork.ProductNumber = DataBaseDataGridView.Rows[selectRow].Cells["ProductNumber"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.OrderNumber = DataBaseDataGridView.Rows[selectRow].Cells["OrderNumber"].Value?.ToString() ?? string.Empty;
            _productMaster.ProductModel = DataBaseDataGridView.Rows[selectRow].Cells["ProductModel"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.Quantity = int.TryParse(DataBaseDataGridView.Rows[selectRow].Cells["Quantity"].Value?.ToString(), out var quantity) ? quantity : 0;
            _productRegisterWork.RegDate = DataBaseDataGridView.Rows[selectRow].Cells["RegDate"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialFirst = DataBaseDataGridView.Rows[selectRow].Cells["SerialFirst"].Value?.ToString() ?? string.Empty;
            _productRegisterWork.SerialLast = DataBaseDataGridView.Rows[selectRow].Cells["SerialLast"].Value?.ToString() ?? string.Empty;

            // [UIスレッド] 前処理: Config読み込み + InputDialog（温度・湿度）
            (CheckSheetGeneratorClosedXml.CheckSheetConfigData configData, string temperature, string humidity)? prepared;
            try {
                prepared = CheckSheetGeneratorClosedXml.PrepareCheckSheet(_productMaster);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (prepared is null) return; // キャンセル

            GenerateCheckSheetButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                            _productMaster, _productRegisterWork,
                            prepared.Value.configData,
                            prepared.Value.temperature,
                            prepared.Value.humidity));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            GenerateCheckSheetButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 編集モードを開始する：ContextMenuStrip を有効化して枠線を赤に変更する
        private void EnterEditMode() {
            DataBaseDataGridView.ContextMenuStrip = EditContextMenuStrip;
            // DataGridView の背景色を変えて編集モード中であることを視覚的に示す
            DataBaseDataGridView.BackgroundColor = Color.MistyRose;
            DataBaseDataGridView.EnableHeadersVisualStyles = false;
            DataBaseDataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;
            編集開始ToolStripMenuItem.Visible = false;
            編集終了ToolStripMenuItem.Visible = true;
            SetBottomControlsEnabled(false);
        }

        // 編集モードを終了する：ContextMenuStrip を無効化して外観を元に戻す
        private void ExitEditMode() {
            DataBaseDataGridView.ContextMenuStrip = null;
            DataBaseDataGridView.BackgroundColor = _originalGridBackColor;
            DataBaseDataGridView.EnableHeadersVisualStyles = true;
            編集開始ToolStripMenuItem.Visible = true;
            編集終了ToolStripMenuItem.Visible = false;
            SetBottomControlsEnabled(true);
        }

        // 下部のボタン・チェックボックスの有効/無効を切り替える
        private void SetBottomControlsEnabled(bool enabled) {
            GenerateReportButton.Enabled = enabled;
            GenerateListButton.Enabled = enabled;
            GenerateCheckSheetButton.Enabled = enabled;
            ShowUsedSubstrateButton.Enabled = enabled;
            AllSubstrateCheckBox.Enabled = enabled;
            StockCheckBox.Enabled = enabled;
            GroupModelCheckBox.Enabled = enabled;
        }

        // 右クリックメニューを開く前にテーブル種別・行選択に応じて項目を制御する
        private void EditContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e) {
            // 右クリックした位置の行インデックスを取得する
            var clientPoint = DataBaseDataGridView.PointToClient(Cursor.Position);
            var hit = DataBaseDataGridView.HitTest(clientPoint.X, clientPoint.Y);

            // ヘッダー行や空白領域ではメニューを表示しない
            if (hit.RowIndex < 0) { e.Cancel = true; return; }

            // 右クリックした行が選択されていなければ単一選択に切り替える
            if (!DataBaseDataGridView.Rows[hit.RowIndex].Selected) {
                DataBaseDataGridView.CurrentCell = DataBaseDataGridView[0, hit.RowIndex];
            }

            // テーブル種別に応じてメニュー項目の有効/無効を設定する
            EditContextMenuItem.Enabled = (_tableName == "Product");
            DeleteContextMenuItem.Enabled = (_tableName != string.Empty);
        }

        private void HistoryWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void 編集開始ToolStripMenuItem_Click(object sender, EventArgs e) { EnterEditMode(); }
        private void 編集終了ToolStripMenuItem_Click(object sender, EventArgs e) { ExitEditMode(); }
        private async void EditContextMenuItem_Click(object sender, EventArgs e) { await EditProductRecord(); }
        private async void DeleteContextMenuItem_Click(object sender, EventArgs e) { await DeleteSelectedRows(); }
        private void 在庫調整ToolStripMenuItem_Click(object sender, EventArgs e) { InventoryAdjustment(); }
        private void ShowUsedSubstrateButton_Click(object sender, EventArgs e) { ShowUsedSubstrateDetails(); }
        private async void GenerateReportButton_Click(object sender, EventArgs e) { await GenerateReport(); }
        private async void GenerateListButton_Click(object sender, EventArgs e) { await GenerateList(); }
        private async void GenerateCheckSheetButton_Click(object sender, EventArgs e) { await GenerateCheckSheet(); }
        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void FilterStringTextBox_TextChanged(object sender, EventArgs e) { HistoryTableFilter(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void StockCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void AllSubstrateCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
        private void GroupModelCheckBox_CheckedChanged(object sender, EventArgs e) { ViewSubstrateRegistrationLog(); }
    }
}
