using Microsoft.Data.Sqlite;
using ProductDatabase.Common;
using ProductDatabase.Data;
using ProductDatabase.Excel;
using ProductDatabase.Models;
using ProductDatabase.Other;
using ProductDatabase.Services;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        private readonly List<string> _listUsedSubstrate = [];
        private readonly List<string> _listUsedProductNumber = [];
        private readonly List<int> _listUsedQuantity = [];

        // LoadEvents中のCheckBox_CheckedChanged誤発火を防ぐフラグ
        private bool _isLoadingEvents = false;
        // 排他グループによる自動OFFで警告が出ないよう抑制するフラグ
        private bool _suppressUncheckedWarning = false;
        private readonly List<string> _checkBoxNames = [
                        "Substrate1CheckBox", "Substrate2CheckBox", "Substrate3CheckBox", "Substrate4CheckBox","Substrate5CheckBox",
                        "Substrate6CheckBox", "Substrate7CheckBox", "Substrate8CheckBox", "Substrate9CheckBox","Substrate10CheckBox",
                        "Substrate11CheckBox", "Substrate12CheckBox", "Substrate13CheckBox", "Substrate14CheckBox","Substrate15CheckBox"
                        ];
        private readonly List<string> _dataGridViewNames = [
                        "Substrate1DataGridView", "Substrate2DataGridView", "Substrate3DataGridView", "Substrate4DataGridView","Substrate5DataGridView",
                        "Substrate6DataGridView", "Substrate7DataGridView", "Substrate8DataGridView", "Substrate9DataGridView","Substrate10DataGridView",
                        "Substrate11DataGridView", "Substrate12DataGridView", "Substrate13DataGridView", "Substrate14DataGridView","Substrate15DataGridView"
                        ];

        public SubstrateChange2(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;
        }

        // フォームロード時にUIを初期化し対象製品の基板在庫・使用状況をDBから取得してDataGridViewに表示する
        private void LoadEvents() {
            try {
                Font = new Font(_appSettings.FontName, _appSettings.FontSize);
                CloseButton.Enabled = true;

                PersonComboBox.SelectedIndexChanged += InputControls_TextChanged;
                ValidateAllInputs();

                // テキストボックスに入力
                OrderNumberTextBox.Text = _productRegisterWork.OrderNumber;
                ManufacturingNumberTextBox.Text = _productRegisterWork.ProductNumber;
                QuantityTextBox.Text = _productRegisterWork.Quantity.ToString();
                RevisionTextBox.Text = _productRegisterWork.Revision;
                CommentTextBox.Text = _productRegisterWork.Comment;

                // ComboBoxへ担当者を追加
                var persons = ProductDatabase.Data.PersonRepository.GetActivePersons();
                PersonComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                PersonComboBox.DataSource = persons;
                PersonComboBox.DisplayMember = "DisplayName";
                PersonComboBox.ValueMember = "PersonID";

                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        _isLoadingEvents = true;
                        try {
                            // 排他グループで先に選択された基板を管理するセット
                            var initializedGroups = new HashSet<int>();

                            for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {
                                var substrate = _productMaster.UseSubstrates[i];
                                var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                                var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;

                                // ラベルは先に設定（在庫データがなくても表示されるように）
                                if (objCbx is not null) {
                                    var splitSubstrateName = substrate.SubstrateName.Split(':');
                                    objCbx.Text = $"{splitSubstrateName.Last()} - {substrate.SubstrateModel}";
                                }

                                // 排他グループ内で2番目以降の基板はOFFで開始
                                var shouldBeChecked = true;
                                if (substrate.ExclusiveGroupID.HasValue
                                    && !initializedGroups.Add(substrate.ExclusiveGroupID.Value)) {
                                    shouldBeChecked = false;
                                }

                                if (objCbx is not null) {
                                    objCbx.Enabled = true;
                                    objCbx.Checked = shouldBeChecked;
                                    objCbx.ForeColor = shouldBeChecked ? Color.Black : Color.Red;
                                }

                                if (objDgv is not null) {
                                    objDgv.Visible = shouldBeChecked;
                                    objDgv.Enabled = shouldBeChecked;
                                    objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                                    objDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                    objDgv.Columns[2].ReadOnly = true;
                                    objDgv.Columns[3].ReadOnly = false;
                                    objDgv.Columns[4].ReadOnly = false;
                                    if (_layoutSettings.TryGetValue(Font.Size, out var layout)) {
                                        objDgv.RowTemplate.Height = layout.RowHeight;
                                        for (int z = 0; z < layout.ColumnWidths.Length; z++) {
                                            objDgv.Columns[z].Width = layout.ColumnWidths[z];
                                        }
                                    }
                                }

                                // 在庫データは全基板で読み込む（後から選択された場合に備えて）
                                var list = SubstrateChangeRepository.GetSubstrateStock(
                                    _productRegisterWork.RowID, substrate.SubstrateID);
                                var j = 0;
                                foreach (var row in list) {
                                    if (objDgv is null) break;
                                    objDgv.Rows.Add();
                                    objDgv.Rows[j].Cells[0].Value = row.SubstrateNumber;
                                    objDgv.Rows[j].Cells[1].Value = row.Stock;
                                    objDgv.Rows[j].Cells[2].Value = row.UsedDecrease;
                                    objDgv.Rows[j].Cells[3].Value = row.UsedDecrease;
                                    objDgv.Rows[j].Cells[4].Value = row.UsedDecrease != 0;
                                    j++;
                                }
                            }
                        } finally {
                            _isLoadingEvents = false;
                        }
                        break;
                }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        private class DataGridViewLayoutSettings {
            public int RowHeight { get; init; }
            public int[] ColumnWidths { get; init; } = [];
        }

        private static readonly Dictionary<float, DataGridViewLayoutSettings> _layoutSettings = new() {
            [9f] = new() { RowHeight = 24, ColumnWidths = [130, 30, 30, 30, 24] },
            [12f] = new() { RowHeight = 24, ColumnWidths = [220, 40, 40, 40, 24] },
            [14f] = new() { RowHeight = 25, ColumnWidths = [230, 60, 60, 60, 25] }
        };
        // 数量チェック後にDB登録を実行し登録完了後はフォームを編集不可にしてリスト印刷ボタンを有効化する
        private void ChangeRegistration() {
            try {
                if (!QuantityCheck()) { return; }

                Registration();

                MessageBox.Show("登録完了");

                // フォームを編集不可にする
                RegistrationDateTimePicker.Enabled = false;
                PersonComboBox.Enabled = false;
                RegisterButton.Enabled = false;
                for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {
                    var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                    if (objCbx is not null) {
                        objCbx.Enabled = false;
                    }

                    var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;
                    if (objDgv is not null) {
                        objDgv.Enabled = false;
                    }
                }
                // リスト印刷ボタンを有効に
                if (_productMaster.IsListPrint) { SubstrateListPrintButton.Enabled = true; }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 各基板のDgvで入力された使用数が在庫範囲内か・必要数と一致するかを検証しOK/NGを返す
        private bool QuantityCheck() {
            try {
                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        if (_productMaster.UseSubstrates is null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {

                            var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxが nullです。");
                            objCbx.Enabled = true;

                            var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvが nullです。");
                            objDgv.Columns[3].ReadOnly = false;
                            objDgv.Columns[4].ReadOnly = false;

                            if (objCbx.Checked) {
                                var intQuantityCheck = _productRegisterWork.Quantity;
                                var dgvRowCnt = objDgv.Rows.Count;

                                for (var j = 0; j < dgvRowCnt; j++) {
                                    var boolCbx = objDgv.Rows[j].Cells[4].Value is not null && (bool)objDgv.Rows[j].Cells[4].Value;
                                    if (boolCbx) {
                                        var stockValue = int.TryParse(objDgv.Rows[j].Cells[1].Value?.ToString(), out var sv1) ? sv1 : 0;
                                        var usedValue = int.TryParse(objDgv.Rows[j].Cells[2].Value?.ToString(), out var usd1) ? usd1 : 0;
                                        var useValue = int.TryParse(objDgv.Rows[j].Cells[3].Value?.ToString(), out var use1) ? use1 : 0;

                                        if (useValue < 0) {
                                            throw new Exception("使用数が0未満の値は入力できません。");
                                        }

                                        if (stockValue + usedValue < useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        intQuantityCheck -= useValue;
                                    }
                                }

                                if (intQuantityCheck != 0) {
                                    var dlg = MessageBox.Show(
                                        $"[{_productMaster.UseSubstrates[i].SubstrateName}]の合計が必要数と一致しませんがよろしいですか？",
                                        "確認",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question
                                    );
                                    if (dlg == DialogResult.No) {
                                        return false;
                                    }
                                }
                            }
                        }
                        break;
                    default:
                        break;
                }
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // トランザクションで基板テーブルの使用数を更新（既存行の修正または新規挿入）し製品テーブルも更新してバックアップとログを記録する
        private void Registration() {
            try {
                _productRegisterWork.RegDate = RegistrationDateTimePicker.Value.ToShortDateString();

                if (PersonComboBox.SelectedValue != null && PersonComboBox.SelectedItem is ProductDatabase.Models.PersonDef selectedPerson) {
                    _productRegisterWork.PersonID = (long?)PersonComboBox.SelectedValue;
                    _productRegisterWork.PersonName = selectedPerson.PersonName;
                } else {
                    throw new Exception("担当者を選択してください。");
                }

                _productRegisterWork.Comment = CommentTextBox.Text;

                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        using (SqliteConnection con = new(ProductRepository.GetConnectionRegistration())) {
                            con.Open();
                            using var transaction = con.BeginTransaction();

                            if (_productMaster.UseSubstrates is null) { throw new Exception("ArrUseSubstrateが nullです。"); }
                            for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {
                                var substrate = _productMaster.UseSubstrates[i];
                                var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxが nullです。");

                                if (objCbx.Checked) {
                                    var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvが nullです。");
                                    var dgvRowCnt = objDgv.Rows.Count;

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var usedValue = int.TryParse(objDgv.Rows[j].Cells[2].Value?.ToString(), out var usd2) ? usd2 : 0;
                                        var boolCbx = objDgv.Rows[j].Cells[4].Value is not null && (bool)objDgv.Rows[j].Cells[4].Value;
                                        var useValue = boolCbx ? (int.TryParse(objDgv.Rows[j].Cells[3].Value?.ToString(), out var use2) ? use2 : 0) : 0;

                                        if (boolCbx) {
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value?.ToString() ?? string.Empty;

                                            var info = SubstrateChangeRepository.GetSubstrateInfo(
                                                con, transaction,
                                                substrate.SubstrateID, substrateNum);

                                            var affectedRows = SubstrateChangeRepository.UpdateSubstrateDecrease(
                                                con, transaction,
                                                useValue,
                                                _productRegisterWork.PersonID,
                                                _productRegisterWork.RegDate,
                                                _productRegisterWork.Comment,
                                                substrateNum,
                                                _productRegisterWork.RowID,
                                                substrate.SubstrateID);

                                            // 更新できない場合は挿入
                                            if (affectedRows == 0) {
                                                SubstrateChangeRepository.InsertSubstrateDecrease(
                                                    con, transaction,
                                                    useValue,
                                                    _productRegisterWork.PersonID,
                                                    _productRegisterWork.RegDate,
                                                    _productRegisterWork.Comment,
                                                    substrateNum,
                                                    info?.OrderNumber ?? "",
                                                    _productRegisterWork.RowID,
                                                    substrate.SubstrateID);
                                            }

                                            if (_productMaster.IsListPrint) {
                                                _listUsedSubstrate.Add(substrate.SubstrateName);
                                                _listUsedProductNumber.Add(substrateNum);
                                                _listUsedQuantity.Add(useValue);
                                            }
                                        }
                                        else if (usedValue != useValue && useValue == 0) {
                                            SubstrateChangeRepository.ClearSubstrateDecrease(
                                                con, transaction,
                                                substrate.SubstrateID,
                                                objDgv.Rows[j].Cells[0].Value?.ToString() ?? string.Empty,
                                                _productRegisterWork.RowID);
                                        }
                                    }
                                } else {
                                    // チェックOFF（排他グループの自動OFF含む）の基板は既存の使用数をすべてクリアする
                                    var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;
                                    if (objDgv is not null) {
                                        for (var j = 0; j < objDgv.Rows.Count; j++) {
                                            var usedValue = int.TryParse(objDgv.Rows[j].Cells[2].Value?.ToString(), out var usd) ? usd : 0;
                                            if (usedValue != 0) {
                                                SubstrateChangeRepository.ClearSubstrateDecrease(
                                                    con, transaction,
                                                    substrate.SubstrateID,
                                                    objDgv.Rows[j].Cells[0].Value?.ToString() ?? string.Empty,
                                                    _productRegisterWork.RowID);
                                            }
                                        }
                                    }
                                }
                            }

                            // 製品テーブル更新
                            SubstrateChangeRepository.UpdateProduct(
                                con, transaction, _productRegisterWork, _productMaster.RevisionGroup);

                            transaction.Commit();
                        }
                        break;
                }

                // バックアップ作成
                BackupManager.CreateBackup();
                // ログ出力
                HistoryAuditLogger.LogSubstrateChange(_productMaster, _productRegisterWork);
            } catch (Exception) {
                throw;
            }
        }
        // 基板変更済み製品情報をもとにExcel製品一覧を生成する
        private async Task GenerateList() {
            SubstrateListPrintButton.Enabled = false;
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

            SubstrateListPrintButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateList)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // チェックボックスのON/OFFに応じて対応するDataGridViewの表示・有効状態を切り替え未チェック時は警告を表示する
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            if (_isLoadingEvents) return;

            var checkBox = (CheckBox)sender;
            var dataGridView = GetDataGridViewForCheckBox(checkBox.Name);

            if (dataGridView is null) return;

            dataGridView.Enabled = checkBox.Checked;
            dataGridView.Visible = checkBox.Checked;
            if (checkBox.Checked) {
                // 非表示中に行を追加されたDataGridViewはスクロール範囲の再計算が行われないため、表示時に強制する
                dataGridView.PerformLayout();
            }
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            // 排他グループの他チェックボックスを自動OFF（チェックON時のみ）
            if (checkBox.Checked) {
                var idx = _checkBoxNames.IndexOf(checkBox.Name);
                if (idx >= 0 && idx < _productMaster.UseSubstrates.Count) {
                    var groupId = _productMaster.UseSubstrates[idx].ExclusiveGroupID;
                    if (groupId.HasValue) {
                        _suppressUncheckedWarning = true;
                        try {
                            for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {
                                if (i == idx) continue;
                                if (_productMaster.UseSubstrates[i].ExclusiveGroupID == groupId) {
                                    if (MainPanel.Controls[_checkBoxNames[i]] is CheckBox otherCbx && otherCbx.Checked)
                                        otherCbx.Checked = false;
                                }
                            }
                        } finally {
                            _suppressUncheckedWarning = false;
                        }
                    }
                }
            }

            if (!checkBox.Checked && !_suppressUncheckedWarning) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "",
                               MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // 担当者選択状態を検証しエラーがあればメッセージ表示と登録ボタン無効化を行う
        private void ValidateAllInputs() {
            ErrorMessageLabel.Text = "";
            RegisterButton.Enabled = true;

            if (PersonComboBox.SelectedIndex == -1 && PersonComboBox.Enabled) {
                ShowError("担当者が選択されていません。");
                return;
            }

        }
        // エラーメッセージラベルに赤字でメッセージを表示し登録ボタンを無効化する
        private void ShowError(string message) {
            ErrorMessageLabel.Text = message;
            ErrorMessageLabel.ForeColor = Color.Red;
            RegisterButton.Enabled = false;
        }

        // 入力コントロール変更時にすべての入力検証を再実行する
        private void InputControls_TextChanged(object? sender, EventArgs e) {
            ValidateAllInputs();
        }

        // チェックボックス名から対応するDataGridViewを返す
        private DataGridView? GetDataGridViewForCheckBox(string checkBoxName) {
            return checkBoxName switch {
                "Substrate1CheckBox" => Substrate1DataGridView,
                "Substrate2CheckBox" => Substrate2DataGridView,
                "Substrate3CheckBox" => Substrate3DataGridView,
                "Substrate4CheckBox" => Substrate4DataGridView,
                "Substrate5CheckBox" => Substrate5DataGridView,
                "Substrate6CheckBox" => Substrate6DataGridView,
                "Substrate7CheckBox" => Substrate7DataGridView,
                "Substrate8CheckBox" => Substrate8DataGridView,
                "Substrate9CheckBox" => Substrate9DataGridView,
                "Substrate10CheckBox" => Substrate10DataGridView,
                "Substrate11CheckBox" => Substrate11DataGridView,
                "Substrate12CheckBox" => Substrate12DataGridView,
                "Substrate13CheckBox" => Substrate13DataGridView,
                "Substrate14CheckBox" => Substrate14DataGridView,
                "Substrate15CheckBox" => Substrate15DataGridView,
                _ => null
            };
        }

        private void SubstrateChange2_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { ChangeRegistration(); }
        private async void SubstrateListPrintButton_Click(object sender, EventArgs e) { await GenerateList(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
