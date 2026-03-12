using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Models;
using ProductDatabase.Other;
using static ProductDatabase.Data.ProductRepository;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        private sealed class SubstrateStockDto {
            public string SubstrateName { get; init; } = string.Empty;
            public string SubstrateModel { get; init; } = string.Empty;
            public string SubstrateNumber { get; init; } = string.Empty;
            public int Stock { get; init; }
            public int UsedDecrease { get; init; }
        }

        private sealed class SubstrateInfoDto {
            public string SubstrateID { get; init; } = string.Empty;
            public string OrderNumber { get; init; } = string.Empty;
        }

        private readonly List<string> _listUsedSubstrate = [];
        private readonly List<string> _listUsedProductNumber = [];
        private readonly List<int> _listUsedQuantity = [];
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
                PersonComboBox.Items.AddRange([.. _appSettings.PersonList]);

                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {
                            var substrateName = _productMaster.UseSubstrates[i].SubstrateName;
                            var substrateModel = _productMaster.UseSubstrates[i].SubstrateModel;
                            var quantity = _productRegisterWork.Quantity;
                            var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                            var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;

                            // チェックボックスとDgvを有効に
                            objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                            if (objCbx is not null) {
                                objCbx.Enabled = true;
                                objCbx.Checked = true;
                            }

                            objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;
                            if (objDgv is not null) {
                                objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                                objDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                objDgv.Columns[2].ReadOnly = true;
                                objDgv.Columns[3].ReadOnly = false;
                                objDgv.Columns[4].ReadOnly = false;
                                if (s_layoutSettings.TryGetValue(Font.Size, out var layout)) {
                                    objDgv.RowTemplate.Height = layout.RowHeight;
                                    for (int z = 0; z < layout.ColumnWidths.Length; z++) {
                                        objDgv.Columns[z].Width = layout.ColumnWidths[z];
                                    }
                                }
                            }
                            using SqliteConnection con = new(GetConnectionRegistration());

                            var sql =
                                $"""
                                WITH Used AS 
                                (
                                    SELECT
                                        SubstrateNumber,
                                        COALESCE(SUM(Decrease), 0) AS UsedDecrease
                                    FROM
                                        {Constants.VSubstrateTableName}
                                    WHERE
                                        SubstrateModel = @SubstrateModel
                                        AND UseID = @ID
                                        AND SubstrateNumber IS NOT NULL
                                        AND IsDeleted = 0
                                    GROUP BY
                                        SubstrateNumber
                                ),

                                Stocked AS 
                                (
                                    SELECT
                                        SubstrateName,
                                        SubstrateModel,
                                        SubstrateNumber,
                                        COALESCE(SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)), 0) AS Stock
                                    FROM
                                        {Constants.VSubstrateTableName}
                                    WHERE
                                        SubstrateModel = @SubstrateModel
                                        AND SubstrateNumber IS NOT NULL
                                        AND IsDeleted = 0
                                    GROUP BY
                                        SubstrateName,
                                        SubstrateModel,
                                        SubstrateNumber
                                )

                                SELECT
                                    s.SubstrateName,
                                    s.SubstrateModel,
                                    s.SubstrateNumber,
                                    s.Stock,
                                    COALESCE(u.UsedDecrease * -1, 0) AS UsedDecrease
                                FROM
                                    Stocked s
                                    LEFT JOIN Used u ON s.SubstrateNumber = u.SubstrateNumber
                                WHERE
                                    s.Stock > 0 
                                    OR u.UsedDecrease IS NOT NULL
                                ORDER BY
                                    CASE WHEN COALESCE(u.UsedDecrease, 0) = 0 THEN 1 ELSE 0 END,
                                    s.SubstrateNumber;
                                """;

                            var list = con.Query<SubstrateStockDto>(sql,
                                new {
                                    ID = _productRegisterWork.RowID,
                                    SubstrateModel = substrateModel
                                });

                            var j = 0;

                            foreach (var row in list) {

                                if (objCbx is not null) {
                                    var splitSubstrateName = substrateName.Split(':');
                                    objCbx.Text = $"{splitSubstrateName.Last()} - {substrateModel}";
                                }

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

        private static readonly Dictionary<float, DataGridViewLayoutSettings> s_layoutSettings = new() {
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
                for (var i = 0; i <= 9; i++) {
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
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value.ToString());
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[3].Value.ToString());

                                        if (useValue < 0) {
                                            throw new Exception("使用数が0以下になっています。");
                                        }

                                        // 使用数がマイナスの場合は確認ダイアログを表示
                                        if (useValue < 0) {
                                            throw new Exception("使用数が0未満になっています。");
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
                _productRegisterWork.Person = PersonComboBox.Text;
                _productRegisterWork.Comment = CommentTextBox.Text;
                if (string.IsNullOrEmpty(_productRegisterWork.Person)) { throw new Exception("担当者を選択してください。"); }

                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        using (SqliteConnection con = new(GetConnectionRegistration())) {
                            con.Open();
                            using var transaction = con.BeginTransaction();

                            if (_productMaster.UseSubstrates is null) { throw new Exception("ArrUseSubstrateが nullです。"); }
                            for (var i = 0; i < _productMaster.UseSubstrates.Count; i++) {

                                var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxが nullです。");

                                if (objCbx.Checked) {
                                    var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvが nullです。");
                                    var dgvRowCnt = objDgv.Rows.Count;

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[4].Value);
                                        var useValue = boolCbx ? Convert.ToInt32(objDgv.Rows[j].Cells[3].Value) : 0;
                                        if (boolCbx) {
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                            var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);

                                            var sql =
                                                $"""
                                                SELECT
                                                    SubstrateID,
                                                    OrderNumber
                                                FROM
                                                    {Constants.VSubstrateTableName}
                                                WHERE
                                                    SubstrateID = @SubstrateID
                                                    AND SubstrateNumber = @SubstrateNumber
                                                    AND IsDeleted = 0
                                                GROUP BY
                                                    SubstrateID,
                                                    SubstrateNumber,
                                                    OrderNumber
                                                ORDER BY
                                                    MIN(ID)
                                                LIMIT 1
                                                """;

                                            var info = con.QueryFirstOrDefault<SubstrateInfoDto>(sql,
                                                new {
                                                    _productMaster.UseSubstrates[i].SubstrateID,
                                                    SubstrateNumber = substrateNum
                                                },
                                                transaction
                                            );

                                            var substrateID = info?.SubstrateID ?? "";
                                            var orderNum = info?.OrderNumber ?? "";

                                            // 基板テーブルを更新
                                            var updateSql =
                                                $"""
                                                UPDATE
                                                    {Constants.TSubstrateTableName}
                                                SET
                                                    Decrease = @Decrease,
                                                    Person = @Person,
                                                    RegDate = @RegDate,
                                                    Comment = @Comment
                                                WHERE
                                                    SubstrateNumber = @SubstrateNumber
                                                    AND IsDeleted = 0
                                                    AND UseID = @UseID
                                                """;

                                            var affectedRows = con.Execute(updateSql,
                                                new {
                                                    Decrease = 0 - useValue,
                                                    Person = ToDbValue(_productRegisterWork.Person),
                                                    RegDate = ToDbValue(_productRegisterWork.RegDate),
                                                    Comment = ToDbValue(_productRegisterWork.Comment),
                                                    SubstrateNumber = objDgv.Rows[j].Cells[0].Value,
                                                    UseID = _productRegisterWork.RowID
                                                },
                                                transaction
                                            );

                                            // 更新できない場合は挿入
                                            if (affectedRows == 0) {
                                                var insertSql =
                                                    $"""
                                                    INSERT INTO {Constants.TSubstrateTableName}
                                                        (
                                                            SubstrateID,
                                                            SubstrateNumber,
                                                            OrderNumber,
                                                            Decrease,
                                                            Person,
                                                            RegDate,
                                                            Comment,
                                                            UseID
                                                        )
                                                    VALUES
                                                        (
                                                            @SubstrateID,
                                                            @SubstrateNumber,
                                                            @OrderNumber,
                                                            @Decrease,
                                                            @Person,
                                                            @RegDate,
                                                            @Comment,
                                                            @UseID
                                                        )
                                                    """;

                                                con.Execute(insertSql,
                                                    new {
                                                        SubstrateID = ToDbValue(substrateID),
                                                        SubstrateNumber = objDgv.Rows[j].Cells[0].Value,
                                                        OrderNumber = ToDbValue(orderNum),
                                                        Decrease = 0 - useValue,
                                                        Person = ToDbValue(_productRegisterWork.Person),
                                                        RegDate = ToDbValue(_productRegisterWork.RegDate),
                                                        Comment = ToDbValue(_productRegisterWork.Comment),
                                                        UseID = _productRegisterWork.RowID
                                                    },
                                                    transaction
                                                );
                                            }

                                            if (_productMaster.IsListPrint) {
                                                _listUsedSubstrate.Add(_productMaster.UseSubstrates[i].SubstrateName);
                                                if (substrateNum is not null) { _listUsedProductNumber.Add(substrateNum); }
                                                _listUsedQuantity.Add(useValue);
                                            }
                                        }
                                        else if (usedValue != useValue && useValue == 0) {
                                            var deleteSql =
                                                $"""
                                                UPDATE
                                                    {Constants.TSubstrateTableName}
                                                SET
                                                    Decrease = @Decrease
                                                WHERE
                                                    SubstrateID = @SubstrateID
                                                    AND SubstrateNumber = @SubstrateNumber
                                                    AND IsDeleted = 0
                                                    AND UseID = @UseID
                                                """;

                                            con.Execute(deleteSql,
                                                new {
                                                    Decrease = 0,
                                                    _productMaster.UseSubstrates[i].SubstrateID,
                                                    SubstrateNumber = objDgv.Rows[j].Cells[0].Value,
                                                    UseID = _productRegisterWork.RowID
                                                },
                                                transaction
                                            );
                                        }
                                    }
                                }
                            }

                            // 製品テーブル更新
                            var updateProductSql =
                                $"""
                                UPDATE
                                    {Constants.TProductTableName}
                                SET
                                    Quantity = @Quantity,
                                    Person = @Person,
                                    RegDate = @RegDate,
                                    Revision = @Revision,
                                    RevisionGroup = @RevisionGroup,
                                    SerialLast = @SerialLast,
                                    SerialLastNumber = @SerialLastNumber,
                                    Comment = @Comment
                                WHERE
                                    ProductNumber = @ProductNumber
                                    AND SerialFirst = @SerialFirst
                                    AND IsDeleted = 0
                                """;

                            con.Execute(updateProductSql,
                                new {
                                    OrderNumber = ToDbValue(_productRegisterWork.OrderNumber),
                                    ProductNumber = ToDbValue(_productRegisterWork.ProductNumber),
                                    _productRegisterWork.Quantity,
                                    Person = ToDbValue(_productRegisterWork.Person),
                                    RegDate = ToDbValue(_productRegisterWork.RegDate),
                                    Revision = ToDbValue(_productRegisterWork.Revision),
                                    _productMaster.RevisionGroup,
                                    SerialFirst = ToDbValue(_productRegisterWork.SerialFirst),
                                    SerialLast = ToDbValue(_productRegisterWork.SerialLast),
                                    _productRegisterWork.SerialLastNumber,
                                    Comment = ToDbValue(_productRegisterWork.Comment)
                                },
                                transaction
                            );

                            transaction.Commit();
                        }
                        break;
                }

                // バックアップ作成
                CommonUtils.BackupManager.CreateBackup();
                // ログ出力
                string[] logMessageArray = [
                    $"[基板変更]",
                    $"[{_productMaster.CategoryName}]",
                    $"ID[{_productRegisterWork.RowID}]",
                    $"注文番号[{_productRegisterWork.OrderNumber}]",
                    $"製造番号[{_productRegisterWork.ProductNumber}]",
                    $"[]",
                    $"製品名[{_productMaster.ProductName}]",
                    $"タイプ[{_productMaster.ProductType}]",
                    $"型式[{_productMaster.ProductModel}]",
                    $"数量[{_productRegisterWork.Quantity}]",
                    $"シリアル先頭[{_productRegisterWork.SerialFirst}]",
                    $"シリアル末尾[{_productRegisterWork.SerialLast}]",
                    $"Revision[{_productRegisterWork.Revision}]",
                    $"登録日[{_productRegisterWork.RegDate}]",
                    $"担当者[{_productRegisterWork.Person}]",
                    $"コメント[{_productRegisterWork.Comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);
            } catch (Exception) {
                throw;
            }
        }

        // 空文字・空白文字列はDBNull.Valueに変換してDB挿入/更新時のNULL処理を統一する
        private static object ToDbValue(string? value) {
            return string.IsNullOrWhiteSpace(value) ? DBNull.Value : value;
        }
        // 基板変更済み製品情報をもとにExcel製品一覧を生成する
        private void GenerateList() {
            try {
                // --- 処理中カーソルに変更 ---
                Cursor.Current = Cursors.WaitCursor;

                ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(_productMaster, _productRegisterWork);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                // --- 通常カーソルに戻す ---
                Cursor.Current = Cursors.Default;
            }
        }

        // チェックボックスのON/OFFに応じて対応するDataGridViewの表示・有効状態を切り替え未チェック時は警告を表示する
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var checkBox = (CheckBox)sender;
            var dataGridView = GetDataGridViewForCheckBox(checkBox.Name);

            if (dataGridView is null) return;

            dataGridView.Enabled = checkBox.Checked;
            dataGridView.Visible = checkBox.Checked;
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
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
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { GenerateList(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
