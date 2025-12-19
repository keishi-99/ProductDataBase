using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Other;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

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
        private CheckBox? _objCbx;
        private DataGridView? _objDgv;

        public SubstrateChange2(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();

            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(_appSettings.FontName, _appSettings.FontSize);
                CloseButton.Enabled = true;

                // テキストボックスに入力
                OrderNumberTextBox.Text = _productRegisterWork.OrderNumber;
                ManufacturingNumberTextBox.Text = _productRegisterWork.ProductNumber;
                QuantityTextBox.Text = _productRegisterWork.Quantity.ToString();
                RevisionTextBox.Text = _productRegisterWork.Revision;
                CommentTextBox.Text = _productRegisterWork.Comment;

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. _appSettings.PersonList]);

                var strQuantity = string.Empty;

                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        for (var i = 0; i <= _productMaster.UseSubstrates.Count - 1; i++) {
                            var substrateName = _productMaster.UseSubstrates[i].SubstrateName;
                            var substrateModel = _productMaster.UseSubstrates[i].SubstrateModel;
                            var quantity = _productRegisterWork.Quantity;

                            // チェックボックスとDgvを有効に
                            _objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                            if (_objCbx is not null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            _objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;
                            if (_objDgv is not null) {
                                _objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                                _objDgv.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                                _objDgv.Columns[2].ReadOnly = true;
                                _objDgv.Columns[3].ReadOnly = false;
                                _objDgv.Columns[4].ReadOnly = false;
                                switch (Font.Size) {
                                    case 9:
                                        _objDgv.RowTemplate.Height = 24;
                                        _objDgv.Columns[0].Width = 130;
                                        _objDgv.Columns[1].Width = 30;
                                        _objDgv.Columns[2].Width = 30;
                                        _objDgv.Columns[3].Width = 30;
                                        _objDgv.Columns[4].Width = 24;
                                        break;
                                    case 12:
                                        _objDgv.RowTemplate.Height = 24;
                                        _objDgv.Columns[0].Width = 220;
                                        _objDgv.Columns[1].Width = 40;
                                        _objDgv.Columns[2].Width = 40;
                                        _objDgv.Columns[3].Width = 40;
                                        _objDgv.Columns[4].Width = 24;
                                        break;
                                    case 14:
                                        _objDgv.RowTemplate.Height = 25;
                                        _objDgv.Columns[0].Width = 230;
                                        _objDgv.Columns[1].Width = 60;
                                        _objDgv.Columns[2].Width = 60;
                                        _objDgv.Columns[3].Width = 60;
                                        _objDgv.Columns[4].Width = 25;
                                        break;
                                }
                            }

                            using SqliteConnection con = new(GetConnectionRegistration());
                            con.Open();
                            using var cmd = con.CreateCommand();

                            cmd.CommandText =
                                $"""
                                -- 使用履歴のある SubstrateNumber ごとの合計を別テーブルで集計
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
                                        AND SubstrateNumber NOTNULL
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
                                        AND SubstrateNumber NOTNULL
                                    GROUP BY
                                        SubstrateName,
                                        SubstrateModel,
                                        SubstrateNumber
                                )

                                -- 最終結合（在庫がある or 使用済み）を条件に絞る
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

                            cmd.Parameters.Clear();
                            cmd.Parameters.Add("@ID", SqliteType.Integer).Value = _productRegisterWork.RowID;
                            cmd.Parameters.Add("@SubstrateModel", SqliteType.Text).Value = substrateModel;

                            using var dr = cmd.ExecuteReader();
                            var j = 0;

                            if (_objCbx is not null) {
                                var splitSubstrateName = substrateName.Split(':');
                                _objCbx.Text = $"{splitSubstrateName.Last()} - {substrateModel}";
                            }
                            while (dr.Read()) {
                                var strSubstrateNum = dr["SubstrateNumber"].ToString() ?? string.Empty;
                                var intStock = Convert.ToInt32(dr["Stock"]);
                                var intUsedQuantity = Convert.ToInt32(dr["UsedDecrease"]); ;

                                if (_objDgv is null) {
                                    break;
                                }

                                _objDgv.Rows.Add();
                                _objDgv.Rows[j].Cells[0].Value = strSubstrateNum;
                                _objDgv.Rows[j].Cells[1].Value = intStock;
                                _objDgv.Rows[j].Cells[2].Value = intUsedQuantity;
                                _objDgv.Rows[j].Cells[3].Value = intUsedQuantity;
                                _objDgv.Rows[j].Cells[4].Value = intUsedQuantity != 0;
                                j++;
                            }
                        }
                        break;
                }

                RegisterButton.Enabled = true;

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
            }
        }
        // 変更登録
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
                    _objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                    if (_objCbx is not null) {
                        _objCbx.Enabled = false;
                    }

                    _objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;
                    if (_objDgv is not null) {
                        _objDgv.Enabled = false;
                    }
                }
                // リスト印刷ボタンを有効に
                if (_productMaster.IsListPrint) { SubstrateListPrintButton.Enabled = true; }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool QuantityCheck() {
            try {
                switch (_productMaster.RegType) {
                    case 2:
                    case 3:
                        if (_productMaster.UseSubstrates is null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i <= _productMaster.UseSubstrates.Count - 1; i++) {

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

                                        if (useValue <= 0) {
                                            throw new Exception("使用数が0以下になっています。");
                                        }
                                        if (stockValue + usedValue < useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        intQuantityCheck -= useValue;
                                    }
                                }

                                if (intQuantityCheck != 0) {
                                    throw new Exception("入力された数量の合計が必要数と一致しません。");
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
                            for (var i = 0; i <= _productMaster.UseSubstrates.Count; i++) {

                                var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxが nullです。");

                                if (objCbx.Checked) {
                                    var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvが nullです。");
                                    var dgvRowCnt = objDgv.Rows.Count;

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[4].Value);
                                        var useValue = boolCbx ? Convert.ToInt32(objDgv.Rows[j].Cells[3].Value) : 0;
                                        if (boolCbx) {
                                            var substrateID = string.Empty;
                                            var orderNum = string.Empty;
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                            var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);

                                            using (var cmd = con.CreateCommand()) {
                                                cmd.CommandText =
                                                    $"""
                                                    SELECT
                                                        SubstrateName,
                                                        SubstrateModel,
                                                        SubstrateNumber,
                                                        OrderNumber,
                                                        SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock,
                                                        SubstrateID
                                                    FROM
                                                        {Constants.VSubstrateTableName}
                                                    WHERE
                                                        SubstrateModel = @SubstrateModel 
                                                        AND SubstrateNumber = @SubstrateNumber
                                                    GROUP BY
                                                        SubstrateName,
                                                        SubstrateModel,
                                                        SubstrateNumber,
                                                        OrderNumber
                                                    ORDER BY
                                                        MIN(ID)
                                                    ;
                                                    """;
                                                cmd.Parameters.Add("@SubstrateModel", SqliteType.Text).Value = _productMaster.UseSubstrates[i].SubstrateModel;
                                                cmd.Parameters.Add("@SubstrateNumber", SqliteType.Text).Value = substrateNum;
                                                using var dr = cmd.ExecuteReader();
                                                while (dr.Read()) {
                                                    substrateID = $"{dr["SubstrateID"]}";
                                                    orderNum = $"{dr["OrderNumber"]}";
                                                }
                                            }
                                            // 基板テーブルを更新、できない場合は挿入
                                            using (var cmdUpdate = con.CreateCommand()) {
                                                // 更新
                                                cmdUpdate.CommandText =
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
                                                        AND UseID = @UseID
                                                    ;
                                                    """;

                                                cmdUpdate.Parameters.Add("@Decrease", SqliteType.Integer).Value = 0 - useValue;
                                                cmdUpdate.Parameters.Add("@Person", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Person) ? DBNull.Value : _productRegisterWork.Person;
                                                cmdUpdate.Parameters.Add("@RegDate", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.RegDate) ? DBNull.Value : _productRegisterWork.RegDate;
                                                cmdUpdate.Parameters.Add("@Comment", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Comment) ? DBNull.Value : _productRegisterWork.Comment;
                                                cmdUpdate.Parameters.Add("@SubstrateNumber", SqliteType.Text).Value = objDgv.Rows[j].Cells[0].Value;
                                                cmdUpdate.Parameters.Add("@UseID", SqliteType.Text).Value = _productRegisterWork.RowID;

                                                var affectedRows = cmdUpdate.ExecuteNonQuery();

                                                if (affectedRows == 0) {
                                                    // 挿入
                                                    using var cmdInsert = con.CreateCommand();
                                                    cmdInsert.CommandText =
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
                                                        ;
                                                        """;

                                                    cmdInsert.Parameters.Add("@SubstrateID", SqliteType.Text).Value = string.IsNullOrWhiteSpace(substrateID) ? DBNull.Value : substrateID;
                                                    cmdInsert.Parameters.Add("@SubstrateNumber", SqliteType.Text).Value = objDgv.Rows[j].Cells[0].Value;
                                                    cmdInsert.Parameters.Add("@OrderNumber", SqliteType.Text).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                                                    cmdInsert.Parameters.Add("@Decrease", SqliteType.Integer).Value = 0 - useValue;
                                                    cmdInsert.Parameters.Add("@Person", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Person) ? DBNull.Value : _productRegisterWork.Person;
                                                    cmdInsert.Parameters.Add("@RegDate", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.RegDate) ? DBNull.Value : _productRegisterWork.RegDate;
                                                    cmdInsert.Parameters.Add("@Comment", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Comment) ? DBNull.Value : _productRegisterWork.Comment;
                                                    cmdInsert.Parameters.Add("@UseID", SqliteType.Text).Value = _productRegisterWork.RowID;

                                                    cmdInsert.ExecuteNonQuery();
                                                }
                                            }

                                            if (_productMaster.IsListPrint) {
                                                _listUsedSubstrate.Add(_productMaster.UseSubstrates[i].SubstrateName);
                                                if (substrateNum is not null) { _listUsedProductNumber.Add(substrateNum); }
                                                _listUsedQuantity.Add(useValue);
                                            }
                                        }
                                        // 使用数が0になった場合、基板テーブルから削除
                                        else if (usedValue != useValue && useValue == 0) {
                                            using var cmdDelete = con.CreateCommand();
                                            cmdDelete.CommandText =
                                                $"""
                                                DELETE
                                                FROM
                                                    {Constants.TSubstrateTableName}
                                                WHERE
                                                    SubstrateNumber = @SubstrateNumber
                                                    AND UseID = @ID
                                                ;
                                                """;
                                            cmdDelete.Parameters.Clear(); // パラメータをクリア
                                            cmdDelete.Parameters.Add("@SubstrateNumber", SqliteType.Text).Value = objDgv.Rows[j].Cells[0].Value;
                                            cmdDelete.Parameters.Add("@ID", SqliteType.Integer).Value = _productRegisterWork.RowID;

                                            cmdDelete.Connection = con;
                                            cmdDelete.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }

                            // 製品テーブル更新
                            using (var cmd = con.CreateCommand()) {
                                cmd.CommandText =
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
                                    ;
                                    """;
                                cmd.Parameters.Add("@OrderNumber", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.OrderNumber) ? DBNull.Value : _productRegisterWork.OrderNumber;
                                cmd.Parameters.Add("@ProductNumber", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.ProductNumber) ? DBNull.Value : _productRegisterWork.ProductNumber;
                                cmd.Parameters.Add("@Quantity", SqliteType.Text).Value = _productRegisterWork.Quantity;
                                cmd.Parameters.Add("@Person", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Person) ? DBNull.Value : _productRegisterWork.Person;
                                cmd.Parameters.Add("@RegDate", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.RegDate) ? DBNull.Value : _productRegisterWork.RegDate;
                                cmd.Parameters.Add("@Revision", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Revision) ? DBNull.Value : _productRegisterWork.Revision;
                                cmd.Parameters.Add("@RevisionGroup", SqliteType.Text).Value = _productMaster.RevisionGroup;
                                cmd.Parameters.Add("@SerialFirst", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.SerialFirst) ? DBNull.Value : _productRegisterWork.SerialFirst;
                                cmd.Parameters.Add("@SerialLast", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.SerialLast) ? DBNull.Value : _productRegisterWork.SerialLast;
                                cmd.Parameters.Add("@SerialLastNumber", SqliteType.Text).Value = _productRegisterWork.SerialLastNumber;
                                cmd.Parameters.Add("@Comment", SqliteType.Text).Value = string.IsNullOrWhiteSpace(_productRegisterWork.Comment) ? DBNull.Value : _productRegisterWork.Comment;

                                cmd.ExecuteNonQuery();
                                transaction.Commit();
                            }
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

        // リスト印刷
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

        // チェックボックスイベント
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var checkBox = (CheckBox)sender;
            DataGridView dataGridView = new();

            switch (checkBox.Name) {
                case "Substrate1CheckBox":
                    dataGridView = Substrate1DataGridView;
                    break;
                case "Substrate2CheckBox":
                    dataGridView = Substrate2DataGridView;
                    break;
                case "Substrate3CheckBox":
                    dataGridView = Substrate3DataGridView;
                    break;
                case "Substrate4CheckBox":
                    dataGridView = Substrate4DataGridView;
                    break;
                case "Substrate5CheckBox":
                    dataGridView = Substrate5DataGridView;
                    break;
                case "Substrate6CheckBox":
                    dataGridView = Substrate6DataGridView;
                    break;
                case "Substrate7CheckBox":
                    dataGridView = Substrate7DataGridView;
                    break;
                case "Substrate8CheckBox":
                    dataGridView = Substrate8DataGridView;
                    break;
                case "Substrate9CheckBox":
                    dataGridView = Substrate9DataGridView;
                    break;
                case "Substrate10CheckBox":
                    dataGridView = Substrate10DataGridView;
                    break;
                case "Substrate11CheckBox":
                    dataGridView = Substrate11DataGridView;
                    break;
                case "Substrate12CheckBox":
                    dataGridView = Substrate12DataGridView;
                    break;
                case "Substrate13CheckBox":
                    dataGridView = Substrate13DataGridView;
                    break;
                case "Substrate14CheckBox":
                    dataGridView = Substrate14DataGridView;
                    break;
                case "Substrate15CheckBox":
                    dataGridView = Substrate15DataGridView;
                    break;
                default:
                    break;
            }

            dataGridView.Enabled = checkBox.Checked;
            dataGridView.Visible = checkBox.Checked;
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SubstrateChange2_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { ChangeRegistration(); }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { GenerateList(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
