using ProductDatabase.Other;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        private string _totalSubstrate = string.Empty;

        private string[] _useSubstrate = [];
        private string[] _usedSubstrate = [];

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

        // プロパティ設定
        private bool IsListPrint => ProductInfo.PrintType is 5 or 6;

        public SubstrateChange2() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            try {
                Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);
                CloseButton.Enabled = true;

                // テキストボックスに入力
                OrderNumberTextBox.Text = ProductInfo.OrderNumber;
                ManufacturingNumberTextBox.Text = ProductInfo.ProductNumber;
                QuantityTextBox.Text = ProductInfo.Quantity.ToString();
                var dtNow = DateTime.Now;
                RegistrationDateMaskedTextBox.Text = dtNow.ToShortDateString();
                RevisionTextBox.Text = ProductInfo.Revision;
                CommentTextBox.Text = ProductInfo.Comment;

                // DB1へ接続し担当者取得
                using (SQLiteConnection con = new(GetConnectionInformation())) {
                    con.Open();
                    using var cmd = con.CreateCommand();
                    // テーブル検索SQL - 担当者をComboboxへ追加
                    cmd.CommandText = "SELECT * FROM Person ORDER BY _rowid_ ASC";
                    using var dr = cmd.ExecuteReader();
                    while (dr.Read()) {
                        PersonComboBox.Items.Add($"{dr["PersonName"]}");
                    }
                }

                _useSubstrate = ProductInfo.UseSubstrate.Split(",");
                _usedSubstrate = ProductInfo.UsedSubstrate.Split(",");
                var strQuantity = string.Empty;
                var strSubstrateName = string.Empty;

                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {
                            var quantity = ProductInfo.Quantity;

                            // チェックボックスとDgvを有効に
                            _objCbx = Controls[_checkBoxNames[i]] as CheckBox;
                            if (_objCbx != null) {
                                _objCbx.Enabled = true;
                                _objCbx.Checked = true;
                            }

                            _objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                            if (_objDgv != null) {
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

                            using SQLiteConnection con = new(GetConnectionRegistration());
                            con.Open();
                            using var cmd = con.CreateCommand();

                            // テーブル検索SQL - [[StockName]_StockView]テーブルから基板型式[Model]で在庫基板を抽出
                            cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateModel = @SubstrateModel""";
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                            using var dr = cmd.ExecuteReader();
                            var j = 0;
                            while (dr.Read()) {
                                var strUsedSubNum = string.Empty;
                                var strUsedQuantity = string.Empty;
                                var intUsedQuantity = 0;

                                // 抽出した行から製造番号,在庫取得
                                var strSubstrateNum = $"{dr["SubstrateNumber"]}";
                                var intStock = int.Parse($"{dr["Stock"]}");
                                strSubstrateName = $"{dr["SubstrateName"]}";
                                if (_objCbx != null) { _objCbx.Text = $"{strSubstrateName} - {_useSubstrate[i]}"; }

                                var subIndex = 0;
                                var b1 = false;

                                for (var i1 = 0; i1 < _usedSubstrate.Length; i1++) {
                                    var d = _usedSubstrate[i1];
                                    b1 = d.Contains(strSubstrateNum);
                                    if (b1 == true) {
                                        break;
                                    }

                                    subIndex++;
                                }
                                if (b1 == true) {
                                    strUsedSubNum = _usedSubstrate[subIndex];
                                    strUsedSubNum = strUsedSubNum[(strUsedSubNum.IndexOf(']') + 1)..];
                                    strUsedSubNum = strUsedSubNum[..strUsedSubNum.IndexOf('(')];

                                    strUsedQuantity = _usedSubstrate[subIndex];
                                    strUsedQuantity = strUsedQuantity[(strUsedQuantity.IndexOf('(') + 1)..];
                                    strUsedQuantity = strUsedQuantity[..strUsedQuantity.IndexOf(')')];
                                    intUsedQuantity = int.Parse(strUsedQuantity);
                                }

                                if (intStock > 0) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[j].Cells[0].Value = strSubstrateNum;
                                    _objDgv.Rows[j].Cells[1].Value = intStock;

                                    if (intUsedQuantity != 0) {
                                        _objDgv.Rows[j].Cells[2].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[3].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[j].Cells[2].Value = 0;
                                        _objDgv.Rows[j].Cells[3].Value = 0;
                                    }
                                    j++;
                                }
                                else if (strUsedSubNum == strSubstrateNum) {
                                    if (_objDgv == null) { break; }
                                    _objDgv.Rows.Add();
                                    _objDgv.Rows[j].Cells[0].Value = strSubstrateNum;
                                    _objDgv.Rows[j].Cells[1].Value = intStock;

                                    if (intUsedQuantity != 0) {
                                        _objDgv.Rows[j].Cells[2].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[3].Value = intUsedQuantity;
                                        _objDgv.Rows[j].Cells[4].Value = true;
                                    }
                                    else {
                                        _objDgv.Rows[j].Cells[2].Value = 0;
                                        _objDgv.Rows[j].Cells[3].Value = 0;
                                    }
                                    j++;
                                }
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

                if (!Registration()) { throw new Exception("登録失敗しました。"); }

                MessageBox.Show("登録完了");

                // フォームを編集不可にする
                RegistrationDateMaskedTextBox.Enabled = false;
                PersonComboBox.Enabled = false;
                RegisterButton.Enabled = false;
                for (var i = 0; i <= 9; i++) {
                    _objCbx = Controls[_checkBoxNames[i]] as CheckBox;
                    if (_objCbx != null) {
                        _objCbx.Enabled = false;
                    }

                    _objDgv = Controls[_dataGridViewNames[i]] as DataGridView;
                    if (_objDgv != null) {
                        _objDgv.Enabled = false;
                    }
                }
                // リスト印刷ボタンを有効に
                if (IsListPrint) { SubstrateListPrintButton.Enabled = true; }

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private bool QuantityCheck() {
            try {
                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateが空です"); }
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {

                            var objCbx = Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");
                            objCbx.Enabled = true;
                            objCbx.Checked = true;

                            var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                            objDgv.Columns[3].ReadOnly = false;
                            objDgv.Columns[4].ReadOnly = false;

                            if (objCbx.Checked) {
                                var intQuantityCheck = ProductInfo.Quantity;
                                var dgvRowCnt = objDgv.Rows.Count;

                                for (var j = 0; j < dgvRowCnt; j++) {
                                    var boolCbx = objDgv.Rows[j].Cells[4].Value != null && (bool)objDgv.Rows[j].Cells[4].Value;
                                    if (boolCbx) {
                                        var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value.ToString());
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value.ToString());
                                        var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[3].Value.ToString());

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
        private bool Registration() {
            try {
                ProductInfo.RegDate = RegistrationDateMaskedTextBox.Text;
                ProductInfo.Person = PersonComboBox.Text;
                ProductInfo.Comment = CommentTextBox.Text;
                if (string.IsNullOrEmpty(ProductInfo.Person)) { throw new Exception("担当者を選択してください。"); }

                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        using (SQLiteConnection con = new(GetConnectionRegistration())) {
                            con.Open();
                            using var transaction = con.BeginTransaction();

                            if (_useSubstrate == null) { throw new Exception("ArrUseSubstrateがnullです。"); }
                            for (var i = 0; i <= _useSubstrate.Length; i++) {

                                var objCbx = Controls[_checkBoxNames[i]] as CheckBox ?? throw new Exception("objCbxがnullです。");

                                if (objCbx.Checked) {
                                    var objDgv = Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvがnullです。");
                                    var dgvRowCnt = objDgv.Rows.Count;
                                    var subTotalTemp = string.Empty;

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[4].Value);
                                        if (boolCbx) {
                                            var substrateName = string.Empty;
                                            var substrateModel = string.Empty;
                                            var orderNum = string.Empty;
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                            var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);
                                            var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                            var useValue = Convert.ToInt32(objDgv.Rows[j].Cells[3].Value);

                                            using (var cmd = con.CreateCommand()) {
                                                cmd.CommandText = $"""SELECT * FROM "{ProductInfo.StockName}_StockView" WHERE SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber""";
                                                cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                                                cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;
                                                using var dr = cmd.ExecuteReader();
                                                while (dr.Read()) {
                                                    substrateName = $"{dr["SubstrateName"]}";
                                                    substrateModel = $"{dr["SubstrateModel"]}";
                                                    orderNum = $"{dr["OrderNumber"]}";
                                                }
                                            }

                                            if (useValue != 0) {
                                                subTotalTemp = string.IsNullOrEmpty(subTotalTemp) ? $"{substrateNum}({useValue})" : $"{subTotalTemp},{substrateNum}({useValue})";
                                            }

                                            // 基板テーブルを更新、できない場合は挿入
                                            using (var cmdUpdate = con.CreateCommand()) {
                                                cmdUpdate.CommandText =
                                                    $"""
                                                    UPDATE {ProductInfo.StockName}_Substrate
                                                    SET
                                                        Decrease = @Decrease,
                                                        Person = @Person,
                                                        RegDate = @RegDate,
                                                        Comment = @Comment
                                                    WHERE
                                                        SubstrateNumber = @SubstrateNumber
                                                    AND
                                                        UseID = @UseID
                                                    """;

                                                cmdUpdate.Parameters.Add("@Decrease", DbType.String).Value = 0 - useValue;
                                                cmdUpdate.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                                cmdUpdate.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                                cmdUpdate.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                                                cmdUpdate.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                                cmdUpdate.Parameters.Add("@UseID", DbType.String).Value = ProductInfo.ProductID;

                                                var affectedRows = cmdUpdate.ExecuteNonQuery();

                                                if (affectedRows == 0) {
                                                    // 基板テーブルを挿入
                                                    using var cmdInsert = con.CreateCommand();
                                                    cmdInsert.CommandText =
                                                    $"""
                                                    INSERT INTO "{ProductInfo.StockName}_Substrate"
                                                        (
                                                        SubstrateName,
                                                        SubstrateModel,
                                                        SubstrateNumber,
                                                        OrderNumber,
                                                        Decrease,
                                                        UsedProductType,
                                                        UsedProductNumber,
                                                        UsedOrderNumber,
                                                        Person,
                                                        RegDate,
                                                        Comment,
                                                        UseID
                                                        )
                                                    VALUES
                                                        (
                                                        @SubstrateName,
                                                        @SubstrateModel,
                                                        @SubstrateNumber,
                                                        @OrderNumber,
                                                        @Decrease,
                                                        @UsedProductType,
                                                        @UsedProductNumber,
                                                        @UsedOrderNumber,
                                                        @Person,
                                                        @RegDate,
                                                        @Comment,
                                                        @UseID
                                                        )
                                                    """;

                                                    cmdInsert.Parameters.Add("@SubstrateName", DbType.String).Value = string.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                                                    cmdInsert.Parameters.Add("@SubstrateModel", DbType.String).Value = string.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                                                    cmdInsert.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                                    cmdInsert.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                                                    cmdInsert.Parameters.Add("@Decrease", DbType.String).Value = 0 - useValue;
                                                    cmdInsert.Parameters.Add("@UsedProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                                    cmdInsert.Parameters.Add("@UsedProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                                    cmdInsert.Parameters.Add("@UsedOrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                                    cmdInsert.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                                    cmdInsert.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                                    cmdInsert.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                                                    cmdInsert.Parameters.Add("@UseID", DbType.String).Value = ProductInfo.ProductID;

                                                    cmdInsert.ExecuteNonQuery();
                                                }
                                            }

                                            if (IsListPrint) {
                                                _listUsedSubstrate.Add(_useSubstrate[i]);
                                                if (substrateNum != null) { _listUsedProductNumber.Add(substrateNum); }
                                                _listUsedQuantity.Add(useValue);
                                            }
                                        }
                                    }
                                    _totalSubstrate = string.IsNullOrEmpty(_totalSubstrate)
                                        ? $"[{_useSubstrate[i]}]{subTotalTemp}"
                                        : $"{_totalSubstrate},{Environment.NewLine}[{_useSubstrate[i]}]{subTotalTemp}";
                                    subTotalTemp = string.Empty;
                                }
                            }

                            // 製品テーブル更新
                            using (var cmd = con.CreateCommand()) {
                                cmd.CommandText =
                                    $"""
                                    UPDATE {ProductInfo.ProductName}_Product
                                    SET
                                        Quantity = @Quantity,
                                        Person = @Person,
                                        RegDate = @RegDate,
                                        Revision = @Revision,
                                        RevisionGroup = @RevisionGroup,
                                        SerialLast = @SerialLast,
                                        SerialLastNumber = @SerialLastNumber,
                                        Comment = @Comment,
                                        UsedSubstrate = @UsedSubstrate
                                    WHERE
                                        ProductNumber = @ProductNumber
                                    AND
                                        SerialFirst = @SerialFirst
                                    """;

                                cmd.Parameters.Add("@ProductType", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductType) ? DBNull.Value : ProductInfo.ProductType;
                                cmd.Parameters.Add("@ProductModel", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductModel) ? DBNull.Value : ProductInfo.ProductModel;
                                cmd.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.OrderNumber) ? DBNull.Value : ProductInfo.OrderNumber;
                                cmd.Parameters.Add("@ProductNumber", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.ProductNumber) ? DBNull.Value : ProductInfo.ProductNumber;
                                cmd.Parameters.Add("@Quantity", DbType.String).Value = ProductInfo.Quantity;
                                cmd.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                cmd.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                cmd.Parameters.Add("@Revision", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Revision) ? DBNull.Value : ProductInfo.Revision;
                                cmd.Parameters.Add("@RevisionGroup", DbType.String).Value = ProductInfo.RevisionGroup;
                                cmd.Parameters.Add("@SerialFirst", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialFirst) ? DBNull.Value : ProductInfo.SerialFirst;
                                cmd.Parameters.Add("@SerialLast", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.SerialLast) ? DBNull.Value : ProductInfo.SerialLast;
                                cmd.Parameters.Add("@SerialLastNumber", DbType.String).Value = ProductInfo.SerialLastNumber;
                                cmd.Parameters.Add("@UsedSubstrate", DbType.String).Value = string.IsNullOrWhiteSpace(_totalSubstrate) ? DBNull.Value : _totalSubstrate;
                                cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

                                cmd.ExecuteNonQuery();
                                transaction.Commit();
                            }
                        }
                        break;
                }

                // バックアップ作成
                BackupManager.CreateBackup();
                // ログ出力
                Logger.AppendLog($";[基板変更];注文番号[{ProductInfo.OrderNumber}];製造番号[{ProductInfo.ProductNumber}];製品名[{ProductInfo.ProductName}];タイプ[{ProductInfo.ProductType}];型式[{ProductInfo.ProductModel}];数量[{ProductInfo.Quantity}];シリアル先頭[{ProductInfo.SerialFirst}];シリアル末尾[{ProductInfo.SerialLast}];Revision[{ProductInfo.Revision}];登録日[{ProductInfo.RegDate}];担当者[{ProductInfo.Person}];");
                return true;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // リスト印刷
        private void GenerationList() {
            try {
                ExcelGenerater.GenerationList(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static class ExcelHelper {
            //public static string GetCellValueOrDefault(IXLWorksheet sheet, int rowIndex, int colIndex, int? defaultRow = null) {
            //    // 対象セルの値を取得
            //    var value = sheet.Cell(rowIndex, colIndex).GetString();

            //    // 値が存在しない場合の処理
            //    if (string.IsNullOrEmpty(value)) {
            //        // デフォルト行が指定されていない場合は空文字を返す
            //        if (defaultRow is null or 0) {
            //            return string.Empty;
            //        }

            //        // デフォルト行の値を取得
            //        return sheet.Cell(defaultRow.Value, colIndex).GetString();
            //    }

            //    // 値が存在すればその値を返す
            //    return value;
            //}
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
            checkBox.ForeColor = checkBox.Checked ? Color.Black : Color.Red;

            if (!checkBox.Checked) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void SubstrateChange2_Load(object sender, EventArgs e) { LoadEvents(); }
        private void RegisterButton_Click(object sender, EventArgs e) { ChangeRegistration(); }
        private void SubstrateListPrintButton_Click(object sender, EventArgs e) { GenerationList(); }
        private void CloseButton_Click(object sender, EventArgs e) { Close(); }
    }
}
