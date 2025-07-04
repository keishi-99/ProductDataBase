using ProductDatabase.Other;
using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Other.CommonUtils;

namespace ProductDatabase {
    public partial class SubstrateChange2 : Form {

        public ProductInformation ProductInfo { get; set; } = new ProductInformation();

        private string[] _useSubstrate = [];

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
                RevisionTextBox.Text = ProductInfo.Revision;
                CommentTextBox.Text = ProductInfo.Comment;

                // ComboBoxへ担当者を追加
                PersonComboBox.Items.AddRange([.. ProductInfo.PersonList]);

                // 使用基板リスト化+名前順にソート
                _useSubstrate = ProductInfo.UseSubstrate.Split(",");
                Array.Sort(_useSubstrate);
                var strQuantity = string.Empty;

                switch (ProductInfo.RegType) {
                    case 2:
                    case 3:
                        for (var i = 0; i <= _useSubstrate.GetUpperBound(0); i++) {
                            var substrateModel = _useSubstrate[i];
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

                            List<(string, List<string>, List<int>)> usedSubstrate = [];

                            using SQLiteConnection con = new(GetConnectionRegistration());
                            con.Open();
                            using var cmd = con.CreateCommand();

                            cmd.CommandText =
                                $"""
                                SELECT
                                	SubstrateModel,
                                	SubstrateNumber,
                                	Decrease
                                FROM
                                	{ProductInfo.CategoryName}_Substrate
                                WHERE
                                	UseID = @ID
                                ORDER BY
                                	SubstrateModel ASC
                                """;
                            cmd.Parameters.Add("@ID", DbType.String).Value = ProductInfo.ProductID;

                            using (var dr = cmd.ExecuteReader()) {
                                while (dr.Read()) {
                                    var usedSubstrateModel = dr.GetString(0);
                                    var usedSubstrateNumber = dr.GetString(1);
                                    var usedDecrease = -1 * dr.GetInt32(2);

                                    // 既存の usedSubstrateModel を検索
                                    var existingSubstrate = usedSubstrate.FirstOrDefault(x => x.Item1 == usedSubstrateModel);

                                    if (existingSubstrate != default) {
                                        // 既存の usedSubstrateModel が見つかった場合、リストに追加
                                        existingSubstrate.Item2.Add(usedSubstrateNumber);
                                        existingSubstrate.Item3.Add(usedDecrease);
                                    }
                                    else {
                                        // 既存の usedSubstrateModel が見つからなかった場合、新しいエントリを追加
                                        List<string> usedSubstrateNumbers = [usedSubstrateNumber];
                                        List<int> usedDecreases = [usedDecrease];
                                        (string, List<string>, List<int>) usedSubstrateData = (usedSubstrateModel, usedSubstrateNumbers, usedDecreases);
                                        usedSubstrate.Add(usedSubstrateData);
                                    }
                                }
                            }
                            // テーブル検索SQL - [[StockName]_StockView]テーブルから基板型式[Model]で在庫基板を抽出
                            cmd.CommandText =
                                $"""
                                SELECT
                                    SubstrateName,
                                    SubstrateModel,
                                    SubstrateNumber,
                                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                FROM
                                    {ProductInfo.CategoryName}_Substrate
                                WHERE
                                    SubstrateModel = @SubstrateModel AND SubstrateNumber NOTNULL
                                GROUP BY
                                    SubstrateName,
                                    SubstrateModel,
                                    SubstrateNumber
                                ORDER BY
                                    MIN(ID);
                                """;
                            cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = substrateModel;

                            using (var dr = cmd.ExecuteReader()) {
                                var j = 0;
                                while (dr.Read()) {
                                    // 抽出した行から製造番号,在庫取得
                                    var strSubstrateNum = $"{dr["SubstrateNumber"]}";
                                    var intStock = int.Parse($"{dr["Stock"]}");
                                    var strSubstrateName = $"{dr["SubstrateName"]}";
                                    if (_objCbx != null) { _objCbx.Text = $"{strSubstrateName} - {substrateModel}"; }

                                    // usedSubstrate から strSubstrateNum を検索
                                    var num = usedSubstrate.FindIndex(substrate => substrate.Item1 == substrateModel);
                                    var usedSubstrateItem = usedSubstrate[num].Item2
                                        .Select((num, index) => new { Num = num, Index = index })
                                        .FirstOrDefault(item => item.Num == strSubstrateNum);

                                    var strUsedSubNum = usedSubstrateItem != null ? strSubstrateNum : string.Empty;
                                    var intUsedQuantity = usedSubstrateItem != null ? usedSubstrate[num].Item3[usedSubstrateItem.Index] : 0;

                                    if (intStock > 0 || strUsedSubNum == strSubstrateNum) {
                                        if (_objDgv == null) {
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
                if (ProductInfo.IsListPrint) { SubstrateListPrintButton.Enabled = true; }

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
                ProductInfo.RegDate = RegistrationDateTimePicker.Value.ToShortDateString();
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

                                    for (var j = 0; j <= dgvRowCnt - 1; j++) {
                                        var usedValue = Convert.ToInt32(objDgv.Rows[j].Cells[2].Value);
                                        var boolCbx = Convert.ToBoolean(objDgv.Rows[j].Cells[4].Value);
                                        var useValue = boolCbx ? Convert.ToInt32(objDgv.Rows[j].Cells[3].Value) : 0;
                                        if (boolCbx) {
                                            var substrateName = string.Empty;
                                            var substrateModel = string.Empty;
                                            var orderNum = string.Empty;
                                            var substrateNum = objDgv.Rows[j].Cells[0].Value.ToString() ?? string.Empty;
                                            var stockValue = Convert.ToInt32(objDgv.Rows[j].Cells[1].Value);

                                            using (var cmd = con.CreateCommand()) {
                                                cmd.CommandText =
                                                    $"""
                                                    SELECT
                                                        SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                                                    FROM
                                                        {ProductInfo.CategoryName}_Substrate
                                                    WHERE
                                                        SubstrateModel = @SubstrateModel AND SubstrateNumber = @SubstrateNumber
                                                    GROUP BY
                                                        SubstrateName, SubstrateModel, SubstrateNumber, OrderNumber
                                                    ORDER BY
                                                        MIN(ID);
                                                    """;
                                                cmd.Parameters.Add("@SubstrateModel", DbType.String).Value = _useSubstrate[i];
                                                cmd.Parameters.Add("@SubstrateNumber", DbType.String).Value = substrateNum;
                                                using var dr = cmd.ExecuteReader();
                                                while (dr.Read()) {
                                                    substrateName = $"{dr["SubstrateName"]}";
                                                    substrateModel = $"{dr["SubstrateModel"]}";
                                                    orderNum = $"{dr["OrderNumber"]}";
                                                }
                                            }
                                            // 基板テーブルを更新、できない場合は挿入
                                            using (var cmdUpdate = con.CreateCommand()) {
                                                cmdUpdate.CommandText =
                                                    $"""
                                                    UPDATE
                                                        {ProductInfo.CategoryName}_Substrate
                                                    SET
                                                        Decrease = @Decrease,Person = @Person,RegDate = @RegDate,Comment = @Comment
                                                    WHERE
                                                        SubstrateNumber = @SubstrateNumber
                                                    AND
                                                        UseID = @UseID
                                                    """;

                                                cmdUpdate.Parameters.Add("@Decrease", DbType.Int64).Value = 0 - useValue;
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
                                                    INSERT INTO "{ProductInfo.CategoryName}_Substrate"
                                                        (StockName,SubstrateName,SubstrateModel,SubstrateNumber,OrderNumber,Decrease,
                                                        Person,RegDate,Comment,UseID)
                                                    VALUES
                                                        (@StockName,@SubstrateName,@SubstrateModel,@SubstrateNumber,@OrderNumber,
                                                        @Decrease,@Person,@RegDate,@Comment,@UseID)
                                                    """;

                                                    cmdInsert.Parameters.Add("@StockName", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.StockName) ? DBNull.Value : ProductInfo.StockName;
                                                    cmdInsert.Parameters.Add("@SubstrateName", DbType.String).Value = string.IsNullOrWhiteSpace(substrateName) ? DBNull.Value : substrateName;
                                                    cmdInsert.Parameters.Add("@SubstrateModel", DbType.String).Value = string.IsNullOrWhiteSpace(substrateModel) ? DBNull.Value : substrateModel;
                                                    cmdInsert.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                                    cmdInsert.Parameters.Add("@OrderNumber", DbType.String).Value = string.IsNullOrWhiteSpace(orderNum) ? DBNull.Value : orderNum;
                                                    cmdInsert.Parameters.Add("@Decrease", DbType.Int64).Value = 0 - useValue;
                                                    cmdInsert.Parameters.Add("@Person", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Person) ? DBNull.Value : ProductInfo.Person;
                                                    cmdInsert.Parameters.Add("@RegDate", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.RegDate) ? DBNull.Value : ProductInfo.RegDate;
                                                    cmdInsert.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;
                                                    cmdInsert.Parameters.Add("@UseID", DbType.String).Value = ProductInfo.ProductID;

                                                    cmdInsert.ExecuteNonQuery();
                                                }
                                            }

                                            if (ProductInfo.IsListPrint) {
                                                _listUsedSubstrate.Add(_useSubstrate[i]);
                                                if (substrateNum != null) { _listUsedProductNumber.Add(substrateNum); }
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
                                                        "{ProductInfo.CategoryName}_Substrate"
                                                    WHERE
                                                        SubstrateNumber = @SubstrateNumber AND UseID = @ID;
                                                    """;
                                            cmdDelete.Parameters.Clear(); // パラメータをクリア
                                            cmdDelete.Parameters.Add("@SubstrateNumber", DbType.String).Value = objDgv.Rows[j].Cells[0].Value;
                                            cmdDelete.Parameters.Add("@ID", DbType.Int64).Value = ProductInfo.ProductID;

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
                                        {ProductInfo.CategoryName}_Product
                                    SET
                                        Quantity = @Quantity,Person = @Person,RegDate = @RegDate,Revision = @Revision,RevisionGroup = @RevisionGroup,SerialLast = @SerialLast,
                                        SerialLastNumber = @SerialLastNumber,Comment = @Comment
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
                                cmd.Parameters.Add("@Comment", DbType.String).Value = string.IsNullOrWhiteSpace(ProductInfo.Comment) ? DBNull.Value : ProductInfo.Comment;

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
                    $"[{ProductInfo.CategoryName}]",
                    $"ID[{ProductInfo.ProductID}]",
                    $"注文番号[{ProductInfo.OrderNumber}]",
                    $"製造番号[{ProductInfo.ProductNumber}]",
                    $"製品名[{ProductInfo.ProductName}]",
                    $"タイプ[{ProductInfo.ProductType}]",
                    $"型式[{ProductInfo.ProductModel}]",
                    $"数量[{ProductInfo.Quantity}]",
                    $"シリアル先頭[{ProductInfo.SerialFirst}]",
                    $"シリアル末尾[{ProductInfo.SerialLast}]",
                    $"Revision[{ProductInfo.Revision}]",
                    $"登録日[{ProductInfo.RegDate}]",
                    $"担当者[{ProductInfo.Person}]",
                    $"コメント[{ProductInfo.Comment}]"
                ];
                CommonUtils.Logger.AppendLog(logMessageArray);
            } catch (Exception) {
                throw;
            }
        }

        // リスト印刷
        private void GenerateList() {
            try {
                ListGenerator.GenerateList(ProductInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
