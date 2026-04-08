using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.ExcelService;
using ProductDatabase.Other;
using ProductDatabase.Print;
using System.Data;
using ProductDatabase.Models;
using static ProductDatabase.Data.ProductRepository;
using static ProductDatabase.Other.CommonUtils;
using static ProductDatabase.Print.PrintManager;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase {

    public partial class ProductRegistration2Window : Form {

        public DocumentPrintSettings ProductPrintSettings { get; set; } = new DocumentPrintSettings();
        public LabelPrintSettings LabelPrintSettings => ProductPrintSettings.LabelPrintSettings ?? new LabelPrintSettings();
        public BarcodePrintSettings BarcodePrintSettings => ProductPrintSettings.BarcodePrintSettings ?? new BarcodePrintSettings();
        public NameplatePrintSettings NameplatePrintSettings => ProductPrintSettings.NameplatePrintSettings ?? new NameplatePrintSettings();

        public string PrintSettingPath { get; set; } = string.Empty;

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;
        private readonly AppSettings _appSettings;

        private int _serialLastNumber;
        private bool _isTransactionCommitted = false;

        private readonly List<string> _serialList = [];
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

        private SqliteConnection? _sqliteConnection;
        private SqliteTransaction? _sqliteTransaction;

        public ProductRegistration2Window(ProductMaster productMaster, ProductRegisterWork productRegisterWork, AppSettings appSettings) {
            InitializeComponent();
            _productMaster = productMaster;
            _productRegisterWork = productRegisterWork;
            _appSettings = appSettings;
        }

        // フォームロード時にDB接続・トランザクション開始・UIと基板データ初期化・印刷設定読み込みを行う
        private void LoadEvents() {
            try {
                _sqliteConnection = new SqliteConnection(GetConnectionRegistration());
                _sqliteConnection.Open();
                _sqliteTransaction = _sqliteConnection.BeginTransaction();
                _isTransactionCommitted = false;

                SetFont();
                InitializeUIControls();

                if (_productMaster.IsSerialGeneration) {
                    SetSerialNumbers();
                }

                switch (_productMaster.RegType) {
                    case 2:
                    case 4:
                    case 3:
                        if (_productMaster.SheetPrintType == 3) {
                            break;
                        }
                        LoadSubstrateData(_sqliteConnection, _sqliteTransaction);
                        break;
                    case 9:
                        using (ServiceForm window = new(ServiceInfo)) {
                            if (window.ShowDialog(this) != DialogResult.OK) {
                                Close();
                            }
                            ServiceInfo = window.ServiceInfo;
                        }
                        LoadSubstrateData(_sqliteConnection, _sqliteTransaction);
                        break;
                    default:
                        HideAllControls();
                        break;
                }

                ConfigurePrintSettings();

            } catch (SqliteException ex) {
                MessageBox.Show($"データベースがロックされています。: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ClosingEvents(shouldCommit: false);
                Close();
            }
        }
        // フォームクローズ時にトランザクションをコミットまたはロールバックしDB接続を解放する
        private void ClosingEvents(bool shouldCommit = true) {
            try {
                if (_sqliteTransaction is not null && !_isTransactionCommitted) {
                    if (shouldCommit) {
                        _sqliteTransaction.Commit();
                    }
                    else {
                        _sqliteTransaction.Rollback();
                    }
                    _isTransactionCommitted = true;
                }
            } finally {
                _sqliteTransaction?.Dispose();
                _sqliteTransaction = null;
                _sqliteConnection?.Close();
                _sqliteConnection?.Dispose();
                _sqliteConnection = null;
                _isTransactionCommitted = false;
            }
        }
        // 設定フォントをフォームに適用する
        private void SetFont() {
            Font = new Font(_appSettings.FontName, _appSettings.FontSize);
        }
        // 登録ボタンと成績書ボタンの初期有効状態を設定する
        private void InitializeUIControls() {
            GenerateReportButton.Enabled = false;
            RegisterButton.Enabled = true;
        }
        // 全ての基板チェックボックスとDataGridViewを非表示にする
        private void HideAllControls() {
            for (var i = 0; i <= 14; i++) {
                if (MainPanel.Controls[_checkBoxNames[i]] is CheckBox objCbx) {
                    objCbx.Visible = false;
                }

                if (MainPanel.Controls[_dataGridViewNames[i]] is DataGridView objDgv) {
                    objDgv.Visible = false;
                }
            }
        }
        // シリアル末尾番号を先頭番号と数量から算出して保持する
        private void SetSerialNumbers() {
            _serialLastNumber = _productRegisterWork.SerialFirstNumber + _productRegisterWork.Quantity - 1;
        }
        // 使用基板一覧をDBから取得してチェックボックスとDataGridViewに表示し在庫不足を警告する
        private void LoadSubstrateData(SqliteConnection connection, SqliteTransaction transaction) {
            // サービス向け登録の場合は、サービス情報を使用する
            var isServiceRegistration = _productMaster.RegType == 9;
            var useSubstrates = (isServiceRegistration ? ServiceInfo.ServiceUseSubstrates : _productMaster.UseSubstrates);

            var shortageSubstrateName = string.Empty;

            for (var i = 0; i < useSubstrates.Count; i++) {
                var objCbx = MainPanel.Controls[_checkBoxNames[i]] as CheckBox;
                var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView;

                var substrateID = useSubstrates[i].SubstrateID;
                var substrateName = useSubstrates[i].SubstrateName;
                var substrateModel = useSubstrates[i].SubstrateModel;

                SetupCheckBox(objCbx, substrateName, substrateModel);
                SetupDataGridView(objDgv);

                if (FetchAndDisplaySubstrateData(connection, objDgv, substrateID, transaction)) {
                    shortageSubstrateName += $"[{substrateName}]{Environment.NewLine}";
                }
            }

            if (!string.IsNullOrEmpty(shortageSubstrateName)) {
                Activate();
                MessageBox.Show($"在庫が足りません。{Environment.NewLine}{shortageSubstrateName}", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // チェックボックスを有効化してONにし基板名と型式を表示テキストにセットする
        private static void SetupCheckBox(CheckBox? objCbx, string substrateName, string substrateModel) {
            if (objCbx is not null) {
                objCbx.Enabled = true;
                objCbx.Checked = true;
                var splitSubstrateName = substrateName.Split(':');
                objCbx.Text = $"{splitSubstrateName.Last()} - {substrateModel}";
            }
        }
        // DataGridViewのセル書式を設定しフォントサイズに合わせた列幅を適用する
        private void SetupDataGridView(DataGridView? objDgv) {
            if (objDgv is not null) {
                objDgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                objDgv.Columns[2].ReadOnly = false;
                objDgv.Columns[3].ReadOnly = false;
                SetDataGridViewSize(objDgv);
            }
        }
        // 現在のフォントサイズに合わせてDataGridViewの行高・列幅を設定する
        private void SetDataGridViewSize(DataGridView objDgv) {
            switch (Font.Size) {
                case 9:
                    objDgv.RowTemplate.Height = 24;
                    objDgv.Columns[0].Width = 130;
                    objDgv.Columns[1].Width = 35;
                    objDgv.Columns[2].Width = 35;
                    objDgv.Columns[3].Width = 24;
                    break;
                case 12:
                    objDgv.RowTemplate.Height = 24;
                    objDgv.Columns[0].Width = 200;
                    objDgv.Columns[1].Width = 50;
                    objDgv.Columns[2].Width = 50;
                    objDgv.Columns[3].Width = 24;
                    break;
                case 14:
                    objDgv.RowTemplate.Height = 30;
                    objDgv.Columns[0].Width = 240;
                    objDgv.Columns[1].Width = 60;
                    objDgv.Columns[2].Width = 60;
                    objDgv.Columns[3].Width = 30;
                    break;
            }
        }
        // 基板IDの在庫データをDBから取得してDataGridViewに表示し在庫が0以下ならtrueを返す
        private bool FetchAndDisplaySubstrateData(SqliteConnection connection, DataGridView? objDgv, long substrateID, SqliteTransaction transaction) {
            var intQuantity = _productRegisterWork.Quantity;

            var commandText =
                $"""
                SELECT
                    SubstrateName,
                    SubstrateNumber,
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID 
                    AND SubstrateNumber IS NOT NULL
                    AND IsDeleted = 0
                GROUP BY
                    SubstrateID,
                    SubstrateNumber
                HAVING
                    Stock > 0
                ORDER BY
                    MIN(ID)
                ;
                """;

            var results = connection.Query(commandText, new { SubstrateID = substrateID }, transaction: transaction);

            foreach (var row in results) {
                var strSubstrateNumber = $"{row.SubstrateNumber}";
                var intStock = int.TryParse(row.Stock?.ToString(), out int stock) ? stock : 0;

                objDgv?.Rows.Add(strSubstrateNumber, intStock);

                if (objDgv is not null) {
                    if (intQuantity >= intStock) {
                        intQuantity -= intStock;
                        objDgv.Rows[^1].Cells[2].Value = intStock;
                        objDgv.Rows[^1].Cells[3].Value = true;
                    }
                    else {
                        if (intQuantity == 0) {
                            objDgv.Rows[^1].Cells[2].Value = 0;
                        }
                        else {
                            objDgv.Rows[^1].Cells[2].Value = intQuantity;
                            objDgv.Rows[^1].Cells[3].Value = true;
                            intQuantity = 0;
                        }
                    }
                }
            }
            return intQuantity > 0;
        }

        // 製品マスターの印刷フラグに応じて印刷ボタン・メニューの表示と有効状態を設定し印刷設定を読み込む
        private void ConfigurePrintSettings() {
            SubstrateListPrintButton.Visible = _productMaster.IsListPrint;
            CheckSheetPrintButton.Visible = _productMaster.IsCheckSheetPrint;
            NameplatePrintButton.Visible = _productMaster.IsNameplatePrint;

            SerialPrintPositionLabel.Visible = _productMaster.IsLabelPrint;
            SerialPrintPositionNumericUpDown.Visible = _productMaster.IsLabelPrint;

            BarcodePrintPositionLabel.Visible = _productMaster.IsBarcodePrint;
            BarcodePrintPositionNumericUpDown.Visible = _productMaster.IsBarcodePrint;

            シリアルラベル印刷プレビューToolStripMenuItem.Enabled = _productMaster.IsLabelPrint;
            シリアルラベル印刷設定ToolStripMenuItem.Enabled = _productMaster.IsLabelPrint;
            バーコード印刷プレビューToolStripMenuItem.Enabled = _productMaster.IsBarcodePrint;
            バーコード印刷設定ToolStripMenuItem.Enabled = _productMaster.IsBarcodePrint;
            銘版印刷設定ToolStripMenuItem.Enabled = _productMaster.IsNameplatePrint;

            if (_productMaster.IsSerialGeneration) {
                LoadSettings();
            }
        }
        // 製品・型式に対応するJSON印刷設定ファイルを読み込んでProductPrintSettingsに反映する
        private void LoadSettings() {
            try {
                ProductPrintSettings = new DocumentPrintSettings();
                PrintSettingPath = Path.Combine(Environment.CurrentDirectory, "config", "Product", _productMaster.CategoryName, _productMaster.ProductName, $"PrintConfig_{_productMaster.ProductName}_{_productMaster.ProductID}_{_productMaster.ProductModel}.json");
                if (!File.Exists(PrintSettingPath)) {
                    throw new DirectoryNotFoundException($"印刷用設定ファイルがありません。");
                }
                var jsonString = File.ReadAllText(PrintSettingPath);
                ProductPrintSettings = System.Text.Json.JsonSerializer.Deserialize<DocumentPrintSettings>(jsonString) ?? new DocumentPrintSettings();
            } catch (Exception ex) {
                MessageBox.Show("設定ファイルの読み込みに失敗しました:\n" + ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 重複チェック・数量検証・シリアル検証を経てDB登録・ログ記録・印刷処理を実行する
        private async Task RegisterCheck() {
            try {
                _serialList.Clear();

                if (_sqliteConnection is null || _sqliteTransaction is null) {
                    throw new InvalidOperationException("編集モード用の接続が初期化されていません。");
                }

                if (!NumberCheck(_sqliteConnection, _sqliteTransaction) || !QuantityCheck()) {
                    return;
                }
                if (_productMaster.IsSerialGeneration) {
                    SerialCheck(_sqliteConnection);
                    GenerateSerialCodes();
                }

                DisableControls();

                using (var overlay = new LoadingOverlay(this)) {
                    await Task.Run(() => {
                        Registration(_sqliteConnection, _sqliteTransaction);

                        LogRegistration(_productMaster, _productRegisterWork);
                        BackupManager.CreateBackup();

                        // 登録チェック
                        RegistrationCheck(_sqliteConnection);
                    });
                }

                // 登録完了メッセージ
                MessageBox.Show("登録しました。");

                // 印刷処理
                await HandleLabelPrinting();
                await HandleBarcodePrinting();

                HandlePostRegistration();
                GenerateReportButton.Enabled = true;

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // RegTypeに応じて製品・シリアル・基板のINSERT処理を呼び出しトランザクションをコミットする
        private void Registration(SqliteConnection connection, SqliteTransaction transaction) {
            try {
                InsertProduct(connection, transaction);

                switch (_productMaster.RegType) {
                    case 0:
                        break;
                    case 1:
                        if (_productMaster.IsSerialGeneration) {
                            InsertSerial(connection, transaction);
                        }
                        break;
                    case 2:
                    case 3:
                    case 4:
                    case 9:
                        if (_productMaster.IsSerialGeneration) {
                            InsertSerial(connection, transaction);
                        }
                        RegisterSubstrate(connection, transaction);
                        break;
                    default:
                        throw new Exception("RegType unknown");
                }

                string? nextSuffix = null;
                if (ShouldIncrementOLesSuffix) {
                    nextSuffix = GetNextOLesSuffix(_productMaster.OLesSerialSuffix);
                    connection.Execute(
                        $"UPDATE {Constants.ProductTableName} SET OLesSerialSuffix = @Suffix WHERE ProductID = @ProductID;",
                        new { Suffix = nextSuffix, _productMaster.ProductID },
                        transaction: transaction);
                }

                transaction.Commit();
                _isTransactionCommitted = true;

                if (nextSuffix != null) {
                    _productMaster.OLesSerialSuffix = nextSuffix;
                }

            } catch (Exception) {
                transaction.Rollback();
                _isTransactionCommitted = true;
                MessageBox.Show("登録に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        // 製品履歴テーブルにレコードをINSERTして生成されたROWIDを作業データに保存する
        private void InsertProduct(SqliteConnection connection, SqliteTransaction transaction) {
            var commandText =
                $"""
                INSERT INTO {Constants.TProductTableName}
                (
                    ProductID,
                    OrderNumber,
                    ProductNumber,
                    OLesNumber,
                    Quantity,
                    Person, 
                    RegDate,
                    Revision, 
                    RevisionGroup,
                    SerialFirst,
                    SerialLast, 
                    SerialLastNumber, 
                    Comment
                )
                VALUES 
                (
                    @ProductID,
                    @OrderNumber,
                    @ProductNumber,
                    @OLesNumber,
                    @Quantity, 
                    @Person,
                    @RegDate,
                    @Revision,
                    @RevisionGroup,
                    @SerialFirst, 
                    @SerialLast, 
                    @SerialLastNumber, 
                    @Comment
                )
                ;
                """;

            var comment = _productMaster.RegType switch {
                9 => $"製品ID[{ServiceInfo.ServiceProductID}],製品名[{ServiceInfo.ServiceProductName}],製品型式[{ServiceInfo.ServiceProductModel}]\n{_productRegisterWork.Comment}",
                _ => _productRegisterWork.Comment,
            };

            connection.Execute(commandText, new {
                _productMaster.ProductID,
                OrderNumber = _productRegisterWork.OrderNumber.NullIfWhiteSpace(),
                ProductNumber = _productRegisterWork.ProductNumber.NullIfWhiteSpace(),
                OLesNumber = _productRegisterWork.OLesNumber.NullIfWhiteSpace(),
                _productRegisterWork.Quantity,
                Person = _productRegisterWork.Person.NullIfWhiteSpace(),
                RegDate = _productRegisterWork.RegDate.NullIfWhiteSpace(),
                Revision = _productRegisterWork.Revision.NullIfWhiteSpace(),
                _productMaster.RevisionGroup,
                SerialFirst = _productRegisterWork.SerialFirst.NullIfWhiteSpace(),
                SerialLast = _productRegisterWork.SerialLast.NullIfWhiteSpace(),
                SerialLastNumber = _productMaster.IsSerialGeneration ? (int?)_serialLastNumber : null,
                Comment = comment.NullIfWhiteSpace()
            }, transaction: transaction);

            _productRegisterWork.RowID = connection.ExecuteScalar<int>("SELECT last_insert_rowid();", transaction: transaction);
        }
        // シリアルリストの各シリアルをシリアルテーブルにINSERTする（フォーマット2有効時はOLesSerialも記録）
        private void InsertSerial(SqliteConnection connection, SqliteTransaction transaction) {
            var commandText =
                $"""
                INSERT INTO {Constants.TSerialTableName}
                (
                    Serial,
                    OLesSerial,
                    UsedID,
                    ProductName
                )
                VALUES
                (
                    @Serial,
                    @OLesSerial,
                    @productRowId,
                    @ProductName
                )
                ;
                """;

            var olesSerialList = (_productMaster.IsOLesLabelPrint && !string.IsNullOrEmpty(LabelPrintSettings.OLesLabelTextFormat))
                ? GenerateSerialListForType(SerialType.OLesLabel)
                : null;

            var serialData = _serialList.Select((serial, i) => new {
                Serial = serial,
                OLesSerial = (olesSerialList != null && i < olesSerialList.Count) ? olesSerialList[i] : null,
                productRowId = _productRegisterWork.RowID,
                _productMaster.ProductName
            }).ToList();

            connection.Execute(commandText, serialData, transaction: transaction);
        }
        // DataGridViewでチェックされた基板番号・使用数を読み取り基板履歴テーブルにINSERTする
        private void RegisterSubstrate(SqliteConnection connection, SqliteTransaction transaction) {
            // サービス向け登録の場合は、サービス情報を使用する
            var isServiceRegistration = _productMaster.RegType == 9;
            var useSubstrates = (isServiceRegistration ? ServiceInfo.ServiceUseSubstrates : _productMaster.UseSubstrates);
            long? useID = _productRegisterWork.RowID;

            for (var i = 0; i < useSubstrates.Count; i++) {
                if (!(MainPanel.Controls[_checkBoxNames[i]] as CheckBox)?.Checked ?? true) {
                    continue;
                }

                if (MainPanel.Controls[_dataGridViewNames[i]] as DataGridView is not DataGridView objDgv) {
                    throw new Exception("DataGridViewが nullです。");
                }

                var substrateID = useSubstrates[i].SubstrateID;

                foreach (DataGridViewRow row in objDgv.Rows) {
                    if (row.Cells[3].Value is not bool isChecked || !isChecked) {
                        continue;
                    }

                    var substrateNumber = row.Cells[0].Value.ToString() ?? string.Empty;
                    var useValue = int.TryParse(row.Cells[2].Value?.ToString(), out var useVal) ? useVal : 0;
                    var orderNumber = GetSubstrateInfo(connection, substrateID, substrateNumber, transaction);
                    InsertSubstrate(connection, substrateID, substrateNumber, orderNumber, useValue, useID, transaction);
                }
            }
        }
        // 基板IDと製造番号から在庫情報を取得して注文番号を返す
        private string GetSubstrateInfo(SqliteConnection connection, long substrateID, string substrateNumber, SqliteTransaction transaction) {
            var commandText =
                $"""
                SELECT
                    SubstrateID,
                    SubstrateName,
                    SubstrateModel,
                    SubstrateNumber, 
                    OrderNumber, 
                    SUM(COALESCE(Increase, 0) + COALESCE(Decrease, 0) + COALESCE(Defect, 0)) AS Stock
                FROM
                    {Constants.VSubstrateTableName}
                WHERE
                    SubstrateID = @SubstrateID
                    AND SubstrateNumber = @SubstrateNumber
                    AND IsDeleted = 0
                GROUP BY
                    OrderNumber
                ORDER
                    BY ID ASC LIMIT 1
                ;
                """;

            var result = connection.QueryFirstOrDefault<SubstrateStockInfo>(
                commandText,
                new { SubstrateID = substrateID, SubstrateNumber = substrateNumber },
                transaction: transaction);

            return result
                ?.OrderNumber
                ?? string.Empty;
        }
        private class SubstrateStockInfo {
            public int SubstrateID { get; set; }
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public string SubstrateNumber { get; set; } = string.Empty;
            public string OrderNumber { get; set; } = string.Empty;
            public int Stock { get; set; }
        }
        // 基板使用履歴テーブルに使用数（Decrease）をINSERTして製品IDと紐づける
        private void InsertSubstrate(SqliteConnection connection, long substrateID, string substrateNumber, string orderNumber, long useValue, long? useID, SqliteTransaction transaction) {
            var commandText =
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

            var comment = _productRegisterWork.Comment;

            connection.Execute(commandText, new {
                SubstrateID = substrateID,
                SubstrateNumber = substrateNumber.NullIfWhiteSpace(),
                OrderNumber = orderNumber.NullIfWhiteSpace(),
                Decrease = 0 - useValue,
                Person = _productRegisterWork.Person.NullIfWhiteSpace(),
                RegDate = _productRegisterWork.RegDate.NullIfWhiteSpace(),
                Comment = comment.NullIfWhiteSpace(),
                UseID = useID
            }, transaction: transaction);
        }

        // 登録したIDがDB上に存在するか確認し見つからない場合は例外をスローする
        private void RegistrationCheck(SqliteConnection connection) {
            var exists = connection.ExecuteScalar<bool>(
                $@"SELECT EXISTS(SELECT 1 FROM {Constants.VProductTableName} WHERE Id = @Id);",
                new { Id = _productRegisterWork.RowID });

            if (!exists) {
                throw new Exception("登録に失敗しました。IDが見つかりません。");
            }
        }
        // 製品登録の操作内容をログファイルに記録する
        private static void LogRegistration(ProductMaster productMaster, ProductRegisterWork productRegisterWork) {
            string[] logMessageArray = [
                $"[製品登録]",
                $"[{productMaster.CategoryName}]",
                $"ID[{productRegisterWork.RowID}]",
                $"注文番号[{productRegisterWork.OrderNumber}]",
                $"製造番号[{productRegisterWork.ProductNumber}]",
                $"OLes番号[{productRegisterWork.OLesNumber}]",
                $"製品名[{productMaster.ProductName}]",
                $"タイプ[{productMaster.ProductType}]",
                $"型式[{productMaster.ProductModel}]",
                $"数量[{productRegisterWork.Quantity}]",
                $"シリアル先頭[{productRegisterWork.SerialFirst}]",
                $"シリアル末尾[{productRegisterWork.SerialLast}]",
                $"Revision[{productRegisterWork.Revision}]",
                $"登録日[{productRegisterWork.RegDate}]",
                $"担当者[{productRegisterWork.Person}]",
                $"コメント[{productRegisterWork.Comment}]"
            ];
            CommonUtils.Logger.AppendLog(logMessageArray);
        }

        // 製番と注文番号の重複登録をDBで確認しユーザーに続行可否を確認する
        private bool NumberCheck(SqliteConnection connection, SqliteTransaction transaction) {

            if (!string.IsNullOrEmpty(_productRegisterWork.ProductNumber)) {
                // 製番が新規かチェック
                var productNumberQuery = $@"SELECT ProductModel FROM {Constants.VProductTableName} WHERE ProductName = @ProductName AND ProductNumber = @ProductNumber AND IsDeleted = 0 LIMIT 1";
                var existingModel = connection.QueryFirstOrDefault<string>(
                    productNumberQuery,
                    new { _productMaster.ProductName, _productRegisterWork.ProductNumber },
                    transaction: transaction); if (existingModel != null) {
                    if (existingModel == _productMaster.ProductModel) {
                        Activate();
                        var result = MessageBox.Show($"製番[{_productRegisterWork.ProductNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                    else {
                        Activate();
                        MessageBox.Show($"[{_productRegisterWork.ProductNumber}]は[{existingModel}]として登録があります。確認してください。", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }

            if (!string.IsNullOrEmpty(_productRegisterWork.OrderNumber)) {
                // 注文番号が新規かチェック
                var orderNumberQuery = $@"SELECT ProductModel FROM {Constants.VProductTableName} WHERE ProductName = @ProductName AND OrderNumber = @OrderNumber AND IsDeleted = 0 LIMIT 1";
                var existingModel = connection.QueryFirstOrDefault<string>(
                    orderNumberQuery,
                    new { _productMaster.ProductName, _productRegisterWork.OrderNumber },
                    transaction: transaction); if (existingModel != null) {
                    if (existingModel == _productMaster.ProductModel) {
                        Activate();
                        var result = MessageBox.Show($"注文番号[{_productRegisterWork.OrderNumber}]は過去に登録があります。再度登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                    else {
                        Activate();
                        var result = MessageBox.Show($"[{_productRegisterWork.OrderNumber}]は[{existingModel}]として登録があります。登録しますか？", "", MessageBoxButtons.YesNo);
                        if (result == DialogResult.No) {
                            return false;
                        }
                    }
                }
            }

            return true;
        }
        // RegTypeと選択された基板数量の整合性を検証する
        private bool QuantityCheck() {
            try {
                switch (_productMaster.RegType) {
                    case 0:
                    case 1:
                        break;
                    case 2:
                    case 3:
                    case 4:
                    case 9:
                        var useSubstrate = _productMaster.RegType != 9
                            ? _productMaster.UseSubstrates
                            : ServiceInfo.ServiceUseSubstrates;
                        for (var i = 0; i <= useSubstrate.Count - 1; i++) {

                            var objCbx = MainPanel.Controls[_checkBoxNames[i]] as System.Windows.Forms.CheckBox ?? throw new Exception("objCbxが nullです。");

                            var objDgv = MainPanel.Controls[_dataGridViewNames[i]] as DataGridView ?? throw new Exception("objDgvが nullです。");
                            objDgv.Columns[2].ReadOnly = false;
                            objDgv.Columns[3].ReadOnly = false;

                            if (objCbx.Checked) {
                                var intQuantityCheck = _productRegisterWork.Quantity;
                                var dgvRowCnt = objDgv.Rows.Count;

                                for (var j = 0; j < dgvRowCnt; j++) {
                                    var boolCbx = objDgv.Rows[j].Cells[3].Value is not null && (bool)objDgv.Rows[j].Cells[3].Value;
                                    if (boolCbx) {
                                        var stockValue = int.TryParse(objDgv.Rows[j].Cells[1].Value?.ToString(), out var sv) ? sv : 0;
                                        if (objDgv.Rows[j].Cells[2].Value is null) {
                                            throw new Exception("使用数が入力されていません。");
                                        }
                                        var useValue = int.TryParse(objDgv.Rows[j].Cells[2].Value?.ToString(), out var uv) ? uv : 0;
                                        if (useValue <= 0) {
                                            throw new Exception("使用数が0以下になっています。");
                                        }

                                        if (stockValue < useValue) {
                                            throw new Exception("在庫より多い数量が入力されています。");
                                        }
                                        intQuantityCheck -= useValue;
                                    }
                                }

                                if (intQuantityCheck != 0) {
                                    if (_productMaster.RegType != 9) {
                                        throw new Exception("入力された数量の合計が必要数と一致しません。");
                                    }
                                    else {
                                        DialogResult result = MessageBox.Show(
                                            "使用数量の合計が製品数と一致しませんがよろしいですか？",
                                            "",
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Exclamation,
                                            MessageBoxDefaultButton.Button2);
                                        if (result == DialogResult.No) {
                                            return false;
                                        }
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
        // シリアルリストを生成しDBに既存シリアルがないか重複チェックする
        private void SerialCheck(SqliteConnection connection, bool print = true) {

            _serialList.Clear();
            CurrentSerialType = GetSerialTypeFromProductMaster();

            for (var i = 0; i < _productRegisterWork.Quantity; i++) {
                _serialList.Add(GenerateCode(_productRegisterWork.SerialFirstNumber + i));
            }


            if (!print) {
                return;
            }

            List<string> strSerialDuplication = [];

            var sql =
                $"""
                SELECT
                    s.Serial
                FROM
                    {Constants.TSerialTableName} AS s
                LEFT JOIN
                    {Constants.VProductTableName} AS p
                ON
                    s.UsedID = p.ID
                WHERE
                    s.ProductName = @ProductName
                AND
                    s.Serial IN @SerialList
                ;
                """;

            var duplicatedSerials = connection.Query<string>(
                sql,
                new {
                    _productMaster.ProductName,
                    SerialList = _serialList.Select(x => x.Trim()).ToList()
                });

            var list = duplicatedSerials.ToList();

            if (list.Count > 0) {
                var message = string.Join(Environment.NewLine, list);
                throw new Exception($"{message}{Environment.NewLine}は既に使用されているシリアルです。");
            }

            // O-Les シリアルの重複チェック
            if (_productMaster.IsOLesLabelPrint && !string.IsNullOrEmpty(LabelPrintSettings.OLesLabelTextFormat)) {
                var olesList = GenerateSerialListForType(SerialType.OLesLabel)
                    .Select(x => x.Trim())
                    .ToList();

                var olesSql =
                    $"""
                    SELECT
                        s.OLesSerial
                    FROM
                        {Constants.TSerialTableName} AS s
                    WHERE
                        s.ProductName = @ProductName
                    AND
                        s.OLesSerial IN @OlesList
                    ;
                    """;

                var duplicatedOlesSerials = connection.Query<string>(
                    olesSql,
                    new {
                        _productMaster.ProductName,
                        OlesList = olesList
                    }).ToList();

                if (duplicatedOlesSerials.Count > 0) {
                    var message = string.Join(Environment.NewLine, duplicatedOlesSerials);
                    throw new Exception($"{message}{Environment.NewLine}は既に使用されているO-Lesシリアルです。");
                }
            }
        }
        // 製品マスターの印刷フラグからDB保存基準のシリアルタイプを決定する（Label > Barcode > Nameplate）
        private SerialType GetSerialTypeFromProductMaster() {
            if (_productMaster.IsLabelPrint) { return SerialType.Label; }
            if (_productMaster.IsBarcodePrint) { return SerialType.Barcode; }
            if (_productMaster.IsNameplatePrint) { return SerialType.Nameplate; }
            return SerialType.Label;
        }
        // IsLabelPrintが有効な場合にシリアルラベルを印刷する
        // IsOLesLabelPrint かつ枚数>0 の場合は Label・OLes を交互に含む1つの印刷ジョブとして実行する
        private async Task HandleLabelPrinting() {
            if (_productMaster.IsLabelPrint) {
                MessageBox.Show("シリアルラベルを印刷します。");
                CurrentSerialType = SerialType.Label;
                var labelList = GenerateLabelList();
                if (!await Print(true, labelList, GenerateOlesListOrNull())) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        // IsBarcodePrintが有効な場合にバーコード専用リストを生成してバーコードラベルの印刷を実行する
        private async Task HandleBarcodePrinting() {
            if (_productMaster.IsBarcodePrint) {
                MessageBox.Show("バーコードラベルを印刷します。");
                CurrentSerialType = SerialType.Barcode;
                var barcodeList = GenerateSerialListForType(SerialType.Barcode);

                if (!await Print(true, barcodeList)) {
                    throw new OperationCanceledException("キャンセルしました。");
                }
            }
        }
        // 登録ボタンと印刷位置入力コントロールを無効化して二重登録を防止する
        private void DisableControls() {
            RegisterButton.Enabled = false;
            SerialPrintPositionNumericUpDown.Enabled = false;
            BarcodePrintPositionNumericUpDown.Enabled = false;
        }
        // 指定タイプのフォーマットでシリアルリストを生成して返す（CurrentSerialType を変更しない）
        private List<string> GenerateSerialListForType(SerialType type) {
            return [.. Enumerable.Range(0, _productRegisterWork.Quantity).Select(i => GenerateCode(_productRegisterWork.SerialFirstNumber + i, type))];
        }
        private List<string> GenerateLabelList() => GenerateSerialListForType(SerialType.Label);
        private List<string>? GenerateOlesListOrNull() =>
            _productMaster.IsOLesLabelPrint ? GenerateSerialListForType(SerialType.OLesLabel) : null;
        // シリアルタイプを決定してシリアル先頭・末尾のコード文字列を生成する
        private void GenerateSerialCodes() {
            CurrentSerialType = GetSerialTypeFromProductMaster();
            _productRegisterWork.SerialFirst = GenerateCode(_productRegisterWork.SerialFirstNumber);
            _productRegisterWork.SerialLast = GenerateCode(_serialLastNumber);
        }
        // 登録完了後に基板コントロールを無効化してリスト・チェックシート・銘版ボタンを有効化する
        private void HandlePostRegistration() {
            foreach (Control control in MainPanel.Controls) {
                switch (control) {
                    case DataGridView dgv:
                        dgv.Enabled = false;
                        break;
                    case System.Windows.Forms.CheckBox chk:
                        chk.Enabled = false;
                        break;
                }
            }
            SubstrateListPrintButton.Enabled = _productMaster.IsListPrint;
            CheckSheetPrintButton.Enabled = _productMaster.IsCheckSheetPrint;
            NameplatePrintButton.Enabled = _productMaster.IsNameplatePrint;
        }

        // サービス向け用処理
        public class ServiceInformation {
            public DataTable ServiceDataTable { get; } = new();
            public long ServiceProductID { get; set; }
            public string ServiceCategoryName { get; set; } = string.Empty;
            public string ServiceProductName { get; set; } = string.Empty;
            public string ServiceProductType { get; set; } = string.Empty;
            public string ServiceProductModel { get; set; } = string.Empty;
            public List<SubstrateInfo> ServiceUseSubstrates { get; set; } = [];
        }
        public ServiceInformation ServiceInfo { get; set; } = new();

        // isPrintがtrueなら印刷ダイアログ経由で印刷しfalseならプレビューを表示する
        // serialList を渡した場合はそのリストを使用する（null の場合は _serialList を使用）
        // olesLabelList を渡した場合は OLes 交互印刷として PrintManager に委譲する
        private async Task<bool> Print(bool isPrint, List<string>? serialList = null, List<string>? olesLabelList = null) {
            try {
                // PrintDocumentオブジェクトの作成
                using System.Drawing.Printing.PrintDocument pd = new();
                var isPreview = !isPrint;
                var printList = serialList ?? _serialList;

                var startLine = CurrentSerialType switch {
                    SerialType.Label => (int)SerialPrintPositionNumericUpDown.Value - 1,
                    SerialType.Barcode => (int)BarcodePrintPositionNumericUpDown.Value - 1,
                    _ => throw new Exception("SerialPrintType unknown")
                };

                pd.BeginPrint += (sender, e) => {
                    PrintManager.ProductInitialize(_productMaster, _productRegisterWork, ProductPrintSettings, printList, olesLabelList);
                };
                pd.PrintPage += (sender, e) => {
                    var hasMore = PrintManager.PrintSerialCommon(e, isPreview, startLine, CurrentSerialType);
                    e.HasMorePages = hasMore;
                };

                switch (isPreview) {
                    case false:
                        var pdlg = new PrintDialog {
                            Document = pd
                        };
                        if (pdlg.ShowDialog() == DialogResult.OK) {
                            using var overlay = new LoadingOverlay(this);
                            await CommonUtils.RunOnStaThreadAsync(() => pd.Print());
                        }
                        else {
                            return false;
                        }
                        return true;
                    case true:
                        var ppd = new PrintPreviewDialog();
                        ppd.Shown += (sender, e) => {
                            var tool = (ToolStrip)ppd.Controls[1];
                            tool.Items[0].Visible = false;
                            if (sender is Form form) {
                                form.WindowState = FormWindowState.Maximized;
                            }
                        };
                        ppd.PrintPreviewControl.Zoom = 3;
                        ppd.Document = pd;
                        ppd.ShowDialog();

                        return true;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        // シリアルコードと日付情報からテキストフォーマットに従ってラベル印字コードを生成する
        // type を指定した場合はそのタイプのフォーマットを使用する（null の場合は CurrentSerialType を使用）
        private string GenerateCode(int serialCode, SerialType? type = null) {
            var regDate = DateTime.TryParse(_productRegisterWork.RegDate, out var parsedDate)
                ? parsedDate
                : DateTime.Today;

            var monthCode = CommonUtils.ToMonthCode(regDate);

            var resolvedType = type ?? CurrentSerialType;
            var outputCode = resolvedType switch {
                SerialType.Label => LabelPrintSettings.LabelTextFormat ?? string.Empty,
                SerialType.Barcode => BarcodePrintSettings.LabelTextFormat ?? string.Empty,
                SerialType.Nameplate => NameplatePrintSettings.NameplateTextFormat ?? string.Empty,
                SerialType.OLesLabel => LabelPrintSettings.OLesLabelTextFormat ?? string.Empty,
                _ => string.Empty
            };

            var map = new Dictionary<string, string> {
                ["{T}"] = _productMaster.Initial,
                ["{OT}"] = _productMaster.OLesInitial,
                ["{Y}"] = regDate.ToString("yy"),
                ["{MM}"] = regDate.ToString("MM"),
                ["{R}"] = _productRegisterWork.Revision,
                ["{M}"] = monthCode[^1..],
                ["{S}"] = serialCode.ToString($"D{_productMaster.SerialDigit}"),
                ["{SA}"] = ShouldApplyNextOLesSuffix ? GetNextOLesSuffix(_productMaster.OLesSerialSuffix) : _productMaster.OLesSerialSuffix
            };

            foreach (var kv in map) {
                outputCode = outputCode.Replace(kv.Key, kv.Value);
            }

            return outputCode;
        }
        private static string GetNextOLesSuffix(string current) {
            if (string.IsNullOrWhiteSpace(current)) return "A";
            var c = char.ToUpperInvariant(current[0]);
            if (c < 'A' || c >= 'Z') return "A";  // 'Z' は 'A' へ循環、範囲外文字も 'A' にフォールバック
            return ((char)(c + 1)).ToString();
        }
        private bool ShouldIncrementOLesSuffix =>
            _productRegisterWork.IsOLesSerialSuffixIncrement
            && _productMaster.IsOLesLabelPrint
            && LabelPrintSettings.OLesLabelTextFormat?.Contains("{SA}") is true;
        // コミット前のみ次サフィックスをプレビューに反映する（コミット後はモデルが更新済みのため不要）
        private bool ShouldApplyNextOLesSuffix =>
            !_isTransactionCommitted && ShouldIncrementOLesSuffix;
        // 基板チェックボックスのOn/Offに連動して対応DataGridViewの表示と有効状態を切り替える
        private void CheckBox_CheckedChanged(object sender, EventArgs e) {
            var checkBox = (System.Windows.Forms.CheckBox)sender;
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

            if (!checkBox.Checked && !_productMaster.IsRegType9) {
                MessageBox.Show("チェックがない場合在庫から引き落とされなくなります。", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        // 登録済み製品情報をもとにExcel成績書を生成する
        private async Task GenerateReport() {
            // [UIスレッド] 前処理: Config読み込み + ファイル選択 + 保存先選択
            (string templateFilePath, string savePath, ExcelServiceClosedXml.ReportGeneratorClosedXml.ReportConfigClosedXml config)? prepared;
            try {
                prepared = ExcelServiceClosedXml.ReportGeneratorClosedXml.PrepareReport(
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
                        ExcelServiceClosedXml.ReportGeneratorClosedXml.ExecuteReport(
                            _productMaster, _productRegisterWork,
                            prepared.Value.templateFilePath,
                            prepared.Value.savePath,
                            prepared.Value.config));
                } catch (Exception ex) {
                    taskException = ex;
                }
            } // オーバーレイをここで除去してからメッセージを表示

            GenerateReportButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateReport)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else {
                MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        // 登録済み製品情報をもとにExcel製品一覧を生成する
        private async Task GenerateList() {
            SubstrateListPrintButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ExcelServiceClosedXml.ListGeneratorClosedXml.GenerateList(
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
        // 登録済み製品情報をもとにExcelチェックシートを生成する
        private async Task GenerateCheckSheet() {
            // [UIスレッド] 前処理: Config読み込み + InputDialog（温度・湿度）
            (ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.CheckSheetConfigData configData, string temperature, string humidity)? prepared;
            try {
                prepared = ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.PrepareCheckSheet(_productMaster);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (prepared is null) return; // キャンセル

            CheckSheetPrintButton.Enabled = false;
            Exception? taskException = null;
            using (var overlay = new LoadingOverlay(this)) {
                try {
                    await CommonUtils.RunOnStaThreadAsync(() =>
                        ExcelServiceClosedXml.CheckSheetGeneratorClosedXml.ExecuteCheckSheet(
                            _productMaster, _productRegisterWork,
                            prepared.Value.configData,
                            prepared.Value.temperature,
                            prepared.Value.humidity));
                } catch (Exception ex) {
                    taskException = ex;
                }
            }

            CheckSheetPrintButton.Enabled = true;
            if (taskException is not null) {
                MessageBox.Show(taskException.Message, $"[{nameof(GenerateCheckSheet)}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 現在の製品・作業データのフィールド値をListViewで確認表示するデバッグ用ダイアログを開く
        private void ShowInfo() {
            var items = new Dictionary<string, string>
                {
                    {"ProductID", $"{_productMaster.ProductID}"},
                    {"ProductName", $"{_productMaster.ProductName}"},
                    {"ProductModel", $"{_productMaster.ProductModel}"},
                    {"ProductType", $"{_productMaster.ProductType}"},
                    {"OrderNumber", $"{_productRegisterWork.OrderNumber}"},
                    {"ProductNumber", $"{_productRegisterWork.ProductNumber}"},
                    {"OLesNumber", $"{_productRegisterWork.OLesNumber}"},
                    {"Revision", $"{_productRegisterWork.Revision}"},
                    {"RegDate", $"{_productRegisterWork.RegDate}"},
                    {"Person", $"{_productRegisterWork.Person}"},
                    {"Quantity", $"{_productRegisterWork.Quantity}"},
                    {"SerialFirstNumber", $"{_productRegisterWork.SerialFirstNumber}"},
                    {"SerialLastNumber", $"{_serialLastNumber}"},
                    {"Initial", $"{_productMaster.Initial}"},
                    {"OLesInitial", $"{_productMaster.OLesInitial}"},
                    {"RegType", $"{_productMaster.RegType}"},
                    {"SerialPrintType", $"{_productMaster.SerialPrintType}"},
                    {"SheetPrintType", $"{_productMaster.SheetPrintType}"},
                    {"SerialDigit", $"{_productMaster.SerialDigit}"}
                };

            var form = new Form {
                Text = "取得情報",
                Width = 300,
                Height = 400,
                AutoSize = true,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var listView = new ListView {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("PlemolJP", _appSettings.FontSize),
            };

            listView.Columns.Add("", 0);   // 値列の幅（調整可）
            listView.Columns.Add("項目", 200, HorizontalAlignment.Right);  // 項目列の幅
            listView.Columns.Add("値", 360);   // 値列の幅（調整可）

            foreach (var kvp in items) {
                var item = new ListViewItem("");  // ダミー1列目
                item.SubItems.Add(kvp.Key);
                item.SubItems.Add(kvp.Value);
                listView.Items.Add(item);
            }
            form.Controls.Add(listView);

            form.Shown += (_, _) => {
                listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            };
            form.ShowDialog();
        }

        private void ProductRegistration2Window_Load(object sender, EventArgs e) {
            LoadEvents();
        }
        private void ProductRegistration2Window_FormClosing(object sender, FormClosingEventArgs e) {
            ClosingEvents();
        }
        private async void RegisterButton_Click(object sender, EventArgs e) {
            await RegisterCheck();
        }
        private void CloseButton_Click(object sender, EventArgs e) {
            Close();
        }
        private void SubstrateCheckBox_CheckedChanged(object sender, EventArgs e) {
            CheckBox_CheckedChanged(sender, e);
        }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) {
            Close();
        }
        private async void シリアルラベル印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Label;
            await Print(false, GenerateLabelList(), GenerateOlesListOrNull());
        }
        private async void バーコード印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Barcode;
            var barcodeList = GenerateSerialListForType(SerialType.Barcode);
            await Print(false, barcodeList);
        }
        private void シリアルラベル印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Label;
            using PrintSettingsWindow ls = new() {
                AppSettings = _appSettings
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void バーコード印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Barcode;
            using PrintSettingsWindow ls = new() {
                AppSettings = _appSettings
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void 銘版印刷設定ToolStripMenuItem_Click(object sender, EventArgs e) {
            CurrentSerialType = SerialType.Nameplate;
            using PrintSettingsWindow ls = new() {
                AppSettings = _appSettings
            };
            ls.ShowDialog(this);
            LoadSettings();
        }
        private void 取得情報ToolStripMenuItem_Click(object sender, EventArgs e) { ShowInfo(); }
        private void NamePlatePrintButton_Click(object sender, EventArgs e) {
            var nameplateList = GenerateSerialListForType(SerialType.Nameplate);
            PrintManager.PrintUsingBPac(NameplatePrintSettings, nameplateList);
        }
        private async void GenerateReportButton_Click(object sender, EventArgs e) { await GenerateReport(); }
        private async void SubstrateListPrintButton_Click(object sender, EventArgs e) { await GenerateList(); }
        private async void CheckSheetPrintButton_Click(object sender, EventArgs e) { await GenerateCheckSheet(); }

    }
    public static class StringExtensions {
        public static string? NullIfWhiteSpace(this string? value) {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }
    }
}
