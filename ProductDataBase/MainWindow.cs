using Dapper;
using Microsoft.Data.Sqlite;
using ProductDatabase.Other;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        public int RadioButtonNumber { get; set; }
        private float _fontSize = SystemFonts.DefaultFont.Size;
        private IEnumerable<DataRow> _currentTargetRows = [];

        private string _dsn = string.Empty;
        private string _uid = string.Empty;
        private string _pwd = string.Empty;

        readonly string _jsonFilePath = Path.Combine(Environment.CurrentDirectory, "Config", "General", "appsettings.json");

        private readonly ProductRepository _productRepository;

        private readonly ProductMaster _productMaster;
        private readonly ProductRegisterWork _productRegisterWork;

        private readonly SubstrateMaster _substrateMaster;
        private readonly SubstrateRegisterWork _substrateRegisterWork;

        private readonly QrSettings _qrSettings;
        private readonly AppSettings _appSettings;

        public MainWindow() {
            InitializeComponent();

            _productRepository = new ProductRepository();
            _productMaster = new ProductMaster();
            _productRegisterWork = new ProductRegisterWork();
            _substrateMaster = new SubstrateMaster();
            _substrateRegisterWork = new SubstrateRegisterWork();
            _qrSettings = new QrSettings();
            _appSettings = new AppSettings();
        }

        private static FileStream? s_lockStream;

        // フォームロード時にファイルロック・設定読み込み・DBデータ取得・日次バックアップ作成を行いUIを初期化する
        private void LoadEvents() {
            try {
                // ファイルロック
                LockSelf();

                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;

                CategoryRadioButton1.Checked = false;
                CategoryRadioButton2.Checked = false;
                CategoryRadioButton3.Checked = false;
                CategoryRadioButton4.Checked = false;

                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                // 設定ファイル読み込み
                if (!File.Exists(_jsonFilePath)) {
                    throw new DirectoryNotFoundException($"'{_jsonFilePath}'\nが見つかりません。");
                }
                var generalSettings = System.Text.Json.JsonSerializer.Deserialize<GeneralSettings>(
                    File.ReadAllText(_jsonFilePath),
                    new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true }
                ) ?? new GeneralSettings();

                // バックアップパス取得
                CommonUtils.BackupPath = generalSettings.BackupFolderPath;

                // バックアップ作成
                CreateDailyBackup();

                // 担当者取得
                _appSettings.PersonList = [.. generalSettings.Persons];

                // 認証ユーザー名を取得
                _appSettings.IsAuthorizedUser = generalSettings.AuthorizedUsers
                    .Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase);

                // 管理者ユーザー名を取得
                _appSettings.IsAdministrator = generalSettings.Administrators
                    .Contains(Environment.UserName, StringComparer.OrdinalIgnoreCase);

                _dsn = generalSettings.DSN;
                _uid = generalSettings.UID;
                _pwd = generalSettings.PWD;

                _productRepository.LoadAll();

                this.Activate();
                QRCodePanel.Enabled = _appSettings.IsAuthorizedUser;
                QRCodeTextBox.Select();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 当日分のバックアップが未作成の場合のみDBをバックアップフォルダにコピーする
        private static void CreateDailyBackup() {
            // フォルダ未設定
            if (string.IsNullOrWhiteSpace(CommonUtils.BackupPath)) {
                MessageBox.Show("フォルダが設定されていません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ネットワークフォルダが見つからない
            if (!Directory.Exists(CommonUtils.BackupPath)) {
                MessageBox.Show($"'{CommonUtils.BackupPath}'\nが見つかりません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // バックアップ処理
            var today = DateTime.Today;
            var year = today.Year;
            var month = today.Month;
            var day = today.Day;

            // パス生成
            string backupFolder = Path.Combine(CommonUtils.BackupPath, "db", "backup", $"{year}", $"{month:00}");
            string backupFile = Path.Combine(backupFolder, $"_bak_{year}-{month:00}-{day:00}.db");
            string productRegistryFile = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");

            // 当日バックアップがなければ作成
            if (!File.Exists(backupFile)) {
                Directory.CreateDirectory(backupFolder);
                File.Copy(productRegistryFile, backupFile, overwrite: false);
            }
        }
        // 各マスター・作業データをリセットしDBを再読み込みする
        private void ResetFields() {
            _productMaster.Reset();
            _productRegisterWork.Reset();
            _productRepository.Clear();
            _substrateMaster.Reset();
            _substrateRegisterWork.Reset();

            _productRepository.LoadAll();

            _appSettings.FontSize = _fontSize;
        }
        // 実行中のEXEファイルをロックして二重起動を防止する
        private static void LockSelf() {
            try {
                string exePath = Application.ExecutablePath;
                s_lockStream = new FileStream(exePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            } catch {
                // 失敗しても無視
            }
        }
        // ラジオボタンのモードに応じて選択品目の基板登録・製品登録・再印刷・基板変更ウィンドウを開く
        private void Registration() {
            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<int> item) { return; }

            switch (RadioButtonNumber) {
                case 1:
                    HandleSubstrateRegistration(item.Id);
                    break;

                case 2:
                    HandleProductRegistration(item.Id, ProductOperationMode.Register);
                    break;

                case 3:
                    HandleProductRegistration(item.Id, ProductOperationMode.RePrint);
                    break;

                case 4:
                    HandleProductRegistration(item.Id, ProductOperationMode.SubstrateChange);
                    break;
            }

            QRCodeTextBox.Text = string.Empty;
        }
        // 指定基板IDのマスターを読み込み基板登録ウィンドウを開く
        private void HandleSubstrateRegistration(int substrateId) {

            var row = _productRepository.GetSubstrateById(substrateId);

            _substrateMaster.LoadFrom(row);

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // 指定製品IDのマスターを読み込みモードに応じた製品操作ウィンドウを開く
        private void HandleProductRegistration(int productId, ProductOperationMode mode) {

            var row = _productRepository.GetProductById(productId);

            _productMaster.LoadFrom(row);
            _productMaster.UseSubstrates = ProductRepository.GetUseSubstrates(_productMaster.ProductID);

            switch (mode) {
                case ProductOperationMode.Register:
                    using (var window = new ProductRegistration1Window(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;

                case ProductOperationMode.RePrint:
                    using (var window = new RePrintWindow(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;

                case ProductOperationMode.SubstrateChange:
                    using (var window = new SubstrateChange1(_productMaster, _productRegisterWork, _appSettings)) {
                        window.ShowDialog(this);
                    }
                    break;
            }
        }
        // 品目選択有無に応じてマスターをセットし履歴ウィンドウをダイアログで開く
        private void History() {

            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<int> item) {
                LoadHistoryWithoutSelection();
            }
            else {
                LoadHistoryWithSelection(item.Id);
            }

            using var window = new HistoryWindow(
                _productMaster,
                _productRegisterWork,
                _substrateMaster,
                _substrateRegisterWork,
                RadioButtonNumber,
                _appSettings);

            window.ShowDialog(this);
        }
        // 品目未選択時にカテゴリ名・製品名のみマスターにセットする
        private void LoadHistoryWithoutSelection() {
            switch (RadioButtonNumber) {
                case 1:
                    _substrateMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _substrateMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
                case 2 or 3 or 4:
                    _productMaster.CategoryName = CategoryListBox1.SelectedItem?.ToString() ?? string.Empty;
                    _productMaster.ProductName = CategoryListBox2.SelectedItem?.ToString() ?? string.Empty;
                    break;
            }
        }
        // 品目選択時に該当マスターデータをDBから読み込む
        private void LoadHistoryWithSelection(int itemId) {
            switch (RadioButtonNumber) {
                case 1:
                    _substrateMaster.LoadFrom(_productRepository.GetSubstrateById(itemId));
                    break;
                case 2 or 3 or 4:
                    _productMaster.LoadFrom(_productRepository.GetProductById(itemId));
                    break;
                default:
                    throw new InvalidOperationException("不正なモードです");
            }
        }

        private record CategoryConfig(string OrderKey, string IdKey, string NameKey);
        private readonly Dictionary<int, CategoryConfig> _categoryConfigs = new() {
            { 1, new CategoryConfig("SubstrateName", "SubstrateID", "SubstrateName") }, // 基板登録
            { 2, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 製品登録
            { 3, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 再印刷
            { 4, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }  // 基板変更
        };
        // ラジオボタン選択時にモードに応じたマスターデータをフィルタしてCategoryListBox1にカテゴリ一覧を表示する
        private void CategorySelect(object sender) {

            RegisterButton.Enabled = false;
            HistoryButton.Enabled = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            // RadioButton の Tag 取得
            if (sender is not RadioButton rb ||
                !int.TryParse(rb.Tag?.ToString(), out int number)) {
                return;
            }

            RadioButtonNumber = number;

            // 未定義モード防止
            if (!_categoryConfigs.ContainsKey(RadioButtonNumber)) {
                _currentTargetRows = [];
                return;
            }

            // データソース切替
            bool isSubstrateMode = RadioButtonNumber == 1;

            var sourceTable = isSubstrateMode
                ? _productRepository.SubstrateDataTable
                : _productRepository.ProductDataTable;

            // フィルタ済み行を保持
            _currentTargetRows = sourceTable
                .AsEnumerable()
                .Where(r => r["Visible"] != DBNull.Value &&
                            Convert.ToInt32(r["Visible"]) == 1)
                .Where(r => RadioButtonNumber switch {
                    1 => true,
                    2 => true,
                    3 => r["SerialPrintType"] != DBNull.Value &&
                         Convert.ToInt32(r["SerialPrintType"]) != 0,
                    4 => r["SheetPrintType"] != DBNull.Value &&
                        (Convert.ToInt32(r["SheetPrintType"]) == 2 ||
                         Convert.ToInt32(r["SheetPrintType"]) == 3),
                    _ => false
                });

            // CategoryName 列の重複除外＋SortNumber 昇順
            var categoryNames = _currentTargetRows
                .Where(r => r["CategoryName"] != DBNull.Value &&
                            r["SortNumber"] != DBNull.Value)
                .OrderBy(r => Convert.ToInt32(r["SortNumber"]))
                .Select(r => r["CategoryName"]!.ToString()!)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct()
                .ToList();

            CategoryListBox1.Items.AddRange([.. categoryNames]);
            HistoryButton.Enabled = RadioButtonNumber != 4;
        }
        // カテゴリ選択時に一致する製品名またはSubstrateNameの一覧をListBox2に表示する
        private void CategoryListBox1Select() {
            RegisterButton.Enabled = false;
            HistoryButton.Enabled = RadioButtonNumber != 4;
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            if (CategoryListBox1.SelectedItem is null) {
                return;
            }

            string selectedCategory = CategoryListBox1.SelectedItem.ToString()!;

            // _currentTargetRows から ProductName を取得
            var productNames = _currentTargetRows
                .Where(r =>
                    r["CategoryName"]?.ToString() == selectedCategory &&
                    r["ProductName"] != DBNull.Value)
                .Select(r => r["ProductName"]!.ToString()!)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            CategoryListBox2.Items.AddRange([.. productNames]);
        }
        // 製品名セレクト：カテゴリ・製品名に一致する品目一覧をListBox3に表示する
        private void CategoryListBox2Select() {
            RegisterButton.Enabled = false;
            HistoryButton.Enabled = RadioButtonNumber != 4;
            CategoryListBox3.Items.Clear();

            if (CategoryListBox1.SelectedItem is null ||
                CategoryListBox2.SelectedItem is null) {
                return;
            }

            if (!_categoryConfigs.TryGetValue(RadioButtonNumber, out var config)) {
                return;
            }

            var selectedRows = _currentTargetRows
                .Where(r =>
                    r["CategoryName"]?.ToString() == CategoryListBox1.SelectedItem?.ToString() &&
                    r["ProductName"]?.ToString() == CategoryListBox2.SelectedItem?.ToString())
                .OrderBy(r => r[config.OrderKey]?.ToString())
                .ToArray();

            var items = selectedRows
                .Select(r => new ListItem<int> {
                    Id = Convert.ToInt32(r[config.IdKey]),
                    Name = r[config.NameKey]?.ToString() ?? string.Empty
                })
                .GroupBy(x => x.Id)
                .Select(g => g.First())
                .OrderBy(x => x.Name)
                .ToList();

            CategoryListBox3.Items.AddRange([.. items.Cast<object>()]);
        }
        // 品目セレクト：選択確定時に登録ボタンを有効化する
        private void CategoryListBox3Select() {
            RegisterButton.Enabled = true;
        }

        // QR/バーコード入力を解析してSQLiteを検索し品目候補に応じた登録ウィンドウを開く
        private void CodeScan() {
            try {
                if (string.IsNullOrWhiteSpace(QRCodeTextBox.Text)) { return; }
                ResetFieldsForCodeScan();

                if (RadioButtonQR.Checked) {
                    if (textToUpperCheckBox.Checked) { QRCodeTextBox.Text = QRCodeTextBox.Text.ToUpper(); }
                    ParseQRCodeInput();
                }
                else if (RadioButtonBarcode.Checked) {
                    BarcodeInput();
                }

                ProcessCategoryItemData();
                FetchDataFromSQLite();

                var listIndex = 0;
                if (_qrSettings.CategoryItemNumber.Count >= 2) {
                    listIndex = ShowDialogWindowForMultipleItems();
                }

                if (listIndex == -1) { return; }
                HandleSelectedItem(listIndex);
                QRCodeTextBox.Text = string.Empty;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                CleanupAfterScan();
            }
        }
        // コードスキャン前にUIと各設定データをリセットしフォームを無効化する
        private void ResetFieldsForCodeScan() {
            CategoryRadioButton1.Checked = CategoryRadioButton2.Checked = CategoryRadioButton3.Checked = CategoryRadioButton4.Checked = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();
            _qrSettings.CategoryItemNumber.Clear();
            _qrSettings.CategoryProductType.Clear();
            _qrSettings.CategoryProductName.Clear();
            _qrSettings.CategorySubstrateName.Clear();
            _qrSettings.CategoryType.Clear();
            ResetFields();
            Enabled = false;
        }
        // QRコードテキストを解析して製番・品目番号・数量・注番を取得する
        private void ParseQRCodeInput() {
            try {
                string[] separator = ["//"];
                var arr = QRCodeTextBox.Text.Split(separator, StringSplitOptions.None);
                if (arr.Length == 1) {
                    _productMaster.ProductModel = QRCodeTextBox.Text;
                    return;
                }
                if (arr.Length != 4) { throw new Exception("QRコードが正しくありません。"); }
                _productRegisterWork.ProductNumber = arr[0];
                _productMaster.ProductModel = arr[1];
                _productRegisterWork.Quantity = int.TryParse(arr[2], out var result) ? result : throw new Exception("数量に数値以外が入力されています。");
                _productRegisterWork.OrderNumber = arr[3];
            } catch (Exception ex) {
                throw new Exception($"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}", ex);
            }
        }
        // バーコードの手配管理番号からODBCで手配情報を取得して各フィールドにセットする
        private void BarcodeInput() {
            using var con = new OdbcConnection($"DSN={_dsn}; UID={_uid}; PWD={_pwd}");
            con.Open();

            var result = con.QueryFirstOrDefault<OrderDto>(
                @"""
                SELECT *
                FROM V_宮崎手配情報
                WHERE 手配管理番号 = ?;
                """,
                new { 手配管理番号 = QRCodeTextBox.Text })
                ?? throw new Exception($"一致する情報がありません。{Environment.NewLine}手配管理番号:{QRCodeTextBox.Text}");

            _productRegisterWork.ProductNumber = result.手配製番 ?? string.Empty;
            _productMaster.ProductModel = result.品目番号 ?? string.Empty;
            _productMaster.ProductName = result.品目名称 ?? string.Empty;
            _productRegisterWork.Quantity = result.手配数;
            _productRegisterWork.OrderNumber = result.請求先注番 ?? string.Empty;
        }
        private sealed class OrderDto {
            public string 手配製番 { get; set; } = string.Empty;
            public string 品目番号 { get; set; } = string.Empty;
            public string 品目名称 { get; set; } = string.Empty;
            public int 手配数 { get; set; }
            public string 請求先注番 { get; set; } = string.Empty;
        }
        // 品目番号から不要なサフィックスを除去して正規化する
        private void ProcessCategoryItemData() {
            var pattern = @"-(?:SMT|H|GH).*";
            var result = Regex.Replace(_productMaster.ProductModel, pattern, string.Empty);
            _productMaster.ProductModel = result
                .Replace("-ACGH", "-AC")
                .Replace("-DCGH", "-DC");
        }
        // 品目番号でSQLiteを検索し一致する基板・製品情報をリストに追加する
        private void FetchDataFromSQLite() {
            using var con = new SqliteConnection(ProductRepository.GetConnectionRegistration());

            var sql =
                $"""
                SELECT
                    s.SubItemNumber,
                    s.SubstrateName,
                    s.ProductName AS sName,
                    p.ProductName AS pName,
                    p.ProductType,
                    p.ProItemNumber
                FROM
                    {Constants.ProductTableName} AS p
                FULL JOIN
                    {Constants.SubstrateTableName} AS s
                ON
                    s.SubItemNumber = p.ProItemNumber
                WHERE
                    s.SubItemNumber LIKE '%' || @ProductModel || '%'
                OR
                    p.ProItemNumber LIKE '%' || @ProductModel || '%'
                """;

            var results = con.Query(sql, new { _productMaster.ProductModel });

            if (!results.Any()) {
                throw new Exception($"品目番号が見つかりません。\n品目番号:[{_productMaster.ProductModel}]");
            }

            foreach (var row in results) {
                var colSubItemNumber = row.SubItemNumber?.ToString() ?? string.Empty;
                var colProItemNumber = row.ProItemNumber?.ToString() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(colSubItemNumber)) {
                    var productName = row.sName?.ToString() ?? string.Empty;
                    var substrateName = row.SubstrateName?.ToString() ?? string.Empty;
                    AddToLists(colSubItemNumber, productName, string.Empty, substrateName, "1");
                }

                if (!string.IsNullOrWhiteSpace(colProItemNumber)) {
                    var productName = row.pName?.ToString() ?? string.Empty;
                    var productType = row.ProductType?.ToString() ?? string.Empty;
                    AddToLists(colProItemNumber, productName, productType, string.Empty, "2");
                }
            }
        }
        // 品目情報を各カテゴリリストに追加する
        private void AddToLists(string itemNumber, string productName, string productType, string substrateName, string type) {
            _qrSettings.CategoryItemNumber.Add(itemNumber);
            _qrSettings.CategoryProductName.Add(productName);
            _qrSettings.CategoryProductType.Add(productType);
            _qrSettings.CategorySubstrateName.Add(substrateName);
            _qrSettings.CategoryType.Add(type);
        }
        // 複数候補がある場合に選択ダイアログを表示し選択インデックスを返す
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(_qrSettings, _appSettings);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
        // 選択された品目のタイプに応じて基板または製品の登録処理を実行する
        private void HandleSelectedItem(int listIndex) {
            var productName = _qrSettings.CategoryProductName[listIndex];
            var productType = _qrSettings.CategoryProductType[listIndex];
            var substrateName = _qrSettings.CategorySubstrateName[listIndex];
            var type = _qrSettings.CategoryType[listIndex];

            switch (type) {
                case "1":
                    HandleSubstrateSelection(productName, substrateName);
                    break;
                case "2":
                    HandleProductSelection(productName, productType);
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{_productMaster.ProductModel}{Environment.NewLine}");
            }
        }
        // 製品名と基板名で基板マスターを検索し基板登録ウィンドウを開く
        private void HandleSubstrateSelection(string productName, string substrateName) {
            var substrateRet = _productRepository.SubstrateDataTable
                .AsEnumerable()
                .Where(r => r["ProductName"]?.ToString() == productName &&
                            r["SubstrateName"]?.ToString() == substrateName)
                .ToArray();
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        // 基板マスターと作業データをセットして基板登録ウィンドウを表示する
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            _substrateMaster.LoadFrom(substrateRet[0]);
            _substrateRegisterWork.ProductNumber = _productRegisterWork.ProductNumber;
            _substrateRegisterWork.OrderNumber = _productRegisterWork.OrderNumber;
            _substrateRegisterWork.AddQuantity = _productRegisterWork.Quantity;
            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // 製品名と型式で製品マスターを検索し製品登録ウィンドウを開く
        private void HandleProductSelection(string productName, string productType) {
            var productRet = _productRepository.ProductDataTable
                .AsEnumerable()
                .Where(r => r["ProductName"]?.ToString() == productName &&
                            r["ProductType"]?.ToString() == productType)
                .ToArray();
            OpenProductRegistrationWindow(productRet);
        }
        // 製品マスターと使用基板情報をセットして製品登録ウィンドウを表示する
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            _productMaster.LoadFrom(productRet[0]);
            _productMaster.UseSubstrates = ProductRepository.GetUseSubstrates(_productMaster.ProductID);
            using ProductRegistration1Window window = new(_productMaster, _productRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        // スキャン後にフォームを再有効化してQRコード入力欄にフォーカスを戻す
        private void CleanupAfterScan() {
            Enabled = true;
            QRCodeTextBox.Focus();
        }

        // ラジオボタン選択に応じてアプリ全体のフォントサイズを変更する
        private void FontChange(object sender) {
            var radioButton = (RadioButton)sender;

            _fontSize = radioButton.Name switch {
                "FontSize9RadioButton" => 9,
                "FontSize12RadioButton" => 12,
                "FontSize14RadioButton" => 14,
                _ => _fontSize
            };

            _appSettings.FontSize = _fontSize;
            Font = new System.Drawing.Font(_appSettings.FontName, _appSettings.FontSize);
        }
        // ExcelのパスをInteropで取得しEXEを直接起動して指定ファイルを開く
        private static void OpenExcel(string filePath) {
            Microsoft.Office.Interop.Excel.Application? xlApp = null;
            try {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                var excelFullPath = Path.Combine(xlApp.Path, "EXCEL.EXE");
                xlApp.Quit();

                using var process = Process.Start(new ProcessStartInfo {
                    FileName = excelFullPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = true
                });
            } finally {
                if (xlApp != null) {
                    Marshal.ReleaseComObject(xlApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        private void MainWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void ReloadToolStripMenuItem_Click(object sender, EventArgs e) { LoadEvents(); }
        private void ConfigReportToolStripMenuItem_Click(object sender, EventArgs e) {
            var reportConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsm");
            OpenExcel(reportConfigPath);
        }
        private void ConfigListToolStripMenuItem_Click(object sender, EventArgs e) {
            var listConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsm");
            OpenExcel(listConfigPath);
        }
        private void ConfigCheckSheetToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsm");
            OpenExcel(checkSheetConfigPath);
        }
        private void ConfigSubstrateInformationToolStripMenuItem_Click(object sender, EventArgs e) {
            var checkSheetConfigPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsm");
            OpenExcel(checkSheetConfigPath);
        }
        private void 終了ToolStripMenuItem_Click(object sender, EventArgs e) { Close(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
        private void HistoryButton_Click(object sender, EventArgs e) { History(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void CategoryRadioButton_CheckedChanged(object sender, EventArgs e) { CategorySelect(sender); }
        private void FontSize_CheckedChanged(object sender, EventArgs e) { FontChange(sender); }
        private void CategoryListBox3_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) { return; }
            Registration();
        }
        private void QRCodeTextBox_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter) { return; }
            CodeScan();
        }
        private void QRCodeButton_Click(object sender, EventArgs e) { CodeScan(); }
        private void QRCodeTextBox_Enter(object sender, EventArgs e) { CommonUtils.Keyboard.CapsDisable(); }

    }
}
