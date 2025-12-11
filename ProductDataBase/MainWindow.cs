using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using ProductDatabase.Other;
using System.Data;
using System.Data.Odbc;
using System.Text.RegularExpressions;

namespace ProductDatabase {
    public partial class MainWindow : Form {

        public int RadioButtonNumber { get; set; }
        private float _fontSize = SystemFonts.DefaultFont.Size;
        private IEnumerable<DataRow> _currentTargetRows = [];

        readonly string _jsonFilePath = Path.Combine(Environment.CurrentDirectory, "Config", "General", "appsettings.json");

        [Flags]
        public enum SerialPrintTypeFlags {
            None = 0,
            Label = 1 << 0,         // 00001
            Barcode = 1 << 1,       // 00010
            Nameplate = 1 << 2,     // 00100
            Underline = 1 << 3,     // 01000
            Last4Digits = 1 << 4    // 10000
        }

        [Flags]
        public enum SheetPrintTypeFlags {
            None = 0,
            CheckSheet = 1 << 0,    // 00001
            List = 1 << 1,          // 00010
        }

        public class ProductRepository {

            public DataTable ProductDataTable { get; } = new();
            public DataTable SubstrateDataTable { get; } = new();
            public DataTable ProductUseSubstrate { get; } = new();

            public void LoadAll() {

                ProductDataTable.Clear();
                SubstrateDataTable.Clear();
                ProductUseSubstrate.Clear();

                using var con = new SqliteConnection(GetConnectionRegistration());
                con.Open();

                using (var cmd = new SqliteCommand("SELECT * FROM M_ProductDef;", con))
                using (var reader = cmd.ExecuteReader()) {
                    ProductDataTable.Load(reader);
                }

                using (var cmd = new SqliteCommand("SELECT * FROM M_SubstrateDef;", con))
                using (var reader = cmd.ExecuteReader()) {
                    SubstrateDataTable.Load(reader);
                }

                using (var cmd = new SqliteCommand("SELECT * FROM V_ProductUseSubstrate;", con))
                using (var reader = cmd.ExecuteReader()) {
                    ProductUseSubstrate.Load(reader);
                }
            }

            public DataRow GetSubstrateById(int id) {
                return SubstrateDataTable
                    .AsEnumerable()
                    .First(r => Convert.ToInt32(r["SubstrateID"]) == id);
            }

            public DataRow GetProductById(int id) {
                return ProductDataTable
                    .AsEnumerable()
                    .First(r => Convert.ToInt32(r["ProductID"]) == id);
            }

            // 製品IDから使用基板リストを取得
            public List<SubstrateInfo> GetUseSubstrates(int productKey) {

                return [.. ProductUseSubstrate.AsEnumerable()
                    .Where(r => Convert.ToInt32(r["PKey"]) == productKey)
                    .Select(r => new SubstrateInfo {
                        SubstrateID = Convert.ToInt32(r["SKey"]),
                        SubstrateName = r["SubstrateName"]?.ToString() ?? string.Empty,
                        SubstrateModel = r["SubstrateModel"]?.ToString() ?? string.Empty
                    })];
            }

            public void Clear() {
                ProductDataTable.Clear();
                SubstrateDataTable.Clear();
                ProductUseSubstrate.Clear();
            }
        }

        public class ProductMaster {
            public int ProductID { get; set; }
            public string CategoryName { get; set; } = string.Empty;
            public string ProductName { get; set; } = string.Empty;
            public string ProductModel { get; set; } = string.Empty;
            public string ProductType { get; set; } = string.Empty;
            public string Initial { get; set; } = string.Empty;
            public int SerialDigitType { get; set; }
            public int SerialDigit => SerialDigitType switch {
                3 or 101 or 102 => 3,
                4 => 4,
                _ => 0
            };
            public int RevisionGroup { get; set; }
            public int CheckBin { get; set; }
            public List<SubstrateInfo> UseSubstrates { get; set; } = [];

            private int _regType;
            public int RegType {
                get => _regType;
                set {
                    _regType = value;
                    UpdatePrintFlags();
                }
            }

            private int _serialPrintType;
            public int SerialPrintType {
                get => _serialPrintType;
                set {
                    _serialPrintType = value;
                    UpdatePrintFlags();
                }
            }

            private int _sheetPrintType;
            public int SheetPrintType {
                get => _sheetPrintType;
                set {
                    _sheetPrintType = value;
                    UpdatePrintFlags();
                }
            }

            // ===== 結果フラグ =====

            public bool IsLabelPrint { get; private set; }
            public bool IsBarcodePrint { get; private set; }
            public bool IsNameplatePrint { get; private set; }
            public bool IsLast4Digits { get; private set; }
            public bool IsUnderlinePrint { get; private set; }

            public bool IsCheckSheetPrint { get; private set; }
            public bool IsListPrint { get; private set; }

            public bool IsSerialGeneration { get; private set; }
            public bool IsRegType9 { get; private set; }

            // ===== 内部更新処理 =====

            private void UpdatePrintFlags() {
                UpdateRegTypeFlags();
                UpdateSerialPrintTypeFlags();
                UpdateSheetPrintTypeFlags();
            }

            private void UpdateRegTypeFlags() {
                IsRegType9 = RegType == 9;
                IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
            }

            private void UpdateSerialPrintTypeFlags() {
                var flags = (SerialPrintTypeFlags)_serialPrintType;

                IsLabelPrint = flags.HasFlag(SerialPrintTypeFlags.Label);
                IsBarcodePrint = flags.HasFlag(SerialPrintTypeFlags.Barcode);
                IsNameplatePrint = flags.HasFlag(SerialPrintTypeFlags.Nameplate);
                IsUnderlinePrint = flags.HasFlag(SerialPrintTypeFlags.Underline);
                IsLast4Digits = flags.HasFlag(SerialPrintTypeFlags.Last4Digits);
            }

            private void UpdateSheetPrintTypeFlags() {
                var flags = (SheetPrintTypeFlags)_sheetPrintType;

                IsCheckSheetPrint = flags.HasFlag(SheetPrintTypeFlags.CheckSheet);
                IsListPrint = flags.HasFlag(SheetPrintTypeFlags.List);
            }

            public void LoadFrom(DataRow row) {
                ProductID = Convert.ToInt32(row["ProductID"]);
                CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
                ProductName = row.Field<string>("ProductName") ?? string.Empty;
                ProductType = row.Field<string>("ProductType") ?? string.Empty;
                ProductModel = row.Field<string>("ProductModel") ?? string.Empty;
                Initial = row.Field<string>("Initial") ?? string.Empty;
                RevisionGroup = Convert.ToInt32(row["RevisionGroup"]);
                RegType = Convert.ToInt32(row["RegType"]);
                CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2);
                SerialDigitType = Convert.ToInt32(row["SerialType"]);
                SerialPrintType = Convert.ToInt32(row["SerialPrintType"]);
                SheetPrintType = Convert.ToInt32(row["SheetPrintType"]);
            }

            public void Reset() {
                ProductID = 0;
                CategoryName = string.Empty;
                ProductName = string.Empty;
                ProductModel = string.Empty;
                ProductType = string.Empty;
                Initial = string.Empty;
                SerialDigitType = 0;
                RevisionGroup = 0;
                CheckBin = 0;
                UseSubstrates = [];
                RegType = 0;
                SerialPrintType = 0;
                SheetPrintType = 0;
            }
        }
        public class SubstrateInfo {
            public int SubstrateID { get; set; }
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
        }
        public class ProductRegisterWork {
            public int RowID { get; set; }
            public int ProductID { get; set; }
            public string ProductNumber { get; set; } = string.Empty;
            public string OrderNumber { get; set; } = string.Empty;
            public string SerialFirst { get; set; } = string.Empty;
            public string SerialLast { get; set; } = string.Empty;
            public int Quantity { get; set; }
            public string RegDate { get; set; } = string.Empty;
            public string Person { get; set; } = string.Empty;
            public string Revision { get; set; } = string.Empty;
            public string Comment { get; set; } = string.Empty;

            public int SerialFirstNumber { get; set; }
            public int SerialLastNumber { get; set; }

            public void Reset() {
                RowID = 0;
                ProductID = 0;
                ProductNumber = string.Empty;
                OrderNumber = string.Empty;
                SerialFirst = string.Empty;
                SerialLast = string.Empty;
                Quantity = 0;
                RegDate = string.Empty;
                Person = string.Empty;
                Revision = string.Empty;
                Comment = string.Empty;
                SerialFirstNumber = 0;
                SerialLastNumber = 0;
            }
        }

        public class SubstrateMaster {
            public int SubstrateID { get; set; }
            public string CategoryName { get; set; } = string.Empty;
            public string ProductName { get; set; } = string.Empty;
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public int CheckBin { get; set; }

            private int _regType;
            public int RegType {
                get => _regType;
                set {
                    _regType = value;
                    UpdatePrintFlags();
                }
            }

            private int _serialPrintType;
            public int SerialPrintType {
                get => _serialPrintType;
                set {
                    _serialPrintType = value;
                    UpdatePrintFlags();
                }
            }

            // ===== 結果フラグ =====

            public bool IsLabelPrint { get; private set; }

            public bool IsSerialGeneration { get; private set; }

            // ===== 内部更新処理 =====

            private void UpdatePrintFlags() {
                UpdateRegTypeFlags();
                UpdateSerialPrintTypeFlags();
            }

            private void UpdateRegTypeFlags() {
                IsSerialGeneration = RegType is 1 or 2 or 3 or 9;
            }

            private void UpdateSerialPrintTypeFlags() {
                var flags = (SerialPrintTypeFlags)_serialPrintType;

                IsLabelPrint = flags.HasFlag(SerialPrintTypeFlags.Label);
            }

            public void LoadFrom(DataRow row) {
                SubstrateID = Convert.ToInt32(row["SubstrateID"]);
                CategoryName = row.Field<string>("CategoryName") ?? string.Empty;
                ProductName = row.Field<string>("ProductName") ?? string.Empty;
                SubstrateName = row.Field<string>("SubstrateName") ?? string.Empty;
                SubstrateModel = row.Field<string>("SubstrateModel") ?? string.Empty;
                RegType = Convert.ToInt32(row["RegType"]);
                CheckBin = Convert.ToInt32(row["Checkbox"].ToString(), 2);
                SerialPrintType = Convert.ToInt32(row["SerialPrintType"]);
            }

            public void Reset() {
                SubstrateID = 0;
                CategoryName = string.Empty;
                ProductName = string.Empty;
                SubstrateName = string.Empty;
                SubstrateModel = string.Empty;
                RegType = 0;
                CheckBin = 0;
                RegType = 0;
                SerialPrintType = 0;
            }
        }
        public class SubstrateRegisterWork {
            public int SubstrateID { get; set; }

            public string ProductNumber { get; set; } = string.Empty;
            public string OrderNumber { get; set; } = string.Empty;
            public int AddQuantity { get; set; }
            public int DefectQuantity { get; set; }
            public int UseQuantity { get; set; }

            public string Person { get; set; } = string.Empty;
            public string RegDate { get; set; } = string.Empty;
            public string Comment { get; set; } = string.Empty;

            public void Reset() {
                SubstrateID = 0;
                ProductNumber = string.Empty;
                OrderNumber = string.Empty;
                AddQuantity = 0;
                DefectQuantity = 0;
                UseQuantity = 0;
                Person = string.Empty;
                RegDate = string.Empty;
                Comment = string.Empty;
            }
        }

        public class QrSettings {
            public List<string> CategoryItemNumber { get; set; } = [];
            public List<string> CategoryProductName { get; set; } = [];
            public List<string> CategorySubstrateName { get; set; } = [];
            public List<string> CategoryProductType { get; set; } = [];
            public List<string> CategoryType { get; set; } = [];
        }

        public class AppSettings {
            public List<string> PersonList { get; set; } = [];
            public string FontName { get; set; } = "Meiryo UI";
            public float FontSize { get; set; } = 9;
        }

        public class Config {
            public string BackupFolderPath { get; set; } = string.Empty;
            public string[] Administrators { get; set; } = [];
            public string[] AuthorizedUsers { get; set; } = [];
            public static bool IsAdministrator { get; set; } = false;
            public static bool IsAuthorizedUser { get; set; } = false;
        }

        public class ListItem<T> {
            public T Id { get; set; } = default!;
            public string Name { get; set; } = string.Empty;

            public override string ToString() => Name;
        }

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

        public static string GetConnectionRegistration() {
            var productRegistryPath = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");
            return !File.Exists(productRegistryPath)
                ? throw new FileNotFoundException("ファイルが見つかりません。", productRegistryPath)
                : new SqliteConnectionStringBuilder() { DataSource = productRegistryPath, Pooling = false }.ToString();
        }

        private static FileStream? s_lockStream;

        // ロードイベント
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
                var jsonText = File.ReadAllText(_jsonFilePath);
                var jsonObj = JObject.Parse(jsonText);

                // CloneFolderPathを取得
                CommonUtils.s_backupPath = jsonObj["BackupFolderPath"]?.ToString() ?? string.Empty;

                // バックアップ作成
                CreateDailyBackup();

                // 担当者取得
                _appSettings.PersonList.Clear();
                foreach (var person in jsonObj["Persons"]!) {
                    _appSettings.PersonList.Add(person.ToString()!);
                }

                // 認証ユーザー名を取得
                var userSet = (jsonObj["AuthorizedUsers"] as JArray)?
                    .Select(u => (string)u!)
                    .Where(u => u != null)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                Config.IsAuthorizedUser = userSet?.Contains(Environment.UserName) ?? false;

                // 管理者ユーザー名を取得
                var adminSet = (jsonObj["Administrators"] as JArray)?
                    .Select(u => (string)u!)
                    .Where(u => u != null)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                Config.IsAdministrator = adminSet?.Contains(Environment.UserName) ?? false;

                _productRepository.LoadAll();

                this.Activate();
                QRCodePanel.Enabled = Config.IsAuthorizedUser;
                QRCodeTextBox.Select();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static void CreateDailyBackup() {
            // フォルダ未設定
            if (string.IsNullOrWhiteSpace(CommonUtils.s_backupPath)) {
                MessageBox.Show("フォルダが設定されていません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ネットワークフォルダが見つからない
            if (!Directory.Exists(CommonUtils.s_backupPath)) {
                MessageBox.Show($"'{CommonUtils.s_backupPath}'\nが見つかりません。バックアップは保存されません。",
                    string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // バックアップ処理
            var today = DateTime.Today;
            var year = today.Year;
            var month = today.Month;
            var day = today.Day;

            // パス生成
            string backupFolder = Path.Combine(CommonUtils.s_backupPath, "db", "backup", $"{year}", $"{month:00}");
            string backupFile = Path.Combine(backupFolder, $"_bak_{year}-{month:00}-{day:00}.db");
            string productRegistryFile = Path.Combine(Environment.CurrentDirectory, "db", "ProductRegistry.db");

            // 当日バックアップがなければ作成
            if (!File.Exists(backupFile)) {
                Directory.CreateDirectory(backupFolder);
                File.Copy(productRegistryFile, backupFile, overwrite: false);
            }
        }
        private void ResetFields() {
            _productMaster.Reset();
            _productRegisterWork.Reset();
            _productRepository.Clear();
            _substrateMaster.Reset();
            _substrateRegisterWork.Reset();

            _productRepository.LoadAll();

            _appSettings.FontSize = _fontSize;
        }
        private static void LockSelf() {
            try {
                string exePath = Application.ExecutablePath;
                s_lockStream = new FileStream(exePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            } catch {
                // 失敗しても無視
            }
        }
        // 登録ボタン処理
        public enum ProductOperationMode {
            Register,
            RePrint,
            SubstrateChange
        }
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
        private void HandleSubstrateRegistration(int substrateId) {

            var row = _productRepository.GetSubstrateById(substrateId);

            _substrateMaster.LoadFrom(row);

            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        private void HandleProductRegistration(int productId, ProductOperationMode mode) {

            var row = _productRepository.GetProductById(productId);

            _productMaster.LoadFrom(row);
            _productMaster.UseSubstrates = _productRepository.GetUseSubstrates(_productMaster.ProductID);

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
        // 履歴ボタン処理
        private void History() {

            ResetFields();

            if (CategoryListBox3.SelectedItem is not ListItem<int> item) {
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
            else {
                DataRow row = RadioButtonNumber switch {
                    1 => _productRepository.GetSubstrateById(item.Id),
                    2 or 3 or 4 => _productRepository.GetProductById(item.Id),
                    _ => throw new InvalidOperationException("不正なモードです")
                };

                if (RadioButtonNumber == 1) {
                    _substrateMaster.LoadFrom(row);
                }
                else {
                    _productMaster.LoadFrom(row);
                }
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

        private record CategoryConfig(string OrderKey, string IdKey, string NameKey);
        private readonly Dictionary<int, CategoryConfig> _categoryConfigs = new() {
            { 1, new CategoryConfig("SubstrateName", "SubstrateID", "SubstrateName") }, // 基板登録
            { 2, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 製品登録
            { 3, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }, // 再印刷
            { 4, new CategoryConfig("ProductType",   "ProductID",   "ProductType")   }  // 基板変更
        };
        // 処理カテゴリセレクト
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
        }
        // 製品カテゴリセレクト
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
        private void CategoryListBox3Select() {
            RegisterButton.Enabled = true;
        }

        // QRコード読み取り
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
        private void ResetFieldsForCodeScan() {
            CategoryRadioButton1.Checked = CategoryRadioButton2.Checked = CategoryRadioButton3.Checked = CategoryRadioButton4.Checked = false;
            CategoryListBox1.Items.Clear();
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();
            _qrSettings.CategoryItemNumber.Clear();
            _qrSettings.CategoryProductType.Clear();
            _qrSettings.CategoryProductName.Clear();
            _qrSettings.CategoryType.Clear();
            ResetFields();
            Enabled = false;
        }
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
                throw new Exception($"{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー{Environment.NewLine}{ex.Message}");
            }
        }
        private void BarcodeInput() {
            using (OdbcConnection con = new("DSN=DrSum_PRONES_YD; UID=YD00; PWD=YD00")) {
                con.Open();
                using OdbcCommand cmd = new($"SELECT * FROM V_宮崎手配情報 WHERE 手配管理番号 = '{QRCodeTextBox.Text}';", con);
                using var dr = cmd.ExecuteReader();
                while (dr.Read()) {
                    _productRegisterWork.ProductNumber = dr["手配製番"].ToString() ?? string.Empty;
                    _productMaster.ProductModel = dr["品目番号"].ToString() ?? string.Empty;
                    _productMaster.ProductName = dr["品目名称"].ToString() ?? string.Empty;
                    _productRegisterWork.Quantity = Convert.ToInt32(dr["手配数"] ?? throw new Exception("手配数 is null"));
                    _productRegisterWork.OrderNumber = dr["請求先注番"].ToString() ?? string.Empty;
                }
            }
            if (string.IsNullOrEmpty(_productRegisterWork.ProductNumber)) { throw new Exception($"一致する情報がありません。{Environment.NewLine}手配管理番号:{QRCodeTextBox.Text}"); }
        }
        private void ProcessCategoryItemData() {
            var pattern = @"-(?:SMT|H|GH).*";
            var result = Regex.Replace(_productMaster.ProductModel, pattern, string.Empty);
            _productMaster.ProductModel = result
                .Replace("-ACGH", "-AC")
                .Replace("-DCGH", "-DC");
        }
        private void FetchDataFromSQLite() {
            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText =
                $"""
                SELECT
                    s.SubItemNumber,
                    s.SubstrateName,
                    s.ProductName AS sName,
                    p.ProductName AS pName,
                    p.ProductType,
                    p.ProItemNumber
                FROM
                    M_SubstrateDef AS s
                FULL JOIN
                    M_ProductDef AS p
                ON
                    s.SubItemNumber = p.ProItemNumber
                WHERE
                    s.SubItemNumber LIKE '%'|| @StrProness2 ||'%'
                OR
                    p.ProItemNumber LIKE '%'|| @StrProness2 ||'%'
                ;
                """;

            cmd.Parameters.Add("@StrProness2", SqliteType.Text).Value = _productMaster.ProductModel;
            using var dr = cmd.ExecuteReader();
            if (!dr.HasRows) { throw new Exception($"品目番号が見つかりません。\n品目番号:[{_productMaster.ProductModel}]"); }
            while (dr.Read()) {
                var colSubItemNumber = dr["SubItemNumber"].ToString() ?? string.Empty;
                var colProItemNumber = dr["ProItemNumber"].ToString() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(colSubItemNumber)) {
                    var productName = dr["sName"]?.ToString() ?? string.Empty;
                    var substrateName = dr["SubstrateName"]?.ToString() ?? string.Empty;
                    AddToLists(colSubItemNumber, productName, string.Empty, substrateName, "1");
                }

                if (!string.IsNullOrWhiteSpace(colProItemNumber)) {
                    var productName = dr["pName"]?.ToString() ?? string.Empty;
                    var productType = dr["ProductType"]?.ToString() ?? string.Empty;
                    AddToLists(colProItemNumber, productName, productType, string.Empty, "2");
                }
            }
        }
        private void AddToLists(string itemNumber, string productName, string productType, string substrateName, string type) {
            _qrSettings.CategoryItemNumber.Add(itemNumber);
            _qrSettings.CategoryProductName.Add(productName);
            _qrSettings.CategoryProductType.Add(productType);
            _qrSettings.CategorySubstrateName.Add(substrateName);
            _qrSettings.CategoryType.Add(type);
        }
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(_qrSettings, _appSettings);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
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
        private void HandleSubstrateSelection(string productName, string substrateName) {
            var substrateRet = _productRepository.SubstrateDataTable.Select($"ProductName = '{productName}' AND SubstrateName = '{substrateName}'");
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            _substrateMaster.CategoryName = substrateRet[0]["CategoryName"].ToString() ?? string.Empty;
            _substrateMaster.ProductName = substrateRet[0]["ProductName"].ToString() ?? string.Empty;
            _substrateMaster.SubstrateName = substrateRet[0]["SubstrateName"].ToString() ?? string.Empty;
            _substrateMaster.SubstrateModel = substrateRet[0]["SubstrateModel"].ToString() ?? string.Empty;
            _substrateMaster.RegType = Convert.ToInt32(substrateRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            _substrateMaster.CheckBin = Convert.ToInt32(substrateRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            _substrateMaster.SerialPrintType = Convert.ToInt32(substrateRet[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
            _substrateRegisterWork.ProductNumber = _productRegisterWork.ProductNumber;
            _substrateRegisterWork.OrderNumber = _productRegisterWork.OrderNumber;
            _substrateRegisterWork.AddQuantity = _productRegisterWork.Quantity;
            using SubstrateRegistrationWindow window = new(_substrateMaster, _substrateRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        private void HandleProductSelection(string productName, string productType) {
            var productRet = _productRepository.ProductDataTable.Select($"ProductName = '{productName}' AND ProductType = '{productType}'");
            OpenProductRegistrationWindow(productRet);
        }
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            _productMaster.CategoryName = productRet[0]["CategoryName"].ToString() ?? string.Empty;
            _productMaster.ProductName = productRet[0]["ProductName"].ToString() ?? string.Empty;
            _productMaster.ProductType = productRet[0]["ProductType"].ToString() ?? string.Empty;
            _productMaster.ProductModel = productRet[0]["ProductModel"].ToString() ?? string.Empty;
            _productMaster.Initial = productRet[0]["Initial"].ToString() ?? string.Empty;
            _productMaster.RevisionGroup = Convert.ToInt32(productRet[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
            _productMaster.RegType = Convert.ToInt32(productRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            _productMaster.CheckBin = Convert.ToInt32(productRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            _productMaster.SerialDigitType = Convert.ToInt32(productRet[0]["SerialType"] ?? throw new Exception("SerialType is null"));
            _productMaster.SerialPrintType = Convert.ToInt32(productRet[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
            _productMaster.SheetPrintType = Convert.ToInt32(productRet[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
            _productMaster.UseSubstrates = _productRepository.GetUseSubstrates(_productMaster.ProductID);
            using ProductRegistration1Window window = new(_productMaster, _productRegisterWork, _appSettings);
            window.ShowDialog(this);
        }
        private void CleanupAfterScan() {
            Enabled = true;
            QRCodeTextBox.Focus();
        }

        // フォント変更
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
        // excel
        private static void OpenExcel(string filePath) {
            // Excel実行ファイルの場所を取得
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var excelFullPath = System.IO.Path.Combine(xlApp.Path, "EXCEL.EXE");
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            // ファイルを実行
            System.Diagnostics.Process.Start(excelFullPath, filePath);
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
