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
        private DataTable MainDataTable { get; } = new();
        private float _fontSize = SystemFonts.DefaultFont.Size;

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

        public class ProductInformation {
            public int ID { get; set; }
            public long ProductID { get; set; } = 0;
            public long SubstrateID { get; set; } = 0;
            public string CategoryName { get; set; } = string.Empty;
            public string ProductName { get; set; } = string.Empty;
            public string StockName { get; set; } = string.Empty;
            public string ProductType { get; set; } = string.Empty;
            public string ProductModel { get; set; } = string.Empty;
            public string SubstrateName { get; set; } = string.Empty;
            public string SubstrateModel { get; set; } = string.Empty;
            public string Initial { get; set; } = string.Empty;
            public List<string> UseSubstrate { get; set; } = [];
            public int SerialType { get; set; }
            public int RevisionGroup { get; set; }
            public int CheckBin { get; set; }
            public string Proness1 { get; set; } = string.Empty;
            public string Proness2 { get; set; } = string.Empty;
            public string Proness3 { get; set; } = string.Empty;
            public int Proness4 { get; set; }
            public string Proness5 { get; set; } = string.Empty;
            public List<string> CategoryItemNumber { get; set; } = [];
            public List<string> CategoryProductName { get; set; } = [];
            public List<string> CategorySubstrateName { get; set; } = [];
            public List<string> CategoryProductType { get; set; } = [];
            public List<string> CategoryType { get; set; } = [];
            public List<string> PersonList { get; set; } = [];
            public string FontName { get; } = "Meiryo UI";
            public float FontSize { get; set; } = 9;

            public string OrderNumber { get; set; } = string.Empty;
            public string ProductNumber { get; set; } = string.Empty;
            public string RegDate { get; set; } = string.Empty;
            public string Person { get; set; } = string.Empty;
            public string Revision { get; set; } = string.Empty;
            public string Comment { get; set; } = string.Empty;
            public string SerialFirst { get; set; } = string.Empty;
            public string SerialLast { get; set; } = string.Empty;

            private int _regType;
            private int _serialPrintType;
            private int _sheetPrintType;

            public int RegType {
                get => _regType;
                set {
                    _regType = value;
                    UpdatePrintFlags();
                }
            }

            public int SerialPrintType {
                get => _serialPrintType;
                set {
                    _serialPrintType = value;
                    UpdatePrintFlags();
                }
            }

            public int SheetPrintType {
                get => _sheetPrintType;
                set {
                    _sheetPrintType = value;
                    UpdatePrintFlags();
                }
            }

            public bool IsLabelPrint { get; private set; }
            public bool IsBarcodePrint { get; private set; }
            public bool IsNameplatePrint { get; private set; }
            public bool IsLast4Digits { get; private set; }
            public bool IsUnderlinePrint { get; private set; }

            public bool IsCheckSheetPrint { get; private set; }
            public bool IsListPrint { get; private set; }

            public bool IsSerialGeneration { get; private set; }
            public bool IsRegType9 { get; private set; }

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

            public int Quantity { get; set; }
            public int SerialFirstNumber { get; set; }
            public int SerialLastNumber { get; set; }

            public DataTable ProductDataTable { get; } = new();
            public DataTable SubstrateDataTable { get; } = new();
            public DataTable ProductUseSubstrate { get; } = new();

            public int SerialDigit => SerialType switch {
                3 or 101 or 102 => 3,
                4 => 4,
                _ => 0
            };
        }
        public class Config {
            public string BackupFolderPath { get; set; } = string.Empty;
            public string[] Administrators { get; set; } = [];
            public string[] AuthorizedUsers { get; set; } = [];
            public static bool IsAdministrator { get; set; } = false;
            public static bool IsAuthorizedUser { get; set; } = false;
        }

        public ProductInformation ProductInfo { get; set; } = new();

        public MainWindow() {
            InitializeComponent();
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
                MainDataTable.Clear();

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

                // 製品・基板データ取得
                using SqliteConnection con = new(GetConnectionRegistration());
                con.Open();
                using (var cmd = new SqliteCommand("SELECT * FROM M_ProductDef;", con)) {
                    using var reader = cmd.ExecuteReader();
                    ProductInfo.ProductDataTable.Load(reader);
                }
                using (var cmd = new SqliteCommand("SELECT * FROM M_SubstrateDef;", con)) {
                    using var reader = cmd.ExecuteReader();
                    ProductInfo.SubstrateDataTable.Load(reader);
                }
                using (var cmd = new SqliteCommand("SELECT * FROM T_ProductUseSubstrate;", con)) {
                    using var reader = cmd.ExecuteReader();
                    ProductInfo.ProductUseSubstrate.Load(reader);
                }

                // 担当者取得
                ProductInfo.PersonList.Clear();
                foreach (var person in jsonObj["Persons"]!) {
                    ProductInfo.PersonList.Add(person.ToString()!);
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
                    "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ネットワークフォルダが見つからない
            if (!Directory.Exists(CommonUtils.s_backupPath)) {
                MessageBox.Show($"'{CommonUtils.s_backupPath}'\nが見つかりません。バックアップは保存されません。",
                    "", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            ProductInfo = new ProductInformation();

            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();
            using (var cmd = new SqliteCommand("SELECT * FROM M_ProductDef;", con)) {
                using var reader = cmd.ExecuteReader();
                ProductInfo.ProductDataTable.Load(reader);
            }
            using (var cmd = new SqliteCommand("SELECT * FROM M_SubstrateDef;", con)) {
                using var reader = cmd.ExecuteReader();
                ProductInfo.SubstrateDataTable.Load(reader);
            }
            using (var cmd = new SqliteCommand("SELECT * FROM T_ProductUseSubstrate;", con)) {
                using var reader = cmd.ExecuteReader();
                ProductInfo.ProductUseSubstrate.Load(reader);
            }

            // 担当者取得
            ProductInfo.PersonList.Clear();
            var jsonText = File.ReadAllText(_jsonFilePath);
            var jsonObj = JObject.Parse(jsonText);
            foreach (var person in jsonObj["Persons"]!) {
                ProductInfo.PersonList.Add(person.ToString()!);
            }

            ProductInfo.FontSize = _fontSize;
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
        private void Registration() {
            try {
                ResetFields();

                if (CategoryListBox3.SelectedIndex == -1) { return; }
                switch (RadioButtonNumber) {
                    case 1:
                        HandleSubstrateRegistration();
                        break;
                    case 2:
                        HandleProductRegistration1();
                        break;
                    case 3:
                        HandleRePrint();
                        break;
                    case 4:
                        HandleSubstrateChange1();
                        break;
                    default:
                        break;
                }

                QRCodeTextBox.Text = string.Empty;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void HandleSubstrateRegistration() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND SubstrateName = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.SubstrateID = Convert.ToInt64(selectedRows[0]["SubstrateID"]);
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.SubstrateName = selectedRows[0]["SubstrateName"].ToString() ?? string.Empty;
                ProductInfo.SubstrateModel = selectedRows[0]["SubstrateModel"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                using SubstrateRegistrationWindow window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleProductRegistration1() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.ProductID = Convert.ToInt64(selectedRows[0]["ProductID"]);
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                ProductInfo.SheetPrintType = Convert.ToInt32(selectedRows[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
                ProductInfo.SerialType = Convert.ToInt32(selectedRows[0]["SerialType"] ?? throw new Exception("SerialType is null"));
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                ProductInfo.Initial = selectedRows[0]["Initial"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                ProductInfo.UseSubstrate = GetUseSubstrate(ProductInfo.ProductID);
                using ProductRegistration1Window window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleRePrint() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.ProductID = Convert.ToInt64(selectedRows[0]["ProductID"]);
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                ProductInfo.SheetPrintType = Convert.ToInt32(selectedRows[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
                ProductInfo.SerialType = Convert.ToInt32(selectedRows[0]["SerialType"] ?? throw new Exception("SerialType is null"));
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.CheckBin = Convert.ToInt32(selectedRows[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
                ProductInfo.Initial = selectedRows[0]["Initial"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                using RePrintWindow window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private void HandleSubstrateChange1() {
            var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ProductInfo.ProductID = Convert.ToInt64(selectedRows[0]["ProductID"]);
                ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                ProductInfo.SheetPrintType = Convert.ToInt32(selectedRows[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ProductInfo.RevisionGroup = Convert.ToInt32(selectedRows[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
                ProductInfo.UseSubstrate = GetUseSubstrate(ProductInfo.ProductID);
                using SubstrateChange1 window = new(ProductInfo);
                window.ShowDialog(this);
            }
        }
        private List<string> GetUseSubstrate(long productId) {
            // 使用基板リスト化+名前順にソート
            var useSubstrate = new List<string>();
            var useSubstrateRows = ProductInfo.ProductUseSubstrate.Select($"ProductID = '{productId}'");

            foreach (var row in useSubstrateRows) {
                var substrateId = row["SubstrateID"];

                var substrateRows = ProductInfo.SubstrateDataTable
                    .Select($"SubstrateID = {substrateId}");

                if (substrateRows.Length > 0) {
                    useSubstrate.Add(substrateRows[0]["SubstrateModel"].ToString() ?? string.Empty);
                }
            }
            useSubstrate.Sort();
            return useSubstrate;
        }
        // 履歴ボタン処理
        private void History() {
            ResetFields();
            try {
                DataRow[]? selectedRows = null;

                var listBox2 = CategoryListBox2.SelectedIndex == -1 ? string.Empty : $"AND ProductName = '{CategoryListBox2.SelectedItem}'";

                var listBox3 = RadioButtonNumber switch {
                    1 => CategoryListBox3.SelectedIndex == -1 ? string.Empty : $"AND SubstrateName = '{CategoryListBox3.SelectedItem}'",
                    2 or 3 or 4 => CategoryListBox3.SelectedIndex == -1 ? string.Empty : $"AND ProductType = '{CategoryListBox3.SelectedItem}'",
                    _ => string.Empty
                };
                selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' {listBox2} {listBox3}");

                if (selectedRows is not null && selectedRows.Length > 0) {
                    switch (RadioButtonNumber) {
                        case 1:
                            ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(listBox2)) {
                                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                            }

                            if (!string.IsNullOrEmpty(listBox3)) {
                                ProductInfo.SubstrateID = Convert.ToInt64(selectedRows[0]["SubstrateID"]);
                                ProductInfo.SubstrateName = selectedRows[0]["SubstrateName"].ToString() ?? string.Empty;
                                ProductInfo.SubstrateModel = selectedRows[0]["SubstrateModel"].ToString() ?? string.Empty;
                                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                            }
                            break;
                        case 2:
                        case 3:
                            ProductInfo.CategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;

                            if (!string.IsNullOrEmpty(listBox2)) {
                                ProductInfo.ProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                                ProductInfo.StockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                            }

                            if (!string.IsNullOrEmpty(listBox3)) {
                                ProductInfo.ProductID = Convert.ToInt64(selectedRows[0]["ProductID"]);
                                ProductInfo.ProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                                ProductInfo.ProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                                ProductInfo.SerialPrintType = Convert.ToInt32(selectedRows[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
                                ProductInfo.SheetPrintType = Convert.ToInt32(selectedRows[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
                                ProductInfo.RegType = Convert.ToInt32(selectedRows[0]["RegType"] ?? throw new Exception("RegType is null"));
                            }
                            break;
                    }

                    using HistoryWindow window = new(ProductInfo, RadioButtonNumber);
                    window.ShowDialog(this);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 処理カテゴリセレクト
        private void CategorySelect(object sender) {
            // Tagプロパティを使用してラジオボタンの識別子を取得
            Dictionary<string, (int Number, string Sql)> radioButtonMap = new() {
                ["1"] = (1, "SELECT * FROM M_SubstrateDef WHERE Visible = 1 ORDER BY SortNumber ASC;"),
                ["2"] = (2, "SELECT * FROM M_ProductDef WHERE Visible = 1 ORDER BY SortNumber ASC;"),
                ["3"] = (3, "SELECT * FROM M_ProductDef WHERE Visible = 1 AND SerialPrintType != 0 ORDER BY SortNumber ASC;"),
                ["4"] = (4, "SELECT * FROM M_ProductDef WHERE Visible = 1 AND (SheetPrintType = 2 OR SheetPrintType = 3) ORDER BY SortNumber ASC;")
            };

            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = false;
                CategoryListBox1.Items.Clear();
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();
                MainDataTable.Clear();

                if (sender is RadioButton selectedRadioButton && radioButtonMap.TryGetValue(selectedRadioButton.Tag?.ToString() ?? "", out var map)) {
                    RadioButtonNumber = map.Number;
                    using var con = new SqliteConnection(GetConnectionRegistration());
                    con.Open();

                    using var cmd = new SqliteCommand(map.Sql, con);
                    using var reader = cmd.ExecuteReader();
                    MainDataTable.Load(reader);
                }

                // CategoryName 列の重複を削除し、ソートする
                var categoryNames = MainDataTable.AsEnumerable()
                    .Select(row => row.Field<string?>("CategoryName") ?? string.Empty)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct()
                    .ToList();

                // リストボックスにアイテムを追加する
                CategoryListBox1.Items.AddRange([.. categoryNames]);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // 製品カテゴリセレクト
        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = RadioButtonNumber != 4;
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = [];

                var selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}'", "ProductName ASC");

                foreach (var row in selectedRows) {
                    var productName = row["ProductName"].ToString() ?? throw new Exception("ProductName is null");
                    productNames.Add(productName);
                }

                CategoryListBox2.Items.AddRange([.. productNames]);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox2Select() {
            try {
                RegisterButton.Enabled = false;
                HistoryButton.Enabled = RadioButtonNumber != 4;
                CategoryListBox3.Items.Clear();

                DataRow[] selectedRows;

                switch (RadioButtonNumber) {
                    case 1:
                        selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "SubstrateName ASC");
                        HashSet<string> substrateNames = [.. selectedRows.AsEnumerable()
                            .Select(x => x.Field<string>("SubstrateName"))
                            .Where(x => x is not null)
                            .Select(x => x!)];

                        CategoryListBox3.Items.AddRange([.. substrateNames]);
                        break;

                    case 2:
                    case 3:
                    case 4:
                        selectedRows = MainDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "ProductType ASC");
                        HashSet<string> productTypes = [.. selectedRows.AsEnumerable()
                            .Select(x => x.Field<string>("ProductType"))
                            .Where(x => x is not null)
                            .Select(x => x!)];

                        CategoryListBox3.Items.AddRange([.. productTypes]);
                        break;

                    default:
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox3Select() {
            try {
                if (CategoryListBox3.SelectedIndex == -1) { return; }
                RegisterButton.Enabled = Config.IsAuthorizedUser;
                HistoryButton.Enabled = true;

                switch (RadioButtonNumber) {
                    case 1:
                        ProductInfo.SubstrateName = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 2:
                    case 3:
                        ProductInfo.ProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        break;

                    case 4:
                        ProductInfo.ProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
                        HistoryButton.Enabled = false;
                        break;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                if (ProductInfo.CategoryItemNumber.Count >= 2) {
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
            ProductInfo.CategoryItemNumber.Clear();
            ProductInfo.CategoryProductType.Clear();
            ProductInfo.CategoryProductName.Clear();
            ProductInfo.CategoryType.Clear();
            ResetFields();
            Enabled = false;
        }
        private void ParseQRCodeInput() {
            try {
                string[] separator = ["//"];
                var arr = QRCodeTextBox.Text.Split(separator, StringSplitOptions.None);
                if (arr.Length == 1) {
                    ProductInfo.Proness2 = QRCodeTextBox.Text;
                    return;
                }
                if (arr.Length != 4) { throw new Exception("QRコードが正しくありません。"); }
                ProductInfo.Proness1 = arr[0];
                ProductInfo.Proness2 = arr[1];
                ProductInfo.Proness4 = int.TryParse(arr[2], out var result) ? result : throw new Exception("数量に数値以外が入力されています。");
                ProductInfo.Proness5 = arr[3];
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
                    ProductInfo.Proness1 = dr["手配製番"].ToString() ?? string.Empty;
                    ProductInfo.Proness2 = dr["品目番号"].ToString() ?? string.Empty;
                    ProductInfo.Proness3 = dr["品目名称"].ToString() ?? string.Empty;
                    ProductInfo.Proness4 = Convert.ToInt32(dr["手配数"] ?? throw new Exception("手配数 is null"));
                    ProductInfo.Proness5 = dr["請求先注番"].ToString() ?? string.Empty;
                }
            }
            if (string.IsNullOrEmpty(ProductInfo.Proness1)) { throw new Exception($"一致する情報がありません。{Environment.NewLine}手配管理番号:{QRCodeTextBox.Text}"); }
        }
        private void ProcessCategoryItemData() {
            var pattern = @"-(?:SMT|H|GH).*";
            var result = Regex.Replace(ProductInfo.Proness2, pattern, "");
            ProductInfo.Proness2 = result
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

            cmd.Parameters.Add("@StrProness2", SqliteType.Text).Value = ProductInfo.Proness2;
            using var dr = cmd.ExecuteReader();
            if (!dr.HasRows) { throw new Exception($"品目番号が見つかりません。\n品目番号:[{ProductInfo.Proness2}]"); }
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
            ProductInfo.CategoryItemNumber.Add(itemNumber);
            ProductInfo.CategoryProductName.Add(productName);
            ProductInfo.CategoryProductType.Add(productType);
            ProductInfo.CategorySubstrateName.Add(substrateName);
            ProductInfo.CategoryType.Add(type);
        }
        private int ShowDialogWindowForMultipleItems() {
            using SeveralDialogWindow window = new(ProductInfo);
            window.ShowDialog(this);
            return window.SelectedIndex;
        }
        private void HandleSelectedItem(int listIndex) {
            var productName = ProductInfo.CategoryProductName[listIndex];
            var productType = ProductInfo.CategoryProductType[listIndex];
            var substrateName = ProductInfo.CategorySubstrateName[listIndex];
            var type = ProductInfo.CategoryType[listIndex];

            switch (type) {
                case "1":
                    HandleSubstrateSelection(productName, substrateName);
                    break;
                case "2":
                    HandleProductSelection(productName, productType);
                    break;
                default:
                    throw new Exception($"一致する情報がありません。{Environment.NewLine}品目番号:{ProductInfo.Proness2}{Environment.NewLine}");
            }
        }
        private void HandleSubstrateSelection(string productName, string substrateName) {
            using SqliteConnection con = new(GetConnectionRegistration());
            con.Open();
            using (var cmd = new SqliteCommand("SELECT * FROM M_SubstrateDef;", con)) {
                using var reader = cmd.ExecuteReader();
                MainDataTable.Load(reader);
            }

            var substrateRet = MainDataTable.Select($"ProductName = '{productName}' AND SubstrateName = '{substrateName}'");
            OpenSubstrateRegistrationWindow(substrateRet);
        }
        private void OpenSubstrateRegistrationWindow(DataRow[] substrateRet) {
            ProductInfo.CategoryName = substrateRet[0]["CategoryName"].ToString() ?? string.Empty;
            ProductInfo.ProductName = substrateRet[0]["ProductName"].ToString() ?? string.Empty;
            ProductInfo.StockName = substrateRet[0]["StockName"].ToString() ?? string.Empty;
            ProductInfo.SubstrateName = substrateRet[0]["SubstrateName"].ToString() ?? string.Empty;
            ProductInfo.SubstrateModel = substrateRet[0]["SubstrateModel"].ToString() ?? string.Empty;
            ProductInfo.Initial = string.Empty;
            ProductInfo.RegType = Convert.ToInt32(substrateRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            ProductInfo.SerialPrintType = Convert.ToInt32(substrateRet[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
            ProductInfo.CheckBin = Convert.ToInt32(substrateRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            using SubstrateRegistrationWindow window = new(ProductInfo);
            window.ShowDialog(this);
        }
        private void HandleProductSelection(string productName, string productType) {
            using SqliteConnection con = new(GetConnectionRegistration());
            con.Open();
            using (var cmd = new SqliteCommand("SELECT * FROM M_ProductDef;", con)) {
                using var reader = cmd.ExecuteReader();
                MainDataTable.Load(reader);
            }

            var productRet = MainDataTable.Select($"ProductName = '{productName}' AND ProductType = '{productType}'");
            OpenProductRegistrationWindow(productRet);
        }
        private void OpenProductRegistrationWindow(DataRow[] productRet) {
            ProductInfo.CategoryName = productRet[0]["CategoryName"].ToString() ?? string.Empty;
            ProductInfo.ProductName = productRet[0]["ProductName"].ToString() ?? string.Empty;
            ProductInfo.StockName = productRet[0]["StockName"].ToString() ?? string.Empty;
            ProductInfo.ProductType = productRet[0]["ProductType"].ToString() ?? string.Empty;
            ProductInfo.ProductModel = productRet[0]["ProductModel"].ToString() ?? string.Empty;
            ProductInfo.Initial = productRet[0]["Initial"].ToString() ?? string.Empty;
            ProductInfo.RevisionGroup = Convert.ToInt32(productRet[0]["RevisionGroup"] ?? throw new Exception("RevisionGroup is null"));
            ProductInfo.RegType = Convert.ToInt32(productRet[0]["RegType"] ?? throw new Exception("RegType is null"));
            ProductInfo.SerialPrintType = Convert.ToInt32(productRet[0]["SerialPrintType"] ?? throw new Exception("SerialPrintType is null"));
            ProductInfo.SheetPrintType = Convert.ToInt32(productRet[0]["SheetPrintType"] ?? throw new Exception("SheetPrintType is null"));
            ProductInfo.CheckBin = Convert.ToInt32(productRet[0]["Checkbox"].ToString() ?? throw new Exception("Checkbox is null"), 2);
            ProductInfo.SerialType = Convert.ToInt32(productRet[0]["SerialType"] ?? throw new Exception("SerialType is null"));
            using ProductRegistration1Window window = new(ProductInfo);
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

            ProductInfo.FontSize = _fontSize;
            Font = new System.Drawing.Font(ProductInfo.FontName, ProductInfo.FontSize);
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
