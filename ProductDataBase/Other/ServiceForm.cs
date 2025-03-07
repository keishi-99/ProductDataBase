using System.Data;
using System.Data.SQLite;
using static ProductDatabase.MainWindow;
using static ProductDatabase.ProductRegistration2Window;

namespace ProductDatabase.Other {
    public partial class ServiceForm : Form {

        public ServiceInformation ServiceInfo { get; }
        public ServiceForm(ServiceInformation serviceInfo) {
            InitializeComponent();
            ServiceInfo = serviceInfo;
        }

        private void LoadEvents() {
            var strSqlQuery = """SELECT * FROM Product WHERE Visible = 1 ORDER BY SortNumber ASC;""";

            using (SQLiteConnection con = new(GetConnectionInformation()))
            using (SQLiteDataAdapter adapter = new(strSqlQuery, con)) {
                adapter.Fill(ServiceInfo.ServiceDataTable);
            }

            // CategoryName 列の重複を削除し、ソートする
            var categoryNames = ServiceInfo.ServiceDataTable.AsEnumerable()
                .Select(row => row.Field<string?>("CategoryName") ?? string.Empty)
                .Distinct()
                .ToList();

            // リストボックスにアイテムを追加する
            CategoryListBox1.Items.AddRange([.. categoryNames]);

            RegisterButton.Enabled = false;
        }

        private void CategoryListBox1Select() {
            try {
                RegisterButton.Enabled = false;
                CategoryListBox2.Items.Clear();
                CategoryListBox3.Items.Clear();

                HashSet<string> productNames = [];

                var selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}'", "ProductName ASC");

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
                CategoryListBox3.Items.Clear();

                DataRow[] selectedRows;

                selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "ProductType ASC");
                HashSet<string> productTypes = [.. selectedRows.AsEnumerable()
                    .Select(x => x.Field<string>("ProductType"))
                    .Where(x => x != null)
                    .Select(x => x!)];

                CategoryListBox3.Items.AddRange([.. productTypes]);

            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CategoryListBox3Select() {
            try {
                if (CategoryListBox3.SelectedIndex == -1) { return; ; }
                RegisterButton.Enabled = true;
                ServiceInfo.ServiceProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private DialogResult Registration() {
            var selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ServiceInfo.ServiceCategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ServiceInfo.ServiceStockName = selectedRows[0]["StockName"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                var useSubstrate = selectedRows[0]["UseSubstrate"].ToString() ?? string.Empty;
                ServiceInfo.ServiveUseSubstrate = useSubstrate.Split(",");
            }
            Close();
            return DialogResult.OK;
        }

        private void ServiceForm_Load(object sender, EventArgs e) { LoadEvents(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
    }
}
