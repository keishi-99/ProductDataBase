using Dapper;
using Microsoft.Data.Sqlite;
using System.Data;
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
            var strSqlQuery = "SELECT * FROM M_ProductDef WHERE Visible = 1 ORDER BY SortNumber ASC;";

            using SqliteConnection con = new(GetConnectionRegistration());
            con.Open();
            using (var reader = con.ExecuteReader(strSqlQuery)) {
                ServiceInfo.ServiceDataTable.Load(reader);
            }

            var categoryNames = ServiceInfo.ServiceDataTable.AsEnumerable()
                .Select(row => row.Field<string?>("CategoryName") ?? string.Empty)
                .Distinct()
                .ToList();

            CategoryListBox1.Items.AddRange([.. categoryNames]);

            RegisterButton.Enabled = false;
        }

        private void CategoryListBox1Select() {
            RegisterButton.Enabled = false;
            CategoryListBox2.Items.Clear();
            CategoryListBox3.Items.Clear();

            HashSet<string> productNames = [];

            var selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductModel <> 'SERVICE'", "ProductName ASC");

            foreach (var row in selectedRows) {
                var productName = row["ProductName"].ToString() ?? throw new Exception("ProductName is null");
                productNames.Add(productName);
            }

            CategoryListBox2.Items.AddRange([.. productNames]);
        }
        private void CategoryListBox2Select() {
            RegisterButton.Enabled = false;
            CategoryListBox3.Items.Clear();

            DataRow[] selectedRows;

            selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}'", "ProductType ASC");
            HashSet<string> productTypes = [.. selectedRows.AsEnumerable()
                    .Select(x => x.Field<string>("ProductType"))
                    .Where(x => x is not null)
                    .Select(x => x!)];

            CategoryListBox3.Items.AddRange([.. productTypes]);
        }
        private void CategoryListBox3Select() {
            if (CategoryListBox3.SelectedIndex == -1) { return; ; }
            RegisterButton.Enabled = true;
            ServiceInfo.ServiceProductType = CategoryListBox3.SelectedItem?.ToString() ?? string.Empty;
        }
        private DialogResult Registration() {
            var selectedRows = ServiceInfo.ServiceDataTable.Select($"CategoryName = '{CategoryListBox1.SelectedItem}' AND ProductName = '{CategoryListBox2.SelectedItem}' AND ProductType = '{CategoryListBox3.SelectedItem}'");

            if (selectedRows.Length > 0) {
                ServiceInfo.ServiceProductID = Convert.ToInt64(selectedRows[0]["ProductID"]);
                ServiceInfo.ServiceCategoryName = selectedRows[0]["CategoryName"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductName = selectedRows[0]["ProductName"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductType = selectedRows[0]["ProductType"].ToString() ?? string.Empty;
                ServiceInfo.ServiceProductModel = selectedRows[0]["ProductModel"].ToString() ?? string.Empty;
                ServiceInfo.ServiceUseSubstrates = GetUseSubstrates(ServiceInfo.ServiceProductID);
                DialogResult = DialogResult.OK;
            }
            else { DialogResult = DialogResult.Cancel; }

            Close();
            return DialogResult.OK;
        }
        private static List<SubstrateInfo> GetUseSubstrates(long productId) {

            var productUseSubstrate = new DataTable();

            using var con = new SqliteConnection(GetConnectionRegistration());
            con.Open();

            using (var reader = con.ExecuteReader($"SELECT * FROM {Constants.VProductUseSubstrate};")) {
                productUseSubstrate.Load(reader);
            }

            return [.. productUseSubstrate.AsEnumerable()
                    .Where(r => Convert.ToInt32(r["P_ProductID"]) == productId)
                    .Select(r => new SubstrateInfo {
                        SubstrateID = Convert.ToInt32(r["S_SubstrateID"]),
                        SubstrateName = r["SubstrateName"]?.ToString() ?? "",
                        SubstrateModel = r["SubstrateModel"]?.ToString() ?? ""
                    })];
        }

        private void ServiceForm_Load(object sender, EventArgs e) { LoadEvents(); }
        private void CategoryListBox1_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox1Select(); }
        private void CategoryListBox2_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox2Select(); }
        private void CategoryListBox3_SelectedIndexChanged(object sender, EventArgs e) { CategoryListBox3Select(); }
        private void RegisterButton_Click(object sender, EventArgs e) { Registration(); }
    }
}
