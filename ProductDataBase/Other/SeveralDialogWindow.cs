using static ProductDatabase.MainWindow;

namespace ProductDatabase {
    public partial class SeveralDialogWindow : Form {

        public ProductInfomation ProductInfo { get; set; } = new ProductInfomation();

        public int SelectedIndex { get; private set; }

        public SeveralDialogWindow() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            SeveralListBox.Items.Clear();
            Font = new Font(ProductInfo.FontName, ProductInfo.FontSize);
            var j = 0;

            foreach (var category in ProductInfo.Category11) {
                var type = (ProductInfo.Category14[j] == "1") ? "基板" : "製品";

                if (!String.IsNullOrEmpty(ProductInfo.Category11[j])) { SeveralListBox.Items.Add($"[{type}]  [{ProductInfo.Category13[j]}]  [{ProductInfo.Category12[j]}]  [{category}]"); }
                j++;
            }

            SeveralListBox.SelectedIndex = 0;
        }

        private void SeveralDialogWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OkButton_Click(object sender, EventArgs e) {
            SelectedIndex = SeveralListBox.SelectedIndex;
            Close();
        }
        private void SeveralListBox_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == (char)Keys.Enter) { OKButton.PerformClick(); }
        }
    }
}
