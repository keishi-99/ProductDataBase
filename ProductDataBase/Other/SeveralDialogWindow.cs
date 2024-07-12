namespace ProductDatabase {
    public partial class SeveralDialogWindow : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public int selectedIndex = 0;

        public List<string> ListCategory11 { get; set; } = [];
        public List<string> ListCategory12 { get; set; } = [];
        public List<string> ListCategory13 { get; set; } = [];
        public List<string> ListCategory14 { get; set; } = [];

        public SeveralDialogWindow() {
            InitializeComponent();
        }

        // ロードイベント
        private void LoadEvents() {
            SeveralListBox.Items.Clear();
            Font = new Font(StrFontName, IntFontSize);
            var j = 0;

            foreach (var category in ListCategory11) {
                var type = (ListCategory14[j] == "1") ? "基板" : "製品";

                if (!string.IsNullOrEmpty(ListCategory11[j])) { SeveralListBox.Items.Add($"[{type}]  [{ListCategory13[j]}]  [{ListCategory12[j]}]  [{category}]"); }
                j++;
            }

            SeveralListBox.SelectedIndex = 0;
        }

        private void SeveralDialogWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OkButton_Click(object sender, EventArgs e) {
            selectedIndex = SeveralListBox.SelectedIndex;
            DialogResult = DialogResult.OK;
            Close();
        }
        private void SeveralListBox_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == (char)Keys.Enter) { OKButton.PerformClick(); }
        }
    }
}
