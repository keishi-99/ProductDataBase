namespace ProductDatabase {
    public partial class SeveralDialogWindow : Form {

        public string StrFontName { get; set; } = "Meiryo UI";
        public int IntFontSize { get; set; } = 9;

        public int SelectedIndex = 0;

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
            int _j = 0;

            foreach (string Category in ListCategory11) {
                string _type = (ListCategory14[_j] == "1") ? "基板" : "製品";

                if (!string.IsNullOrEmpty(ListCategory11[_j])) { SeveralListBox.Items.Add($"[{_type}]  [{ListCategory13[_j]}]  [{ListCategory12[_j]}]  [{Category}]"); }
                _j++;
            }

            SeveralListBox.SelectedIndex = 0;
        }

        private void SeveralDialogWindow_Load(object sender, EventArgs e) { LoadEvents(); }
        private void OkButton_Click(object sender, EventArgs e) {
            SelectedIndex = SeveralListBox.SelectedIndex;
            DialogResult = DialogResult.OK;
            Close();
        }
        private void SeveralListBox_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == (char)Keys.Enter) { OKButton.PerformClick(); }
        }
    }
}
