using ProductDatabase.Data;
using ProductDatabase.Models;

namespace ProductDatabase.MasterManagement {
    public partial class PersonManagementWindow : Form {

        public PersonManagementWindow() {
            InitializeComponent();
        }

        private void PersonManagementWindow_Load(object sender, EventArgs e) {
            LoadPersonList();
        }

        private void LoadPersonList() {
            try {
                var persons = PersonRepository.GetAll();
                var displayData = persons.Select(p => new {
                    p.PersonID,
                    p.PersonName,
                    IsActive = p.IsActive != 0
                }).ToList();
                PersonDataGridView.DataSource = displayData;

                if (PersonDataGridView.Rows.Count > 0) {
                    PersonDataGridView.Rows[0].Selected = true;
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddButton_Click(object sender, EventArgs e) {
            using (var dialog = new PersonMasterEditDialog(null)) {
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    LoadPersonList();
                }
            }
        }

        private void EditButton_Click(object sender, EventArgs e) {
            if (PersonDataGridView.SelectedRows.Count == 0) {
                MessageBox.Show("編集する担当者を選択してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selectedRow = PersonDataGridView.SelectedRows[0];
            var personId = (long)selectedRow.Cells["IdColumn"].Value;
            var personName = selectedRow.Cells["NameColumn"].Value.ToString() ?? string.Empty;
            var isActive = (bool)selectedRow.Cells["IsActiveColumn"].Value ? 1 : 0;

            var person = new PersonDef {
                PersonID = personId,
                PersonName = personName,
                IsActive = isActive
            };

            using (var dialog = new PersonMasterEditDialog(person)) {
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    LoadPersonList();
                }
            }
        }

    }
}
