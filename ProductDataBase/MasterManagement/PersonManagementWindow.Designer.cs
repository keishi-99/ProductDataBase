using System.Windows.Forms;

namespace ProductDatabase.MasterManagement {
    partial class PersonManagementWindow {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.PersonDataGridView = new DataGridView();
            this.IdColumn = new DataGridViewTextBoxColumn();
            this.NameColumn = new DataGridViewTextBoxColumn();
            this.IsActiveColumn = new DataGridViewCheckBoxColumn();
            this.ButtonPanel = new Panel();
            this.AddButton = new Button();
            this.EditButton = new Button();
            this.DeleteButton = new Button();
            this.CloseButton = new Button();
            ((System.ComponentModel.ISupportInitialize)(this.PersonDataGridView)).BeginInit();
            this.ButtonPanel.SuspendLayout();
            this.SuspendLayout();
            //
            // PersonDataGridView
            //
            this.PersonDataGridView.AllowUserToAddRows = false;
            this.PersonDataGridView.AllowUserToDeleteRows = false;
            this.PersonDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.PersonDataGridView.Columns.AddRange(new DataGridViewColumn[] {
            this.IdColumn,
            this.NameColumn,
            this.IsActiveColumn});
            this.PersonDataGridView.Dock = DockStyle.Fill;
            this.PersonDataGridView.Location = new Point(0, 0);
            this.PersonDataGridView.Name = "PersonDataGridView";
            this.PersonDataGridView.ReadOnly = true;
            this.PersonDataGridView.RowTemplate.Height = 25;
            this.PersonDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.PersonDataGridView.Size = new Size(484, 311);
            this.PersonDataGridView.TabIndex = 0;
            //
            // IdColumn
            //
            this.IdColumn.DataPropertyName = "PersonID";
            this.IdColumn.HeaderText = "ID";
            this.IdColumn.Name = "IdColumn";
            this.IdColumn.ReadOnly = true;
            this.IdColumn.Width = 50;
            //
            // NameColumn
            //
            this.NameColumn.DataPropertyName = "PersonName";
            this.NameColumn.HeaderText = "名前";
            this.NameColumn.Name = "NameColumn";
            this.NameColumn.ReadOnly = true;
            this.NameColumn.Width = 250;
            //
            // IsActiveColumn
            //
            this.IsActiveColumn.DataPropertyName = "IsActive";
            this.IsActiveColumn.HeaderText = "有効";
            this.IsActiveColumn.Name = "IsActiveColumn";
            this.IsActiveColumn.ReadOnly = true;
            this.IsActiveColumn.Width = 60;
            //
            // ButtonPanel
            //
            this.ButtonPanel.Controls.Add(this.AddButton);
            this.ButtonPanel.Controls.Add(this.EditButton);
            this.ButtonPanel.Controls.Add(this.DeleteButton);
            this.ButtonPanel.Controls.Add(this.CloseButton);
            this.ButtonPanel.Dock = DockStyle.Bottom;
            this.ButtonPanel.Location = new Point(0, 311);
            this.ButtonPanel.Name = "ButtonPanel";
            this.ButtonPanel.Size = new Size(484, 48);
            this.ButtonPanel.TabIndex = 1;
            //
            // AddButton
            //
            this.AddButton.Location = new Point(12, 10);
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new Size(100, 28);
            this.AddButton.TabIndex = 0;
            this.AddButton.Text = "追加";
            this.AddButton.UseVisualStyleBackColor = true;
            this.AddButton.Click += this.AddButton_Click;
            //
            // EditButton
            //
            this.EditButton.Location = new Point(120, 10);
            this.EditButton.Name = "EditButton";
            this.EditButton.Size = new Size(100, 28);
            this.EditButton.TabIndex = 1;
            this.EditButton.Text = "編集";
            this.EditButton.UseVisualStyleBackColor = true;
            this.EditButton.Click += this.EditButton_Click;
            //
            // DeleteButton
            //
            this.DeleteButton.Location = new Point(228, 10);
            this.DeleteButton.Name = "DeleteButton";
            this.DeleteButton.Size = new Size(100, 28);
            this.DeleteButton.TabIndex = 2;
            this.DeleteButton.Text = "削除";
            this.DeleteButton.UseVisualStyleBackColor = true;
            this.DeleteButton.Click += this.DeleteButton_Click;
            //
            // CloseButton
            //
            this.CloseButton.DialogResult = DialogResult.Cancel;
            this.CloseButton.Location = new Point(372, 10);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new Size(100, 28);
            this.CloseButton.TabIndex = 3;
            this.CloseButton.Text = "閉じる";
            this.CloseButton.UseVisualStyleBackColor = true;
            //
            // PersonManagementWindow
            //
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new Size(484, 359);
            this.Controls.Add(this.PersonDataGridView);
            this.Controls.Add(this.ButtonPanel);
            this.Font = new Font("Meiryo UI", 9F);
            this.Name = "PersonManagementWindow";
            this.ShowIcon = false;
            this.Text = "担当者マスター";
            this.Load += this.PersonManagementWindow_Load;
            ((System.ComponentModel.ISupportInitialize)(this.PersonDataGridView)).EndInit();
            this.ButtonPanel.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private DataGridView PersonDataGridView;
        private DataGridViewTextBoxColumn IdColumn;
        private DataGridViewTextBoxColumn NameColumn;
        private DataGridViewCheckBoxColumn IsActiveColumn;
        private Panel ButtonPanel;
        private Button AddButton;
        private Button EditButton;
        private Button DeleteButton;
        private Button CloseButton;
    }
}
