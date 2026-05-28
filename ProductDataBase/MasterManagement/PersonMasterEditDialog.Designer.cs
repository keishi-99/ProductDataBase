namespace ProductDatabase.MasterManagement {
    partial class PersonMasterEditDialog {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.PersonNameLabel = new Label();
            this.PersonNameTextBox = new TextBox();
            this.IsActiveCheckBox = new CheckBox();
            this.SaveButton = new Button();
            this.DialogCancelButton = new Button();
            this.SuspendLayout();
            //
            // PersonNameLabel
            //
            this.PersonNameLabel.AutoSize = true;
            this.PersonNameLabel.Location = new Point(12, 15);
            this.PersonNameLabel.Name = "PersonNameLabel";
            this.PersonNameLabel.Size = new Size(40, 15);
            this.PersonNameLabel.TabIndex = 0;
            this.PersonNameLabel.Text = "名前:";
            //
            // PersonNameTextBox
            //
            this.PersonNameTextBox.Location = new Point(120, 12);
            this.PersonNameTextBox.MaxLength = 50;
            this.PersonNameTextBox.Name = "PersonNameTextBox";
            this.PersonNameTextBox.Size = new Size(200, 23);
            this.PersonNameTextBox.TabIndex = 1;
            //
            // IsActiveCheckBox
            //
            this.IsActiveCheckBox.AutoSize = true;
            this.IsActiveCheckBox.Location = new Point(120, 48);
            this.IsActiveCheckBox.Name = "IsActiveCheckBox";
            this.IsActiveCheckBox.Size = new Size(54, 19);
            this.IsActiveCheckBox.TabIndex = 2;
            this.IsActiveCheckBox.Text = "有効";
            this.IsActiveCheckBox.UseVisualStyleBackColor = true;
            //
            // SaveButton
            //
            this.SaveButton.Location = new Point(120, 83);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new Size(80, 28);
            this.SaveButton.TabIndex = 3;
            this.SaveButton.Text = "保存";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += this.SaveButton_Click;
            //
            // DialogCancelButton
            //
            this.DialogCancelButton.Location = new Point(210, 83);
            this.DialogCancelButton.Name = "DialogCancelButton";
            this.DialogCancelButton.Size = new Size(80, 28);
            this.DialogCancelButton.TabIndex = 4;
            this.DialogCancelButton.Text = "キャンセル";
            this.DialogCancelButton.UseVisualStyleBackColor = true;
            this.DialogCancelButton.Click += this.DialogCancelButton_Click;
            //
            // PersonMasterEditDialog
            //
            this.AcceptButton = this.SaveButton;
            this.CancelButton = this.DialogCancelButton;
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(340, 126);
            this.Controls.Add(this.PersonNameLabel);
            this.Controls.Add(this.PersonNameTextBox);
            this.Controls.Add(this.IsActiveCheckBox);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.DialogCancelButton);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PersonMasterEditDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "担当者";
            this.Load += this.PersonMasterEditDialog_Load;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Label PersonNameLabel;
        private TextBox PersonNameTextBox;
        private CheckBox IsActiveCheckBox;
        private Button SaveButton;
        private Button DialogCancelButton;
    }
}
