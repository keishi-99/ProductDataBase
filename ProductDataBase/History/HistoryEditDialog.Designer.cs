namespace ProductDatabase.History {
    partial class HistoryEditDialog {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.ReadOnlyGroupBox = new GroupBox();
            this.IdLabel = new Label();
            this.IdValueLabel = new Label();
            this.CategoryNameLabel = new Label();
            this.CategoryNameValueLabel = new Label();
            this.ProductNameLabel = new Label();
            this.ProductNameValueLabel = new Label();
            this.ProductModelLabel = new Label();
            this.ProductModelValueLabel = new Label();
            this.RegDateLabel = new Label();
            this.RegDateValueLabel = new Label();
            this.RevisionLabel = new Label();
            this.RevisionValueLabel = new Label();
            this.RevisionGroupLabel = new Label();
            this.RevisionGroupValueLabel = new Label();
            this.EditGroupBox = new GroupBox();
            this.OrderNumberLabel = new Label();
            this.OrderNumberTextBox = new TextBox();
            this.ProductNumberLabel = new Label();
            this.ProductNumberTextBox = new TextBox();
            this.OLesNumberLabel = new Label();
            this.OLesNumberTextBox = new TextBox();
            this.PersonLabel = new Label();
            this.PersonTextBox = new TextBox();
            this.CommentLabel = new Label();
            this.CommentTextBox = new TextBox();
            this.SaveButton = new Button();
            this.DialogCancelButton = new Button();
            this.ReadOnlyGroupBox.SuspendLayout();
            this.EditGroupBox.SuspendLayout();
            this.SuspendLayout();
            //
            // ReadOnlyGroupBox
            //
            this.ReadOnlyGroupBox.Controls.Add(this.IdLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.IdValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.CategoryNameLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.CategoryNameValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.ProductNameLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.ProductNameValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.ProductModelLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.ProductModelValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RegDateLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RegDateValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RevisionLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RevisionValueLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RevisionGroupLabel);
            this.ReadOnlyGroupBox.Controls.Add(this.RevisionGroupValueLabel);
            this.ReadOnlyGroupBox.Location = new Point(12, 12);
            this.ReadOnlyGroupBox.Name = "ReadOnlyGroupBox";
            this.ReadOnlyGroupBox.Size = new Size(460, 192);
            this.ReadOnlyGroupBox.TabIndex = 0;
            this.ReadOnlyGroupBox.TabStop = false;
            this.ReadOnlyGroupBox.Text = "参照";
            //
            // IdLabel
            //
            this.IdLabel.AutoSize = true;
            this.IdLabel.Location = new Point(8, 24);
            this.IdLabel.Name = "IdLabel";
            this.IdLabel.Size = new Size(24, 15);
            this.IdLabel.TabIndex = 0;
            this.IdLabel.Text = "ID:";
            //
            // IdValueLabel
            //
            this.IdValueLabel.AutoSize = true;
            this.IdValueLabel.Location = new Point(130, 24);
            this.IdValueLabel.Name = "IdValueLabel";
            this.IdValueLabel.Size = new Size(14, 15);
            this.IdValueLabel.TabIndex = 1;
            this.IdValueLabel.Text = "-";
            //
            // CategoryNameLabel
            //
            this.CategoryNameLabel.AutoSize = true;
            this.CategoryNameLabel.Location = new Point(8, 48);
            this.CategoryNameLabel.Name = "CategoryNameLabel";
            this.CategoryNameLabel.Size = new Size(56, 15);
            this.CategoryNameLabel.TabIndex = 2;
            this.CategoryNameLabel.Text = "カテゴリ:";
            //
            // CategoryNameValueLabel
            //
            this.CategoryNameValueLabel.AutoSize = true;
            this.CategoryNameValueLabel.Location = new Point(130, 48);
            this.CategoryNameValueLabel.Name = "CategoryNameValueLabel";
            this.CategoryNameValueLabel.Size = new Size(14, 15);
            this.CategoryNameValueLabel.TabIndex = 3;
            this.CategoryNameValueLabel.Text = "-";
            //
            // ProductNameLabel
            //
            this.ProductNameLabel.AutoSize = true;
            this.ProductNameLabel.Location = new Point(8, 72);
            this.ProductNameLabel.Name = "ProductNameLabel";
            this.ProductNameLabel.Size = new Size(47, 15);
            this.ProductNameLabel.TabIndex = 4;
            this.ProductNameLabel.Text = "製品名:";
            //
            // ProductNameValueLabel
            //
            this.ProductNameValueLabel.AutoSize = true;
            this.ProductNameValueLabel.Location = new Point(130, 72);
            this.ProductNameValueLabel.Name = "ProductNameValueLabel";
            this.ProductNameValueLabel.Size = new Size(14, 15);
            this.ProductNameValueLabel.TabIndex = 5;
            this.ProductNameValueLabel.Text = "-";
            //
            // ProductModelLabel
            //
            this.ProductModelLabel.AutoSize = true;
            this.ProductModelLabel.Location = new Point(8, 96);
            this.ProductModelLabel.Name = "ProductModelLabel";
            this.ProductModelLabel.Size = new Size(55, 15);
            this.ProductModelLabel.TabIndex = 6;
            this.ProductModelLabel.Text = "製品型式:";
            //
            // ProductModelValueLabel
            //
            this.ProductModelValueLabel.AutoSize = true;
            this.ProductModelValueLabel.Location = new Point(130, 96);
            this.ProductModelValueLabel.Name = "ProductModelValueLabel";
            this.ProductModelValueLabel.Size = new Size(14, 15);
            this.ProductModelValueLabel.TabIndex = 7;
            this.ProductModelValueLabel.Text = "-";
            //
            // RegDateLabel
            //
            this.RegDateLabel.AutoSize = true;
            this.RegDateLabel.Location = new Point(8, 120);
            this.RegDateLabel.Name = "RegDateLabel";
            this.RegDateLabel.Size = new Size(47, 15);
            this.RegDateLabel.TabIndex = 8;
            this.RegDateLabel.Text = "登録日:";
            //
            // RegDateValueLabel
            //
            this.RegDateValueLabel.AutoSize = true;
            this.RegDateValueLabel.Location = new Point(130, 120);
            this.RegDateValueLabel.Name = "RegDateValueLabel";
            this.RegDateValueLabel.Size = new Size(14, 15);
            this.RegDateValueLabel.TabIndex = 9;
            this.RegDateValueLabel.Text = "-";
            //
            // RevisionLabel
            //
            this.RevisionLabel.AutoSize = true;
            this.RevisionLabel.Location = new Point(8, 144);
            this.RevisionLabel.Name = "RevisionLabel";
            this.RevisionLabel.Size = new Size(31, 15);
            this.RevisionLabel.TabIndex = 10;
            this.RevisionLabel.Text = "Rev:";
            //
            // RevisionValueLabel
            //
            this.RevisionValueLabel.AutoSize = true;
            this.RevisionValueLabel.Location = new Point(130, 144);
            this.RevisionValueLabel.Name = "RevisionValueLabel";
            this.RevisionValueLabel.Size = new Size(14, 15);
            this.RevisionValueLabel.TabIndex = 11;
            this.RevisionValueLabel.Text = "-";
            //
            // RevisionGroupLabel
            //
            this.RevisionGroupLabel.AutoSize = true;
            this.RevisionGroupLabel.Location = new Point(8, 168);
            this.RevisionGroupLabel.Name = "RevisionGroupLabel";
            this.RevisionGroupLabel.Size = new Size(47, 15);
            this.RevisionGroupLabel.TabIndex = 12;
            this.RevisionGroupLabel.Text = "Group:";
            //
            // RevisionGroupValueLabel
            //
            this.RevisionGroupValueLabel.AutoSize = true;
            this.RevisionGroupValueLabel.Location = new Point(130, 168);
            this.RevisionGroupValueLabel.Name = "RevisionGroupValueLabel";
            this.RevisionGroupValueLabel.Size = new Size(14, 15);
            this.RevisionGroupValueLabel.TabIndex = 13;
            this.RevisionGroupValueLabel.Text = "-";
            //
            // EditGroupBox
            //
            this.EditGroupBox.Controls.Add(this.OrderNumberLabel);
            this.EditGroupBox.Controls.Add(this.OrderNumberTextBox);
            this.EditGroupBox.Controls.Add(this.ProductNumberLabel);
            this.EditGroupBox.Controls.Add(this.ProductNumberTextBox);
            this.EditGroupBox.Controls.Add(this.OLesNumberLabel);
            this.EditGroupBox.Controls.Add(this.OLesNumberTextBox);
            this.EditGroupBox.Controls.Add(this.PersonLabel);
            this.EditGroupBox.Controls.Add(this.PersonTextBox);
            this.EditGroupBox.Controls.Add(this.CommentLabel);
            this.EditGroupBox.Controls.Add(this.CommentTextBox);
            this.EditGroupBox.Location = new Point(12, 214);
            this.EditGroupBox.Name = "EditGroupBox";
            this.EditGroupBox.Size = new Size(460, 227);
            this.EditGroupBox.TabIndex = 1;
            this.EditGroupBox.TabStop = false;
            this.EditGroupBox.Text = "編集";
            //
            // OrderNumberLabel
            //
            this.OrderNumberLabel.AutoSize = true;
            this.OrderNumberLabel.Location = new Point(8, 28);
            this.OrderNumberLabel.Name = "OrderNumberLabel";
            this.OrderNumberLabel.Size = new Size(55, 15);
            this.OrderNumberLabel.TabIndex = 0;
            this.OrderNumberLabel.Text = "注文番号:";
            //
            // OrderNumberTextBox
            //
            this.OrderNumberTextBox.Location = new Point(130, 25);
            this.OrderNumberTextBox.MaxLength = 50;
            this.OrderNumberTextBox.Name = "OrderNumberTextBox";
            this.OrderNumberTextBox.Size = new Size(150, 23);
            this.OrderNumberTextBox.TabIndex = 0;
            //
            // ProductNumberLabel
            //
            this.ProductNumberLabel.AutoSize = true;
            this.ProductNumberLabel.Location = new Point(8, 57);
            this.ProductNumberLabel.Name = "ProductNumberLabel";
            this.ProductNumberLabel.Size = new Size(55, 15);
            this.ProductNumberLabel.TabIndex = 2;
            this.ProductNumberLabel.Text = "製造番号:";
            //
            // ProductNumberTextBox
            //
            this.ProductNumberTextBox.Location = new Point(130, 54);
            this.ProductNumberTextBox.MaxLength = 50;
            this.ProductNumberTextBox.Name = "ProductNumberTextBox";
            this.ProductNumberTextBox.Size = new Size(150, 23);
            this.ProductNumberTextBox.TabIndex = 1;
            //
            // OLesNumberLabel
            //
            this.OLesNumberLabel.AutoSize = true;
            this.OLesNumberLabel.Location = new Point(8, 86);
            this.OLesNumberLabel.Name = "OLesNumberLabel";
            this.OLesNumberLabel.Size = new Size(63, 15);
            this.OLesNumberLabel.TabIndex = 4;
            this.OLesNumberLabel.Text = "OLes番号:";
            //
            // OLesNumberTextBox
            //
            this.OLesNumberTextBox.Location = new Point(130, 83);
            this.OLesNumberTextBox.MaxLength = 50;
            this.OLesNumberTextBox.Name = "OLesNumberTextBox";
            this.OLesNumberTextBox.Size = new Size(150, 23);
            this.OLesNumberTextBox.TabIndex = 2;
            //
            // PersonLabel
            //
            this.PersonLabel.AutoSize = true;
            this.PersonLabel.Location = new Point(8, 115);
            this.PersonLabel.Name = "PersonLabel";
            this.PersonLabel.Size = new Size(47, 15);
            this.PersonLabel.TabIndex = 6;
            this.PersonLabel.Text = "担当者:";
            //
            // PersonTextBox
            //
            this.PersonTextBox.Location = new Point(130, 112);
            this.PersonTextBox.MaxLength = 50;
            this.PersonTextBox.Name = "PersonTextBox";
            this.PersonTextBox.Size = new Size(150, 23);
            this.PersonTextBox.TabIndex = 3;
            //
            // CommentLabel
            //
            this.CommentLabel.AutoSize = true;
            this.CommentLabel.Location = new Point(8, 148);
            this.CommentLabel.Name = "CommentLabel";
            this.CommentLabel.Size = new Size(55, 15);
            this.CommentLabel.TabIndex = 8;
            this.CommentLabel.Text = "コメント:";
            //
            // CommentTextBox
            //
            this.CommentTextBox.Location = new Point(130, 145);
            this.CommentTextBox.MaxLength = 200;
            this.CommentTextBox.Multiline = true;
            this.CommentTextBox.Name = "CommentTextBox";
            this.CommentTextBox.ScrollBars = ScrollBars.Vertical;
            this.CommentTextBox.Size = new Size(320, 60);
            this.CommentTextBox.TabIndex = 4;
            //
            // SaveButton
            //
            this.SaveButton.Location = new Point(290, 451);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new Size(88, 28);
            this.SaveButton.TabIndex = 5;
            this.SaveButton.Text = "保存";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += this.SaveButton_Click;
            //
            // DialogCancelButton
            //
            this.DialogCancelButton.Location = new Point(384, 451);
            this.DialogCancelButton.Name = "DialogCancelButton";
            this.DialogCancelButton.Size = new Size(88, 28);
            this.DialogCancelButton.TabIndex = 6;
            this.DialogCancelButton.Text = "キャンセル";
            this.DialogCancelButton.UseVisualStyleBackColor = true;
            this.DialogCancelButton.Click += this.CancelButton_Click;
            //
            // HistoryEditDialog
            //
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(484, 491);
            this.Controls.Add(this.ReadOnlyGroupBox);
            this.Controls.Add(this.EditGroupBox);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.DialogCancelButton);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HistoryEditDialog";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "製品履歴編集";
            this.Load += this.HistoryEditDialog_Load;
            this.ReadOnlyGroupBox.ResumeLayout(false);
            this.ReadOnlyGroupBox.PerformLayout();
            this.EditGroupBox.ResumeLayout(false);
            this.EditGroupBox.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        // 参照グループ
        private GroupBox ReadOnlyGroupBox;
        private Label IdLabel;
        private Label IdValueLabel;
        private Label CategoryNameLabel;
        private Label CategoryNameValueLabel;
        private Label ProductNameLabel;
        private Label ProductNameValueLabel;
        private Label ProductModelLabel;
        private Label ProductModelValueLabel;
        private Label RegDateLabel;
        private Label RegDateValueLabel;
        private Label RevisionLabel;
        private Label RevisionValueLabel;
        private Label RevisionGroupLabel;
        private Label RevisionGroupValueLabel;

        // 編集グループ
        private GroupBox EditGroupBox;
        private Label OrderNumberLabel;
        private TextBox OrderNumberTextBox;
        private Label ProductNumberLabel;
        private TextBox ProductNumberTextBox;
        private Label OLesNumberLabel;
        private TextBox OLesNumberTextBox;
        private Label PersonLabel;
        private TextBox PersonTextBox;
        private Label CommentLabel;
        private TextBox CommentTextBox;

        // ボタン
        private Button SaveButton;
        private Button DialogCancelButton;
    }
}
