namespace ProductDatabase {
    partial class ProductAttributeSelectWindow {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent() {
            this.PowerLabel = new Label();
            this.PowerListBox = new ListBox();
            this.MountLabel = new Label();
            this.MountListBox = new ListBox();
            this.AdiLabel = new Label();
            this.AdiListBox = new ListBox();
            this.CommLabel = new Label();
            this.CommListBox = new ListBox();
            this.VerLabel = new Label();
            this.VerListBox = new ListBox();
            this.OKButton = new Button();
            this.DialogCancelButton = new Button();
            this.SuspendLayout();
            //
            // PowerLabel
            //
            this.PowerLabel.AutoSize = true;
            this.PowerLabel.Location = new Point(12, 15);
            this.PowerLabel.Name = "PowerLabel";
            this.PowerLabel.Size = new Size(60, 15);
            this.PowerLabel.TabIndex = 0;
            this.PowerLabel.Text = "電源";
            //
            // PowerListBox
            //
            this.PowerListBox.IntegralHeight = false;
            this.PowerListBox.Location = new Point(12, 33);
            this.PowerListBox.Name = "PowerListBox";
            this.PowerListBox.SelectionMode = SelectionMode.One;
            this.PowerListBox.Size = new Size(110, 100);
            this.PowerListBox.TabIndex = 1;
            this.PowerListBox.SelectedIndexChanged += this.AttributeListBox_SelectedIndexChanged;
            //
            // MountLabel
            //
            this.MountLabel.AutoSize = true;
            this.MountLabel.Location = new Point(132, 15);
            this.MountLabel.Name = "MountLabel";
            this.MountLabel.Size = new Size(60, 15);
            this.MountLabel.TabIndex = 2;
            this.MountLabel.Text = "設置形態";
            //
            // MountListBox
            //
            this.MountListBox.IntegralHeight = false;
            this.MountListBox.Location = new Point(132, 33);
            this.MountListBox.Name = "MountListBox";
            this.MountListBox.SelectionMode = SelectionMode.One;
            this.MountListBox.Size = new Size(110, 100);
            this.MountListBox.TabIndex = 3;
            this.MountListBox.SelectedIndexChanged += this.AttributeListBox_SelectedIndexChanged;
            //
            // AdiLabel
            //
            this.AdiLabel.AutoSize = true;
            this.AdiLabel.Location = new Point(252, 15);
            this.AdiLabel.Name = "AdiLabel";
            this.AdiLabel.Size = new Size(60, 15);
            this.AdiLabel.TabIndex = 4;
            this.AdiLabel.Text = "ADI";
            //
            // AdiListBox
            //
            this.AdiListBox.IntegralHeight = false;
            this.AdiListBox.Location = new Point(252, 33);
            this.AdiListBox.Name = "AdiListBox";
            this.AdiListBox.SelectionMode = SelectionMode.One;
            this.AdiListBox.Size = new Size(110, 100);
            this.AdiListBox.TabIndex = 5;
            this.AdiListBox.SelectedIndexChanged += this.AttributeListBox_SelectedIndexChanged;
            //
            // CommLabel
            //
            this.CommLabel.AutoSize = true;
            this.CommLabel.Location = new Point(372, 15);
            this.CommLabel.Name = "CommLabel";
            this.CommLabel.Size = new Size(60, 15);
            this.CommLabel.TabIndex = 6;
            this.CommLabel.Text = "通信方式";
            //
            // CommListBox
            //
            this.CommListBox.IntegralHeight = false;
            this.CommListBox.Location = new Point(372, 33);
            this.CommListBox.Name = "CommListBox";
            this.CommListBox.SelectionMode = SelectionMode.One;
            this.CommListBox.Size = new Size(110, 100);
            this.CommListBox.TabIndex = 7;
            this.CommListBox.SelectedIndexChanged += this.AttributeListBox_SelectedIndexChanged;
            //
            // VerLabel
            //
            this.VerLabel.AutoSize = true;
            this.VerLabel.Location = new Point(492, 15);
            this.VerLabel.Name = "VerLabel";
            this.VerLabel.Size = new Size(60, 15);
            this.VerLabel.TabIndex = 8;
            this.VerLabel.Text = "バージョン";
            //
            // VerListBox
            //
            this.VerListBox.IntegralHeight = false;
            this.VerListBox.Location = new Point(492, 33);
            this.VerListBox.Name = "VerListBox";
            this.VerListBox.SelectionMode = SelectionMode.One;
            this.VerListBox.Size = new Size(110, 100);
            this.VerListBox.TabIndex = 9;
            this.VerListBox.SelectedIndexChanged += this.AttributeListBox_SelectedIndexChanged;
            //
            // OKButton
            //
            this.OKButton.Enabled = false;
            this.OKButton.Location = new Point(427, 148);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(85, 27);
            this.OKButton.TabIndex = 10;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.OKButton_Click;
            //
            // DialogCancelButton
            //
            this.DialogCancelButton.Location = new Point(517, 148);
            this.DialogCancelButton.Name = "DialogCancelButton";
            this.DialogCancelButton.Size = new Size(85, 27);
            this.DialogCancelButton.TabIndex = 11;
            this.DialogCancelButton.Text = "キャンセル";
            this.DialogCancelButton.UseVisualStyleBackColor = true;
            this.DialogCancelButton.Click += this.CancelButton_Click;
            //
            // ProductAttributeSelectWindow
            //
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(614, 187);
            this.Controls.Add(this.PowerLabel);
            this.Controls.Add(this.PowerListBox);
            this.Controls.Add(this.MountLabel);
            this.Controls.Add(this.MountListBox);
            this.Controls.Add(this.AdiLabel);
            this.Controls.Add(this.AdiListBox);
            this.Controls.Add(this.CommLabel);
            this.Controls.Add(this.CommListBox);
            this.Controls.Add(this.VerLabel);
            this.Controls.Add(this.VerListBox);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.DialogCancelButton);
            this.CancelButton = this.DialogCancelButton;
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProductAttributeSelectWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "仕様選択";
            this.Load += this.ProductAttributeSelectWindow_Load;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Label PowerLabel;
        private ListBox PowerListBox;
        private Label MountLabel;
        private ListBox MountListBox;
        private Label AdiLabel;
        private ListBox AdiListBox;
        private Label CommLabel;
        private ListBox CommListBox;
        private Label VerLabel;
        private ListBox VerListBox;
        private Button OKButton;
        private Button DialogCancelButton;
    }
}
