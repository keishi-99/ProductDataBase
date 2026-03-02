namespace ProductDatabase {
    partial class SubstrateChange1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.SubstrateChangeDataGridView = new DataGridView();
            this.label1 = new Label();
            this.FilterStringTextBox = new TextBox();
            this.OKButton = new Button();
            ((System.ComponentModel.ISupportInitialize)this.SubstrateChangeDataGridView).BeginInit();
            this.SuspendLayout();
            // 
            // SubstrateChangeDataGridView
            // 
            this.SubstrateChangeDataGridView.AllowUserToAddRows = false;
            this.SubstrateChangeDataGridView.AllowUserToDeleteRows = false;
            this.SubstrateChangeDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.SubstrateChangeDataGridView.Location = new Point(0, 0);
            this.SubstrateChangeDataGridView.MultiSelect = false;
            this.SubstrateChangeDataGridView.Name = "SubstrateChangeDataGridView";
            this.SubstrateChangeDataGridView.ReadOnly = true;
            this.SubstrateChangeDataGridView.Size = new Size(1184, 300);
            this.SubstrateChangeDataGridView.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new Point(10, 315);
            this.label1.Name = "label1";
            this.label1.Size = new Size(98, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "製造番号フィルター";
            // 
            // FilterStringTextBox
            // 
            this.FilterStringTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.FilterStringTextBox.Location = new Point(12, 330);
            this.FilterStringTextBox.MaxLength = 50;
            this.FilterStringTextBox.Name = "FilterStringTextBox";
            this.FilterStringTextBox.Size = new Size(100, 23);
            this.FilterStringTextBox.TabIndex = 3;
            // 
            // OKButton
            // 
            this.OKButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.OKButton.Location = new Point(118, 328);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(75, 25);
            this.OKButton.TabIndex = 4;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.OKButton_Click;
            // 
            // SubstrateChange1
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(1184, 361);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.FilterStringTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SubstrateChangeDataGridView);
            this.Font = new Font("Meiryo UI", 9F);
            this.MinimizeBox = false;
            this.Name = "SubstrateChange1";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "変更候補";
            this.Load += this.SubstrateChange1_Load;
            ((System.ComponentModel.ISupportInitialize)this.SubstrateChangeDataGridView).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private DataGridView SubstrateChangeDataGridView;
        private Label label1;
        private TextBox FilterStringTextBox;
        private Button OKButton;
    }
}