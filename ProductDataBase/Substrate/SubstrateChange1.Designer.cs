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
            SubstrateChangeDataGridView = new DataGridView();
            label1 = new Label();
            FilterStringTextBox = new TextBox();
            OKButton = new Button();
            ((System.ComponentModel.ISupportInitialize)SubstrateChangeDataGridView).BeginInit();
            SuspendLayout();
            // 
            // SubstrateChangeDataGridView
            // 
            SubstrateChangeDataGridView.AllowUserToAddRows = false;
            SubstrateChangeDataGridView.AllowUserToDeleteRows = false;
            SubstrateChangeDataGridView.AllowUserToResizeRows = false;
            SubstrateChangeDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            SubstrateChangeDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            SubstrateChangeDataGridView.Location = new Point(0, 0);
            SubstrateChangeDataGridView.Name = "SubstrateChangeDataGridView";
            SubstrateChangeDataGridView.ReadOnly = true;
            SubstrateChangeDataGridView.RowHeadersVisible = false;
            SubstrateChangeDataGridView.RowTemplate.Height = 21;
            SubstrateChangeDataGridView.Size = new Size(1184, 300);
            SubstrateChangeDataGridView.TabIndex = 1;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label1.AutoSize = true;
            label1.Location = new Point(10, 315);
            label1.Name = "label1";
            label1.Size = new Size(98, 15);
            label1.TabIndex = 2;
            label1.Text = "製造番号フィルター";
            // 
            // FilterStringTextBox
            // 
            FilterStringTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            FilterStringTextBox.Location = new Point(12, 330);
            FilterStringTextBox.MaxLength = 50;
            FilterStringTextBox.Name = "FilterStringTextBox";
            FilterStringTextBox.Size = new Size(100, 23);
            FilterStringTextBox.TabIndex = 3;
            // 
            // OKButton
            // 
            OKButton.Location = new Point(118, 329);
            OKButton.Name = "OKButton";
            OKButton.Size = new Size(75, 23);
            OKButton.TabIndex = 4;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            // 
            // SubstrateChange1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1184, 361);
            Controls.Add(OKButton);
            Controls.Add(FilterStringTextBox);
            Controls.Add(label1);
            Controls.Add(SubstrateChangeDataGridView);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "SubstrateChange1";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "変更候補";
            ((System.ComponentModel.ISupportInitialize)SubstrateChangeDataGridView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView SubstrateChangeDataGridView;
        private Label label1;
        private TextBox FilterStringTextBox;
        private Button OKButton;
    }
}