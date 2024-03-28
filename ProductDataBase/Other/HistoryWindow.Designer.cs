namespace ProductDatabase {
    partial class HistoryWindow {
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
            DataBaseDataGridView = new DataGridView();
            CategoryComboBox = new ComboBox();
            FilterStringTextBox = new TextBox();
            ExportCsvButton = new Button();
            ((System.ComponentModel.ISupportInitialize)DataBaseDataGridView).BeginInit();
            SuspendLayout();
            // 
            // DataBaseDataGridView
            // 
            DataBaseDataGridView.AllowUserToAddRows = false;
            DataBaseDataGridView.AllowUserToDeleteRows = false;
            DataBaseDataGridView.AllowUserToResizeRows = false;
            DataBaseDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            DataBaseDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataBaseDataGridView.Location = new Point(0, 0);
            DataBaseDataGridView.Name = "DataBaseDataGridView";
            DataBaseDataGridView.ReadOnly = true;
            DataBaseDataGridView.RowHeadersVisible = false;
            DataBaseDataGridView.RowTemplate.Height = 21;
            DataBaseDataGridView.Size = new Size(1184, 476);
            DataBaseDataGridView.TabIndex = 0;
            // 
            // CategoryComboBox
            // 
            CategoryComboBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            CategoryComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            CategoryComboBox.FormattingEnabled = true;
            CategoryComboBox.Location = new Point(12, 482);
            CategoryComboBox.Name = "CategoryComboBox";
            CategoryComboBox.Size = new Size(121, 23);
            CategoryComboBox.TabIndex = 1;
            // 
            // FilterStringTextBox
            // 
            FilterStringTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            FilterStringTextBox.Location = new Point(139, 483);
            FilterStringTextBox.MaxLength = 50;
            FilterStringTextBox.Name = "FilterStringTextBox";
            FilterStringTextBox.Size = new Size(100, 23);
            FilterStringTextBox.TabIndex = 2;
            FilterStringTextBox.TextChanged += FilterStringTextBox_TextChanged;
            // 
            // ExportCsvButton
            // 
            ExportCsvButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            ExportCsvButton.Location = new Point(1097, 483);
            ExportCsvButton.Name = "ExportCsvButton";
            ExportCsvButton.Size = new Size(75, 23);
            ExportCsvButton.TabIndex = 3;
            ExportCsvButton.Text = "CSV出力";
            ExportCsvButton.UseVisualStyleBackColor = true;
            ExportCsvButton.Click += ExportCsvButton_Click;
            // 
            // HistoryWindow
            // 
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1184, 511);
            Controls.Add(ExportCsvButton);
            Controls.Add(FilterStringTextBox);
            Controls.Add(CategoryComboBox);
            Controls.Add(DataBaseDataGridView);
            Font = new Font("Meiryo UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MinimizeBox = false;
            Name = "HistoryWindow";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "データベース";
            Load += HistoryWindow_Load;
            ((System.ComponentModel.ISupportInitialize)DataBaseDataGridView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView DataBaseDataGridView;
        private ComboBox CategoryComboBox;
        private TextBox FilterStringTextBox;
        private Button ExportCsvButton;
    }
}