namespace ProductDatabase {
    partial class PrintSettingsWindow {
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
            this.OKButton = new Button();
            this.CloseButton = new Button();
            this.PrintPropertyGrid = new PropertyGrid();
            this.SuspendLayout();
            // 
            // OKButton
            // 
            this.OKButton.Location = new Point(471, 581);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new Size(75, 25);
            this.OKButton.TabIndex = 32;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += this.OKButton_Click;
            // 
            // CloseButton
            // 
            this.CloseButton.Location = new Point(552, 581);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new Size(75, 25);
            this.CloseButton.TabIndex = 33;
            this.CloseButton.Text = "キャンセル";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += this.CloseButton_Click;
            // 
            // PrintPropertyGrid
            // 
            this.PrintPropertyGrid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.PrintPropertyGrid.Location = new Point(12, 12);
            this.PrintPropertyGrid.Name = "PrintPropertyGrid";
            this.PrintPropertyGrid.PropertySort = PropertySort.Categorized;
            this.PrintPropertyGrid.Size = new Size(630, 557);
            this.PrintPropertyGrid.TabIndex = 34;
            this.PrintPropertyGrid.ToolbarVisible = false;
            // 
            // PrintSettingsWindow
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new Size(654, 618);
            this.Controls.Add(this.PrintPropertyGrid);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.CloseButton);
            this.Font = new Font("Meiryo UI", 9F);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintSettingsWindow";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "印刷設定";
            this.Load += this.ProductBarcodePrintSetting_Load;
            this.ResumeLayout(false);
        }

        #endregion

        private Button OKButton;
        private Button CloseButton;
        private PropertyGrid PrintPropertyGrid;
    }
}