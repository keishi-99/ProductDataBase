namespace ProductDatabase.Other {
    partial class ServiceForm {
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
            this.panelCategory1 = new Panel();
            this.CategoryListBox1 = new ListBox();
            this.panelCategory2 = new Panel();
            this.CategoryListBox2 = new ListBox();
            this.panelCategory3 = new Panel();
            this.CategoryListBox3 = new ListBox();
            this.RegisterButton = new Button();
            this.panelCategory1.SuspendLayout();
            this.panelCategory2.SuspendLayout();
            this.panelCategory3.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelCategory1
            // 
            this.panelCategory1.Controls.Add(this.CategoryListBox1);
            this.panelCategory1.Location = new Point(31, 12);
            this.panelCategory1.Name = "panelCategory1";
            this.panelCategory1.Size = new Size(60, 259);
            this.panelCategory1.TabIndex = 604;
            // 
            // CategoryListBox1
            // 
            this.CategoryListBox1.Dock = DockStyle.Fill;
            this.CategoryListBox1.FormattingEnabled = true;
            this.CategoryListBox1.HorizontalScrollbar = true;
            this.CategoryListBox1.ItemHeight = 15;
            this.CategoryListBox1.Location = new Point(0, 0);
            this.CategoryListBox1.Name = "CategoryListBox1";
            this.CategoryListBox1.Size = new Size(60, 259);
            this.CategoryListBox1.TabIndex = 5;
            this.CategoryListBox1.SelectedIndexChanged += this.CategoryListBox1_SelectedIndexChanged;
            // 
            // panelCategory2
            // 
            this.panelCategory2.Controls.Add(this.CategoryListBox2);
            this.panelCategory2.Location = new Point(142, 12);
            this.panelCategory2.Name = "panelCategory2";
            this.panelCategory2.Size = new Size(210, 259);
            this.panelCategory2.TabIndex = 605;
            // 
            // CategoryListBox2
            // 
            this.CategoryListBox2.Dock = DockStyle.Fill;
            this.CategoryListBox2.FormattingEnabled = true;
            this.CategoryListBox2.HorizontalScrollbar = true;
            this.CategoryListBox2.ItemHeight = 15;
            this.CategoryListBox2.Location = new Point(0, 0);
            this.CategoryListBox2.Name = "CategoryListBox2";
            this.CategoryListBox2.Size = new Size(210, 259);
            this.CategoryListBox2.TabIndex = 6;
            this.CategoryListBox2.SelectedIndexChanged += this.CategoryListBox2_SelectedIndexChanged;
            // 
            // panelCategory3
            // 
            this.panelCategory3.Controls.Add(this.CategoryListBox3);
            this.panelCategory3.Location = new Point(403, 12);
            this.panelCategory3.Name = "panelCategory3";
            this.panelCategory3.Size = new Size(350, 259);
            this.panelCategory3.TabIndex = 606;
            // 
            // CategoryListBox3
            // 
            this.CategoryListBox3.Dock = DockStyle.Fill;
            this.CategoryListBox3.FormattingEnabled = true;
            this.CategoryListBox3.HorizontalScrollbar = true;
            this.CategoryListBox3.ItemHeight = 15;
            this.CategoryListBox3.Location = new Point(0, 0);
            this.CategoryListBox3.Name = "CategoryListBox3";
            this.CategoryListBox3.Size = new Size(350, 259);
            this.CategoryListBox3.TabIndex = 7;
            this.CategoryListBox3.SelectedIndexChanged += this.CategoryListBox3_SelectedIndexChanged;
            // 
            // RegisterButton
            // 
            this.RegisterButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.RegisterButton.Font = new Font("Meiryo UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 128);
            this.RegisterButton.Location = new Point(344, 295);
            this.RegisterButton.Name = "RegisterButton";
            this.RegisterButton.Size = new Size(96, 54);
            this.RegisterButton.TabIndex = 607;
            this.RegisterButton.Text = "OK";
            this.RegisterButton.UseVisualStyleBackColor = true;
            this.RegisterButton.Click += this.RegisterButton_Click;
            // 
            // ServiceForm
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(784, 361);
            this.Controls.Add(this.RegisterButton);
            this.Controls.Add(this.panelCategory1);
            this.Controls.Add(this.panelCategory2);
            this.Controls.Add(this.panelCategory3);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ServiceForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Load += this.ServiceForm_Load;
            this.panelCategory1.ResumeLayout(false);
            this.panelCategory2.ResumeLayout(false);
            this.panelCategory3.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private Panel panelCategory1;
        private ListBox CategoryListBox1;
        private Panel panelCategory2;
        private ListBox CategoryListBox2;
        private Panel panelCategory3;
        private ListBox CategoryListBox3;
        private Button RegisterButton;
    }
}