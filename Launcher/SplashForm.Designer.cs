namespace Launcher {
    partial class SplashForm {
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(SplashForm));
            this.SuspendLayout();
            // 
            // SplashForm
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.Green;
            this.BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            this.BackgroundImageLayout = ImageLayout.Zoom;
            this.ClientSize = new Size(400, 400);
            this.ControlBox = false;
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Icon = (Icon)resources.GetObject("$this.Icon");
            this.Name = "SplashForm";
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "SplashForm";
            this.TopMost = true;
            this.TransparencyKey = Color.Green;
            this.ResumeLayout(false);
        }

        #endregion
    }
}