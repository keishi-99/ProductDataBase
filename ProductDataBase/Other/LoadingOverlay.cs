using System.Drawing.Drawing2D;

namespace ProductDatabase.Other {
    // 親フォームに重ねて表示する半透明ローディングオーバーレイ
    // using 構文で使用し、Dispose 時に親フォームから自動的に除去される
    internal sealed class LoadingOverlay : IDisposable {
        private readonly Control _parent;
        private readonly OverlayPanel _panel;

        public LoadingOverlay(Control parent) {
            _parent = parent;
            _panel = new OverlayPanel {
                Bounds = parent.ClientRectangle
            };
            parent.Controls.Add(_panel);
            _panel.BringToFront();
        }

        public void Dispose() {
            _parent.Controls.Remove(_panel);
            _panel.Dispose();
        }

        // 半透明オーバーレイと円弧スピナーを描画するパネル
        private sealed class OverlayPanel : Panel {
            private readonly System.Windows.Forms.Timer _timer;
            private float _angle;

            public OverlayPanel() {
                SetStyle(
                    ControlStyles.SupportsTransparentBackColor |
                    ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.UserPaint |
                    ControlStyles.OptimizedDoubleBuffer,
                    true);
                BackColor = Color.Transparent;

                // 約60fps でスピナーを回転
                _timer = new System.Windows.Forms.Timer { Interval = 16 };
                _timer.Tick += (_, _) => {
                    _angle = (_angle + 6f) % 360f;
                    Invalidate();
                };
                _timer.Start();
            }

            protected override void OnPaint(PaintEventArgs e) {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                // 半透明グレーオーバーレイ
                using var overlayBrush = new SolidBrush(Color.FromArgb(130, 50, 50, 50));
                g.FillRectangle(overlayBrush, ClientRectangle);

                // 中央の白いカード（角丸）
                const int cardW = 130;
                const int cardH = 110;
                int cardX = (Width - cardW) / 2;
                int cardY = (Height - cardH) / 2;
                var cardRect = new RectangleF(cardX, cardY, cardW, cardH);

                using (var cardBrush = new SolidBrush(Color.White))
                using (var cardPath = CreateRoundedRect(cardRect, 10f)) {
                    g.FillPath(cardBrush, cardPath);
                }

                // スピナー（薄い背景円 + 回転する円弧）
                const int spinnerSize = 48;
                int spinnerX = cardX + (cardW - spinnerSize) / 2;
                int spinnerY = cardY + 18;
                var spinnerRect = new RectangleF(spinnerX, spinnerY, spinnerSize, spinnerSize);

                using (var bgPen = new Pen(Color.FromArgb(220, 220, 220), 5f)) {
                    bgPen.StartCap = LineCap.Round;
                    bgPen.EndCap = LineCap.Round;
                    g.DrawEllipse(bgPen, spinnerRect);
                }

                using (var arcPen = new Pen(Color.SteelBlue, 5f)) {
                    arcPen.StartCap = LineCap.Round;
                    arcPen.EndCap = LineCap.Round;
                    g.DrawArc(arcPen, spinnerRect, _angle, 90f);
                }

                // 「処理中...」テキスト
                using var font = new Font("Meiryo UI", 9f);
                using var textBrush = new SolidBrush(Color.FromArgb(80, 80, 80));
                const string text = "処理中...";
                var textSize = g.MeasureString(text, font);
                g.DrawString(
                    text, font, textBrush,
                    cardX + (cardW - textSize.Width) / 2f,
                    cardY + cardH - 26f);
            }

            // 角丸四角形のパスを生成する
            private static GraphicsPath CreateRoundedRect(RectangleF bounds, float radius) {
                var path = new GraphicsPath();
                var d = radius * 2;
                path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }

            protected override void Dispose(bool disposing) {
                if (disposing) _timer.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
