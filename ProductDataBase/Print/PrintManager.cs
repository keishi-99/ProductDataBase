using System.ComponentModel;
using System.Drawing.Printing;
using System.Text.Json;
using System.Text.Json.Serialization;
using ZXing;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDatabase.Print {
    // シリアル印刷管理クラス
    public static class PrintManager {

        // 状態保持用プロパティ（外部からは読み取り専用）
        public static ProductInformation ProductInfo { get; private set; } = default!;
        public static DocumentPrintSettings ProductPrintSettings { get; private set; } = new();

        // 各種設定へのアクセスプロパティ（nullチェック済）
        public static LabelPrintSettings LabelPrintSettings => ProductPrintSettings.LabelPrintSettings ?? throw new InvalidOperationException("LabelPageSettings が null です。");
        public static BarcodePrintSettings BarcodePrintSettings => ProductPrintSettings.BarcodePrintSettings ?? throw new InvalidOperationException("BarcodePageSettings が null です。");
        private static DocumentPrintSettings PrintSettings => ProductPrintSettings ?? throw new InvalidOperationException("ProductPrintSettings が null です。");

        // 印刷状態を保持するプロパティ
        public static int PrintCount { get; private set; }
        public static int CopiesRemainingPerSerial { get; private set; }
        public static int PageCount { get; private set; }

        public static int PrintType => ProductInfo?.PrintType ?? throw new InvalidOperationException("PrintManager が初期化されていません。");

        private static List<string> s_serialList = [];

        public static bool IsUnderlinePrint => ProductInfo.IsUnderlinePrint;

        // 4桁以上の型式番号の下4桁を取得するプロパティ
        public static string Last4ProductModel =>
            ProductInfo.ProductModel.Length >= 4
                ? ProductInfo.ProductModel[^4..]
                : string.Empty;

        public static void Initialize(ProductInformation productInfo, DocumentPrintSettings productPrintSettings, List<string> serialList) {
            ProductInfo = productInfo;
            ProductPrintSettings = productPrintSettings ?? throw new ArgumentNullException(nameof(productPrintSettings));

            s_serialList = serialList;
            PageCount = 1;
            PrintCount = 0;
        }

        // ミリメートルをピクセルに変換するヘルパーメソッド
        private static float ConvertMmToPixel(double mm, float dpi) {
            const float MmPerInch = 25.4f;
            return (float)(mm / MmPerInch * dpi);
        }

        public static bool PrintSerialCommon(PrintPageEventArgs e, bool isPreview, int startLine, string serialType) {
            try {
                if (e.Graphics is null) { throw new InvalidOperationException("Graphics オブジェクトを取得できません。"); }

                e.Graphics.PageUnit = GraphicsUnit.Pixel;

                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;

                PrintSettingsBase printSettings = serialType switch {
                    "Label" => PrintSettings.LabelPrintSettings!,
                    "Barcode" => PrintSettings.BarcodePrintSettings!,
                    "Substrate" => PrintSettings.LabelPrintSettings!,
                    _ => throw new ArgumentException($"不明なシリアルタイプ: {serialType}")
                };

                var labelCountX = printSettings.LabelsPerColumn;
                var labelCountY = printSettings.LabelsPerRow;
                var labelWidthPx = ConvertMmToPixel(printSettings.LabelWidth, dpiX);
                var labelHeightPx = ConvertMmToPixel(printSettings.LabelHeight, dpiY);
                var marginXPx = ConvertMmToPixel(printSettings.MarginX, dpiX);
                var marginYPx = ConvertMmToPixel(printSettings.MarginY, dpiY);
                var intervalXPx = ConvertMmToPixel(printSettings.IntervalX, dpiX);
                var intervalYPx = ConvertMmToPixel(printSettings.IntervalY, dpiY);
                var headerPositionXPx = ConvertMmToPixel(printSettings.HeaderPositionX, dpiX);
                var headerPositionYPx = ConvertMmToPixel(printSettings.HeaderPositionY, dpiY);
                var headerString = ConvertHeaderString(printSettings.HeaderTextFormat);
                var headerFont = printSettings.HeaderFont;

                var copiesPerLabel = printSettings.CopiesPerLabel;
                CopiesRemainingPerSerial = copiesPerLabel;
                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) {
                    throw new Exception("印刷設定が異常です。");
                }

                var (hardMarginX, hardMarginY) = isPreview
                    ? (0, 0)
                    : (e.PageSettings.HardMarginX * dpiX / 100.0f, e.PageSettings.HardMarginY * dpiY / 100.0f);

                //if (marginXPx <= hardMarginX) { throw new Exception($"左余白の値がプリンター左余白より小さい値になっています。\n{marginXPx} <= {hardMarginX}"); }
                //if (marginYPx <= hardMarginY) { throw new Exception($"上余白の値がプリンター上余白より小さい値になっています。\n{marginYPx} <= {hardMarginY}"); }

                if (PageCount >= 2) { startLine = 0; }

                var verticalOffsetPx = PageCount == 1 ? startLine * (intervalYPx + labelHeightPx) : 0;
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, headerPositionXPx, verticalOffsetPx + headerPositionYPx - hardMarginY);

                for (var y = startLine; y < labelCountY; y++) {
                    for (var x = 0; x < labelCountX; x++) {
                        var posX = marginXPx - hardMarginX + (x * (intervalXPx + labelWidthPx));
                        var posY = marginYPx - hardMarginY + (y * (intervalYPx + labelHeightPx));

                        var isLastCopy = CopiesRemainingPerSerial == 1;

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = IsUnderlinePrint && isLastCopy;

                        var printText = PrintType switch {
                            9 when isLastCopy => Last4ProductModel,
                            _ => s_serialList[PrintCount]
                        };

                        using var labelImage = MakeLabelImage(printText, serialType, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY, isPreview);

                        e.Graphics.DrawImage(labelImage, posX, posY, labelWidthPx, labelHeightPx);

                        CopiesRemainingPerSerial--;

                        if (CopiesRemainingPerSerial <= 0) {
                            CopiesRemainingPerSerial = copiesPerLabel;
                            PrintCount++;

                            if (s_serialList.Count <= PrintCount) {
                                // 最後の行にマークを描画
                                DrawFinalRowMark(e.Graphics, y + 1, 0, posY, 0, labelHeightPx, headerFont);
                                return false;
                            }
                        }
                    }
                }

                if (s_serialList.Count > PrintCount) {
                    PageCount++;
                    return true;
                }

                return false;
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private static Bitmap MakeLabelImage(string text, string serialType, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY, bool isPreview) {

            // ビットマップのサイズをピクセル単位で計算
            var widthPx = (int)Math.Round(labelWidthPx);
            var heightPx = (int)Math.Round(labelHeightPx);

            var labelImage = new Bitmap(widthPx, heightPx);
            labelImage.SetResolution(dpiX, dpiY);

            // usingステートメントでGraphicsオブジェクトを確実に破棄
            using (var g = Graphics.FromImage(labelImage)) {
                // すべての描画操作をピクセル単位で行うように設定
                g.PageUnit = GraphicsUnit.Pixel;

                // 高品質な描画設定
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                // 印刷タイプに応じた描画処理
                switch (serialType) {
                    case "Label":
                        DrawLabel(g, text, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                    case "Barcode":
                        DrawBarcode(g, text, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                    case "Substrate":
                        DrawLabel(g, text, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                }

                // プレビュー用の枠を描画
                if (isPreview) {
                    // 0.1mmの黒いペンで枠線を描画
                    using var p = new Pen(Color.Black, ConvertMmToPixel(0.1, dpiX));
                    g.DrawRectangle(p, 0, 0, widthPx - 1, heightPx - 1); // 枠線がはみ出さないように-1
                }
            }

            return labelImage;
        }
        private static void DrawLabel(Graphics g, string text, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            if (PrintSettings.LabelPrintSettings is null) { return; }
            // フォントサイズはポイント単位でそのまま使用
            var fontName = PrintSettings.LabelPrintSettings.TextFont.Name;
            var fontSize = PrintSettings.LabelPrintSettings.TextFont.SizeInPoints;
            var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;

            using var textFont = new Font(fontName, fontSize, style);
            // テキストの配置設定
            using var sf = new StringFormat {
                Alignment = PrintSettings.LabelPrintSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = PrintSettings.LabelPrintSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
            };
            // 描画領域をピクセル単位で計算
            // TextPositionX, TextPositionY はミリメートル単位なのでピクセルに変換
            var textPosX = PrintSettings.LabelPrintSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(PrintSettings.LabelPrintSettings.TextPositionX, dpiX);
            var textPosY = PrintSettings.LabelPrintSettings.AlignTextCenterY ? 0f : ConvertMmToPixel(PrintSettings.LabelPrintSettings.TextPositionY, dpiY);

            var layoutRect = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

            g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);
        }
        private static void DrawBarcode(Graphics g, string text, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            if (PrintSettings.BarcodePrintSettings is null) { return; }
            // --- テキストの描画 ---
            // フォントサイズはポイント単位でそのまま使用
            var fontName = PrintSettings.BarcodePrintSettings.TextFont.Name;
            var fontSize = PrintSettings.BarcodePrintSettings.TextFont.SizeInPoints;

            using (var textFont = new Font(fontName, fontSize, FontStyle.Regular)) {
                // テキストの配置設定
                using var sf = new StringFormat {
                    Alignment = PrintSettings.BarcodePrintSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near
                };

                // TextPositionX, TextPositionY はミリメートル単位なのでピクセルに変換
                var textPosX = PrintSettings.BarcodePrintSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(PrintSettings.BarcodePrintSettings.TextPositionX, dpiX);
                var textPosY = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.TextPositionY, dpiY);

                var layoutRectString = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

                g.DrawString(text, textFont, Brushes.Black, layoutRectString, sf);
            }

            // --- バーコードの描画 ---
            // ZXingはピクセル単位で画像を生成するため、mmから pixelへの変換が必要
            var barcodePixelWidth = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeWidth, dpiX);
            var barcodePixelHeight = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeHeight, dpiY);

            // ZXing用に int型に変換
            var qrWidthPx = (int)Math.Round(barcodePixelWidth);
            var qrHeightPx = (int)Math.Round(barcodePixelHeight);

            var writer = new BarcodeWriter<Bitmap> {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions {
                    Width = qrWidthPx,
                    Height = qrHeightPx,
                    PureBarcode = true // テキストを含まないバーコードのみを生成
                },
                Renderer = new BitmapRenderer()
            };

            using var barcodeBitmap = writer.Write(text);
            // BarcodePositionX, BarcodePositionY はミリメートル単位なのでピクセルに変換
            var barcodePosX = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodePositionX, dpiX);
            var barcodePosY = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodePositionY, dpiY);
            var barcodeWidthPx = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeWidth, dpiX);
            var barcodeHeightPx = ConvertMmToPixel(PrintSettings.BarcodePrintSettings.BarcodeHeight, dpiY);

            // X座標を中央に調整 (ピクセル単位で計算)
            if (PrintSettings.BarcodePrintSettings.AlignBarcodeCenterX) {
                barcodePosX = (labelWidthPx / 2f) - (barcodeWidthPx / 2f);
            }

            var layoutRectBarcode = new RectangleF(barcodePosX, barcodePosY, barcodeWidthPx, barcodeHeightPx);
            // g.DrawImageでミリメートル単位の座標とサイズを指定して描画
            g.DrawImage(barcodeBitmap, layoutRectBarcode);
        }
        private static string ConvertHeaderString(string s) {
            if (ProductInfo is null) { throw new Exception("ProductInfoが nullです。"); }
            s = s.Replace("%P", ProductInfo.ProductName)
                 .Replace("%T", ProductInfo.ProductModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", ProductInfo.ProductNumber)
                 .Replace("%O", ProductInfo.OrderNumber)
                 .Replace("%N", ProductInfo.Quantity.ToString())
                 .Replace("%U", ProductInfo.Person);
            return s;
        }
        private static void DrawFinalRowMark(Graphics graphics, int rowNumber, float posX, float posY, float width, float height, Font font) {
            var sf = new StringFormat {
                Alignment = StringAlignment.Near,
                LineAlignment = StringAlignment.Center
            };
            var layoutRect = new RectangleF(posX, posY, width, height);
            graphics.DrawString(rowNumber.ToString(), font, Brushes.Black, layoutRect, sf);
        }
    }
    public class PrintOptions {

        public class DocumentPrintSettings {
            public LabelPrintSettings? LabelPrintSettings { get; set; }
            public BarcodePrintSettings? BarcodePrintSettings { get; set; }

            public DocumentPrintSettings() {
                LabelPrintSettings = new LabelPrintSettings();
                BarcodePrintSettings = new BarcodePrintSettings();
            }

            public void SetSettingsType(bool isLabelPrint, bool isBarcodePrint) {
                LabelPrintSettings = isLabelPrint ? (LabelPrintSettings ?? new LabelPrintSettings()) : null;
                BarcodePrintSettings = isBarcodePrint ? (BarcodePrintSettings ?? new BarcodePrintSettings()) : null;
            }
        }

        public abstract class PrintSettingsBase {
            [Category("\t用紙設定"), DisplayName("ラベル幅 (mm)")]
            public double LabelWidth { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル高さ (mm)")]
            public double LabelHeight { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (行)")]
            public int LabelsPerRow { get; set; }
            [Category("\t用紙設定"), DisplayName("配置数 (列)")]
            public int LabelsPerColumn { get; set; }
            [Category("\t用紙設定"), DisplayName("余白左 (mm)")]
            public double MarginX { get; set; }
            [Category("\t用紙設定"), DisplayName("余白上 (mm)")]
            public double MarginY { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 横 (mm)")]
            public double IntervalX { get; set; }
            [Category("\t用紙設定"), DisplayName("ラベル間隔 縦 (mm)")]
            public double IntervalY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 上 (mm)")]
            public double HeaderPositionX { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー開始位置 左 (mm)")]
            public double HeaderPositionY { get; set; }
            [Category("\t用紙設定"), DisplayName("ヘッダー フォーマット"), Description("%D(印刷日)  %M(製番)  %O(注番)  %T(型式)  %P(製品名)  %N(台数)  %U(担当者)")]
            public string HeaderTextFormat { get; set; } = string.Empty;

            [Category("\t用紙設定"), DisplayName("ヘッダー フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;

            [Category("印字設定"), DisplayName("発行枚数")]
            public int CopiesPerLabel { get; set; }
            [Category("印字設定"), DisplayName("印字位置 左 (mm)")]
            public double TextPositionX { get; set; }
            [Category("印字設定"), DisplayName("印字位置 上 (mm)")]
            public double TextPositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignTextCenterX { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (縦)")]
            public bool AlignTextCenterY { get; set; }
            [Category("印字設定"), DisplayName("印字フォーマット"), Description("%T(型式)  %R(リビジョン)  %S(シリアル)  %Y(製造年)  %M(製造月(1桁))  %MM(製造月(2桁))")]
            public string TextFormat { get; set; } = string.Empty;

            [Category("印字設定"), DisplayName("印字フォント"), JsonConverter(typeof(FontJsonConverter))]
            public Font TextFont { get; set; } = SystemFonts.DefaultFont;

        }

        public class LabelPrintSettings : PrintSettingsBase {
            // ラベル固有の設定（今後増える場合にここに追加）
        }

        public class BarcodePrintSettings : PrintSettingsBase {
            [Category("印字設定"), DisplayName("バーコード幅 (mm)")]
            public double BarcodeWidth { get; set; }
            [Category("印字設定"), DisplayName("バーコード高さ (mm)")]
            public double BarcodeHeight { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 左 (mm)")]
            public double BarcodePositionX { get; set; }
            [Category("印字設定"), DisplayName("バーコード位置 上 (mm)")]
            public double BarcodePositionY { get; set; }
            [Category("印字設定"), DisplayName("中心に配置 (横)")]
            public bool AlignBarcodeCenterX { get; set; }
        }

        public class FontJsonConverter : JsonConverter<Font> {
            public override Font Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) {
                var fontString = reader.GetString()!;
                return (Font)TypeDescriptor.GetConverter(typeof(Font)).ConvertFromString(fontString)!;
            }

            public override void Write(Utf8JsonWriter writer, Font value, JsonSerializerOptions options) {
                var fontString = TypeDescriptor.GetConverter(typeof(Font)).ConvertToString(value)!;
                writer.WriteStringValue(fontString);
            }
        }
    }
}
