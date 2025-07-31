using System.ComponentModel;
using System.Drawing.Printing;
using System.Text.Json;
using System.Text.Json.Serialization;
using ZXing;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;
using static ProductDatabase.Other.PrintOptions;

namespace ProductDatabase.Other {
    // シリアル印刷管理クラス
    public static class PrintManager {

        // 状態保持用プロパティ（外部からは読み取り専用）
        public static ProductInformation ProductInfo { get; private set; } = default!;
        public static DocumentPrintSettings  ProductPrintSettings { get; private set; } = new();

        // 各種設定へのアクセスプロパティ（nullチェック済）
        public static PrintPageSettings LabelPageSettings =>
            ProductPrintSettings.LabelPageSettings ?? throw new InvalidOperationException("LabelPageSettings が null です。");

        public static PrintLayoutSettings LabelLayoutSettings =>
            ProductPrintSettings.LabelLayoutSettings ?? throw new InvalidOperationException("LabelLayoutSettings が null です。");

        public static PrintPageSettings BarcodePageSettings =>
            ProductPrintSettings.BarcodePageSettings ?? throw new InvalidOperationException("BarcodePageSettings が null です。");

        public static PrintLayoutSettings BarcodeLayoutSettings =>
            ProductPrintSettings.BarcodeLayoutSettings ?? throw new InvalidOperationException("BarcodeLayoutSettings が null です。");

        // 印刷状態を保持するプロパティ
        public static int RemainingLabelCount { get; private set; }
        public static int CopiesRemainingPerSerial { get; private set; }
        public static int SerialNumber { get; private set; }
        public static int PageCount { get; private set; }
        public static int PrintType => ProductInfo?.PrintType ?? throw new InvalidOperationException("PrintManager が初期化されていません。");

        public static bool IsUnderlinePrint => ProductInfo.IsUnderlinePrint;

        // 4桁以上の型式番号の下4桁を取得するプロパティ
        public static string Last4ProductModel =>
            ProductInfo.ProductModel.Length >= 4
                ? ProductInfo.ProductModel[^4..]
                : string.Empty;

        public static void Initialize(ProductInformation productInfo, DocumentPrintSettings  productPrintSettings) {
            ProductInfo = productInfo;
            ProductPrintSettings = productPrintSettings ?? throw new ArgumentNullException(nameof(productPrintSettings));

            PageCount = 1;
            RemainingLabelCount = productInfo.Quantity;
            SerialNumber = productInfo.SerialFirstNumber;
        }

        // ミリメートルをピクセルに変換するヘルパーメソッド
        private static float ConvertMmToPixel(double mm, float dpi) {
            const float MmPerInch = 25.4f;
            return (float)(mm / MmPerInch * dpi);
        }

        public static bool PrintSerial(PrintPageEventArgs e, bool isPreview, string serialType, int startLine) {
            try {
                if (e.Graphics == null) { throw new Exception("e.Graphicsがnullです。"); }

                // Graphicsオブジェクトの描画単位をピクセルに設定
                e.Graphics.PageUnit = GraphicsUnit.Pixel;

                // プリンターのDPIを取得
                var dpiX = e.Graphics.DpiX;
                var dpiY = e.Graphics.DpiY;

                var pageSettings = serialType switch {
                    "Label" => LabelPageSettings,
                    "Barcode" => BarcodePageSettings,
                    _ => throw new ArgumentException($"不明なシリアルタイプ: {serialType}")
                };

                var labelCountX = pageSettings.LabelsPerColumn;
                var labelCountY = pageSettings.LabelsPerRow;
                var labelWidthPx = ConvertMmToPixel(pageSettings.LabelWidth, dpiX);
                var labelHeightPx = ConvertMmToPixel(pageSettings.LabelHeight, dpiY);
                var marginXPx = ConvertMmToPixel(pageSettings.MarginX, dpiX);
                var marginYPx = ConvertMmToPixel(pageSettings.MarginY, dpiY);
                var intervalXPx = ConvertMmToPixel(pageSettings.IntervalX, dpiX);
                var intervalYPx = ConvertMmToPixel(pageSettings.IntervalY, dpiY);
                var headerPositionXPx = ConvertMmToPixel(pageSettings.HeaderPositionX, dpiX);
                var headerPositionYPx = ConvertMmToPixel(pageSettings.HeaderPositionY, dpiY);
                var headerString = ConvertHeaderString(pageSettings.HeaderTextFormat);
                var headerFont = LabelPageSettings.HeaderFont;
                var copiesPerLabel = LabelLayoutSettings.CopiesPerLabel;

                if (labelCountX == 0 || labelCountY == 0 || copiesPerLabel == 0) { throw new Exception("印刷設定が異常です。"); }

                // ハードマージンをpixelに変換
                var hardMarginX = 0f;
                var hardMarginY = 0f;
                if (!isPreview) {
                    (hardMarginX, hardMarginY) = (e.PageSettings.HardMarginX * e.Graphics.DpiX / 100.0f, e.PageSettings.HardMarginY * e.Graphics.DpiY / 100.0f);
                    //(hardMarginX, hardMarginY) = _printerName switch {
                    //    _ => (e.PageSettings.HardMarginX * e.Graphics.DpiX / 100.0f, e.PageSettings.HardMarginY * e.Graphics.DpiY / 100.0f)
                    //};
                }

                if (PageCount == 1) {
                    CopiesRemainingPerSerial = copiesPerLabel;
                }
                if (PageCount >= 2) { startLine = 0; }

                // 最初のページのみオフセットを調整
                var verticalOffsetPx = PageCount == 1 ? startLine * (intervalYPx + labelHeightPx) : 0;
                // ヘッダーの描画
                e.Graphics.DrawString(headerString, headerFont, Brushes.Gray, headerPositionXPx, (float)(verticalOffsetPx + headerPositionYPx - hardMarginY));

                var y = 0;
                for (y = startLine; y < labelCountY; y++) {
                    var x = 0;
                    for (x = 0; x < labelCountX; x++) {
                        // ピクセル単位の座標を使用
                        var posX = marginXPx - hardMarginX + (x * (intervalXPx + labelWidthPx));
                        var posY = marginYPx - hardMarginY + (y * (intervalYPx + labelHeightPx));

                        // タイプ4で残り1の場合、最後のラベルに下線をつける
                        var fontUnderline = IsUnderlinePrint && CopiesRemainingPerSerial == 1;

                        // シリアル生成、PrintTypeが9かつ最終行の場合は型式下4桁、それ以外はシリアルを生成
                        string generatedCode;
                        if (PrintType != 9 || CopiesRemainingPerSerial != 1) {
                            generatedCode = GenerateCode(SerialNumber, serialType); // シリアルコードを生成
                        }
                        else {
                            generatedCode = Last4ProductModel; // 型式の下4桁を使用
                        }

                        // MakeLabelImageにdpiX, dpiYを渡す
                        using var labelImage = MakeLabelImage(generatedCode, serialType, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY, isPreview);
                        // DrawImageにピクセル単位の座標とサイズを渡す
                        e.Graphics.DrawImage(labelImage, posX, posY, labelWidthPx, labelHeightPx);

                        CopiesRemainingPerSerial--;
                        if (CopiesRemainingPerSerial <= 0) {
                            SerialNumber++;
                            RemainingLabelCount--;
                            //印刷するラベルがなくなった場合の処理
                            if (RemainingLabelCount <= 0) {
                                // 最終行の行番号を表示
                                var sf = new StringFormat {
                                    Alignment = StringAlignment.Near,
                                    LineAlignment = StringAlignment.Center
                                };
                                var layoutRect = new RectangleF(0, posY, 0, (float)labelHeightPx);
                                var rowNumber = (y + 1).ToString();
                                e.Graphics.DrawString(rowNumber, headerFont, Brushes.Black, layoutRect, sf);

                                //e.HasMorePages = false;
                                PageCount = 1;
                                RemainingLabelCount = 0;
                                return false;
                            }
                            CopiesRemainingPerSerial = copiesPerLabel;
                        }
                    }
                }

                if (RemainingLabelCount > 0) {
                    PageCount++;
                    return true;
                }
                else {
                    return false;
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private static Bitmap MakeLabelImage(string text, string serialType, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY, bool isPreview) {

            // ビットマップのサイズをピクセル単位で計算
            var widthPx = (int)Math.Round(labelWidthPx);
            var heightPx = (int)Math.Round(labelHeightPx);

            var labelImage = new Bitmap(widthPx, heightPx);
            labelImage.SetResolution(dpiX, dpiY);

            // 'using'ステートメントでGraphicsオブジェクトを確実に破棄
            using (var g = Graphics.FromImage(labelImage)) {
                // すべての描画操作をピクセル単位で行うように設定
                g.PageUnit = GraphicsUnit.Pixel;

                // 高品質な描画設定
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                // --- 3. 印刷タイプに応じた描画処理 ---
                switch (serialType) {
                    case "Label":
                        DrawLabel(g, text, fontUnderline, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;

                    case "Barcode":
                        DrawBarcode(g, text, labelWidthPx, labelHeightPx, dpiX, dpiY);
                        break;
                }

                // --- 4. プレビュー用の枠を描画 ---
                if (isPreview) {
                    // 0.1mmの黒いペンで枠線を描画
                    using var p = new Pen(Color.Black, ConvertMmToPixel(0.1, dpiX));
                    g.DrawRectangle(p, 0, 0, widthPx - 1, heightPx - 1); // 枠線がはみ出さないように-1
                }
            }

            return labelImage;
        }
        private static void DrawLabel(Graphics g, string text, bool fontUnderline, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            // フォントサイズはポイント単位でそのまま使用
            var fontName = LabelLayoutSettings.TextFont.Name;
            var fontSize = LabelLayoutSettings.TextFont.SizeInPoints;
            var style = fontUnderline ? FontStyle.Underline : FontStyle.Regular;

            using var textFont = new System.Drawing.Font(fontName, fontSize, style);
            // テキストの配置設定
            using var sf = new StringFormat {
                Alignment = LabelLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near,
                LineAlignment = LabelLayoutSettings.AlignTextCenterY ? StringAlignment.Center : StringAlignment.Near
            };
            // 描画領域をピクセル単位で計算
            // TextPositionX, TextPositionY はミリメートル単位なのでピクセルに変換
            var textPosX = LabelLayoutSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(LabelLayoutSettings.TextPositionX, dpiX);
            var textPosY = LabelLayoutSettings.AlignTextCenterY ? 0f : ConvertMmToPixel(LabelLayoutSettings.TextPositionY, dpiY);

            var layoutRect = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

            g.DrawString(text, textFont, Brushes.Black, layoutRect, sf);
        }
        private static void DrawBarcode(Graphics g, string text, float labelWidthPx, float labelHeightPx, float dpiX, float dpiY) {

            // --- テキストの描画 ---
            // フォントサイズはポイント単位でそのまま使用
            var fontName = BarcodeLayoutSettings.TextFont.Name;
            var fontSize = BarcodeLayoutSettings.TextFont.SizeInPoints;

            using (var textFont = new System.Drawing.Font(fontName, fontSize, FontStyle.Regular)) {
                // テキストの配置設定
                using var sf = new StringFormat {
                    Alignment = BarcodeLayoutSettings.AlignTextCenterX ? StringAlignment.Center : StringAlignment.Near
                };

                // TextPositionX, TextPositionY はミリメートル単位なのでピクセルに変換
                var textPosX = BarcodeLayoutSettings.AlignTextCenterX ? 0f : ConvertMmToPixel(BarcodeLayoutSettings.TextPositionX, dpiX);
                var textPosY = ConvertMmToPixel(BarcodeLayoutSettings.TextPositionY, dpiY);

                var layoutRectString = new RectangleF(textPosX, textPosY, labelWidthPx - textPosX, labelHeightPx - textPosY);

                g.DrawString(text, textFont, Brushes.Black, layoutRectString, sf);
            }

            // --- バーコードの描画 ---

            // ZXingはピクセル単位で画像を生成するため、mmからpixelへの変換が必要
            var barcodePixelWidth = ConvertMmToPixel(BarcodeLayoutSettings.BarcodeWidth, dpiX);
            var barcodePixelHeight = ConvertMmToPixel(BarcodeLayoutSettings.BarcodeHeight, dpiY);

            // ZXing用にint型に変換
            var qrWidthPx = (int)Math.Round(barcodePixelWidth);
            var qrHeightPx = (int)Math.Round(barcodePixelHeight);

            var writer = new BarcodeWriter<Bitmap> {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions {
                    Height = qrWidthPx,
                    Width = qrHeightPx,
                    PureBarcode = true // テキストを含まないバーコードのみを生成
                },
                Renderer = new BitmapRenderer()
            };

            using var barcodeBitmap = writer.Write(text);
            // BarcodePositionX, BarcodePositionY はミリメートル単位なのでピクセルに変換
            var barcodePosX = ConvertMmToPixel(BarcodeLayoutSettings.BarcodePositionX, dpiX);
            var barcodePosY = ConvertMmToPixel(BarcodeLayoutSettings.BarcodePositionY, dpiY);
            var barcodeWidthPx = ConvertMmToPixel(BarcodeLayoutSettings.BarcodeWidth, dpiX);
            var barcodeHeightPx = ConvertMmToPixel(BarcodeLayoutSettings.BarcodeHeight, dpiY);

            // X座標を中央に調整 (ピクセル単位で計算)
            if (BarcodeLayoutSettings.AlignBarcodeCenterX) {
                barcodePosX = (labelWidthPx / 2f) - (barcodeWidthPx / 2f);
            }

            var layoutRectBarcode = new RectangleF(barcodePosX, barcodePosY, barcodeWidthPx, barcodeHeightPx);
            // g.DrawImageでミリメートル単位の座標とサイズを指定して描画
            g.DrawImage(barcodeBitmap, layoutRectBarcode);
        }
        private static string ConvertHeaderString(string s) {
            if (ProductInfo == null) { throw new Exception("ProductInfoがnullです。"); }
            s = s.Replace("%P", ProductInfo.ProductName)
                 .Replace("%T", ProductInfo.ProductModel)
                 .Replace("%D", DateTime.Today.ToShortDateString())
                 .Replace("%M", ProductInfo.ProductNumber)
                 .Replace("%O", ProductInfo.OrderNumber)
                 .Replace("%N", ProductInfo.Quantity.ToString())
                 .Replace("%U", ProductInfo.Person);
            return s;
        }
        private static string GenerateCode(int serialCode, string serialType) {
            if (ProductInfo == null) { throw new Exception("ProductInfoがnullです。"); }
            var monthCode = DateTime.Parse(ProductInfo.RegDate).ToString("MM");

            monthCode = monthCode switch {
                "10" => "X",
                "11" => "Y",
                "12" => "Z",
                _ => monthCode
            };

            var outputCode = serialType switch {
                "Label" => LabelLayoutSettings.TextFormat ?? string.Empty,
                "Barcode" => BarcodeLayoutSettings.TextFormat ?? string.Empty,
                _ => string.Empty
            };

            outputCode = outputCode.Replace("%Y", DateTime.Parse(ProductInfo.RegDate).ToString("yy"))
                                    .Replace("%MM", DateTime.Parse(ProductInfo.RegDate).ToString("MM"))
                                    .Replace("%T", ProductInfo.Initial)
                                    .Replace("%R", ProductInfo.Revision)
                                    .Replace("%M", monthCode[^1..])
                                    .Replace("%S", Convert.ToInt32(serialCode).ToString($"D{ProductInfo.SerialDigit}"));

            return outputCode;
        }
    }
    public class PrintOptions {

        public class DocumentPrintSettings  {
            public PrintPageSettings? LabelPageSettings { get; set; }
            public PrintLayoutSettings? LabelLayoutSettings { get; set; }
            public PrintPageSettings? BarcodePageSettings { get; set; }
            public PrintLayoutSettings? BarcodeLayoutSettings { get; set; }

            public DocumentPrintSettings () {
                LabelPageSettings = new PrintPageSettings();
                LabelLayoutSettings = new PrintLayoutSettings();
                BarcodePageSettings = new PrintPageSettings();
                BarcodeLayoutSettings = new PrintLayoutSettings();
            }

            public bool ShouldSerializeLabelPageSettings() {
                return LabelPageSettings != null;
            }
            public bool ShouldSerializeLabelLayoutSettings() {
                return LabelLayoutSettings != null;
            }
            public bool ShouldSerializeBarcodePageSettings() {
                return BarcodePageSettings != null;
            }
            public bool ShouldSerializeBarcodeLayoutSettings() {
                return BarcodeLayoutSettings != null;
            }

            public void SetSettingsType(bool isLabelPrint, bool isBarcodePrint) {
                if (!isLabelPrint) {
                    LabelPageSettings = null;
                    LabelLayoutSettings = null;
                }
                else {
                    LabelPageSettings ??= new PrintPageSettings();
                    LabelLayoutSettings ??= new PrintLayoutSettings();
                }

                if (!isBarcodePrint) {
                    BarcodePageSettings = null;
                    BarcodeLayoutSettings = null;
                }
                else {
                    BarcodePageSettings ??= new PrintPageSettings();
                    BarcodeLayoutSettings ??= new PrintLayoutSettings();
                }
            }
        }

        public class PrintPageSettings {
            public int LabelsPerRow { get; set; }
            public int LabelsPerColumn { get; set; }
            public double LabelWidth { get; set; }
            public double LabelHeight { get; set; }
            public double MarginX { get; set; }
            public double MarginY { get; set; }
            public double IntervalX { get; set; }
            public double IntervalY { get; set; }
            public string HeaderTextFormat { get; set; } = string.Empty;
            public double HeaderPositionX { get; set; }
            public double HeaderPositionY { get; set; }

            [JsonConverter(typeof(FontJsonConverter))]
            public Font HeaderFont { get; set; } = SystemFonts.DefaultFont;
        }
        public class PrintLayoutSettings {
            public string TextFormat { get; set; } = string.Empty;
            public bool AlignTextCenterX { get; set; }
            public bool AlignTextCenterY { get; set; }
            public double TextPositionX { get; set; }
            public double TextPositionY { get; set; }
            public bool AlignBarcodeCenterX { get; set; }
            public double BarcodeHeight { get; set; }
            public double BarcodeWidth { get; set; }
            public double BarcodePositionX { get; set; }
            public double BarcodePositionY { get; set; }
            public int CopiesPerLabel { get; set; }

            [JsonConverter(typeof(FontJsonConverter))]
            public Font TextFont { get; set; } = SystemFonts.DefaultFont;
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
