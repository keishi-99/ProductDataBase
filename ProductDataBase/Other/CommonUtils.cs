using NPOI.SS.UserModel;
using NPOI.SS.Util;
using OfficeOpenXml;
using ProductDatabase.Product;
using System.Data;
using System.Data.SQLite;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Windows.Compatibility;
using static ProductDatabase.MainWindow;

namespace ProductDatabase.Other {

    internal static partial class NativeMethods {
        [LibraryImport("user32.dll", SetLastError = true)]
        public static partial void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
    }

    internal partial class CommonUtils {
        public static string s_networkPath = string.Empty; // ClonePathを保持する静的変数
        /// <summary>
        /// ファイルをコピーします。
        /// </summary>
        private static void CopyWithRetry(string sourceFileName, string destFileName, bool overwrite, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(sourceFileName)) {
                        File.Copy(sourceFileName, destFileName, overwrite);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルがコピーできません。"); // 最後の試行で失敗したら例外を投げる
                    }
                    Thread.Sleep(delayMilliseconds); // 一定時間待機してリトライ
                }
            }
        }
        /// <summary>
        /// ファイルを削除します。
        /// </summary>
        private static void DeleteWithRetry(string filePath, int retryCount = 5, int delayMilliseconds = 2000) {
            for (var attempt = 1; attempt <= retryCount; attempt++) {
                try {
                    if (File.Exists(filePath)) {
                        File.Delete(filePath);
                    }
                    return;
                } catch (IOException) {
                    if (attempt == retryCount) {
                        throw new Exception("バックアップファイルが削除できません。"); // 最後の試行で失敗したら例外を投げる
                    }
                    Thread.Sleep(delayMilliseconds); // 一定時間待機してリトライ
                }
            }
        }
        // ログ作成
        public static class Logger {
            private static readonly string s_logDirectory = Path.Combine(Environment.CurrentDirectory, "db", "logs"); // ログを保存するディレクトリ
            private static readonly object s_lockObject = new();

            /// <summary>
            /// 作業ログを追記します。
            /// </summary>
            /// <param name="message">記録する作業内容</param>
            public static void AppendLog(string[] message) {
                try {
                    lock (s_lockObject) {
                        // ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_logDirectory)) {
                            Directory.CreateDirectory(s_logDirectory);
                        }

                        //// 年と月を含むログファイル名を生成
                        var logFileName = $"log_{DateTime.Now:yyyyMM}.csv";
                        var logFilePath = Path.Combine(s_logDirectory, logFileName);

                        // CSV形式でログ内容をファイルの末尾に追記
                        var logEntry = $"\"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\",{string.Join(",", message.Select(m => $"\"{m.Replace("\"", "\"\"").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ")}\""))}";
                        // ログ内容をファイルの末尾に追記
                        File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

                        if (!string.IsNullOrEmpty(s_networkPath)) {
                            var cloneFilePath = Path.Combine(s_networkPath, "db", "logs", logFileName);
                            if (cloneFilePath != logFilePath) {
                                CopyWithRetry(logFilePath, cloneFilePath, true);
                            }
                        }
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"ログの書き込み中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // バックアップ作成
        public static class BackupManager {
            private static readonly string s_backupDirectory = Path.Combine(Environment.CurrentDirectory, "db", "backup"); // バックアップを保存するディレクトリ
            private static readonly string s_originalFilePath = Path.Combine(Environment.CurrentDirectory, "db", "registration.db"); // 元ファイルパス
            private static readonly int s_maxBackupFiles = 20; // 最大バックアップファイル数
            private static readonly object s_lockObject = new();

            /// <summary>
            /// バックアップを作成します。
            /// </summary>
            public static void CreateBackup() {
                try {
                    lock (s_lockObject) {
                        // バックアップ用ディレクトリが存在しない場合は作成
                        if (!Directory.Exists(s_backupDirectory)) {
                            Directory.CreateDirectory(s_backupDirectory);
                        }

                        // 日付と時間をファイル名に付加
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var backupFileName = $"registration_{timestamp}.db";
                        var backupFilePath = Path.Combine(s_backupDirectory, backupFileName);

                        // 元ファイルをバックアップにコピー
                        CopyWithRetry(s_originalFilePath, backupFilePath, true);
                        // バックアップファイルを管理
                        ManageBackupFiles();

                        // ネットワークにバックアップ
                        var networkFilePath = Path.Combine(s_networkPath, "db", "registration.db");
                        if (Environment.CurrentDirectory != s_networkPath) {
                            CopyWithRetry(s_originalFilePath, networkFilePath, true);
                        }

                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            /// <summary>
            /// 古いバックアップファイルを削除します。
            /// </summary>
            private static void ManageBackupFiles() {
                try {
                    var backupFiles = Directory.GetFiles(s_backupDirectory, "registration_*.db")
                        .OrderBy(File.GetCreationTime) // 作成日時順に並べる
                        .ToList();

                    while (backupFiles.Count > s_maxBackupFiles) {
                        var oldestFile = backupFiles.First();
                        DeleteWithRetry(oldestFile);
                        backupFiles.RemoveAt(0);
                    }
                } catch (Exception ex) {
                    MessageBox.Show($"バックアップの作成中にエラーが発生しました: {ex.Message}", $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        // CapsLockがオンになっていたらCapsLockを解除する
        public static partial class Keyboard {
            private const byte VK_CAPITAL = 0x14; // CapsLock の仮想キーコード
            private const int KEYEVENTF_EXTENDEDKEY = 0x1;
            private const int KEYEVENTF_KEYUP = 0x2;

            // CapsLock の状態を切り替える
            public static void CapsDisable() {
                if (Control.IsKeyLocked(Keys.CapsLock)) {
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_EXTENDEDKEY, 0);
                    NativeMethods.keybd_event(VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0);
                }
            }
        }

        // シリアル印刷管理クラス
        public static class PrintManager {

            // 状態保持用プロパティ（外部からは読み取り専用）
            public static ProductInformation ProductInfo { get; private set; } = default!;
            public static ProductPrintSettings ProductPrintSettings { get; private set; } = new();

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

            public static void Initialize(ProductInformation productInfo, ProductPrintSettings productPrintSettings) {
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

        // 成績書生成データを保持するクラス
        public class ReportConfig {
            public required string DirectoryPath { get; set; }
            public required string FileName { get; set; }
            public required string FileExtension { get; set; }
            public required string SearchFileName { get; set; }
            public required string SheetName { get; set; }
            public string? ProductNumberRange { get; set; }
            public string? OrderNumberRange { get; set; }
            public string? QuantityRange { get; set; }
            public string? SerialFirstRange { get; set; }
            public string? SerialLastRange { get; set; }
            public string? ProductModelRange { get; set; }
            public string? SaveDirectory { get; set; }
        }
        // 成績書生成
        public static class ReportGenerator {
            // レポートを生成するメインメソッド
            public static void GenerateReport(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var reportConfig = GetReportConfig(configWorkbook, productInfo.ProductModel);

                    // 2. レポートテンプレートを読み込み
                    var reportWorkbook = LoadReportTemplate(reportConfig.DirectoryPath, reportConfig.SearchFileName);

                    // 3. レポートシートにデータを挿入
                    PopulateReportSheet(reportWorkbook, productInfo, reportConfig);

                    // 4. レポートを保存
                    SaveReport(reportWorkbook, productInfo, reportConfig);

                    MessageBox.Show("成績書が正常に生成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    // エラーメッセージをユーザーに表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigReport.xlsx");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static ReportConfig GetReportConfig(ExcelPackage configWorkbook, string productModel) {
                var workSheetMain = configWorkbook.Workbook.Worksheets["Sheet1"] ?? throw new Exception("設定ファイルに 'Sheet1' が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productModel)
                    ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得し、ReportConfigオブジェクトに格納
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 3].Value?.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(directoryPath)) { throw new Exception("Configのファイルパスが無効です。"); }
                if (!Directory.Exists(directoryPath)) { throw new FileNotFoundException($"指定されたフォルダが存在しません: {directoryPath}"); }

                var searchName = workSheetMain.Cells[searchAddressResultRow, 4].Value?.ToString()?.Trim('"');
                if (string.IsNullOrWhiteSpace(searchName)) { throw new Exception("Configのファイル名が無効です。"); }

                var filePaths = Directory.GetFiles(directoryPath, $"*{searchName}*", SearchOption.TopDirectoryOnly);
                var filePath = filePaths[0];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var fileExtension = Path.GetExtension(filePath).ToLower();

                var sheetName = workSheetMain.Cells[searchAddressResultRow, 5].Value?.ToString();
                return string.IsNullOrWhiteSpace(sheetName)
                    ? throw new Exception("シート名がありません。")
                    : new ReportConfig {
                        DirectoryPath = directoryPath,
                        FileName = fileName,
                        FileExtension = fileExtension,
                        SearchFileName = searchName,
                        SheetName = sheetName,
                        ProductNumberRange = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString() ?? string.Empty,
                        OrderNumberRange = workSheetMain.Cells[searchAddressResultRow, 7].Value?.ToString() ?? string.Empty,
                        QuantityRange = workSheetMain.Cells[searchAddressResultRow, 8].Value?.ToString() ?? string.Empty,
                        SerialFirstRange = workSheetMain.Cells[searchAddressResultRow, 9].Value?.ToString() ?? string.Empty,
                        SerialLastRange = workSheetMain.Cells[searchAddressResultRow, 10].Value?.ToString() ?? string.Empty,
                        ProductModelRange = workSheetMain.Cells[searchAddressResultRow, 11].Value?.ToString() ?? string.Empty,
                        SaveDirectory = workSheetMain.Cells[searchAddressResultRow, 12].Value?.ToString() ?? string.Empty
                    };
            }

            // レポートテンプレートExcelワークブックを読み込む
            private static ExcelPackage LoadReportTemplate(string directoryPath, string searchFileName) {
                var filePaths = Directory.GetFiles(directoryPath, $"*{searchFileName}*", SearchOption.TopDirectoryOnly);
                if (filePaths.Length == 0) {
                    throw new FileNotFoundException($"指定されたファイル名 '{searchFileName}' のファイルが '{directoryPath}' に見つかりません。");
                }
                var filePath = filePaths[0]; // 最初のファイルを使用

                try {
                    FileStream fileStreamReport = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamReport);
                } catch (Exception ex) {
                    throw new Exception($"レポートテンプレートの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // レポートシートに製品情報を挿入する
            private static void PopulateReportSheet(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfig config) {
                var workSheetTemp = reportWorkbook.Workbook.Worksheets[config.SheetName] ?? throw new Exception($"シート '{config.SheetName}' が見つかりません。");

                // セルに値を挿入
                if (!string.IsNullOrEmpty(config.ProductNumberRange) && !string.IsNullOrEmpty(productInfo.ProductNumber)) {
                    var productNumbers = productInfo.ProductNumber.Split('-');
                    workSheetTemp.Cells[config.ProductNumberRange].Value = productNumbers[0];
                }
                if (!string.IsNullOrEmpty(config.OrderNumberRange) && !string.IsNullOrEmpty(productInfo.OrderNumber)) {
                    workSheetTemp.Cells[config.OrderNumberRange].Value = productInfo.OrderNumber;
                }
                if (!string.IsNullOrEmpty(config.QuantityRange) && productInfo.Quantity != 0) {
                    workSheetTemp.Cells[config.QuantityRange].Value = productInfo.Quantity;
                }
                if (!string.IsNullOrEmpty(config.SerialFirstRange) && !string.IsNullOrEmpty(productInfo.SerialFirst)) {
                    workSheetTemp.Cells[config.SerialFirstRange].Value = productInfo.SerialFirst;
                }
                if (!string.IsNullOrEmpty(config.SerialLastRange) && !string.IsNullOrEmpty(productInfo.SerialLast)) {
                    workSheetTemp.Cells[config.SerialLastRange].Value = productInfo.SerialLast;
                }
                if (!string.IsNullOrEmpty(config.ProductModelRange) && !string.IsNullOrEmpty(productInfo.ProductModel)) {
                    workSheetTemp.Cells[config.ProductModelRange].Value = productInfo.ProductModel;
                }
            }

            // 変更されたレポートをファイルに保存する
            private static void SaveReport(ExcelPackage reportWorkbook, ProductInformation productInfo, ReportConfig config) {
                var fileName = config.FileName;
                var fileExtension = config.FileExtension;
                var initialDirectory = config.SaveDirectory;

                using SaveFileDialog saveFileDialog = new() {
                    Filter = $"Excel Files (*{fileExtension})|*{fileExtension}|All Files (*.*)|*.*",
                    FileName = $"{fileName} のコピー{productInfo.ProductNumber}{fileExtension}",
                    Title = "保存先を選択してください",
                    InitialDirectory = initialDirectory ?? Environment.CurrentDirectory // Nullの場合はデフォルトディレクトリを使用
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                    var outputPath = saveFileDialog.FileName;
                    var fileInfo = new FileInfo(outputPath);
                    reportWorkbook.SaveAs(fileInfo);
                }
                else {
                    // ユーザーが保存をキャンセルした場合の処理
                    throw new OperationCanceledException("キャンセルされました。");
                }
            }
        }

        // リスト生成
        public static class ListGenerator {
            // メインのオーケストレーションメソッド
            public static void GenerateList(ProductInformation productInfo) {
                try {
                    // 1. Excel設定の読み込みとワークブックの準備
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigList.xlsx");
                    using var workBook = new ExcelPackage(new FileInfo(configPath));
                    var (targetSheetName, resultRow, workSheetMain) = LoadExcelConfiguration(workBook, productInfo.ProductModel);

                    // 2. 製品情報の設定とExcelへの書き込み
                    var productCellRanges = GetProductCellRanges(workSheetMain, resultRow);
                    var workSheetTemp = workBook.Workbook.Worksheets[targetSheetName] ?? throw new Exception($"テンプレートシート:[{targetSheetName}]が見つかりません。");
                    PopulateProductDetails(workSheetTemp, productInfo, productCellRanges);

                    // 3. データベースから使用済み基板情報を取得
                    var usedSubstrate = GetUsedSubstrateData(productInfo);

                    // 4. 基板情報のExcelへの書き込み
                    UpdateSubstrateDetailsInExcel(workSheetMain, workSheetTemp, resultRow, usedSubstrate);

                    // 5. QRコードの生成と埋め込み
                    if (!string.IsNullOrEmpty(productCellRanges.QrCodeRange)) {
                        GenerateAndEmbedQrCode(workSheetTemp, productInfo, productCellRanges.QrCodeRange);
                    }

                    // 6. Excelファイルの保存と印刷
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyList.xlsx");
                    SaveAndPrintExcel(workBook, temporarilyPath, targetSheetName);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定を読み込むメソッド
            private static (string sheetName, int resultRow, ExcelWorksheet workSheetMain) LoadExcelConfiguration(ExcelPackage workBook, string productModel) {
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName] ?? throw new Exception($"設定シート:[{targetSheetName}]が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productModel)
                                          ?? throw new Exception($"Configに品目番号:[{productModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                var sheetName = workSheetMain.Cells[resultRow, 2].Value?.ToString()
                                ?? throw new Exception($"設定シートの行 {resultRow}, 列 2 にシート名が設定されていません。");

                return (sheetName, resultRow, workSheetMain);
            }

            // 製品情報に関連するExcelのセル範囲を取得するヘルパーメソッド
            private static ProductCellRanges GetProductCellRanges(ExcelWorksheet workSheetMain, int resultRow) {
                return new ProductCellRanges {
                    ProductNameRange = workSheetMain.Cells[resultRow, 4].Value?.ToString() ?? string.Empty,
                    ProductNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString() ?? string.Empty,
                    OrderNumberRange = workSheetMain.Cells[resultRow, 6].Value?.ToString() ?? string.Empty,
                    RegDateRange = workSheetMain.Cells[resultRow, 7].Value?.ToString() ?? string.Empty,
                    ProductModelConfig = workSheetMain.Cells[resultRow, 8].Value?.ToString() ?? string.Empty,
                    ProductModelRange = workSheetMain.Cells[resultRow, 9].Value?.ToString() ?? string.Empty,
                    QuantityRange = workSheetMain.Cells[resultRow, 10].Value?.ToString() ?? string.Empty,
                    SerialFirstRange = workSheetMain.Cells[resultRow, 11].Value?.ToString() ?? string.Empty,
                    SerialLastRange = workSheetMain.Cells[resultRow, 12].Value?.ToString() ?? string.Empty,
                    CommentRange = workSheetMain.Cells[resultRow, 13].Value?.ToString() ?? string.Empty,
                    QrCodeRange = workSheetMain.Cells[resultRow, 14].Value?.ToString() ?? string.Empty
                };
            }

            // ProductCellRanges クラスを定義
            private class ProductCellRanges {
                public string? ProductNameRange { get; set; }
                public string? ProductNumberRange { get; set; }
                public string? OrderNumberRange { get; set; }
                public string? RegDateRange { get; set; }
                public string? ProductModelConfig { get; set; }
                public string? ProductModelRange { get; set; }
                public string? QuantityRange { get; set; }
                public string? SerialFirstRange { get; set; }
                public string? SerialLastRange { get; set; }
                public string? CommentRange { get; set; }
                public string? QrCodeRange { get; set; }
            }


            // 製品情報をExcelシートに書き込むメソッド
            private static void PopulateProductDetails(ExcelWorksheet workSheetTemp, ProductInformation productInfo, ProductCellRanges ranges) {
                if (!string.IsNullOrEmpty(ranges.ProductNameRange)) { workSheetTemp.Cells[ranges.ProductNameRange].Value = productInfo.ProductModel; } // ProductNameはProductModelから取得
                if (!string.IsNullOrEmpty(ranges.ProductNumberRange)) { workSheetTemp.Cells[ranges.ProductNumberRange].Value = productInfo.ProductNumber; }
                if (!string.IsNullOrEmpty(ranges.OrderNumberRange)) { workSheetTemp.Cells[ranges.OrderNumberRange].Value = productInfo.OrderNumber; }
                if (!string.IsNullOrEmpty(ranges.RegDateRange)) { workSheetTemp.Cells[ranges.RegDateRange].Value = productInfo.RegDate; }
                if (!string.IsNullOrEmpty(ranges.ProductModelRange)) { workSheetTemp.Cells[ranges.ProductModelRange].Value = productInfo.ProductModel; }
                if (!string.IsNullOrEmpty(ranges.QuantityRange)) { workSheetTemp.Cells[ranges.QuantityRange].Value = productInfo.Quantity; }
                if (!string.IsNullOrEmpty(ranges.SerialFirstRange)) { workSheetTemp.Cells[ranges.SerialFirstRange].Value = productInfo.SerialFirst; }
                if (!string.IsNullOrEmpty(ranges.SerialLastRange)) { workSheetTemp.Cells[ranges.SerialLastRange].Value = productInfo.SerialLast; }
                if (!string.IsNullOrEmpty(ranges.CommentRange)) { workSheetTemp.Cells[ranges.CommentRange].Value = productInfo.Comment; }
            }

            // データベースから使用済み基板情報を取得するメソッド
            private static List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> GetUsedSubstrateData(ProductInformation productInfo) {
                List<(string, List<string>, List<int>)> usedSubstrate = [];

                using SQLiteConnection con = new(GetConnectionRegistration());
                con.Open();
                using var cmd = con.CreateCommand();

                var tableName = $"[{productInfo.CategoryName}_Substrate]";
                cmd.CommandText =
                    $"""
                    SELECT
                        SubstrateModel,
                        SubstrateNumber,
                        Decrease
                    FROM
                        {tableName}
                    WHERE
                        UseID = @ID
                    ORDER BY
                        SubstrateModel ASC
                    ;
                    """;
                cmd.Parameters.Add("@ID", DbType.String).Value = productInfo.ProductID;
                using var dr = cmd.ExecuteReader();

                while (dr.Read()) {
                    var substrateModel = dr.GetString(0);
                    var substrateNumber = dr.GetString(1);
                    var decrease = -1 * dr.GetInt32(2);

                    // 既存の substrateModel を検索し、見つかればリストに追加、なければ新しいエントリを作成
                    var existingSubstrateIndex = usedSubstrate.FindIndex(x => x.Item1 == substrateModel);

                    if (existingSubstrateIndex != -1) {
                        usedSubstrate[existingSubstrateIndex].Item2.Add(substrateNumber);
                        usedSubstrate[existingSubstrateIndex].Item3.Add(decrease);
                    }
                    else {
                        List<string> substrateNumbers = [substrateNumber];
                        List<int> decreases = [decrease];
                        usedSubstrate.Add((substrateModel, substrateNumbers, decreases));
                    }
                }
                return usedSubstrate;
            }

            // 基板情報をExcelシートに書き込むメソッド
            private static void UpdateSubstrateDetailsInExcel(ExcelWorksheet workSheetMain, ExcelWorksheet workSheetTemp, int resultRow, List<(string SubstrateModel, List<string> SubstrateNumbers, List<int> Decreases)> usedSubstrate) {
                foreach (var (substrateModel, substrateNumbers, decreases) in usedSubstrate) {
                    var targetRow = resultRow;
                    var searchValue = substrateModel;
                    var foundColumn = 0;

                    var searchAddressResult2 = workSheetMain.Cells
                                                .Where(x => x.Start.Row == targetRow && x.Value?.ToString() == searchValue)
                                                .FirstOrDefault();

                    foundColumn = searchAddressResult2 != null
                        ? searchAddressResult2.Start.Column
                        : throw new Exception($"Configに基板モデル:[{searchValue}]が見つかりません。");

                    var mainCellValue = workSheetMain.Cells[resultRow, foundColumn + 1].Value?.ToString()
                        ?? throw new Exception($"Configシートの行 {resultRow}, 列 {foundColumn + 1} にセル範囲が設定されていません。");

                    var tempCellValue = string.Join("   ", substrateNumbers.Select((subStrateNumber, k) => $"{subStrateNumber}({decreases[k]})"));

                    workSheetTemp.Cells[mainCellValue].Value = tempCellValue;
                }
            }

            // QRコードを生成し、Excelシートに埋め込むメソッド
            private static void GenerateAndEmbedQrCode(ExcelWorksheet workSheetTemp, ProductInformation productInfo, string qrCodeRange) {
                BarcodeWriter<Bitmap> qr = new() {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new QrCodeEncodingOptions {
                        ErrorCorrection = ErrorCorrectionLevel.L,
                        CharacterSet = "Shift_JIS",
                        Width = 90,
                        Height = 90,
                    },
                    Renderer = new BitmapRenderer {
                        Foreground = Color.Gray,
                        Background = Color.White,
                    }
                };

                // QRコードのデータ文字列を構築
                var qrData = $"{productInfo.OrderNumber};{productInfo.ProductNumber};{productInfo.ProductModel};{productInfo.Quantity};{productInfo.SerialFirst};{productInfo.SerialLast}";
                var qrBitmap = qr.Write(qrData);

                // Excelに画像を埋め込む
                var picture = workSheetTemp.Drawings.AddPicture("QR", qrBitmap);
                // QRコードのセル範囲の左上隅に画像を配置
                picture.SetPosition(workSheetTemp.Cells[qrCodeRange].Start.Row - 1, 0, workSheetTemp.Cells[qrCodeRange].Start.Column - 1, 0);
            }

            // Excelファイルを保存し、印刷処理を行うメソッド
            private static void SaveAndPrintExcel(ExcelPackage workBook, string temporarilyPath, string sheetName) {
                var fileInfo = new FileInfo(temporarilyPath);
                workBook.SaveAs(fileInfo);

                // 印刷 (COM Interop)
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;
                Microsoft.Office.Interop.Excel.Sheets? xlSheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? xlSheet = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    // ワークブック開く
                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(temporarilyPath, ReadOnly: true);

                    // ワークシート選択
                    xlSheets = xlBook.Sheets;
                    xlSheet = xlSheets[sheetName];

                    // ワークシート表示
                    xlSheet.Activate();

                    // 印刷処理が必要な場合は、ここに xlSheet.PrintOut() などを追加します。
                    // 例: xlSheet.PrintOut();

                    // ワークブックを閉じてExcelを終了 (必要に応じてコメント解除)
                    // xlBook.Close(false);
                    // xlApp.Quit();
                } finally {
                    // COMオブジェクトの解放
                    if (xlSheet != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                    }

                    if (xlSheets != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                    }

                    if (xlBook != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                    }

                    if (xlBooks != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                    }

                    if (xlApp != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    }
                }
            }
        }

        // チェックシート生成データを保持するクラス
        public class CheckSheetConfigData {
            public string? ProductModelRange { get; set; }
            public string? ProductNumberRange { get; set; }
            public string? OrderNumberRange { get; set; }
            public string? QuantityRange { get; set; }
            public string? SerialFirstRange { get; set; }
            public string? SerialLastRange { get; set; }
            public string? RegDateRange { get; set; }
            public string? DateFormat { get; set; }
            public string? RegTemperatureRange { get; set; }
            public string? RegHumidityRange { get; set; }
            public List<string> SheetNames { get; set; } = [];
        }
        // チェックシート生成
        public static class CheckSheetGenerator {
            // Excelチェックシートを生成し、データを書き込み、印刷します。
            public static void GenerateCheckSheetNPOI(ProductInformation productInfo) {
                try {
                    // 設定ファイルのパスを構築
                    var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigCheckSheet.xlsx");
                    var temporarilyPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "temporarilyCheckSheet.xlsx");

                    // 1. 設定ファイルの読み込みとメインシートの取得
                    // EPPlusを使用して設定ファイルを読み込み、必要な情報を抽出
                    var excelData = LoadAndExtractConfig(configPath, productInfo);

                    // 2. 温度・湿度入力ダイアログの表示と値の取得
                    (var temperature, var humidity) = GetTemperatureAndHumidity(excelData);

                    // 3. 日付のフォーマット
                    var formattedDate = FormatDate(productInfo.RegDate, excelData.DateFormat);

                    // 4. NPOIでExcelファイルを編集
                    // 設定ファイルからNPOIワークブックを作成
                    using (var workBookNPOI = WorkbookFactory.Create(configPath)) {
                        // セル参照をCellReferenceオブジェクトとして作成
                        var cellReferences = CreateCellReferences(excelData);

                        // 各シートに対して値を書き込む
                        PopulateExcelSheets(workBookNPOI, productInfo, cellReferences, temperature, humidity, formattedDate, excelData.SheetNames);

                        // 不要なシートを非表示にする
                        HideSheets(workBookNPOI, excelData.SheetNames);

                        // ブックを保存
                        SaveWorkbook(workBookNPOI, temporarilyPath);
                    }

                    // 5. Excel Interopを使用して印刷
                    PrintExcelFile(temporarilyPath);
                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Excel設定ファイルを読み込み、メインシートから設定データを抽出します。
            private static CheckSheetConfigData LoadAndExtractConfig(string configPath, ProductInformation productInfo) {
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                using FileStream fileStream = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var workBook = new ExcelPackage(fileStream);

                // 既存ワークシートを取得
                var sheet = workBook.Workbook.Worksheets;
                var targetSheetName = "Sheet1";
                var workSheetMain = sheet[targetSheetName];

                // セル検索: ProductModelに基づいて行を特定
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == productInfo.ProductModel)
                                          ?? throw new Exception($"Configに品目番号:[{productInfo.ProductModel}]が見つかりません。");
                var resultRow = searchAddressResult.Start.Row;

                // ワークシートのセルから値を取得し、ExcelConfigDataオブジェクトに格納
                var excelData = new CheckSheetConfigData {
                    ProductModelRange = workSheetMain.Cells[resultRow, 3].Value?.ToString() ?? string.Empty,
                    ProductNumberRange = workSheetMain.Cells[resultRow, 4].Value?.ToString() ?? string.Empty,
                    OrderNumberRange = workSheetMain.Cells[resultRow, 5].Value?.ToString() ?? string.Empty,
                    QuantityRange = workSheetMain.Cells[resultRow, 6].Value?.ToString() ?? string.Empty,
                    SerialFirstRange = workSheetMain.Cells[resultRow, 7].Value?.ToString() ?? string.Empty,
                    SerialLastRange = workSheetMain.Cells[resultRow, 8].Value?.ToString() ?? string.Empty,
                    RegDateRange = workSheetMain.Cells[resultRow, 9].Value?.ToString() ?? string.Empty,
                    DateFormat = workSheetMain.Cells[resultRow, 10].Value?.ToString() ?? string.Empty,
                    RegTemperatureRange = workSheetMain.Cells[resultRow, 11].Value?.ToString() ?? string.Empty,
                    RegHumidityRange = workSheetMain.Cells[resultRow, 12].Value?.ToString() ?? string.Empty,
                    SheetNames = [.. Enumerable.Range(13, 20)
                                       .Select(column => workSheetMain.Cells[resultRow, column].Value?.ToString() ?? string.Empty)
                                       .TakeWhile(sheetName => !string.IsNullOrWhiteSpace(sheetName))]
                };

                return excelData.SheetNames.Count == 0 ? throw new Exception("対象シートがConfigファイルに設定されていません。") : excelData;
            }

            // 温度と湿度の入力ダイアログを表示し、ユーザーからの値を取得します。
            private static (string temperature, string humidity) GetTemperatureAndHumidity(CheckSheetConfigData excelData) {
                var temperature = string.Empty;
                var humidity = string.Empty;

                // 温度セルか湿度セルが設定されている場合、ダイアログを表示
                if (!string.IsNullOrEmpty(excelData.RegTemperatureRange) || !string.IsNullOrEmpty(excelData.RegHumidityRange)) {
                    var dialog = new InputDialog1();
                    var result = dialog.ShowDialog();
                    if (result != DialogResult.OK) {
                        // キャンセルされた場合は処理を中断
                        throw new OperationCanceledException("ユーザーによって温度・湿度入力がキャンセルされました。");
                    }
                    temperature = dialog.Temperature;
                    humidity = dialog.Humidity;
                }
                return (temperature, humidity);
            }

            // 日付を指定されたフォーマットで文字列に変換します。
            private static string FormatDate(string dateString, string? dateFormat) {
                var formattedDate = string.Empty;
                if (DateTime.TryParse(dateString, out var date)) {
                    formattedDate = dateFormat switch {
                        "1" => date.ToString("yyyy年MM月dd日"),
                        "2" => date.ToString("yyyy-MM-dd"),
                        _ => formattedDate
                    };
                }
                return formattedDate;
            }

            // ExcelConfigDataからNPOIのCellReferenceオブジェクトを作成します。
            private static dynamic CreateCellReferences(CheckSheetConfigData excelData) {
                // セル参照をCellReferenceオブジェクトとして作成
                return new {
                    ProductModel = string.IsNullOrEmpty(excelData.ProductModelRange) ? null : new CellReference(excelData.ProductModelRange),
                    ProductNumber = string.IsNullOrEmpty(excelData.ProductNumberRange) ? null : new CellReference(excelData.ProductNumberRange),
                    OrderNumber = string.IsNullOrEmpty(excelData.OrderNumberRange) ? null : new CellReference(excelData.OrderNumberRange),
                    Quantity = string.IsNullOrEmpty(excelData.QuantityRange) ? null : new CellReference(excelData.QuantityRange),
                    SerialFirst = string.IsNullOrEmpty(excelData.SerialFirstRange) ? null : new CellReference(excelData.SerialFirstRange),
                    SerialLast = string.IsNullOrEmpty(excelData.SerialLastRange) ? null : new CellReference(excelData.SerialLastRange),
                    RegDate = string.IsNullOrEmpty(excelData.RegDateRange) ? null : new CellReference(excelData.RegDateRange),
                    RegTemperature = string.IsNullOrEmpty(excelData.RegTemperatureRange) ? null : new CellReference(excelData.RegTemperatureRange),
                    RegHumidity = string.IsNullOrEmpty(excelData.RegHumidityRange) ? null : new CellReference(excelData.RegHumidityRange),
                };
            }

            // 指定されたNPOIワークブックの各シートに製品情報を書き込みます。
            private static void PopulateExcelSheets(
                IWorkbook workBookNPOI,
                ProductInformation productInfo,
                dynamic cellReferences,
                string temperature,
                string humidity,
                string formattedDate,
                List<string> sheetNames) {
                foreach (var sheetName in sheetNames) {
                    var sheetNPOI = workBookNPOI.GetSheet(sheetName) ?? throw new Exception($"シート[{sheetName}]が見つかりません。");
                    sheetNPOI.ForceFormulaRecalculation = true; // 数式の再計算を強制

                    // 各セルに値を書き込む
                    WriteCellValue(sheetNPOI, cellReferences.ProductModel, productInfo.ProductModel);
                    WriteCellValue(sheetNPOI, cellReferences.ProductNumber, productInfo.ProductNumber);
                    WriteCellValue(sheetNPOI, cellReferences.OrderNumber, productInfo.OrderNumber);
                    WriteCellValue(sheetNPOI, cellReferences.Quantity, productInfo.Quantity.ToString());
                    WriteCellValue(sheetNPOI, cellReferences.SerialFirst, productInfo.SerialFirst);
                    WriteCellValue(sheetNPOI, cellReferences.SerialLast, productInfo.SerialLast);
                    WriteCellValue(sheetNPOI, cellReferences.RegDate, formattedDate);
                    WriteCellValue(sheetNPOI, cellReferences.RegTemperature, temperature);
                    WriteCellValue(sheetNPOI, cellReferences.RegHumidity, humidity);
                }
            }

            // 指定されたシートのセルに値を書き込みます。
            private static void WriteCellValue(ISheet sheet, CellReference? reference, string value) {
                if (reference == null) {
                    return; // 参照がnullの場合は何もしない
                }

                // 指定した行を取得できない時はエラーとならないよう新規作成している
                var row = sheet.GetRow(reference.Row) ?? sheet.CreateRow(reference.Row);
                // 一行上の処理の列版
                var cell = row.GetCell(reference.Col) ?? row.CreateCell(reference.Col);
                cell.SetCellValue(value);
            }

            // 指定されたNPOIワークブックの不要なシートを非表示にします。
            private static void HideSheets(IWorkbook workBookNPOI, List<string> sheetsToKeep) {
                // 非表示にするシートのインデックスを特定
                var allSheetNames = new List<string>();
                for (var i = 0; i < workBookNPOI.NumberOfSheets; i++) {
                    allSheetNames.Add(workBookNPOI.GetSheetName(i));
                }

                var sheetIndicesToHide = allSheetNames.Select((name, index) => new { Name = name, Index = index })
                                                      .Where(sheet => !sheetsToKeep.Contains(sheet.Name))
                                                      .Select(sheet => sheet.Index)
                                                      .ToList();

                // シート（"Sheet1"）を非表示にする
                if (workBookNPOI.GetSheetIndex("Sheet1") != -1 && !sheetIndicesToHide.Contains(workBookNPOI.GetSheetIndex("Sheet1"))) {
                    sheetIndicesToHide.Add(workBookNPOI.GetSheetIndex("Sheet1"));
                }

                // シートを非表示に設定
                foreach (var sheetIndex in sheetIndicesToHide) {
                    workBookNPOI.SetSheetHidden(sheetIndex, SheetVisibility.VeryHidden);
                }
            }

            // NPOIワークブックをファイルに保存します。
            private static void SaveWorkbook(IWorkbook workBookNPOI, string outputPath) {
                try {
                    using var fs = new FileStream(outputPath, FileMode.Create);
                    workBookNPOI.Write(fs);
                } catch (IOException ex) {
                    throw new IOException($"Excelファイルの保存に失敗しました: {outputPath}. 詳細: {ex.Message}", ex);
                }
            }

            // Excel Interopを使用して指定されたExcelファイルを開きます。
            private static void PrintExcelFile(string filePath) {
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(filePath, ReadOnly: true);

                    //// ワークブックを閉じる (falseは変更を保存しない)
                    //xlBook.Close(false);
                    //// Excelアプリケーションを終了
                    //xlApp.Quit();
                } catch (Exception ex) {
                    throw new Exception($"エラーが発生しました。詳細: {ex.Message}", ex);
                } finally {
                    // COMオブジェクトの解放
                    if (xlBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook); }
                    if (xlBooks != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks); }
                    if (xlApp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp); }
                }
            }
        }

        // 基板設定を開く
        public static class SubstrateInformation {
            // 基板設定ファイルを開くメソッド
            public static void OpenSubstrateInformation(ProductInformation productInfo) {
                try {
                    // 1. 設定ファイルを読み込み、レポート設定を取得
                    var configWorkbook = LoadConfigWorkbook();
                    var (filePaths, fileName) = GetSubstrateConfig(configWorkbook, productInfo.SubstrateModel);

                    OpenExcel(filePaths, fileName);

                } catch (Exception ex) {
                    // エラーメッセージを表示
                    MessageBox.Show(ex.Message, $"[{System.Reflection.MethodBase.GetCurrentMethod()?.Name ?? "不明なメソッド"}]エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // 設定Excelワークブックを読み込む
            private static ExcelPackage LoadConfigWorkbook() {
                var configPath = Path.Combine(Environment.CurrentDirectory, "config", "General", "Excel", "ConfigSubstrateInformation.xlsx");
                if (!File.Exists(configPath)) {
                    throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
                }

                try {
                    FileStream fileStreamConfig = new(configPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    return new ExcelPackage(fileStreamConfig);
                } catch (Exception ex) {
                    throw new Exception($"設定ファイルの読み込み中にエラーが発生しました: {ex.Message}", ex);
                }
            }

            // 設定ワークブックから製品に対応する設定を抽出する
            private static (string filePaths, string sheetName) GetSubstrateConfig(ExcelPackage configWorkbook, string substrateModel) {
                var workSheetMain = configWorkbook.Workbook.Worksheets["Sheet1"] ?? throw new Exception("設定ファイルのシールに Sheet1 が見つかりません。");

                // セル検索
                var searchAddressResult = workSheetMain.Cells.FirstOrDefault(x => x.Start.Column == 1 && x.Value?.ToString() == substrateModel)
                    ?? throw new Exception($"Configに型式:[{substrateModel}]が見つかりません。");
                var searchAddressResultRow = searchAddressResult.Start.Row;

                // ディレクトリパスとファイル名を取得(ディレクトリパスが空の場合は、Configの2行目の値を使用する)
                var directoryPath = workSheetMain.Cells[searchAddressResultRow, 4].Value?.ToString()?.Trim('"') ?? workSheetMain.Cells[2, 4].Value?.ToString()?.Trim('"') ?? string.Empty;
                var fileName = workSheetMain.Cells[searchAddressResultRow, 5].Value?.ToString()?.Trim('"') ?? string.Empty;
                if (string.IsNullOrEmpty(fileName)) { throw new Exception($"設定ファイルの型式 {substrateModel} の シート名 が空です。"); }
                var filePaths = Path.Combine(directoryPath, fileName);
                if (!File.Exists(filePaths)) { throw new FileNotFoundException($"指定されたファイルが存在しません: {filePaths}"); }

                var sheetName = workSheetMain.Cells[searchAddressResultRow, 6].Value?.ToString()?.Trim('"') ?? string.Empty;

                return (filePaths, sheetName);
            }

            // Excelファイルを開くメソッド
            private static void OpenExcel(string filePath, string sheetName) {

                // COM Interopを使用してExcelを開く
                Microsoft.Office.Interop.Excel.Application? xlApp = null;
                Microsoft.Office.Interop.Excel.Workbooks? xlBooks = null;
                Microsoft.Office.Interop.Excel.Workbook? xlBook = null;
                Microsoft.Office.Interop.Excel.Sheets? xlSheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? xlSheet = null;

                try {
                    xlApp = new Microsoft.Office.Interop.Excel.Application {
                        Visible = true // Excelウィンドウを表示します。
                    };

                    // ワークブック開く
                    xlBooks = xlApp.Workbooks;
                    xlBook = xlBooks.Open(filePath, ReadOnly: true);

                    // ワークシート選択
                    xlSheets = xlBook.Sheets;
                    xlSheet = string.IsNullOrEmpty(sheetName) ? xlSheets[1] : xlSheets[sheetName];

                    // ワークシート表示
                    xlSheet.Activate();

                } finally {
                    // COMオブジェクトの解放
                    if (xlSheet != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                    }

                    if (xlSheets != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                    }

                    if (xlBook != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                    }

                    if (xlBooks != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                    }

                    if (xlApp != null) {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    }
                }
            }

        }
    }
}
