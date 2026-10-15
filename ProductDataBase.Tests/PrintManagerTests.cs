using ProductDatabase.Models;
using ProductDatabase.Print;
using System.Drawing;
using System.Drawing.Printing;
using static ProductDatabase.Print.PrintOptions;

namespace ProductDataBase.Tests {
    public class PrintManagerTests {

        private const int Columns = 4;
        private const int Rows = 2;
        private const float Dpi = 96f;
        private const double LabelWidthMm = 20;
        private const double LabelHeightMm = 10;
        private const double MarginMm = 10;
        private const double IntervalMm = 2;

        private static float MmToPx(double mm) => (float)(mm / 25.4 * Dpi);

        // 4列x2行(1ページ8枚)のラベル設定でPrintManagerを初期化する
        private static PrintManager CreateManager(int copiesPerLabel, List<string> serials) {
            var settings = new DocumentPrintSettings {
                LabelPrintSettings = new LabelPrintSettings {
                    LabelWidth = LabelWidthMm,
                    LabelHeight = LabelHeightMm,
                    LabelsPerColumn = Columns,
                    LabelsPerRow = Rows,
                    MarginX = MarginMm,
                    MarginY = MarginMm,
                    IntervalX = IntervalMm,
                    IntervalY = IntervalMm,
                    CopiesPerLabel = copiesPerLabel,
                }
            };
            var manager = new PrintManager { CurrentSerialType = PrintManager.SerialType.Label };
            manager.ProductInitialize(new ProductMaster(), new ProductRegisterWork(), settings, serials);
            return manager;
        }

        // プレビュー描画した1ページ分のBitmapから、枠線が描かれたセル数を数える
        private static int PrintOnePageAndCountLabels(PrintManager manager, out bool hasMore) {
            using var bitmap = new Bitmap(500, 300);
            bitmap.SetResolution(Dpi, Dpi);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.White);

            var bounds = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var args = new PrintPageEventArgs(graphics, bounds, bounds, new PageSettings());
            hasMore = manager.PrintSerialCommon(args, isPreview: true, startLine: 0, PrintManager.SerialType.Label);

            var count = 0;
            for (var y = 0; y < Rows; y++) {
                for (var x = 0; x < Columns; x++) {
                    var originX = (int)(MmToPx(MarginMm) + x * (MmToPx(IntervalMm) + MmToPx(LabelWidthMm)));
                    var originY = (int)(MmToPx(MarginMm) + y * (MmToPx(IntervalMm) + MmToPx(LabelHeightMm)));
                    if (HasDarkPixel(bitmap, originX, originY)) { count++; }
                }
            }
            return count;
        }

        // セル左上の数ピクセル内に枠線(暗い画素)があるかを返す
        private static bool HasDarkPixel(Bitmap bitmap, int originX, int originY) {
            for (var dy = 0; dy < 3; dy++) {
                for (var dx = 0; dx < 3; dx++) {
                    if (bitmap.GetPixel(originX + dx, originY + dy).R < 100) { return true; }
                }
            }
            return false;
        }

        private static List<int> PrintAllPages(PrintManager manager) {
            var labelsPerPage = new List<int>();
            bool hasMore;
            do {
                labelsPerPage.Add(PrintOnePageAndCountLabels(manager, out hasMore));
            } while (hasMore && labelsPerPage.Count < 10);
            return labelsPerPage;
        }

        [Fact]
        public void PrintSerialCommon_シリアルがページをまたぐ_合計枚数が発行枚数どおりになる() {
            var manager = CreateManager(copiesPerLabel: 3, ["S1", "S2", "S3"]);

            var labelsPerPage = PrintAllPages(manager);

            Assert.Equal([8, 1], labelsPerPage);
        }

        [Fact]
        public void PrintSerialCommon_シリアルがページをまたがない_1ページで収まる() {
            var manager = CreateManager(copiesPerLabel: 2, ["S1", "S2", "S3"]);

            var labelsPerPage = PrintAllPages(manager);

            Assert.Equal([6], labelsPerPage);
        }
    }
}
