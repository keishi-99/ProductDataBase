using ClosedXML.Excel;

namespace ProductDatabase.Excel {
    // Excelサービスクラス間で共有するユーティリティメソッド
    internal static class ExcelHelper {

        // 保存前に全シートのビューを改ページプレビューに強制設定する
        // ClosedXML が view 属性を正しく読み取れないバグへの暫定対応
        // バグ修正後は CaptureSheetViews / RestoreSheetViews に切り替える
        public static void ForcePageBreakPreview(XLWorkbook workbook) {
            foreach (var ws in workbook.Worksheets) {
                ws.SheetView.View = XLSheetViewOptions.PageBreakPreview;
            }
        }

        // COMオブジェクトを安全に解放する
        public static void ReleaseComObject(object? comObj) {
            if (comObj is not null) {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObj);
            }
        }
    }
}
