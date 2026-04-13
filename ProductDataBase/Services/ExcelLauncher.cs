using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ProductDatabase.Services {
    internal static class ExcelLauncher {

        // ExcelのパスをInteropで取得しEXEを直接起動して指定ファイルを開く
        public static void Open(string filePath) {
            Microsoft.Office.Interop.Excel.Application? xlApp = null;
            try {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                var excelFullPath = Path.Combine(xlApp.Path, "EXCEL.EXE");
                xlApp.Quit();
                using var process = Process.Start(new ProcessStartInfo {
                    FileName = excelFullPath,
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = true
                });
            } finally {
                if (xlApp != null) {
                    Marshal.ReleaseComObject(xlApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
    }
}
