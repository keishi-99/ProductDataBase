using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Data;
using ProductWebViewer.Models;

namespace ProductWebViewer.Pages;

public class OperationLogModel : PageModel {
    private readonly LogRecordRepository _logRepo;

    public OperationLogModel(LogRecordRepository logRepo) {
        _logRepo = logRepo;
    }

    [BindProperty(SupportsGet = true)] public string? YearMonth { get; set; }
    [BindProperty(SupportsGet = true)] public string? OperationType { get; set; }
    [BindProperty(SupportsGet = true)] public string? Keyword { get; set; }

    public IReadOnlyList<string> AvailableMonths { get; private set; } = [];
    public IReadOnlyList<string> OperationTypes { get; private set; } = [];
    public IReadOnlyList<LogEntry> Entries { get; private set; } = [];

    public void OnGet() {
        AvailableMonths = _logRepo.GetAvailableMonths();
        if (string.IsNullOrEmpty(YearMonth) || !AvailableMonths.Contains(YearMonth))
            YearMonth = AvailableMonths.FirstOrDefault();

        if (YearMonth is null) return;

        var monthEntries = _logRepo.GetLogEntries(YearMonth);
        OperationTypes = monthEntries.Select(e => e.OperationType).Distinct().OrderBy(s => s).ToList();

        IEnumerable<LogEntry> filtered = monthEntries;
        if (!string.IsNullOrEmpty(OperationType))
            filtered = filtered.Where(e => e.OperationType == OperationType);
        if (!string.IsNullOrWhiteSpace(Keyword)) {
            var keyword = Keyword.Trim();
            filtered = filtered.Where(e =>
                e.OperationType.Contains(keyword) || e.Category.Contains(keyword) || e.Id.Contains(keyword) ||
                e.OrderNumber.Contains(keyword) || e.ProductNumber.Contains(keyword) || e.OLesNumber.Contains(keyword) ||
                e.ProductName.Contains(keyword) || e.ProductType.Contains(keyword) || e.ProductModel.Contains(keyword) ||
                e.Quantity.Contains(keyword) || e.SerialFirst.Contains(keyword) || e.SerialLast.Contains(keyword) ||
                e.Revision.Contains(keyword) || e.RegDate.Contains(keyword) || e.PersonInfo.Contains(keyword) ||
                e.Comment.Contains(keyword));
        }

        Entries = filtered.OrderByDescending(e => e.Timestamp).ToList();
    }

    // 操作種別に応じた行の背景色クラスを返す（メインアプリのLogViewerWindow.GetRowColorに準拠）
    public static string RowCssClass(string operationType) {
        if (operationType.Contains("削除")) return "table-danger";
        if (operationType.Contains("編集")) return "table-warning";
        if (operationType.Contains("製品登録")) return "table-info";
        if (operationType.Contains("基板登録")) return "table-success";
        return string.Empty;
    }
}
