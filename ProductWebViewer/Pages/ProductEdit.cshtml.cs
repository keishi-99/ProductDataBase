using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Data;
using ProductWebViewer.Models;

namespace ProductWebViewer.Pages;

[Authorize]
public class ProductEditModel : PageModel {
    private readonly ProductWriteRepository _writeRepo;
    private readonly AuditLogger _auditLogger;
    private readonly ILogger<ProductEditModel> _logger;

    public ProductEditModel(ProductWriteRepository writeRepo, AuditLogger auditLogger, ILogger<ProductEditModel> logger) {
        _writeRepo = writeRepo;
        _auditLogger = auditLogger;
        _logger = logger;
    }

    [BindProperty] public ProductRecord Record { get; set; } = new();
    public string? ErrorMessage { get; private set; }

    public IActionResult OnGet(long id) {
        var record = _writeRepo.GetById(id);
        if (record is null) return NotFound();
        Record = record;
        return Page();
    }

    public IActionResult OnPostAsync(long id) {
        var before = _writeRepo.GetById(id);
        if (before is null) return NotFound();

        // 他の操作で既に削除されている場合は更新せず競合として扱う
        if (!_writeRepo.UpdateProduct(id, Record.OrderNumber, Record.ProductNumber, Record.OLesNumber, Record.Comment)) {
            ErrorMessage = "この製品登録は他の操作で既に削除されているため、更新できませんでした。";
            Record = before;
            return Page();
        }

        AuditLogging.TryLog(_logger, () => _auditLogger.LogProductEdit(before, Record.OrderNumber, Record.ProductNumber, Record.OLesNumber, Record.Comment));

        return RedirectToPage("/Index");
    }

    public IActionResult OnPostDelete(long id) {
        var before = _writeRepo.GetById(id);
        if (before is null) return NotFound();

        var result = _writeRepo.DeleteProduct(id);
        if (!result.Success) {
            ErrorMessage = "この製品登録は他の操作で既に削除されています。";
            Record = before;
            return Page();
        }

        AuditLogging.TryLog(_logger, () => {
            _auditLogger.LogProductDelete(before);
            _auditLogger.LogProductSubstrateDelete(result.DeletedSubstrates, before.CategoryName);
            _auditLogger.LogProductSerialDelete(result.DeletedSerials, before.CategoryName);
        });

        return RedirectToPage("/Index");
    }
}
