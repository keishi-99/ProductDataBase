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

    public ProductEditModel(ProductWriteRepository writeRepo, AuditLogger auditLogger) {
        _writeRepo = writeRepo;
        _auditLogger = auditLogger;
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

        _writeRepo.UpdateProduct(id, Record.OrderNumber, Record.ProductNumber, Record.OLesNumber, Record.RegDate, Record.Revision, Record.Comment);
        _auditLogger.LogProductEdit(User.Identity?.Name ?? "管理者", before, Record);

        return RedirectToPage("/Index");
    }

    public IActionResult OnPostDelete(long id) {
        var before = _writeRepo.GetById(id);
        if (before is null) return NotFound();

        _writeRepo.DeleteProduct(id);
        _auditLogger.LogProductDelete(User.Identity?.Name ?? "管理者", before);

        return RedirectToPage("/Index");
    }
}
