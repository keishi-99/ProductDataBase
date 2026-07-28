using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ProductWebViewer.Data;
using ProductWebViewer.Models;

namespace ProductWebViewer.Pages;

[Authorize]
public class SubstrateDeleteModel : PageModel {
    private readonly SubstrateWriteRepository _writeRepo;
    private readonly AuditLogger _auditLogger;

    public SubstrateDeleteModel(SubstrateWriteRepository writeRepo, AuditLogger auditLogger) {
        _writeRepo = writeRepo;
        _auditLogger = auditLogger;
    }

    public SubstrateRecord Record { get; set; } = new();

    public IActionResult OnGet(long id) {
        var record = _writeRepo.GetById(id);
        if (record is null) return NotFound();
        Record = record;
        return Page();
    }

    public IActionResult OnPost(long id) {
        var before = _writeRepo.GetById(id);
        if (before is null) return NotFound();

        _writeRepo.DeleteSubstrate(id);
        _auditLogger.LogSubstrateDelete(User.Identity?.Name ?? "管理者", before);

        return RedirectToPage("/Index", new { tab = "substrate" });
    }
}
