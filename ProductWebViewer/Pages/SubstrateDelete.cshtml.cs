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
    private readonly ILogger<SubstrateDeleteModel> _logger;

    public SubstrateDeleteModel(SubstrateWriteRepository writeRepo, AuditLogger auditLogger, ILogger<SubstrateDeleteModel> logger) {
        _writeRepo = writeRepo;
        _auditLogger = auditLogger;
        _logger = logger;
    }

    public SubstrateRecord Record { get; set; } = new();
    public string? ErrorMessage { get; private set; }

    public IActionResult OnGet(long id) {
        var record = _writeRepo.GetById(id);
        if (record is null) return NotFound();
        Record = record;
        return Page();
    }

    public IActionResult OnPost(long id) {
        var before = _writeRepo.GetById(id);
        if (before is null) return NotFound();

        if (!_writeRepo.DeleteSubstrate(id)) {
            ErrorMessage = "この基板登録は他の操作で既に削除されています。";
            Record = before;
            return Page();
        }

        AuditLogging.TryLog(_logger, () => _auditLogger.LogSubstrateDelete(before));

        return RedirectToPage("/Index", new { tab = "substrate" });
    }
}
