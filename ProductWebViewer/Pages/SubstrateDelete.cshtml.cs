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

        // 監査ログの書き込み失敗（ディスク容量不足等）で削除自体が失敗扱いにならないようにする
        try {
            _auditLogger.LogSubstrateDelete(before);
        } catch (Exception ex) {
            _logger.LogError(ex, "監査ログの記録に失敗しました。");
        }

        return RedirectToPage("/Index", new { tab = "substrate" });
    }
}
