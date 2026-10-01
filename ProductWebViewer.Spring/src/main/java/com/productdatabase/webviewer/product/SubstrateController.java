package com.productdatabase.webviewer.product;

import java.util.List;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/substrates")
public class SubstrateController {

    private final SubstrateRepository substrateRepository;
    private final AuditLogRepository auditLogRepository;

    public SubstrateController(SubstrateRepository substrateRepository, AuditLogRepository auditLogRepository) {
        this.substrateRepository = substrateRepository;
        this.auditLogRepository = auditLogRepository;
    }

    @GetMapping
    public List<Substrate> list(
        @RequestParam(required = false) String category,
        @RequestParam(required = false) String productName,
        @RequestParam(required = false) String substrateName) {
        return substrateRepository.findAll(category, productName, substrateName);
    }

    @GetMapping("/categories")
    public List<String> categories() {
        return substrateRepository.findCategories();
    }

    // カスケードリストボックス用: カテゴリで絞った製品名一覧
    @GetMapping("/names")
    public List<String> names(@RequestParam(required = false) String category) {
        return substrateRepository.findProductNames(category);
    }

    // カスケードリストボックス用: カテゴリ・製品名で絞った基板名一覧
    @GetMapping("/substrate-names")
    public List<String> substrateNames(
        @RequestParam(required = false) String category,
        @RequestParam(required = false) String productName) {
        return substrateRepository.findSubstrateNames(category, productName);
    }

    @GetMapping("/{id}")
    public ResponseEntity<Substrate> get(@PathVariable long id) {
        var substrate = substrateRepository.findById(id);
        if (substrate == null) return ResponseEntity.notFound().build();
        return ResponseEntity.ok(substrate);
    }

    @DeleteMapping("/{id}")
    public ResponseEntity<Void> delete(@PathVariable long id) {
        var before = substrateRepository.findById(id);
        if (before == null) return ResponseEntity.notFound().build();

        var deleted = substrateRepository.softDelete(id);
        if (!deleted) return ResponseEntity.status(HttpStatus.CONFLICT).build();

        auditLogRepository.log("DELETE", "SUBSTRATE", id, "基板削除: %s (%s)".formatted(before.substrateName(), before.orderNumber()));
        return ResponseEntity.ok().build();
    }
}
