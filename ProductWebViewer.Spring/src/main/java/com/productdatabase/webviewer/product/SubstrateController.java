package com.productdatabase.webviewer.product;

import java.util.List;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
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
    public List<Substrate> list() {
        return substrateRepository.findAll();
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
