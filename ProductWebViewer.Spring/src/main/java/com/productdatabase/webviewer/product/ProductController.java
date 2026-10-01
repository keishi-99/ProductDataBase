package com.productdatabase.webviewer.product;

import java.util.List;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/products")
public class ProductController {

    private final ProductRepository productRepository;
    private final AuditLogRepository auditLogRepository;

    public ProductController(ProductRepository productRepository, AuditLogRepository auditLogRepository) {
        this.productRepository = productRepository;
        this.auditLogRepository = auditLogRepository;
    }

    @GetMapping
    public List<Product> list(
        @RequestParam(required = false) String category,
        @RequestParam(required = false) String productName,
        @RequestParam(required = false) String productType,
        @RequestParam(required = false) String keyword) {
        return productRepository.findAll(category, productName, productType, keyword);
    }

    @GetMapping("/categories")
    public List<String> categories() {
        return productRepository.findCategories();
    }

    // カスケードリストボックス用: カテゴリで絞った製品名一覧
    @GetMapping("/names")
    public List<String> names(@RequestParam(required = false) String category) {
        return productRepository.findProductNames(category);
    }

    // カスケードリストボックス用: カテゴリ・製品名で絞った種別一覧
    @GetMapping("/types")
    public List<String> types(
        @RequestParam(required = false) String category,
        @RequestParam(required = false) String productName) {
        return productRepository.findProductTypes(category, productName);
    }

    @GetMapping("/{id}")
    public ResponseEntity<Product> get(@PathVariable long id) {
        var product = productRepository.findById(id);
        if (product == null) return ResponseEntity.notFound().build();
        return ResponseEntity.ok(product);
    }

    @PutMapping("/{id}")
    public ResponseEntity<Void> update(@PathVariable long id, @RequestBody ProductEditRequest request) {
        var before = productRepository.findById(id);
        if (before == null) return ResponseEntity.notFound().build();

        var updated = productRepository.update(id, request.orderNumber(), request.productNumber(), request.olesNumber(), request.comment());
        if (!updated) return ResponseEntity.status(HttpStatus.CONFLICT).build();

        auditLogRepository.log("EDIT", "PRODUCT", id, "注文番号:%s→%s, 製造番号:%s→%s, OLES番号:%s→%s, コメント:%s→%s".formatted(
            before.orderNumber(), request.orderNumber(),
            before.productNumber(), request.productNumber(),
            before.olesNumber(), request.olesNumber(),
            before.comment(), request.comment()
        ));
        return ResponseEntity.ok().build();
    }

    @DeleteMapping("/{id}")
    public ResponseEntity<Void> delete(@PathVariable long id) {
        var before = productRepository.findById(id);
        if (before == null) return ResponseEntity.notFound().build();

        var deleted = productRepository.softDelete(id);
        if (!deleted) return ResponseEntity.status(HttpStatus.CONFLICT).build();

        auditLogRepository.log("DELETE", "PRODUCT", id, "製品削除: %s (%s)".formatted(before.productName(), before.orderNumber()));
        return ResponseEntity.ok().build();
    }

    record ProductEditRequest(String orderNumber, String productNumber, String olesNumber, String comment) {}
}
