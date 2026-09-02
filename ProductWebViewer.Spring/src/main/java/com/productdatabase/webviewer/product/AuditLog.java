package com.productdatabase.webviewer.product;

public record AuditLog(
    long id,
    String action,
    long productId,
    String productName,
    String detail,
    String createdAt
) {}
