package com.productdatabase.webviewer.product;

public record AuditLog(
    long id,
    String action,
    String targetType,
    long targetId,
    String targetName,
    String detail,
    String createdAt
) {}
