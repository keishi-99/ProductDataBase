package com.productdatabase.webviewer.product;

public record Product(
    long id,
    String categoryName,
    String productName,
    String productModel,
    String productType,
    String orderNumber,
    String productNumber,
    String olesNumber,
    Long quantity,
    String regDate,
    String comment,
    String createdAt
) {}
