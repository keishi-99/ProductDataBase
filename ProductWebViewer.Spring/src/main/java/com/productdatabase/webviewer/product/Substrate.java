package com.productdatabase.webviewer.product;

public record Substrate(
    long id,
    String categoryName,
    String productName,
    String substrateName,
    String substrateModel,
    String orderNumber,
    String substrateNumber,
    Long increase,
    Long decrease,
    Long defect,
    String personInfo,
    String regDate,
    String comment,
    String createdAt
) {}
