package com.productdatabase.webviewer.product;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

@Repository
public class AuditLogRepository {

    private final JdbcTemplate jdbcTemplate;

    public AuditLogRepository(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public void log(String action, long productId, String detail) {
        jdbcTemplate.update(
            "INSERT INTO audit_logs (action, product_id, detail, created_at) VALUES (?, ?, ?, datetime('now', 'localtime'))",
            action, productId, detail
        );
    }
}
