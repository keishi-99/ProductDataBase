package com.productdatabase.webviewer.product;

import java.util.List;
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

    public List<AuditLog> findAll() {
        var sql = """
            SELECT a.id, a.action, a.product_id, p.product_name, a.detail, a.created_at
            FROM audit_logs a
            LEFT JOIN products p ON a.product_id = p.id
            ORDER BY a.created_at DESC
            """;
        return jdbcTemplate.query(sql, (rs, rowNum) -> new AuditLog(
            rs.getLong("id"),
            rs.getString("action"),
            rs.getLong("product_id"),
            rs.getString("product_name"),
            rs.getString("detail"),
            rs.getString("created_at")
        ));
    }
}
