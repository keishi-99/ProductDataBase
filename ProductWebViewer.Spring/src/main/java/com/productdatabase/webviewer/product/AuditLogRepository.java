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

    // targetType は "PRODUCT" または "SUBSTRATE"
    public void log(String action, String targetType, long targetId, String detail) {
        jdbcTemplate.update(
            "INSERT INTO audit_logs (action, target_type, target_id, detail, created_at) VALUES (?, ?, ?, ?, datetime('now', 'localtime'))",
            action, targetType, targetId, detail
        );
    }

    public List<AuditLog> findAll() {
        // target_typeに応じてproducts/substratesのどちらかとだけ一致するようJOIN条件で振り分ける
        var sql = """
            SELECT a.id, a.action, a.target_type, a.target_id,
                   COALESCE(p.product_name, s.substrate_name) AS target_name,
                   a.detail, a.created_at
            FROM audit_logs a
            LEFT JOIN products p ON a.target_type = 'PRODUCT' AND a.target_id = p.id
            LEFT JOIN substrates s ON a.target_type = 'SUBSTRATE' AND a.target_id = s.id
            ORDER BY a.created_at DESC
            """;
        return jdbcTemplate.query(sql, (rs, rowNum) -> new AuditLog(
            rs.getLong("id"),
            rs.getString("action"),
            rs.getString("target_type"),
            rs.getLong("target_id"),
            rs.getString("target_name"),
            rs.getString("detail"),
            rs.getString("created_at")
        ));
    }
}
