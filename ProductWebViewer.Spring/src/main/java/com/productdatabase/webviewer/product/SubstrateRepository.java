package com.productdatabase.webviewer.product;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

@Repository
public class SubstrateRepository {

    private static final String SELECT_COLUMNS = """
        id, category_name, product_name, substrate_name, substrate_model,
        order_number, substrate_number, increase, decrease, defect, person_info, reg_date, comment, created_at
        """;

    private final JdbcTemplate jdbcTemplate;

    public SubstrateRepository(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public List<Substrate> findAll() {
        var sql = "SELECT " + SELECT_COLUMNS + " FROM substrates WHERE is_deleted = 0 ORDER BY id DESC";
        return jdbcTemplate.query(sql, SubstrateRepository::mapRow);
    }

    public Substrate findById(long id) {
        var sql = "SELECT " + SELECT_COLUMNS + " FROM substrates WHERE id = ? AND is_deleted = 0";
        return jdbcTemplate.query(sql, SubstrateRepository::mapRow, id).stream().findFirst().orElse(null);
    }

    public boolean softDelete(long id) {
        var sql = "UPDATE substrates SET is_deleted = 1, deleted_at = datetime('now', 'localtime') WHERE id = ? AND is_deleted = 0";
        var affected = jdbcTemplate.update(sql, id);
        return affected > 0;
    }

    private static Substrate mapRow(ResultSet rs, int rowNum) throws SQLException {
        return new Substrate(
            rs.getLong("id"),
            rs.getString("category_name"),
            rs.getString("product_name"),
            rs.getString("substrate_name"),
            rs.getString("substrate_model"),
            rs.getString("order_number"),
            rs.getString("substrate_number"),
            rs.getObject("increase") != null ? rs.getLong("increase") : null,
            rs.getObject("decrease") != null ? rs.getLong("decrease") : null,
            rs.getObject("defect") != null ? rs.getLong("defect") : null,
            rs.getString("person_info"),
            rs.getString("reg_date"),
            rs.getString("comment"),
            rs.getString("created_at")
        );
    }
}
