package com.productdatabase.webviewer.product;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
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

    public List<Substrate> findAll(String category, String productName, String substrateName) {
        var conditions = new ArrayList<String>();
        var params = new ArrayList<Object>();
        conditions.add("is_deleted = 0");

        if (category != null && !category.isBlank()) {
            conditions.add("category_name = ?");
            params.add(category);
        }
        if (productName != null && !productName.isBlank()) {
            conditions.add("product_name = ?");
            params.add(productName);
        }
        if (substrateName != null && !substrateName.isBlank()) {
            conditions.add("substrate_name = ?");
            params.add(substrateName);
        }

        var sql = "SELECT " + SELECT_COLUMNS + " FROM substrates WHERE " + String.join(" AND ", conditions) + " ORDER BY id DESC";
        return jdbcTemplate.query(sql, SubstrateRepository::mapRow, params.toArray());
    }

    public List<String> findCategories() {
        return jdbcTemplate.queryForList(
            "SELECT DISTINCT category_name FROM substrates WHERE is_deleted = 0 ORDER BY category_name",
            String.class
        );
    }

    // カスケードリストボックス用: カテゴリで絞った製品名一覧(未指定なら全件)
    public List<String> findProductNames(String category) {
        var conditions = new ArrayList<String>();
        var params = new ArrayList<Object>();
        conditions.add("is_deleted = 0");
        if (category != null && !category.isBlank()) {
            conditions.add("category_name = ?");
            params.add(category);
        }
        var sql = "SELECT DISTINCT product_name FROM substrates WHERE " + String.join(" AND ", conditions) + " ORDER BY product_name";
        return jdbcTemplate.queryForList(sql, String.class, params.toArray());
    }

    // カスケードリストボックス用: カテゴリ・製品名で絞った基板名一覧(未指定なら全件)
    public List<String> findSubstrateNames(String category, String productName) {
        var conditions = new ArrayList<String>();
        var params = new ArrayList<Object>();
        conditions.add("is_deleted = 0");
        if (category != null && !category.isBlank()) {
            conditions.add("category_name = ?");
            params.add(category);
        }
        if (productName != null && !productName.isBlank()) {
            conditions.add("product_name = ?");
            params.add(productName);
        }
        var sql = "SELECT DISTINCT substrate_name FROM substrates WHERE " + String.join(" AND ", conditions) + " ORDER BY substrate_name";
        return jdbcTemplate.queryForList(sql, String.class, params.toArray());
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
