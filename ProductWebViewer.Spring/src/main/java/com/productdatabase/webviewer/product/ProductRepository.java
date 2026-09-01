package com.productdatabase.webviewer.product;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;

@Repository
public class ProductRepository {

    private static final String SELECT_COLUMNS = """
        id, category_name, product_name, product_model, product_type,
        order_number, product_number, oles_number, quantity, reg_date, comment, created_at
        """;

    private final JdbcTemplate jdbcTemplate;

    public ProductRepository(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public List<Product> findAll(String category, String keyword) {
        var conditions = new ArrayList<String>();
        var params = new ArrayList<Object>();
        conditions.add("is_deleted = 0");

        if (category != null && !category.isBlank()) {
            conditions.add("category_name = ?");
            params.add(category);
        }
        if (keyword != null && !keyword.isBlank()) {
            conditions.add("(product_name LIKE ? OR order_number LIKE ? OR product_number LIKE ?)");
            var like = "%" + keyword + "%";
            params.add(like);
            params.add(like);
            params.add(like);
        }

        var sql = "SELECT " + SELECT_COLUMNS + " FROM products WHERE " + String.join(" AND ", conditions) + " ORDER BY id DESC";
        return jdbcTemplate.query(sql, ProductRepository::mapRow, params.toArray());
    }

    public Product findById(long id) {
        var sql = "SELECT " + SELECT_COLUMNS + " FROM products WHERE id = ? AND is_deleted = 0";
        return jdbcTemplate.query(sql, ProductRepository::mapRow, id).stream().findFirst().orElse(null);
    }

    public List<String> findCategories() {
        return jdbcTemplate.queryForList(
            "SELECT DISTINCT category_name FROM products WHERE is_deleted = 0 ORDER BY category_name",
            String.class
        );
    }

    public boolean update(long id, String orderNumber, String productNumber, String olesNumber, String comment) {
        var sql = """
            UPDATE products
            SET order_number = ?, product_number = ?, oles_number = ?, comment = ?
            WHERE id = ? AND is_deleted = 0
            """;
        var affected = jdbcTemplate.update(sql, orderNumber, productNumber, olesNumber, comment, id);
        return affected > 0;
    }

    public boolean softDelete(long id) {
        var sql = "UPDATE products SET is_deleted = 1, deleted_at = datetime('now', 'localtime') WHERE id = ? AND is_deleted = 0";
        var affected = jdbcTemplate.update(sql, id);
        return affected > 0;
    }

    private static Product mapRow(ResultSet rs, int rowNum) throws SQLException {
        return new Product(
            rs.getLong("id"),
            rs.getString("category_name"),
            rs.getString("product_name"),
            rs.getString("product_model"),
            rs.getString("product_type"),
            rs.getString("order_number"),
            rs.getString("product_number"),
            rs.getString("oles_number"),
            rs.getObject("quantity") != null ? rs.getLong("quantity") : null,
            rs.getString("reg_date"),
            rs.getString("comment"),
            rs.getString("created_at")
        );
    }
}
