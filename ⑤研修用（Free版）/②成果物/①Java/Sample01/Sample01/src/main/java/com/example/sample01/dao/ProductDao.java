package com.example.sample01.dao;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import com.example.sample01.model.Product;
import com.example.sample01.util.DbUtil;

public class ProductDao {

    /** 注文は (注文日, 注文番号) の組で数える。注文が無い商品は 0 件。 */
    private static final String SQL =
            "SELECT p.商品コード, p.商品名, p.単価, p.商品区分, p.関連商品コード, "
          + "       COUNT(DISTINCT (o.注文日, o.注文番号)) FILTER (WHERE o.注文番号 IS NOT NULL) AS 件数 "
          + "  FROM 商品 p "
          + "  LEFT JOIN 注文 o ON p.商品コード = o.商品コード "
          + " GROUP BY p.商品コード, p.商品名, p.単価, p.商品区分, p.関連商品コード "
          + " ORDER BY p.商品コード ASC";

    private static final int NAME_MAX = 25;

    public List<Product> findAll() throws SQLException {
        List<Product> list = new ArrayList<>();
        try (Connection con = DbUtil.getConnection();
             PreparedStatement ps = con.prepareStatement(SQL);
             ResultSet rs = ps.executeQuery()) {
            while (rs.next()) {
                String related = rs.getString("関連商品コード");
                if (related != null) {
                    related = related.trim();
                    if (related.isEmpty()) {
                        related = null;
                    }
                }
                list.add(new Product(
                        rs.getString("商品コード").trim(),
                        truncate(rs.getString("商品名"), NAME_MAX),
                        String.format(Locale.US, "%,d", rs.getInt("単価")),
                        rs.getString("商品区分").trim(),
                        related,
                        rs.getInt("件数")));
            }
        }
        return list;
    }

    /** 25文字まで(サロゲートペアも1文字として数える)。 */
    private static String truncate(String s, int max) {
        if (s == null) {
            return "";
        }
        if (s.codePointCount(0, s.length()) <= max) {
            return s;
        }
        return s.substring(0, s.offsetByCodePoints(0, max));
    }
}
