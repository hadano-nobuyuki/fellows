package com.example.sample01.dao;

import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import com.example.sample01.model.Order;
import com.example.sample01.util.DbUtil;

public class OrderDao {

    private static final String SQL =
            "SELECT 注文日, 注文番号, 注文枝番, 数量, クーポン割引料 "
          + "  FROM 注文 "
          + " WHERE 商品コード = ? "
          + " ORDER BY 注文日 DESC, 注文番号 DESC, 注文枝番 ASC";

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    public List<Order> findByProductCode(String code) throws SQLException {
        List<Order> list = new ArrayList<>();
        try (Connection con = DbUtil.getConnection();
             PreparedStatement ps = con.prepareStatement(SQL)) {
            ps.setString(1, code);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    Date d = rs.getDate("注文日");
                    int coupon = rs.getInt("クーポン割引料");
                    boolean couponNull = rs.wasNull();

                    String couponText = "";
                    boolean minus = false;
                    if (!couponNull) {
                        if (coupon > 0) {
                            couponText = "+" + fmt(coupon);
                        } else {
                            couponText = fmt(coupon);   // 0 または負数(負数は "-" 付き)
                            minus = coupon < 0;
                        }
                    }
                    list.add(new Order(
                            d.toLocalDate().format(DATE_FMT),
                            rs.getString("注文番号").trim(),
                            rs.getInt("注文枝番"),
                            fmt(rs.getInt("数量")),
                            couponText,
                            minus));
                }
            }
        }
        return list;
    }

    private static String fmt(int n) {
        return String.format(Locale.US, "%,d", n);
    }
}
