package com.example.sample01.dao;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import com.example.sample01.util.DbUtil;
import com.example.sample01.util.HashUtil;

public class UserDao {

    private static final String SQL =
            "SELECT 1 FROM ユーザー情報 WHERE eメール = ? AND パスワード = ?";

    /** ユーザ名(eメール)とパスワード(SHA-256 ハッシュ)が一致するユーザーがいれば true。 */
    public boolean authenticate(String username, String password) throws SQLException {
        String hash = HashUtil.sha256Hex(password);
        try (Connection con = DbUtil.getConnection();
             PreparedStatement ps = con.prepareStatement(SQL)) {
            ps.setString(1, username);
            ps.setString(2, hash);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }
}
