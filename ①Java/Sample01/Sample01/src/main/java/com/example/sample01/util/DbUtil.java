package com.example.sample01.util;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

/** db.properties(クラスパス)から接続情報を読み込み、Connection を返す。 */
public final class DbUtil {

    private static final Properties PROPS = new Properties();

    static {
        try (InputStream in = DbUtil.class.getClassLoader().getResourceAsStream("db.properties")) {
            if (in == null) {
                throw new IllegalStateException("db.properties がクラスパスに見つかりません。");
            }
            PROPS.load(in);
        } catch (IOException e) {
            throw new IllegalStateException("db.properties の読み込みに失敗しました。", e);
        }
    }

    private DbUtil() {
    }

    public static Connection getConnection() throws SQLException {
        try {
            // WAR 同梱のドライバは自動登録されないことがあるため明示的にロードする
            Class.forName(PROPS.getProperty("db.driver"));
        } catch (ClassNotFoundException e) {
            throw new SQLException("JDBC ドライバが見つかりません: " + PROPS.getProperty("db.driver"), e);
        }
        return DriverManager.getConnection(
                PROPS.getProperty("db.url"),
                PROPS.getProperty("db.user"),
                PROPS.getProperty("db.password"));
    }
}
