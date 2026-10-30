package com.example.sample01.servlet;

import java.io.IOException;
import java.sql.SQLException;

import com.example.sample01.dao.UserDao;

import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

@WebServlet("/login")
public class LoginServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;
    private static final String VIEW = "/WEB-INF/jsp/login.jsp";

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        request.getRequestDispatcher(VIEW).forward(request, response);
    }

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        String username = request.getParameter("username");
        String password = request.getParameter("password");

        boolean ok = false;
        if (username != null && !username.isEmpty() && password != null && !password.isEmpty()) {
            try {
                ok = new UserDao().authenticate(username, password);
            } catch (SQLException e) {
                throw new ServletException("認証処理中にDBエラーが発生しました。", e);
            }
        }

        if (ok) {
            // セッション固定化対策: 既存セッションがあれば ID を更新する
            HttpSession session = request.getSession(false);
            if (session != null) {
                request.changeSessionId();
            } else {
                session = request.getSession(true);
            }
            session.setAttribute("loginUser", username);
            response.sendRedirect(request.getContextPath() + "/products");
        } else {
            request.setAttribute("errorMessage", "ユーザ名またはパスワードが正しくありません。");
            request.getRequestDispatcher(VIEW).forward(request, response);
        }
    }
}
