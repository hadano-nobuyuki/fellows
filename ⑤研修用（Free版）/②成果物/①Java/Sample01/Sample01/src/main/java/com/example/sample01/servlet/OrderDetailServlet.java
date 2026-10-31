package com.example.sample01.servlet;

import java.io.IOException;
import java.sql.SQLException;

import com.example.sample01.dao.OrderDao;

import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

@WebServlet("/orders")
public class OrderDetailServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("loginUser") == null) {
            response.sendRedirect(request.getContextPath() + "/login");
            return;
        }
        String code = request.getParameter("code");
        if (code == null || code.isBlank()) {
            response.sendRedirect(request.getContextPath() + "/products");
            return;
        }
        code = code.trim();
        try {
            request.setAttribute("orders", new OrderDao().findByProductCode(code));
        } catch (SQLException e) {
            throw new ServletException("注文情報の取得中にDBエラーが発生しました。", e);
        }
        request.setAttribute("code", code);
        request.getRequestDispatcher("/WEB-INF/jsp/order_detail.jsp").forward(request, response);
    }
}
