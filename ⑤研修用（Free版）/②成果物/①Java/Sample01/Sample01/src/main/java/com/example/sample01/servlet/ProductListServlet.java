package com.example.sample01.servlet;

import java.io.IOException;
import java.sql.SQLException;

import com.example.sample01.dao.ProductDao;

import jakarta.servlet.ServletException;
import jakarta.servlet.annotation.WebServlet;
import jakarta.servlet.http.HttpServlet;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;

@WebServlet("/products")
public class ProductListServlet extends HttpServlet {

    private static final long serialVersionUID = 1L;

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        HttpSession session = request.getSession(false);
        if (session == null || session.getAttribute("loginUser") == null) {
            response.sendRedirect(request.getContextPath() + "/login");
            return;
        }
        try {
            request.setAttribute("products", new ProductDao().findAll());
        } catch (SQLException e) {
            throw new ServletException("商品一覧の取得中にDBエラーが発生しました。", e);
        }
        request.getRequestDispatcher("/WEB-INF/jsp/product_list.jsp").forward(request, response);
    }
}
