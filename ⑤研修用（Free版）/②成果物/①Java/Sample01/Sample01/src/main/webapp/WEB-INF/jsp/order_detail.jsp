<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>注文情報 | ClaudeCode 研修用プログラム</title>
  <link rel="stylesheet" href="${pageContext.request.contextPath}/style.css">
</head>
<body class="list-page">
<header class="app-header">
  <img class="logo" src="${pageContext.request.contextPath}/logo.svg" alt="ロゴ">
  <span class="app-name">ClaudeCode 研修用プログラム</span>
  <span class="spacer"></span>
  <form method="post" action="${pageContext.request.contextPath}/logout">
    <button type="submit" class="btn btn-outline">ログアウト</button>
  </form>
</header>

<main class="content">
  <div class="page-title">
    <h1>注文情報詳細画面</h1>
    <span class="sub">商品コード:<strong class="code"><c:out value="${code}"/></strong></span>
  </div>

  <div class="table-card">
    <table>
      <thead>
        <tr>
          <th>注文日</th>
          <th>注文番号</th>
          <th class="num">注文枝番</th>
          <th class="num">数量</th>
          <th class="num">クーポン割引料</th>
        </tr>
      </thead>
      <tbody>
        <c:forEach var="o" items="${orders}">
          <tr>
            <td><c:out value="${o.orderDate}"/></td>
            <td class="code"><c:out value="${o.orderNo}"/></td>
            <td class="num"><c:out value="${o.branchNo}"/></td>
            <td class="num"><c:out value="${o.quantity}"/></td>
            <td class="num${o.couponMinus ? ' minus' : ''}"><c:out value="${o.coupon}"/></td>
          </tr>
        </c:forEach>
      </tbody>
    </table>
  </div>

  <div class="page-footer">
    <a class="back-link" href="${pageContext.request.contextPath}/products">&larr; 商品一覧に戻る</a>
  </div>
</main>
</body>
</html>
