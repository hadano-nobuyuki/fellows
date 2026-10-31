<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>商品一覧 | ClaudeCode 研修用プログラム</title>
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
    <h1>商品一覧画面</h1>
  </div>

  <div class="table-card">
    <table>
      <thead>
        <tr>
          <th>商品コード</th>
          <th>商品名</th>
          <th class="num">単価</th>
          <th>商品区分</th>
          <th>関連商品コード</th>
          <th>注文状況</th>
        </tr>
      </thead>
      <tbody>
        <c:forEach var="p" items="${products}">
          <tr>
            <td class="code"><c:out value="${p.code}"/></td>
            <td class="name"><c:out value="${p.name}"/></td>
            <td class="num"><c:out value="${p.unitPrice}"/></td>
            <td><span class="badge"><c:out value="${p.category}"/></span></td>
            <td>
              <c:choose>
                <c:when test="${not empty p.relatedCode}"><span class="code"><c:out value="${p.relatedCode}"/></span></c:when>
                <c:otherwise><span class="muted">―</span></c:otherwise>
              </c:choose>
            </td>
            <td>
              <c:choose>
                <c:when test="${p.orderCount > 0}">
                  <c:url var="orderUrl" value="/orders">
                    <c:param name="code" value="${p.code}"/>
                  </c:url>
                  <a class="order-link" href="${orderUrl}">注文(<c:out value="${p.orderCount}"/> 件)</a>
                </c:when>
                <c:otherwise><span class="muted">注文なし</span></c:otherwise>
              </c:choose>
            </td>
          </tr>
        </c:forEach>
      </tbody>
    </table>
  </div>
</main>
</body>
</html>
