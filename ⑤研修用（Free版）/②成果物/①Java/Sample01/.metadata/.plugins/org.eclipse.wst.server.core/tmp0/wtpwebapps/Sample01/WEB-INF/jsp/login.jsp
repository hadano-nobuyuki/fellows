<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ログイン | ClaudeCode 研修用プログラム</title>
  <link rel="stylesheet" href="${pageContext.request.contextPath}/style.css">
</head>
<body>
<main class="login-page">
  <div class="login-card">
    <img class="logo" src="${pageContext.request.contextPath}/logo.svg" alt="ロゴ">
    <p class="app-name">ClaudeCode 研修用プログラム</p>
    <h1>ログイン画面</h1>

    <c:if test="${not empty errorMessage}">
      <p class="error-message"><c:out value="${errorMessage}"/></p>
    </c:if>

    <form method="post" action="${pageContext.request.contextPath}/login">
      <div class="field">
        <label for="username">ユーザ名</label>
        <input type="text" id="username" name="username" autocomplete="username" required autofocus>
      </div>
      <div class="field">
        <label for="password">パスワード</label>
        <input type="password" id="password" name="password" autocomplete="current-password" required>
      </div>
      <button type="submit" class="btn btn-primary">ログイン</button>
    </form>
  </div>
</main>
</body>
</html>
