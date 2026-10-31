import hashlib
from functools import wraps

from django.db import connection
from django.shortcuts import redirect, render
from django.views.decorators.http import require_http_methods, require_POST

from .models import Order, UserInfo

SESSION_KEY = 'loginUser'
LOGIN_ERROR = 'ユーザ名またはパスワードが正しくありません。'


def login_required_session(view):
    """セッションに "loginUser" が無ければログイン画面へリダイレクトする。"""
    @wraps(view)
    def wrapper(request, *args, **kwargs):
        if not request.session.get(SESSION_KEY):
            return redirect('login')
        return view(request, *args, **kwargs)
    return wrapper


def index(request):
    return redirect('login')


@require_http_methods(['GET', 'POST'])
def login_view(request):
    if request.method == 'GET':
        return render(request, 'app/login.html')

    username = request.POST.get('username', '')
    password = request.POST.get('password', '')
    # SHA-256(UTF-8) の小文字16進数64桁。ソルトなし(Java 版と同方式)。
    hashed = hashlib.sha256(password.encode('utf-8')).hexdigest()

    if UserInfo.objects.filter(eメール=username, パスワード=hashed).exists():
        request.session.cycle_key()
        request.session[SESSION_KEY] = username
        return redirect('products')

    return render(request, 'app/login.html', {'error': LOGIN_ERROR})


@require_POST
def logout_view(request):
    request.session.flush()
    return redirect('login')


PRODUCT_SQL = """
    SELECT p.商品コード,
           p.商品名,
           p.単価,
           p.商品区分,
           p.関連商品コード,
           COUNT(DISTINCT (o.注文日, o.注文番号))
               FILTER (WHERE o.注文番号 IS NOT NULL) AS 件数
      FROM 商品 p
      LEFT JOIN 注文 o ON o.商品コード = p.商品コード
     GROUP BY p.商品コード, p.商品名, p.単価, p.商品区分, p.関連商品コード
     ORDER BY p.商品コード
"""


@login_required_session
def product_list(request):
    with connection.cursor() as cursor:
        cursor.execute(PRODUCT_SQL)
        rows = cursor.fetchall()

    products = [
        {
            'code': r[0],
            'name': r[1],
            'price': r[2],
            'category': r[3],
            'related_code': r[4],
            'order_count': r[5],
        }
        for r in rows
    ]
    return render(request, 'app/product_list.html', {'products': products})


@login_required_session
def order_detail(request):
    code = request.GET.get('code', '').strip()
    if not code:
        return redirect('products')

    orders = (
        Order.objects.filter(商品コード=code)
        .order_by('-注文日', '-注文番号', '注文枝番')
    )
    rows = [
        {
            'date': o.注文日,
            'number': o.注文番号,
            'branch': o.注文枝番,
            'quantity': o.数量,
            'coupon': o.クーポン割引料,
        }
        for o in orders
    ]
    return render(request, 'app/order_detail.html', {'code': code, 'orders': rows})
