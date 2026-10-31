from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('login', views.login_view, name='login'),
    path('logout', views.logout_view, name='logout'),
    path('products', views.product_list, name='products'),
    path('orders', views.order_detail, name='orders'),
]
