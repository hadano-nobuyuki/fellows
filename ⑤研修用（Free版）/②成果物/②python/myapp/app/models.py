"""既存テーブル(テーブル定義書 myappdb)に対応するモデル。

すべて managed = False。テーブル・列の作成や更新は行わず、SELECT のみ行う。
[廃番商品] テーブルは本アプリでは使わないためモデルを作らない。
"""
from django.db import models


class Product(models.Model):
    商品コード = models.CharField(max_length=5, primary_key=True, db_column='商品コード')
    商品名 = models.CharField(max_length=100, db_column='商品名')
    単価 = models.IntegerField(db_column='単価')
    商品区分 = models.CharField(max_length=1, db_column='商品区分')
    関連商品コード = models.CharField(max_length=5, null=True, blank=True, db_column='関連商品コード')

    class Meta:
        managed = False
        db_table = '商品'


class Order(models.Model):
    pk = models.CompositePrimaryKey('注文日', '注文番号', '注文枝番')
    注文日 = models.DateField(db_column='注文日')
    注文番号 = models.CharField(max_length=12, db_column='注文番号')
    注文枝番 = models.IntegerField(db_column='注文枝番')
    商品コード = models.CharField(max_length=5, db_column='商品コード')
    数量 = models.IntegerField(db_column='数量')
    クーポン割引料 = models.IntegerField(null=True, blank=True, db_column='クーポン割引料')

    class Meta:
        managed = False
        db_table = '注文'


class UserInfo(models.Model):
    ユーザーid = models.AutoField(primary_key=True, db_column='ユーザーid')
    eメール = models.CharField(max_length=255, db_column='eメール')
    パスワード = models.CharField(max_length=255, db_column='パスワード')

    class Meta:
        managed = False
        db_table = 'ユーザー情報'
