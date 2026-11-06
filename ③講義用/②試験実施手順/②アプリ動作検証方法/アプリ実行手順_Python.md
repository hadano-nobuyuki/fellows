# アプリ実行手順（Python）

Python アプリの生成から動作確認までの手順です。

## 起動方法

1. **プログラム用資材をアップロード** する
2. **Claude でプログラミング** する
3. Claude が、作成したプログラムを **zip で圧縮** する
4. ダウンロード後、**解凍** する

### 参考: cmd と PowerShell で使えるコマンドの比較

| やりたいこと | cmd | PowerShell |
|---|---|---|
| ファイルの中身を表示 | `type` | `Get-Content` |
| 環境変数を設定 | `set DB_HOST=localhost` | `$env:DB_HOST="localhost"` |
| venv を有効化 | `venv\Scripts\activate` | `.\venv\Scripts\Activate.ps1` |
| ファイルに書き出す | `> file`(そのまま使える) | `Out-File -Encoding ascii`(`>` は UTF-16 になる) |

以降は cmd で説明します。

### 1. 仮想環境の作成

ターミナルを起動し、プロジェクトのディレクトリ(myapp フォルダ)に移動します。

- 初回のみ: 仮想環境を作成する

  ```
  python -m venv venv
  ```

- 毎回実施: 仮想環境を有効化する

  ```
  venv\Scripts\activate.bat
  ```

- 作業終了時に実施: 仮想環境を終了(無効化)する

  ```
  deactivate
  ```

### 2. インストール

```
pip install -r requirements.txt
```

### 3. 環境変数の設定

この設定を忘れると、`KeyError: 'DJANGO_SECRET_KEY'` になります。設定した値は、その画面(ターミナル)を閉じると消えるので、**新しい画面を開くたびに設定**してください。

```
set DJANGO_SECRET_KEY=local-dev-key-change-me-0123456789
set DB_HOST=localhost
set DB_PORT=5432
set DB_NAME=myappdb
set DB_USER=myapp
set DB_PASSWORD=1234
```

### 4. マイグレーション

初回、または環境が変わったときに実行します。

```
python manage.py migrate
```

### 5. Django 製の内蔵 Web サーバーを起動

```
python manage.py runserver
```

### 6. URL へアクセス

- http://127.0.0.1:8000/

| 項目 | 値 |
|---|---|
| ユーザ名 | `user@example.com` |
| パスワード | `1234` |

### 7. requirements.txt の生成

環境を作ったときに実行します。

```
python -m pip freeze > requirements.txt
```

中身の確認:

```
type requirements.txt
```
