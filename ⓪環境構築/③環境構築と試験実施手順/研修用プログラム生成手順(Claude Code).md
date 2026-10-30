# 研修用プログラム生成手順

Oct 25, 2026 · @波多野伸幸

## 1. 目的と全体像

受講生は、Claude(Free プランのチャット)に「指示書」と画面 HTML・テーブル定義書を添付して、Java と Python のアプリを生成し、同じ PostgreSQL のユーザーで相互にログインできることを確認します。VBA は生成に使用量がかかるため、講師が自分のアカウントで生成した Sample01.xlsm を素材として渡し、講義では説明のみとします。題材は検証で使った「やくそくん」とは別の、商品・注文を扱う新しいアプリ(プロジェクト名 Sample01)です。

| フェーズ | 場所 | 生成・確認するもの |
| --- | --- | --- |
| ① スタンドアロン | 受講生のパソコン | A5M2 で作ったローカル PostgreSQL(myappdb)に、Eclipse(Java)、VSCode(Python / Django)から接続し、ログイン認証と情報表示を確認 |
| ② WEB | AWS(講師が配布する AMI) | ①で作った Java を Tomcat(`http://IP/Sample01/`)、Python を Django(`http://IP:8081/`)にデプロイして動作確認 |

アプリの機能は「ログイン → 商品一覧(商品ごとの注文件数) → 注文情報詳細」と、各画面のログアウトです。ユーザー登録画面は作りません。DB の初期データにある 1 ユーザーで、Java・Python・VBA のどれにもログインできること(相互ログイン)が、この研修の到達点です。

**全体の流れ**

1. 事前準備(ソフトのインストール、教材フォルダの受け取り)
2. A5M2 で myappdb を作り、`テーブル定義書（myappdb）.sql` を実行してテーブルと初期データを作る
3. Claude のチャットに指示書と資材を添付し、Java → Python の順に生成して各自で動かす
4. 同じユーザーで Java・Python(・講師の VBA)にログインして相互ログインを確認する
5. フェーズ②で AWS にデプロイする(別手順書)

## 2. 事前準備

| 用途 | ソフト | 注意点 |
| --- | --- | --- |
| 生成 | Claude(Free プランのチャット) | ログイン済みであること。Free は 5 時間単位の使用量制限があるので、1 言語ずつ生成する |
| DB | PostgreSQL 18.4(ローカル)、A5M2 | DB 名 myappdb、ユーザー myapp |
| Java | Eclipse、Maven(Eclipse 同梱の m2e でよい)、Tomcat 11、JDK 25 | Maven の war プロジェクトとして取り込む。サーバーは Tomcat 11 / Corretto 25 |
| Python | VSCode、Python 3.12 以上 | Django 6.1、psycopg 3 を venv に入れる |
| VBA | (講師のみ)Excel、psqlODBC | 受講生は不要。ODBC ドライバは Excel と同じ bit 数(64bit)を入れる |

**講師から渡すもの**

- 教材フォルダ「⑤研修用（Free版）」(下の構成)
- 講師が生成して組み込んだ Sample01.xlsm(VBA の説明用)

**フォルダ構成**

```
⑤研修用（Free版）\
  ①作業指示\
    ①html\            login.html, product_list.html, order_detail.html, style.css, logo.svg, 画面html.md
    ②テーブル定義書\    テーブル定義書（myappdb）.sql, テーブル定義書（myappdb）.md
    ③指示書\
      ①Java\  ②python\  ③VBA\     各フォルダに 指示書.xlsx と 指示書.md
  ②成果物\            ①Java\  ②python\  ③VBA\(生成した zip の展開先)
```

講師用の有料版「④研修用」も同じ構成です。有料版の指示書は、Claude Code から `①作業指示` を直接参照し、`②成果物` へ出力する前提で書かれています。Free 版は、ファイルをチャットに添付し、zip で受け取る前提です。

## 3. 題材とテーブル定義

### 3.1 接続先(フェーズ①)

| 項目 | 値 |
| --- | --- |
| ホスト / ポート | `localhost` / `5432` |
| DB 名 | `myappdb`(AWS 側と同じ) |
| ユーザー | `myapp` |
| パスワード | 研修用の値(指示書に記載)。Java は `db.properties`、Python は環境変数 `DB_PASSWORD` に書き、AWS へ載せる前に差し替える |

### 3.2 テーブル定義書(sql)

テーブル定義は Excel ではなく、DDL と初期データ(INSERT)を書いた `テーブル定義書（myappdb）.sql` です。A5M2 の SQL 実行で流すと、テーブルと初期データができます。生成したプログラムは SELECT だけを行い、テーブルの作成やデータの投入はしません。

| テーブル | 主な列 | 備考 |
| --- | --- | --- |
| 商品 | 商品コード(PK)、商品名、単価、商品区分、関連商品コード | 42 件 |
| 注文 | 注文日・注文番号・注文枝番(複合 PK)、商品コード、数量、クーポン割引料 | 68 件。クーポン割引料は NULL あり |
| ユーザー情報 | ユーザーid(連番 PK)、eメール、パスワード | 1 件(user@example.com / パスワード 1234) |
| 廃番商品 | 商品コード(PK)、廃番日、売上個数 など | 9 件。本アプリでは使わない |

> **注意**: この SQL は先頭で各テーブルを `DROP TABLE … CASCADE` してから作り直します。実行するとデータが初期化されるので、研修用 DB(myappdb)以外では実行しません。

パスワード列には、入力値を UTF-8 の SHA-256 でハッシュ化した小文字 16 進 64 桁(ソルトなし)を入れます。1234 のハッシュは `03ac674216f3e15c761ee1a5e255f067953623c8b388b4459e13f978d7c846f4` です。

### 3.3 画面仕様

| 画面 | URL(Java / Python) | 内容 |
| --- | --- | --- |
| ログイン | `/login` | eメールとパスワードで認証。失敗時は「ユーザ名またはパスワードが正しくありません。」 |
| 商品一覧 | `/products` | 商品コード順。注文状況に「注文(n 件)」を表示し、注文情報詳細へリンク |
| 注文情報詳細 | `/orders?code=商品コード` | その商品の注文を、注文日降順・注文番号降順・注文枝番昇順で表示 |
| ログアウト | `POST /logout` | セッションを破棄してログイン画面へ |

- 未ログインで商品一覧・注文情報詳細を開いたら、ログイン画面へ戻す
- 注文件数は `COUNT(DISTINCT (o.注文日, o.注文番号)) FILTER (WHERE o.注文番号 IS NOT NULL)` で数え、0 件は「注文なし」(リンクなし)
- 関連商品コードが無い商品は「―」、クーポン割引料が NULL の注文は空欄
- 画面の項目・並び・表示形式は `①html` の HTML に合わせる

## 4. 指示書の構成

生成ルールは、以前の CLAUDE.md ではなく、言語ごとの指示書(`指示書.xlsx` と、同内容の `指示書.md`)に書いてあります。Claude には md 版を添付します。

| 指示内容 | 中身 |
| --- | --- |
| １ 作業フォルダ | Free 版は添付するファイルと zip の作り方。有料版は参照フォルダと出力先 |
| ２ 生成物の条件 | 言語・バージョン、プロジェクト構成、URL、接続情報の持ち方、テーブル定義の参照 |
| ３ ログイン画面 | SHA-256 による認証方法 |
| ４ 商品一覧画面 | 注文件数の数え方、表示形式 |
| ５ 注文情報詳細画面 | 抽出条件と並び順 |
| ６ 動作確認 | 生成物の起動手順(Java・Python のみ) |

言語ごとの要点は次のとおりです。

- **Java**: Maven の war、Java 25、Servlet / JSP(`jakarta.servlet`)、Tomcat 11。JDBC ドライバは WAR に同梱し、接続前に `Class.forName` を呼ぶ。接続情報は `db.properties`。リンクとリダイレクトには `request.getContextPath()` を付ける
- **Python**: Django 6.1、psycopg 3。`INSTALLED_APPS` は sessions・staticfiles・app のみ。モデルは `managed = False` で、注文は `CompositePrimaryKey`。接続情報と `DJANGO_SECRET_KEY` は環境変数から読む
- **VBA**: Excel のシート上のログイン画面(ActiveX)、ADO + psqlODBC。パスワード照合は DB 側の `sha256()` で行い、VBA ではハッシュ計算しない。接続情報は「設定」シートから読む

## 5. 生成手順(Free プランのチャット)

### 5.1 共通の進め方

1. Claude のチャットを新しく開く(言語ごとに別のチャットにする)
2. 添付ファイルを付ける: 指示書.md、login.html、product\_list.html、order\_detail.html、style.css、logo.svg、テーブル定義書（myappdb）.sql(「\_1」などが付いても同じファイルとして扱われる)
3. 次の文を送る
4. 出力された zip をダウンロードし、`②成果物` の該当フォルダに展開する
5. 動かしてエラーが出たら、エラー全文を貼って修正させる(パスワードは伏せる)
6. 動いたら 6 章の確認項目を通す

```
添付の指示書.md に従って、【言語】のアプリを生成してください。
添付の html・style.css・logo.svg・テーブル定義書(sql)を使い、完成したら zip にまとめて出力してください。
```

### 5.2 Java(Eclipse / Tomcat)

1. 【言語】を「Java」にして送り、zip を受け取って展開する(中に Sample01 フォルダがある)
2. Eclipse で「インポート → 既存の Maven プロジェクト」から Sample01 を取り込む
3. `src/main/resources/db.properties` の接続情報がローカル DB と合っているか確認する
4. Tomcat 11 のサーバーにプロジェクトを追加して起動し、`http://localhost:8080/Sample01/login` を開く

### 5.3 Python(VSCode / Django)

1. 【言語】を「Python(Django)」にして送り、zip を展開する(直下に manage.py がある)
2. そのフォルダで venv を作り、依存を入れる

```
py -3.12 -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

3. 環境変数 DB\_HOST・DB\_PORT・DB\_NAME・DB\_USER・DB\_PASSWORD・DJANGO\_SECRET\_KEY を設定する(ターミナルを開き直したら設定し直す。cmd は `set DB_HOST=localhost` のように `=` の前後に空白を入れない)
4. `python manage.py migrate`(django\_session と django\_migrations だけが作られるのが正常。makemigrations は実行しない)
5. `python manage.py runserver` で `http://127.0.0.1:8000/login` を開く
6. 動作確認後、`python -m pip freeze` でバージョン付きの requirements.txt を作り直す(PowerShell は `| Out-File -Encoding ascii requirements.txt`。`>` は UTF-16 になるので使わない)

### 5.4 VBA(講師のみ)

受講生は生成しません。講師は有料版(Claude Code)または Free 版の指示書で生成し、Sample01.xlsm を素材として配布します。Free 版の VBA 指示書は、コードを .bas / .txt と組込手順.txt に分けて zip で出力させ、講師が空のブックに組み込む形です。

- 組み込み前に psqlODBC(64bit)を入れ、「設定」シートに接続情報を入力する
- ログイン欄は ActiveX のテキストボックスです。Tab は自動で移動しないため、KeyDown イベントで実装されています

## 6. 動作確認

各言語の単体確認を通してから、同じユーザーで相互ログインを確認します。講師は研修前に同じ手順を一度通しておきます。

### 6.1 単体確認(各言語)

- [ ] 初期データのユーザー(user@example.com / 1234)でログインできる
- [ ] 間違ったパスワードではログインできず、エラー文言が出る
- [ ] 未ログインで商品一覧・注文情報詳細の URL を直接開くと、ログイン画面に戻る
- [ ] 商品一覧の「注文(n 件)」が、A5M2 で集計した件数と一致する。注文の無い商品は「注文なし」でリンクが無い
- [ ] 注文情報詳細が注文日降順・注文番号降順・注文枝番昇順で、クーポン割引料が NULL の行は空欄
- [ ] 「← 商品一覧に戻る」とログアウトが動き、ログアウト後に商品一覧へ戻れない

### 6.2 相互ログイン

ユーザー登録画面は無いため、DB の初期データにある 1 ユーザーで、Java・Python(・講師の VBA)の全部にログインできることを確認します。ログインできないときは、DB 側のハッシュを次の SQL で確認します(`ok` が true ならハッシュ形式は正しい)。

```sql
SELECT eメール,
       length(パスワード) AS len,  -- 64 であること
       パスワード = encode(sha256(convert_to('1234', 'UTF8')), 'hex') AS ok
FROM ユーザー情報;
```

相互ログインできないときは、ほぼ次のどれかです: 大文字の 16 進、Base64 形式、ソルト付き(Django の `make_password` など)、UTF-8 以外でのバイト変換、列名の不一致。

## 7. フェーズ② AWS への接続

①で動いたコードは、修正なしで AWS に載せられるのが目標です。本番専用の設定は受講生のコードに入れず、AMI 側(講師が用意)で上書きします。詳細は受講生向けデプロイ手順書(Django 編・Tomcat 編、別途作成)にまとめます。

|  | Java | Python |
| --- | --- | --- |
| 公開 URL | `http://IP/Sample01/` | `http://IP:8081/` |
| 送るもの | WAR ファイル(JDBC ドライバ同梱) | `myapp.tar.gz`(config、app、static、templates、manage.py、requirements.txt) |
| 反映方法 | Tomcat の `webapps` に配置(Apache の `tomcat_proxy.conf` にアプリ名を追加) | FileZilla で `/home/ec2-user/` へ送り `bash deploy.sh` |
| 本番設定 | `db.properties` の接続情報を AWS 用に書き換えてから WAR を作る | `/opt/myapp-conf/settings_prod.py` と `/etc/myapp/env` が上書き |

①で `request.getContextPath()`、`{% url %}`、`{% static %}`、環境変数のルールを守っていれば、ここでコードを直す必要はありません。Java の `db.password` は WAR に含まれるため、AWS へ載せる前に必ず書き換えます。

## 8. よくある落とし穴

検証で実際に踏んだものです。受講生が詰まったらまずこの表を見ます。

| 症状 | 原因 | 対処 |
| --- | --- | --- |
| Java で登録(または DB の)ユーザーが Python でログインできない | ハッシュ形式の不一致 | 6.2 の SQL で確認し、指示書のハッシュ仕様を示して修正させる |
| 注文の無い商品が「注文(1 件)」になる | LEFT JOIN で (NULL, NULL) の組が 1 件と数えられる | `FILTER (WHERE o.注文番号 IS NOT NULL)` を付ける |
| Django で makemigrations しても何も起きない | `MIGRATION_MODULES = {'app': None}` による仕様 | 正常。テーブルは sql で作る |
| Django のログインで 500 エラー | django\_session が無い | `python manage.py migrate` |
| 環境変数で KeyError | 環境変数が未設定(ターミナルを開き直した、cmd で `=` の前後に空白を入れた) | 同じターミナルで設定し直す |
| ログイン後にリダイレクトがループする | `request.path` で比較している | `request.path_info` で比較する |
| requirements.txt が作れない / pip install が失敗 | UTF-16 になっている、または pip を実行できない環境で作った | `Out-File -Encoding ascii`。無ければ Django と psycopg の 2 行を手で書く |
| Java の画面遷移が AWS でだけ 404 | リンクが `/login` など絶対パスの直書き | `request.getContextPath()` を付ける |
| Tomcat 11 で ClassNotFound / ドライバが見つからない | `javax.servlet` を使っている、または Class.forName が無い | `jakarta.servlet` に直し、接続前に `Class.forName` を呼ぶ |
| sql を流したらデータが消えた | sql 先頭の `DROP TABLE … CASCADE` | 仕様。再実行すると初期データに戻る。研修用 DB 以外では流さない |
| VBA で「データ ソース名が見つかりません」 | psqlODBC の bit 数・ドライバ名の違い | ODBC データソース管理でドライバ名を確認する(`PostgreSQL Unicode(x64)`) |
| VBA で「変数が定義されていません」 | 宣言されていない定数、モジュール名の誤り | 定数を Const で宣言し、モジュール名を確認して VBE でコンパイルする |
| VBA で NumberFormat が 1004 エラー | 書式に日本語の `[赤]` を使った | 英語の `[Red]` を使う |
| VBA のログイン欄で Tab が効かない | ActiveX は Tab で自動移動しない | KeyDown で実装する(Shift+Tab は戻る) |
| コマンドが通らない | PowerShell / cmd / bash の取り違え | 本書のコマンドは、ローカル = PowerShell(cmd は併記)、サーバー = bash |

## 9. 未決事項

- [x] テーブル構成と初期データ(`テーブル定義書（myappdb）.sql` で確定)
- [x] 研修用 DB を myappdb に集約する(ローカルと AWS で同名)
- [x] VBA のパスワード照合は DB 側の `sha256()` にする
- [x] VBA にユーザー登録画面は付けない(ユーザー登録画面は Java・Python にも無い)
- [x] Java のプロジェクト形式は Maven の war(`jakarta.servlet`、Tomcat 11 で動作確認済み)
- [x] VBA は講師が生成して配布する(Free プランの使用量の都合)
- [x] 研修用アプリの正式名称と Tomcat のアプリ名(現状は Sample01)→ Sample01 で確定(2026/10/30)
- [x] Java の接続情報の渡し方(`db.properties` のままか、環境変数か context.xml か)。AMI 側の設定と合わせる→ Eclipse・AWS とも db.properties で確定(2026/10/30)。AWS へ載せる前に接続情報を書き換えて WAR を作り直す
- [ ] 講師による事前生成と相互ログイン確認(AMI 準備の前に実施)
- [x] Free 版 order\_detail.html の余計な 1 行(「25文字まで表示…」)の削除→ 削除済み(2026/10/30)
