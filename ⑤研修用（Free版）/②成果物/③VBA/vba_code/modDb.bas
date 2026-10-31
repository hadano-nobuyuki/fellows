Attribute VB_Name = "modDb"
Option Explicit

'==============================================================
' modDb : DB 接続と SQL 実行(ADO + psqlODBC、遅延バインディング)
'   - 接続情報は「設定」シートの A 列(項目名)/ B 列(値)から読み込む
'   - SELECT のみ実行する(DDL/DML は実行しない)
'   - SQL の値は ADODB.Command のパラメータ(?)で渡す
'==============================================================

' ADO の定数(参照設定を使わないため、値をここで宣言する)
Private Const adVarWChar As Long = 202
Private Const adParamInput As Long = 1
Private Const adCmdText As Long = 1
Private Const adStateOpen As Long = 1

Private Const SHEET_SETTINGS As String = "設定"
Private Const SETTINGS_MAX_ROW As Long = 50
Private Const ODBC_DRIVER As String = "PostgreSQL Unicode(x64)"

' 「設定」シートから項目名に対応する値(B 列)を返す。項目が無ければ空文字。
Private Function GetSetting(ByVal itemName As String) As String
    Dim ws As Worksheet
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    For r = 1 To SETTINGS_MAX_ROW
        If Trim$(CStr(ws.Cells(r, 1).Value)) = itemName Then
            GetSetting = CStr(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    GetSetting = ""
End Function

' 必須の設定値を返す。未入力ならエラーにする。
Private Function RequireSetting(ByVal itemName As String) As String
    Dim v As String

    v = Trim$(GetSetting(itemName))
    If Len(v) = 0 Then
        Err.Raise vbObjectError + 1001, "modDb.RequireSetting", _
                  "「" & SHEET_SETTINGS & "」シートの " & itemName & " が入力されていません。"
    End If
    RequireSetting = v
End Function

Private Function BuildConnectionString() As String
    Dim dbHost As String
    Dim dbPort As String
    Dim dbName As String
    Dim dbUser As String
    Dim dbPassword As String

    dbHost = RequireSetting("DB_HOST")
    dbPort = RequireSetting("DB_PORT")
    dbName = RequireSetting("DB_NAME")
    dbUser = RequireSetting("DB_USER")
    dbPassword = GetSetting("DB_PASSWORD")

    BuildConnectionString = "Driver={" & ODBC_DRIVER & "};" & _
                            "Server=" & dbHost & ";" & _
                            "Port=" & dbPort & ";" & _
                            "Database=" & dbName & ";" & _
                            "Uid=" & dbUser & ";" & _
                            "Pwd=" & dbPassword & ";"
End Function

' Recordset を 1 始まりの 2 次元配列(行, 列)に変換する。0 件なら Empty を返す。
Private Function RecordsetToArray(ByVal rs As Object) As Variant
    Dim raw As Variant
    Dim result() As Variant
    Dim r As Long
    Dim c As Long
    Dim nRows As Long
    Dim nCols As Long

    If rs Is Nothing Then Exit Function
    If rs.State <> adStateOpen Then Exit Function
    If rs.EOF Then Exit Function

    raw = rs.GetRows
    nCols = UBound(raw, 1) - LBound(raw, 1) + 1
    nRows = UBound(raw, 2) - LBound(raw, 2) + 1

    ReDim result(1 To nRows, 1 To nCols)
    For r = 1 To nRows
        For c = 1 To nCols
            result(r, c) = raw(LBound(raw, 1) + c - 1, LBound(raw, 2) + r - 1)
        Next c
    Next r

    RecordsetToArray = result
End Function

' SELECT を実行して結果を 2 次元配列で返す(0 件なら Empty)。
' params は Array(...) で渡す。パラメータが無い場合は Array() を渡す。
' 接続は成功・失敗にかかわらず必ず閉じる。
Private Function QueryRows(ByVal sql As String, ByVal params As Variant) As Variant
    Dim conn As Object
    Dim cmd As Object
    Dim prm As Object
    Dim rs As Object
    Dim i As Long
    Dim v As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    Set conn = CreateObject("ADODB.Connection")
    conn.Open BuildConnectionString()

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = adCmdText

    If IsArray(params) Then
        For i = LBound(params) To UBound(params)
            v = CStr(params(i))
            Set prm = cmd.CreateParameter("p" & CStr(i - LBound(params) + 1), adVarWChar, adParamInput, Len(v) + 1, v)
            cmd.Parameters.Append prm
        Next i
    End If

    Set rs = cmd.Execute
    QueryRows = RecordsetToArray(rs)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set rs = Nothing
    Set prm = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    On Error GoTo 0
    If errNum <> 0 Then
        Err.Raise errNum, "modDb.QueryRows", errDesc
    End If
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Function

' 認証: eメールとパスワードが一致するユーザーが 1 件以上あれば True。
' パスワードのハッシュ化(SHA-256, UTF-8, 小文字16進)は DB 側で行う。
Public Function Authenticate(ByVal userName As String, ByVal password As String) As Boolean
    Dim sql As String
    Dim recs As Variant

    sql = "SELECT COUNT(*) FROM ユーザー情報 " & _
          "WHERE eメール = ? " & _
          "AND パスワード = encode(sha256(convert_to(?, 'UTF8')), 'hex')"

    recs = QueryRows(sql, Array(userName, password))
    If IsEmpty(recs) Then
        Authenticate = False
    Else
        Authenticate = (CLng(recs(1, 1)) >= 1)
    End If
End Function

' 商品一覧情報: 商品を全件取得し、注文と LEFT JOIN して注文件数(注文日+注文番号の組)を数える。
' 列: 1=商品コード 2=商品名 3=単価 4=商品区分 5=関連商品コード 6=注文件数
Public Function GetProducts() As Variant
    Dim sql As String

    sql = "SELECT p.商品コード, p.商品名, p.単価, p.商品区分, p.関連商品コード, " & _
          "COUNT(DISTINCT (o.注文日, o.注文番号)) FILTER (WHERE o.注文番号 IS NOT NULL) AS 件数 " & _
          "FROM 商品 p " & _
          "LEFT JOIN 注文 o ON o.商品コード = p.商品コード " & _
          "GROUP BY p.商品コード, p.商品名, p.単価, p.商品区分, p.関連商品コード " & _
          "ORDER BY p.商品コード"

    GetProducts = QueryRows(sql, Array())
End Function

' 注文情報: 指定した商品コードの注文を取得する。
' 列: 1=注文日 2=注文番号 3=注文枝番 4=数量 5=クーポン割引料
Public Function GetOrders(ByVal productCode As String) As Variant
    Dim sql As String

    sql = "SELECT 注文日, 注文番号, 注文枝番, 数量, クーポン割引料 " & _
          "FROM 注文 " & _
          "WHERE 商品コード = ? " & _
          "ORDER BY 注文日 DESC, 注文番号 DESC, 注文枝番 ASC"

    GetOrders = QueryRows(sql, Array(productCode))
End Function
