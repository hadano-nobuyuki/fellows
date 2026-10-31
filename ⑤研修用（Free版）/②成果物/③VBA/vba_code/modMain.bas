Attribute VB_Name = "modMain"
Option Explicit

'==============================================================
' modMain : ログイン状態の管理、各シートの表示、ログアウト
'==============================================================

' ログイン中のユーザー(eメール)。空ならログインしていない扱い。
Public gLoginUser As String

Private Const SHEET_LOGIN As String = "ログイン"
Private Const SHEET_PRODUCTS As String = "商品一覧"
Private Const SHEET_ORDERS As String = "注文情報詳細"

Private Const CTL_USER As String = "txtUserName"
Private Const CTL_PASSWORD As String = "txtPassword"
Private Const NAME_MESSAGE As String = "LoginMessage"

Private Const PRODUCT_FIRST_ROW As Long = 4
Private Const ORDER_FIRST_ROW As Long = 5
Private Const PRODUCT_NAME_MAX_LEN As Long = 25

Private Const MSG_LOGIN_FAILED As String = "ユーザ名またはパスワードが正しくありません。"
Private Const MSG_INPUT_REQUIRED As String = "ユーザ名とパスワードを入力してください。"
Private Const DASH_MARK As String = "―"

'--------------------------------------------------------------
' エラー表示(プロシージャ名とエラー番号を含める)
'--------------------------------------------------------------
Private Sub ReportError(ByVal procName As String, ByVal whatFailed As String, _
                        ByVal errNum As Long, ByVal errDesc As String)
    MsgBox whatFailed & vbCrLf & vbCrLf & _
           "プロシージャ: " & procName & vbCrLf & _
           "エラー番号: " & CStr(errNum) & vbCrLf & _
           "内容: " & errDesc, vbExclamation, "エラー"
End Sub

'--------------------------------------------------------------
' ブックを開いたときの初期化(ThisWorkbook.Workbook_Open から呼ぶ)
'--------------------------------------------------------------
Public Sub InitWorkbook()
    On Error GoTo EH

    gLoginUser = ""
    ClearDataRows SHEET_PRODUCTS, PRODUCT_FIRST_ROW
    ClearDataRows SHEET_ORDERS, ORDER_FIRST_ROW
    ClearTitles
    ClearLoginInputs
    ShowLogin
    Exit Sub

EH:
    ReportError "modMain.InitWorkbook", "初期化処理でエラーが発生しました。", Err.Number, Err.Description
End Sub

'--------------------------------------------------------------
' ログイン画面の表示
'--------------------------------------------------------------
Public Sub ShowLogin(Optional ByVal message As String = "")
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_LOGIN)
    ws.Activate

    With ws.Range(NAME_MESSAGE)
        .Value = message
        .Font.Color = RGB(220, 38, 38)
    End With

    ' フォーカスを ユーザ名 欄へ(失敗しても処理は続ける)
    On Error Resume Next
    ws.OLEObjects(CTL_USER).Activate
    On Error GoTo 0
End Sub

Private Sub ClearPassword()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_LOGIN)
    ws.OLEObjects(CTL_PASSWORD).Object.Text = ""
End Sub

Private Sub ClearLoginInputs()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_LOGIN)
    ws.OLEObjects(CTL_USER).Object.Text = ""
    ws.OLEObjects(CTL_PASSWORD).Object.Text = ""
End Sub

'--------------------------------------------------------------
' ログイン実行(ログインボタン・Enter キーから呼ぶ)
'--------------------------------------------------------------
Public Sub DoLogin()
    Dim ws As Worksheet
    Dim userName As String
    Dim password As String
    Dim isOk As Boolean
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    Set ws = ThisWorkbook.Worksheets(SHEET_LOGIN)
    userName = Trim$(CStr(ws.OLEObjects(CTL_USER).Object.Text))
    password = CStr(ws.OLEObjects(CTL_PASSWORD).Object.Text)

    If Len(userName) = 0 Or Len(password) = 0 Then
        ClearPassword
        ShowLogin MSG_INPUT_REQUIRED
        Exit Sub
    End If

    Application.Cursor = xlWait
    isOk = modDb.Authenticate(userName, password)
    Application.Cursor = xlDefault

    ClearPassword

    If isOk Then
        gLoginUser = userName
        ws.Range(NAME_MESSAGE).Value = ""
        ShowProducts
    Else
        ShowLogin MSG_LOGIN_FAILED
    End If
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    On Error Resume Next
    ClearPassword
    On Error GoTo 0
    ReportError "modMain.DoLogin", "ログイン処理でエラーが発生しました。", errNum, errDesc
End Sub

'--------------------------------------------------------------
' 商品一覧シートの表示(一覧を DB から読み込む)
'--------------------------------------------------------------
Public Sub ShowProducts()
    Dim errNum As Long
    Dim errDesc As String

    If Len(gLoginUser) = 0 Then
        ShowLogin
        Exit Sub
    End If

    On Error GoTo EH
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(SHEET_PRODUCTS).Activate
    LoadProducts
    Application.ScreenUpdating = True
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = True
    ReportError "modMain.ShowProducts", "商品一覧の読み込みでエラーが発生しました。", errNum, errDesc
End Sub

' 「← 商品一覧に戻る」ボタン用: 商品一覧シートを表示する
Public Sub BackToProducts()
    If Len(gLoginUser) = 0 Then
        ShowLogin
        Exit Sub
    End If
    ThisWorkbook.Worksheets(SHEET_PRODUCTS).Activate
End Sub

'--------------------------------------------------------------
' 注文情報詳細シートの表示(指定した商品コードの注文を読み込む)
'--------------------------------------------------------------
Public Sub ShowOrderDetail(ByVal productCode As String)
    Dim errNum As Long
    Dim errDesc As String

    If Len(gLoginUser) = 0 Then
        ShowLogin
        Exit Sub
    End If

    On Error GoTo EH
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(SHEET_ORDERS).Activate
    LoadOrders productCode
    Application.ScreenUpdating = True
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = True
    ReportError "modMain.ShowOrderDetail", "注文情報の読み込みでエラーが発生しました。", errNum, errDesc
End Sub

'--------------------------------------------------------------
' ログアウト(ログアウトボタンに登録)
'--------------------------------------------------------------
Public Sub Logout()
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    gLoginUser = ""
    ClearDataRows SHEET_PRODUCTS, PRODUCT_FIRST_ROW
    ClearDataRows SHEET_ORDERS, ORDER_FIRST_ROW
    ClearTitles
    ClearLoginInputs
    ShowLogin
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    ReportError "modMain.Logout", "ログアウト処理でエラーが発生しました。", errNum, errDesc
End Sub

'--------------------------------------------------------------
' 以下、内部処理
'--------------------------------------------------------------

' 指定行以降(データ行)の内容・書式・ハイパーリンクを消す
Private Sub ClearDataRows(ByVal sheetName As String, ByVal firstRow As Long)
    Dim ws As Worksheet
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set rng = ws.Rows(CStr(firstRow) & ":" & CStr(ws.Rows.Count))

    On Error Resume Next
    rng.Hyperlinks.Delete
    On Error GoTo 0

    rng.Clear
End Sub

' ログイン中ユーザー・商品コードの表示を消す
Private Sub ClearTitles()
    ThisWorkbook.Worksheets(SHEET_PRODUCTS).Range("D1").ClearContents
    ThisWorkbook.Worksheets(SHEET_ORDERS).Range("A2").ClearContents
End Sub

' データ行の下線(見やすさのため)
Private Sub ApplyBodyBorders(ByVal rng As Range)
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
    End With
    If rng.Rows.Count > 1 Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Color = RGB(226, 232, 240)
        End With
    End If
End Sub

' 商品一覧シートに DB の内容を出力する
Private Sub LoadProducts()
    Dim ws As Worksheet
    Dim recs As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim n As Long
    Dim cnt As Long
    Dim relCode As String
    Dim lastRow As Long
    Dim rng As Range
    Dim cel As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUCTS)
    ws.Range("D1").Value = "ログイン中:" & gLoginUser

    ' 前回のデータを消す
    ClearDataRows SHEET_PRODUCTS, PRODUCT_FIRST_ROW

    recs = modDb.GetProducts()
    If IsEmpty(recs) Then Exit Sub

    n = UBound(recs, 1)
    ReDim outArr(1 To n, 1 To 6)

    For i = 1 To n
        outArr(i, 1) = Trim$(CStr(recs(i, 1)))
        outArr(i, 2) = Left$(CStr(recs(i, 2)), PRODUCT_NAME_MAX_LEN)
        outArr(i, 3) = CLng(recs(i, 3))
        outArr(i, 4) = Trim$(CStr(recs(i, 4)))

        If IsNull(recs(i, 5)) Then
            relCode = ""
        Else
            relCode = Trim$(CStr(recs(i, 5)))
        End If
        If Len(relCode) = 0 Then
            outArr(i, 5) = DASH_MARK
        Else
            outArr(i, 5) = relCode
        End If

        cnt = CLng(recs(i, 6))
        If cnt > 0 Then
            outArr(i, 6) = "注文(" & CStr(cnt) & " 件)"
        Else
            outArr(i, 6) = "注文なし"
        End If
    Next i

    lastRow = PRODUCT_FIRST_ROW + n - 1
    Set rng = ws.Range(ws.Cells(PRODUCT_FIRST_ROW, 1), ws.Cells(lastRow, 6))

    ' 文字列として出力する(先頭 0 などが消えないように)。単価のみ数値。
    rng.NumberFormat = "@"
    ws.Range(ws.Cells(PRODUCT_FIRST_ROW, 3), ws.Cells(lastRow, 3)).NumberFormat = "#,##0"
    rng.Value = outArr

    rng.HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(PRODUCT_FIRST_ROW, 3), ws.Cells(lastRow, 3)).HorizontalAlignment = xlRight
    ApplyBodyBorders rng

    ' 注文がある商品の「注文状況」にハイパーリンクを設定する(自セルへのリンク)
    For i = 1 To n
        If CLng(recs(i, 6)) > 0 Then
            Set cel = ws.Cells(PRODUCT_FIRST_ROW + i - 1, 6)
            ws.Hyperlinks.Add Anchor:=cel, Address:="", _
                SubAddress:="'" & SHEET_PRODUCTS & "'!F" & CStr(PRODUCT_FIRST_ROW + i - 1), _
                TextToDisplay:=CStr(outArr(i, 6))
            cel.Font.Color = RGB(14, 116, 144)
        End If
    Next i
End Sub

' 注文情報詳細シートに DB の内容を出力する
Private Sub LoadOrders(ByVal productCode As String)
    Dim ws As Worksheet
    Dim recs As Variant
    Dim outArr() As Variant
    Dim i As Long
    Dim n As Long
    Dim lastRow As Long
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_ORDERS)
    ws.Range("A2").Value = "商品コード:" & productCode

    ' 前回のデータを消す
    ClearDataRows SHEET_ORDERS, ORDER_FIRST_ROW

    recs = modDb.GetOrders(productCode)
    If IsEmpty(recs) Then Exit Sub

    n = UBound(recs, 1)
    ReDim outArr(1 To n, 1 To 5)

    For i = 1 To n
        outArr(i, 1) = CDate(recs(i, 1))
        outArr(i, 2) = Trim$(CStr(recs(i, 2)))
        outArr(i, 3) = CLng(recs(i, 3))
        outArr(i, 4) = CLng(recs(i, 4))
        If IsNull(recs(i, 5)) Then
            outArr(i, 5) = Empty
        Else
            outArr(i, 5) = CLng(recs(i, 5))
        End If
    Next i

    lastRow = ORDER_FIRST_ROW + n - 1
    Set rng = ws.Range(ws.Cells(ORDER_FIRST_ROW, 1), ws.Cells(lastRow, 5))

    ' 書式は NumberFormat に英語の書式コードで設定する
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 1), ws.Cells(lastRow, 1)).NumberFormat = "yyyy/mm/dd"
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 2), ws.Cells(lastRow, 2)).NumberFormat = "@"
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 3), ws.Cells(lastRow, 3)).NumberFormat = "0"
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 4), ws.Cells(lastRow, 4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 5), ws.Cells(lastRow, 5)).NumberFormat = "#,##0;[Red]-#,##0"

    rng.Value = outArr

    ' 注文日・注文番号は左寄せ(日付は標準で右寄せになるため明示する)、数値列は右寄せ
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 1), ws.Cells(lastRow, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(ORDER_FIRST_ROW, 3), ws.Cells(lastRow, 5)).HorizontalAlignment = xlRight
    ApplyBodyBorders rng
End Sub
