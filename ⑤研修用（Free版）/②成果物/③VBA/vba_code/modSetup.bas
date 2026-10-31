Attribute VB_Name = "modSetup"
Option Explicit

'==============================================================
' modSetup : シートと画面部品の作成
'   空のブックで BuildWorkbook を 1 回実行すると、次を作成する。
'     - シート「ログイン」「商品一覧」「注文情報詳細」「設定」と見出し
'     - 名前付き範囲 LoginMessage
'     - ActiveX: txtUserName / txtPassword / cmdLogin
'     - ボタン(図形): ログアウト・商品一覧に戻る
'     - 「設定」シート A 列の項目名
'==============================================================

' MSForms の定数(参照設定に頼らないため、値をここで宣言する)
Private Const fmBorderStyleSingle As Long = 1
Private Const fmSpecialEffectFlat As Long = 0
Private Const fmIMEModeDisable As Long = 3

Private Const SHEET_LOGIN As String = "ログイン"
Private Const SHEET_PRODUCTS As String = "商品一覧"
Private Const SHEET_ORDERS As String = "注文情報詳細"
Private Const SHEET_SETTINGS As String = "設定"

Private Const APP_NAME As String = "ClaudeCode 研修用プログラム"
Private Const BUTTON_HEIGHT As Double = 26

'--------------------------------------------------------------
' メイン: ブックを完成させる
'--------------------------------------------------------------
Public Sub BuildWorkbook()
    Dim wsLogin As Worksheet
    Dim wsProd As Worksheet
    Dim wsOrder As Worksheet
    Dim wsSet As Worksheet
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH
    Application.ScreenUpdating = False

    ' 既定のフォントを統一する(Normal スタイル)
    On Error Resume Next
    ThisWorkbook.Styles("Normal").Font.Name = "Meiryo UI"
    ThisWorkbook.Styles("Normal").Font.Size = 10
    On Error GoTo EH

    Set wsLogin = EnsureSheet(SHEET_LOGIN, True)
    Set wsProd = EnsureSheet(SHEET_PRODUCTS, False)
    Set wsOrder = EnsureSheet(SHEET_ORDERS, False)
    Set wsSet = EnsureSheet(SHEET_SETTINGS, False)

    BuildLoginSheet wsLogin
    BuildProductSheet wsProd
    BuildOrderSheet wsOrder
    BuildSettingsSheet wsSet

    wsLogin.Activate
    Application.ScreenUpdating = True

    MsgBox "シートと画面部品を作成しました。" & vbCrLf & vbCrLf & _
           "次の作業:" & vbCrLf & _
           "1. Sheet_ログイン.txt / Sheet_商品一覧.txt を各シートのコードに貼り付ける" & vbCrLf & _
           "2. 「設定」シートの B 列に接続情報を入力する" & vbCrLf & _
           "3. 保存して、ブックを開き直す" & vbCrLf & vbCrLf & _
           "※ ActiveX が動かないときは [開発] タブの「デザイン モード」を OFF にしてください。", _
           vbInformation, "BuildWorkbook"
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = True
    MsgBox "BuildWorkbook でエラーが発生しました。" & vbCrLf & vbCrLf & _
           "プロシージャ: modSetup.BuildWorkbook" & vbCrLf & _
           "エラー番号: " & CStr(errNum) & vbCrLf & _
           "内容: " & errDesc, vbCritical, "エラー"
End Sub

'--------------------------------------------------------------
' シートの取得(無ければ作る)。isFirst = True のときは既存の 1 枚目を改名する。
'--------------------------------------------------------------
Private Function EnsureSheet(ByVal sheetName As String, ByVal isFirst As Boolean) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        If isFirst Then
            Set ws = ThisWorkbook.Worksheets(1)
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        End If
        ws.Name = sheetName
    End If

    Set EnsureSheet = ws
End Function

Private Sub DeleteOleIfExists(ByVal ws As Worksheet, ByVal ctlName As String)
    Dim ole As OLEObject

    For Each ole In ws.OLEObjects
        If ole.Name = ctlName Then
            ole.Delete
            Exit For
        End If
    Next ole
End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim shp As Shape

    For Each shp In ws.Shapes
        If shp.Name = shapeName Then
            shp.Delete
            Exit For
        End If
    Next shp
End Sub

'--------------------------------------------------------------
' ActiveX テキストボックス(ユーザ名・パスワードで同じ見た目にそろえる)
'--------------------------------------------------------------
Private Function AddTextBox(ByVal ws As Worksheet, ByVal ctlName As String, ByVal anchor As Range) As OLEObject
    Dim ole As OLEObject

    DeleteOleIfExists ws, ctlName

    Set ole = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
                                Left:=anchor.Left, Top:=anchor.Top + 1, _
                                Width:=anchor.Width, Height:=anchor.Height - 2)
    ole.Name = ctlName
    ole.Placement = xlMove

    With ole.Object
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .ForeColor = RGB(30, 27, 75)
        .BackColor = RGB(248, 250, 252)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(203, 213, 225)
        .SpecialEffect = fmSpecialEffectFlat
        .IMEMode = fmIMEModeDisable
        .MaxLength = 255
    End With

    Set AddTextBox = ole
End Function

'--------------------------------------------------------------
' 図形のボタン(高さ 24pt 以上、文字は上下左右とも中央)
'--------------------------------------------------------------
Private Sub AddButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
                      ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPt As Double, _
                      ByVal macroName As String, ByVal isFilled As Boolean)
    Dim shp As Shape
    Dim fontColor As Long

    DeleteShapeIfExists ws, shapeName

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthPt, BUTTON_HEIGHT)
    shp.Name = shapeName
    shp.Placement = xlMove
    shp.OnAction = macroName
    shp.Adjustments.Item(1) = 0.5

    If isFilled Then
        shp.Fill.ForeColor.RGB = RGB(79, 70, 229)
        shp.Line.Visible = msoFalse
        fontColor = RGB(255, 255, 255)
    Else
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Line.ForeColor.RGB = RGB(79, 70, 229)
        shp.Line.Weight = 1
        fontColor = RGB(79, 70, 229)
    End If
    shp.Shadow.Visible = msoFalse

    With shp.TextFrame2
        .WordWrap = msoFalse
        .MarginTop = 0
        .MarginBottom = 0
        .MarginLeft = 6
        .MarginRight = 6
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = caption
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = fontColor
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

' 見出しセルの書式
Private Sub StyleHeaderRow(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .Interior.Color = RGB(248, 250, 252)
        .VerticalAlignment = xlCenter
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
    End With
End Sub

' 見出し行の下を固定する
Private Sub FreezeBelow(ByVal ws As Worksheet, ByVal firstDataRow As Long)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Cells(firstDataRow, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

'--------------------------------------------------------------
' 「ログイン」シート
'--------------------------------------------------------------
Private Sub BuildLoginSheet(ByVal ws As Worksheet)
    Dim ole As OLEObject

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Cells.Interior.Color = RGB(244, 246, 251)
    ws.Range("B2:B13").Interior.Color = RGB(255, 255, 255)

    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 40

    ws.Rows(1).RowHeight = 12
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 8
    ws.Rows(5).RowHeight = 32
    ws.Rows(6).RowHeight = 8
    ws.Rows(7).RowHeight = 20
    ws.Rows(8).RowHeight = 28
    ws.Rows(9).RowHeight = 10
    ws.Rows(10).RowHeight = 20
    ws.Rows(11).RowHeight = 28
    ws.Rows(12).RowHeight = 14
    ws.Rows(13).RowHeight = 34

    With ws.Range("B2")
        .Value = APP_NAME
        .Font.Color = RGB(100, 116, 139)
        .Font.Bold = True
    End With
    With ws.Range("B3")
        .Value = "ログイン画面"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(30, 27, 75)
        .VerticalAlignment = xlCenter
    End With

    ' メッセージセル(名前付き範囲 LoginMessage)
    With ws.Range("B5")
        .Value = ""
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Font.Color = RGB(220, 38, 38)
    End With
    ThisWorkbook.Names.Add Name:="LoginMessage", RefersTo:="='" & ws.Name & "'!$B$5"

    ' ラベル(login.html に準拠)
    With ws.Range("B7")
        .Value = "ユーザ名"
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("B10")
        .Value = "パスワード"
        .Font.Bold = True
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlBottom
    End With

    ' ActiveX テキストボックス(同じ大きさ・枠線・背景色・フォント)
    Set ole = AddTextBox(ws, "txtUserName", ws.Range("B8"))
    Set ole = AddTextBox(ws, "txtPassword", ws.Range("B11"))
    ole.Object.PasswordChar = "●"

    ' ActiveX ログインボタン
    DeleteOleIfExists ws, "cmdLogin"
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=ws.Range("B13").Left, Top:=ws.Range("B13").Top + 1, _
                                Width:=ws.Range("B13").Width, Height:=ws.Range("B13").Height - 2)
    ole.Name = "cmdLogin"
    ole.Placement = xlMove
    With ole.Object
        .Caption = "ログイン"
        .Font.Name = "Meiryo UI"
        .Font.Size = 12
        .Font.Bold = True
        .BackColor = RGB(79, 70, 229)
        .ForeColor = RGB(255, 255, 255)
    End With

    ws.Range("A1").Select
End Sub

'--------------------------------------------------------------
' 「商品一覧」シート
'--------------------------------------------------------------
Private Sub BuildProductSheet(ByVal ws As Worksheet)
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 18

    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 32
    ws.Rows(3).RowHeight = 24

    With ws.Range("A1")
        .Value = "商品一覧画面"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(30, 27, 75)
    End With
    ws.Range("A1:F1").VerticalAlignment = xlCenter
    ws.Range("D1").Font.Color = RGB(100, 116, 139)

    ws.Range("A3").Value = "商品コード"
    ws.Range("B3").Value = "商品名"
    ws.Range("C3").Value = "単価"
    ws.Range("D3").Value = "商品区分"
    ws.Range("E3").Value = "関連商品コード"
    ws.Range("F3").Value = "注文状況"
    StyleHeaderRow ws.Range("A3:F3")
    ws.Range("C3").HorizontalAlignment = xlRight

    AddButton ws, "btnLogout", "ログアウト", ws.Range("A2").Left + 2, ws.Range("A2").Top + 3, 90, _
              "modMain.Logout", False

    FreezeBelow ws, 4
End Sub

'--------------------------------------------------------------
' 「注文情報詳細」シート
'--------------------------------------------------------------
Private Sub BuildOrderSheet(ByVal ws As Worksheet)
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 18

    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 24

    With ws.Range("A1")
        .Value = "注文情報詳細画面"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(30, 27, 75)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A2").Value = "商品コード:"
    ws.Range("A2").Font.Color = RGB(100, 116, 139)

    ws.Range("A4").Value = "注文日"
    ws.Range("B4").Value = "注文番号"
    ws.Range("C4").Value = "注文枝番"
    ws.Range("D4").Value = "数量"
    ws.Range("E4").Value = "クーポン割引料"
    StyleHeaderRow ws.Range("A4:E4")
    ws.Range("C4:E4").HorizontalAlignment = xlRight

    AddButton ws, "btnBack", ChrW(8592) & " 商品一覧に戻る", ws.Range("A3").Left + 2, ws.Range("A3").Top + 3, 130, _
              "modMain.BackToProducts", True
    AddButton ws, "btnLogout", "ログアウト", ws.Range("A3").Left + 140, ws.Range("A3").Top + 3, 90, _
              "modMain.Logout", False

    FreezeBelow ws, 5
End Sub

'--------------------------------------------------------------
' 「設定」シート(A 列に項目名。B 列の値は利用者が入力する)
'--------------------------------------------------------------
Private Sub BuildSettingsSheet(ByVal ws As Worksheet)
    Dim items As Variant
    Dim i As Long

    ws.Activate

    items = Array("DB_HOST", "DB_PORT", "DB_NAME", "DB_USER", "DB_PASSWORD")
    For i = LBound(items) To UBound(items)
        ws.Cells(i + 1, 1).Value = items(i)
    Next i

    ' 値は文字列として保持する(先頭 0 などが消えないように)
    ws.Range("B1:B5").NumberFormat = "@"

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 30
    ws.Range("A1:A5").Font.Bold = True

    ws.Range("D1").Value = "B 列に接続情報を入力してください。"
    ws.Range("D2").Value = "※ パスワードがブックに保存されます。ブックを他人に渡すときは B 列の値を消してください。"
    ws.Range("D1:D2").Font.Color = RGB(100, 116, 139)

    ws.Range("A1").Select
End Sub
