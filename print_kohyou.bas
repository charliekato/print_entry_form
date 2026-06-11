Attribute VB_Name = "print_kohyou"
    Dim Names, Names_2 As Variant

    Dim Names2, Names2_2 As Variant
Sub showMenu()
    Menu.Show
End Sub

Sub print_run()
Attribute print_run.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' print_run Macro
'
' Keyboard Shortcut: Ctrl+p
'
    If Menu.CheckBox4Debug Then
        Exit Sub
    End If
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
Sub insert_data(rangeName As Variant, rs As Object)

    Range(rangeName) = rs(rangeName)
    
End Sub
Sub insert_page2(rangeName As Variant, rs As Object)
    Range(rangeName + "_2") = rs(rangeName)
End Sub

Sub clear_data()
    Dim name As Variant
    For Each name In Names
        Range(name) = ""
    Next name
    For Each name In Names2
        Range(name) = ""
    Next name
    For Each name In Names_2
        Range(name) = ""
    Next name
    For Each name In Names2_2
        Range(name) = ""
    Next name
End Sub




Sub go_ahead()
    Dim i As Long

    Dim counter As Integer
    
    Call preparation

    'Names = Array("氏名", "フリガナ", "学校名", "学年", "泳力", "保護者氏名", "住所", "郵便番号", "緊急連絡先1", _
    '              "緊急連絡先2", "緊急連絡先3", _
    '              "備考", "メールアドレス", "前年度最終泳力")
    Names = Array("ID", "氏名", "フリガナ", "学校名", "学年", "泳力", "保護者氏名", "住所", "郵便番号", "緊急連絡先1", _
                  "緊急連絡先2", "緊急連絡先3", _
                  "備考", "メールアドレス")

    
    Names_2 = Array("ID_2", "氏名_2", "フリガナ_2", "学校名_2", "学年_2", "泳力_2", "保護者氏名_2", "住所_2", "郵便番号_2", "緊急連絡先1_2", _
                  "緊急連絡先2_2", "緊急連絡先3_2", _
                  "備考_2", "メールアドレス")
                  

    Names2 = Array("泳力1", "泳力2", "泳力3")
    Names2_2 = Array("泳力1_2", "泳力2_2", "泳力3_2")
    counter = 1
    For i = Menu.TextBoxFrom.Value To Menu.TextBoxTo.Value
        If counter = 1 Then
            Call clear_data
        End If
        Call print_specific_id(i)
        If counter = 2 Then
            counter = 0
            Call print_run
        End If
        counter = counter + 1
    Next i
    If counter = 2 Then
        Call print_run
    End If
End Sub


Sub change_eiryoku_description()
    Cells.Replace What:="泳げる --> 何メートル泳げるか記入", Replacement:="3", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="顔をつけて浮くことができる(伏し浮きやけのびができる)", Replacement:="2", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="顔が水につけられない(水が怖い)", Replacement:="1", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub
Sub insert_all(rs As Object, id As Long)
    Dim name As Variant

    If id Mod 2 = 0 Then
        For Each name In Names
            If (name <> "住所") Or (Not Menu.CheckBox4Instructor.Value) Then
           
                    Call insert_page2(name, rs)

            End If
            
        Next name
        Range("泳力" & rs("泳力C") & "_2") = "〇"
    Else
        For Each name In Names
            If (name <> "住所") Or (Not Menu.CheckBox4Instructor.Value) Then
                Call insert_data(name, rs)
            End If
        Next name
        Range("泳力" & rs("泳力C")) = "〇"
    End If

        
End Sub
Sub open_meibo(fname As String)
     Workbooks.Open Filename:=fname
End Sub
Function last_row(column As Integer) As Integer
    last_row = Cells(Rows.Count, column).End(xlUp).row
End Function

Sub change_sheet_name()

    Sheets("フォームの回答 1").name = "Sheet1"
End Sub
Sub insert_id()
    Dim row As Long
    Dim endline As Long
    endline = last_row(1)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ID"
    
    For row = 2 To endline
        Cells(row, 1) = row - 1
    Next row
    
End Sub


Sub change_header()
    Cells(1, 4).Value = "氏名"
    Cells(1, 8).Value = "泳力C"
    Cells(1, 9).Value = "泳力"
    Cells(1, 10).Value = "備考"
End Sub
Sub preparation()

    If Menu.CheckBox4SkipPreparation.Value Then
        Exit Sub
    End If
    Call open_meibo(Menu.TextBoxFileName.Value)
    ' column を挿入しheader に "ID" として番号を振る
    Call insert_id
    Call change_sheet_name  ' sheet 名を Sheet1 に
    Call change_header
    Call change_eiryoku_description
    ActiveWorkbook.Close
End Sub
Sub print_specific_id(id As Long)
    Const adOpenKeyset = 1
    Const adLockReadOnly = 1
    Dim cn As Object
    Dim rs As Object
    Dim SQLstr As String
    Dim row As Integer
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    cn.Provider = "Microsoft.ACE.OLEDB.12.0"

    cn.Properties("Extended Properties") = "Excel 12.0;HDR=YES;IMEX=1"
    cn.Open Menu.TextBoxFileName.Value

    SQLstr = "SELECT * FROM [Sheet1$] where ID = " & id & ";"
    rs.Open SQLstr, cn, adOpenKeyset, adLockReadOnly
    Do Until rs.EOF
        Call insert_all(rs, id)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub


Sub create_parent_list()
    Dim cn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim strFilePath As String
    Dim ws As Worksheet
    Dim i As Integer
    
    ' 接続文字列の設定
    strFilePath = Range("受講者名簿のファイル名")
    
    ' ADOオブジェクトの作成
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Excelファイルに接続
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                          "Data Source=" & strFilePath & ";" & _
                          "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    cn.Open
    
    ' SQLクエリの設定 (重複を削除)
    strSQL = "SELECT 受付番号, 住所, 郵便番号, 保護者氏名 FROM [Sheet1$] WHERE キャンセル='N';"
    
    ' レコードセットのオープン
       rs.Open strSQL, cn, adOpenKeyset, adlockoptimistic
    
    ' 結果を現在のWorkbookのSheet1にコピー
    Set ws = ThisWorkbook.Sheets("保護者名簿")
    
    ' フィールド名をコピー
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).name
    Next i
    
' コレクションを使って重複を削除
    Set uniqueData = New Collection
    On Error Resume Next
    For j = 0 To UBound(Data, 2)
        Key = Data(0, j) & "|" & Data(1, j) & "|" & Data(2, j) ' キーとして結合
        uniqueData.Add Data, Key
    Next j
    On Error GoTo 0
        ' データをコピー
    ws.Cells(2, 1).CopyFromRecordset rs
    
    ' ユニークなデータをシートに書き込む
    For j = 1 To uniqueData.Count
        For i = 0 To UBound(Data, 1)
            ws.Cells(j + 1, i + 1).Value = uniqueData(j)(i, j - 1)
        Next i
    Next j
    
    ' オブジェクトのクリーンアップ
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub
