VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConstrat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "轨道衡与流量计对比"
   ClientHeight    =   7095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton tareBtn 
      Caption         =   "取皮重"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   6240
      Width           =   855
   End
   Begin staticWeight.GdhPrintWeight GdhPrintWeight1 
      Left            =   6720
      Top             =   3600
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin VB.CommandButton printBtn 
      Caption         =   "打印"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   6240
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid indexGrid 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10398
      _Version        =   393216
      Rows            =   60000
      Cols            =   5
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   6720
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6855
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12091
      _Version        =   393216
      Rows            =   120
      Cols            =   11
      FixedCols       =   0
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退出"
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "保存"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   6240
      Width           =   855
   End
End
Attribute VB_Name = "frmConstrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim headerTitle(conFixedCols) As String
Dim grow As Integer
Dim gcol As Integer
Dim KeyAscii As Integer
Dim strDateTime As String
Dim curRow As Integer

Const ASC_ENTER = 13

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
    MSFlexGrid1.Text = Combo1.Text
    Combo1.Visible = True
    Combo1.ZOrder 0
    Combo1.SetFocus
    If KeyAscii <> ASC_ENTER Then
        SendKeys Chr$(KeyAscii) '判断键盘所按下的键是否是回车键
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim strin As String
    Dim StrMSFG As Single
    Dim tmpVal As Double
    Dim Percent As Integer
    
    If KeyAscii = vbKeyEscape Then
        Combo1.Visible = False
        MSFlexGrid1.SetFocus
        Exit Sub
    End If
    If KeyAscii = ASC_ENTER Then
        MSFlexGrid1.TextMatrix(grow, gcol) = Combo1.Text
        Combo1.Visible = False
        MSFlexGrid1.SetFocus
        KeyAscii = 0
           
        If gcol = 2 Then
            tmpVal = Val(MSFlexGrid1.TextMatrix(grow, gcol)) - Val(MSFlexGrid1.TextMatrix(grow, 7))
            MSFlexGrid1.TextMatrix(grow, 8) = Format(tmpVal, "###0.000")
        End If
        
        If gcol = 6 Then
            tmpVal = Val(MSFlexGrid1.TextMatrix(grow, 5)) - Val(MSFlexGrid1.TextMatrix(grow, gcol))
            MSFlexGrid1.TextMatrix(grow, 7) = Format(tmpVal, "###0.000")
            tmpVal = Val(MSFlexGrid1.TextMatrix(grow, 2)) - Val(MSFlexGrid1.TextMatrix(grow, 7))
            MSFlexGrid1.TextMatrix(grow, 8) = Format(tmpVal, "###0.000")
        End If
        
        If Val(MSFlexGrid1.TextMatrix(grow, 7)) <> 0 Then
            tmpVal = Val(MSFlexGrid1.TextMatrix(grow, 8))
            Percent = (tmpVal / Val(MSFlexGrid1.TextMatrix(grow, 7))) * 1000
            MSFlexGrid1.TextMatrix(grow, 9) = CStr(Percent)
        End If
        
        If gcol = 3 Then
            If Val(MSFlexGrid1.TextMatrix(grow, 7)) <> 0 Then
                tmpVal = Val(MSFlexGrid1.TextMatrix(grow, gcol)) - Val(MSFlexGrid1.TextMatrix(grow, 7))
                Percent = (tmpVal / Val(MSFlexGrid1.TextMatrix(grow, 7))) * 1000
                MSFlexGrid1.TextMatrix(grow, 10) = CStr(Percent)
            End If
        End If
        
        If gcol = 4 Then
            MSFlexGrid1.TextMatrix(grow, 4) = Combo1.Text
        End If
        
        Dim tmpRow As Integer
        Dim tmpCol As Integer
        
        tmpRow = MSFlexGrid1.Row
        tmpCol = MSFlexGrid1.Col
        MSFlexGrid1.Row = grow
        MSFlexGrid1.Col = gcol
        MSFlexGrid1.Text = Combo1.Text
        Combo1.SelStart = 0
        Combo1.Visible = False
        MSFlexGrid1.Row = tmpRow + 1
        MSFlexGrid1.Col = gcol
        '***************************************************************
    
    End If
End Sub

Private Sub Combo1_LostFocus()
    Dim tmpRow As Integer
    Dim tmpCol As Integer
    
    tmpRow = MSFlexGrid1.Row
    tmpCol = MSFlexGrid1.Col
    
    MSFlexGrid1.Row = grow
    MSFlexGrid1.Col = gcol
    
    MSFlexGrid1.Text = Combo1.Text
    'inputTxt.SelStart = 0
    Combo1.Visible = False
    MSFlexGrid1.Row = tmpRow
'    if tmpCol = me.MSFlexGrid1.R
    MSFlexGrid1.Col = tmpCol
    
End Sub

Private Sub Form_Load()
    Dim i As Integer

    curRow = 1
    With MSFlexGrid1
        For i = 0 To conFixedCols - 1
            .ColAlignment(CLng(i)) = 1
            .ColWidth(CLng(i)) = 1000
        Next i
        .ColWidth(0) = 600
        .ColWidth(1) = 600
        
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "位号"
        .TextMatrix(0, 2) = "流量计值"
        .TextMatrix(0, 3) = "检尺值"
        .TextMatrix(0, 4) = "车号"
        .TextMatrix(0, 5) = "毛重(T)"
        .TextMatrix(0, 6) = "皮重(T)"
        .TextMatrix(0, 7) = "净重(T)"
        .TextMatrix(0, 8) = "流减净(T)"
        .TextMatrix(0, 9) = "表差/净(‰)"
        .TextMatrix(0, 10) = "尺差/净(‰)"
        
    End With
    MSFlexGrid1.rows = 2
    
    With indexGrid
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .ColWidth(2) = 500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "日期时间"
        .TextMatrix(0, 2) = "节数"
        .TextMatrix(0, 3) = "总重"
        .TextMatrix(0, 4) = "方向"
    End With
    indexGrid.rows = 2
    
 '   indexGrid.col = 2
 '   indexGrid.Sort = 9
    
    Call initialIndex
            
End Sub

Private Sub indexGrid_Click()
    Dim Row As Integer
    Dim key As String
    Dim name As String
    Dim i As Integer
    
    Row = indexGrid.Row
    key = Trim(indexGrid.TextMatrix(Row, 1))
    curRow = Row
    strDateTime = key
    g_TareStartDate = key
    

    If key <> "" Then
        queryDetail key
    End If
End Sub

Private Sub indexGrid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    Dim i As Integer
    Dim rowStart, rowEnd As Integer
    Dim Col As Integer
    
    rowStart = Row1
    rowEnd = Row2
    Col = Cmp
    
End Sub

Private Sub MSFlexGrid1_Click()
    Dim strin As String
    
    If MSFlexGrid1.Row = MSFlexGrid1.rows - 1 Then
        Exit Sub
    End If
    
    Combo1.Clear
     
    Combo1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top '移动组合框到网格当前的地方
    Combo1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    
    grow = MSFlexGrid1.Row '保存网格行和列的位置
    If (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Or MSFlexGrid1.Col = 6 Or MSFlexGrid1.Col = 4) Then
        gcol = MSFlexGrid1.Col
        Combo1.Width = MSFlexGrid1.CellWidth - 2 * Screen.TwipsPerPixelX '设置文本大小和网格当前的大小一致
        Combo1.Text = MSFlexGrid1.Text '把网格中的内容放到组合框中
        Combo1.Visible = True
        Combo1.ZOrder 0 ' 把 Combo1 放到最前面！
        Combo1.SetFocus
        If KeyAscii <> ASC_ENTER Then
            SendKeys Chr$(KeyAscii)
        End If
    End If
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    Combo1.Clear
    
    If MSFlexGrid1.Row = MSFlexGrid1.rows - 1 Then
        Exit Sub
    End If
     
    Combo1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top '移动组合框到网格当前的地方
    Combo1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    
    grow = MSFlexGrid1.Row '保存网格行和列的位置
    If (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 3 Or MSFlexGrid1.Col = 6 Or MSFlexGrid1.Col = 4) Then
        gcol = MSFlexGrid1.Col
        Combo1.Width = MSFlexGrid1.CellWidth - 2 * Screen.TwipsPerPixelX '设置文本大小和网格当前的大小一致
        Combo1.Text = MSFlexGrid1.Text '把网格中的内容放到组合框中
        Combo1.Visible = True
        Combo1.ZOrder 0 ' 把 Combo1 放到最前面！
        Combo1.SetFocus
        If KeyAscii <> ASC_ENTER Then
            SendKeys Chr$(KeyAscii)
        End If

    End If
End Sub

Private Sub OKButton_Click()
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i, j As Integer
    Dim query As String
    Dim table_Name As String
    Dim dbName As String
    Dim fullDBPath As String
    Dim cnt As Integer
    Dim strDate_Time As String
    
    On Error GoTo ok
    
    If MsgBox("确认保存表格数据", vbOKCancel, "保存") = vbCancel Then
        GoTo ok
    End If
        
    cnt = 0
    While MSFlexGrid1.TextMatrix(cnt, 0) <> ""
        cnt = cnt + 1
    Wend
    
    If cnt = 0 Or MSFlexGrid1.TextMatrix(1, 0) = "" Then
        GoTo ok
    End If
    
    dbName = "gdhys.mdb"
    table_Name = "constrat"
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    'strDate_Time = Format(Date, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss")
    query = "select * from " & table_Name & " where 日期时间='" & strDateTime & "'"
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
    End If
    
    For i = 1 To cnt - 1
        rs.AddNew
        For j = 0 To 10
            If j = 0 Or j = 9 Or j = 10 Then
                rs.Fields(j) = Int(Val(MSFlexGrid1.TextMatrix(i, j)))
            Else
                rs.Fields(j) = MSFlexGrid1.TextMatrix(i, j)
            End If
        Next j
        rs.Fields("日期时间") = strDateTime
        rs.Update
    Next i
    rs.Close
    
    '添加索引记录
    query = "select * from gdhindex where (日期时间='" & strDateTime & "'" & " and 表名='constrat')"
'    Set rs = db.OpenRecordset(Query)
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
    Else
        rs.AddNew
        rs.Fields("表名") = "constrat"
        rs.Fields("日期时间") = indexGrid.TextMatrix(curRow, 1)
        rs.Fields("车数") = indexGrid.TextMatrix(curRow, 2)
        rs.Fields("总重") = indexGrid.TextMatrix(curRow, 3)
        rs.Fields("方向") = indexGrid.TextMatrix(curRow, 4)
        rs.Update
    End If
    rs.Close
    db.Close
    
ok:

End Sub

Private Sub initialIndex()
    Dim dbName As String
    Dim tableName As String
    Dim fullDBPath As String
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim rows As Integer
    Dim i, curRow As Integer
    
    On Error GoTo ok

    indexGrid.rows = 2
        
    dbName = "gdhys.mdb"
    tableName = "gdhindex"
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    query = "select * from " & tableName & " where 表名='gdhys'"
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        
        rows = 0
        curRow = 1
        Do While Not rs.EOF
            rows = rows + 1
            
            If rows = 1 Then
                With indexGrid
                    .TextMatrix(rows, 0) = CStr(rows)
                    .TextMatrix(rows, 1) = CStr(rs.Fields("日期时间").Value)
                    .TextMatrix(rows, 2) = CStr(rs.Fields("车数").Value)
                    .TextMatrix(rows, 3) = CStr(rs.Fields("总重").Value)
                    .TextMatrix(rows, 4) = CStr(rs.Fields("方向").Value)
                End With
            Else
                For i = 1 To indexGrid.rows - 2
                    If (CDate(rs.Fields("日期时间")) > CDate(Trim(CStr(indexGrid.TextMatrix(i, 1))))) Then
                        curRow = i
                        Exit For
                    End If
                Next i
                
                With indexGrid
                    For i = indexGrid.rows - 2 To curRow Step -1
                        .TextMatrix(i + 1, 0) = CStr(i + 1)
                        .TextMatrix(i + 1, 1) = .TextMatrix(i, 1)
                        .TextMatrix(i + 1, 2) = .TextMatrix(i, 2)
                        .TextMatrix(i + 1, 3) = .TextMatrix(i, 3)
                        .TextMatrix(i + 1, 4) = .TextMatrix(i, 4)
                    Next i
                    
                    .TextMatrix(curRow, 0) = CStr(curRow)
                    .TextMatrix(curRow, 1) = CStr(rs.Fields("日期时间").Value)
                    .TextMatrix(curRow, 2) = CStr(rs.Fields("车数").Value)
                    .TextMatrix(curRow, 3) = CStr(rs.Fields("总重").Value)
                    .TextMatrix(curRow, 4) = CStr(rs.Fields("方向").Value)
                End With
            End If
            indexGrid.rows = indexGrid.rows + 1
            
            curRow = 1
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    db.Close
ok:
    
End Sub

Private Sub queryDetail(key As String)
    Dim dbName As String
    Dim tableName As String
    Dim fullDBPath As String
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim rows As Integer
    Dim cols As Integer
    Dim j As Integer
    
    On Error GoTo ok

    MSFlexGrid1.rows = 2
        
    dbName = "gdhys.mdb"
    tableName = "gdhys"
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    query = "select * from " & tableName & " where 日期时间='" & key & "'"

    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst

        rows = 0
        Do While Not rs.EOF
            rows = rows + 1
            With MSFlexGrid1
                For j = 0 To MSFlexGrid1.cols - 1
                    .TextMatrix(rows, j) = ""
                Next j
                
                .TextMatrix(rows, 0) = CStr(rows)
                .TextMatrix(rows, 4) = CStr(rs.Fields(2).Value)
                .TextMatrix(rows, 5) = CStr(rs.Fields(3).Value)
            End With
            MSFlexGrid1.rows = MSFlexGrid1.rows + 1
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    db.Close
ok:
    
End Sub


Private Sub printBtn_Click()
    GdhPrintWeight1.PrintConstratData strDateTime, MSFlexGrid1
    
End Sub

Private Sub tareBtn_Click()
    Dim TareFile As String
    Dim dbName As String
    Dim tableName As String
    Dim fullDBPath As String
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim rows As Integer
    Dim Row As Integer
    Dim tare(200, 1) As String
    Dim key As String
    Dim grsw As String
    Dim i As Integer
    Dim tmp As Single
    'tareDlg.Show vbModal
    
    On Error GoTo ok
    
    'g_QueryMethod = QueryMethod.Constrat
    'frmQuery.Show vbModal
    tareDlg.Show vbModal
    If tareDlg.exitStatus Then
        TareFile = tareDlg.TareFile

        dbName = "gdhys.mdb"
        tableName = "gdhys"
        fullDBPath = App.Path & "\" & dbName
        
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
        
        query = "select * from " & tableName & " where 日期时间='" & TareFile & "'"
    
        rs.Open query, db, adOpenDynamic, adLockOptimistic
        
        '提取数据库中的皮重数据
        Row = 0
        If Not rs.BOF And Not rs.EOF Then
            rs.MoveFirst
            
            Do While Not rs.EOF
                tare(Row, 0) = rs.Fields("车号").Value
                tare(Row, 1) = rs.Fields("毛重").Value
                Row = Row + 1
                rs.MoveNext
            Loop
        End If
        rs.Close
        
        rows = Row
        For i = 1 To MSFlexGrid1.rows - 2
            key = MSFlexGrid1.TextMatrix(i, 4)
            grsw = findTare(tare, rows, key)
            
            MSFlexGrid1.TextMatrix(i, 6) = grsw
            MSFlexGrid1.TextMatrix(i, 7) = CStr(Val(MSFlexGrid1.TextMatrix(i, 5)) - Val(grsw))
            MSFlexGrid1.TextMatrix(i, 8) = CStr(Val(MSFlexGrid1.TextMatrix(i, 2)) - Val(MSFlexGrid1.TextMatrix(i, 7)))
            
            tmp = Val(MSFlexGrid1.TextMatrix(i, 8)) / Val(MSFlexGrid1.TextMatrix(i, 7))
            tmp = tmp * 1000
            MSFlexGrid1.TextMatrix(i, 9) = CStr(CInt(tmp))
            
            tmp = (Val(MSFlexGrid1.TextMatrix(i, 3)) / Val(MSFlexGrid1.TextMatrix(i, 7))) - 1
            tmp = tmp * 1000
            MSFlexGrid1.TextMatrix(i, 10) = CStr(CInt(tmp))
            
        Next i
        
        '删除记录和索引
        'query = "select * from " & tableName & " where 日期时间='" & TareFile & "'"
        'rs.Open query, db, adOpenDynamic, adLockOptimistic
        'If rs.RecordCount > 0 Then
        '    rs.Delete
        '    rs.Update
        'End If
        'rs.Close
        
        'query = "select * from gdhindex" & " where 日期时间='" & TareFile & "'"
        'rs.Open query, db, adOpenDynamic, adLockOptimistic
        'If rs.RecordCount > 0 Then
        '    rs.Delete
        '    rs.Update
        'End If
        
        rs.Close
        db.Close
    End If
ok:
    If tareDlg.exitStatus Then
        MsgBox Error, vbOKOnly, "提示"
    End If

End Sub

Private Function findTare(res() As String, cnt As Integer, find As String) As String
    Dim i As Integer
    Dim grs As String
    Dim key As String
    
    key = find
    For i = 0 To cnt
        If res(i, 0) <> "" And key = res(i, 0) Then
            grs = res(i, 1)
            GoTo ok
        End If
    Next i
    
ok:
    findTare = grs

End Function
