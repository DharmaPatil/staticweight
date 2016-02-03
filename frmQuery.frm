VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQuery 
   Caption         =   "数据查询"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "query"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11250
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton okBtn 
      Caption         =   "确定"
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin staticWeight.GdhPrintWeight GdhPrintWeight1 
      Left            =   2040
      Top             =   1080
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin VB.CommandButton deleteBtn 
      Caption         =   "退出"
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton printBtn 
      Caption         =   "打印"
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid detailGrid 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   120
      Cols            =   12
   End
   Begin MSFlexGridLib.MSFlexGrid indexGrid 
      Height          =   1815
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   60000
      Cols            =   5
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询条件"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton cndAll 
         Caption         =   "全部数据"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton cndConstrat 
         Caption         =   "对比数据"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton cndWeight 
         Caption         =   "过衡数据"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim titleWeight As Variant
Dim titleConstrat As Variant
Dim tableKey() As String
Dim strDateTime As String
Dim strPrintType As String
Dim strDirection As String
Dim m_tareFile As String
Dim m_ok As Boolean

'常量定义
Private Const conWeightTitleLen = 7
Private Const conConstratTitleLen = conFixedCols

Public Property Get TareFile() As String
    TareFile = m_tareFile
End Property
Public Property Get exitStatus() As Boolean
    exitStatus = m_ok
End Property

Private Sub cndAll_Click()
    cndWeight.Value = False
    cndConstrat.Value = False
    cndAll.Value = True

    Call setHeader("all")
    Call queryIndex("all")
    
    detailGrid.rows = 2
End Sub

Private Sub cndConstrat_Click()
    cndWeight.Value = False
    cndConstrat.Value = True
    cndAll.Value = False

    Call setHeader("constrat")
    Call queryIndex("constrat")
    
    detailGrid.rows = 2
End Sub

Private Sub cndWeight_Click()
    cndWeight.Enabled = True
    cndConstrat.Value = False
    cndAll.Value = False
    
    Call setHeader("gdhys")
    Call queryIndex("gdhys")
    
    detailGrid.rows = 2
End Sub

Private Sub deleteBtn_Click()
    m_ok = False
    Unload Me
End Sub

Private Sub Form_Initialize()
    titleWeight = Array("序号", "车型", "车号", "毛重", "速度", "方向", "日期时间")
    titleConstrat = Array("序号", "位号", "流量计值", "检尺值", "车号", "毛重(T)", "皮重(T)", "净重(T)", "流减净(T)", "表差/净(‰)", "尺差/净(‰)")
    
    cndWeight.Value = True
    cndConstrat.Value = False
    cndAll.Value = False
    m_ok = False
    
End Sub

Private Sub Form_Load()
    Dim i, j As Integer
    
    With indexGrid
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        
        .ColWidth(0) = 600
        .ColWidth(1) = 2000
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
    End With
    
    Call setIndexHeader
    indexGrid.rows = 2
    
    With detailGrid
        For i = 0 To conFixedCols - 1
            .ColAlignment(i) = 1
            .ColWidth(i) = 1200
        Next i
        .ColWidth(0) = 600
    End With
    
    Call setHeader("gdhys")
    detailGrid.rows = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub indexGrid_Click()
    Dim Row As Integer
    Dim key As String
    Dim name As String
    
    Row = indexGrid.Row
    key = Trim(indexGrid.TextMatrix(Row, 1))
    strDirection = Trim(indexGrid.TextMatrix(Row, 4))
    strDateTime = key
    m_tareFile = key
    
    If UBound(tableKey) >= Row Then
        name = tableKey(Row)
        strPrintType = name
        Call queryDetail(name, key)
    End If
End Sub

Private Sub okBtn_Click()
    m_ok = True
    Unload Me
End Sub

Private Sub printBtn_Click()
    Dim i As Integer
    
    If strPrintType = "constrat" Then
        GdhPrintWeight1.PrintConstratData strDateTime, detailGrid
    Else
        GdhPrintWeight1.PrintOriginalData strDateTime, strDirection, detailGrid
    End If
    'MsgBox ("打印数据")
End Sub
Private Sub setHeader(key As String)
    Dim i As Integer
    
    detailGrid.Clear
    With detailGrid
        If ((cndWeight.Value = True) Or (cndAll.Value = True And key = "gdhys")) Then
            For i = 0 To conWeightTitleLen - 1
                .TextMatrix(0, i) = titleWeight(i)
            Next i
        Else
            For i = 0 To conConstratTitleLen - 1
                .TextMatrix(0, i) = titleConstrat(i)
            Next i
        End If
    End With
End Sub
Private Sub setIndexHeader()
    Dim i As Integer
    
    indexGrid.Clear
    With indexGrid
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "日期时间"
        .TextMatrix(0, 2) = "节数"
        .TextMatrix(0, 3) = "总重"
        .TextMatrix(0, 4) = "方向"
    End With
End Sub

Private Sub clearTableName()
    'Dim i As Integer
    'For i = CInt(LBound(tableName)) To CInt(UBound(tableName))
    '    tableName(i) = ""
    'Next i
End Sub
Private Sub queryIndex(table As String)
    Dim dbName As String
    Dim tableName As String
    Dim fullDBPath As String
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim rows As Integer
    Dim i, curRow As Integer
    
    On Error GoTo ok
    
    Call setIndexHeader
    indexGrid.rows = 2
        
    dbName = "gdhys.mdb"
    tableName = "gdhindex"
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    If table = "all" Then
        query = "select * from " & tableName
    Else
        query = "select * from " & tableName & " where 表名='" & table & "'"
    End If
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        ReDim tableKey(1 To rs.RecordCount)
        
        rows = 0
        curRow = 1
        
        Do While Not rs.EOF
            
            rows = rows + 1
            'With indexGrid
            '    .TextMatrix(rows, 0) = CStr(rows)
            '    .TextMatrix(rows, 1) = CStr(rs.Fields("日期时间").Value)
            '    .TextMatrix(rows, 2) = CStr(rs.Fields("车数").Value)
            '    .TextMatrix(rows, 3) = CStr(rs.Fields("总重").Value)
            '    .TextMatrix(rows, 4) = CStr(rs.Fields("方向").Value)
            'End With
            
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
            
            tableKey(rows) = rs.Fields("表名").Value
            indexGrid.rows = indexGrid.rows + 1
            curRow = 1
            
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    db.Close
ok:
End Sub

Private Sub queryDetail(table As String, key As String)
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
    
    Call setHeader(table)
    detailGrid.rows = 2
        
    dbName = "gdhys.mdb"
    tableName = table
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    query = "select * from " & tableName & " where 日期时间='" & key & "'"

    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        cols = rs.Fields.Count
                
        rows = 0
        Do While Not rs.EOF
            rows = rows + 1
            With detailGrid
                .TextMatrix(rows, 0) = CStr(rows)
                For j = 1 To cols - 1
                    .TextMatrix(rows, j) = CStr(rs.Fields(j).Value)
                Next j
            End With
            detailGrid.rows = detailGrid.rows + 1
            
            rs.MoveNext
        Loop
    End If
    
ok:
    rs.Close
    db.Close
End Sub


