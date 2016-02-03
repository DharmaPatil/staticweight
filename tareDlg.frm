VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tareDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "皮重数据"
   ClientHeight    =   5715
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid detailGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   200
      Cols            =   7
   End
   Begin VB.ComboBox rangeCbx 
      Height          =   300
      ItemData        =   "tareDlg.frx":0000
      Left            =   7080
      List            =   "tareDlg.frx":0002
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid indexGrid 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   60000
      Cols            =   5
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "时间范围"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "tareDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_tareFile As String
Private m_ok As Boolean
Private header As Variant

Public Property Get TareFile() As String
    TareFile = m_tareFile
End Property
Public Property Get exitStatus() As Boolean
    exitStatus = m_ok
End Property
Private Sub CancelButton_Click()
    m_ok = False
    Unload Me
End Sub

Private Sub Form_Initialize()
    header = Array("序号", "车型", "车号", "毛重", "速度", "方向", "日期时间")
    m_ok = False
End Sub

Private Sub indexGrid_Click()
    Dim Row As Integer
    Dim key As String
    
    Row = indexGrid.Row
    m_tareFile = Trim(indexGrid.TextMatrix(Row, 1))
    
    queryDetail "gdhys", m_tareFile
End Sub

Private Sub OKButton_Click()
    m_ok = True
    Unload Me
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
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
        
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "日期时间"
        .TextMatrix(0, 2) = "节数"
        .TextMatrix(0, 3) = "总重"
        .TextMatrix(0, 4) = "方向"
    End With
    indexGrid.rows = 2

    With detailGrid
        For i = 0 To 6
            .ColAlignment(i) = 1
            .ColWidth(i) = 1200
            .TextMatrix(0, i) = header(i)
        Next i
        .ColWidth(0) = 600
    End With
    detailGrid.rows = 2
    
    rangeCbx.AddItem ("1")
    rangeCbx.AddItem ("2")
    rangeCbx.AddItem ("3")
    
    rangeCbx.Text = "3"
    rangeCbx.TabIndex = 2
    
    queryIndex "3"
End Sub


Private Sub queryIndex(range As String)
    Dim dbName As String
    Dim tableName As String
    Dim fullDBPath As String
    Dim db As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim rows As Integer
    Dim strDate As Variant
    Dim selDate As Date
    Dim strSelKey As String
    Dim selDateRange(3) As String
    Dim strTemp As Variant
    Dim i, curRow As Integer
    Dim findFlag As Boolean
    
    On Error GoTo ok
    
    indexGrid.rows = 2
    curRow = 1
        
    dbName = "gdhys.mdb"
    tableName = "gdhindex"
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
    query = "select * from " & tableName
    
    strDate = Split(Trim(g_TareStartDate), " ")
    selDate = CDate(strDate(0))
    
    For i = 0 To CInt(range)
        strSelKey = Format((selDate + CInt(i)), "yyyy-mm-dd")
        selDateRange(i) = strSelKey
    Next i
    
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    findFlag = False
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        rows = 0
        Do While Not rs.EOF
            strTemp = Split(CStr(rs.Fields("日期时间").Value), " ")
            For i = 0 To CInt(range)
                If selDateRange(i) = strTemp(0) Then
                    findFlag = True
                    Exit For
                End If
            Next i
            
            If findFlag Then
                rows = rows + 1
                
                'With indexGrid
                '    .TextMatrix(rows, 0) = CStr(rows)
                '    .TextMatrix(rows, 1) = CStr(rs.Fields("日期时间").Value)
                '    .TextMatrix(rows, 2) = CStr(rs.Fields("车数").Value)
                '    .TextMatrix(rows, 3) = CStr(rs.Fields("总重").Value)
                '    .TextMatrix(rows, 4) = CStr(rs.Fields("方向").Value)
                'End With
                'indexGrid.rows = indexGrid.rows + 1
                
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
                
                findFlag = False
                curRow = 1
            End If
            
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
    
    'Call setHeader(table)
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

Private Sub rangeCbx_Click()
    Dim sel As String
    
    sel = rangeCbx.Text
    queryIndex sel
End Sub

Private Sub setHeader()
    Dim i As Integer
    
    detailGrid.Clear
    With detailGrid
        For i = 0 To 6
            .TextMatrix(0, i) = header(i)
        Next i
    End With
End Sub
