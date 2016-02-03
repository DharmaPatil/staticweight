VERSION 5.00
Begin VB.Form adduserDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加用户"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtFieldom 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtPassWord 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "权  限"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "密  码"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "adduserDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum PopedomEnum
    Administrator = 0
    User = 1
    Guest = 2
End Enum

'
Private m_IsLogin As Boolean
Private m_UserName As String
Private m_UserPopedom As PopedomEnum
Private m_return As Integer
Private adoConnection As ADODB.Connection
Private adoRecordset As ADODB.Recordset

Public Property Get ReturnStatus()
    ReturnStatus = m_return
End Property
Private Sub CancelButton_Click()
    m_return = StatusExit
    adoRecordset.Close
    Unload Me
End Sub

Private Sub Form_Initialize()
    m_IsLogin = False
    m_return = StatusExit
End Sub

Private Sub Form_Load()
    Dim strCommand As String
    
    m_IsLogin = False
    m_UserName = ""
    m_UserPopedom = PopedomEnum.Guest
    
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    adoConnection.CursorLocation = adUseClient
    adoConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\gdh.mdb"
    strCommand = "SELECT * FROM tblUser"
    If adoConnection.State = 0 Then
        MsgBox "连接超时"
        End
    End If
    
    adoRecordset.Open strCommand, adoConnection, adOpenDynamic, adLockOptimistic
End Sub

Private Sub OKButton_Click()
    Dim name As String
    Dim pass As String
    Dim field As String

    name = Trim(txtUserName.Text)
    pass = Trim(txtPassWord.Text)
    field = m_UserPopedom
    
    If name = "" Or pass = "" Or field = "" Then
        GoTo ok
    End If
    
    adoRecordset.MoveFirst
    Do While Not adoRecordset.EOF
        If adoRecordset.Fields("UserName") = name Then
            m_IsLogin = True
            If (MsgBox("该用户已经存在,是否要修改密码？", vbYesNo, "提示") = vbYes) Then
                adoRecordset.Fields("PassWord") = pass
                adoRecordset.Update
            End If
            m_return = StatusOk
            GoTo ok
        End If
        
        adoRecordset.MoveNext
    Loop
    
    If Not m_IsLogin Then
        With adoRecordset
            .AddNew
            .Fields("UserName") = name
            .Fields("PassWord") = pass
            .Fields("Popedom") = field
            .Update
        End With
    End If
    m_return = StatusOk

ok:
    adoRecordset.Close
    Unload Me
End Sub
