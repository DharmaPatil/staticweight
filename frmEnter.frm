VERSION 5.00
Begin VB.Form frmEnter 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton adduserBtn 
      Caption         =   "添加用户(&A)"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "登录(&O)"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtPassWord 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户密码："
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户姓名："
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum PopedomEnum
    Administrator = 0
    User = 1
    Guest = 2
End Enum

'
Private m_IsLogin As Boolean
Private m_UserName As String
Private m_UserPopedom As PopedomEnum
Private adoConnection As ADODB.Connection
Private adoRecordset As ADODB.Recordset

Public Property Get IsLogin() As Boolean
    IsLogin = m_IsLogin
End Property
Public Property Get LoginName() As String
    LoginName = m_UserName
End Property


'解密处理
Private Function DecodeString(ByVal sString As String) As String
    Dim nLoopCount As Integer
    Dim sFinal As String
    
    sFinal = ""
    For nLoopCount = 1 To Len(sString)
        sFinal = sFinal + Chr$(Asc(Mid$(sString, nLoopCount, 1)) - 70)
    Next nLoopCount
    
    DecodeString = sFinal
End Function

Private Sub adduserBtn_Click()
    Dim strCommand As String
    
    Me.Hide
    adduserDlg.Show vbModal
    If adduserDlg.ReturnStatus = ReturnStatus.StatusOk Then
        'adoRecordset.Close
        'strCommand = "SELECT * FROM tblUser"
        'adoRecordset.Open strCommand, adoConnection, adOpenDynamic, adLockOptimistic
        MsgBox "程序将关闭,请选择重新运行程序！", vbInformation, "警告"
        Unload Me
        GoTo ok
    End If
    Me.Show
ok:
End Sub

Private Sub cmdCancel_Click()
    m_IsLogin = False
    m_UserName = ""
    m_UserPopedom = Guest
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim name As String
    Dim pass As String
    
    name = Trim(txtUserName.Text)
    pass = Trim(txtPassWord.Text)
    
    adoRecordset.MoveFirst
    Do While Not adoRecordset.EOF
        If adoRecordset.Fields("UserName") = name And (adoRecordset.Fields("PassWord") & "") = pass Then
            m_UserName = name
            m_UserPopedom = adoRecordset.Fields("Popedom")
            m_IsLogin = True
            Exit Do
        Else
            m_IsLogin = False
            adoRecordset.MoveNext
        End If
    Loop
    
    If m_IsLogin Then
        'Unload Me
        If conFactory = Factory.gsh Then
            If (m_UserName = "Administrator") Then
                g_LoginUser = m_UserName
                g_SuperOk = True
            Else
                g_LoginUser = m_UserName
                g_SuperOk = False
            End If
            
            If g_StartLogin Then
                Unload Me
                frmStaticWeight.Show vbModal
            End If
        End If
        
        Unload Me
    Else
        MsgBox "您输入的用户名或密码错误,请重新输入!", vbOKOnly + vbInformation, "提示"
    End If
End Sub

Private Sub Form_Initialize()
    g_StartLogin = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK_Click
        Case vbKeyEscape
            cmdCancel_Click
        Case Else
    End Select
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
