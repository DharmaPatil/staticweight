VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStaticWeight 
   Caption         =   "静态称重程序1.00"
   ClientHeight    =   8970
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11985
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox GdhSamp1 
      Height          =   480
      Left            =   8400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   70
      Top             =   8280
      Width           =   1200
   End
   Begin VB.CommandButton com_Ctrl 
      Caption         =   "打开串口"
      Height          =   375
      Left            =   6600
      TabIndex        =   60
      Top             =   8280
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   11280
      TabIndex        =   57
      Text            =   "Combo1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9840
      Top             =   8280
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   1455
      Index           =   0
      Left            =   10200
      TabIndex        =   55
      Top             =   5520
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "数据已经保存"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   -840
         TabIndex        =   56
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "行车方向"
      Height          =   855
      Left            =   120
      TabIndex        =   52
      Top             =   8040
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "右行"
         Height          =   495
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "左行"
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7095
      Left            =   8040
      TabIndex        =   50
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12515
      _Version        =   393216
      Rows            =   100
      Cols            =   5
   End
   Begin VB.Frame reslutpanel 
      Caption         =   "结果:"
      Height          =   2175
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   7815
      Begin VB.CommandButton Command3 
         Caption         =   "保存系数"
         Height          =   375
         Left            =   3240
         TabIndex        =   64
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存数据"
         Height          =   375
         Left            =   720
         TabIndex        =   51
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "记录BC台面车重"
         Height          =   495
         Left            =   6000
         TabIndex        =   49
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "记录AB台面车重"
         Height          =   495
         Left            =   6000
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox boardAB 
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         Text            =   "0.000"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox boardCB 
         Height          =   375
         Left            =   3960
         TabIndex        =   42
         Text            =   "0.000"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "吨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   59
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "吨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   58
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "A 台面 + B 台面重量:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "C 台面 + B 台面重量:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "通道平衡系数"
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   7815
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   11
         Left            =   6120
         TabIndex        =   30
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   10
         Left            =   6120
         TabIndex        =   29
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   9
         Left            =   4920
         TabIndex        =   28
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   8
         Left            =   4920
         TabIndex        =   27
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   7
         Left            =   3720
         TabIndex        =   26
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   6
         Left            =   3720
         TabIndex        =   25
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   24
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   23
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   22
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "传感器11"
         Height          =   255
         Left            =   6120
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "传感器10"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "传感器9"
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "传感器8"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "传感器7"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "传感器6"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "传感器5"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "传感器4"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "传感器3"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "传感器2"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "传感器1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "传感器0"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "台面C"
      Height          =   3285
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
      Begin VB.CommandButton factorc 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   480
         TabIndex        =   63
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Ctozero 
         Caption         =   "台面C回零"
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtScaleC 
         Height          =   375
         Left            =   1080
         TabIndex        =   37
         Text            =   "1000.0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox board 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "C 台面重量"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame boardB 
      Caption         =   "台面B"
      Height          =   3285
      Left            =   2640
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
      Begin VB.CommandButton factorb 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   720
         TabIndex        =   62
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Btozero 
         Caption         =   "台面B回零"
         Height          =   375
         Left            =   720
         TabIndex        =   46
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtScaleB 
         Height          =   375
         Left            =   1440
         TabIndex        =   36
         Text            =   "1000.0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox board 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Text            =   "0.0"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "B 台面重量"
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame boardA 
      Caption         =   "台面A"
      Height          =   3285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
      Begin VB.CommandButton factora 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   360
         TabIndex        =   61
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton Atozero 
         Caption         =   "台面A回零"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   35
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtScaleA 
         Height          =   285
         HideSelection   =   0   'False
         Left            =   1320
         TabIndex        =   34
         Text            =   "1000.0"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox board 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Text            =   "0.0"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "A 台面重量"
         Height          =   375
         Left            =   720
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame 串口设置 
      Caption         =   "串口设置"
      Height          =   855
      Left            =   2520
      TabIndex        =   65
      Top             =   8040
      Width           =   5415
      Begin VB.ComboBox Combo_baud 
         Height          =   300
         Left            =   2760
         TabIndex        =   66
         Text            =   "57600"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo_ComNum 
         Height          =   315
         Left            =   960
         TabIndex        =   67
         Text            =   "COM1"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label 波特率 
         Caption         =   "波特率："
         Height          =   255
         Left            =   2040
         TabIndex        =   68
         Top             =   360
         Width           =   735
      End
      Begin VB.Label 串口号 
         Caption         =   "串口号："
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   360
         Width           =   735
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10440
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   57600
   End
End
Attribute VB_Name = "frmStaticWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim objFileSystem As Object
Dim ReadAveTimeFile As String      '读取平均次数
Dim strComBuf As String
Dim iAvrTimes As Integer
Dim iRawBuf(100, 16) As Long
Dim iFrameCnt As Integer
Dim iSensor(16) As Long
Dim iBoard(4) As Long
Dim fSensor(16) As Double
Dim fBoard(4) As Double
Dim balance(16) As Double
Dim fScaleA, fScaleB, fScaleC As Double
Dim iRef(16) As Integer
Dim fRef(16) As Double
Dim senn As Boolean
Dim CountRecord As Integer    '记录每页行记录数
Dim num, A, B, C, AB, CB, TrainOnTime As Integer '打印时标题每个标题的长度变量
Dim strPrint As String  '打印标题字符串变量
Dim avgsum(3) As Double
Dim msfrows As Integer
Dim msfrow As Integer
Private m_SaveFilePath As String
Private m_IsCalcRef As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim gdh_Weight_Dbpath As String
Dim ObjSystem As Object
    Dim strDirection As String
    Dim strDirection1 As String
    Dim grow As Integer
    Dim gcol As Integer
    Dim KeyAscii As Integer
    Const ASC_ENTER = 13
Private Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_factora(conMaxNonlinerFactors - 1) As Long
Private m_factorb(conMaxNonlinerFactors - 1) As Long
Private m_factorc(conMaxNonlinerFactors - 1) As Long

Dim startFlag As Boolean



'Private Sub avrTime_Change()
'    If Val(avrTime.Text) > 20 Or Val(avrTime.Text) < 1 Then
'        MsgBox ("平均次数应设置在1至20间!")
'    Else
'        iAvrTimes = Val(avrTime.Text)
'    End If
'End Sub

Private Sub bal_Change(Index As Integer)
    Dim pos As Integer
    
    If Not startFlag Then GoTo ok
    
    If Val(bal(Index).Text) < 0 Or Val(bal(Index).Text) > 10000 Then
        MsgBox ("平均次数应设置在0至10000间!")
    Else
        balance(Index) = Val(bal(Index).Text)
        If Index >= 0 And Index <= 3 Then
            pos = conBalanceAPos + Index Mod 4
        ElseIf Index >= 4 And Index <= 7 Then
            pos = conBalanceBPos + Index Mod 4
        Else
            pos = conBalanceCPos + Index Mod 4
        End If
        
        g_NonlineFactors(pos) = bal(Index).Text
    End If
    
ok:
End Sub

Private Sub cmdSave_Click()
'    Dim sFileName As String
'
'    sFileName = GetSaveFileName
'
'    SaveDataToFile sFileName, MSFlexGrid1
'    FileCopy sFileName, App.Path & "\tempfile.tpr"
    SaveData
End Sub
Public Sub AddLine(data() As Long)
    Dim i As Integer, j As Integer, K As Integer
    
            For i = 0 To 3
               avgsum(1) = avgsum(1) + data(i)
            Next i
            For j = 4 To 7
               avgsum(2) = avgsum(2) + data(j)
            Next j
            For K = 8 To 11
               avgsum(3) = avgsum(3) + data(K)
            Next K
            
    Call FrameExp(data())

End Sub
Public Sub AddTrainCode(sCode As String)
    Call TrainCodeExp(sCode)

End Sub

Private Sub com_ctrl_Click()
    Dim portNum As Integer
    
    portNum = Mid(Combo_ComNum.Text, 4)
    
    On Error GoTo ErrorHandler
    
    If MSComm1.PortOpen = False Then
        Call commOpen
    Else
        Call commClose
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        Case 8005
            MsgBox ("串口" & portNum & "已打开！")
        Case 8002
            MsgBox ("无效串口号！")
        Case Else
            MsgBox (Err.Description)
    End Select
    
    Exit Sub
End Sub

Private Sub cmdStaticWeight_Click()

End Sub

Private Sub Command1_Click()
   Call StaticWeight_Exp(boardAB.Text)
End Sub

Private Sub Command2_Click()
   Call StaticWeight_Exp(boardCB.Text)
End Sub


Private Sub Atozero_Click(Index As Integer)
  Dim i As Integer
    For i = 0 To 3
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i

End Sub


Private Sub Btozero_Click()
    Dim i As Integer
    
    For i = 4 To 7
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i

End Sub


Private Sub Command3_Click()
    Call SaveFactorFile
End Sub

Private Sub Ctozero_Click()
Dim i As Integer
    For i = 8 To 11
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i
End Sub

Private Sub factora_Click()
    g_CurrentConfigPlantform = 0
    
    frmEnter.Show 1, Me
    If frmEnter.IsLogin = True Then
        factorDlg.Show vbModal
    End If
    
    
End Sub

Private Sub factorb_Click()
    g_CurrentConfigPlantform = 1
    
    frmEnter.Show 1, Me
    If frmEnter.IsLogin = True Then
        factorDlg.Show vbModal
    End If

End Sub

Private Sub factorc_Click()
    g_CurrentConfigPlantform = 2
    
    frmEnter.Show 1, Me
    If frmEnter.IsLogin = True Then
        factorDlg.Show vbModal
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim AveTime As String
    Dim temp As String
    Dim ChannelVar As String
    Dim timer  As Integer
    Dim j As Integer
    Dim strLine As String
    Dim Index As Integer
    Dim FileName As String
    Dim pos As Integer
    
    m_DynWeight = False
            
    On Error GoTo ok:

    'baud rate setting
    Combo_baud.AddItem (1200)
    Combo_baud.AddItem (2400)
    Combo_baud.AddItem (4800)
    Combo_baud.AddItem (9600)
    Combo_baud.AddItem (19200)
    Combo_baud.AddItem (38400)
    Combo_baud.AddItem (57600)
    Combo_baud.AddItem (115200)
    Combo_baud.Text = 9600
    
    Combo_ComNum.AddItem ("COM1")
    Combo_ComNum.AddItem ("COM2")
    Combo_ComNum.AddItem ("COM3")
    Combo_ComNum.AddItem ("COM4")
    Combo_ComNum.Text = "COM2"
    
    If MSComm1.PortOpen = True Then
        'shape_comState.FillColor = &HFF00&
        com_Ctrl.Caption = "关闭串口"
    End If
    
    timer = 1
    CountRecord = 1
    Frame2(1).Visible = True
    startFlag = False
'    cmdStaticWeight.Visible = False

    GdhSamp1.Run = True
       

    Index = 0
    Open App.Path & "\AveTime.txt" For Input As #33
         For j = 0 To 17
           Line Input #33, strLine
           If j = 1 Then
            iAvrTimes = strLine
           End If
           
           g_NonlineFactors(Index) = strLine
           Index = Index + 1
           
         Next j
         Line Input #33, strLine
         g_NonlineFactors(Index) = strLine
         Index = Index + 1
         txtScaleA.Text = strLine
         fScaleA = Val(txtScaleA.Text)
         
         Line Input #33, strLine
         g_NonlineFactors(Index) = strLine
         Index = Index + 1
         txtScaleB.Text = strLine
         fScaleB = Val(txtScaleB.Text)
         
         Line Input #33, strLine
         g_NonlineFactors(Index) = strLine
         Index = Index + 1
         txtScaleC.Text = strLine
         fScaleC = Val(txtScaleC.Text)
         
         Do While Not EOF(33) And Index < conMaxFactorNum - 1
            Line Input #33, strLine
            g_NonlineFactors(Index) = strLine
            Index = Index + 1
            
         Loop
    Close #33
    
    For i = 0 To conMaxNonlinerFactors - 1
        m_factora(i) = g_NonlineFactors(conFactorAStartPos + i)
        m_factorb(i) = g_NonlineFactors(conFactorBStartPos + i)
        m_factorc(i) = g_NonlineFactors(conFactorCStartPos + i)
    Next i
    
    m_IsCalcRef = False
    iAvrTimes = g_NonlineFactors(1)
    If (iAvrTimes) > 1000 Or iAvrTimes < 1 Then
        MsgBox ("平均次数应设置在1至1000间!")
    End If
     
    'For i = 0 To 11
    '   balance(i) = g_NonlineFactors(conBalanceAPos + i / 4 + 1)
    'Next i
    For i = 0 To 2
        If i = 0 Then
            pos = conBalanceAPos
        ElseIf i = 1 Then
            pos = conBalanceBPos
        Else
            pos = conBalanceCPos
        End If
        
        For j = 0 To 3
            balance(i * 4 + j) = g_NonlineFactors(pos + j)
        Next j
    Next i

    iFrameCnt = 0
    
    For i = 0 To 11
        bal(i).Text = Format(Val(balance(i)), "###0.0")
    Next i
    
'    txtScaleA.Enabled = False
'    txtScaleB.Enabled = False
'    txtScaleC.Enabled = False
    txtScaleA.Text = g_NonlineFactors(conGradStartPos)
    txtScaleB.Text = g_NonlineFactors(conGradStartPos + 1)
    txtScaleC.Text = g_NonlineFactors(conGradStartPos + 2)
    fScaleA = Val(txtScaleA.Text)
    fScaleB = Val(txtScaleB.Text)
    fScaleC = Val(txtScaleC.Text)
    For i = 0 To 2
        board(i).Text = Format(0, "###0.00") & "吨"
    Next i
    
    m_SaveFilePath = String(255, " ")
    i = GetPrivateProfileString("weight", "savepath", "D:\RW\DATA\VEHICLE", m_SaveFilePath, 255, App.Path & "\weightconfig.ini")
    m_SaveFilePath = Trim(m_SaveFilePath)

    With MSFlexGrid1    '以下是配置表格控件
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        
      
        
        .ColWidth(0) = 600
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
       

        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "车号"
        .TextMatrix(0, 2) = "车型"
        .TextMatrix(0, 3) = "重量"
        .TextMatrix(0, 4) = "速度"

    End With
    MSFlexGrid1.Rows = 2
    
    startFlag = True
    
ok:
    Close #33
'txtScaleA.Enabled = True
'txtScaleB.Enabled = True
'txtScaleC.Enabled = True
'txtScaleA.Visible = True
    If Option1.Value = True Then
        Option2.Value = False
        strDirection = "L"
        strDirection1 = "<--"
    Else
     
        Option2.Value = True
        strDirection = "R"
        strDirection = "-->"
    End If
 gdh_Weight_Dbpath = App.Path & "\"
 Set objFileSystem = CreateObject("scripting.filesystemobject")
 Set ObjSystem = CreateObject("Scripting.FileSystemObject")
End Sub

Public Sub FrameExp(data() As Long)
    Dim i, j As Integer
    Dim tSensor(12) As Long           '保存iSensor(i) - iRef(i)的差值
    Dim sumA, sumB, sumC, sumAB, sumCB As Long

    For i = 0 To 11
        iRawBuf(iFrameCnt, i) = data(i)
    Next i

    If iFrameCnt = iAvrTimes - 1 Then
    
       'calc boradA sensor data
       For i = 0 To 3
           sumA = 0
           For j = 0 To iAvrTimes - 1
               sumA = sumA + iRawBuf(j, i)
           Next j
           
           
            iSensor(i) = sumA / iAvrTimes
            
'            If Not m_IsCalcRef Then
'                iRef(I) = iSensor(I)
'            End If
           
            fSensor(i) = ((iSensor(i) - iRef(i)) * balance(i)) / fScaleA
            tSensor(i) = iSensor(i) - iRef(i) '保存iSensor(i) - iRef(i)的差值
       Next i
       'calc boradB sensor data
       For i = 4 To 7
           sumB = 0
           For j = 0 To iAvrTimes - 1
               sumB = sumB + iRawBuf(j, i)
           Next j
           
           iSensor(i) = sumB / iAvrTimes
'            If Not m_IsCalcRef Then
'                iRef(I) = iSensor(I)
'            End If
           fSensor(i) = ((iSensor(i) - iRef(i)) * balance(i)) / fScaleB
           tSensor(i) = iSensor(i) - iRef(i) '保存iSensor(i) - iRef(i)的差值
       Next i
       'calc boradA sensor data
       For i = 8 To 11
           sumC = 0
           For j = 0 To iAvrTimes - 1
               sumC = sumC + iRawBuf(j, i)
           Next j
           iSensor(i) = sumC / iAvrTimes
'            If Not m_IsCalcRef Then
'                iRef(I) = iSensor(I)
'            End If
           fSensor(i) = ((iSensor(i) - iRef(i)) * balance(i)) / fScaleC
           tSensor(i) = iSensor(i) - iRef(i) '保存iSensor(i) - iRef(i)的差值
       Next i
    
       'calc board data
       For i = 0 To 2
           iBoard(i) = iSensor(4 * i) + iSensor(4 * i + 1) + iSensor(4 * i + 2) + iSensor(4 * i + 3)
           fBoard(i) = fSensor(4 * i) + fSensor(4 * i + 1) + fSensor(4 * i + 2) + fSensor(4 * i + 3)

           fBoard(i) = AdjustData(CInt(i), fBoard(i))
           board(i).Text = Format(fBoard(i), "###0.00")
       Next i
       
      'calc A + B,C + B 台面重量
       boardAB.Text = Format(Val(Mid(board(0).Text, 1, 6)) + Val(Mid(board(1).Text, 1, 6)), "###0.00  ")
       boardCB.Text = Format(Val(Mid(board(2).Text, 1, 6)) + Val(Mid(board(1).Text, 1, 6)), "###0.00  ")
        
'        If Not m_IsCalcRef Then
'            m_IsCalcRef = True
'        End If
       iFrameCnt = 0
    End If
    
    iFrameCnt = iFrameCnt + 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmWeight.RunDynWeight = True
    'Call SaveFactorFile
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If
End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        Case comEvReceive
            strComBuf = strComBuf & MSComm1.Input
            Call ComDataDeal
        Case Else
            Exit Sub
    End Select
End Sub
Public Function commOpen() As Boolean
    On Error GoTo ErrorHandler
    
    Dim portNum As Integer
    commOpen = False
    portNum = Mid(Combo_ComNum.Text, 4)

    MSComm1.CommPort = portNum
    MSComm1.Settings = Combo_baud.Text & ",n,8,1"
    MSComm1.PortOpen = True
   ' shape_comState.FillColor = &HFF00&
    com_Ctrl.Caption = "关闭串口"
    Combo_ComNum.Enabled = False
    Combo_baud.Enabled = False
  
    commOpen = True
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 8005
            MsgBox ("串口" & portNum & "已打开！")
        Case 8002
            MsgBox ("无效串口号！")
        Case Else
            MsgBox (Err.Description)
    End Select
    
    Exit Function
    
End Function

Public Function commClose() As Boolean
    commClose = False
    
    On Error GoTo ErrorHandler
    
    MSComm1.PortOpen = False
   ' shape_comState.FillColor = &H8000000F
    com_Ctrl.Caption = "打开串口"
    Combo_ComNum.Enabled = True
    Combo_baud.Enabled = True
    
    commClose = True
    Exit Function
    
ErrorHandler:
    Select Case Err.Number
        Case 8002
            MsgBox ("无效串口号！")
        Case Else
            MsgBox (Err.Description)
    End Select
    
    Exit Function
    
End Function

Public Sub ComDataDeal()
    Dim PtrHead As Integer
    Dim PtrTail As Integer
    Dim strFrame As String
    Dim length As Integer
    
    Do While InStr(strComBuf, "@") <> 0                     'There is frame header "#" in buffer
        PtrHead = InStr(strComBuf, "@")
        strComBuf = Mid(strComBuf, PtrHead)     'Trim the charactor before "@"
        
        PtrTail = InStr(strComBuf, "&")
        If PtrTail <> 0 Then                                            'Find the frame tailer. Find a full frame
            strFrame = Mid(strComBuf, 1, PtrTail) 'Save the frame to strRcvFrame()
            Call TnFrameExp(strFrame)                                   'Call the frame data analyze function
            strComBuf = Mid(strComBuf, PtrTail + 1)
            length = Len(strComBuf)
        Else                                                            'Find a frame header without tailer
            If Len(strComBuf) >= 200 Then                   'Buffer overflow
                strComBuf = ""                              'Clear buffer
                Exit Sub
            Else
                Exit Sub                                                'Waite the tailer
            End If
        End If
    Loop
    
    If InStr(strComBuf, "&") <> 0 Then                    'There is "&" in buffer without "@",frame error
        strComBuf = ""                                      'Clear buffer
        Exit Sub
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    Dim strin As String
    
    Combo1.Clear
    Select Case MSFlexGrid1.Col
        Case 1
            Combo1.Clear
'            Open "d:\rw\file\发站.txt" For Input As #2
'                Do While Not EOF(2)
'                Line Input #2, strin
'                    Combo1.AddItem strin
'                Loop
'            Close #2
        Case 2
            Combo1.Clear
'            Open "d:\rw\file\到站.txt" For Input As #2
'                Do While Not EOF(2)
'                Line Input #2, strin
'                    Combo1.AddItem strin
'                Loop
'            Close #2
       
    End Select
    
    Combo1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top '移动组合框到网格当前的地方
    Combo1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    
    grow = MSFlexGrid1.Row '保存网格行和列的位置
    If (MSFlexGrid1.Col = 3 Or 5 < MSFlexGrid1.Col And MSFlexGrid1.Col < 11) Then
        GoTo ok
    Else
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
ok:
End Sub

Private Sub Option1_Click()
   If Option1.Value = False Then
      
      Option2.Value = True
      strDirection = "R"
      strDirection1 = "-->"
   Else
      
      Option2.Value = False
      strDirection = "L"
      strDirection1 = "<--"
   End If
End Sub

Private Sub Option2_Click()
   If Option1.Value = False Then
      
      Option2.Value = True
      strDirection = "R"
      strDirection1 = "-->"
   Else
      
      Option2.Value = False
      strDirection = "L"
      strDirection1 = "<--"
   End If
End Sub

Private Sub Timer2_Timer()
Frame2(0).Visible = False
Frame2(1).Visible = False
Timer2.Enabled = False
End Sub

Private Sub txtScaleA_Change()
    If Not startFlag Then GoTo ok
     
    If Val(txtScaleA.Text) < 0 Or Val(txtScaleA.Text) > 10000 Then
        MsgBox ("比例系数应设置在0至10000间!")
    Else
        fScaleA = Val(txtScaleA.Text)
        g_NonlineFactors(conGradStartPos) = txtScaleA.Text
    End If
ok:
End Sub

Private Sub txtScaleB_Change()
    If Not startFlag Then GoTo ok
    
    If Val(txtScaleB.Text) < 0 Or Val(txtScaleB.Text) > 10000 Then
        MsgBox ("比例系数应设置在0至10000间!")
    Else
        fScaleB = Val(txtScaleB.Text)
        g_NonlineFactors(conGradStartPos + 1) = txtScaleB.Text
    End If

ok:
End Sub

Private Sub txtScaleC_Change()
     If Not startFlag Then GoTo ok
     
    If Val(txtScaleC.Text) < 0 Or Val(txtScaleC.Text) > 10000 Then
        MsgBox ("比例系数应设置在0至10000间!")
    Else
        fScaleC = Val(txtScaleC.Text)
        g_NonlineFactors(conGradStartPos + 2) = txtScaleC.Text
    End If
ok:
End Sub
Public Function StaticWeight_Exp(data As String)

     msfrow = msfrow + 1
     MSFlexGrid1.TextMatrix(msfrow, 0) = msfrow
     MSFlexGrid1.TextMatrix(msfrow, 3) = data
     MSFlexGrid1.TextMatrix(msfrow, 4) = str(0)
     
     MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    
End Function
Public Function TrainCodeExp(sCode As String)
    Dim strType As String
    
    strType = Mid(sCode, 1, 7)
    If Mid(strType, 1, 1) = "T" Then
        strType = Mid(strType, 2)
    End If
    
    MSFlexGrid1.TextMatrix(msfrow, 1) = strType
    MSFlexGrid1.TextMatrix(msfrow, 2) = Mid(sCode, 8, 7)

End Function
Function SaveDataToFile(FilePath As String, Grid As MSFlexGrid) ', vDateTime As String, vDirection As String)
    Dim strLine As String
    Dim temp As String
    Dim j As Integer
    Dim i As Integer
    Dim vDatetime As String
    Dim FileNo As Integer
'    On Error GoTo ok

    If Dir(FilePath) <> "" Then
        Kill FilePath
    End If

    If Grid.Rows = 2 Or Grid.TextMatrix(1, 0) = "" Then
        Exit Function
    End If

    FileNo = FreeFile
    
    temp = Right$(FilePath, 16)
    vDatetime = Mid(temp, 1, 4) & "-" & Mid(temp, 5, 2) & "-" & Mid(temp, 7, 2)
    vDatetime = vDatetime & " " & Mid(temp, 9, 2) & ":" & Mid(temp, 11, 2)
    Open "d:\staticweight.txt" For Output As #22
 '   Open FilePath For Output As #fileNo
'    txtDateTime.Caption = "正在保存数据，请等待．．．"
    '存标志
    Print #22, "GDHW"
    '存时间
    Print #22, vDatetime
    '存方向
'    Print #FileNo, lblDirection.Caption
   ' Print #22, "-->"
    '存数量
    Print #22, Trim(str(Grid.Rows - 2))
    '存表头
    strLine = "序号" + "|" + "车号" + "|" + "车型" + "|" + "毛重" + "|" + "速度" + "|"
    '存数据
    Print #22, strLine
    For i = 1 To Grid.Rows - 2
        strLine = ""
        strLine = strLine + Trim(Grid.TextMatrix(i, 0)) + "|"
        strLine = strLine + Trim(Grid.TextMatrix(i, 1)) + "|"
        strLine = strLine + Trim(Grid.TextMatrix(i, 2)) + "|"
        strLine = strLine + Trim(Grid.TextMatrix(i, 3)) + "|"
        strLine = strLine + Trim(Grid.TextMatrix(i, 4)) + "|"
        Print #22, strLine
    Next i
    Close #22
'    Sleep (10000)
ok:
End Function
Private Function GetSaveFileName() As String
    Dim strtemp As String
    
    strtemp = Format(Now, "yyyymmddhhmm")
    strtemp = m_SaveFilePath & "\" & strtemp & ".txt"
    
    GetSaveFileName = strtemp
End Function
Private Sub SaveData()
    Dim strtime As String
    Dim strTime1 As String
    Dim strTime2 As String

      Dim intRow As Integer
    Dim i, j, length As Integer
    Dim strser As String, str As String, strBC As String
    Dim buffer(1 To 200) As String
    Dim strChoiceFile As String
    Dim str_FilePath As String
    Dim Sign As Boolean
    Dim strtemp As String
    On Error GoTo ok
    
    
   ' strtime = Mid(GdhCarriage1.TrainOnTime, 1, 2) & Mid(GdhCarriage1.TrainOnTime, 4, 2) & Mid(GdhCarriage1.TrainOnTime, 7, 2)
   ' strTime1 = Format(Now, "yyyymmdd") & strtime
    strtemp = Format(Date, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss")
  
    
                    

'    Call CreatDB(Mid(Trim(strtemp), 1, 4))
    Sign = Save_Data_to_gdhys(Trim(strtemp), Trim(strDirection1), MSFlexGrid1, gdh_Weight_Dbpath)
            

   ' response = MsgBox("原始数据已经保存!", vbExclamation, "提示")
    If Sign = True Then
        Frame2(0).Visible = True
        Timer2.Enabled = True
    Else
        Frame2(1).Visible = True
        Timer2.Enabled = True
    End If
     
ok:
   ' Call GdhSaveWeight1.SaveOriginalData(strTime1, strDirection, GdhCarriage1.Grid)
End Sub
Function Save_Data_to_gdhys(strDate_Time As String, v_Direction As String, GD As MSFlexGrid, dbsavePath As String) As Boolean
    Dim db As New ADODB.Connection, rs As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim query As String, table_Name As String, dbName As String
    Dim FulldbPath As String
    Dim K As Integer
    On Error GoTo ok
    
    While GD.TextMatrix(K, 0) <> ""
        K = K + 1
    Wend
    
    If K = 0 Or GD.TextMatrix(1, 0) = "" Then
        Exit Function
    End If
    
    Select Case Mid(strDate_Time, 6, 2)
        Case "01"
            table_Name = "gdh" & "01"
        Case "02"
            table_Name = "gdh" & "02"
        Case "03"
            table_Name = "gdh" & "03"
        Case "04"
            table_Name = "gdh" & "04"
        Case "05"
            table_Name = "gdh" & "05"
        Case "06"
            table_Name = "gdh" & "06"
        Case "07"
            table_Name = "gdh" & "07"
        Case "08"
            table_Name = "gdh" & "08"
        Case "09"
            table_Name = "gdh" & "09"
        Case "10"
            table_Name = "gdh" & "10"
        Case "11"
            table_Name = "gdh" & "11"
        Case "12"
            table_Name = "gdh" & "12"
        Case Else
            Exit Function
    End Select
    
'    dbName = "gdhys" & Mid(strDate_Time, 1, 4) & ".mdb"
    dbName = "gdhys.mdb"
    table_Name = "gdhys"
    FulldbPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FulldbPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
'    Set db = OpenDatabase(dbsavePath & dbName, False, False, "; pwd=1")
    query = "select * from " & table_Name & " where 日期时间='" & strDate_Time & "'"
'    Set rs = db.OpenRecordset(Query)
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
    End If
    
    For i = 1 To K - 1
        rs.AddNew
        rs.Fields("序号") = Int(Val(GD.TextMatrix(i, 0)))
        rs.Fields("车型") = Trim(GD.TextMatrix(i, 2))
        rs.Fields("车号") = Trim(GD.TextMatrix(i, 1))
       ' If Val(Trim(GD.TextMatrix(I, 3))) > 50 Then
        rs.Fields("毛重") = Trim(GD.TextMatrix(i, 3))
       ' Else
       ' rs.Fields("皮重") = Trim(GD.TextMatrix(I, 3))
       ' End If
        rs.Fields("速度") = Trim(GD.TextMatrix(i, 4))
'        If Val(Trim(GD.TextMatrix(I, 3))) > 50 Then
'        rs.Fields("轻重车") = "√"
'        End If
        rs.Fields("方向") = v_Direction
        rs.Fields("日期时间") = strDate_Time
        rs.Update
    Next i
    rs.Close
    
    '添加索引记录
    query = "select * from gdhindex where 日期时间='" & strDate_Time & "'"
'    Set rs = db.OpenRecordset(Query)
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
    Else
        rs.AddNew
        rs.Fields("车数") = Trim(str(K - 1))
        rs.Fields("日期时间") = strDate_Time
        rs.Fields("方向") = v_Direction
        rs.Update
    End If
    rs.Close
    
    db.Close
    Save_Data_to_gdhys = True
    Exit Function
ok:
End Function
Function CreatDB(strYear As String)
    Dim dbName As String
    Dim DBFullPath As String
    On Error GoTo ok
    
    If Len(strYear) <> 4 Then Exit Function
    
    dbName = "gdhys" + strYear + ".mdb"
    DBFullPath = gdh_Weight_Dbpath + dbName
    If ObjSystem.FileExists(DBFullPath) = False Then
        If ObjSystem.FileExists(App.Path + "\gdhys.mdb") = True Then
            FileCopy App.Path + "\gdhys.mdb", DBFullPath
        End If
    End If
    
'    dbName = "gdhdatamy" + strYear + ".mdb"
'    DBFullPath = gdh_Weight_Dbpath + dbName
'    If ObjSystem.FileExists(DBFullPath) = False Then
'        If ObjSystem.FileExists(App.Path + "\" + "dbs" + "\gdhdatamy.mdb") = True Then
'            FileCopy App.Path + "\" + "dbs" + "\gdhdatamy.mdb", DBFullPath
'        End If
'    End If
'
'    dbName = "gdhdataother" + strYear + ".mdb"
'    DBFullPath = gdh_Weight_Dbpath + dbName
'    If ObjSystem.FileExists(DBFullPath) = False Then
'        If ObjSystem.FileExists(App.Path + "\" + "dbs" + "\gdhdataother.mdb") = True Then
'            FileCopy App.Path + "\" + "dbs" + "\gdhdataother.mdb", DBFullPath
'        End If
'    End If
    
    Exit Function
ok:
'    MsgBox "原始文件不存在,请查看"
End Function

Private Sub Combo1_Click()
    MSFlexGrid1.Text = Combo1.Text
    Combo1.Visible = True
    Combo1.ZOrder 0 ' 把 Combo1 放到最前面！
    Combo1.SetFocus
    If KeyAscii <> ASC_ENTER Then
        SendKeys Chr$(KeyAscii) '判断键盘所按下的键是否是回车键
    End If
End Sub
'************************************************************************************
'Combo1 KeyPress down "enter" form display
'************************************************************************************
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim strin As String
    Dim StrMSFG As Single
    
    If KeyAscii = vbKeyEscape Then '如果按下的是“ESC”键
        Combo1.Visible = False '隐藏Combo1
        MSFlexGrid1.SetFocus '获得焦点
        Exit Sub
    End If
    If KeyAscii = ASC_ENTER Then '如果按下的是“ENTER”键
        MSFlexGrid1.TextMatrix(grow, gcol) = Combo1.Text '列表框的值附给网格单元
        Combo1.Visible = False '隐藏Combo1
        MSFlexGrid1.SetFocus '获得焦点
        KeyAscii = 0 ' 忽视按下键
      
        
        
        Dim tmpRow As Integer
        Dim tmpCol As Integer
        
        tmpRow = MSFlexGrid1.Row '保存网格的当前行和列
        tmpCol = MSFlexGrid1.Col
        MSFlexGrid1.Row = grow '在文本框失去焦点之后，设置网格行和列
        MSFlexGrid1.Col = gcol
        MSFlexGrid1.Text = Combo1.Text '隐藏组合框
        Combo1.SelStart = 0 ' Return caret to beginning.
        Combo1.Visible = False ' Disable text box.
        MSFlexGrid1.Row = tmpRow + 1 '回到行和列的内容
        MSFlexGrid1.Col = gcol
        '***************************************************************
    
    End If
End Sub
'*****************************************************************************
'Combo1 LostFocus
'*****************************************************************************
Private Sub Combo1_LostFocus()
    Dim tmpRow As Integer
    Dim tmpCol As Integer
    
    tmpRow = MSFlexGrid1.Row '保存网格的当前行和列
    tmpCol = MSFlexGrid1.Col
    
    MSFlexGrid1.Row = grow '在文本框失去焦点之后，设置网格行和列
    MSFlexGrid1.Col = gcol
    
    MSFlexGrid1.Text = Combo1.Text '隐藏组合框
    Combo1.SelStart = 0 ' Return caret to beginning.
    Combo1.Visible = False ' Disable text box.
    MSFlexGrid1.Row = tmpRow '回到行和列的内容
'    if tmpCol = me.MSFlexGrid1.R
    MSFlexGrid1.Col = tmpCol
End Sub

Private Sub SaveFactorFile()
    Dim i As Integer
    
    Open App.Path & "\AveTime.txt" For Output As #33
    
    For i = 0 To conMaxFactorNum
        Print #33, g_NonlineFactors(i)
    Next i
    
    Close #33
End Sub

Private Function AdjustData(pt As Integer, Value As Double) As Double
    Dim ret As Double
    Dim Index As Integer
    Dim rev As Double
    
    Index = CInt((Value - 10) / 2)
    
    If pt = 0 Then
        rev = CDbl(m_factora(Index) / 10000)
    ElseIf pt = 1 Then
        rev = CDbl(m_factorb(Index) / 10000)
    Else
        rev = CDbl(m_factorc(Index) / 10000)
    End If
    
    ret = Value * rev
    
    AdjustData = ret
    Exit Function
End Function

Private Sub TnFrameExp(data As String)

    Dim i, j As Integer
End Sub

Private Sub GdhSamp1_OnRcvLine(data() As Long, ByVal cnt As Long)
    AddLine data
ok:
End Sub
