VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStaticWeight 
   Caption         =   "静态称重程序1.01"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13800
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13800
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame6 
      Height          =   1215
      Left            =   0
      TabIndex        =   73
      Top             =   1080
      Width           =   7935
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   11
         Left            =   6360
         TabIndex        =   85
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   10
         Left            =   6360
         TabIndex        =   84
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   9
         Left            =   5280
         TabIndex        =   83
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   8
         Left            =   5280
         TabIndex        =   82
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   7
         Left            =   3840
         TabIndex        =   81
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   80
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   79
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   4
         Left            =   2760
         TabIndex        =   78
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   77
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   76
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   75
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   74
         Top             =   240
         Width           =   855
      End
   End
   Begin staticWeight.GdhSamp GdhSamp1 
      Left            =   10560
      Top             =   8040
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin staticWeight.GdhCode GdhCode1 
      Left            =   9720
      Top             =   8160
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin staticWeight.GdhPrintWeight GdhPrintWeight1 
      Left            =   11280
      Top             =   8040
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存数据"
      Height          =   615
      Left            =   3000
      TabIndex        =   63
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "保存系数"
      Height          =   615
      Left            =   4560
      TabIndex        =   62
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   10200
      TabIndex        =   56
      Text            =   "Combo1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8280
      Top             =   8160
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   1455
      Index           =   0
      Left            =   8880
      TabIndex        =   54
      Top             =   5040
      Visible         =   0   'False
      Width           =   3735
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
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "行车方向"
      Height          =   1095
      Left            =   8040
      TabIndex        =   51
      Top             =   1200
      Width           =   5655
      Begin VB.OptionButton Option2 
         Caption         =   "右行"
         Height          =   495
         Left            =   2880
         TabIndex        =   53
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "左行"
         Height          =   495
         Left            =   840
         TabIndex        =   52
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6855
      Left            =   8040
      TabIndex        =   50
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12091
      _Version        =   393216
      Rows            =   100
      Cols            =   5
   End
   Begin VB.Frame reslutpanel 
      Caption         =   "结果:"
      Height          =   1695
      Left            =   0
      TabIndex        =   41
      Top             =   2400
      Width           =   7935
      Begin VB.CommandButton Command2 
         Caption         =   "记录BC台面车重"
         Height          =   495
         Left            =   6000
         TabIndex        =   49
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "记录AB台面车重"
         Height          =   495
         Left            =   6000
         TabIndex        =   48
         Top             =   360
         Width           =   1815
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
         TabIndex        =   58
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
         TabIndex        =   57
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
      Left            =   0
      TabIndex        =   6
      Top             =   7680
      Width           =   7935
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   11
         Left            =   6600
         TabIndex        =   30
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   10
         Left            =   6600
         TabIndex        =   29
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   9
         Left            =   5400
         TabIndex        =   28
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   8
         Left            =   5400
         TabIndex        =   27
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   7
         Left            =   3960
         TabIndex        =   26
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   6
         Left            =   3960
         TabIndex        =   25
         Text            =   "1.0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   24
         Text            =   "1.0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox bal 
         Height          =   285
         Index           =   4
         Left            =   2880
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
         Left            =   6600
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "传感器10"
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "传感器9"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "传感器8"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "传感器7"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "传感器6"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "传感器5"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "传感器4"
         Height          =   255
         Left            =   2880
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
      Left            =   5400
      TabIndex        =   2
      Top             =   4200
      Width           =   2535
      Begin VB.TextBox txtNonlinear 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   92
         Text            =   "1.0000"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton factorc 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   480
         TabIndex        =   61
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
         Top             =   1200
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
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "补偿系数"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   91
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "C 台面重量"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame boardB 
      Caption         =   "台面B"
      Height          =   3285
      Left            =   2520
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
      Begin VB.TextBox txtNonlinear 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   90
         Text            =   "1.0000"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton factorb 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   720
         TabIndex        =   60
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
         Left            =   1200
         TabIndex        =   36
         Text            =   "1000.0"
         Top             =   1200
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
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Text            =   "0.0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "补偿系数"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   89
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "B 台面重量"
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame boardA 
      Caption         =   "台面A"
      Height          =   3285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
      Begin VB.TextBox txtNonlinear 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   88
         Text            =   "1.0000"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton factora 
         Caption         =   "系数调整"
         Height          =   375
         Left            =   360
         TabIndex        =   59
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
         Height          =   315
         HideSelection   =   0   'False
         Left            =   1200
         TabIndex        =   34
         Text            =   "1000.0"
         Top             =   1200
         Width           =   975
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
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "补偿系数"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   87
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "A 台面重量"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "比例系数:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   855
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9000
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   57600
   End
   Begin VB.Frame Frame4 
      Caption         =   "功能操作"
      Height          =   1095
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton zeroBtn 
         Caption         =   "台面回零"
         Height          =   615
         Left            =   1680
         TabIndex        =   86
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton loginBtn 
         Caption         =   "用户登录"
         Height          =   615
         Left            =   360
         TabIndex        =   72
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton portSetting 
         Caption         =   "端口设置"
         Height          =   615
         Left            =   6120
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "退出"
         Height          =   615
         Left            =   12360
         TabIndex        =   70
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Index           =   1
         Left            =   7440
         TabIndex        =   69
         Top             =   120
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Index           =   0
         Left            =   12120
         TabIndex        =   68
         Top             =   120
         Width           =   15
      End
      Begin VB.CommandButton printBtn 
         Caption         =   "打印计量单"
         Height          =   615
         Left            =   10680
         TabIndex        =   67
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton contrastBtn 
         Caption         =   "数据对比"
         Height          =   615
         Left            =   7680
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton queryBtn 
         Caption         =   "数据查询"
         Height          =   615
         Left            =   9240
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmStaticWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conBanlancePos = 3

Dim objFileSystem As Object
Dim ReadAveTimeFile As String      '读取平均次数
Dim strComBuf As String
Dim iAvrTimes As Long
Dim iRawBuf(1000, 16) As Long
Dim iFrameCnt As Long
Dim iSensor(16) As Long
Dim fSensor(16) As Double
Dim iBoard(4) As Long
Dim fBoard(4) As Double
Dim balance(16) As Double
Dim fScaleA, fScaleB, fScaleC As Double
Dim iRef(16) As Integer
Dim fRef(16) As Double
Dim refWeight(4) As Double
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
Dim strtime As String

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

Private m_nonlinear(3) As String
Private m_smooth(3, 5) As Double

Dim startFlag As Boolean
Dim printFlag As Boolean
Dim refFlag As Boolean
Dim smoothCnt As Integer

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
    
    MSFlexGrid1.Clear
    With MSFlexGrid1
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 1) = "车号"
        .TextMatrix(0, 2) = "车型"
        .TextMatrix(0, 3) = "重量"
        .TextMatrix(0, 4) = "速度"
    End With
    MSFlexGrid1.rows = 2
    
    msfrow = 0
End Sub
Public Sub AddLine(Data() As Long)
    Dim i As Integer, j As Integer, K As Integer
    
            For i = 0 To 3
               avgsum(1) = avgsum(1) + Data(i)
            Next i
            For j = 4 To 7
               avgsum(2) = avgsum(2) + Data(j)
            Next j
            For K = 8 To 11
               avgsum(3) = avgsum(3) + Data(K)
            Next K
            
            For i = 0 To 11
                Text1(i).Text = CStr(Data(i))
            Next i
            
    Call FrameExp(Data())

End Sub
Public Sub AddTrainCode(sCode As String)
    Call TrainCodeExp(sCode)

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
    refWeight(0) = fBoard(0)
End Sub


Private Sub Btozero_Click()
    Dim i As Integer
    
    For i = 4 To 7
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i
    refWeight(1) = fBoard(1)
End Sub

Private Sub Command3_Click()
    Call SaveFactorFile
End Sub

Private Sub Command4_Click()
    Unload Me
    
End Sub

Private Sub Command5_Click()
    
End Sub

Private Sub contrastBtn_Click()
    frmConstrat.Show vbModal
End Sub

Private Sub Ctozero_Click()
Dim i As Integer
    For i = 8 To 11
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i
    
    refWeight(2) = fBoard(2)
End Sub

Private Sub factora_Click()
    Dim i As Integer
    
    g_CurrentConfigPlantform = 0
    factorDlg.Show vbModal
    For i = 0 To conMaxNonlinerFactors - 1
        m_factora(i) = g_NonlineFactors(conFactorAStartPos + i)
    Next i
End Sub

Private Sub factorb_Click()
    Dim i As Integer
    
    g_CurrentConfigPlantform = 1
    factorDlg.Show vbModal
    For i = 0 To conMaxNonlinerFactors - 1
        m_factorb(i) = g_NonlineFactors(conFactorBStartPos + i)
    Next i
End Sub

Private Sub factorc_Click()
    Dim i As Integer
    
    g_CurrentConfigPlantform = 2
    factorDlg.Show vbModal
    For i = 0 To conMaxNonlinerFactors - 1
        m_factorc(i) = g_NonlineFactors(conFactorCStartPos + i)
    Next i
End Sub

Private Sub Form_Initialize()
    Dim i As Integer
    
    startFlag = False
    printFlag = False
    g_StartLogin = False
    refFlag = True
   
    For i = 0 To 3
        refWeight(i) = 0
        m_nonlinear(i) = Format(1, "#0.0000")
    Next i

    msfrow = 0
    smoothCnt = 0
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
    Dim fileName As String
    Dim pos As Integer
    
        
    On Error GoTo ok:
    
    GdhSamp1.Run = True
    GdhCode1.Run = True
    
    
    timer = 1
    CountRecord = 1
    Frame2(1).Visible = True
    startFlag = False
'    cmdStaticWeight.Visible = False

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
        board(i).Text = Format(0, "###0.000") & "吨"
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
    MSFlexGrid1.rows = 2
    
    startFlag = True
    
ok:
    Close #33

    
    If Option1.Value = True Then
        Option2.Value = False
        strDirection = "L"
        strDirection1 = "<--"
    Else
     
        Option2.Value = True
        strDirection = "R"
        strDirection = "-->"
    End If
    
    If g_LoginUser <> "Administrator" And Not g_SuperOk Then
        factora.Enabled = False
        factorb.Enabled = False
        factorc.Enabled = False
        Command3.Enabled = False
        portSetting.Enabled = False
        
        txtScaleA.Enabled = False
        txtScaleB.Enabled = False
        txtScaleC.Enabled = False
        
        For i = 0 To 11
            bal(i).Enabled = False
        Next i
        For i = 0 To 2
            txtNonlinear(i).Enabled = False
        Next i
    Else
        factora.Enabled = True
        factorb.Enabled = True
        factorc.Enabled = True
        Command3.Enabled = True
        portSetting.Enabled = True
        
        txtScaleA.Enabled = True
        txtScaleB.Enabled = True
        txtScaleC.Enabled = True
        
        For i = 0 To 11
            bal(i).Enabled = True
            'bal(i).ForeColor = QColor(2 * (i - 4))
        Next i
        For i = 0 To 2
            txtNonlinear(i).Enabled = True
        Next i
    End If
    
    If conFactory = Factory.gsh Then
        Command1.Enabled = False
    End If
    
 gdh_Weight_Dbpath = App.Path & "\"
 Set objFileSystem = CreateObject("scripting.filesystemobject")
 Set ObjSystem = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub GdhCode1_OnCode(sCode As String)
    On Error GoTo ok
    AddTrainCode sCode
ok:
End Sub

Private Sub GdhSamp1_OnRcvLine(Data() As Long, ByVal cnt As Long)
    On Error GoTo ok
    AddLine Data
ok:
End Sub
Private Sub GdhSamp1_OnTimer()
On Error GoTo ok
    If GdhCode1.Run Then
        GdhCode1.Receive
    End If
ok:
End Sub

Public Sub FrameExp(Data() As Long)
    Dim i, j As Integer
    Dim tSensor(12) As Long           '保存iSensor(i) - iRef(i)的差值
    Dim sumA, sumB, sumC, sumAB, sumCB As Long

    For i = 0 To 11
        iRawBuf(iFrameCnt, i) = Data(i)
    Next i

    If iFrameCnt = iAvrTimes - 1 Then
    
       'calc boradA sensor data
       For i = 0 To 3
           sumA = 0
           For j = 0 To iAvrTimes - 1
               sumA = sumA + iRawBuf(j, i)
           Next j
 
            iSensor(i) = sumA / iAvrTimes
           
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
            
           fSensor(i) = ((iSensor(i) - iRef(i)) * balance(i)) / fScaleC
           tSensor(i) = iSensor(i) - iRef(i) '保存iSensor(i) - iRef(i)的差值
       Next i
       
       If refFlag Then
            For i = 0 To 15
                 iRef(i) = iSensor(i)
                 fSensor(i) = 0
            Next i
            
            For i = 0 To 4
                m_smooth(0, i) = fBoard(0)
                m_smooth(1, i) = fBoard(1)
                m_smooth(2, i) = fBoard(2)
            Next i
            refFlag = False
        End If
       
        'calc board data
        For i = 0 To 2
            iBoard(i) = iSensor(4 * i) + iSensor(4 * i + 1) + iSensor(4 * i + 2) + iSensor(4 * i + 3)
            fBoard(i) = fSensor(4 * i) + fSensor(4 * i + 1) + fSensor(4 * i + 2) + fSensor(4 * i + 3)

            fBoard(i) = AdjustData(CInt(i), fBoard(i))
            board(i).Text = Format(fBoard(i), "###0.000")
            txtNonlinear(i).Text = m_nonlinear(i)
        Next i
        
        'For i = 0 To 2
        '    smoothCnt = smoothCnt Mod 5
        '    m_smooth(i, smoothCnt) = fBoard(i)
            
        '    fBoard(i) = (m_smooth(i, 0) + m_smooth(i, 1) + m_smooth(i, 2) + m_smooth(i, 3) + m_smooth(i, 4)) / 5
        '    board(i).Text = Format(fBoard(i), "###0.000")
            
        '    smoothCnt = smoothCnt + 1
        'Next i
       
       
      'calc A + B,C + B 台面重量
       boardAB.Text = Format(Val(Mid(board(0).Text, 1, 6)) + Val(Mid(board(1).Text, 1, 6)), "###0.000  ")
       boardCB.Text = Format(Val(Mid(board(2).Text, 1, 6)) + Val(Mid(board(1).Text, 1, 6)), "###0.000  ")
        
       iFrameCnt = 0
    End If
    
    iFrameCnt = iFrameCnt + 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmWeight.RunDynWeight = True
    'Call SaveFactorFile
    'Call commClose
    GdhSamp1.Run = False
    GdhCode1.Run = False
    
End Sub

Private Sub loginBtn_Click()
    Dim i As Integer
    
    frmEnter.Show 1, Me
    If frmEnter.IsLogin = True Then
        g_LoginUser = frmEnter.LoginName

        If g_LoginUser = "Administrator" Then
            g_SuperOk = True
            
            factora.Enabled = True
            factorb.Enabled = True
            factorc.Enabled = True
            Command3.Enabled = True
            portSetting.Enabled = True
            
            txtScaleA.Enabled = True
            txtScaleB.Enabled = True
            txtScaleC.Enabled = True
            
            For i = 0 To 11
                bal(i).Enabled = True
            Next i
            For i = 0 To 2
                txtNonlinear(i).Enabled = True
            Next i
        Else
            g_SuperOk = False
            
            factora.Enabled = False
            factorb.Enabled = False
            factorc.Enabled = False
            Command3.Enabled = False
            portSetting.Enabled = False
            
            txtScaleA.Enabled = False
            txtScaleB.Enabled = False
            txtScaleC.Enabled = False
            
            For i = 0 To 11
                bal(i).Enabled = False
            Next i
            For i = 0 To 2
            txtNonlinear(i).Enabled = False
        Next i
        End If
    End If
End Sub

Private Sub MSFlexGrid1_Click()
    Dim strin As String
    
    If MSFlexGrid1.Row = MSFlexGrid1.rows - 1 Then
        Exit Sub
    End If
    
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

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    Combo1.Clear
    
    If MSFlexGrid1.Row = MSFlexGrid1.rows - 1 Then
        Exit Sub
        
    End If
     
    Combo1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top '移动组合框到网格当前的地方
    Combo1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    
    grow = MSFlexGrid1.Row '保存网格行和列的位置
    If (MSFlexGrid1.Col = 1 Or MSFlexGrid1.Col = 2 Or MSFlexGrid1.Col = 4) Then
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

Private Sub portSetting_Click()
    GdhSamp1.Run = False
    GdhCode1.Run = False
    settingDialog.Show vbModal
    
    If settingDialog.RetStatus Then
        MsgBox "请关闭程序，重新启动！", vbOKOnly, "提示"
    End If
End Sub

Private Sub printBtn_Click()
    If printFlag Then
        Call PrintReport
        printFlag = False
    Else
        MsgBox "请先保存数据，然后打印！", vbExclamation, "提示"
    End If
End Sub

Private Sub queryBtn_Click()
    g_QueryMethod = QueryMethod.Other
    
    frmQuery.Show vbModal
    
End Sub

'Private Sub Timer2_Timer()
'Frame2(0).Visible = False
'Frame2(1).Visible = False
'Timer2.Enabled = False

'    If Not Run Then Exit Sub
'    RaiseEvent OnTimer
'    Receive
'End Sub

Private Sub txtScaleA_Change()
    If Not startFlag Then GoTo ok
     
    If Val(txtScaleA.Text) < 0 Or Val(txtScaleA.Text) > 10000 Then
        MsgBox ("比例系数应设置在0至10000间!")
    Else
        fScaleA = Val(txtScaleA.Text)
        g_NonlineFactors(conGradStartPos) = txtScaleA.Text
        g_GridA = txtScaleA.Text
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
        g_GridB = txtScaleB.Text
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
        g_GridC = txtScaleC.Text
        
    End If
ok:
End Sub
Public Function StaticWeight_Exp(Data As String)

     msfrow = msfrow + 1
     MSFlexGrid1.TextMatrix(msfrow, 0) = msfrow
     MSFlexGrid1.TextMatrix(msfrow, 3) = Data
     MSFlexGrid1.TextMatrix(msfrow, 4) = str(0)
     
     MSFlexGrid1.rows = MSFlexGrid1.rows + 1
    
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
Function SaveDataToFile(FilePath As String, grid As MSFlexGrid) ', vDateTime As String, vDirection As String)
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

    If grid.rows = 2 Or grid.TextMatrix(1, 0) = "" Then
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
    Print #22, Trim(str(grid.rows - 2))
    '存表头
    strLine = "序号" + "|" + "车号" + "|" + "车型" + "|" + "毛重" + "|" + "速度" + "|"
    '存数据
    Print #22, strLine
    For i = 1 To grid.rows - 2
        strLine = ""
        strLine = strLine + Trim(grid.TextMatrix(i, 0)) + "|"
        strLine = strLine + Trim(grid.TextMatrix(i, 1)) + "|"
        strLine = strLine + Trim(grid.TextMatrix(i, 2)) + "|"
        strLine = strLine + Trim(grid.TextMatrix(i, 3)) + "|"
        strLine = strLine + Trim(grid.TextMatrix(i, 4)) + "|"
        Print #22, strLine
    Next i
    Close #22
'    Sleep (10000)
ok:
End Function
Private Function GetSaveFileName() As String
    Dim strTemp As String
    
    strTemp = Format(Now, "yyyymmddhhmm")
    strTemp = m_SaveFilePath & "\" & strTemp & ".txt"
    
    GetSaveFileName = strTemp
End Function
Private Sub SaveData()
    Dim strTime1 As String
    Dim strTime2 As String

      Dim intRow As Integer
    Dim i, j, length As Integer
    Dim strser As String, str As String, strBC As String
    Dim buffer(1 To 200) As String
    Dim strChoiceFile As String
    Dim str_FilePath As String
    Dim Sign As Boolean
    Dim strTemp As String
    On Error GoTo ok
    
    
   ' strtime = Mid(GdhCarriage1.TrainOnTime, 1, 2) & Mid(GdhCarriage1.TrainOnTime, 4, 2) & Mid(GdhCarriage1.TrainOnTime, 7, 2)
   ' strTime1 = Format(Now, "yyyymmdd") & strtime
    strTemp = Format(Date, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss")
    strtime = strTemp

'    Call CreatDB(Mid(Trim(strtemp), 1, 4))
    Sign = Save_Data_to_gdhys(Trim(strTemp), Trim(strDirection1), MSFlexGrid1, gdh_Weight_Dbpath)
            
    If Sign Then
        printFlag = True
        MsgBox "过衡数据已经保存!", vbExclamation, "提示"
    Else
        MsgBox "无可保存的数据!", vbExclamation, "提示"
    End If
    
   ' If Sign = True Then
   '     Frame2(0).Visible = True
   '     Timer2.Enabled = True
   ' Else
   '     Frame2(1).Visible = True
   '     Timer2.Enabled = True
   ' End If
     
ok:
   ' Call GdhSaveWeight1.SaveOriginalData(strTime1, strDirection, GdhCarriage1.Grid)
End Sub
Function Save_Data_to_gdhys(strDate_Time As String, v_Direction As String, GD As MSFlexGrid, dbsavePath As String) As Boolean
    Dim db As New ADODB.Connection, rs As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim query As String, table_Name As String, dbName As String
    Dim fullDBPath As String
    Dim K As Integer
    Dim total As Double
    
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
    fullDBPath = App.Path & "\" & dbName
    
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & fullDBPath & ";Jet OLEDB:Database Password=dfrw2306;"
    
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
    
    total = CDbl(0)
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
        
        total = total + CDbl(Trim(GD.TextMatrix(i, 3)))
    Next i
    rs.Close
    
    '添加索引记录
    query = "select * from gdhindex where 日期时间='" & strDate_Time & "'"
'    Set rs = db.OpenRecordset(Query)
    rs.Open query, db, adOpenDynamic, adLockOptimistic
    
    If Not rs.BOF And Not rs.EOF Then
    Else
        rs.AddNew
        rs.Fields("表名") = "gdhys"
        rs.Fields("车数") = Trim(str(K - 1))
        rs.Fields("日期时间") = strDate_Time
        rs.Fields("方向") = v_Direction
        rs.Fields("总重") = Format(total, "###0.000")
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
    Dim i As Integer

    If (Abs(Value) > 10) Then
    
        Index = CInt((Abs(Value) - 10) / 2)
        If Index > 15 Then
            Index = 15
        End If
        If Index < 0 Then
            Index = 0
        End If
        
        If pt = 0 Then
            rev = CDbl(m_factora(Index) / 10000)
            m_nonlinear(0) = Format(rev, "#0.0000")
        ElseIf pt = 1 Then
            rev = CDbl(m_factorb(Index) / 10000)
            m_nonlinear(1) = Format(rev, "#0.0000")
        Else
            rev = CDbl(m_factorc(Index) / 10000)
            m_nonlinear(2) = Format(rev, "#0.0000")
        End If
        
        ret = Value * rev
    Else
        ret = Value
        For i = 0 To 2
            m_nonlinear(i) = Format(1, "#0.0000")
        Next i
    End If
    
    AdjustData = ret
    Exit Function
End Function
Private Sub PrintReport()
    Call GdhPrintWeight1.PrintOriginalData(strtime, strDirection, MSFlexGrid1)
End Sub

Private Sub zeroBtn_Click()
    Dim i As Integer
    
    For i = 0 To 15
        iRef(i) = iSensor(i)
        fRef(i) = fSensor(i)
    Next i
    
    For i = 0 To 3
        refWeight(i) = fBoard(i)
    Next i
    
End Sub
