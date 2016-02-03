VERSION 5.00
Begin VB.Form settingDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "端口设置"
   ClientHeight    =   3120
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "车号设置"
      Height          =   1935
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox tnPortEn 
         Caption         =   "端口使能"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox tnAttr 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox portCbx 
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "端口号"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "属性"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "通道设置"
      Height          =   1935
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox sampPortEn 
         Caption         =   "端口使能"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox sampAttr 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox portCbx 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label setting 
         Caption         =   "属性"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "端口号"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "保存"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "settingDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_sampPortNum As String
Private m_sampAttribute As String
Private m_tnPortNum As String
Private m_tnAttribute As String
Private portNum As Variant

Private m_Config As CConfig
Private m_selPort(2) As String

Private m_retFlag As Boolean
Public Property Get RetStatus() As Boolean
    RetStatus = m_retFlag
End Property
Private Sub CancelButton_Click()
    m_retFlag = False
    Unload Me
End Sub

Private Sub Form_Initialize()
    m_sampPortNum = gGdhIni.Samp.port
    m_sampAttribute = gGdhIni.Samp.Settings
    m_tnPortNum = gGdhIni.Code.port
    m_tnAttribute = gGdhIni.Code.Settings
    portNum = Array("COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8")
    
    Set m_Config = New CConfig
    
    m_retFlag = False
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    '初始化
    m_Config.FileName = App.Path + "\gdh.bin"
    
    sampAttr.Text = m_sampAttribute
    tnAttr.Text = m_tnAttribute
    
    For i = 0 To 1
        portCbx(i).AddItem ("COM1")
        portCbx(i).AddItem ("COM2")
        portCbx(i).AddItem ("COM3")
        portCbx(i).AddItem ("COM4")
        portCbx(i).AddItem ("COM5")
        portCbx(i).AddItem ("COM6")
        portCbx(i).AddItem ("COM7")
        portCbx(i).AddItem ("COM8")
      
        sampPortEn.Value = vbChecked
        sampPortEn.Enabled = False
        tnPortEn.Value = vbChecked
        tnPortEn.Enabled = False
   Next i
   
   portCbx(0).Text = portNum(CInt(Trim(m_sampPortNum)) - 1)
   portCbx(1).Text = portNum(CInt(Trim(m_tnPortNum)) - 1)
   
   m_selPort(0) = Trim(m_sampPortNum)
   m_selPort(1) = Trim(m_tnPortNum)

End Sub

Private Sub OKButton_Click()
    m_Config.SaveString "samp", "port", m_selPort(0)
    m_Config.SaveString "samp", "settings", sampAttr.Text
    m_Config.SaveString "code", "port", m_selPort(1)
    m_Config.SaveString "code", "settings", tnAttr.Text
    
    m_retFlag = True
    Unload Me
End Sub

Private Sub portCbx_Click(Index As Integer)
    Dim port As String
    
    port = portCbx(Index).Text
    m_selPort(Index) = CStr(portCbx(Index).ListIndex + 1)
    
End Sub
