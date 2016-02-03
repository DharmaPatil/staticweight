VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form configDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   6210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10200
   DrawMode        =   0  'Blackness
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab setting 
      Height          =   5175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9128
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "端口设置"
      TabPicture(0)   =   "configDialog.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "通道平衡系数"
      TabPicture(1)   =   "configDialog.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text2(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text2(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text2(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text2(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text2(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text2(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text2(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text2(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text2(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text2(11)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "非线性系数"
      TabPicture(2)   =   "configDialog.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   11
         Left            =   -68160
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   10
         Left            =   -68160
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   9
         Left            =   -69240
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   8
         Left            =   -69240
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   7
         Left            =   -70680
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   6
         Left            =   -70680
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   5
         Left            =   -71760
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   -73080
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   -73080
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   -74040
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1215
         Left            =   -74520
         TabIndex        =   9
         Top             =   600
         Width           =   7695
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   4
            Left            =   2760
            TabIndex        =   14
            Text            =   "Text2"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   10
            Text            =   "Text2"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "通道端口"
         Height          =   3015
         Left            =   -74160
         TabIndex        =   3
         Top             =   600
         Width           =   2895
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   375
            Left            =   480
            TabIndex        =   8
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Text            =   "Combo1"
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "属性"
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "端口"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "configDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub
