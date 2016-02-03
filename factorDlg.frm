VERSION 5.00
Begin VB.Form factorDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分段系数调整"
   ClientHeight    =   5355
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Exit 
      Caption         =   "退出"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "标准曲线"
      Height          =   4575
      Left            =   4560
      TabIndex        =   35
      Top             =   120
      Width           =   2295
      Begin VB.TextBox intercept 
         Height          =   375
         Left            =   960
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox grad 
         Height          =   375
         Left            =   960
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "截距"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "斜率"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "分段线性系数"
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   15
         Left            =   2520
         TabIndex        =   34
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   14
         Left            =   2520
         TabIndex        =   25
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   13
         Left            =   2520
         TabIndex        =   24
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   12
         Left            =   2520
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   11
         Left            =   2520
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   10
         Left            =   2520
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   9
         Left            =   2520
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   3
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label fa 
         Caption         =   "14-"
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   33
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "16-"
         Height          =   255
         Index           =   14
         Left            =   2160
         TabIndex        =   32
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "15-"
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   31
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "13-"
         Height          =   255
         Index           =   12
         Left            =   2160
         TabIndex        =   30
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "12-"
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   29
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "11-"
         Height          =   255
         Index           =   10
         Left            =   2160
         TabIndex        =   28
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "10-"
         Height          =   255
         Index           =   9
         Left            =   2160
         TabIndex        =   27
         Top             =   960
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "9-"
         Height          =   255
         Index           =   8
         Left            =   2160
         TabIndex        =   26
         Top             =   480
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "3-"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "1-"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "2-"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "4-"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "5-"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "6-"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "7-"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   11
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label fa 
         Caption         =   "8-"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   10
         Top             =   3840
         Width           =   255
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "恢复"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "保存"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "factorDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'==================================================================================
'私有变量定义
'==================================================================================
Private m_factor(conMaxNonlinerFactors - 1) As Long
Private m_factorDefault(conMaxNonlinerFactors - 1) As Long
Private m_factorFileName As String
Private m_currIndex As Integer
Private m_currVal As String

Private Function OpenFactorFile(fileName As String) As Boolean

End Function

Private Sub SaveData()
    Dim i As Integer
    Dim pos As Integer
    
    If g_CurrentConfigPlantform = 0 Then
        pos = conFactorAStartPos
    ElseIf g_CurrentConfigPlantform = 1 Then
        pos = conFactorBStartPos
    Else
        pos = conFactorCStartPos
    End If
    
    For i = 0 To conMaxNonlinerFactors - 1
        g_NonlineFactors(pos + i) = m_factor(i)
    Next i
    
    g_NonlineFactors(g_CurrentConfigPlantform + conGradStartPos) = grad.Text
    
End Sub

Private Sub CancelButton_Click()
    Text1(m_currIndex).Text = m_factorDefault(m_currIndex)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim pos As Integer
        
    If g_CurrentConfigPlantform = 0 Then
        pos = conFactorAStartPos
    ElseIf g_CurrentConfigPlantform = 1 Then
        pos = conFactorBStartPos
    Else
        pos = conFactorCStartPos
    End If

    For i = 0 To conMaxNonlinerFactors - 1
        m_factorDefault(i) = g_NonlineFactors(pos + i)
        m_factor(i) = g_NonlineFactors(pos + i)
    Next i
    
    For i = 0 To conMaxNonlinerFactors - 1
        Text1(i).Text = m_factor(i)
    Next i
    
    grad.Text = g_NonlineFactors(conGradStartPos + g_CurrentConfigPlantform)
    intercept.Text = 0
    
End Sub

Private Sub OKButton_Click()
    Call SaveData
    Unload Me
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Text1_Change(Index As Integer)
    m_currIndex = Index
    m_factor(Index) = Text1(Index).Text
    
End Sub
