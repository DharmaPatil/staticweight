VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl GdhCode 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GdhCode.ctx":0000
   ScaleHeight     =   750
   ScaleWidth      =   1080
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   120
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   20480
      RTSEnable       =   -1  'True
      BaudRate        =   57600
   End
End
Attribute VB_Name = "GdhCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================��


'=================================================================================��
'�����¼�
'==================================================================================
Public Event OnCode(sCode As String)

'=================================================================================��
'���Ա���
'==================================================================================
Private m_DataCount As Long         '
'==================================================================================
'ȱʡ����ֵ:

'���Ա���:

'=================================================================================��
'���ر���
'==================================================================================
'���մ����û�����
Dim mRcvBuf As String

'=================================================================================��
'����
'==================================================================================
'������Ž��ջ�����
Public Sub ClearBuff()
    If MSComm1.PortOpen Then
        MSComm1.InBufferCount = 0
        mRcvBuf = ""
    End If
End Sub
'���մ���
Public Sub Receive()

'���г���
    Const conLineSize = 26

    Dim posBegin  As Integer, posEnd As Integer
    Dim sLine As String, sCode As String

    If Not MSComm1.PortOpen Then Exit Sub
    If MSComm1.InBufferCount = 0 Then Exit Sub

    'ȡ����
    MSComm1.InputLen = conLineSize
    mRcvBuf = mRcvBuf + MSComm1.Input
    If Len(mRcvBuf) < conLineSize Then Exit Sub

    posBegin = InStr(1, mRcvBuf, "@")
    If posBegin > 0 Then
        posEnd = InStr(posBegin + 1, mRcvBuf, "&")
        If posEnd > 0 Then
            If posEnd - posBegin + 1 = conLineSize Then
                sCode = Mid(mRcvBuf, posBegin + 1, 14)
                RaiseEvent OnCode(sCode)
            End If
            mRcvBuf = Mid(mRcvBuf, posEnd + 1)
        Else
            mRcvBuf = Mid(mRcvBuf, posBegin)
        End If
    Else
        mRcvBuf = ""
    End If


'ʢ������


End Sub

Public Sub OpenPower()
    If Not MSComm1.PortOpen Then Exit Sub
    MSComm1.Output = "@on&"
     MSComm1.Output = "2"
End Sub

Public Sub ClosePower()
    If Not MSComm1.PortOpen Then Exit Sub
    MSComm1.Output = "@off&"
     MSComm1.Output = "3"
End Sub


'=================================================================================��
'����
'==================================================================================



Private Sub MSComm1_OnComm()
'2007-07-18
'    Select Case MSComm1.CommEvent
'        Case comEvReceive
'            Receive
'        ' ����
''        Case comOmronEventBreak, comOmronEventCDTO, comOmronEventCTSTO, comOmronEventDSRTO, comOmronEventFrame, _
''             comOmronEventOverrun, comOmronEventRxOver, comOmronEventRxParity, comOmronEventTxFull, comOmronEventDCB
''
''            mOmronEvent = omronEvCommErr
''            RaiseEvent OnOmron
'    End Select
'    'Debug.Print "Mscomm event: " & MSComm1.CommEvent
End Sub

'2007-07-16
'Private Sub Timer1_Timer()
'    Receive
'End Sub

Private Sub Class_Terminate()
End Sub

''==============================����====================================
'Public Property Get CommPort() As Integer
'    CommPort = MSComm1.CommPort
'End Property
'
'Public Property Let CommPort(ByVal New_CommPort As Integer)
'    MSComm1.CommPort() = New_CommPort
'    PropertyChanged "CommPort"
'End Property
'
'Public Property Get CommSettings() As String
'    CommSettings = MSComm1.Settings
'End Property
'
'Public Property Let CommSettings(ByVal New_CommSettings As String)
'    MSComm1.Settings() = New_CommSettings
'    PropertyChanged "CommSettings"
'End Property
'
''Ϊ�û��ؼ���ʼ������
'Private Sub UserControl_InitProperties()
'End Sub
'
''�Ӵ������м�������ֵ
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    MSComm1.CommPort = PropBag.ReadProperty("CommPort", 1)
'    MSComm1.Settings = PropBag.ReadProperty("CommSettings", "115200,n,8,1")
'End Sub
'
''������ֵд���洢��
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("CommPort", MSComm1.CommPort, 1)
'    Call PropBag.WriteProperty("CommSettings", MSComm1.Settings, "115200,n,8,1")
'End Sub

Public Property Let DataCount(ByVal vData As Long)
    m_DataCount = vData
End Property

Public Property Get Run() As Boolean
    Run = MSComm1.PortOpen
End Property

Public Property Let Run(ByVal New_Run As Boolean)
    On Error GoTo RunErr
        
    If MSComm1.PortOpen = New_Run Then Exit Property
    
    Timer1.Enabled = False  ' �ؼ�ʱ��
    '���ö˿ںͲ���
    If Not MSComm1.PortOpen Then
        MSComm1.CommPort = gGdhIni.Code.Port
        MSComm1.Settings = gGdhIni.Code.Settings
    End If
    MSComm1.PortOpen = New_Run
    
    If MSComm1.PortOpen Then
        MSComm1.InBufferCount = 0
        m_DataCount = 0
    End If
    
    mRcvBuf = ""
    Timer1.Enabled = MSComm1.PortOpen
    
    Exit Property
RunErr:
    MsgBox "�򿪳��Ž��ն˿�ʱ���ִ���" & vbCrLf & "�˿ںţ�" & MSComm1.CommPort & "��" & Err.Description, vbInformation, "��ʾ"
    Err.Clear
End Property
