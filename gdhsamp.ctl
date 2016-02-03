VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl GdhSamp 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   InvisibleAtRuntime=   -1  'True
   Picture         =   "gdhsamp.ctx":0000
   ScaleHeight     =   840
   ScaleWidth      =   1035
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
      BaudRate        =   115200
      InputMode       =   1
   End
End
Attribute VB_Name = "GdhSamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================��
Const conSampChannelNum = 12    '2007-01-23


'=================================================================================��
'�����¼�
'==================================================================================
Public Event OnRcvLine(Data() As Long, ByVal cnt As Long)     '���յ���������
Public Event OnTimer()

Private m_DataCount As Long         '���к��������
'==================================================================================
'ȱʡ����ֵ:

'���Ա���:




'================���ر���=========================================
'������һ������
Dim m_LineData(0 To conSampChannelNum - 1) As Long

'���մ����û�����
'2007-01-23
Dim m_byteBuf(300000) As Byte
Dim m_Count As Long

'���մ���
Public Sub Receive()
    
    Const conLineSize = conSampChannelNum * 2 + 2
    
    Dim varRcv As Variant
    Dim I As Long, pos As Long, flagPos As Long
    Dim num As Long
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    
    'On Error GoTo rcvErrHandler
    
    If Not MSComm1.PortOpen Then Exit Sub
    If MSComm1.InBufferCount = 0 Then Exit Sub
    
    'ȡ����,���뵽�ֽ�����
    varRcv = MSComm1.Input
    For I = LBound(varRcv) To UBound(varRcv)
        m_byteBuf(m_Count + I) = varRcv(I)
    Next I
    m_Count = m_Count + UBound(varRcv) - LBound(varRcv) + 1
    
    If m_Count < conLineSize + 2 Then Exit Sub
    Do While Not bExit
        '����ff ff
        bFind = False
        For I = pos To m_Count - (conLineSize + 2)
            If (m_byteBuf(I) = &HFF And m_byteBuf(I + 1) = &HFF) Then
                flagPos = I
                bFind = True
                Exit For
            End If
        Next I
        
        If bFind Then
            If m_byteBuf(flagPos + conLineSize) = &HFF And m_byteBuf(flagPos + conLineSize + 1) = &HFF Then
                '�ж�����ȷ��ffff��ת��һ��
                For I = 0 To conSampChannelNum - 1
                    num = m_byteBuf(flagPos + 2 * I + 2)
                    num = num * 256 + m_byteBuf(flagPos + 2 * I + 3)
                    If num > 32767 Then
                        num = num - 65536
                    End If
                    m_LineData(I) = num
                Next I
                pos = flagPos + conLineSize
                '
                m_DataCount = m_DataCount + 1
                RaiseEvent OnRcvLine(m_LineData, m_DataCount)
            Else
                pos = flagPos + 1
            End If
        Else
            pos = m_Count - (conLineSize + 2)   '2007-01-11
            bExit = True
        End If
    Loop
    
    '�ƶ�buf
    m_Count = m_Count - pos
    If pos > 0 Then
        For I = 0 To m_Count - 1
            m_byteBuf(I) = m_byteBuf(pos + I)
        Next I
    End If
    Exit Sub
    
'������
'rcvErrHandler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "���ݽ���-Receive"
    
End Sub



'==============================����====================================
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "����/����ͨѶ�˿ںš�"
    CommPort = MSComm1.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    MSComm1.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Public Property Get CommSettings() As String
Attribute CommSettings.VB_Description = "����/���ز����ʡ���żУ�顢����λ��ֹͣλ������"
    CommSettings = MSComm1.Settings
End Property

Public Property Let CommSettings(ByVal New_CommSettings As String)
    MSComm1.Settings() = New_CommSettings
    PropertyChanged "CommSettings"
End Property


Private Sub MSComm1_OnComm()
    '���ͨ��״̬
'    Select Case MSComm1.CommEvent
        ' ����
'        Case comOmronEventBreak, comOmronEventCDTO, comOmronEventCTSTO, comOmronEventDSRTO, comOmronEventFrame, _
'             comOmronEventOverrun, comOmronEventRxOver, comOmronEventRxParity, comOmronEventTxFull, comOmronEventDCB
'
'            mOmronEvent = omronEvCommErr
'            RaiseEvent OnOmron
'    End Select
    Debug.Print "Mscomm event: " & MSComm1.CommEvent
End Sub

Private Sub Timer1_Timer()
    If Not Run Then Exit Sub
    RaiseEvent OnTimer
    Receive
 End Sub

Private Sub Class_Terminate()
End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MSComm1.CommPort = PropBag.ReadProperty("CommPort", 1)
    MSComm1.Settings = PropBag.ReadProperty("CommSettings", "115200,n,8,1")
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CommPort", MSComm1.CommPort, 1)
    Call PropBag.WriteProperty("CommSettings", MSComm1.Settings, "115200,n,8,1")
End Sub

Public Property Get DataCount() As Long
    DataCount = m_DataCount
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
        MSComm1.CommPort = gGdhIni.Samp.Port
        MSComm1.Settings = gGdhIni.Samp.Settings
    End If
    MSComm1.PortOpen = New_Run
    
    If MSComm1.PortOpen Then
        m_Count = 0
        m_DataCount = 0
        MSComm1.InBufferCount = 0
    End If
    
    Timer1.Enabled = MSComm1.PortOpen
    
    Exit Property
RunErr:
    MsgBox "�򿪲����˿�ʱ���ִ���" & vbCrLf & "�˿ںţ�" & MSComm1.CommPort & "��" & Err.Description
    Err.clear
End Property


