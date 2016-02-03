VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSectionSamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'���ݲɼ�������
Option Explicit
'=========================================================================================================
'����
'=========================================================================================================
Const conSection = "Samp"

'=========================================================================================================
'�ֲ�����
'=========================================================================================================

'=========================================================================================================
'����
'=========================================================================================================
Private m_Ini As CGdhINI    'ָ�����������ini�ļ�

Private m_Port As Integer       '�˿�
Private m_Settings As String   '�˿�����

'=========================================================================================================
'����
'=========================================================================================================
'���ļ��ж�����key
Public Sub GetAllKeys()
    m_Port = m_Ini.GetLong(conSection, "Port", 1)
    m_Settings = m_Ini.GetString(conSection, "Settings", "115200,n,8,1")
End Sub

'=========================================================================================================
'����
'=========================================================================================================
'ini
Public Property Set Ini(vData As CGdhINI)
    Set m_Ini = vData
End Property

'Port
Public Property Let Port(vData As Integer)
    m_Port = vData
    m_Ini.SaveInteger conSection, "Port", m_Port
End Property

Public Property Get Port() As Integer
    Port = m_Port
End Property

'Settings
Public Property Let Settings(vData As String)
    m_Settings = vData
    m_Ini.SaveString conSection, "Settings", m_Settings
End Property

Public Property Get Settings() As String
    Settings = m_Settings
End Property
