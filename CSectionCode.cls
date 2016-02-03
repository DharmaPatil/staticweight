VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSectionCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'车号配置项
Option Explicit
'=========================================================================================================
'常量
'=========================================================================================================
Const conSection = "Code"

'=========================================================================================================
'局部声明
'=========================================================================================================

'=========================================================================================================
'变量
'=========================================================================================================
Private m_Ini As CGdhINI    '指向包含自身的ini文件

Private m_Port As Integer       '端口
Private m_Settings As String   '端口设置
Private m_BoardNo As Integer    '车号安装位置

'=========================================================================================================
'方法
'=========================================================================================================
'从文件中读所有key
Public Sub GetAllKeys()
    m_Port = m_Ini.GetLong(conSection, "Port", 2)
    m_Settings = m_Ini.GetString(conSection, "Settings", "115200,n,8,1")
    m_BoardNo = m_Ini.GetInteger(conSection, "TNBoardNo", 2)
    If m_BoardNo < 0 Or m_BoardNo > 2 Then
        m_BoardNo = 2
    End If
End Sub

'=========================================================================================================
'属性
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

'BoardNo
Public Property Let BoardNo(vData As Integer)
    m_BoardNo = vData
    m_Ini.SaveInteger conSection, "TNBoardNo", m_BoardNo
End Property

Public Property Get BoardNo() As Integer
    BoardNo = m_BoardNo
End Property


