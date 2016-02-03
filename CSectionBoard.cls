VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSectionBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'数据预处理配置项
Option Explicit
'=========================================================================================================
'常量
'=========================================================================================================
Const conSection = "Board"

Const conChannelNum = 12    '2007-01-09
Const conNonlinearNum = 15  '2007-01-09
'=========================================================================================================
'局部声明
'=========================================================================================================

'=========================================================================================================
'变量
'=========================================================================================================
Private m_Ini As CGdhINI    '指向包含自身的ini文件


Private m_BoardLen As Single
Private m_Weight As Single
Private m_DirL(0 To conBordNum - 1) As Single
Private m_DirR(0 To conBordNum - 1) As Single
Private m_Percent As Single
Private m_Channel(0 To conChannelNum - 1) As Single
Private m_Nonlinear(1 To conNonlinearNum) As Single
Private m_Nonlinearb(1 To conNonlinearNum) As Single

'=========================================================================================================
'方法
'=========================================================================================================
'从文件中读所有key
Public Sub GetAllKeys()
    Dim I As Integer
    Dim key As String
    
    m_BoardLen = m_Ini.GetSingle(conSection, "BoardLen", 3.5)
    '2007-01-11
    For I = 0 To conBordNum - 1
        key = "DirL" & I
        m_DirL(I) = m_Ini.GetSingle(conSection, key, 1)
        key = "DirR" & I
        m_DirR(I) = m_Ini.GetSingle(conSection, key, 1)
    Next I

    m_Weight = m_Ini.GetSingle(conSection, "Weight", 1000)
    m_Percent = m_Ini.GetSingle(conSection, "Percent", 1)
    For I = 0 To conChannelNum - 1
        m_Channel(I) = m_Ini.GetSingle(conSection, "Channel" & I, 1)
    Next I
    For I = 1 To conNonlinearNum
        m_Nonlinear(I) = m_Ini.GetSingle(conSection, "Nonlinear" & I, 1)
        m_Nonlinearb(I) = m_Ini.GetSingle(conSection, "Nonlinearb" & I, 1)
    Next I
End Sub

'=========================================================================================================
'属性
'=========================================================================================================
'ini
Public Property Set Ini(vData As CGdhINI)
    Set m_Ini = vData
End Property


'BoardLen
Public Property Let BoardLen(vData As Single)
    m_BoardLen = vData
    m_Ini.SaveSingle conSection, "BoardLen", m_BoardLen
End Property

Public Property Get BoardLen() As Single
    BoardLen = m_BoardLen
End Property

'Weight
Public Property Let Weight(vData As Single)
    m_Weight = vData
    m_Ini.SaveSingle conSection, "Weight", m_Weight
End Property

Public Property Get Weight() As Single
    Weight = m_Weight
End Property

'DirL
Public Property Let DirL(vNo As Integer, vData As Single)
    m_DirL(vNo) = vData
    m_Ini.SaveSingle conSection, "DirL" & vNo, m_DirL(vNo)
End Property

Public Property Get DirL(vNo As Integer) As Single
    DirL = m_DirL(vNo)
End Property

'DirR
Public Property Let DirR(vNo As Integer, vData As Single)
    m_DirR(vNo) = vData
    m_Ini.SaveSingle conSection, "DirR" & vNo, m_DirR(vNo)
End Property

Public Property Get DirR(vNo As Integer) As Single
    DirR = m_DirR(vNo)
End Property

'Percent
Public Property Let Percent(vData As Single)
    m_Percent = vData
    m_Ini.SaveSingle conSection, "Percent", m_Percent
End Property

Public Property Get Percent() As Single
    Percent = m_Percent
End Property

'Channel
Public Property Let Channel(vIndex As Integer, vData As Single)
    m_Channel(vIndex) = vData
    m_Ini.SaveSingle conSection, "Channel" & vIndex, m_Channel(vIndex)
End Property

Public Property Get Channel(vIndex As Integer) As Single
    Channel = m_Channel(vIndex)
End Property

'Nonlinear
Public Property Let Nonlinear(vIndex As Integer, vData As Single)
    m_Nonlinear(vIndex) = vData
    m_Ini.SaveSingle conSection, "Nonlinear" & vIndex, m_Nonlinear(vIndex)
End Property

Public Property Get Nonlinear(vIndex As Integer) As Single
    Nonlinear = m_Nonlinear(vIndex)
End Property

Public Property Let Nonlinearb(vIndex As Integer, vData As Single)
    m_Nonlinearb(vIndex) = vData
    m_Ini.SaveSingle conSection, "Nonlinearb" & vIndex, m_Nonlinearb(vIndex)
End Property

Public Property Get Nonlinearb(vIndex As Integer) As Single
    Nonlinearb = m_Nonlinearb(vIndex)
End Property



