VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSectionProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'数据预处理配置项
Option Explicit
'=========================================================================================================
'常量
'=========================================================================================================
Const conSection = "Process"

'=========================================================================================================
'局部声明
'=========================================================================================================

'=========================================================================================================
'变量
'=========================================================================================================
Private m_Ini As CGdhINI    '指向包含自身的ini文件

Private m_TrainOnThreshold As Long      '上称阀值
Private m_TrainOffDelayTime As Long     '下称判定时间
Private m_DifThreshold As Long           '差分域值
Private m_DifThresholdStep              '差分域值步长

'=========================================================================================================
'方法
'=========================================================================================================
'从文件中读所有key
Public Sub GetAllKeys()
    m_TrainOnThreshold = m_Ini.GetLong(conSection, "TrainOnThreshold", 1280)
    m_TrainOffDelayTime = m_Ini.GetLong(conSection, "TrainOffDelayTime", 5)
    m_DifThreshold = m_Ini.GetLong(conSection, "DifThreshold", 500)
    m_DifThresholdStep = m_Ini.GetLong(conSection, "DifThresholdStep", 100)
End Sub

'=========================================================================================================
'属性
'=========================================================================================================
'ini
Public Property Set Ini(vData As CGdhINI)
    Set m_Ini = vData
End Property

'TrainOnThreshold
Public Property Let TrainOnThreshold(vData As Long)
    m_TrainOnThreshold = vData
    m_Ini.SaveLong conSection, "TrainOnThreshold", m_TrainOnThreshold
End Property

Public Property Get TrainOnThreshold() As Long
    TrainOnThreshold = m_TrainOnThreshold
End Property

'TrainOffDelayTime
Public Property Let TrainOffDelayTime(vData As Long)
    m_TrainOffDelayTime = vData
    m_Ini.SaveLong conSection, "TrainOffDelayTime", m_TrainOffDelayTime
End Property

Public Property Get TrainOffDelayTime() As Long
    TrainOffDelayTime = m_TrainOffDelayTime
End Property

Public Property Get DifThreshold() As Long
    DifThreshold = m_DifThreshold
End Property

Public Property Get DifThresholdStep() As Long
    DifThresholdStep = m_DifThresholdStep
End Property
