VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDebugCarriage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'本地变量
Private m_Row As Integer
Private m_Col As Integer
Private m_Code(1 To 5) As String
Private m_Weight(1 To 20, 1 To 5) As Single
Private m_MaxWeight(1 To 5) As Single
Private m_MinWeight(1 To 5) As Single
Private m_AvgWeight(1 To 5) As Single
Private m_Count As Integer


'属性
'行
Public Property Let Row(vData As Integer)
    m_Row = vData
End Property

Public Property Get Row() As Integer
    Row = m_Row
End Property

'列
Public Property Let Col(vData As Integer)
    m_Col = vData
End Property

Public Property Get Col() As Integer
    Col = m_Col
End Property

'车号
Public Property Let Code(vIndex As Integer, vData As String)
    m_Code(vIndex) = vData
End Property

Public Property Get Code(vIndex As Integer) As String
    Code = m_Code(vIndex)
End Property

'重量
Public Property Let Weight(vRow As Integer, vCol As Integer, vData As Single)
    m_Weight(vRow, vCol) = vData
End Property

Public Property Get Weight(vRow As Integer, vCol As Integer) As Single
    Weight = m_Weight(vRow, vCol)
End Property

'最大值
Public Property Let MaxWeight(vIndex As Integer, vData As Single)
    m_MaxWeight(vIndex) = vData
End Property

Public Property Get MaxWeight(vIndex As Integer) As Single
    MaxWeight = m_MaxWeight(vIndex)
End Property

'最小值
Public Property Let MinWeight(vIndex As Integer, vData As Single)
    m_MinWeight(vIndex) = vData
End Property

Public Property Get MinWeight(vIndex As Integer) As Single
    MinWeight = m_MinWeight(vIndex)
End Property

'平均值
Public Property Let AvgWeight(vIndex As Integer, vData As Single)
    m_AvgWeight(vIndex) = vData
End Property

Public Property Get AvgWeight(vIndex As Integer) As Single
    AvgWeight = m_AvgWeight(vIndex)
End Property

'数量
Public Property Let Count(vData As Integer)
    m_Count = vData
End Property

Public Property Get Count() As Integer
    Count = m_Count
End Property


