Attribute VB_Name = "GobalModule"
Option Explicit

Public Enum QueryMethod
    Constrat = 0
    Other = 1
End Enum

Public Enum ReturnStatus
    StatusOk = 0
    StatusExit = 1
End Enum

Public Enum Factory
    sjz = 0
    gsh = 1
End Enum

'2015-12-15 add by qingenjian
Public g_NonlineFactors(100) As String
Public g_CurrentConfigPlantform As Integer
Public g_GridA As String
Public g_GridB As String
Public g_GridC As String
Public g_LoginUser As String
Public g_StartLogin As Boolean
Public g_SuperOk As Boolean

Public g_QueryMethod As String
Public g_TareStartDate As String

'==================================================================================
'常量
'==================================================================================
Public Const conMaxNonlinerFactors = 16
Public Const conFactorAStartPos = 23
Public Const conFactorBStartPos = 40
Public Const conFactorCStartPos = 57
Public Const conGradStartPos = 18
Public Const conMaxFactorNum = 100
Public Const conBalanceAPos = 4
Public Const conBalanceBPos = 9
Public Const conBalanceCPos = 14

Public Const conUseCh = 12
Public Const conBordNum = 3

Public Const conFixedCols = 11

Public Const conFactory = gsh   ' !!! 编译时注意修改一下厂家

'==================================================================================
'全局变量
'==================================================================================
Public gGdhIni As New CGdhINI  'ini文件

'==================================================================================
'公共函数
'==================================================================================
Public Function ContentSort(ms As MSFlexGrid, cnt As Integer, ByRef ret() As String) As Boolean
    ReDim ret(cnt) As String
    Dim i As Integer
    Dim key As Date
    
    For i = 0 To cnt
        
    Next i
End Function
