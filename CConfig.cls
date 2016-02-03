VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'配置文件类
Option Explicit
'=========================================================================================================
'常量
'=========================================================================================================
Const conMaxKeySize = 255

'=========================================================================================================
'局部声明
'=========================================================================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSection As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

'=========================================================================================================
'变量
'=========================================================================================================
Private m_FileName As String    'ini文件名称
'
'属性
'
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(Value As String)
    m_FileName = Value
End Property

'string类型
Public Sub SaveString(SSection$, SKey$, sValue$)
    If WritePrivateProfileString(SSection, SKey, sValue, m_FileName) = 0 Then
        MsgBox "写ini文件错误。" & " Section=" & SSection & " Key=" & SKey, vbCritical, "错误"
    End If
End Sub

Public Function GetString(SSection$, SKey$, sDefault$) As String
    Dim sReturn As String
    Dim I As Integer
    
    sReturn = String(conMaxKeySize + 1, " ")
    If GetPrivateProfileString(SSection, SKey, sDefault, sReturn, conMaxKeySize, m_FileName) = 0 Then
        MsgBox "读ini文件错误。" & " Section=" & SSection & " Key=" & SKey, vbCritical, "错误"
    End If
    For I = 1 To Len(sReturn)
        If Mid(sReturn, I, 1) = Chr(0) Then
           sReturn = Left(sReturn, I - 1)
           Exit For
        End If
    Next I
    GetString = Trim(sReturn)
End Function

'boolean
Public Function GetBoolean(SSection$, SKey$, bDefault As Boolean) As Boolean
    Dim sVal$, sDefault$
    If bDefault Then
        sDefault = "1"
    Else
        sDefault = "0"
    End If
    sVal = GetString(SSection, SKey, sDefault)
    If CLng(sVal) = 0 Then
        GetBoolean = False
    Else
        GetBoolean = True
    End If
End Function

Public Sub SaveBoolean(SSection$, SKey$, bValue As Boolean)
    Dim sVal$
    If bValue Then
        sVal = "1"
    Else
        sVal = "0"
    End If
     SaveString SSection, SKey, sVal
End Sub

'Integer
Public Function GetInteger(SSection$, SKey$, vDefault As Integer) As Integer
    GetInteger = CInt(GetString(SSection, SKey, CStr(vDefault)))
End Function

Public Sub SaveInteger(SSection$, SKey$, vValue As Integer)
     SaveString SSection, SKey, CStr(vValue)
End Sub

'Long
Public Function GetLong(SSection$, SKey$, vDefault As Long) As Long
    GetLong = CLng(GetString(SSection, SKey, CStr(vDefault)))
End Function

Public Sub SaveLong(SSection$, SKey$, vValue As Long)
     SaveString SSection, SKey, CStr(vValue)
End Sub

'Single
Public Function GetSingle(SSection$, SKey$, vDefault As Single) As Single
    GetSingle = CSng(GetString(SSection, SKey, CStr(vDefault)))
End Function

Public Sub SaveSingle(SSection$, SKey$, vValue As Single)
     SaveString SSection, SKey, CStr(vValue)
End Sub

Private Sub Class_Terminate()
'
End Sub

