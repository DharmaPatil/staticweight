VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSectionSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'数据预处理配置项
Option Explicit
'=========================================================================================================
'常量
'=========================================================================================================
Const conSection = "Save"

'=========================================================================================================
'局部声明
'=========================================================================================================

'=========================================================================================================
'变量
'=========================================================================================================
Private m_Ini As CGdhINI    '指向包含自身的ini文件

Private m_Path As String

Private m_SaveRaw As Boolean
Private m_SaveRef As Boolean
Private m_SaveFlt As Boolean
Private m_SaveDif As Boolean
Private m_SaveIdx As Boolean
Private m_SaveRst As Boolean

'=========================================================================================================
'方法
'=========================================================================================================
'从文件中读所有key
Public Sub GetAllKeys()
    m_Path = m_Ini.GetString(conSection, "Path", "c:\")
    m_SaveRaw = m_Ini.GetBoolean(conSection, "SaveRaw", False)
    m_SaveRef = m_Ini.GetBoolean(conSection, "SaveRef", False)
    m_SaveFlt = m_Ini.GetBoolean(conSection, "SaveFlt", False)
    m_SaveDif = m_Ini.GetBoolean(conSection, "SaveDif", False)
    m_SaveIdx = m_Ini.GetBoolean(conSection, "SaveIdx", False)
    m_SaveRst = m_Ini.GetBoolean(conSection, "SaveRst", False)
End Sub

'=========================================================================================================
'属性
'=========================================================================================================
'ini
Public Property Set Ini(vData As CGdhINI)
    Set m_Ini = vData
End Property

'Path
Public Property Let Path(vData As String)
    m_Path = vData
    m_Ini.SaveString conSection, "Path", m_Path
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

'SaveRaw
Public Property Let SaveRaw(vData As Boolean)
    m_SaveRaw = vData
    m_Ini.SaveBoolean conSection, "SaveRaw", m_SaveRaw
End Property

Public Property Get SaveRaw() As Boolean
    SaveRaw = m_SaveRaw
End Property

'SaveRef
Public Property Let SaveRef(vData As Boolean)
    m_SaveRef = vData
    m_Ini.SaveBoolean conSection, "SaveRef", m_SaveRef
End Property

Public Property Get SaveRef() As Boolean
    SaveRef = m_SaveRef
End Property

'SaveFlt
Public Property Let SaveFlt(vData As Boolean)
    m_SaveFlt = vData
    m_Ini.SaveBoolean conSection, "SaveFlt", m_SaveFlt
End Property

Public Property Get SaveFlt() As Boolean
    SaveFlt = m_SaveFlt
End Property

'SaveDif
Public Property Let SaveDif(vData As Boolean)
    m_SaveDif = vData
    m_Ini.SaveBoolean conSection, "SaveDif", m_SaveDif
End Property

Public Property Get SaveDif() As Boolean
    SaveDif = m_SaveDif
End Property

'SaveIdx
Public Property Let SaveIdx(vData As Boolean)
    m_SaveIdx = vData
    m_Ini.SaveBoolean conSection, "SaveIdx", m_SaveIdx
End Property

Public Property Get SaveIdx() As Boolean
    SaveIdx = m_SaveIdx
End Property

'SaveRst
Public Property Let SaveRst(vData As Boolean)
    m_SaveRst = vData
    m_Ini.SaveBoolean conSection, "SaveRst", m_SaveRst
End Property

Public Property Get SaveRst() As Boolean
    SaveRst = m_SaveRst
End Property

