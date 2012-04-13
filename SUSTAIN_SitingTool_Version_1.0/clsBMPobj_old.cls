VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BMPobj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'----------------------------------------------------------------------------------------
' Class : clsBMP
'
' Encapsulates the definition of a table used by a feature class.
' Holds a collection of field definitions.
' As each field can hold a current value it can be used as a "record" by an editor.
'----------------------------------------------------------------------------------------

' Object data table name
Private m_strBMPType As String
' ID of the object data table
Private m_lngBMPID As Long
'Properties.......
Private m_DR_EL As Boolean
Private m_DR_SL As Boolean
Private m_DR_UL As Boolean
Private m_DR_WT As Boolean
Private m_DR_RD As Boolean

Private m_DC_DA As String
Private m_DC_DS As String
Private m_DC_LS As String
Private m_DC_HG As String
Private m_DC_WT As String
Private m_DC_RD As String

Private m_SU_ND As Boolean
Private m_SU_PL As Boolean
Private m_SU_RD As Boolean
Private m_SU_DW As Boolean

'----------------------------------------------------------------------------------------
' Properties
'----------------------------------------------------------------------------------------

' Name of the OD table

Public Property Get BMPType() As String
    BMPType = m_strBMPType
End Property

Public Property Let BMPType(strNewValue As String)
    m_strBMPType = strNewValue
End Property

' ID of the OD table

Public Property Get BMPId() As Long
    BMPId = m_lngBMPID
End Property

Public Property Let BMPId(lngNewValue As Long)
    m_lngBMPID = lngNewValue
End Property

' Flags.........
Public Property Get DR_EL() As Boolean
    DR_EL = m_DR_EL
End Property

Public Property Let DR_EL(boolNewValue As Boolean)
    m_DR_EL = boolNewValue
End Property

Public Property Get DR_SL() As Boolean
    DR_SL = m_DR_SL
End Property

Public Property Let DR_SL(boolNewValue As Boolean)
    m_DR_SL = boolNewValue
End Property

Public Property Get DR_UL() As Boolean
    DR_UL = m_DR_UL
End Property

Public Property Let DR_UL(boolNewValue As Boolean)
    m_DR_UL = boolNewValue
End Property

Public Property Get DR_WT() As Boolean
    DR_WT = m_DR_WT
End Property

Public Property Let DR_WT(boolNewValue As Boolean)
    m_DR_WT = boolNewValue
End Property

Public Property Get DR_RD() As Boolean
    DR_RD = m_DR_RD
End Property

Public Property Let DR_RD(boolNewValue As Boolean)
    m_DR_RD = boolNewValue
End Property

'Design Criteria....
Public Property Get DC_DA() As String
    DC_DA = m_DC_DA
End Property

Public Property Let DC_DA(strNewValue As String)
    m_DC_DA = strNewValue
End Property

Public Property Get DC_DS() As String
    DC_DS = m_DC_DS
End Property

Public Property Let DC_DS(strNewValue As String)
    m_DC_DS = strNewValue
End Property

Public Property Get DC_LS() As String
    DC_LS = m_DC_LS
End Property

Public Property Let DC_LS(strNewValue As String)
    m_DC_LS = strNewValue
End Property

Public Property Get DC_HG() As String
    DC_HG = m_DC_HG
End Property

Public Property Let DC_HG(strNewValue As String)
    m_DC_HG = strNewValue
End Property

Public Property Get DC_WT() As String
    DC_WT = m_DC_WT
End Property

Public Property Let DC_WT(strNewValue As String)
    m_DC_WT = strNewValue
End Property

Public Property Get DC_RD() As String
    DC_RD = m_DC_RD
End Property

Public Property Let DC_RD(strNewValue As String)
    m_DC_RD = strNewValue
End Property


'Suitability......
Public Property Get SU_ND() As Boolean
    SU_ND = m_SU_ND
End Property

Public Property Let SU_ND(boolNewValue As Boolean)
    m_SU_ND = boolNewValue
End Property

Public Property Get SU_PL() As Boolean
    SU_PL = m_SU_PL
End Property

Public Property Let SU_PL(boolNewValue As Boolean)
    m_SU_PL = boolNewValue
End Property

Public Property Get SU_RD() As Boolean
    SU_RD = m_SU_RD
End Property

Public Property Let SU_RD(boolNewValue As Boolean)
    m_SU_RD = boolNewValue
End Property

Public Property Get SU_DW() As Boolean
    SU_DW = m_SU_DW
End Property

Public Property Let SU_DW(boolNewValue As Boolean)
    m_SU_DW = boolNewValue
End Property



'----------------------------------------------------------------------------------------
' Methods
'----------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    
End Sub


