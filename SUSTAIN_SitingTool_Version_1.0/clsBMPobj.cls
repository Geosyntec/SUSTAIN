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
' Class : BMPObj
'----------------------------------------------------------------------------------------

' Object data table name
Private m_strBMPName As String
' ID of the object data table
Private m_lngBMPID As Long
' BMP Type....
Private m_strBMPType As String
'Properties.......
Private m_DC_DA As String
Private m_DC_DS As String
Private m_DC_IMP As String
Private m_DC_HG As String
Private m_DC_WT As String
Private m_DC_RB As String
Private m_DC_BB As String
Private m_DC_SB As String
'State properties.....
Private m_DC_DA_State As Byte
Private m_DC_DS_State As Byte
Private m_DC_IMP_State As Byte
Private m_DC_HG_State As Byte
Private m_DC_WT_State As Byte
Private m_DC_RB_State As Byte
Private m_DC_BB_State As Byte
Private m_DC_SB_State As Byte


'----------------------------------------------------------------------------------------
' Properties
'----------------------------------------------------------------------------------------

' Name of the BMP

Public Property Get BMPName() As String
    BMPName = m_strBMPName
End Property

Public Property Let BMPName(strNewValue As String)
    m_strBMPName = strNewValue
End Property

' Type of the BMP

Public Property Get BMPType() As String
    BMPType = m_strBMPType
End Property

Public Property Let BMPType(strNewValue As String)
    m_strBMPType = strNewValue
End Property

' ID of the BMP

Public Property Get BMPId() As Long
    BMPId = m_lngBMPID
End Property

Public Property Let BMPId(lngNewValue As Long)
    m_lngBMPID = lngNewValue
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

Public Property Get DC_HG() As String
    DC_HG = m_DC_HG
End Property

Public Property Let DC_HG(strNewValue As String)
    m_DC_HG = strNewValue
End Property

Public Property Get DC_IMP() As String
    DC_IMP = m_DC_IMP
End Property

Public Property Let DC_IMP(strNewValue As String)
    m_DC_IMP = strNewValue
End Property

Public Property Get DC_WT() As String
    DC_WT = m_DC_WT
End Property

Public Property Let DC_WT(strNewValue As String)
    m_DC_WT = strNewValue
End Property

Public Property Get DC_RB() As String
    DC_RB = m_DC_RB
End Property

Public Property Let DC_RB(strNewValue As String)
    m_DC_RB = strNewValue
End Property

Public Property Get DC_BB() As String
    DC_BB = m_DC_BB
End Property

Public Property Let DC_BB(strNewValue As String)
    m_DC_BB = strNewValue
End Property

Public Property Get DC_SB() As String
    DC_SB = m_DC_SB
End Property

Public Property Let DC_SB(strNewValue As String)
    m_DC_SB = strNewValue
End Property

'State Values.....
Public Property Get DC_DA_State() As Byte
    DC_DA_State = m_DC_DA_State
End Property

Public Property Let DC_DA_State(strNewValue As Byte)
    m_DC_DA_State = strNewValue
End Property

Public Property Get DC_DS_State() As Byte
    DC_DS_State = m_DC_DS_State
End Property

Public Property Let DC_DS_State(strNewValue As Byte)
    m_DC_DS_State = strNewValue
End Property

Public Property Get DC_IMP_State() As Byte
    DC_IMP_State = m_DC_IMP_State
End Property

Public Property Let DC_IMP_State(strNewValue As Byte)
    m_DC_IMP_State = strNewValue
End Property

Public Property Get DC_HG_State() As Byte
    DC_HG_State = m_DC_HG_State
End Property

Public Property Let DC_HG_State(strNewValue As Byte)
    m_DC_HG_State = strNewValue
End Property

Public Property Get DC_WT_State() As Byte
    DC_WT_State = m_DC_WT_State
End Property

Public Property Let DC_WT_State(strNewValue As Byte)
    m_DC_WT_State = strNewValue
End Property

Public Property Get DC_RB_State() As Byte
    DC_RB_State = m_DC_RB_State
End Property

Public Property Let DC_RB_State(strNewValue As Byte)
    m_DC_RB_State = strNewValue
End Property

Public Property Get DC_BB_State() As Byte
    DC_BB_State = m_DC_BB_State
End Property

Public Property Let DC_BB_State(strNewValue As Byte)
    m_DC_BB_State = strNewValue
End Property

Public Property Get DC_SB_State() As Byte
    DC_SB_State = m_DC_SB_State
End Property

Public Property Let DC_SB_State(strNewValue As Byte)
    m_DC_SB_State = strNewValue
End Property

'----------------------------------------------------------------------------------------
' Methods
'----------------------------------------------------------------------------------------


