Attribute VB_Name = "ModuleMenu"

'******************************************************************************
'   Application: Sustain - BMP Siting Tool
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Arun Raj
'   Developer:   Arun Raj
'******************************************************************************


Option Explicit
Option Base 0
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\ModuleMenu.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms




Public Function EnableExtension() As Boolean
  On Error GoTo ErrorHandler

    EnableExtension = False
   
    Dim u As New UID
    u.Value = "BMP_Siting_Tool.BMPExtension"
    Dim m_pExt As IExtensionConfig
    Set m_pExt = gApplication.FindExtensionByCLSID(u)
    'EnableExtension = (m_pExt.State = esriESEnabled)
    EnableExtension = True

  Exit Function
ErrorHandler:
  HandleError True, "EnableExtension " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function
