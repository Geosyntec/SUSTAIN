VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BMPCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: Sustain - BMP Siting Tool
'   Company:     Tetra Tech, Inc
'   History:     Created:
'******************************************************************************

Option Explicit
Implements ICommand
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("BMP", vbResBitmap)
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "&Place BMPs"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "BMP_Siting_Tool"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableExtension
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Place BMPs"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "BMPCommand"
End Property

Private Sub ICommand_OnClick()
       
    ' Initialize....
    Call InitializeMapDocument
    
    If Not CheckMapDocumentSavedStatus_ST Then Exit Sub
    ' Set the BMP Type.....
    BMP_Siting_Tool.frmBMPCriteria.m_BMPType = "*"
    BMP_Siting_Tool.frmBMPCriteria.Show vbModal
    
End Sub

Private Sub ICommand_OnCreate(ByVal Hook As Object)
    Set gApplication = Hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "BMP Siting Tool"
End Property












