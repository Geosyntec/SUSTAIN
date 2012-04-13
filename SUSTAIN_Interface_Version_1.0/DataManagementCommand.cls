VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataManagementCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   History:     Created:
'******************************************************************************

Option Explicit
Implements ICommand
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    ''Set m_Bitmap = LoadPicture
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    'ICommand_Bitmap = m_Bitmap
End Property



Private Property Get ICommand_Caption() As String
    ICommand_Caption = "&Define Data Path "
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = EnableSustain 'True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Data Management/Working Directory"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "DataManagementCommand"
End Property

Private Sub ICommand_OnClick()
    Call ModuleUtility.InitializeMapDocument
    
    'Read data layer information from src file
    'Call SetDataDirectory
    
    'If Not CheckMapDocumentSavedStatus Then Exit Sub
    
    FrmDataManagement.Show vbModal
    
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Data Management/Working Directory"
End Property









