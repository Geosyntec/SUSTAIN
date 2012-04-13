VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActivatePostProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 03/29/2005 - mira chokshi added comments to project
'
'******************************************************************************

Option Explicit
Implements ICommand
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("Excel", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "View Results In MS Excel"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableSimulation
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "View Results In MS Excel"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "View Results In MS Excel"
End Property

Private Sub ICommand_OnClick()
    Dim pbars As ICommandBars
    Set pbars = gApplication.Document.CommandBars
    
    'show the Edit BMP toolbar
    Dim pUID As UID
    Set pUID = New UID
    'Below is a toolbar in this project (BMPTool..."
    pUID.value = "SUSTAIN.SUSTAINToolBar"
    Dim pbar As ICommandBar
    Set pbar = pbars.Find(pUID)
    pbar.Dock esriDockShow
        
    'Make the define assessment point tool as active tool
    Dim pUID1 As UID
    Set pUID1 = New UID
    'Below is a toolbar in this project (BMPTool..."
    pUID1.value = "SUSTAIN.PostProcessor"
    
    Dim pSelectTool As ICommandItem
    Set pSelectTool = pbars.Find(pUID1)
     
    'Set the current tool of the application to be the Select Graphics Tool
    Set gApplication.CurrentTool = pSelectTool

End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "View Results In MS Excel"
End Property



