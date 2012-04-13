VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bmpToolWP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Implements ITool

Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("bmpWP", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Wet Pond"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableIndividualBMPTool("WetPond")
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Wet Pond"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "WetPond"
End Property

Private Sub ICommand_OnClick()

End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Wet Pond"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("bmpWP", vbResCursor)
End Property

Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
  
End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    '*** Call common function to add this BMP
    AddBMPFromToolbox Button, Shift, X, Y, "WetPond"
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub
















