VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SustainToolbar"
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
Implements IToolBarDef

Private Sub Class_Terminate()
    
    Set gMxDoc = Nothing
    Set gMap = Nothing
    
End Sub

Private Property Get IToolBarDef_Caption() As String
    ' Toolbar's caption
    IToolBarDef_Caption = "BMP Siting Tool"
End Property
Private Property Get IToolBarDef_ItemCount() As Long
    'Define how many commands will be in the toolbar
    IToolBarDef_ItemCount = 1
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Define the commands that will be on the toolbar.
    ' These should all be classes defined within this project.
    ' ID sets the classID, Group determines if this iterm starts a new group
Select Case Pos
    Case 0
        itemDef.ID = "BMP_Siting_Tool.BMPCommand"
        itemDef.Group = False
    End Select
End Sub

Private Property Get IToolBarDef_Name() As String
    ' Toolbar's name
    IToolBarDef_Name = "BMP_Siting_Tool"
End Property



