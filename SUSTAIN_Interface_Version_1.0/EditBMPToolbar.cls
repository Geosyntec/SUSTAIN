VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditBMPToolbar"
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
Implements IToolBarDef

Private Sub Class_Terminate()
    Call ModuleUtility.CleanUpMemory
End Sub

Private Property Get IToolBarDef_Caption() As String
    ' Toolbar's caption
    IToolBarDef_Caption = "SUSTAIN Edit BMP Toolbar"
End Property
Private Property Get IToolBarDef_ItemCount() As Long
    'Define how many commands will be in the toolbar
    IToolBarDef_ItemCount = 5
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Define the commands that will be on the toolbar.
    ' These should all be classes defined within this project.
    ' ID sets the classID, Group determines if this iterm starts a new group
Select Case Pos
    Case 0
        itemDef.ID = "SUSTAIN.bmpToolMove"
        itemDef.Group = False
    Case 1
        itemDef.ID = "SUSTAIN.DeleteBMPs"
        itemDef.Group = False
    Case 2
        itemDef.ID = "SUSTAIN.EditBMPs"
        itemDef.Group = False
    Case 3
        itemDef.ID = "SUSTAIN.DefineAssessPoint"
        itemDef.Group = True
    Case 4
        itemDef.ID = "SUSTAIN.DeleteAssessPoint"
        itemDef.Group = False
End Select

End Sub

Private Property Get IToolBarDef_Name() As String
    ' Toolbar's name
    IToolBarDef_Name = "SustainEditBMPToolbar"
End Property





