VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BMPToolbar"
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
Implements IToolBarDef

Private Sub Class_Terminate()
    Call ModuleUtility.CleanUpMemory
End Sub

Private Property Get IToolBarDef_Caption() As String
    ' Toolbar's caption
    IToolBarDef_Caption = "SUSTAIN BMP Toolbar"
End Property
Private Property Get IToolBarDef_ItemCount() As Long
    'Define how many commands will be in the toolbar
    IToolBarDef_ItemCount = 14
End Property

Private Sub IToolBarDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Define the commands that will be on the toolbar.
    ' These should all be classes defined within this project.
    ' ID sets the classID, Group determines if this iterm starts a new group
Select Case Pos
    Case 0
        itemDef.ID = "SUSTAIN.bmpToolBR"
        itemDef.Group = False
    Case 1
        itemDef.ID = "SUSTAIN.bmpToolDP"
        itemDef.Group = False
    Case 2
        itemDef.ID = "SUSTAIN.bmpToolWP"
        itemDef.Group = False
    Case 3
        itemDef.ID = "SUSTAIN.bmpToolRB"
        itemDef.Group = False
    Case 4
        itemDef.ID = "SUSTAIN.bmpToolC"
        itemDef.Group = False
    Case 5
        itemDef.ID = "SUSTAIN.bmpToolIT"
        itemDef.Group = False
    Case 6
        itemDef.ID = "SUSTAIN.bmpToolVS"
        itemDef.Group = False
    Case 7
        itemDef.ID = "SUSTAIN.bmpToolGR"
        itemDef.Group = False
    Case 8
        itemDef.ID = "SUSTAIN.bmpToolPP"
        itemDef.Group = False
    Case 9
        itemDef.ID = "SUSTAIN.AddAggregate"
        itemDef.Group = False
    Case 10
        itemDef.ID = "SUSTAIN.bmpToolREG"
        itemDef.Group = False
    Case 11
        itemDef.ID = "SUSTAIN.AddJunction"
        itemDef.Group = False
    Case 12
        itemDef.ID = "SUSTAIN.AddVOutlet"
        itemDef.Group = True
    Case 13
        itemDef.ID = "SUSTAIN.AddVFS"
        itemDef.Group = True
    End Select
End Sub

Private Property Get IToolBarDef_Name() As String
    ' Toolbar's name
    IToolBarDef_Name = "SustainBMPToolBar"
End Property



