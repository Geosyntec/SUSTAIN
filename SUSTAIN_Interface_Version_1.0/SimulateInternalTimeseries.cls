VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimulateInternalTimeseries"
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
Implements IMenuDef
Implements IRootLevelMenu

Private Sub Class_Terminate()
    Call CleanUpMemory
End Sub


Private Property Get IMenuDef_Caption() As String
    ' Menu caption
    IMenuDef_Caption = "Simulate Timeseries for Subwatersheds"
End Property

Private Sub IMenuDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Menu items
    Select Case Pos
    Case 0
        itemDef.ID = "SUSTAIN.SWMMDefineClimatology"
        itemDef.Group = False
    Case 1
        itemDef.ID = "SUSTAIN.SWMMDefinePollutants"
        itemDef.Group = False
    Case 2
        itemDef.ID = "SUSTAIN.SWMMDefineLanduseProperties"
        itemDef.Group = False
    Case 3
        itemDef.ID = "SUSTAIN.SWMMDefineRainGages"
        itemDef.Group = False
    Case 4
        itemDef.ID = "SUSTAIN.SWMMDefineAquifers"
        itemDef.Group = True
    Case 5
        itemDef.ID = "SUSTAIN.SWMMDefineSnowPack"
        itemDef.Group = False
    Case 6
        itemDef.ID = "SUSTAIN.SWMMActivateSubBasinTool"
        itemDef.Group = False
    Case 7
        itemDef.ID = "SUSTAIN.SWMMDefineOptions"
        itemDef.Group = True
    Case 8
        itemDef.ID = "SUSTAIN.SWMMEditInputFile"
        itemDef.Group = False
    Case 9
        itemDef.ID = "SUSTAIN.SWMMRunSimulation"
        itemDef.Group = False
    End Select
End Sub

Private Property Get IMenuDef_ItemCount() As Long
    ' Menu count
    IMenuDef_ItemCount = 10
End Property

Private Property Get IMenuDef_Name() As String
    'Menu name
    IMenuDef_Name = "SUSTAIN"
End Property






