VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SUSTAINMenu"
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
Implements IMenuDef
Implements IRootLevelMenu

Private Sub Class_Terminate()
    Call ModuleUtility.CleanUpMemory
End Sub

Private Property Get IMenuDef_Caption() As String
    ' Menu caption
    IMenuDef_Caption = "SUSTAIN"
End Property

Private Sub IMenuDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Menu items
    Select Case Pos
    Case 0
        itemDef.ID = "SUSTAIN.DataManageMenu"
        itemDef.Group = True
    Case 1
        itemDef.ID = "SUSTAIN.SimulationOption"
        itemDef.Group = True
'    Case 1
'        itemDef.ID = "SUSTAIN.DefinePollutants"
'        itemDef.Group = False
    Case 2
        itemDef.ID = "SUSTAIN.ExternalSimulationOptionsMenu"
        itemDef.Group = False
    Case 3
        itemDef.ID = "SUSTAIN.InternalSimulationOptionsMenu"
        itemDef.Group = False
    Case 4
        itemDef.ID = "SUSTAIN.DefineBMPs"
        itemDef.Group = True
    Case 5
        itemDef.ID = "SUSTAIN.AddBMP"
        itemDef.Group = False
    Case 6
        itemDef.ID = "SUSTAIN.AddOnlineBMP"
        itemDef.Group = False
    Case 7
        itemDef.ID = "SUSTAIN.ActivateBufStrip"
        itemDef.Group = False
    Case 8
        itemDef.ID = "SUSTAIN.Delineate"
        itemDef.Group = True
    Case 9
        itemDef.ID = "SUSTAIN.ActivateBasinBMPTool"
        itemDef.Group = False
    Case 10
        itemDef.ID = "SUSTAIN.ActivateBMPRoutingTool"
        itemDef.Group = False
    Case 11
        itemDef.ID = "SUSTAIN.CreateNetwork"
        itemDef.Group = False
'    Case 11
'        itemDef.ID = "SUSTAIN.AssignExternalTimeseries"
'        itemDef.Group = False
    Case 12
        itemDef.ID = "SUSTAIN.SimulateInternalTimeseries"
        itemDef.Group = True
    Case 13
        itemDef.ID = "SUSTAIN.DistributeLanduse"
        itemDef.Group = True
    Case 14
        itemDef.ID = "SUSTAIN.DefineAssessPointMenu"
        itemDef.Group = True
    Case 15
        itemDef.ID = "SUSTAIN.OptimizationMenu"
        itemDef.Group = False
    Case 16
        itemDef.ID = "SUSTAIN.CreateInputFile"
        itemDef.Group = True
    Case 17
        itemDef.ID = "SUSTAIN.EditInputFile"
        itemDef.Group = False
    Case 18
        itemDef.ID = "SUSTAIN.Simulation"
        itemDef.Group = False
    Case 19
        itemDef.ID = "SUSTAIN.ResultsToolbar"
        itemDef.Group = False
    Case 20
        itemDef.ID = "SUSTAIN.AboutSUSTAIN"
        itemDef.Group = True
    End Select
End Sub

Private Property Get IMenuDef_ItemCount() As Long
    ' Menu count
    IMenuDef_ItemCount = 21
End Property

Private Property Get IMenuDef_Name() As String
    'Menu name
    IMenuDef_Name = "SUSTAIN"
End Property
