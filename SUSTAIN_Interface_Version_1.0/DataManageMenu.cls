VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataManageMenu"
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
Implements IMenuDef
Implements IRootLevelMenu

Private Sub Class_Terminate()
    'Call CleanUpMemory
End Sub

Private Property Get IMenuDef_Caption() As String
    ' Menu caption
    IMenuDef_Caption = "Data Management"
End Property

Private Sub IMenuDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Menu items
    Select Case Pos
    Case 0
        itemDef.ID = "SUSTAIN.DataManagementCommand"
        itemDef.Group = False
    Case 1
        itemDef.ID = "SUSTAIN.LoadDataCommand"
        itemDef.Group = False
    Case 2
        itemDef.ID = "SUSTAIN.ManageData"
        itemDef.Group = False
    Case 3
        itemDef.ID = "SUSTAIN.AddDataCommand"
        itemDef.Group = False
    End Select
End Sub

Private Property Get IMenuDef_ItemCount() As Long
    ' Menu count
    IMenuDef_ItemCount = 4
End Property

Private Property Get IMenuDef_Name() As String
    'Menu name
    IMenuDef_Name = "SUSTAIN"
End Property




