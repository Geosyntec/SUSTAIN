VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SustainMenu"
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
Implements IMenuDef
Implements IRootLevelMenu

Private Sub Class_Terminate()

    Set gMxDoc = Nothing
    Set gMap = Nothing
    
End Sub

Private Property Get IMenuDef_Caption() As String
    ' Menu caption
    IMenuDef_Caption = "BMP Siting Tool"
End Property

Private Sub IMenuDef_GetItemInfo(ByVal Pos As Long, ByVal itemDef As IItemDef)
    ' Menu items
    Select Case Pos
    Case 0
        itemDef.ID = "BMP_Siting_Tool.DataManageCommand"
        itemDef.Group = False
    Case 1
        itemDef.ID = "BMP_Siting_Tool.BMPCommand"
        itemDef.Group = False
    Case 2
        itemDef.ID = "BMP_Siting_Tool.AboutBMPCommand"
        itemDef.Group = True
    End Select
End Sub

Private Property Get IMenuDef_ItemCount() As Long
    ' Menu count
    IMenuDef_ItemCount = 3
End Property

Private Property Get IMenuDef_Name() As String
    'Menu name
    IMenuDef_Name = "BMP_Siting_Tool"
End Property

