VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeleteWatershed"
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_pEditor As IEditor
Private m_pEditSketch As IEditSketch
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("DELETE", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
    Set m_pEditor = Nothing
    Set m_pEditSketch = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Delete Watershed"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableDelineation
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Delete Watershed"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Delete Watershed"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Delete Watershed"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE

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

'*** Get feature selection and merge with selected watershed
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

On Error GoTo ShowError

  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Dim pPoint As IPoint
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
 
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass

  Dim pEnvelope As IEnvelope
  Set pEnvelope = pPoint.Envelope
  ExpandPointEnvelope pEnvelope
    
  Dim pSpatialFilter As ISpatialFilter
  Set pSpatialFilter = New SpatialFilter
  Set pSpatialFilter.Geometry = pEnvelope
  pSpatialFilter.SpatialRel = esriSpatialRelIntersects
            
  Dim pFeatureCursor As IFeatureCursor
  Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, True)
  Dim pFeature As IFeature
  Set pFeature = pFeatureCursor.NextFeature

  Dim pDeleteSet As ISet
  Set pDeleteSet = New esriSystem.Set
  Dim pFeatureEdit As IFeatureEdit
  
  Dim boolDelete
  If (Not pFeature Is Nothing) Then
      FlashSelectedFeature pFeature
      boolDelete = MsgBox("Are you sure you want to delete this subwatershed ?", vbYesNo, "Delete Watershed")
      If (boolDelete = vbYes) Then
            Call StartEditingFeatureLayer("Watershed")
            pDeleteSet.add pFeature
            pDeleteSet.Reset
            Set pFeatureEdit = pDeleteSet.Next
            If Not pFeatureEdit Is Nothing Then
              pFeatureEdit.DeleteSet pDeleteSet
            End If
            Call StopEditingFeatureLayer
      End If
  End If

  Set pActiveView = Nothing
  Set pPoint = Nothing
  Set pFeatureLayer = Nothing
  Set pFeatureclass = Nothing
  Set pEnvelope = Nothing
  Set pSpatialFilter = Nothing
  Set pFeatureCursor = Nothing
  Set pFeature = Nothing
  Set pDeleteSet = Nothing
  Set pFeatureEdit = Nothing

  'Renumber watershed ID's
  RenumberWatershedFeatures
  
  Exit Sub
  
ShowError:
    MsgBox "DeleteWatershed_ITool_OnMouseDown: " & Err.description & vbTab & Err.Number
'Cleanup again, in case of error
  Set pActiveView = Nothing
  Set pPoint = Nothing
  Set pFeatureLayer = Nothing
  Set pFeatureclass = Nothing
  Set pEnvelope = Nothing
  Set pSpatialFilter = Nothing
  Set pFeatureCursor = Nothing
  Set pFeature = Nothing
  Set pDeleteSet = Nothing
  Set pFeatureEdit = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub

