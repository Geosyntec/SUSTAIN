VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MergeWatershed"
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
Implements ITool

Private m_pFeedbackEnv As INewEnvelopeFeedback
Private m_pPoint As IPoint
Private m_bIsMouseDown As Boolean
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("MERGE", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_pFeedbackEnv = Nothing
    Set m_pPoint = Nothing
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
  ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
  ICommand_Caption = "Merge Selected Pictures"
End Property

Private Property Get ICommand_Category() As String
  ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean
  ICommand_Checked = False
End Property

Private Property Get ICommand_Enabled() As Boolean
  ICommand_Enabled = ModuleMenu.EnableDelineation
End Property

Private Property Get ICommand_HelpContextID() As Long
  
End Property

Private Property Get ICommand_HelpFile() As String
  
End Property

Private Property Get ICommand_Message() As String
  ICommand_Message = "Merge Selected Features"
End Property

Private Property Get ICommand_Name() As String
  ICommand_Name = "Merge Selected Features"
End Property

Private Sub ICommand_OnClick()
    InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
  Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
  ICommand_Tooltip = "Merge Selected Features"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
  If Not m_bIsMouseDown Then ' not in the middle of rubber banding
    ITool_Cursor = vbDefault
  Else
    ITool_Cursor = vbCrosshair
  End If
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
  ' Get the ActiveView for the map
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.FocusMap
  
  'Store current point, set mousedown flag
  Set m_pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  m_bIsMouseDown = True
  
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler
  
  If Not m_bIsMouseDown Then Exit Sub
  
  ' Get the ActiveView for the map
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.FocusMap
  
  ' Create a rubber banding box, if it hasn't been created already
  If (m_pFeedbackEnv Is Nothing) Then
    Set m_pFeedbackEnv = New NewEnvelopeFeedback
    Set m_pFeedbackEnv.Display = pActiveView.ScreenDisplay
    m_pFeedbackEnv.Start m_pPoint
  End If
  
  'Store current point, and use to move rubberband
  Set m_pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  m_pFeedbackEnv.MoveTo m_pPoint

  Exit Sub
ErrorHandler:
  MsgBox "MergeWatershed_ITool_OnMouseMove: " & Err.description
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  Dim pEnv As IEnvelope
  Dim pActiveView As IActiveView
  
  On Error GoTo ErrorHandler
  
  ' Get the ActiveView for the map
  Set pActiveView = gMxDoc.FocusMap

 ' Refresh the selections
  pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing

  ' If user dragged an envelope...
  If (Not m_pFeedbackEnv Is Nothing) Then
    'Use it to calculate new extent
    Set pEnv = m_pFeedbackEnv.Stop
    
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    If (pFeatureLayer Is Nothing) Then
        Exit Sub
    End If

    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pSpatialFilter.Geometry = pEnv
    'Merge Selected Features
    Dim pFeatureSelection As IFeatureSelection
    Set pFeatureSelection = pFeatureLayer
    'Merge Selected Features spatially
    pFeatureSelection.SelectFeatures pSpatialFilter, esriSelectionResultNew, False

    Set pSpatialFilter = Nothing
    Set pFeatureSelection = Nothing
    Set pFeatureLayer = Nothing
    
    ' Refresh the selections
    pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
  
    'Prompt user to merge all selected features ?
    Dim boolMerge
    boolMerge = MsgBox("Do you want to merge selected features ?", vbYesNo)
    
    If (boolMerge = vbYes) Then
        MergeSelectedFeatures pEnv
        'Renumber watershed ID's
        RenumberWatershedFeatures
    End If
  End If
  
  'Reset rubberband and mousedown state
   Set pEnv = Nothing
  Set m_pFeedbackEnv = Nothing
  m_bIsMouseDown = False

  Exit Sub

ErrorHandler:
 MsgBox "MergeWatershed_ITool_OnMouseUp: " & Err.description
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub


