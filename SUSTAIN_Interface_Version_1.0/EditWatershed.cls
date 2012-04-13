VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditWatershed"
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

Private m_pEditor As IEditor
Private m_pEditSketch As IEditSketch
Private m_Bitmap As IPictureDisp
Private m_pPoint As IPoint
Private m_pFeedback As IDisplayFeedback
Private m_bInUse As Boolean
Private m_pAnchorPoint As IPoint
Private m_pEditFeature As IFeature
Private m_pCursor As Object

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("EDIT", vbResBitmap)
    Set m_pCursor = LoadResPicture("Edit", vbResCursor)
End Sub

Private Sub Class_Terminate()
    Set m_pEditor = Nothing
    Set m_pEditSketch = Nothing
    Set m_Bitmap = Nothing
    Set m_pPoint = Nothing
    Set m_pFeedback = Nothing
    Set m_pAnchorPoint = Nothing
    Set m_pEditFeature = Nothing
    Set m_pCursor = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Edit Delineation"
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
    ICommand_Message = "Edit Delineation"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "EditDelineation"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Call StartEditingFeatureLayer("Watershed")
    Call EditWatershedEdges
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Edit Delineation"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = m_pCursor
End Property

Private Function ITool_Deactivate() As Boolean
     ITool_Deactivate = True
    'Get the graphics container of the map and clear it
     Dim pGraphicContainer As IGraphicsContainer
     Set pGraphicContainer = gMap
     pGraphicContainer.DeleteAllElements
     gMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
     Set pGraphicContainer = Nothing
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
On Error GoTo ShowError

  m_bInUse = True
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Dim pPoint As IPoint
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  ' first attempt to add a vertex or change an existing one
  ' if that fails then move the object. Only possible to edit a feature
  ' if it is preselected
  If (Not EditFeatureLineOrVertex(pPoint)) Then
    ' ensure that there is at least one selected in display
    If (Not EnsureOneSelected(pPoint)) Then
      ' We are in edit mode but there are no suitable features to edit
      ' so the tool cannot be active
      m_bInUse = False
    End If
  End If

  GoTo CleanUp
ShowError:
  MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp:
  Set pPoint = Nothing
  Set pActiveView = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError

  Dim pPoint As IPoint
  Dim pEnvFeedback As INewEnvelopeFeedback
  Dim pHitPoint As IPoint
  Dim hitDist As Double
  Dim vertex As Boolean
  Dim vertexIndex As Long
  Dim partIndex As Long
  Dim gotHit As Boolean
  Dim pSelected As IEnumFeature
  Dim pFeature As IFeature
  Dim pGeom As IGeometry
  Dim pMove As IMoveGeometryFeedback
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  
  ' Update the co-ord display
  If ((m_bInUse = False) And (gMap.SelectionCount > 0)) Then
    ' The edit tool is active, but are not currently editing so provide some
    ' feedback to the user regards if the vertex or line is hit
    ' perform the hit test against ALL the selected features, abort
    ' when the first hit is found
    Set pSelected = gMap.FeatureSelection
    pSelected.Reset
    Set pFeature = pSelected.Next
    Do While (Not pFeature Is Nothing)
      If (TestGeometryHit(ConvertPixelsToMapUnits(gMap, 4), pPoint, pFeature, pHitPoint, _
                        hitDist, partIndex, vertexIndex, vertex) = True) Then
        Set pFeature = Nothing
        If (vertex) Then
          Set m_pCursor = LoadResPicture("VertexEdit", vbResCursor)
        Else
          Set m_pCursor = LoadResPicture("LineEdit", vbResCursor)
        End If
        gotHit = True
      Else
        Set pFeature = pSelected.Next
      End If
    Loop
    If (Not gotHit) Then Set m_pCursor = LoadResPicture("Edit", vbResCursor)
    Exit Sub
  ElseIf (m_bInUse = False) Then
    Exit Sub
  End If
  
  If (m_pFeedback Is Nothing) Then
    Set m_pFeedback = New MoveGeometryFeedback
    Set m_pFeedback.Display = pActiveView.ScreenDisplay
    Set pMove = m_pFeedback
    Set pSelected = gMap.FeatureSelection
    pSelected.Reset
    Set pFeature = pSelected.Next
    Do While (Not pFeature Is Nothing)
      Set pGeom = pFeature.ShapeCopy
      pMove.AddGeometry pGeom

      Set pFeature = pSelected.Next
    Loop
    pMove.Start pPoint
    Set m_pAnchorPoint = pPoint
    Set m_pCursor = LoadResPicture("EditMove", vbResCursor)
  End If

  If (Not m_pFeedback Is Nothing) Then m_pFeedback.MoveTo pPoint

  GoTo CleanUp
  
ShowError:
  MsgBox "ITool_OnMouseMove: " & Err.description
CleanUp:
  Set pPoint = Nothing
  Set pEnvFeedback = Nothing
  Set pHitPoint = Nothing
  Set pSelected = Nothing
  Set pFeature = Nothing
  Set pGeom = Nothing
  Set pMove = Nothing
  Set pActiveView = Nothing
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  EditFeatures pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
 
  m_bInUse = False
  Set m_pFeedback = Nothing
  
  '*** Highlight selected nodes
  Call HighlightSelectedNodes

  GoTo CleanUp
ShowError:
  MsgBox "ITool_OnMouseUp: " & Err.description
CleanUp:
  Set pActiveView = Nothing
End Sub



Private Function EditFeatureLineOrVertex(pPoint As IPoint) As Boolean
On Error GoTo ShowError

  Dim vertexIndex As Long
  Dim pGeom As IGeometry
  Dim pHitPoint As IPoint
  Dim hitDist As Double
  Dim partIndex As Long
  Dim pFeature As IFeature
  Dim pObjectClass As IObjectClass
  Dim pPath As IPath
  Dim tol As Double
  Dim vertex As Boolean
  Dim pPointColn As IPointCollection
  Dim numVertices As Long
  Dim pLineMove As ILineMovePointFeedback
  Dim pPolyMove As IPolygonMovePointFeedback
  
 
  If (gMap.SelectionCount < 1) Then Exit Function
  
  ' tolerance in pixels for line hits
  tol = ConvertPixelsToMapUnits(gMap, 4)
 
  ' find the closest feature to the cursor position
  GetClosestSelectedFeature pPoint, pFeature
  If (pFeature Is Nothing) Then Exit Function
  
  ' the feedback action / edit action depends on the geometry type
  ' and the location of point within the geometry
  Set pGeom = pFeature.Shape
  Set pObjectClass = pFeature.Class
  
  ' Nothing here to do if the feature is a point
  If ((pGeom.GeometryType = esriGeometryPoint) Or _
      (pGeom.GeometryType = esriGeometryMultipoint)) Then Exit Function

  Set m_pEditFeature = pFeature
  ' if control is depressed edit operation is a reshape
  ' otherwise it is move poly point or a stretch line
    
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Select Case pGeom.GeometryType
    Case esriGeometryPolyline
      If (TestGeometryHit(tol, pPoint, pFeature, pHitPoint, hitDist, partIndex, vertexIndex, vertex) = True) Then
        If (Not vertex) Then
          ' get the path, add a point to it and vertex edit that newly added point
          
          Set pPath = GetPart(pGeom, partIndex)
          Set pPointColn = pPath
          numVertices = pPointColn.PointCount
          
          If vertexIndex = 0 Then
            pPointColn.AddPoint pPoint, 1
          Else
            pPointColn.AddPoint pPoint, , vertexIndex
          End If
          ' reset the index pointer to the new index
          TestGeometryHit tol, pPoint, pFeature, pHitPoint, hitDist, partIndex, vertexIndex, vertex
        End If
        Set m_pFeedback = New LineMovePointFeedback
        Set m_pFeedback.Display = pActiveView.ScreenDisplay
        Set pLineMove = m_pFeedback
        pLineMove.Start pGeom, vertexIndex, pPoint
      Else
        Exit Function
      End If
    Case esriGeometryPolygon
      If (TestGeometryHit(tol, pPoint, pFeature, pHitPoint, hitDist, partIndex, vertexIndex, vertex) = True) Then
        If (Not vertex) Then
          ' get the path, add a point to it and vertex edit that newly added point
          
          Set pPath = GetPart(pGeom, partIndex)
          Set pPointColn = pPath
          numVertices = pPointColn.PointCount
          
          ' reset the index pointer to the new index
          If vertexIndex = 0 Then
            pPointColn.AddPoint pPoint, 1
          Else
            pPointColn.AddPoint pPoint, , vertexIndex
          End If
          
          ' reset the index pointer to the new index
          TestGeometryHit tol, pPoint, pFeature, pHitPoint, hitDist, partIndex, vertexIndex, vertex
        End If
        Set m_pFeedback = New PolygonMovePointFeedback
        Set m_pFeedback.Display = pActiveView.ScreenDisplay
        Set pPolyMove = m_pFeedback
        pPolyMove.Start pGeom, vertexIndex, pPoint
      Else
        Exit Function
      End If
  End Select

  EditFeatureLineOrVertex = True

  GoTo CleanUp
ShowError:
  MsgBox "EditFeatureLineOrVertex " & Err.description
CleanUp:
  Set pGeom = Nothing
  Set pHitPoint = Nothing
  Set pFeature = Nothing
  Set pObjectClass = Nothing
  Set pPath = Nothing
  Set pPointColn = Nothing
  Set pLineMove = Nothing
  Set pPolyMove = Nothing
  Set pActiveView = Nothing
End Function


Public Function ConvertPixelsToMapUnits(pActiveView As IActiveView, pixelUnits As Double) As Double
On Error GoTo ShowError

  Dim realWorldDisplayExtent As Double
  Dim pixelExtent As Integer
  Dim sizeOfOnePixel As Double
  
  pixelExtent = pActiveView.ScreenDisplay.DisplayTransformation.DeviceFrame.Right - pActiveView.ScreenDisplay.DisplayTransformation.DeviceFrame.Left
  realWorldDisplayExtent = pActiveView.ScreenDisplay.DisplayTransformation.VisibleBounds.Width
  sizeOfOnePixel = realWorldDisplayExtent / pixelExtent
  ConvertPixelsToMapUnits = pixelUnits * sizeOfOnePixel

  Exit Function
ShowError:
  MsgBox "ConvertPixelsToMapUnits " & Err.description
End Function


Private Function EnsureOneSelected(pPoint As IPoint) As Boolean
On Error GoTo ShowError

  Dim pEnv As IEnvelope
  Dim searchRadius As Double

  If (gMap.SelectionCount > 0) Then
    EnsureOneSelected = True
    Exit Function
  End If
  searchRadius = ConvertPixelsToMapUnits(gMap, 16)
        
  Set pEnv = pPoint.Envelope
  pEnv.Height = searchRadius
  pEnv.Width = searchRadius
  pEnv.CenterAt pPoint
  gMap.SelectByShape pEnv, Nothing, True
  
  If (gMap.SelectionCount < 1) Then Exit Function
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
  EnsureOneSelected = True

  GoTo CleanUp
ShowError:
  MsgBox "EnsureOneSelected: " & Err.description
CleanUp:
  Set pEnv = Nothing
  Set pActiveView = Nothing
End Function


Private Sub GetClosestSelectedFeature(pPoint As IPoint, pFeature As IFeature)
On Error GoTo ShowError

  Dim testDist As Double
  Dim pProximity As IProximityOperator
  Dim pGeom As IGeometry
  Dim pTestFeature As IFeature
  Dim tempDist As Double
  Dim pSelected As IEnumFeature

  ' find the closest feature to the cursor position
  Set pProximity = pPoint

  Set pSelected = gMap.FeatureSelection
  pSelected.Reset

  ' get the first point to test the rest against
  Set pTestFeature = pSelected.Next
  Set pGeom = pTestFeature.ShapeCopy
  testDist = pProximity.ReturnDistance(pGeom)
  Set pFeature = pTestFeature
  Do
    Set pTestFeature = pSelected.Next
    If (Not pTestFeature Is Nothing) Then
      Set pGeom = pTestFeature.Shape
      tempDist = pProximity.ReturnDistance(pGeom)
      If (tempDist < testDist) Then
        testDist = tempDist
        Set pFeature = pTestFeature
      End If
    End If
  Loop While (Not pTestFeature Is Nothing)
   
  GoTo CleanUp
ShowError:
  MsgBox "GetClosestSelectedFeature " & Err.description
CleanUp:
  Set pProximity = Nothing
  Set pGeom = Nothing
  Set pTestFeature = Nothing
  Set pSelected = Nothing
End Sub

Private Sub HighlightSelectedNodes()
On Error GoTo ShowError

  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureSelection As IFeatureSelection
  Set pFeatureSelection = pFeatureLayer
  Dim pSelectionSet As ISelectionSet
  Set pSelectionSet = pFeatureSelection.SelectionSet
  Dim pFeatureCursor As IFeatureCursor
  pSelectionSet.Search Nothing, False, pFeatureCursor
  Dim pFeature As IFeature
  Set pFeature = pFeatureCursor.NextFeature
  
  'Get the graphics container of the map and clear it
  Dim pGraphicContainer As IGraphicsContainer
  Set pGraphicContainer = gMap
  pGraphicContainer.DeleteAllElements
  
  'Define marker symbol and color
  Dim pRGB As IRgbColor
  Set pRGB = New RgbColor
  pRGB.Red = 0
  pRGB.Green = 210
  pRGB.Blue = 0
  Dim pSymbol As ISimpleMarkerSymbol
  Set pSymbol = New SimpleMarkerSymbol
  pSymbol.Color = pRGB
  pSymbol.Size = 4
  pSymbol.Style = esriSMSSquare
  
  Dim pElement As IElement
  Dim pPolygon As IPolygon
  Dim pPoint As IPoint
  Dim pPointCollection As IPointCollection
  Dim pMarkerElement As IMarkerElement
  Dim iPt As Long

  Do While Not pFeature Is Nothing
    Set pPolygon = pFeature.Shape
    Set pPointCollection = pPolygon
    For iPt = 0 To pPointCollection.PointCount - 1
          Set pElement = New MarkerElement
          Set pPoint = pPointCollection.Point(iPt)
          pElement.Geometry = pPoint
          Set pMarkerElement = pElement
          pMarkerElement.Symbol = pSymbol
          pGraphicContainer.AddElement pElement, 0
    Next
    Set pFeature = pFeatureCursor.NextFeature
    Set pPolygon = Nothing
    Set pPoint = Nothing
    Set pElement = Nothing
    Set pMarkerElement = Nothing
  Loop
      
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
  
  GoTo CleanUp
  
ShowError:
    MsgBox "HighlightSelectedNodes: " & Err.description
CleanUp:
  Set pFeatureLayer = Nothing
  Set pFeatureSelection = Nothing
  Set pSelectionSet = Nothing
  Set pFeatureCursor = Nothing
  Set pFeature = Nothing
  Set pGraphicContainer = Nothing
  Set pRGB = Nothing
  Set pSymbol = Nothing
  Set pElement = Nothing
  Set pPolygon = Nothing
  Set pPoint = Nothing
  Set pPointCollection = Nothing
  Set pMarkerElement = Nothing
  Set pActiveView = Nothing
End Sub

Private Function TestGeometryHit(tolerance As Double, _
                                 pPoint As IPoint, _
                                 ByVal pFeature As IFeature, _
                                 pHitPoint As IPoint, _
                                 hitDist As Double, _
                                 partIndex As Long, _
                                 vertexIndex As Long, _
                                 vertexHit As Boolean) As Boolean
On Error GoTo ShowError

  Dim pGeom As IGeometry
  Dim pPolyline As IPolyline
  Dim pPolygon As IPolygon
  Dim pHitTest As IHitTest
  Dim pCurve As ICurve
  Dim pPolyCurve As IPolycurve
      
  ' Function returns TRUE is a feature's shape is hit and further
  ' defines if the a vertex lies within the tolorance
  Set pGeom = pFeature.Shape
 
  ' don't bother testing points since we aren't interested in them
  If ((pGeom.GeometryType = esriGeometryPoint) Or _
      (pGeom.GeometryType = esriGeometryMultipoint)) Then Exit Function
    
  Set pHitTest = pGeom
  Set pHitPoint = New Point
  If (pHitTest.HitTest(pPoint, _
                       tolerance, _
                       esriGeometryPartVertex, _
                       pHitPoint, _
                       hitDist, _
                       partIndex, _
                       vertexIndex, _
                       True) = True) Then
    TestGeometryHit = True
    vertexHit = True
  Else
    If (pHitTest.HitTest(pPoint, _
                         tolerance, _
                         esriGeometryPartBoundary, _
                         pHitPoint, _
                         hitDist, _
                         partIndex, _
                         vertexIndex, _
                         True) = True) Then
      TestGeometryHit = True
      vertexHit = False
    End If
  End If


  GoTo CleanUp
ShowError:
  MsgBox "TestGeometryHit: " & Err.description
CleanUp:
  Set pGeom = Nothing
  Set pPolyline = Nothing
  Set pPolygon = Nothing
  Set pHitTest = Nothing
  Set pCurve = Nothing
  Set pPolyCurve = Nothing
End Function


Private Function GetPart(pGeom As IGeometry, partIndex As Long) As IGeometry
On Error GoTo ShowError

  Dim pGeomColn As IGeometryCollection
  
  Set pGeomColn = pGeom
  Set GetPart = pGeomColn.Geometry(partIndex)

  GoTo CleanUp
ShowError:
  MsgBox "GetPart: " & Err.description
CleanUp:
  Set pGeomColn = Nothing
End Function


Private Sub EditFeatures(pPoint As IPoint)
On Error GoTo ShowError

  Dim deltaX As Double
  Dim deltaY As Double
  Dim pSelected As IEnumFeature
  Dim pTrans As ITransform2D
  Dim pGeometry As IGeometry
  Dim pLineMove As ILineMovePointFeedback
  Dim pPolyMove As IPolygonMovePointFeedback
  Dim needsRefresh As Boolean
  Dim pFeature As IFeature
  
  ' If no feedback no edit
  If (m_pFeedback Is Nothing) Then Exit Sub
  
  ' Now determine what type of edit operation has taken place
  ' do this by testing what type of feedback interface we have
  If (TypeOf m_pFeedback Is IMoveGeometryFeedback) Then
    ' find out the amount of movement and then alter the geometry accordingly
    deltaX = pPoint.X - m_pAnchorPoint.X
    deltaY = pPoint.Y - m_pAnchorPoint.Y
    
    Set pSelected = gMap.FeatureSelection
    pSelected.Reset
    Set pFeature = pSelected.Next
    Do While (Not pFeature Is Nothing)
      Set pTrans = pFeature.ShapeCopy
      pTrans.Move deltaX, deltaY
      UpdateFeature pFeature, pTrans, needsRefresh
      
      Set pFeature = pSelected.Next
    Loop
  ElseIf (TypeOf m_pFeedback Is IPolygonMovePointFeedback) Then
    Set pPolyMove = m_pFeedback
      Set pGeometry = pPolyMove.Stop
      UpdateFeature m_pEditFeature, pGeometry, needsRefresh
  ElseIf (TypeOf m_pFeedback Is ILineMovePointFeedback) Then
    Set pLineMove = m_pFeedback
    Set pGeometry = pLineMove.Stop
    UpdateFeature m_pEditFeature, pGeometry, needsRefresh
  End If
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  pActiveView.Refresh

  GoTo CleanUp
ShowError:
  MsgBox "EditFeatures: " & Err.description
CleanUp:
  Set pSelected = Nothing
  Set pTrans = Nothing
  Set pGeometry = Nothing
  Set pLineMove = Nothing
  Set pPolyMove = Nothing
  Set pFeature = Nothing
  Set pActiveView = Nothing
End Sub


Private Function UpdateFeature(pFeature As IFeature, pGeometry As IGeometry, needsRefresh As Boolean) As Boolean
On Error GoTo ShowError

  Dim pDataset As IDataset
  Dim pWorkspaceEdit As IWorkspaceEdit
  Dim giveWarning As Boolean
  
  ' make sure we are actually editing this layer.   If not give a warning
  Set pDataset = pFeature.Class
  If ((pDataset.Type = esriDTFeatureClass) Or _
      (pDataset.Type = esriDTFeatureDataset)) Then
    Set pWorkspaceEdit = pDataset.Workspace
    If (Not pWorkspaceEdit.IsBeingEdited) Then giveWarning = True
  Else
    giveWarning = True
  End If
  
  If (giveWarning) Then
    Beep
    MsgBox "This feature is in a layer not in edit mode." & vbCrLf & "Edit cannot be made.   Start edit and try again.", vbExclamation + vbOKOnly, "Warning"
    Exit Function
  End If

  pWorkspaceEdit.StartEditOperation
  Set pFeature.Shape = pGeometry
  pFeature.Store
  pWorkspaceEdit.StopEditOperation
  
  If (needsRefresh = False) Then needsRefresh = True
  UpdateFeature = True

  GoTo CleanUp
ShowError:
  MsgBox "UpdateFeature " & Err.description
CleanUp:
  Set pDataset = Nothing
  Set pWorkspaceEdit = Nothing
End Function

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub



