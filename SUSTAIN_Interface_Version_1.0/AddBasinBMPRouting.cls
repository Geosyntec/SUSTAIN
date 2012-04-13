VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddBasinBMPRouting"
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
Private m_bInUse As Boolean
Private m_pPoint As IPoint
Private m_pFeedback As IDisplayFeedback

Private pOutletType1 As Integer
Private pOutletType2 As Integer
Private FromPointID As Integer
Private ToPointID As Integer

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("BASINROUTE", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_pEditor = Nothing
    Set m_pEditSketch = Nothing
    Set m_Bitmap = Nothing
    Set m_pPoint = Nothing
    Set m_pFeedback = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Drainage Area to BMP Connection"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableDrainageAreaBMPConnection
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Drainage Area to BMP Connection"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Define Drainage Area to BMP Connection"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Dim pBasinRoutingFLayer As IFeatureLayer
    Set pBasinRoutingFLayer = GetInputFeatureLayer("BasinRouting")
    Dim boolChange
    boolChange = vbYes
    If Not (pBasinRoutingFLayer Is Nothing) Then
        boolChange = MsgBox("Do you want to reset existing drainage area to BMP connection ?", vbYesNo, "Drainage Area to BMP Connection")
    End If
    If (boolChange = vbYes) Then
        'Call subroutine to define default subbasin to bmp routing
        Call CreateSubBasinToBMPRouting
        'Call subroutine to define default subbasin to vfs routing
        Call CreateSubBasinToVFSRouting
        '*** Call the subroutine to render the route to both bmp and vfs
        Call RenderBasintoBMPRouting
    End If

End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Drainage Area to BMP Connection"
End Property


Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("DIGITIZE", vbResCursor)
End Property

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

  Dim pPoint As IPoint
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Dim pLineFeed As INewLineFeedback
  Dim pLayer As ILayer

  If (Not m_bInUse) Then
        m_bInUse = True
        Set m_pFeedback = New NewLineFeedback
        Set pLineFeed = m_pFeedback
        If (Not m_pFeedback Is Nothing) Then
            
            Dim pSLnSym As ISimpleLineSymbol
            Set pSLnSym = m_pFeedback.Symbol
            ' Create a new RGBColor and set it up
            Dim pRGB As IColor
            Set pRGB = New RgbColor
            pRGB.RGB = RGB(0, 112, 255)   'light blue
            
            ' Set the Color and Style for the Feedback's Symbol
            pSLnSym.Color = pRGB
            pSLnSym.Style = esriSLSDash

            Set m_pFeedback.Display = pActiveView.ScreenDisplay
            pLineFeed.Start pPoint
            
        End If
  Else
    
      Set pLineFeed = m_pFeedback
            
      'Add the shifted point
      pLineFeed.AddPoint pPoint
      m_bInUse = False
      'Got a line with two points, stop it get the line
      Dim pPolyline As IPolyline
      Set pPolyline = pLineFeed.Stop
      Dim pGraphicsContainer As IGraphicsContainer
      Set pGraphicsContainer = gMap
      
      Dim pElement As IElement
      Set pElement = New LineElement
      pElement.Geometry = pPolyline
      Dim pLineElement As ILineElement
      Set pLineElement = pElement
      Dim pLineSymbol As ILineSymbol
      Set pLineSymbol = ReturnBasintoBMPRouteSymbol
      pLineElement.Symbol = pLineSymbol
      pGraphicsContainer.AddElement pElement, 0
      gMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
    
      '*** Call the subroutine to check for route validity and save it
      Call CheckBasinBMPValidityAndAddBasinRoute(pPolyline.FromPoint, pPolyline.ToPoint)
      
      pGraphicsContainer.DeleteAllElements
      gMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
  End If
 
   
  'Exit Sub ' Sabu Paul, Jan 17, 2005
  GoTo CleanUp  ' Sabu Paul, Jan 17, 2005
ErrorHandler:
  MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp:  ' Sabu Paul, Jan 17, 2005 -- Cleanup section was added
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pPolyline = Nothing
    Set pGraphicsContainer = Nothing
    Set pElement = Nothing
    Set pLineElement = Nothing
    Set pLineSymbol = Nothing
    Set pLineFeed = Nothing
    Set pSLnSym = Nothing
    Set pRGB = Nothing
End Sub


Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

  If (Not m_bInUse) Then Exit Sub
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  m_pFeedback.MoveTo pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Set m_pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

  GoTo CleanUp
ErrorHandler:
  MsgBox "ITool_OnMouseMove: " & Err.description
CleanUp:
  Set pActiveView = Nothing
End Sub

Private Sub ITool_OnDblClick()

End Sub

Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean

End Function

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub

'Subroutine to get the bmp route type total, weir, outlet or underdrain
Public Function CheckBasinBMPValidityAndAddBasinRoute(ByRef pStartPoint As IPoint, ByRef pEndPoint As IPoint) As Integer
On Error GoTo ShowError
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If (pWatershedFLayer Is Nothing) Then
        MsgBox "Watershed feature layer not found."
        Exit Function
    End If
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    If (pBMPFLayer Is Nothing) Then
        MsgBox "BMPs feature layer not found."
        Exit Function
    End If
    
    'Get the bmp feature class, find if the end point intersects any bmp
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    'Expand the envelope of the user input point 5 times
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pEndPoint.Envelope
    ExpandPointEnvelope pEnvelope
    'Search for BMPs in the expanded envelope
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    'Run the spatial search query
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pBMPFClass.Search(pSpatialFilter, False)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Dim iBmpIdFld As Long
    iBmpIdFld = pFeatureCursor.FindField("ID")
    Dim pBMPID As Integer
    pBMPID = -1
    If Not pFeature Is Nothing Then
        pBMPID = pFeature.value(iBmpIdFld)
        Set pEndPoint = pFeature.Shape
    End If
        
        
    'Get the VFS feature class, find if the end point intersects any vfs
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSID As Integer
    pVFSID = -1
    If Not (pVFSFLayer Is Nothing) Then
        Dim pVFSFClass As IFeatureClass
        Set pVFSFClass = pVFSFLayer.FeatureClass
        'Expand the envelope of the user input point 5 times
        Set pEnvelope = pEndPoint.Envelope
        ExpandPointEnvelope pEnvelope
        'Search for VFSs in the expanded envelope
        Set pSpatialFilter = New SpatialFilter
        Set pSpatialFilter.Geometry = pEnvelope
        pSpatialFilter.SpatialRel = esriSpatialRelIntersects
        'Run the spatial search query
        Set pFeatureCursor = pVFSFClass.Search(pSpatialFilter, False)
        Set pFeature = pFeatureCursor.NextFeature
        Dim iVfsIdFld As Long
        iVfsIdFld = pFeatureCursor.FindField("ID")
        Dim pPolyline As IPolyline
        If Not pFeature Is Nothing Then
            pVFSID = pFeature.value(iVfsIdFld)
            Set pPolyline = pFeature.Shape
            pPolyline.QueryPoint esriNoExtension, (pPolyline.Length / 2), False, pEndPoint
        End If
    End If
    
    'Get the watershed feature class, find if the start point lies within a watershed
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    'Search for watershed containing the start point
    Set pSpatialFilter.Geometry = pStartPoint
    pSpatialFilter.SpatialRel = esriSpatialRelWithin
    'Run the spatial search query
    Set pFeatureCursor = Nothing
    Set pFeatureCursor = pWatershedFClass.Search(pSpatialFilter, False)
    Set pFeature = Nothing
    Set pFeature = pFeatureCursor.NextFeature
    Dim iBasinIDFld As Long
    iBasinIDFld = pFeatureCursor.FindField("ID")
    Dim iBasinBMPIDFld As Long
    iBasinBMPIDFld = pFeatureCursor.FindField("BMPID")
    Dim pBasinID As Integer
    pBasinID = -1
    'If watershed containing start point found, update bmp flowing into and save it
    If Not pFeature Is Nothing Then
        Dim boolChange
        boolChange = MsgBox("Do you want to save this drainage area to BMP connection ?", vbYesNo, "Drainage Area to BMP Connection")
        If (boolChange = vbYes) Then
            'Save the watershed feature's draining BMP ID or VFS ID
            pBasinID = pFeature.value(iBasinIDFld)
            If (pBMPID > 0) Then
                pFeature.value(iBasinBMPIDFld) = pBMPID
            ElseIf (pVFSID > 0) Then
                pFeature.value(iBasinBMPIDFld) = pVFSID
            End If
            pFeature.Store
        End If
    End If
    
    'If BMP and Basin is found
    If (pBasinID > 0 And (pBMPID > 0 Or pVFSID > 0)) Then
        'Define a feature class to store the routing information for basin to BMP
        Dim pBasinToBMPRoutingLayer As IFeatureLayer
        Set pBasinToBMPRoutingLayer = GetInputFeatureLayer("BasinRouting")
        Dim pBasinToBMPRoutingClass As IFeatureClass
        If (pBasinToBMPRoutingLayer Is Nothing) Then
            Set pBasinToBMPRoutingClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, "basinroute")
            Set pBasinToBMPRoutingLayer = New FeatureLayer
            Set pBasinToBMPRoutingLayer.FeatureClass = pBasinToBMPRoutingClass
            AddLayerToMap pBasinToBMPRoutingLayer, "BasinRouting"
        End If
        
        'Create a new line feature class
        Set pBasinToBMPRoutingClass = pBasinToBMPRoutingLayer.FeatureClass
        Dim pBasinToBMPFeature As IFeature
        Dim iIDFld As Long
        iIDFld = pBasinToBMPRoutingClass.FindField("ID")
        Dim pPointCollection As IPointCollection
        Set pPointCollection = New Polyline
        'Add start and end points to a line
        pPointCollection.AddPoint pStartPoint
        pPointCollection.AddPoint pEndPoint
    
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "ID = " & pBasinID
        Dim pBasinToBMPFCursor As IFeatureCursor
        Set pBasinToBMPFCursor = pBasinToBMPRoutingClass.Search(pQueryFilter, True)
        Set pBasinToBMPFeature = pBasinToBMPFCursor.NextFeature
        If (pBasinToBMPFeature Is Nothing) Then
            'Create a new basin to bmp route and save it
            Set pBasinToBMPFeature = pBasinToBMPRoutingClass.CreateFeature
        End If
        'Update shape and save it
        Set pBasinToBMPFeature.Shape = pPointCollection
        pBasinToBMPFeature.value(iIDFld) = pBasinID
        pBasinToBMPFeature.Store
    
        ' create a new simple line renderer
        Dim pRen As ISimpleRenderer
        Dim pGeoFeatLyr As IGeoFeatureLayer
        Set pGeoFeatLyr = pBasinToBMPRoutingLayer
        Set pRen = pGeoFeatLyr.Renderer
        Set pRen.Symbol = ReturnBasintoBMPRouteSymbol
        gMxDoc.ActiveView.Refresh
        gMxDoc.UpdateContents
    End If
    
    'Cleanup variables
    GoTo CleanUp
ShowError:
    MsgBox "CheckBasinBMPValidityAndAddBasinRoute: " & Err.description
CleanUp:
        
    Set pWatershedFLayer = Nothing
    Set pBMPFLayer = Nothing
    Set pBMPFClass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pWatershedFClass = Nothing
    Set pBasinToBMPRoutingLayer = Nothing
    Set pBasinToBMPRoutingClass = Nothing
    Set pBasinToBMPFeature = Nothing
    Set pPointCollection = Nothing
    Set pQueryFilter = Nothing
    Set pBasinToBMPFCursor = Nothing
    Set pRen = Nothing
    Set pGeoFeatLyr = Nothing
    Set pPolyline = Nothing
End Function



