VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoBMPNetwork"
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
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("AUTOROUTE", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Automatic Routing Network"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableAutoBMPNetworkRouting
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Automatic Routing Network"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Define Automatic Routing Network"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Automatic Routing Network"
End Property


Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("DIGITIZE", vbResCursor)
End Property

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Dim pPoint As IPoint
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

  '* call the subroutine to automatically route the layer
  DefineAutomaticNetworkForBMP pPoint

  GoTo CleanUp
ErrorHandler:
  MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pPoint = Nothing
End Sub

Private Function CreateLineElement(pGeometry As IPolyline, pSymbolType, pRGBColor As IRgbColor, pSize As Integer) As ILineElement

    Dim pLElement As ILineElement
    Dim pElement As IElement
    Set pLElement = New LineElement
    Set pElement = pLElement
    Dim pLineSymbol As ILineSymbol
    Set pLineSymbol = New SimpleLineSymbol
    pLineSymbol.Color = pRGBColor
    pLineSymbol.Width = pSize
    pLElement.Symbol = pLineSymbol
    pElement.Geometry = pGeometry
  Set CreateLineElement = pElement

End Function


Private Function CreateMarkerElement(pGeometry As IPoint, pRGBColor As IRgbColor, pSize As Integer, pSymbolType As Integer) As IMarkerElement
  
    'Create the Simple Marker
    Dim pSmplMarker As ISimpleMarkerSymbol
    Set pSmplMarker = New SimpleMarkerSymbol
    pSmplMarker.Style = pSymbolType
    pSmplMarker.Size = pSize
    pSmplMarker.Color = pRGBColor
    
    Dim pElement As IElement
    Dim pMElement As IMarkerElement
    
    Set pMElement = New MarkerElement
    Set pElement = pMElement
    pMElement.Symbol = pSmplMarker
    pElement.Geometry = pGeometry
    
  Set CreateMarkerElement = pElement
  
End Function



Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

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


'******************************************************************************
'Subroutine: DefineAutomaticNetworkForBMP
'Author:     Mira Chokshi
'Purpose:    Get the point clicked by the user, intersect with a bmp and for
'            that bmp, find the cost path. Find next point on that cost path
'******************************************************************************
Public Sub DefineAutomaticNetworkForBMP(pSearchPoint As IPoint)
On Error GoTo ShowError

    '** get VFS feature layer
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    Dim pVFSCount As Integer
    pVFSCount = 0
    If Not (pVFSFLayer Is Nothing) Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
        pVFSCount = pVFSFClass.FeatureCount(Nothing)    ' Get vfs feature count
    End If
    
    '** get BMP feature layer
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    '** get feature class and cursor for bmp
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
              
    Dim pBMPCount As Integer
    pBMPCount = pBMPFClass.FeatureCount(Nothing)
    
    '** define field names
    Dim iID As Long
    iID = pBMPFClass.FindField("ID")
    Dim iLabel As Long
    iLabel = pBMPFClass.FindField("LABEL")
    
    '** define from point variables
    Dim pFromBMPPoint As IPoint
    Dim pFromBMPID As Integer
    Dim pFromBMPLabel As String
    
    '** intersect the input point with bmps feature layer
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pSearchPoint.Envelope
    ExpandPointEnvelope pEnvelope   'Expand the search polygon
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    
    '** apply the spatial filter to find the bmp feature
    Dim pBMPFeatureCursor As IFeatureCursor
    Set pBMPFeatureCursor = pBMPFClass.Search(pSpatialFilter, True)
    Dim pBMPFeature As IFeature
    Set pBMPFeature = pBMPFeatureCursor.NextFeature
    If (Not pBMPFeature Is Nothing) Then
        'Get the feature id and label
        Set pFromBMPPoint = pBMPFeature.Shape
        pFromBMPID = pBMPFeature.value(iID)
        pFromBMPLabel = pBMPFeature.value(iLabel)
    End If
    
    '** if BMP is not found, then may be its a VFS
    Dim pVFSLayer As IFeatureLayer
    Set pVFSLayer = GetInputFeatureLayer("VFS")
    Dim pVFSClass As IFeatureClass
    If (Not pVFSFLayer Is Nothing) Then
        Set pVFSClass = pVFSLayer.FeatureClass
    End If
    Dim pVFSPolyline As IPolyline
    Dim iVFSID As Long
    iVFSID = pVFSClass.FindField("ID")
    Dim iVFSLabel As Long
    iVFSLabel = pVFSClass.FindField("LABEL")
    '** query VFS layer
    If (Not pVFSClass Is Nothing) Then
        Set pBMPFeatureCursor = pVFSClass.Search(pSpatialFilter, True)
        Set pBMPFeature = pBMPFeatureCursor.NextFeature
        If (Not pBMPFeature Is Nothing) Then
            Set pVFSPolyline = pBMPFeature.Shape
            'Get the feature id and label
            Set pFromBMPPoint = pVFSPolyline.ToPoint
            pFromBMPID = pBMPFeature.value(iVFSID)
            pFromBMPLabel = pBMPFeature.value(iVFSLabel)
        End If
    End If
        
    '** get the map point from the feature
    Dim pRoutePath As IPolyline
    If Not (pFromBMPPoint Is Nothing) Then
        Set pRoutePath = DefineCostPathForBMP(pFromBMPPoint)
        If (pRoutePath Is Nothing) Then
            Exit Sub
        End If
    
        'Check if the bmp has an outlet, get outlet value
        Dim pOUTLETType As Integer
        pOUTLETType = GetBMPRouteType(pFromBMPPoint, pFromBMPID, "BMPs", True)
        If (pOUTLETType > 1) Then     'If its Weir, Orifice,Underdrain
             'If it returns 2, disable the underdrain option
             FrmOutlet.UnderDrain.Enabled = True
             If (pOUTLETType = 2) Then
                 FrmOutlet.UnderDrain.Enabled = False
             End If
             FrmOutlet.Show vbModal
             pOUTLETType = gBMPOutletType
        End If
    End If
         
    Dim pPoint As IPoint
    Set pPoint = New Point
    Dim iLen As Integer
    Dim pTotalSegs As Integer
    Dim pDistanceAlongCurve As Double
    
    Dim pPointCollection As IPointCollection
    Set pPointCollection = New Polyline
    
    Dim pToBMPPoint As IPoint
    Dim pToBMPID As Integer
    Dim pToBMPLabel As String
    pToBMPID = -1
    pToBMPLabel = ""
    

    If Not (pRoutePath Is Nothing) Then
        pPointCollection.AddPoint pFromBMPPoint
        pTotalSegs = pRoutePath.Length / gCellSize
        For iLen = 1 To pTotalSegs
            pDistanceAlongCurve = iLen * gCellSize
            pRoutePath.QueryPoint esriNoExtension, pDistanceAlongCurve, False, pPoint
            
            'add that point to the point colleciton
            pPointCollection.AddPoint pPoint
            
            'Increase the buffer around the point to search
            Dim pBufferPoly As IPolygon
            Dim pTopologicalOp As ITopologicalOperator
            Set pTopologicalOp = pPoint
            Set pBufferPoly = pTopologicalOp.Buffer(2 * gCellSize)

            Set pSpatialFilter = New SpatialFilter
            Set pSpatialFilter.Geometry = pBufferPoly
            pSpatialFilter.SpatialRel = esriSpatialRelIntersects
            Set pBMPFeatureCursor = pBMPFClass.Search(pSpatialFilter, True)
            Set pBMPFeature = pBMPFeatureCursor.NextFeature

            'Check if the point has a bmp close by, if yes exit the for
            If Not (pBMPFeature Is Nothing) Then
                Set pToBMPPoint = pBMPFeature.Shape
                pToBMPID = pBMPFeature.value(iID)
                pToBMPLabel = pBMPFeature.value(iLabel)
                If (pToBMPID <> pFromBMPID) Then
                    pPointCollection.AddPoint pToBMPPoint
                    iLen = pRoutePath.Length + 100  'to exit the for
                    Exit For
                End If
            End If

        Next
                
        If (pToBMPID <> -1 And pToBMPID <> pFromBMPID) Then
        
            'Update the bmp network table and bmp feature class table
            UpdateBMPNetworkTableRoute pFromBMPID, pToBMPID, pOUTLETType, "Conduits", "BMPs"
            UpdateBMPFeatureClassInformation pFromBMPID, pToBMPID
                        
            'Create a line from the point collection, display it
            Dim pPolyline As IPolyline
            Set pPolyline = pPointCollection
            DisplayRoutePathOnMap pPolyline
        
            'Ask the user, if they think this route is okay
            Dim bContinue
            bContinue = MsgBox("Do you want to save this route between " & pFromBMPLabel & _
                                " and " & pToBMPLabel & " ?", vbYesNo)
                                
            If (bContinue = vbYes) Then
                '** Save the conduit feature and open property box
                OpenConduitPropertyDialogBox pPolyline, pFromBMPID, pFromBMPLabel, _
                                pToBMPID, pToBMPLabel, pOUTLETType, pBMPCount, pVFSCount
            
            End If
        End If
    
        '** Erase the graphics from the map
        ClearRoutePathOnMap
    
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "DefineAutomaticNetworkForBMP: " & Err.description & " " & Err.Number
CleanUp:

End Sub



Public Sub OpenConduitPropertyDialogBox(pPolyline As IPolyline, pID As Integer, pIDLabel As String, pDSID As Integer, pDSIDLabel As String, pOUTLETType As Integer, pBMPFeatureCount As Integer, pVFSCount As Integer)
On Error GoTo ShowError

    
    'Delete Conduits feature layer if present
    Dim pConduitNetworkLayer As IFeatureLayer
    Set pConduitNetworkLayer = GetInputFeatureLayer("Conduits")
    
    'define the line feature class
    Dim pConduitNetworkFClass As IFeatureClass
    
    'if conduit network layer is not found, create it
    If (pConduitNetworkLayer Is Nothing) Then
        Set pConduitNetworkFClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, "conduit")
        'Add a new layer to map
        Set pConduitNetworkLayer = New FeatureLayer
        Set pConduitNetworkLayer.FeatureClass = pConduitNetworkFClass
        AddLayerToMap pConduitNetworkLayer, "Conduits"
    End If
    
    'Get the feature class
    Set pConduitNetworkFClass = pConduitNetworkLayer.FeatureClass
    
    Dim pConduitFeatureCount As Integer
    pConduitFeatureCount = pConduitNetworkFClass.FeatureCount(Nothing)
    
   '** Check if a conduit already exists, if so, get that id instead of new id
   Dim pQueryFilter As IQueryFilter
   Set pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "CFROM = " & pID & " AND OUTLETTYPE = " & pOUTLETType
   Dim pConduitFeatureCursor As IFeatureCursor
   Set pConduitFeatureCursor = pConduitNetworkFClass.Search(pQueryFilter, False)
   Dim pConduitFeature As IFeature
   Set pConduitFeature = pConduitFeatureCursor.NextFeature
   If (pConduitFeature Is Nothing) Then
       gConduitIDValue = pBMPFeatureCount + pConduitFeatureCount + pVFSCount + 1
       Set pConduitFeature = pConduitNetworkFClass.CreateFeature
   Else
      'Conduit ID value overwritten
      gConduitIDValue = pConduitFeature.value(pConduitNetworkFClass.FindField("ID"))
   End If
   Set pConduitFeature.Shape = pPolyline
   'Update values in the conduits feature layer
   pConduitFeature.value(pConduitNetworkFClass.FindField("ID")) = gConduitIDValue
   pConduitFeature.value(pConduitNetworkFClass.FindField("CFROM")) = pID
   pConduitFeature.value(pConduitNetworkFClass.FindField("CTO")) = pDSID
   pConduitFeature.value(pConduitNetworkFClass.FindField("OUTLETTYPE")) = pOUTLETType
   pConduitFeature.value(pConduitNetworkFClass.FindField("TYPEDESC")) = GetRouteDesc(pOUTLETType)
   pConduitFeature.value(pConduitNetworkFClass.FindField("LABEL")) = pIDLabel & "-" & pDSIDLabel
   pConduitFeature.Store
   
   'Refresh the screen, and render the route layer
   RenderSchematicRouteLayer pConduitNetworkLayer
   gMxDoc.ActiveView.Refresh


   '** Get start and end elevation of the ppolyline
   Dim pStartElevation As Double
   Dim pStopElevation As Double
   pStartElevation = GetElevationForPoint(pPolyline.FromPoint)
   pStopElevation = GetElevationForPoint(pPolyline.ToPoint)
   
   '** update the form
   Dim pConduitLength As Double
   pConduitLength = Format(pPolyline.Length * gMetersPerUnit * 3.28, "#.##")

   '**** Mira Chokshi, Feb 4 2005, To input conduit cross-section
   FrmConduitCSection.txtSlopeEntrance.Text = pStartElevation
   FrmConduitCSection.txtSlopeExit.Text = pStopElevation
   FrmConduitCSection.txtLength.Text = CStr(pConduitLength)
   FrmConduitCSection.Show vbModal
    
   GoTo CleanUp
   
ShowError:
    MsgBox "OpenConduitPropertyDialogBox: " & Err.description
CleanUp:

End Sub

'******************************************************************************
'Subroutine: DefineCostPathForBMP
'Author:     Mira Chokshi
'Purpose:    trace the cost path of a bmp point
'******************************************************************************
Public Function DefineCostPathForBMP(pMapPoint As IPoint) As IPolyline
On Error GoTo ShowError

        '** Initialize Raster Operator
        InitializeOperators
        
        Dim pRasterFlowDir As IRaster
        Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
        
        Dim pRasterFlowAccu As IRaster
        Set pRasterFlowAccu = OpenRasterDatasetFromDisk("FlowAccu")
        
        If (pRasterFlowDir Is Nothing Or pRasterFlowAccu Is Nothing) Then
            MsgBox "Flow Direction & Flow Accumulation required for Automatic BMP Routing."
            Exit Function
        End If
        
        Dim pPointCol As IPointCollection
        Set pPointCol = New Multipoint
        pPointCol.AddPoint pMapPoint

        'Obain rain drop path
        Dim pGeomCol As IGeometryCollection
        Set pGeomCol = gRasterDistanceOp.CostPathAsPolyline(pPointCol, pRasterFlowAccu, pRasterFlowDir)
        
        'get the polyline from the geometry collection
        Dim pPolyline As IPolyline
        Set pPolyline = pGeomCol.Geometry(0)
        
        'return the polyline back to calling function
        Set DefineCostPathForBMP = pPolyline
        GoTo CleanUp
        
ShowError:
    MsgBox "DefineCostPathForBMP: " & Err.description
CleanUp:
        Set pRasterFlowAccu = Nothing
        Set pRasterFlowDir = Nothing
        Set pGeomCol = Nothing
        Set pPolyline = Nothing
End Function


'******************************************************************************
'Subroutine: ClearRoutePathOnMap
'Author:     Mira Chokshi
'Purpose:    display the green route on the map
'******************************************************************************
Public Sub ClearRoutePathOnMap()

      '** flash the line for now
      Dim pActiveView As IActiveView
      Set pActiveView = gMxDoc.ActiveView
      
      '** add a Marker to the map
      Dim pGraphCon As IGraphicsContainer
      Set pGraphCon = pActiveView.GraphicsContainer
      pGraphCon.DeleteAllElements
        
      '** refresh the screen display
      Dim pScreenDisp As IScreenDisplay
      Set pScreenDisp = pActiveView.ScreenDisplay
      pScreenDisp.Invalidate pActiveView.Extent.Envelope, True, 0
      '** refresh the active view
      pActiveView.Refresh
      
      '** clean up
      Set pScreenDisp = Nothing
      Set pGraphCon = Nothing
      Set pActiveView = Nothing
      
End Sub


'******************************************************************************
'Subroutine: GetElevationForPoint
'Author:     Mira Chokshi
'Purpose:    get the elevation value for point
'******************************************************************************
Public Function GetElevationForPoint(pPoint As IPoint) As Double
On Error GoTo ShowError

    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = GetInputRasterLayer("DEM")
    '** exit if DEM raster layer is not found.
    If (pRasterLayer Is Nothing) Then
        Exit Function
    End If
        
    '** use the identify object to get the raster
    Dim pIdentify As IIdentify
    Set pIdentify = pRasterLayer
    
    '** get RasterIdentifyObject on that point
    Dim pIDArray As IArray
    Set pIDArray = pIdentify.Identify(pPoint)
         
         
    Dim pRIDObj As IRasterIdentifyObj
    If Not pIDArray Is Nothing Then
        Set pRIDObj = pIDArray.Element(0)
        GetElevationForPoint = pRIDObj.name
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "GetElevationForPoint: " & Err.description
CleanUp:
    Set pRasterLayer = Nothing
    Set pIdentify = Nothing
    Set pIDArray = Nothing
    Set pRIDObj = Nothing
End Function
   

'******************************************************************************
'Subroutine: DisplayRoutePathOnMap
'Author:     Mira Chokshi
'Purpose:    display the green route on the map
'******************************************************************************
Public Sub DisplayRoutePathOnMap(pRoutePath As IPolyline)
On Error GoTo ShowError

      '** flash the line for now
      Dim pActiveView As IActiveView
      Set pActiveView = gMxDoc.ActiveView
      
      '** add a Marker to the map
      Dim pGraphCon As IGraphicsContainer
      Set pGraphCon = pActiveView.GraphicsContainer
      pGraphCon.DeleteAllElements
        
      Dim pScreenDisp As IScreenDisplay
      Set pScreenDisp = pActiveView.ScreenDisplay
      pScreenDisp.Invalidate pActiveView.Extent.Envelope, True, 0
        
      '** define the color
      Dim pColor1 As IRgbColor
      Set pColor1 = New RgbColor
      pColor1.Blue = 70
      pColor1.Red = 20
      pColor1.Green = 255
        
      '** define line symbol
      Dim pSymL As ISimpleLineSymbol
      Set pSymL = New SimpleLineSymbol
      With pSymL
           .Color = pColor1
           .Style = esriSLSSolid
           .Width = 2
      End With
        
     '** create graphics and add it into map
      Dim pLineElem As IElement
      Set pLineElem = CreateLineElement(pRoutePath, esriSLSSolid, pColor1, 2)
      pGraphCon.AddElement pLineElem, 0
      With pScreenDisp
        .ActiveCache = 0
        .StartDrawing pScreenDisp.hDC, 0
        .SetSymbol pSymL
        pLineElem.Draw pScreenDisp, Nothing
        .FinishDrawing
      End With
        
      '** cleanup
      GoTo CleanUp
        
ShowError:
    MsgBox "DisplayRoutePathOnMap: " & Err.description
CleanUp:
    
End Sub

