Attribute VB_Name = "ModuleRouting"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleRouting
'   Purpose:     This module contains the main function to generate subwatersheds
'                for BMPs selected by user. Also contains utility functions to
'                define network between bmps & create a line feature layer
'                representing the network.
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 08/19/2004 - mira chokshi
'
'******************************************************************************

Option Explicit
Option Base 0

'******************************************************************************
'Subroutine: GenerateSubWatersheds
'Author:     Mira Chokshi
'Purpose:    This Function is entry point for DELINEATE WATERSHEDS command.
'            It calls the module to check flow direction & accumulation rasters
'            It Asks the user for Snapping Distance and creates another
'            feature layer called SnapPoints that represents points moved by
'            snapping to flow accumulation. It then uses SnapPoints to determine
'            Routine between BMPs and creates the BMPNetwork table. Finally it
'            open the form to allow users to modify downstream BMPs for splitter.
'******************************************************************************
Public Sub GenerateSubWatersheds()

On Error GoTo ShowError
    
    'Call the module to fill dem, generate flow direction & flow accumulation
    ModuleFlowDirAccu.RunSTREAMAgreeDEMForFlowDirAndAccu

    'Delete subwatershed layer if present
    Dim pSubWaterRLayer As IRasterLayer
    Set pSubWaterRLayer = GetInputRasterLayer("SubWatershed")
    If Not (pSubWaterRLayer Is Nothing) Then
        DeleteLayerFromMap ("SubWatershed")
    End If
    'Delete snap points layer if present
    DeleteLayerFromMap "SnapPoints"
    
    'Get BMPs feature layer
    Dim pPourPointsFLayer As IFeatureLayer
    Set pPourPointsFLayer = GetInputFeatureLayer("BMPs")
    If (Not pPourPointsFLayer Is Nothing) Then
    
        'Open Flow Direction
        Dim pFlowDirRaster As IRaster
        Set pFlowDirRaster = OpenRasterDatasetFromDisk("FlowDir")
        'Convert bmps feature to raster dataset
        Dim pPourPointFClass As IFeatureClass
        Set pPourPointFClass = pPourPointsFLayer.FeatureClass
        Dim pFlowAccuRaster As IRaster
        Set pFlowAccuRaster = OpenRasterDatasetFromDisk("FlowAccu")
        'Discard the Summing Points before snapping pour points -- Sabu Paul: Aug 15, 2004
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "TYPE <> 'VirtualOutlet'"
        
        Dim pPourPointsDS As IGeoDataset
        Set pPourPointsDS = ConvertFeatureToRaster(pPourPointFClass, "ID", "POUR", pQueryFilter)  ' Nothing)
        'Get snapping distance from user
        Dim userInput
        userInput = InputBox("Enter the snapping distance for BMPs. (" & gLinearUnitName & ")", "Delineation", CInt((2 * gCellSize) + 1))
        Dim pSnapDistance As Double
        If (userInput = "" Or (Not IsNumeric(userInput))) Then
            MsgBox "Auto-Delineation cancelled or Incorrect snapping distance entered.", vbExclamation
            Exit Sub
        End If
        pSnapDistance = CDbl(userInput)
        
        'Get the snapping output result
        Dim pPourPointRaster As IRaster
        'Snap the point based the flow accumulation method without STREAMAgree -- Sabu Paul, Sep 9 2004
        Dim pTmpFlowAccuRaster As IRaster
        Set pTmpFlowAccuRaster = gHydrologyOp.FlowAccumulation(pFlowDirRaster)
        'Set pPourPointRaster = gHydrologyOp.SnapPourPoint(pPourPointsDS, pFlowAccuRaster, pSnapDistance)
        Set pPourPointRaster = gHydrologyOp.SnapPourPoint(pPourPointsDS, pTmpFlowAccuRaster, pSnapDistance)
        'Convert it to a point feature layer
        Dim pFeatureLayer As IFeatureLayer
        Set pFeatureLayer = ConvertRasterToFeature(pPourPointRaster, "Value", "SnapPoints.shp", "Point")
        AddLayerToMap pFeatureLayer, "SnapPoints"
    End If

    ModuleFlowDirAccu.RunFromFlowDir
    
''    'Generate delineated subwatershed with snapped points, add it to map
''    Dim pSubwaterRaster As IRaster
''    Set pSubwaterRaster = gHydrologyOp.WATERSHED(pFlowDirRaster, pPourPointRaster)
''    gAlgebraOp.BindRaster pSubwaterRaster, "SUB"
''    Set pSubwaterRaster = gAlgebraOp.Execute("Float([SUB])")
''    gAlgebraOp.UnbindRaster "SUB"
''    AddRasterToMap pSubwaterRaster, "SubWatershed", False
    
    'Call subroutine to create a feature layer for watershed
    Call CreateWatershedForEditing
    
    '* Rearrange layers for viewing
    Call ReArrangeLayersForViewing
   
    'Set the flag to set the manual delineation tools
    gManualDelineationFlag = True
    
    GoTo CleanUp

ShowError:
    MsgBox "GenerateSubWatersheds: " & Err.description
CleanUp:
    Set pPourPointsFLayer = Nothing
    Set pSubWaterRLayer = Nothing
    Set pFlowDirRaster = Nothing
    Set pPourPointFClass = Nothing
    Set pFlowAccuRaster = Nothing
    Set pPourPointsDS = Nothing
    Set pPourPointRaster = Nothing
    Set pFeatureLayer = Nothing
'    Set pSubwaterRaster = Nothing
End Sub

'******************************************************************************
'Subroutine: ReArrangeLayersForViewing
'Author:     Mira Chokshi
'Purpose:    Move the layers in a proper order to allow easy viewing of layers.
'******************************************************************************

Public Sub ReArrangeLayersForViewing()

  '* Get the layer id of BMP feature layer
  Dim pNextLayerIndex As Integer
  
  '* Move Streams Feature layer below BMPs layer
  pNextLayerIndex = GetInputLayerIndex("BMPs")
  MoveLayerToIndex "STREAM", pNextLayerIndex
  
  '* Move Watershed Feature layer below streams layer
  pNextLayerIndex = GetInputLayerIndex("STREAM")
  MoveLayerToIndex "Watershed", pNextLayerIndex
  
  '* Move SubWatershed Raster layer below Watershed feature layer
  pNextLayerIndex = GetInputLayerIndex("Watershed")
  MoveLayerToIndex "SubWatershed", pNextLayerIndex

  '* Move SnapPoints feature layer below Subwatershed raster layer
  pNextLayerIndex = GetInputLayerIndex("SubWatershed")
  MoveLayerToIndex "SnapPoints", pNextLayerIndex
  
End Sub

'''''******************************************************************************
'''''Subroutine: DefineAssessmentPointNetwork
'''''Author:     Mira Chokshi
'''''Purpose:    This Function is entry point for DELINEATE WATERSHEDS command.
'''''            It calls the module to check flow direction & accumulation rasters
'''''            It Asks the user for Snapping Distance and creates another
'''''            feature layer called SnapPoints that represents points moved by
'''''            snapping to flow accumulation. It then uses SnapPoints to determine
'''''            Routine between BMPs and creates the BMPNetwork table. Finally it
'''''            open the form to allow users to modify downstream BMPs for splitter.
'''''******************************************************************************
''''Public Sub DefineAssessmentPointNetwork(pRasterSubWater As IRaster)
''''
''''On Error GoTo ShowError
''''    'Open flow direction from disk
''''    Dim pRasterFlowDir As IRaster
''''    Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
''''    If pRasterFlowDir Is Nothing Then
''''        MsgBox "Not found the Flow Direction dataset."
''''        GoTo CleanUp
''''    End If
''''    'Get SnapPoints feature layer
''''    Dim pSnapPointsFLayer As IFeatureLayer
''''    Set pSnapPointsFLayer = GetInputFeatureLayer("SnapPoints")
''''    If (pSnapPointsFLayer Is Nothing) Then
''''        MsgBox "SnapPoints feature layer not found."
''''        GoTo CleanUp
''''    End If
''''    'Get BMPNetwork table
''''    Dim pBMPNetworkTable As iTable
''''    Set pBMPNetworkTable = GetInputDataTable("BMPNetwork")
''''    If (pBMPNetworkTable Is Nothing) Then
''''        MsgBox "BMPNetwork table not found."
''''        Exit Sub
''''    End If
''''    'Get the BMPs feature layer
''''    Dim pBMPFLayer As IFeatureLayer
''''    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
''''    If (pBMPFLayer Is Nothing) Then
''''        MsgBox "BMPs feature layer not found."
''''        GoTo CleanUp
''''    End If
''''
''''    Dim pBMPFClass As IFeatureClass
''''    Set pBMPFClass = pBMPFLayer.FeatureClass
''''    Dim pBMPFeatureCursor As IFeatureCursor
''''    Dim pBMPFeature As IFeature
''''
''''    'Add ID fields to SnapPoints feature layer
''''    Dim pSnapPointFClass As IFeatureClass
''''    Set pSnapPointFClass = pSnapPointsFLayer.FeatureClass
''''
''''    'Define variables for cell based computing
''''    Dim pPixelBlockFlowDir As IPixelBlock3
''''    Dim pPixelBlockSubWater As IPixelBlock3
''''    Dim pRasterPropFlowDir As IRasterProps
''''    Dim pRasterPropSubWater As IRasterProps
''''    Dim vPixelDataFlowDir As Variant
''''    Dim vPixelDataSubWater As Variant
''''    Dim pOrg As IPoint
''''    Dim pCellSize As Double
''''    Dim pOrigin As IPnt
''''    Dim pSize As IPnt
''''    Dim pLocation As IPnt
''''    Dim iCol As Integer
''''    Dim iRow As Integer
''''    Dim cCol As Integer
''''    Dim cRow As Integer
''''    Dim pValueFlowDir As Single
''''    Dim pValueSubWater As Single
''''    Dim pValueNewSubWater As Single
''''    ' get raster properties
''''    Set pRasterPropFlowDir = pRasterFlowDir
''''    Set pRasterPropSubWater = pRasterSubWater
''''    Dim pSubWaterNoDataValue As Single
''''    pSubWaterNoDataValue = pRasterPropSubWater.NoDataValue(0)
''''    'compare raster properties
''''    If (pRasterPropFlowDir.Width <> pRasterPropSubWater.Width) Or (pRasterPropFlowDir.Height <> pRasterPropSubWater.Height) Then
''''        MsgBox "Row and/or column number doesn't match for SubWatershed and Flow Direction raster layers."
''''        GoTo CleanUp
''''    End If
''''    'get raster extent, cell size
''''    Set pOrg = New Point
''''    pOrg.X = pRasterPropSubWater.Extent.XMin
''''    pOrg.Y = pRasterPropSubWater.Extent.YMax
''''    pCellSize = (pRasterPropSubWater.MeanCellSize.X + pRasterPropSubWater.MeanCellSize.Y) / 2
''''    ' create a DblPnt to hold the PixelBlock size
''''    Set pSize = New DblPnt
''''    pSize.SetCoords pRasterPropFlowDir.Width, pRasterPropFlowDir.Height
''''    ' create pixelblock the size of the input raster
''''    Set pPixelBlockFlowDir = pRasterFlowDir.CreatePixelBlock(pSize)
''''    Set pPixelBlockSubWater = pRasterSubWater.CreatePixelBlock(pSize)
''''    ' get vb supported pixel type
''''    pRasterPropFlowDir.PixelType = GetVBSupportedPixelType(pRasterPropFlowDir.PixelType)
''''    pRasterPropSubWater.PixelType = GetVBSupportedPixelType(pRasterPropSubWater.PixelType)
''''    'get status bar
''''    Dim pStatusBar As esriSystem.IStatusBar
''''    Set pStatusBar = gApplication.StatusBar
''''    Dim pStepProgressor As IStepProgressor
''''    Set pStepProgressor = pStatusBar.ProgressBar
''''    pStepProgressor.Show
''''    ' get pixeldata
''''    Set pOrigin = New DblPnt
''''    pOrigin.SetCoords 0, 0
''''    pRasterFlowDir.Read pOrigin, pPixelBlockFlowDir
''''    vPixelDataFlowDir = pPixelBlockFlowDir.PixelDataByRef(0)
''''    pRasterSubWater.Read pOrigin, pPixelBlockSubWater
''''    vPixelDataSubWater = pPixelBlockSubWater.PixelDataByRef(0)
''''    'begin processing
''''    pStepProgressor.Message = "Reading SubWatershed & Flow Direction Raster ... "
''''    Dim pCursor As ICursor
''''    Dim pRow As iRow
''''    Dim pFeatID As Integer
''''    Dim pQueryFilter As IQueryFilter
''''
''''    Dim iDSID As Long
''''    'Find next bmp for each bmp
''''    Dim pFeatureCursor As IFeatureCursor
''''    Set pFeatureCursor = pSnapPointFClass.Update(Nothing, True)
''''    Dim pFeature As IFeature
''''    Set pFeature = pFeatureCursor.NextFeature
''''    Dim pPoint As IPoint
''''    Do While Not pFeature Is Nothing
''''        'get the value and look for next downstream point
''''        Set pPoint = pFeature.Shape
''''        iRow = ((pOrg.Y - pPoint.Y) / pCellSize) - 0.5
''''        iCol = ((pPoint.X - pOrg.X) / pCellSize) - 0.5
''''        pValueSubWater = vPixelDataSubWater(iCol, iRow)
''''        cCol = iCol
''''        cRow = iRow
''''        pValueNewSubWater = vPixelDataSubWater(cCol, cRow)
''''        Do While (pValueNewSubWater = pValueSubWater)
''''            pValueFlowDir = vPixelDataFlowDir(cCol, cRow)
''''            Select Case pValueFlowDir
''''                Case 1
''''                    cCol = cCol + 1
''''                Case 2
''''                    cRow = cRow + 1
''''                    cCol = cCol + 1
''''                Case 4
''''                    cRow = cRow + 1
''''                Case 8
''''                    cRow = cRow + 1
''''                    cCol = cCol - 1
''''                Case 16
''''                    cCol = cCol - 1
''''                Case 32
''''                    cRow = cRow - 1
''''                    cCol = cCol - 1
''''                Case 64
''''                    cRow = cRow - 1
''''                Case 128
''''                    cRow = cRow - 1
''''                    cCol = cCol + 1
''''                Case Else
''''                    pValueNewSubWater = 0
''''                    Exit Do
''''            End Select
''''           pValueNewSubWater = vPixelDataSubWater(cCol, cRow)
''''           If (pValueNewSubWater = pSubWaterNoDataValue) Then
''''                pValueNewSubWater = 0
''''           End If
''''        Loop
''''
''''        'Update the network routing in BMPNetwork table
''''        pFeatID = pFeature.value(pFeatureCursor.FindField("GRID_CODE"))
''''        Set pQueryFilter = New QueryFilter
''''        pQueryFilter.WhereClause = "ID = " & pFeatID
''''        Set pCursor = pBMPNetworkTable.Update(pQueryFilter, True)
''''        iDSID = pBMPNetworkTable.FindField("DSID")
''''        Set pRow = pCursor.NextRow
''''        Do While Not pRow Is Nothing
''''            pRow.value(iDSID) = pValueNewSubWater
''''            pCursor.UpdateRow pRow
''''            Set pRow = pCursor.NextRow
''''        Loop
''''        Set pRow = Nothing
''''        Set pCursor = Nothing
''''
''''        'Update the downstream BMP in the BMPs feature layer
''''        Set pBMPFeatureCursor = pBMPFClass.Search(pQueryFilter, False)
''''        Set pBMPFeature = pBMPFeatureCursor.NextFeature
''''        If (Not pBMPFeature Is Nothing) Then
''''            pBMPFeature.value(pBMPFeatureCursor.FindField("DSID")) = pValueNewSubWater
''''            pBMPFeature.Store
''''        End If
''''        Set pBMPFeature = Nothing
''''        Set pBMPFeatureCursor = Nothing
''''        GoTo ContinueNext
''''
''''
''''ContinueNext:
''''        'Check next feature
''''        Set pFeature = pFeatureCursor.NextFeature
''''    Loop
''''    pStepProgressor.Hide
''''
''''    GoTo CleanUp
''''
''''ShowError:
''''    MsgBox "DefineAssessmentPointNetwork: " & Err.description
''''CleanUp:
''''    pStepProgressor.Hide
''''    Set pRasterFlowDir = Nothing
''''    Set pSnapPointsFLayer = Nothing
''''    Set pBMPNetworkTable = Nothing
''''    Set pSnapPointFClass = Nothing
''''    Set pPixelBlockFlowDir = Nothing
''''    Set pPixelBlockSubWater = Nothing
''''    Set pRasterPropFlowDir = Nothing
''''    Set pRasterPropSubWater = Nothing
''''    Set vPixelDataFlowDir = Nothing
''''    Set vPixelDataSubWater = Nothing
''''    Set pOrg = Nothing
''''    Set pOrigin = Nothing
''''    Set pSize = Nothing
''''    Set pLocation = Nothing
''''    Set pCursor = Nothing
''''    Set pRow = Nothing
''''    Set pQueryFilter = Nothing
''''    Set pFeatureCursor = Nothing
''''    Set pFeature = Nothing
''''    Set pPoint = Nothing
''''    Set pBMPFeature = Nothing
''''    Set pBMPFeatureCursor = Nothing
''''    Set pBMPFClass = Nothing
''''    Set pBMPFLayer = Nothing
''''End Sub

'******************************************************************************
'Subroutine: CreateFeatureClassForLineShapeFile
'Author:     Mira Chokshi
'Purpose:    Creates line shape feature class for conduit network class.
'            Adds integer fields: ID, CFROM, CTO. Sets the spatial reference of
'            the feature class same as dem's spatial reference
'******************************************************************************
Public Function CreateFeatureClassForLineShapeFile(DirName As String, FileName As String) As IFeatureClass

On Error GoTo ShowError
    'Create a unique file name for feature class
    Dim pFileName As String
    pFileName = CreateUniqueTableName(DirName, FileName)
    Dim strFolder As String
    strFolder = DirName
    Dim strShapeFieldName As String
    strShapeFieldName = "Shape"
    ' Open the folder to contain the shapefile as a workspace
    Dim pFWS As IFeatureWorkspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Set pFWS = pWorkspaceFactory.OpenFromFile(strFolder, 0)
    ' Set up a simple fields collection
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = New esriGeoDatabase.Fields
    
    Dim pFieldsEdit As IFieldsEdit
    Set pFieldsEdit = pFields
       
    ' Make the shape field
    ' it will need a geometry definition, with a spatial reference
    Dim pFieldShape As esriGeoDatabase.IField
    Dim pFieldEditShape As IFieldEdit
    Set pFieldShape = New esriGeoDatabase.Field
    Set pFieldEditShape = pFieldShape
    pFieldEditShape.name = strShapeFieldName
    pFieldEditShape.Type = esriFieldTypeGeometry
    
    'Define ID field
    Dim pFieldID As esriGeoDatabase.IField
    Dim pFieldEditID As IFieldEdit
    Set pFieldID = New esriGeoDatabase.Field
    Set pFieldEditID = pFieldID
    pFieldEditID.name = "ID"
    pFieldEditID.Type = esriFieldTypeInteger
    pFieldEditID.IsNullable = True
    
    'Define CFROM field
    Dim pFieldFrom As esriGeoDatabase.IField
    Dim pFieldEditFrom As IFieldEdit
    Set pFieldFrom = New esriGeoDatabase.Field
    Set pFieldEditFrom = pFieldFrom
    pFieldEditFrom.name = "CFROM"
    pFieldEditFrom.Type = esriFieldTypeInteger
    pFieldEditFrom.IsNullable = True
    
    'Define CTO field
    Dim pFieldTo As esriGeoDatabase.IField
    Dim pFieldEditTo As IFieldEdit
    Set pFieldTo = New esriGeoDatabase.Field
    Set pFieldEditTo = pFieldTo
    pFieldEditTo.name = "CTO"
    pFieldEditTo.Type = esriFieldTypeInteger
    pFieldEditTo.IsNullable = True
    
    'Define OUTLETTYPE field
    Dim pFieldOUTTYPE As esriGeoDatabase.IField
    Dim pFieldEditOUTTYPE As IFieldEdit
    Set pFieldOUTTYPE = New esriGeoDatabase.Field
    Set pFieldEditOUTTYPE = pFieldOUTTYPE
    pFieldEditOUTTYPE.name = "OUTLETTYPE"
    pFieldEditOUTTYPE.Type = esriFieldTypeInteger
    pFieldEditOUTTYPE.IsNullable = True
    
    'Define OUTLET Description field
    Dim pFieldOUTDESC As esriGeoDatabase.IField
    Dim pFieldEditOUTDESC As IFieldEdit
    Set pFieldOUTDESC = New esriGeoDatabase.Field
    Set pFieldEditOUTDESC = pFieldOUTDESC
    pFieldEditOUTDESC.name = "TYPEDESC"
    pFieldEditOUTDESC.Type = esriFieldTypeString
    pFieldEditOUTDESC.Length = 30
    pFieldEditOUTDESC.IsNullable = True
    
    'Define Label field
    Dim pFieldLABEL As esriGeoDatabase.IField
    Dim pFieldEditLABEL As IFieldEdit
    Set pFieldLABEL = New esriGeoDatabase.Field
    Set pFieldEditLABEL = pFieldLABEL
    pFieldEditLABEL.name = "LABEL"
    pFieldEditLABEL.Type = esriFieldTypeString
    pFieldEditLABEL.Length = 30
    pFieldEditLABEL.IsNullable = True
    
    'Get DEM raster properties
    Dim pRasterDEMProps As IRasterAnalysisProps
    If Not GetInputRasterLayer("DEM") Is Nothing Then
        Set pRasterDEMProps = GetInputRasterLayer("DEM").Raster
    End If
    'if DEM is optional use Land use
    Dim pRasterLUProps As IRasterAnalysisProps
    Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster
  
    'Get spatial reference properties
    Dim pGeomDef As IGeometryDef
    Dim pGeomDefEdit As IGeometryDefEdit
    Set pGeomDef = New GeometryDef
    Set pGeomDefEdit = pGeomDef
    With pGeomDefEdit
      .GeometryType = esriGeometryPolyline
      .HasM = False
      .HasZ = False
      If Not GetInputRasterLayer("DEM") Is Nothing Then
        Set .SpatialReference = pRasterDEMProps.AnalysisExtent.SpatialReference
      Else
        Set .SpatialReference = pRasterLUProps.AnalysisExtent.SpatialReference
      End If
    End With
    Set pFieldEditShape.GeometryDef = pGeomDef
    pFieldsEdit.AddField pFieldShape
    
    'Add other fields
    pFieldsEdit.AddField pFieldID
    pFieldsEdit.AddField pFieldFrom
    pFieldsEdit.AddField pFieldTo
    pFieldsEdit.AddField pFieldOUTTYPE
    pFieldsEdit.AddField pFieldOUTDESC
    pFieldsEdit.AddField pFieldLABEL
   
    ' Create the shapefile some parameters apply to geodatabase options and can be defaulted as Nothing
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = pFWS.CreateFeatureClass(pFileName, pFields, Nothing, Nothing, esriFTSimple, strShapeFieldName, "")
    ' Return the value
    Set CreateFeatureClassForLineShapeFile = pFeatClass
      
  GoTo CleanUp
ShowError:
    MsgBox "CreateFeatureClassForLineShapeFile: " & Err.description
CleanUp:
    Set pFWS = Nothing
    Set pWorkspaceFactory = Nothing
    Set pFields = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldShape = Nothing
    Set pFieldEditShape = Nothing
    Set pFieldOUTTYPE = Nothing
    Set pFieldOUTDESC = Nothing
    Set pFieldEditOUTTYPE = Nothing
    Set pFieldEditOUTDESC = Nothing
    Set pFieldFrom = Nothing
    Set pFieldEditFrom = Nothing
    Set pFieldTo = Nothing
    Set pFieldEditTo = Nothing
    Set pRasterDEMProps = Nothing
    Set pRasterLUProps = Nothing
    Set pGeomDef = Nothing
    Set pGeomDefEdit = Nothing
    Set pFeatClass = Nothing
End Function


'Subroutine to get the bmp route type total, weir, outlet or underdrain
Public Function GetBMPRouteType(ByRef pPoint As IPoint, ByRef pointID As Integer, ByVal pBMPLayerName As String, bVFSUpstream As Boolean) As Integer
On Error GoTo ShowError
    
    Dim pSchematicLayer As IFeatureLayer
    Set pSchematicLayer = GetInputFeatureLayer(pBMPLayerName)
    If (pSchematicLayer Is Nothing) Then
        MsgBox pBMPLayerName & " layer not found."
        Exit Function
    End If
    Dim pBMPNetwork As iTable
    Set pBMPNetwork = GetInputDataTable("BMPNetwork")
    If (pBMPNetwork Is Nothing) Then
        MsgBox "BMPNetwork table not present."
        Exit Function
    End If
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pSchematicLayer.FeatureClass
    
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pPoint.Envelope
    ExpandPointEnvelope pEnvelope
    
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, False)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    
    Dim iID As Long
    iID = pFeatureCursor.FindField("ID")
        
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iOUTLETTYPE As Long
    iOUTLETTYPE = pBMPNetwork.FindField("OutletType")
    Dim pOutletCount As Integer
    pOutletCount = -1
    If (Not pFeature Is Nothing) Then
        pointID = pFeature.value(iID)
        pQueryFilter.WhereClause = "ID = " & pointID
        pOutletCount = pBMPNetwork.RowCount(pQueryFilter)
        'Set the point value back
        Set pPoint = pFeature.ShapeCopy
    End If
        
    '** If pOutletCount = -1, maybe the user is trying to click a VFS
    '** If the bmp feature name passed is BMPs, then the extent is
    '** not a schematic layer.
    Dim pVFSLayer As IFeatureLayer
    
    If (bVFSUpstream = True) Then 'Restrict the VFS to upstream only - Sabu Paul - June 11, 2007
        Set pVFSLayer = GetInputFeatureLayer("VFS")
        If ((pOutletCount = -1) And (pBMPLayerName = "BMPs") And (Not pVFSLayer Is Nothing)) Then
            '** Intersect the vfs with the point to find the route
            Dim pVFSClass As IFeatureClass
            Set pVFSClass = pVFSLayer.FeatureClass
            Set pFeatureCursor = pVFSClass.Search(pSpatialFilter, False)
            Set pFeature = pFeatureCursor.NextFeature
            iID = pFeatureCursor.FindField("ID")
            pOutletCount = -1
            Dim pPolyline As IPolyline
            If (Not pFeature Is Nothing) Then
                Set pPolyline = pFeature.Shape
                pointID = pFeature.value(iID)
                pOutletCount = 1
    ''            'Set the point value back
    ''            If (bVFSUpstream = True) Then
    ''               Set pPoint = pPolyline.ToPoint
    ''            Else
    ''               Set pPoint = pPolyline.FromPoint
    ''            End If
            End If
        
        End If
    End If
    'Return the route type
    GetBMPRouteType = pOutletCount
    
    GoTo CleanUp
     
ShowError:
    MsgBox "GetBMPRouteType: " & Err.description
CleanUp:
    Set pSchematicLayer = Nothing
    Set pBMPNetwork = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pVFSLayer = Nothing
    Set pPolyline = Nothing
End Function


Public Sub UpdateBMPNetworkTableRoute(FromPointID As Integer, ToPointID As Integer, pOUTLETType As Integer, pRouteLayerName As String, pBMPLayerName As String)
On Error GoTo ShowError

    Dim pConduitsFLayer As IFeatureLayer
    Set pConduitsFLayer = GetInputFeatureLayer("Conduits")
    
    Dim pBMPNetwork As iTable
    Set pBMPNetwork = GetInputDataTable("BMPNetwork")
    If (pBMPNetwork Is Nothing) Then
        MsgBox "BMPNetwork table not present."
        Exit Sub
    End If
    
    'Check the reverse orientation
    Dim iOUTLETTYPE As Long
    iOUTLETTYPE = pBMPNetwork.FindField("OutletType")
    Dim pQueryFilter As IQueryFilter
    Dim iID As Long
    iID = pBMPNetwork.FindField("ID")
    Dim iDSID As Long
    iDSID = pBMPNetwork.FindField("DSID")
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    'If conduits layer is not found, it will be re-created, means set dsid = 0 for all bmps
    If (pConduitsFLayer Is Nothing) Then
        Set pCursor = pBMPNetwork.Search(Nothing, True)
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            pRow.value(iDSID) = 0
            pRow.Store
            Set pRow = pCursor.NextRow
        Loop
    End If
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & ToPointID & " AND DSID = " & FromPointID & " AND OutletType = " & pOUTLETType
    Set pCursor = pBMPNetwork.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    If (Not pRow Is Nothing) Then
        pRow.value(iDSID) = 0
        pRow.Store
    End If
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & FromPointID & " AND OutletType = " & pOUTLETType
    Set pCursor = pBMPNetwork.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    If (Not pRow Is Nothing) Then
        pRow.value(iDSID) = ToPointID
        pRow.Store  'Update the network table
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "UpdateBMPNetworkTableRoute :" & Err.description
CleanUp:
   Set pConduitsFLayer = Nothing
   Set pBMPNetwork = Nothing
   Set pQueryFilter = Nothing
   Set pCursor = Nothing
   Set pRow = Nothing
   Set pQueryFilter = Nothing
End Sub


'11/26/2008 Ying Cao: update VFS class as well if IDs are from VFS layer
Public Sub UpdateBMPFeatureClassInformation(FromPointID As Integer, ToPointID As Integer)
On Error GoTo ShowError
    'for BMP layer
    Dim pBMPFeatureLayer As IFeatureLayer
    Set pBMPFeatureLayer = GetInputFeatureLayer("BMPs")
    Dim pBMPFeatureClass As IFeatureClass
    Set pBMPFeatureClass = pBMPFeatureLayer.FeatureClass
    
    Dim iID As Long
    iID = pBMPFeatureClass.FindField("ID")
    Dim iDSID As Long
    iDSID = pBMPFeatureClass.FindField("DSID")
    
    'for VFS layer
    Dim pVFSFeatureLayer As IFeatureLayer
    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFeatureClass As IFeatureClass
    Dim iVFSID As Long
    Dim iVFSDSID As Long
    
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    
    Dim pVFSFeatureCursor As IFeatureCursor
    Dim pVFSFeature As IFeature
    
    'BMP and VFS share same filter
    Dim pQueryFilter As IQueryFilter
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & ToPointID & " AND DSID = " & FromPointID
    Set pFeatureCursor = pBMPFeatureClass.Search(pQueryFilter, False)
    Set pFeature = pFeatureCursor.NextFeature
    
    If Not pVFSFeatureLayer Is Nothing Then
        Set pVFSFeatureClass = pVFSFeatureLayer.FeatureClass
        iVFSID = pVFSFeatureClass.FindField("ID")
        iVFSDSID = pVFSFeatureClass.FindField("DSID")
        Set pVFSFeatureCursor = pVFSFeatureClass.Search(pQueryFilter, False)
        Set pVFSFeature = pVFSFeatureCursor.NextFeature
    End If
    
    If (Not pFeature Is Nothing) Then
        pFeature.value(iDSID) = 0
        pFeature.Store
    ElseIf (Not pVFSFeature Is Nothing) Then
        pVFSFeature.value(iVFSDSID) = 0
        pVFSFeature.Store
    End If
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & FromPointID
    Set pFeatureCursor = pBMPFeatureClass.Search(pQueryFilter, False)
    Set pFeature = pFeatureCursor.NextFeature
    
    If Not pVFSFeatureLayer Is Nothing Then
        Set pVFSFeatureCursor = pVFSFeatureClass.Search(pQueryFilter, False)
        Set pVFSFeature = pVFSFeatureCursor.NextFeature
    End If
    
    If (Not pFeature Is Nothing) Then
        pFeature.value(iDSID) = ToPointID
        pFeature.Store  'Update the network table
    ElseIf (Not pVFSFeature Is Nothing) Then
        pVFSFeature.value(iVFSDSID) = ToPointID
        pVFSFeature.Store
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "UpdateBMPFeatureClassInformation :" & Err.description
CleanUp:  'Sabu Paul Jan 17, 2005 -- Cleanup section added
   Set pBMPFeatureLayer = Nothing
   Set pBMPFeatureClass = Nothing
   Set pFeature = Nothing
   Set pFeatureCursor = Nothing
   Set pVFSFeatureLayer = Nothing
   Set pVFSFeatureClass = Nothing
   Set pVFSFeature = Nothing
   Set pVFSFeatureCursor = Nothing
   Set pQueryFilter = Nothing
   
End Sub

'Added routing to VFS 11/21/2008 (Ying Cao)
Public Sub RenderBasintoBMPRouting()

    'Get watershed feature layer
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If pWatershedFLayer Is Nothing Then
        MsgBox "Watershed feature layer not found."
        Exit Sub
    End If
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    Dim iBasinIDFld As Long
    iBasinIDFld = pWatershedFClass.FindField("ID")
    Dim iBmpIdFld As Long
    iBmpIdFld = pWatershedFClass.FindField("BMPID")
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pQueryFilterV As IQueryFilter
    Set pQueryFilterV = New QueryFilter
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pWatershedFClass.Search(Nothing, True)
    Dim pWatershedFeature As IFeature
    Set pWatershedFeature = pFeatureCursor.NextFeature
    
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    If pBMPFLayer Is Nothing Then
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    Dim pBMPFeatureCursor As IFeatureCursor
    Dim pBMPFeature As IFeature
    
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    If Not pVFSFLayer Is Nothing Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
    End If
    
    Dim pVFSFeatureCursor As IFeatureCursor
    Dim pVFSFeature As IFeature
  
    'ID for VFS
    'Dim iVfsIdFld As Long
    'iVfsIdFld = pWatershedFClass.FindField("BMPID")
    Dim pMidPt As IPoint
    Dim pCurve As ICurve
        
    'Define a feature class to store the routing information for basin to BMP
    Dim pBasinToBMPRoutingLayer As IFeatureLayer
    Set pBasinToBMPRoutingLayer = GetInputFeatureLayer("BasinRouting")
    If Not (pBasinToBMPRoutingLayer Is Nothing) Then
        DeleteLayerFromMap ("BasinRouting")
    End If
    'Create a new line feature class
    Dim pBasinToBMPRoutingClass As IFeatureClass
    Set pBasinToBMPRoutingClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, "basinroute")
    Dim pBasinToBMPFeature As IFeature
    Dim iIDFld As Long
    iIDFld = pBasinToBMPRoutingClass.FindField("ID")
    
    Dim pPointCollection As IPointCollection
    Dim pCenterPt As IPoint
    Dim pEndPt As IPoint
    
    'search for BMP
    Do While Not pWatershedFeature Is Nothing
        pQueryFilter.WhereClause = "ID = " & pWatershedFeature.value(iBmpIdFld)
        Set pBMPFeatureCursor = Nothing
        Set pBMPFeatureCursor = pBMPFClass.Search(pQueryFilter, True)
        Set pBMPFeature = Nothing
        Set pBMPFeature = pBMPFeatureCursor.NextFeature
        
        
        pQueryFilterV.WhereClause = "ID = " & pWatershedFeature.value(iBmpIdFld)
        Set pVFSFeatureCursor = Nothing
        If Not pVFSFLayer Is Nothing Then
            Set pVFSFeatureCursor = pVFSFClass.Search(pQueryFilterV, True)
            Set pVFSFeature = Nothing
            Set pVFSFeature = pVFSFeatureCursor.NextFeature
        End If
        
        If Not (pBMPFeature Is Nothing) Then
            'Add the watershed center as the start point
            Set pPointCollection = New Polyline
            Set pCenterPt = New Point
            pCenterPt.X = (pWatershedFeature.Shape.Envelope.XMax + pWatershedFeature.Shape.Envelope.XMin) / 2
            pCenterPt.Y = (pWatershedFeature.Shape.Envelope.YMax + pWatershedFeature.Shape.Envelope.YMin) / 2
            pPointCollection.AddPoint pCenterPt

            'Add the bmp as the end point
            Set pEndPt = pBMPFeature.Shape
            pPointCollection.AddPoint pEndPt
            'Create a new basin to bmp route and save it
            Set pBasinToBMPFeature = pBasinToBMPRoutingClass.CreateFeature
            Set pBasinToBMPFeature.Shape = pPointCollection
            pBasinToBMPFeature.value(iIDFld) = pWatershedFeature.value(iBasinIDFld)
            pBasinToBMPFeature.Store
        ElseIf Not (pVFSFeature Is Nothing) Then
            'Add the watershed center as the start point
            Set pPointCollection = New Polyline
            Set pMidPt = New Point
            Set pCenterPt = New Point
            pCenterPt.X = (pWatershedFeature.Shape.Envelope.XMax + pWatershedFeature.Shape.Envelope.XMin) / 2
            pCenterPt.Y = (pWatershedFeature.Shape.Envelope.YMax + pWatershedFeature.Shape.Envelope.YMin) / 2
            pPointCollection.AddPoint pCenterPt

            'Add the midpoint of VFS as the end point
            Set pCurve = pVFSFeature.Shape
            pCurve.QueryPoint esriNoExtension, 0.5, True, pMidPt
            pPointCollection.AddPoint pMidPt
        
            'Create a new basin to VFS route and save it
            Set pBasinToBMPFeature = pBasinToBMPRoutingClass.CreateFeature
            Set pBasinToBMPFeature.Shape = pPointCollection
            pBasinToBMPFeature.value(iIDFld) = pWatershedFeature.value(iBasinIDFld)
            pBasinToBMPFeature.Store
            
        End If
        Set pWatershedFeature = pFeatureCursor.NextFeature
    Loop

    'Create a new feature layer for basinroutine
    Set pBasinToBMPRoutingLayer = New FeatureLayer
    Set pBasinToBMPRoutingLayer.FeatureClass = pBasinToBMPRoutingClass
    AddLayerToMap pBasinToBMPRoutingLayer, "BasinRouting"

    ' create a new simple line renderer
    Dim pRen As ISimpleRenderer
    Dim pGeoFeatLyr As IGeoFeatureLayer
    Set pGeoFeatLyr = pBasinToBMPRoutingLayer
    Set pRen = pGeoFeatLyr.Renderer
    Set pRen.Symbol = ReturnBasintoBMPRouteSymbol
    gMxDoc.ActiveView.Refresh
    gMxDoc.UpdateContents
        
End Sub


'Subroutine to find the most upstream BMP in each watershed
Public Sub CreateSubBasinToBMPRouting()
On Error GoTo ShowError

  'Get watershed feature layer
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    
   
    If pWatershedFLayer Is Nothing Then
        MsgBox "Watershed feature layer not found."
        Exit Sub
    End If
    'Get BMPs feature layer
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    
    If pBMPFLayer Is Nothing Then
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If

    'Define a dictionary to save bmp to downstream bmp info
    Dim pDSBMPDictionary As Scripting.Dictionary
    Set pDSBMPDictionary = CreateObject("Scripting.Dictionary")
    
    
    'Define variables for BMPs feature layer access
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    Dim pBMPFCursor As IFeatureCursor
    Dim pBMPFeature As IFeature
    Dim pSpatialFilter As ISpatialFilter
    Dim iID As Long
    iID = pBMPFClass.FindField("ID")
    
    'Define variables for BMPNetwork table
    Dim pBMPTable As iTable
    Set pBMPTable = GetInputDataTable("BMPNetwork")
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iTableDSID As Long
    iTableDSID = pBMPTable.FindField("DSID")
    
    'Define variables for Watershed feature layer access
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    Dim iIDFld As Long
    iIDFld = pWatershedFClass.FindField("ID")
    Dim iBmpIdFld As Long
    iBmpIdFld = pWatershedFClass.FindField("BMPID")
    If (iBmpIdFld < 0) Then
       'Add BMPID field
       Dim pFieldEditBMPID As IFieldEdit
       Set pFieldEditBMPID = New esriGeoDatabase.Field
       pFieldEditBMPID.name = "BMPID"
       pFieldEditBMPID.Type = esriFieldTypeInteger
       pWatershedFClass.AddField pFieldEditBMPID
       iBmpIdFld = pWatershedFClass.FindField("BMPID")
    End If
    
    'Query the watershed layer
    Dim pWatershedFCursor As IFeatureCursor
    Dim pWatershedFeature As IFeature
    Set pWatershedFCursor = pWatershedFClass.Search(Nothing, False)
    Set pWatershedFeature = pWatershedFCursor.NextFeature
    'For each watershed feature, find the most upstream BMP
    Do While Not (pWatershedFeature Is Nothing)
        'Search all BMPs in selected watershed
        Set pSpatialFilter = New SpatialFilter
        Set pSpatialFilter.Geometry = pWatershedFeature.Shape
        pSpatialFilter.SpatialRel = esriSpatialRelIntersects
        Set pBMPFCursor = pBMPFClass.Search(pSpatialFilter, True)
        Set pBMPFeature = pBMPFCursor.NextFeature
        Do While Not pBMPFeature Is Nothing
            'For this BMP find all Downstream BMP's
            pQueryFilter.WhereClause = "ID = " & pBMPFeature.value(iID)
            'Add this bmp to dictionary if it does not exist
            If Not (pDSBMPDictionary.Exists(pBMPFeature.value(iID))) Then
                pDSBMPDictionary.add pBMPFeature.value(iID), True
            End If
            Set pCursor = pBMPTable.Search(pQueryFilter, True)
            Set pRow = pCursor.NextRow
            Do While Not pRow Is Nothing
                pDSBMPDictionary.Item(pRow.value(iTableDSID)) = False
                Set pRow = pCursor.NextRow
            Loop
            Set pBMPFeature = pBMPFCursor.NextFeature
        Loop
        'Iterate over the dictionary to find the key with value = TRUE
        Dim pKeys
        pKeys = pDSBMPDictionary.keys
        Dim pkey
        Dim pUsBMP As Integer
        Dim i As Integer
        For i = 0 To (pDSBMPDictionary.Count - 1)
            pkey = pKeys(i)
            If (pDSBMPDictionary.Item(pkey) = False) Then
                pDSBMPDictionary.Remove (pkey)
            Else
                pUsBMP = pkey
            End If
        Next
    
        'Update the watershed feature, with draining bmp if found, else zero
        pWatershedFeature.value(iBmpIdFld) = 0
        If (pDSBMPDictionary.Count = 1) Then
            pWatershedFeature.value(iBmpIdFld) = pUsBMP
        End If
        pWatershedFeature.Store
        
        'Purge the dictionary and move to next watershed feature
        pDSBMPDictionary.RemoveAll
        Set pWatershedFeature = pWatershedFCursor.NextFeature
    Loop

    GoTo CleanUp

CleanUp:
    Set pWatershedFLayer = Nothing
    Set pBMPFLayer = Nothing
    Set pDSBMPDictionary = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFCursor = Nothing
    Set pBMPFeature = Nothing
    Set pSpatialFilter = Nothing
    Set pBMPTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pWatershedFClass = Nothing
    Set pFieldEditBMPID = Nothing
    Set pWatershedFCursor = Nothing
    Set pWatershedFeature = Nothing
    Set pKeys = Nothing
    Exit Sub
ShowError:
    MsgBox "CreateSubBasinToBMPVFSRouting: " & Err.description
End Sub


'Subroutine to find the most upstream VFS in each watershed
'history: 11/21/2008 Ying Cao
Public Sub CreateSubBasinToVFSRouting()
On Error GoTo ShowError

  'Get watershed feature layer
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")

    If pWatershedFLayer Is Nothing Then
        MsgBox "Watershed feature layer not found."
        Exit Sub
    End If
    
    'Get VFS feature layer
    Dim pVFSFeatureLayer As IFeatureLayer
    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")
    
    If pVFSFeatureLayer Is Nothing Then
        'MsgBox "VFS feature layer not found."
        Exit Sub
    End If
   
    'Define variable for VFSs feature layer access
    Dim pVFSFClass As IFeatureClass
    Set pVFSFClass = pVFSFeatureLayer.FeatureClass
    Dim pVFSFCursor As IFeatureCursor
    Dim pVFSFeature As IFeature
'    Dim pSpatialFilter As ISpatialFilter
    Dim iID As Long
    iID = pVFSFClass.FindField("ID")
    
    '** find nearest VFS
    Dim pProximityOp As IProximityOperator
    Dim curDis As Double, minDis As Double, thresholdDis As Double
    thresholdDis = 999999
    minDis = thresholdDis
    Dim iVFSID As Long
    
    'Define variable for BMPNetwork table
    Dim pVFSTable As iTable
    Set pVFSTable = GetInputDataTable("BMPNetwork")
    Dim pQueryFilter As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iTableDSID As Long
    iTableDSID = pVFSTable.FindField("DSID")

    'Define variables for Watershed feature layer access
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    
    Dim iVfsIdFld As Long
    iVfsIdFld = pWatershedFClass.FindField("BMPID")  'save both BMP and VFS ID
    
    'Query the watershed layer
    Dim pWatershedFCursor As IFeatureCursor
    Dim pWatershedFeature As IFeature
    Set pWatershedFCursor = pWatershedFClass.Search(Nothing, True)
    Set pWatershedFeature = pWatershedFCursor.NextFeature

    'For each watershed feature, find the most upstream VFS
    Do While Not (pWatershedFeature Is Nothing)
        If pWatershedFeature.value(iVfsIdFld) = 0 Then  'if BMPID!=0, the watershed is already assigned to a BMP
            Set pVFSFCursor = pVFSFClass.Search(Nothing, True)
            Set pVFSFeature = pVFSFCursor.NextFeature
            
            Do While Not pVFSFeature Is Nothing
                Set pProximityOp = pVFSFeature.Shape
                curDis = pProximityOp.ReturnDistance(pWatershedFeature.Shape)
                
                If curDis < minDis And curDis >= 0 Then    'VFS either intersect or outside the corresponding watershed
                  minDis = curDis
                  iVFSID = pVFSFeature.value(iID)
                  MsgBox "found nearest VFS ID" & iVFSID
                End If
            Set pVFSFeature = pVFSFCursor.NextFeature
            Loop
        
            'Update the watershed feature, with draining VFS if found, else zero
            pWatershedFeature.value(iVfsIdFld) = 0
            If (minDis < thresholdDis And minDis >= 0) Then
                MsgBox "save value: " & iVFSID
                pWatershedFeature.value(iVfsIdFld) = iVFSID
            End If
            pWatershedFeature.Store
            
            minDis = thresholdDis
        End If
        Set pWatershedFeature = pWatershedFCursor.NextFeature
    Loop

    GoTo CleanUp

CleanUp:
    Set pWatershedFLayer = Nothing
    Set pVFSFeatureLayer = Nothing
    Set pVFSFClass = Nothing
    Set pVFSFCursor = Nothing
    Set pVFSFeature = Nothing
    Set pVFSTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pWatershedFClass = Nothing
    Set pWatershedFCursor = Nothing
    Set pWatershedFeature = Nothing
    Exit Sub
ShowError:
    MsgBox "CreateSubBasinToVFSRouting: " & Err.description
End Sub
