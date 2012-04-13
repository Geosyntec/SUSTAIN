Attribute VB_Name = "ModuleRasterUtils"

'******************************************************************************
'   Application: Sustain - BMP Siting Tool
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Arun Raj
'   Developer:   Arun Raj
'******************************************************************************


Option Explicit
Option Base 0
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\ModuleRasterUtils.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

'******************************************************************************
'Subroutine: CheckSpatialAnalystLicense
'Purpose:    Check the availability of Spatial Analyst license, returns
'            TRUE if SA license found, else returns FALSE. This subroutine
'            should be called at the beginning of any process using SA.
'******************************************************************************
Public Function CheckSpatialAnalystLicense() As Boolean
On Error GoTo ShowError
    
    Dim pLicManager As IExtensionManager
    Set pLicManager = New ExtensionManager
    
    Dim pLicAdmin As IExtensionManagerAdmin
    Set pLicAdmin = pLicManager
    
    Dim saUID As Variant
    saUID = "esriSpatialAnalystUI.SAExtension.1"
    
    Dim pUID As New UID
    pUID.Value = saUID
    
    Dim v As Variant
    Call pLicAdmin.AddExtension(pUID, v)
    
    Dim pExtension As IExtension
    Set pExtension = pLicManager.FindExtension(pUID)
    
    Dim pExtensionConfig As IExtensionConfig
    Set pExtensionConfig = pExtension
    pExtensionConfig.State = esriESEnabled
    
    CheckSpatialAnalystLicense = True
    GoTo Cleanup

ShowError:
    MsgBox "Failed in License Checking - " & Err.Description
Cleanup:
    Set pLicManager = Nothing
    Set pLicAdmin = Nothing
    Set saUID = Nothing
    Set pUID = Nothing
    Set v = Nothing
    Set pExtension = Nothing
    Set pExtensionConfig = Nothing
End Function

'##########################################################################
' RASTER OPERATIONS
'##########################################################################

Public Sub Calculate_Slope(ByVal pRasterLayer As IRasterLayer)
    
    On Error GoTo ErrorHandler
    
    Dim pRasterWS As IWorkspace
    Dim pRasterWSFact As IWorkspaceFactory
    Set pRasterWSFact = New RasterWorkspaceFactory
    Set pRasterWS = pRasterWSFact.OpenFromFile(gRasterfolder, 0)
    
    'Create a GPUtilities object
    Dim pGPUtils As IGPUtilities
    Set pGPUtils = New GPUtilities
    
    'perform slope
    'Create a RasterSurfaceOp operator
    Dim pSurfaceOp As ISurfaceOp
    Set pSurfaceOp = New RasterSurfaceOp
    
    'Set output workspace
    Dim pEnv As IRasterAnalysisEnvironment
    Set pEnv = pSurfaceOp
    Set pEnv.OutWorkspace = pRasterWS
    
    Dim pRaster As IRaster
    Set pRaster = pRasterLayer.Raster 'get raster from rasterlayer
    
    'Perform spatial operation
    Dim pOutRaster As IRaster
    Set pOutRaster = pSurfaceOp.Slope(pRaster, esriGeoAnalysisSlopePercentrise)
    
    ' Create a integer raster......................................
    gMapAlgebraOp.BindRaster pOutRaster, "SLOPE"
    Set pRaster = gMapAlgebraOp.Execute("Int([SLOPE])")
    gMapAlgebraOp.UnbindRaster "SLOPE"

    'Write it to the disk
    WriteRasterDatasetToDisk pRaster, "SLOPE"
    Set pRaster = Nothing ' Destroy the Create Grid.....

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Calculate_Slope " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
End Sub

Public Sub Create_FlowDirectionandAccumulation(ByVal pRaster As IRaster)

    On Error GoTo ErrorHandler
    'fill raw dem
    Dim pFillDEMRaster As IRaster
    Set pFillDEMRaster = FillRawDEM(pRaster)
    'Call the subroutine to create flow direction
    Dim pFlowdirRaster As IRaster
    Set pFlowdirRaster = gHydrologyOp.FlowDirection(pFillDEMRaster, True, True)
    'Call the subroutine to create flow accumulation
    Dim pFlowAccRaster As IRaster
    Set pFlowAccRaster = gHydrologyOp.FlowAccumulation(pFlowdirRaster)
    
    'Write it to the disk
    WriteRasterDatasetToDisk pFlowAccRaster, "FLOW"
    Set pFlowAccRaster = Nothing ' Destroy the Create Grid.....

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Create_FlowDirectionandAccumulation " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
End Sub


'******************************************************************************
'Subroutine: InitializeOperators
'Purpose:    Initializes Global Map Algebra, Hydrology, Neighborhood, Reclass
'            Operators.
'            Set extent, cell size to dem raster's extent and cell size
'******************************************************************************
Public Function InitializeOperators() As Boolean

    On Error GoTo EH
  
    Dim pDEMRasterLayer As IRasterLayer
    Set pDEMRasterLayer = GetInputFeatureLayer(gDEMdata)
    
    If pDEMRasterLayer Is Nothing Then GoTo Cleanup
    
    Dim pDEMRaster As IRaster
    Set pDEMRaster = pDEMRasterLayer.Raster
        
    'Get the raster props
    Dim pDEMRasterProps As IRasterProps
    Set pDEMRasterProps = pDEMRaster
    
    'Get the raster cell size
    gCellSize = (pDEMRasterProps.MeanCellSize.x + pDEMRasterProps.MeanCellSize.y) / 2
        
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pRWS As IRasterWorkspace2
    Set pRWS = pWSF.OpenFromFile(gWorkingfolder, 0)
    Dim pRAEnv As IRasterAnalysisEnvironment

    ' Create the global gAlgebraOp object
    Set gMapAlgebraOp = New RasterMapAlgebraOp
    Set pRAEnv = gMapAlgebraOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.x
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = pDEMRaster

    ' Create the global gHydrologyOp object
    Set gHydrologyOp = New RasterHydrologyOp
    Set pRAEnv = gHydrologyOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.x
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = pDEMRaster
    
    InitializeOperators = True
    GoTo Cleanup
    
EH:
    MsgBox "Failed in Initializing the system - " & Err.Description

Cleanup:
    Set pDEMRasterLayer = Nothing
    Set pDEMRasterProps = Nothing
    Set pRWS = Nothing
    Set pWSF = Nothing
    Set pRAEnv = Nothing
End Function

Public Function Calc_Slope_MapAlgebra(ByVal pRasLayer As IRasterLayer) As IRasterLayer
    
    On Error GoTo ErrorHandler
    Dim pInRaster As IRaster
    Set pInRaster = pRasLayer.Raster
         
    'Set output workspace
    Dim pWs As IWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Set pWs = pWSF.OpenFromFile(gWorkingfolder, 0)
    
    'Bind a raster
    gMapAlgebraOp.BindRaster pInRaster, "R1"
    
    'Execute the Map Algebra expression to calculate slope of the input raster
    Dim pOutRaster As IRaster
    Set pOutRaster = gMapAlgebraOp.Execute("Slope([R1])")
    
    'Add output into ArcMap as a raster layer
    Dim pOutRasLayer As IRasterLayer
    Set pOutRasLayer = New RasterLayer
    pOutRasLayer.CreateFromRaster pOutRaster
    
    Set Calc_Slope_MapAlgebra = pOutRasLayer
    
Cleanup:

  Exit Function
ErrorHandler:
  HandleError True, "Calc_Slope_MapAlgebra " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function

'******************************************************************************
'Subroutine: FillRawDEM
'Purpose:    This module uses hydrology operator to fill raw dem. Input parameter
'            is raw dem and output is the filled dem
'******************************************************************************
Private Function FillRawDEM(pDEMRaster As IRaster) As IRaster
On Error GoTo ShowError

    'Fill the raw dem
    
    Dim pFillDEMRaster As IRaster
    Set pFillDEMRaster = gHydrologyOp.Fill(pDEMRaster)
    'Write it to the disk
    WriteRasterDatasetToDisk pFillDEMRaster, "FillDEM"
    Set FillRawDEM = pFillDEMRaster
    GoTo Cleanup
    
ShowError:
    MsgBox "FillRawDEM: " & Err.Description
Cleanup:
    Set pFillDEMRaster = Nothing
End Function

'******************************************************************************
'Subroutine: OpenRasterDatasetFromDisk
'Purpose:    Opens raster dataset from disk. Requires the name of the raster
'            dataset. This function does not require the directory path.
'            It assums the directory path as TEMP directory.
'******************************************************************************
Public Function OpenRasterDatasetFromDisk(pRasterName As String) As IRaster
On Error GoTo ShowError
    ' check if raster dataset exist
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    If Not fsObj.FolderExists(gRasterfolder & "\" & pRasterName) Then
        Set OpenRasterDatasetFromDisk = Nothing
        GoTo Cleanup
    End If
    Set fsObj = Nothing
          
    'Open workspace
    Dim pWF As IWorkspaceFactory
    Set pWF = New RasterWorkspaceFactory
    Dim pRW As IRasterWorkspace
    Set pRW = pWF.OpenFromFile(gRasterfolder, 0)
    Dim pRDS As IRasterDataset
    If (pRW.IsWorkspace(gRasterfolder)) Then
      Set pRDS = pRW.OpenRasterDataset(LCase(pRasterName))
    End If
    If pRDS Is Nothing Then
      GoTo Cleanup
    End If
    'Get Raster from the raster dataset
    Dim pRaster As IRaster
    Set pRaster = pRDS.CreateDefaultRaster
    'Return raster
    Set OpenRasterDatasetFromDisk = pRaster
    GoTo Cleanup
ShowError:
    MsgBox "OpenRasterDatasetFromDisk: " & Err.Description
Cleanup:
    Set fsObj = Nothing
    Set pWF = Nothing
    Set pRW = Nothing
    Set pRDS = Nothing
    Set pRaster = Nothing
End Function


'******************************************************************************
'Subroutine: WriteRasterDatasetToDisk
'Purpose:    Writes the temporary raster (in memory) to the disk. This function
'            requires the name of raster file, assumes the output directory as
'            TEMP directory.
'******************************************************************************
Public Sub WriteRasterDatasetToDisk(ByRef pRaster As IRaster, pOutName As String)
On Error GoTo ShowError
    ' Create a raster workspace
    Dim pRWS As IRasterWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Set pRWS = pWSF.OpenFromFile(gRasterfolder, 0)
    'Delete the raster dataset if present on disk
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(gRasterfolder & "\" & pOutName)) Then
        Dim pRasterDataset As IRasterDataset
        Set pRasterDataset = pRWS.OpenRasterDataset(pOutName)
        If Not (pRasterDataset Is Nothing) Then
            Dim pDataset As IDataset
            Set pDataset = pRasterDataset
            If pDataset.CanDelete Then pDataset.Delete
            Set pDataset = Nothing
            Set pRasterDataset = Nothing
        End If
    End If
    ' SaveAs the projected raster
    Dim pDS As IDataset
    Dim pRasBandCol As IRasterBandCollection
    Set pRasBandCol = pRaster
    Set pDS = pRasBandCol.SaveAs(pOutName, pRWS, "GRID")
    GoTo Cleanup:
ShowError:
    MsgBox "WriteRasterDatasetToDisk: " & Err.Description, vbExclamation, pOutName
Cleanup:
    Set pRWS = Nothing
    Set pWSF = Nothing
    Set fso = Nothing
    Set pRasterDataset = Nothing
    Set pDataset = Nothing
    Set pDS = Nothing
    Set pRasBandCol = Nothing
End Sub
Public Sub Read_Raster()

    On Error GoTo ErrorHandler
    Dim x, y
    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    Dim pPoint As IPoint
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(x, y)
    
    Dim pBlockSize As IPnt
    Set pBlockSize = New DblPnt
    pBlockSize.SetCoords 1#, 1#
    
    Dim pLayer As IRasterLayer
    Dim pPixelBlock As IPixelBlock 'number of bands
    Dim vValue As Variant
    Dim i As Long, j As Long
    Dim sPixelVals As String
    sPixelVals = "No Raster"
    Dim pRasterProps As IRasterProps
    Dim dXSize As Double, dYSize As Double
    Dim pPixel As IPnt
    Set pPixel = New DblPnt
    
    For i = 0 To gMap.LayerCount - 1
      If (TypeOf gMap.Layer(i) Is IRasterLayer) Then
        Set pLayer = gMap.Layer(i) 'if a raster layer then set it
        Set pPixelBlock = pLayer.Raster.CreatePixelBlock(pBlockSize)
    
        Set pRasterProps = pLayer.Raster
        dXSize = pRasterProps.Extent.XMax - pRasterProps.Extent.XMin
        dYSize = pRasterProps.Extent.YMax - pRasterProps.Extent.YMin
        dXSize = dXSize / pRasterProps.Width
        dYSize = dYSize / pRasterProps.Height
    
        pPixel.x = (pPoint.x - pRasterProps.Extent.XMin) / dXSize
        pPixel.y = (pRasterProps.Extent.YMax - pPoint.y) / dYSize
    
        pLayer.Raster.Read pPixel, pPixelBlock
        For j = 0 To pPixelBlock.Planes - 1
          If (sPixelVals = "No Raster") Then
            sPixelVals = "("
          Else
            sPixelVals = sPixelVals & ", "
          End If
          vValue = pPixelBlock.GetVal(j, 0, 0)
          sPixelVals = sPixelVals & CStr(vValue)
        Next j
        If (sPixelVals <> "No Raster") Then sPixelVals = sPixelVals & ")"
        Exit For
      End If
    Next i

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Read_Raster " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub



Public Sub zone(ByVal pFeatureLayer As IFeatureLayer, ByVal pRasterLayer As IRasterLayer)
    
  On Error GoTo ErrorHandler
  Dim pEnumLayer As IEnumLayer
  Dim pFeature As IFeature
  Dim pFeatureCursor As IFeatureCursor
  Dim pFeatureSelection As IFeatureSelection
  Dim pSelectionSet As ISelectionSet
  Dim pUID As IUID
  Dim pTable As ITable
  Dim pRow As IRow

    'Loop through the selected features per layer
    Set pFeatureSelection = pFeatureLayer 'QI
    Set pSelectionSet = pFeatureSelection.SelectionSet
    'Can use Nothing keyword if you don't want to draw them,
    'otherwise, the spatial reference might not match the Map's
    pSelectionSet.Search Nothing, False, pFeatureCursor
    Set pFeature = pFeatureCursor.NextFeature
    
    Do While Not pFeature Is Nothing

      If TypeOf pFeatureLayer Is IFeatureLayer Then
        'Get value of selected feature for specific field
        Dim Wat_Name As String
        Dim strField As String
        strField = "Watershed_Name"
        Wat_Name = pFeature.Value(pFeature.Fields.FindField(strField))
      
        'Create Filter
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.SubFields = strField
        pQueryFilter.WhereClause = strField & "=" & Wat_Name

        Dim pGeoDS As IGeoDataset
      
        'Pick the zone field
        Set pGeoDS = pFeatureLayer.FeatureClass
        Dim pFDesc As IFeatureClassDescriptor
        Set pFDesc = New FeatureClassDescriptor
        'strField = InputBox("Enter the zone field:")
        'pFDesc.Create pGeoDs, Nothing, strField
        pFDesc.Create pGeoDS, pQueryFilter, strField
        Set pGeoDS = pFDesc
      Else
        If TypeOf pFeatureLayer Is IRasterLayer Then
          Dim pRLayer As IRasterLayer
          Set pRLayer = pFeatureLayer
          Set pGeoDS = pRLayer.Raster
          Dim pRDesc As IRasterDescriptor
          Set pRDesc = New RasterDescriptor
          'strField = InputBox("Enter the zone field:")
          pRDesc.Create pGeoDS, pQueryFilter, strField
          Set pGeoDS = pRDesc
        Else
          MsgBox "Exit Error"
          Exit Sub
        End If
      End If
      
      
      
    ''''''''''''''''''''''''''''''''''''''Get raster values within selected feature or zone
    
    Dim pLayer1 As ILayer
    Set pLayer1 = pRasterLayer
    Dim pGeoDs1 As IGeoDataset
    Dim pRLayer1 As IRasterLayer
    Set pRLayer1 = pLayer1
    Set pGeoDs1 = pRLayer1.Raster
    
    'Create a Spatial operator
    Dim pZoneOp As IZonalOp
    Set pZoneOp = New RasterZonalOp

    'Set output workspace
    Dim pEnv As IRasterAnalysisEnvironment
    Set pEnv = pZoneOp
    Dim pWs As IWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Set pWs = pWSF.OpenFromFile(gWorkingfolder, 0)
    Set pEnv.OutWorkspace = pWs


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Clip vegetation layer to selected watershed and obtain veg class percentages
    ' Create the RasterExtractionOp object
    Dim pExtractionOp As IExtractionOp
    Set pExtractionOp = New RasterExtractionOp

    'Perform the selection
    pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    Dim pPolygon As IPolygon
    Set pPolygon = pFeature.Shape

    ' Call the method
    Dim pOutputDataset As IGeoDataset
    Set pOutputDataset = pExtractionOp.Polygon(pGeoDs1, pPolygon, True)

    'Band stats
    Dim pBand As IRasterBand
    Dim pBandCol As IRasterBandCollection
    Dim pRasBand As IRasterBand
    Dim pStats As IRasterStatistics

    Set pBandCol = pOutputDataset
    Set pBand = pBandCol.Item(0)
    Set pRasBand = pBand

    ' Get the raster table
    Dim ExistTable As Boolean
    pBand.HasTable ExistTable
    If ExistTable = False Then
      Exit Sub
    End If
    Set pTable = pRasBand.AttributeTable

    'Add the table into ArcMap
    'Dim pTWindow As ITableWindow
    'Set pTWindow = New TableWindow
    'Set pTWindow.Table = pTable
    'Set pTWindow.Application = Application
    'pTWindow.Show True
    
    ' Get field index
    Dim FieldIndex As Integer, i As Integer
    FieldIndex = pTable.FindField("Count")
    
    'Sum all pixels counts per class
    Dim ValueResult As Variant
    Dim pSum As Long
    Dim pSumT As Long
    pSum = 0 'pixel count for specified veg class
    pSumT = 0 'pixel count: sum for all veg classes combined
    For i = 0 To 8
      Set pRow = pTable.GetRow(i)
      pSum = pRow.Value(FieldIndex)
      pSumT = pSum + pSumT
    Next i
    
    i = 0
    For i = 0 To 8
      Set pRow = pTable.GetRow(i)
      'Calculates percent coverage per class
      Select Case i
        Case 0 'Coniferous Wood
          MsgBox "Coniferous Wood(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 1 'Deciduous Wood
          MsgBox "Deciduous Weed(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 2 'Bushes
          MsgBox "Bushes(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 3 'Grassland_0-10
          MsgBox "Grassland_0-10(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 4 'Grassland_10-30
          MsgBox "Grassland_10-30(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 5 'Grassland_30-60
          MsgBox "Grassland_30-60(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 6 'Grassland_60-100
          MsgBox "Grassland_60-100(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 7 'Wetland
          MsgBox "Wetland = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        Case 8 'Dam/Basin
          MsgBox "Dam/Basin(Acres/Percent) = " & ((pRow.Value(FieldIndex) / pSumT) * 100)
        End Select
      Next i  'End of vegetation loop
    
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Set pFeature = pFeatureCursor.NextFeature
    Loop

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Zone " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
  
End Sub


'******************************************************************************
'Subroutine: ConvertFeatureToRaster
'Purpose:    Converts feature class to raster dataset. Required parameters
'            include featureclass, name of the field used for conversion,
'            name of the raster file name. Returns a RasterDataset.
'******************************************************************************
Public Function ConvertFeatureToRaster(pFeatureClass As IFeatureClass, pFieldName As String, pFileName As String, pQueryFilter As IQueryFilter) As IRasterDataset
   
On Error GoTo ShowError:
    
    Dim pDEMRLayer As IRasterLayer
    Set pDEMRLayer = GetInputFeatureLayer(gDEMdata)
    
    'Create a workspace
    Dim pWSF As IWorkspaceFactory
    Dim pWs As IWorkspace
    Set pWSF = New RasterWorkspaceFactory
    Set pWs = pWSF.OpenFromFile(gWorkingfolder, 0)
    
    'Select all features of the feature class
    Dim pSelectionSet As ISelectionSet
    ' Use the query filter to select features from STREAM feature layer
    Set pSelectionSet = pFeatureClass.Select(pQueryFilter, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
    
    ' Define the featureclassdescriptor
    Dim pGeoDataDescriptor As IFeatureClassDescriptor
    Set pGeoDataDescriptor = New FeatureClassDescriptor
    ' Get the selection set
    pGeoDataDescriptor.CreateFromSelectionSet pSelectionSet, Nothing, pFieldName
    
    Dim pGeoDS As IGeoDataset
    Set pGeoDS = pGeoDataDescriptor
    ' Delete Old Files
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    
    Dim pRasterPropsDEM As IRasterProps
    If Not pDEMRLayer Is Nothing Then
        Set pRasterPropsDEM = pDEMRLayer.Raster
    End If
    
    '*** Create the conversion object
    Dim pConvert As IConversionOp
    Set pConvert = New RasterConversionOp
    
    'Get the conversion environment
    Dim pEnv As IRasterAnalysisEnvironment
    Dim pCellSize As Double
    Set pEnv = pConvert
    If Not pDEMRLayer Is Nothing Then
        Set pEnv.OutSpatialReference = pRasterPropsDEM.SpatialReference
        pEnv.SetExtent esriRasterEnvValue, pRasterPropsDEM.Extent
        pCellSize = pRasterPropsDEM.MeanCellSize.x
        pEnv.SetCellSize esriRasterEnvValue, pCellSize
        Set pEnv.Mask = pDEMRLayer.Raster
    End If
    
    'Create a new raster dataset
    Dim pConRaster As IRasterDataset
    Set pConRaster = pConvert.ToRasterDataset(pGeoDS, "GRID", pWs, pFileName)
    
    'Return the value
    Set ConvertFeatureToRaster = pConRaster
    GoTo Cleanup

ShowError:
    MsgBox "ConvertFeatureToRaster: " & Err.Description
Cleanup:
    Set pDEMRLayer = Nothing
    Set pWSF = Nothing
    Set pWs = Nothing
    Set pSelectionSet = Nothing
    Set pGeoDataDescriptor = Nothing
    Set pGeoDS = Nothing
    Set fsObj = Nothing
    Set pRasterPropsDEM = Nothing
    Set pConvert = Nothing
    Set pEnv = Nothing
    Set pConRaster = Nothing
End Function


Public Function ConvertRastertoFeature(ByVal strWorkspace As String, ByVal strRaster As String, ByVal pFilter As Boolean, ByVal pParse As Boolean) As IFeatureClass
    
    On Error GoTo ErrorHandler
    'Open the grid
    Dim pRasterWspFact As IWorkspaceFactory
    Set pRasterWspFact = New RasterWorkspaceFactory
    Dim pRasterWorksp As IRasterWorkspace
    Set pRasterWorksp = pRasterWspFact.OpenFromFile(strWorkspace, 0)
    
    Dim pRasterDataset As IRasterDataset
    Set pRasterDataset = pRasterWorksp.OpenRasterDataset(strRaster)
    
    ' Filter the Raster for Flow Criteria.....
    If pFilter Then
        Dim pRaster As IRaster
        Set pRaster = pRasterDataset.CreateDefaultRaster

        Dim strExp As String
        strExp = Parse_Expression(gDACriteria, pParse)

        gMapAlgebraOp.BindRaster pRaster, "Raster_Con"
        Dim pOutputRaster As IRaster
        Set pOutputRaster = gMapAlgebraOp.Execute("con([Raster_Con] " & strExp & ", 1,0)")
        gMapAlgebraOp.UnbindRaster "Raster_Con"

        'Write it to the disk
        WriteRasterDatasetToDisk pOutputRaster, "Raster_Con"
        Set pRasterDataset = pRasterWorksp.OpenRasterDataset("Raster_Con")
    End If
    
    'now convert the raster dataset to a feature dataset
    'start by making a new, empty shapefile
    Dim pFeatWspFact As IWorkspaceFactory
    Set pFeatWspFact = New ShapefileWorkspaceFactory
    Dim pFeatWsp As IFeatureWorkspace
    Set pFeatWsp = pFeatWspFact.OpenFromFile(strWorkspace, 0)
    
    'make an IConversionOp object
    Dim pConvOp As IConversionOp
    Set pConvOp = New RasterConversionOp
    'Delete the shape file if exists.......
    Call Delete_Dataset_ST(strWorkspace, strRaster & "_Ras")
    
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = pConvOp.RasterDataToPolygonFeatureData(pRasterDataset, pFeatWsp, strRaster & "_Ras", True)
    
    'Delete the shape file if exists.......
    Call Delete_Dataset_ST(gWorkingfolder, strRaster & "_Ras")
    Dim pDataset As IDataset
    Set pDataset = pFeatClass
    If pDataset.CanCopy Then
        Dim pWkspace As IWorkspace
        Set pWkspace = GetWorkspace(gWorkingfolder)
        pDataset.Copy strRaster & "_Ras", pWkspace
        Set pFeatClass = OpenShapeFile(gWorkingfolder, strRaster & "_Ras")
    End If
    
    Set ConvertRastertoFeature = pFeatClass

Cleanup:

Set pConvOp = Nothing
Set pDataset = Nothing
Set pFeatClass = Nothing
Set pFeatWsp = Nothing
Set pWkspace = Nothing
Set pFeatWspFact = Nothing
Set pOutputRaster = Nothing
Set pRaster = Nothing
Set pRasterDataset = Nothing
Set pRasterWorksp = Nothing
Set pRasterWspFact = Nothing


  Exit Function
ErrorHandler:
  HandleError True, "ConvertRastertoFeature " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function


Public Function ConditionalOperationGrid(sPath As String, sFileName As String, _
                        dblCompareNumber As Double, dblNbLargerThan As Double, dblNbSmallerThan As Double, _
                        pRasterInputDataset As IGeoDataset) As IGeoDataset
    
    '~~~~~~~~~PREPARE WORKSPACE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Open raster workspace
    Dim pWSF As IWorkspaceFactory2
    Set pWSF = New RasterWorkspaceFactory
    Dim pRW2 As IRasterWorkspace2
    Set pRW2 = pWSF.OpenFromFile(sPath, 0)
    
    ' Calls function to open a raster dataset from disk
    
    Dim pRasterDataset As IRasterDataset
    Set pRasterDataset = New RasterDataset
    Set pRasterDataset = pRW2.OpenRasterDataset(sFileName)
    Dim pExtentGeodataset As IGeoDataset
    Set pExtentGeodataset = pRasterDataset 'QI
    
    Dim pRaster As IRaster
    Set pRaster = pRasterDataset.CreateDefaultRaster
    Dim pRasterAnalProp As IRasterAnalysisProps
    Set pRasterAnalProp = pRaster 'QI
    
    'set the raster analysis environment of the object as new default
    'Environment
    
    Dim pRAE As IRasterAnalysisEnvironment
    Set pRAE = New RasterAnalysis
    pRAE.SetExtent esriRasterEnvValue, pExtentGeodataset.Extent
    pRAE.SetCellSize esriRasterEnvValue, pRasterAnalProp.PixelWidth
    pRAE.SetAsNewDefaultEnvironment
    
    
    '~~~~~~~~CREATE NEW RASTER OBJECTS~~~~~~~~~~
     'Create the RasterConditionalOp object
    Dim pConditionalOp As IConditionalOp
    Set pConditionalOp = New RasterConditionalOp
    
    'the greaterthan method can not accept an integer, only another geodataset
    'must create a raster that has each cell populated with the value
    Dim pMakerOp As IRasterMakerOp
    Set pMakerOp = New RasterMakerOp
    
    'set the comparison raster to the constant value required
    Dim pMSL As IRaster
    Set pMSL = pMakerOp.MakeConstant(dblCompareNumber, True)
    
    'Set the TrueRaster with a constant of 1
    Dim pTrueRaster As IRaster
    Set pTrueRaster = pMakerOp.MakeConstant(dblNbLargerThan, True)
    
    'Set the TrueRaster with a constant of 0
    Dim pTrueRaster2 As IRaster
    Set pTrueRaster2 = pMakerOp.MakeConstant(dblNbSmallerThan, True)
    
    '~~~~~~~~COMPARE and REPLACE~~~~~~~~~~
    'set up for the logical operator greaterthan method
    Dim pLogicalOp As ILogicalOp
    Set pLogicalOp = New RasterMathOps
    
    Dim pInputDataset1 As IGeoDataset
    Set pInputDataset1 = pRasterInputDataset
    
    Dim pInputDataset2 As IGeoDataset
    Set pInputDataset2 = pMSL
    
    Dim pCondRaster As IGeoDataset
    Set pCondRaster = pLogicalOp.GreaterThan(pInputDataset1, pInputDataset2)
    
    'set up for the logical operator smallerthan method
    Dim pLogicalOp2 As ILogicalOp
    Set pLogicalOp2 = New RasterMathOps
    
    Dim pCondRaster2 As IGeoDataset
    Set pCondRaster2 = pLogicalOp2.LessThanEqual(pInputDataset1, pInputDataset2)
    
    
    
    'Will take a look at the pCondRaster and look for values of 1 (everything
    'over 20), will match it up with values
    'in pTrueRaster
    
    Dim pOutputDataset As IGeoDataset
    Dim pOutputDataset2 As IGeoDataset
    Set pOutputDataset = pConditionalOp.Con(pCondRaster, pTrueRaster)
    Set pOutputDataset2 = pConditionalOp.Con(pCondRaster2, pTrueRaster2)

End Function


