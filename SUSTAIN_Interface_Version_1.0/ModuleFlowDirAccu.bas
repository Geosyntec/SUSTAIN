Attribute VB_Name = "ModuleFlowDirAccu"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleFlowDirAccu
'   Purpose:     This module contains functions pertaining to the site hydrology.
'                It fills raw dem, burns dem to the stream, generates flow direction
'                and flow accumulation. It saves all generates rasters to disk.
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 08/19/2004 - mira chokshi
'
'******************************************************************************

Option Explicit
Option Base 0

'******************************************************************************
'Subroutine: RunSTREAMAgreeDEMForFlowDirAndAccu
'Author:     Mira Chokshi
'Purpose:    This subroutine manages creation of burned dem, fills raw dem,
'            uses raw dem to call function to create flow direction & flow
'            accumulation.
'******************************************************************************
Public Sub RunSTREAMAgreeDEMForFlowDirAndAccu()
 
 On Error GoTo ShowError
    'define steps number
    Dim nBurnSteps As Integer
    nBurnSteps = 2  'CInt(strBurnSteps)
    If nBurnSteps < 0 Or nBurnSteps > 3 Then
        Err.description = "Invalid burn steps value. Should be an integer in the range from 0 to 3."
        GoTo ShowError
    End If
    'get STREAM feature layer
    Dim pSTREAMFeatureLayer As IFeatureLayer
    Dim streamlayerName As String
    streamlayerName = gLayerNameDictionary.Item("STREAM")
    'Set pSTREAMFeatureLayer = GetInputFeatureLayer(gStreamLayer)
    Set pSTREAMFeatureLayer = GetInputFeatureLayer(streamlayerName)
    If pSTREAMFeatureLayer Is Nothing Then
        Err.description = "STREAM feature layer not found. "
        GoTo ShowError
    End If
    Dim pSTREAMFeatureClass As IFeatureClass
    Set pSTREAMFeatureClass = pSTREAMFeatureLayer.FeatureClass
    'get dem raster layer
    Dim pDEMRasterLayer As IRasterLayer
    Dim demlayerName As String
    demlayerName = gLayerNameDictionary.Item("DEM")
    'Set pDEMRasterLayer = GetInputRasterLayer(gDEMLayer)
    Set pDEMRasterLayer = GetInputRasterLayer(demlayerName)
    If pDEMRasterLayer Is Nothing Then
        Err.description = "Cannot find raster layer: DEM."
        GoTo ShowError
    End If
    
    'define temp workspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pRWS As IRasterWorkspace2
    Set pRWS = pWSF.OpenFromFile(gMapTempFolder, 0)
    'open burned dem from disk, create it if not found
    Dim pBurnDEMRaster As IRaster
    Set pBurnDEMRaster = OpenRasterDatasetFromDisk("burn_dem")
    If (pBurnDEMRaster Is Nothing) Then
        gAlgebraOp.BindRaster pDEMRasterLayer.Raster, "R"
        Set pBurnDEMRaster = gAlgebraOp.Execute("[R]")
        gAlgebraOp.UnbindRaster "R"
        Dim pRasterProps As IRasterProps
        Set pRasterProps = pBurnDEMRaster
        'call subroutine to create burned dem raster
        Call BurnReachInDem(pSTREAMFeatureClass, pBurnDEMRaster, nBurnSteps, 3 * pRasterProps.MeanCellSize.X)
    End If
    'get stream0 raster
    Dim pStreamDS As IRasterDataset
    Set pStreamDS = pRWS.OpenRasterDataset("stream0")
    Dim pStreamRaster As IRaster
    Set pStreamRaster = pStreamDS.CreateDefaultRaster
    'fill raw dem
    Dim pFillDEMRaster As IRaster
    Set pFillDEMRaster = FillRawDEM(pBurnDEMRaster)
    'Call the subroutine to create flow direction and flow accumulation
    CreateFlowDirectionAndAccumulation_Tilebased pFillDEMRaster, pStreamRaster
    GoTo CleanUp
    
ShowError:
    MsgBox "RunSTREAMAgreeDEMForFlowDirAndAccu: " & Err.description
CleanUp:
    Set pSTREAMFeatureLayer = Nothing
    Set pSTREAMFeatureClass = Nothing
    Set pDEMRasterLayer = Nothing
    Set pWSF = Nothing
    Set pRWS = Nothing
    Set pBurnDEMRaster = Nothing
    Set pRasterProps = Nothing
    Set pStreamDS = Nothing
    Set pStreamRaster = Nothing
    Set pFillDEMRaster = Nothing
End Sub

'******************************************************************************
'Subroutine: BurnReachInDem
'Author:     Haihong Yang
'Purpose:    This module iterates to generate a reach file burned into dem.
'******************************************************************************
Private Sub BurnReachInDem(pFC As IFeatureClass, pRaster As IRaster, nBurnSteps As Integer, nBurnDepth As Double)

On Error GoTo ShowError
    'get temp workspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pRWS As IRasterWorkspace2
    Set pRWS = pWSF.OpenFromFile(gMapTempFolder, 0)
    'initialize status progress bar
    Dim pStatusBar As esriSystem.IStatusBar
    Set pStatusBar = gApplication.StatusBar
    Dim pStepProgressor As IStepProgressor
    Set pStepProgressor = pStatusBar.ProgressBar
    pStepProgressor.Show
    'define burn size
    Dim nNbSize As Integer
    nNbSize = nBurnSteps * 4 + 1
    pStepProgressor.Message = "Calculating minimum elevation in a neighborhood " & nNbSize & " x " & nNbSize & " ..."
    ' Create the neighborhood object
    Dim pNbr As IRasterNeighborhood
    Set pNbr = New RasterNeighborhood
    pNbr.SetRectangle nNbSize, nNbSize, esriUnitsCells
    'compute focal statistics
    Dim pRasterMinDem As IRaster
    Set pRasterMinDem = gNeighborhoodOp.FocalStatistics(pRaster, esriGeoAnalysisStatsMinimum, pNbr, True)
    pStepProgressor.Message = "Burn in stream itself ..."
    'create stream0 raster
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pRasterDS As IRasterDataset
    If (fso.FolderExists(gMapTempFolder & "\stream0")) Then
        Set pRasterDS = pRWS.OpenRasterDataset("stream0")
    Else
        If pFC.FindField("FID") >= 0 Then
            Set pRasterDS = ConvertFeatureToRaster(pFC, "FID", "stream0", Nothing)
        Else
            Set pRasterDS = ConvertFeatureToRaster(pFC, "OBJECTID", "stream0", Nothing)
        End If
    End If
    'get raster buf
    Dim pRasterBuf As IRaster
    Set pRasterBuf = pRasterDS.CreateDefaultRaster
    gAlgebraOp.BindRaster pRaster, "dem"
    gAlgebraOp.BindRaster pRasterMinDem, "min_dem"
    gAlgebraOp.BindRaster pRasterBuf, "buf"
    Set pRaster = gAlgebraOp.Execute("Con(IsNull([buf]), [dem], [min_dem] - " & nBurnDepth & ")")
    gAlgebraOp.UnbindRaster "dem"
    gAlgebraOp.UnbindRaster "min_dem"
    gAlgebraOp.UnbindRaster "buf"
    'error checking
    If nBurnSteps = 0 Then
        GoTo CleanUp
    End If
    'define variables for processing
    Dim pRasterProps As IRasterProps
    Set pRasterProps = pRaster
    Dim pClone As IClone
    Set pClone = pFC.Fields
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = pClone.Clone
    Dim FieldCount As Integer
    Dim pField As esriGeoDatabase.IField
    Dim pGeometryDefEdit As IGeometryDefEdit
    For FieldCount = 0 To pFields.FieldCount - 1  'skip OID and geometry
        Set pField = pFields.Field(FieldCount)
        If pField.Type = esriFieldTypeGeometry Then
            Set pGeometryDefEdit = pField.GeometryDef
            pGeometryDefEdit.GeometryType = esriGeometryPolygon
            Exit For
        End If
    Next FieldCount
    Dim pGeoColl As IGeometryCollection
    Set pGeoColl = New GeometryBag
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFC.Search(Nothing, False)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Do Until pFeature Is Nothing
        pGeoColl.AddGeometry pFeature.ShapeCopy
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    If pGeoColl.GeometryCount = 0 Then
        Err.description = "No feature could be used for burning-in operation."
        GoTo ShowError
    End If
    
    Dim pTopoOp As ITopologicalOperator
    Set pTopoOp = pGeoColl
    Dim pFWSF As IWorkspaceFactory
    Set pFWSF = New ShapefileWorkspaceFactory
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pFWSF.OpenFromFile(gMapTempFolder, 0)
    
    Dim pNewFC As IFeatureClass
    Dim pInsertFeatureCursor As IFeatureCursor
    Dim pInsertFeatureBuffer As IFeatureBuffer
    Dim pGeometry As IGeometry
    Dim pGeometryType As esriGeometryType
    Dim pPolygon As IPolygon
    Dim pTopoOp1 As ITopologicalOperator
  
    Dim nCurBufDis As Integer
    Dim nGeoIndex As Long
    'continue iterating
    For nCurBufDis = 1 To nBurnSteps
        pStepProgressor.Message = "Burn in stream with a buffer of " & (2 * nCurBufDis) & " pixels ..."
        Dim pstrbufnm As String
        pstrbufnm = CreateUniqueTableName(gMapTempFolder, "strbuf" & nCurBufDis)
        Set pNewFC = pFWS.CreateFeatureClass(pstrbufnm, pFields, Nothing, Nothing, esriFTSimple, pFC.ShapeFieldName, "")
        Set pInsertFeatureCursor = pNewFC.Insert(True)
        Set pInsertFeatureBuffer = pNewFC.CreateFeatureBuffer
        Set pGeometry = pTopoOp.Buffer(2 * nCurBufDis * pRasterProps.MeanCellSize.X)
        pGeometryType = pGeometry.GeometryType
        If pGeometryType = esriGeometryBag Then
            Set pGeoColl = pGeometry
            For nGeoIndex = 0 To pGeoColl.GeometryCount - 1
                Set pInsertFeatureBuffer.Shape = pGeoColl.Geometry(nGeoIndex)
                pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
            Next
        Else
            Set pPolygon = pGeometry
            Set pInsertFeatureBuffer.Shape = pPolygon
            pInsertFeatureCursor.InsertFeature pInsertFeatureBuffer
        End If
        pInsertFeatureCursor.Flush
        'If pNewFC.FindField("FID") > 0 Then
        '    Set pRasterDS = ConvertFeatureToRaster(pNewFC, "FID", "stream", Nothing)
        'Else
        '    Set pRasterDS = ConvertFeatureToRaster(pNewFC, "OBJECTID", "stream", Nothing)
        'End If
        Set pRasterDS = ConvertFeatureToRaster(pNewFC, "FID", "stream", Nothing)
        Set pRasterBuf = pRasterDS.CreateDefaultRaster
        gAlgebraOp.BindRaster pRaster, "dem"
        gAlgebraOp.BindRaster pRasterBuf, "buf"
        Set pRaster = gAlgebraOp.Execute("Con(IsNull([buf]), [dem], [dem] - " & nBurnDepth & ")")
        gAlgebraOp.UnbindRaster "dem"
        gAlgebraOp.UnbindRaster "buf"
    Next
       
    '** Save the burned dem
    WriteRasterDatasetToDisk pRaster, "burn_dem"
        
    GoTo CleanUp
ShowError:
    MsgBox "BurnReachInDem: " & Err.description
CleanUp:
    pStepProgressor.Hide
    Set pWSF = Nothing
    Set pRWS = Nothing
    Set pStatusBar = Nothing
    Set pStepProgressor = Nothing
    Set pNbr = Nothing
    Set pRasterMinDem = Nothing
    Set fso = Nothing
    Set pRasterDS = Nothing
    Set pRasterBuf = Nothing
    Set pRasterProps = Nothing
    Set pClone = Nothing
    Set pFields = Nothing
    Set pField = Nothing
    Set pGeometryDefEdit = Nothing
    Set pGeoColl = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pTopoOp = Nothing
    Set pFWSF = Nothing
    Set pFWS = Nothing
    Set pNewFC = Nothing
    Set pInsertFeatureCursor = Nothing
    Set pInsertFeatureBuffer = Nothing
    Set pGeometry = Nothing
    Set pPolygon = Nothing
    Set pTopoOp1 = Nothing

End Sub


'******************************************************************************
'Subroutine: FillRawDEM
'Author:     Mira Chokshi
'Purpose:    This module uses hydrology operator to fill raw dem. Input parameter
'            is raw dem and output is the filled dem
'******************************************************************************
Private Function FillRawDEM(pDEMRaster As IRaster) As IRaster
On Error GoTo ShowError
    'get filled dem from disk
    Dim pFillDEMRaster As IRaster
    Set pFillDEMRaster = OpenRasterDatasetFromDisk("FillDEM")
    If Not (pFillDEMRaster Is Nothing) Then
        Set FillRawDEM = pFillDEMRaster
        GoTo CleanUp
    End If
    'Fill the raw dem
    Set pFillDEMRaster = gHydrologyOp.Fill(pDEMRaster)
    'Write it to the disk
    WriteRasterDatasetToDisk pFillDEMRaster, "FillDEM"
    Set FillRawDEM = pFillDEMRaster
    GoTo CleanUp
    
ShowError:
    MsgBox "FillRawDEM: " & Err.description
CleanUp:
    Set pFillDEMRaster = Nothing
End Function

'******************************************************************************
'Subroutine: CreateFlowDirectionAndAccumulation
'Author:     Mira Chokshi
'Purpose:    This module creates flow direction using filled dem. It uses an
'            cell based computation to compute flow accumulation. Its a iterative
'            process, trying to align the streams with flow accumulation.
'            It saves both the rasters to the disk.
'******************************************************************************
Private Sub CreateFlowDirectionAndAccumulation_Tilebased(pFilledDEMRaster As IRaster, pRasterStream As IRaster)
On Error GoTo ShowError
    
    Dim pStatusBar As esriSystem.IStatusBar
    Set pStatusBar = gApplication.StatusBar
    Dim pStepProgressor As IStepProgressor
    Set pStepProgressor = pStatusBar.ProgressBar
    pStepProgressor.Show
       
    Dim pRasterFlowDir As IRaster
    Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
    If (pRasterFlowDir Is Nothing) Then
        pStepProgressor.Message = "Calculating flow direction ..."
        Set pRasterFlowDir = gHydrologyOp.FlowDirection(pFilledDEMRaster, False, True)
        'Save Flow Direction Raster DataSet
        WriteRasterDatasetToDisk pRasterFlowDir, "FlowDir"
        Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
    End If
    
    Dim pRasterFlowAccu As IRaster
    Set pRasterFlowAccu = OpenRasterDatasetFromDisk("FlowAccu")
    If (Not pRasterFlowAccu Is Nothing) Then
        Exit Sub
    End If
    
    pStepProgressor.Message = "Calculating flow accumulation ..."
    Set pRasterFlowAccu = gHydrologyOp.FlowAccumulation(pRasterFlowDir)
    
    pStepProgressor.Message = "Reclassifying stream raster ..."
    gAlgebraOp.BindRaster pRasterStream, "stream"
    Set pRasterStream = gAlgebraOp.Execute("Con(IsNull([stream]), 0, 1)")
    gAlgebraOp.UnbindRaster "stream"

    Dim pPixelBlockStream As IPixelBlock3
    Dim pPixelBlockFlowDir As IPixelBlock3
    Dim pPixelBlockFlowAccu As IPixelBlock3

    Dim pRasterPropStream As IRasterProps
    Dim pRasterPropFlowDir As IRasterProps
    Dim pRasterPropFlowAccu As IRasterProps

    Dim vPixelDataStream As Variant
    Dim vPixelDataFlowDir As Variant
    Dim vPixelDataFlowAccu As Variant

    Dim pSize As IPnt
    Dim pOrigin As IPnt
    
    ' get raster properties
    Set pRasterPropStream = pRasterStream
    Set pRasterPropFlowDir = pRasterFlowDir
    Set pRasterPropFlowAccu = pRasterFlowAccu
    
    ' get vb supported pixel type
    pRasterPropStream.PixelType = GetVBSupportedPixelType(pRasterPropStream.PixelType)
    pRasterPropFlowDir.PixelType = GetVBSupportedPixelType(pRasterPropFlowDir.PixelType)
    pRasterPropFlowAccu.PixelType = GetVBSupportedPixelType(pRasterPropFlowAccu.PixelType)
    
    If pRasterPropStream.Width <> pRasterPropFlowDir.Width _
        Or pRasterPropFlowDir.Width <> pRasterPropFlowAccu.Width _
        Or pRasterPropStream.Height <> pRasterPropFlowDir.Height _
        Or pRasterPropFlowDir.Height <> pRasterPropFlowAccu.Height Then
        Err.description = "Extent mismatch"
        GoTo ShowError
    End If
    
    Dim pBandCol As IRasterBandCollection
    Set pBandCol = pRasterFlowAccu
    
    Dim pBand As IRasterBand
    Set pBand = pBandCol.Item(0)
    
    Dim pRawPixel As IRawPixels
    Set pRawPixel = pBand
    
    pStepProgressor.Message = "Reading raster data ..."

    ' create a DblPnt to hold the PixelBlock size
    Set pSize = New DblPnt
    Set pOrigin = New DblPnt

    Dim lBlockSize As Long
    Dim lTileNum As Long
    Dim lTile As Long
    Dim lStartRow As Long
    
    ' *************************************************************
    ' Split the Raster into Tiles for reading......
    ' *************************************************************
    
    lBlockSize = CLng(1048576 / pRasterPropStream.Width)
    If lBlockSize < 1 Then lBlockSize = 1
    lTileNum = CLng(pRasterPropStream.Height / lBlockSize)
    If lTileNum * lBlockSize < pRasterPropStream.Height Then lTileNum = lTileNum + 1
     
    Set pOrigin = New DblPnt
    For lTile = 0 To lTileNum - 1
        lStartRow = lTile * lBlockSize
        If lStartRow + lBlockSize > pRasterPropStream.Height Then lBlockSize = pRasterPropStream.Height - lStartRow
        pOrigin.SetCoords 0, lStartRow
        pSize.SetCoords pRasterPropStream.Width, lBlockSize
   
        Set pPixelBlockFlowAccu = pRasterFlowAccu.CreatePixelBlock(pSize)
        pRasterFlowAccu.Read pOrigin, pPixelBlockFlowAccu
        vPixelDataFlowAccu = pPixelBlockFlowAccu.PixelDataByRef(0)
        Set pPixelBlockStream = pRasterStream.CreatePixelBlock(pSize)
        pRasterStream.Read pOrigin, pPixelBlockStream
        vPixelDataStream = pPixelBlockStream.PixelDataByRef(0)
        Set pPixelBlockFlowDir = pRasterFlowDir.CreatePixelBlock(pSize)
        pRasterFlowDir.Read pOrigin, pPixelBlockFlowDir
        vPixelDataFlowDir = pPixelBlockFlowDir.PixelDataByRef(0)
        
        Dim noDataValueFlowAccu As Double
        noDataValueFlowAccu = pRasterPropFlowAccu.NoDataValue(0)
        
        Dim iCol As Integer
        Dim iRow As Integer
        Dim cCol As Integer
        Dim cRow As Integer
        Dim pValueFlowDir As Integer
        Dim nProcessedCount As Long
        Dim bCalc As Boolean
        Dim steps As Integer
        Dim vCount As Integer
        Dim vMeanValue As Double
        Dim loopCount As Integer
        loopCount = 1
        
        Dim strCount As String
        Dim dInterval As Double
        dInterval = pRasterPropStream.Width / 100
    
        pStepProgressor.MinRange = 0
        pStepProgressor.MaxRange = pRasterPropStream.Width - 1
        pStepProgressor.StepValue = dInterval
        nProcessedCount = 0
      
        Do
            If nProcessedCount > 0 Then
                strCount = "Last round " & nProcessedCount & " cells were processed. "
            Else
                strCount = ""
            End If
            
            pStepProgressor.Message = strCount & "Filling gaps for flow accumulation (round " & loopCount & " ) ..."
            nProcessedCount = 0
            
            For iCol = 0 To pRasterPropStream.Width - 1
                pStepProgressor.Position = iCol
                pStepProgressor.Step
                
                For iRow = 0 To lBlockSize - 1
                    ' Process only stream cells
                    If vPixelDataStream(iCol, iRow) = 1 Then
                        ' Calculate average flow accumulation in a 3x3 cells neighborhood
                        vCount = 0
                        vMeanValue = 0
                        
                        For cCol = iCol - 1 To iCol + 1
                            For cRow = iRow - 1 To iRow + 1
                                If cCol >= 0 And cCol < pRasterPropStream.Width And cRow >= 0 And cRow < lBlockSize Then
                                    If vPixelDataFlowAccu(cCol, cRow) <> noDataValueFlowAccu Then
                                        vCount = vCount + 1
                                        vMeanValue = vMeanValue + vPixelDataFlowAccu(cCol, cRow)
                                    End If
                                End If
                            Next
                        Next
                        
                        ' If no valid cell present, skip this cell for next iteration
                        If vCount = 0 Then
                            vPixelDataStream(iCol, iRow) = 0
                        ' Otherwise, process this cell to fill up gaps
                        Else
                            ' Calculate average flow accumulation in the 3x3 neighborhood
                            vMeanValue = vMeanValue / vCount
                            
                            cCol = iCol
                            cRow = iRow
                            steps = 1
                            
                            Do
                                ' If the current flow accumulation is larger than the average value in neighborhood, done with this cell
                                If vPixelDataFlowAccu(iCol, iRow) >= vMeanValue Then
                                    Exit Do
                                End If
                                    
                                ' Otherwise, find a downstream cell iteratively until a flow accumulation value is large enough
                                bCalc = True
                                pValueFlowDir = vPixelDataFlowDir(cCol, cRow)
                                If pValueFlowDir = 64 Then
                                    cRow = cRow - 1
                                ElseIf pValueFlowDir = 128 Then
                                    cCol = cCol + 1
                                    cRow = cRow - 1
                                ElseIf pValueFlowDir = 1 Then
                                    cCol = cCol + 1
                                ElseIf pValueFlowDir = 2 Then
                                    cCol = cCol + 1
                                    cRow = cRow + 1
                                ElseIf pValueFlowDir = 4 Then
                                    cRow = cRow + 1
                                ElseIf pValueFlowDir = 8 Then
                                    cCol = cCol - 1
                                    cRow = cRow + 1
                                ElseIf pValueFlowDir = 16 Then
                                    cCol = cCol - 1
                                ElseIf pValueFlowDir = 32 Then
                                    cCol = cCol - 1
                                    cRow = cRow - 1
                                Else
                                    bCalc = False
                                End If
                
                                If bCalc And cCol >= 0 And cCol < pRasterPropStream.Width _
                                    And cRow >= 0 And cRow < lBlockSize Then
                                    If vPixelDataFlowAccu(cCol, cRow) <> noDataValueFlowAccu Then
                                        ' Assign the current flow accumulation to be the downstream cell flow accumulation lesser than the steps
                                        vPixelDataFlowAccu(iCol, iRow) = vPixelDataFlowAccu(cCol, cRow) - steps
                                        steps = steps + 1
                                        nProcessedCount = nProcessedCount + 1
                                    Else
                                        vPixelDataStream(iCol, iRow) = 0
                                        Exit Do
                                    End If
                                Else
                                    vPixelDataStream(iCol, iRow) = 0
                                    Exit Do
                                End If
                            Loop
                        End If
                    End If
                Next
            Next
        
            loopCount = loopCount + 1
        Loop While nProcessedCount > 0 And loopCount < 6
            
            pRawPixel.Write pOrigin, pPixelBlockFlowAccu ' Write the modified Pixels to the Raster.....
    Next
    
    ' Write back the final pixel block to stream grid
    pStepProgressor.Message = "Saving flow accumulation raster ..."
    pBand.ComputeStatsAndHist
    
    'Save Flow Accumulation Raster to Disk
    WriteRasterDatasetToDisk pRasterFlowAccu, "FlowAccu"

    GoTo CleanUp
ShowError:
    MsgBox "Tiled based Flow accumulation calculation error: " & Err.description, vbExclamation
CleanUp:
    pStepProgressor.Hide
    Set pStatusBar = Nothing
    Set pStepProgressor = Nothing
    Set pRasterFlowDir = Nothing
    Set pRasterFlowAccu = Nothing
    Set pPixelBlockStream = Nothing
    Set pPixelBlockFlowDir = Nothing
    Set pPixelBlockFlowAccu = Nothing
    Set pRasterPropStream = Nothing
    Set pRasterPropFlowDir = Nothing
    Set pRasterPropFlowAccu = Nothing
    Set vPixelDataStream = Nothing
    Set vPixelDataFlowDir = Nothing
    Set vPixelDataFlowAccu = Nothing
    Set pSize = Nothing
    Set pOrigin = Nothing
    Set pBandCol = Nothing
    Set pBand = Nothing
    Set pRawPixel = Nothing

End Sub

'******************************************************************************
'Subroutine: CreateFlowDirectionAndAccumulation
'Author:     Mira Chokshi
'Purpose:    This module creates flow direction using filled dem. It uses an
'            cell based computation to compute flow accumulation. Its a iterative
'            process, trying to align the streams with flow accumulation.
'            It saves both the rasters to the disk.
'******************************************************************************
Private Sub CreateFlowDirectionAndAccumulation(pFilledDEMRaster As IRaster, pRasterStream As IRaster)
On Error GoTo ShowError
    
    Dim pStatusBar As esriSystem.IStatusBar
    Set pStatusBar = gApplication.StatusBar
    Dim pStepProgressor As IStepProgressor
    Set pStepProgressor = pStatusBar.ProgressBar
    pStepProgressor.Show
   
    Dim pRasterFlowDir As IRaster
    Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
    If (pRasterFlowDir Is Nothing) Then
        pStepProgressor.Message = "Calculating flow direction ..."
        Set pRasterFlowDir = gHydrologyOp.FlowDirection(pFilledDEMRaster, False, True)
        'Save Flow Direction Raster DataSet
        WriteRasterDatasetToDisk pRasterFlowDir, "FlowDir"
    End If
    
    Dim pRasterFlowAccu As IRaster
    Set pRasterFlowAccu = OpenRasterDatasetFromDisk("FlowAccu")
    If (Not pRasterFlowAccu Is Nothing) Then
        Exit Sub
    End If
    
    pStepProgressor.Message = "Calculating flow accumulation ..."
    Set pRasterFlowAccu = gHydrologyOp.FlowAccumulation(pRasterFlowDir)
    
    pStepProgressor.Message = "Reclassifying stream raster ..."
    gAlgebraOp.BindRaster pRasterStream, "stream"
    Set pRasterStream = gAlgebraOp.Execute("Con(IsNull([stream]), 0, 1)")
    gAlgebraOp.UnbindRaster "stream"

    Dim pPixelBlockStream As IPixelBlock3
    Dim pPixelBlockFlowDir As IPixelBlock3
    Dim pPixelBlockFlowAccu As IPixelBlock3

    Dim pRasterPropStream As IRasterProps
    Dim pRasterPropFlowDir As IRasterProps
    Dim pRasterPropFlowAccu As IRasterProps

    Dim vPixelDataStream As Variant
    Dim vPixelDataFlowDir As Variant
    Dim vPixelDataFlowAccu As Variant

    Dim pSize As IPnt
    Dim pOrigin As IPnt

    ' get raster properties
    Set pRasterPropStream = pRasterStream
    Set pRasterPropFlowDir = pRasterFlowDir
    Set pRasterPropFlowAccu = pRasterFlowAccu

    ' get vb supported pixel type
    pRasterPropStream.PixelType = GetVBSupportedPixelType(pRasterPropStream.PixelType)
    pRasterPropFlowDir.PixelType = GetVBSupportedPixelType(pRasterPropFlowDir.PixelType)
    pRasterPropFlowAccu.PixelType = GetVBSupportedPixelType(pRasterPropFlowAccu.PixelType)

    If pRasterPropStream.Width <> pRasterPropFlowDir.Width _
        Or pRasterPropFlowDir.Width <> pRasterPropFlowAccu.Width _
        Or pRasterPropStream.Height <> pRasterPropFlowDir.Height _
        Or pRasterPropFlowDir.Height <> pRasterPropFlowAccu.Height Then
        Err.description = "Extent mismatch"
        GoTo ShowError
    End If
    
    pStepProgressor.Message = "Reading raster data ..."

    ' create a DblPnt to hold the PixelBlock size
    Set pSize = New DblPnt
    pSize.SetCoords pRasterPropStream.Width, pRasterPropStream.Height
    Set pPixelBlockStream = pRasterStream.CreatePixelBlock(pSize)
    Set pPixelBlockFlowDir = pRasterFlowDir.CreatePixelBlock(pSize)
    Set pPixelBlockFlowAccu = pRasterFlowAccu.CreatePixelBlock(pSize)

    Set pOrigin = New DblPnt
    pOrigin.SetCoords 0, 0

    pRasterStream.Read pOrigin, pPixelBlockStream
    vPixelDataStream = pPixelBlockStream.PixelDataByRef(0)
    pRasterFlowDir.Read pOrigin, pPixelBlockFlowDir
    vPixelDataFlowDir = pPixelBlockFlowDir.PixelDataByRef(0)
    pRasterFlowAccu.Read pOrigin, pPixelBlockFlowAccu
    vPixelDataFlowAccu = pPixelBlockFlowAccu.PixelDataByRef(0)

    Dim noDataValueFlowAccu As Double
    noDataValueFlowAccu = pRasterPropFlowAccu.NoDataValue ' (0)
    
    Dim iCol As Integer
    Dim iRow As Integer
    Dim cCol As Integer
    Dim cRow As Integer
    Dim pValueFlowDir As Integer
    Dim nProcessedCount As Long
    Dim bCalc As Boolean
    Dim steps As Integer
    Dim vCount As Integer
    Dim vMeanValue As Double
    Dim loopCount As Integer
    loopCount = 1
    
    Dim strCount As String
    Dim dInterval As Double
    dInterval = pRasterPropStream.Width / 100

    pStepProgressor.MinRange = 0
    pStepProgressor.MaxRange = pRasterPropStream.Width - 1
    pStepProgressor.StepValue = dInterval
    nProcessedCount = 0
  
    Do
        If nProcessedCount > 0 Then
            strCount = "Last round " & nProcessedCount & " cells were processed. "
        Else
            strCount = ""
        End If
        
        pStepProgressor.Message = strCount & "Filling gaps for flow accumulation (round " & loopCount & " ) ..."
        nProcessedCount = 0
        
        For iCol = 0 To pRasterPropStream.Width - 1
            pStepProgressor.Position = iCol
            pStepProgressor.Step
            
            For iRow = 0 To pRasterPropStream.Height - 1
                ' Process only stream cells
                If vPixelDataStream(iCol, iRow) = 1 Then
                    ' Calculate average flow accumulation in a 3x3 cells neighborhood
                    vCount = 0
                    vMeanValue = 0
                    
                    For cCol = iCol - 1 To iCol + 1
                        For cRow = iRow - 1 To iRow + 1
                            If cCol >= 0 And cCol < pRasterPropStream.Width And cRow >= 0 And cRow < pRasterPropStream.Height Then
                                If vPixelDataFlowAccu(cCol, cRow) <> noDataValueFlowAccu Then
                                    vCount = vCount + 1
                                    vMeanValue = vMeanValue + vPixelDataFlowAccu(cCol, cRow)
                                End If
                            End If
                        Next
                    Next
                    
                    ' If no valid cell present, skip this cell for next iteration
                    If vCount = 0 Then
                        vPixelDataStream(iCol, iRow) = 0
                    ' Otherwise, process this cell to fill up gaps
                    Else
                        ' Calculate average flow accumulation in the 3x3 neighborhood
                        vMeanValue = vMeanValue / vCount
                        
                        cCol = iCol
                        cRow = iRow
                        steps = 1
                        
                        Do
                            ' If the current flow accumulation is larger than the average value in neighborhood, done with this cell
                            If vPixelDataFlowAccu(iCol, iRow) >= vMeanValue Then
                                Exit Do
                            End If
                                
                            ' Otherwise, find a downstream cell iteratively until a flow accumulation value is large enough
                            bCalc = True
                            pValueFlowDir = vPixelDataFlowDir(cCol, cRow)
                            If pValueFlowDir = 64 Then
                                cRow = cRow - 1
                            ElseIf pValueFlowDir = 128 Then
                                cCol = cCol + 1
                                cRow = cRow - 1
                            ElseIf pValueFlowDir = 1 Then
                                cCol = cCol + 1
                            ElseIf pValueFlowDir = 2 Then
                                cCol = cCol + 1
                                cRow = cRow + 1
                            ElseIf pValueFlowDir = 4 Then
                                cRow = cRow + 1
                            ElseIf pValueFlowDir = 8 Then
                                cCol = cCol - 1
                                cRow = cRow + 1
                            ElseIf pValueFlowDir = 16 Then
                                cCol = cCol - 1
                            ElseIf pValueFlowDir = 32 Then
                                cCol = cCol - 1
                                cRow = cRow - 1
                            Else
                                bCalc = False
                            End If
            
                            If bCalc And cCol >= 0 And cCol < pRasterPropStream.Width _
                                And cRow >= 0 And cRow < pRasterPropStream.Height Then
                                If vPixelDataFlowAccu(cCol, cRow) <> noDataValueFlowAccu Then
                                    ' Assign the current flow accumulation to be the downstream cell flow accumulation lesser than the steps
                                    vPixelDataFlowAccu(iCol, iRow) = vPixelDataFlowAccu(cCol, cRow) - steps
                                    steps = steps + 1
                                    nProcessedCount = nProcessedCount + 1
                                Else
                                    vPixelDataStream(iCol, iRow) = 0
                                    Exit Do
                                End If
                            Else
                                vPixelDataStream(iCol, iRow) = 0
                                Exit Do
                            End If
                        Loop
                    End If
                End If
            Next
        Next
    
        loopCount = loopCount + 1
    Loop While nProcessedCount > 0 And loopCount < 6
    
    ' Write back the final pixel block to stream grid
    pStepProgressor.Message = "Saving flow accumulation raster ..."
    Dim pBandCol As IRasterBandCollection
    Set pBandCol = pRasterFlowAccu
    
    Dim pBand As IRasterBand
    Set pBand = pBandCol.Item(0)
    
    Dim pRawPixel As IRawPixels
    Set pRawPixel = pBand
    pRawPixel.Write pOrigin, pPixelBlockFlowAccu
    pBand.ComputeStatsAndHist
    
    'Save Flow Accumulation Raster to Disk
    WriteRasterDatasetToDisk pRasterFlowAccu, "FlowAccu"

    GoTo CleanUp
ShowError:
    MsgBox "Flow accumulation calculation error: " & Err.description, vbExclamation
CleanUp:
    pStepProgressor.Hide
    Set pStatusBar = Nothing
    Set pStepProgressor = Nothing
    Set pRasterFlowDir = Nothing
    Set pRasterFlowAccu = Nothing
    Set pPixelBlockStream = Nothing
    Set pPixelBlockFlowDir = Nothing
    Set pPixelBlockFlowAccu = Nothing
    Set pRasterPropStream = Nothing
    Set pRasterPropFlowDir = Nothing
    Set pRasterPropFlowAccu = Nothing
    Set vPixelDataStream = Nothing
    Set vPixelDataFlowDir = Nothing
    Set vPixelDataFlowAccu = Nothing
    Set pSize = Nothing
    Set pOrigin = Nothing
    Set pBandCol = Nothing
    Set pBand = Nothing
    Set pRawPixel = Nothing

End Sub


Sub RunFromFlowDir()
On Error GoTo ShowError
    
    Dim pNHDFeatureLayer As IFeatureLayer
    Set pNHDFeatureLayer = GetInputFeatureLayer("STREAM")
    If pNHDFeatureLayer Is Nothing Then
        Err.description = "Streams feature layer not found."
        GoTo ShowError
    End If
    
    Dim pNHDFeatureClass As IFeatureClass
    Set pNHDFeatureClass = pNHDFeatureLayer.FeatureClass
    
    Dim pBMPFeatureLayer As IFeatureLayer
    Set pBMPFeatureLayer = GetInputFeatureLayer("SnapPoints")
    Dim pBMPFeatureClass As IFeatureClass
    If Not pBMPFeatureLayer Is Nothing Then
        Set pBMPFeatureClass = pBMPFeatureLayer.FeatureClass
    End If
    
    Dim pVFSFeatureLayer As IFeatureLayer
    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFeatureClass As IFeatureClass
    If Not pVFSFeatureLayer Is Nothing Then
        Set pVFSFeatureClass = pVFSFeatureLayer.FeatureClass
    End If

    '** Show an error if both BMPs and VFS are not present
    If (pBMPFeatureLayer Is Nothing And pVFSFeatureLayer Is Nothing) Then
        Err.description = "BMPs and VFS feature layers not found."
        GoTo ShowError
    End If
    
    Dim pDEMRasterLayer As IRasterLayer
    Set pDEMRasterLayer = GetInputRasterLayer("DEM")
    If pDEMRasterLayer Is Nothing Then
        Err.description = "DEM raster layer not found."
        GoTo ShowError
    End If
           
    Dim pFlowDirRaster As IRaster
    Set pFlowDirRaster = OpenRasterDatasetFromDisk("FlowDir")
    If pFlowDirRaster Is Nothing Then
        Err.description = "Flow Direction raster not found in temp folder."
        GoTo ShowError
    End If

    Dim pDrainageRaster As IRaster
    Set pDrainageRaster = TraceBMPDrainageArea(pFlowDirRaster, pBMPFeatureClass, pVFSFeatureClass)
    
    Dim pDrainageRasterLayer As IRasterLayer
    Set pDrainageRasterLayer = New RasterLayer
    pDrainageRasterLayer.CreateFromRaster pDrainageRaster
    AddLayerToMap pDrainageRasterLayer, "SubWatershed"
    
    GoTo CleanUp
    
ShowError:
    MsgBox "Unsuccessul in finish this operation. " & Err.description
CleanUp:
    Set pNHDFeatureLayer = Nothing
    Set pNHDFeatureClass = Nothing
    Set pDEMRasterLayer = Nothing
    Set pFlowDirRaster = Nothing
    Set pDrainageRaster = Nothing
    Set pDrainageRasterLayer = Nothing
End Sub


Private Function TraceBMPDrainageArea(pRasterFlowDir As IRaster, pFCBMP As IFeatureClass, pFCVFS As IFeatureClass) As IRaster
On Error GoTo ShowError
    
    'MsgBox "From TraceBMPDrainageArea 1"
    Dim pStatusBar As esriSystem.IStatusBar
    Set pStatusBar = gApplication.StatusBar
    Dim pStepProgressor As IStepProgressor
    Set pStepProgressor = pStatusBar.ProgressBar

    Dim pRasterBMP As IRaster
    Set pRasterBMP = DetermineBufferStrip(pFCBMP, pFCVFS)
    
    'MsgBox "From TraceBMPDrainageArea 2"
       
    ' create a DblPnt to hold the PixelBlock size
    Dim pSize As IPnt
    Set pSize = New DblPnt
    
    Dim pOrigin As IPnt
    Set pOrigin = New DblPnt
    
    Dim pPixelBlockFlowDir As IPixelBlock3
    Dim pPixelBlockBMP As IPixelBlock3
    Dim pPixelBlockDrain As IPixelBlock3

    Dim pRasterPropFlowDir As IRasterProps
    Dim pRasterPropBMP As IRasterProps
    Dim pRasterPropDrain As IRasterProps

    Dim vPixelDataFlowDir As Variant
    Dim vPixelDataBMP As Variant
    Dim vPixelDataDrain As Variant
    
    ' get raster properties
    Set pRasterPropFlowDir = pRasterFlowDir
    Set pRasterPropBMP = pRasterBMP

    ' get vb supported pixel type
    pRasterPropFlowDir.PixelType = GetVBSupportedPixelType(pRasterPropFlowDir.PixelType)
    pRasterPropBMP.PixelType = GetVBSupportedPixelType(pRasterPropBMP.PixelType)

   ' If pRasterPropFlowDir.Width <> pRasterPropBMP.Width Or pRasterPropFlowDir.Height <> pRasterPropBMP.Height Then
   '     Err.Description = "Extent mismatch"
   '     GoTo ShowError
   ' End If
    
    ' create CumSlopeLength raster
    Dim pOrg As IPoint
    Set pOrg = New Point
    pOrg.X = pRasterPropBMP.Extent.XMin
    pOrg.Y = pRasterPropBMP.Extent.YMin

    Dim pRWS As IRasterWorkspace2
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Set pRWS = pWSF.OpenFromFile(gMapTempFolder, 0)

''    Dim pRasterDrain As IRaster
''    Set pRasterDrain = gAlgebraOp.Execute("0.0")
    
    'MsgBox "From TraceBMPDrainageArea 3"
    
    Delete_Raster gMapTempFolder, "Drainage"  ' Arun Raj
    
    'MsgBox "From TraceBMPDrainageArea 4"
    
    Dim pRasterDSDrain As IRasterDataset
    Set pRasterDSDrain = pRWS.CreateRasterDataset("Drainage", "GRID", pOrg, _
                          pRasterPropBMP.Width, pRasterPropBMP.Height, _
                          pRasterPropBMP.MeanCellSize.X, pRasterPropBMP.MeanCellSize.Y, _
                          1, PT_FLOAT, pRasterPropBMP.SpatialReference, True)
    
    'MsgBox "From TraceBMPDrainageArea 5"
    
    Dim pRasterDrain As IRaster
    Set pRasterDrain = pRasterDSDrain.CreateDefaultRaster

    Dim pRasterBandDrain As IRasterBandCollection
    Set pRasterBandDrain = pRasterDrain
    Dim pRawpixelDrain As IRawPixels
    Set pRawpixelDrain = pRasterBandDrain.Item(0)
    Set pRasterPropDrain = pRawpixelDrain
    
    'MsgBox "From TraceBMPDrainageArea 5_1"
    ' get vb supported pixel type
    'pRasterPropDrain.PixelType = GetVBSupportedPixelType(pRasterPropDrain.PixelType)
    
    pSize.SetCoords pRasterPropDrain.Width, pRasterPropDrain.Height
    Set pPixelBlockDrain = pRawpixelDrain.CreatePixelBlock(pSize)

    'MsgBox "From TraceBMPDrainageArea 5_2"
    ' Read pixelblock
    pOrigin.SetCoords 0, 0
    pRawpixelDrain.Read pOrigin, pPixelBlockDrain

    'MsgBox "From TraceBMPDrainageArea 5_3"
    ' Get pixeldata array
    Dim pPixelDataDrain As Variant
    pPixelDataDrain = pPixelBlockDrain.PixelDataByRef(0)
        
    'MsgBox "From TraceBMPDrainageArea 5_4"
    Dim iCol As Integer
    Dim iRow As Integer
    Dim cCol As Integer
    Dim cRow As Integer
    Dim pValueFlowDir As Double
    Dim pValueBMP As Double
    Dim pValueDrain As Double
    
    Dim noDataForBMP As Double
    'Dim noDataForBMP As Integer
    
    'Error here--------------------------
    noDataForBMP = pRasterPropBMP.NoDataValue(0)
    
    'MsgBox "From TraceBMPDrainageArea 5_5_1  noDataForBMP = " & noDataForBMP
        
    Dim noDataForDrain  As Double
    'Set pRasterPropDrain = pRasterDrain
    noDataForDrain = pRasterPropDrain.NoDataValue '(0)
    
    'MsgBox "From TraceBMPDrainageArea 5_5_2 noDataForDrain" & noDataForDrain
    
    
    pStepProgressor.MinRange = 0
    pStepProgressor.MaxRange = pRasterPropBMP.Width - 1
    pStepProgressor.StepValue = pRasterPropBMP.Width / 100

    pSize.SetCoords pRasterPropBMP.Width, pRasterPropBMP.Height
    Set pPixelBlockFlowDir = pRasterFlowDir.CreatePixelBlock(pSize)
    Set pPixelBlockBMP = pRasterBMP.CreatePixelBlock(pSize)
   ' Set pPixelBlockDrain = pRasterDrain.CreatePixelBlock(pSize)
    
    'MsgBox "From TraceBMPDrainageArea 5_6"
    pOrigin.SetCoords 0, 0
    pRasterFlowDir.Read pOrigin, pPixelBlockFlowDir
    vPixelDataFlowDir = pPixelBlockFlowDir.PixelDataByRef(0)
    pRasterBMP.Read pOrigin, pPixelBlockBMP
    vPixelDataBMP = pPixelBlockBMP.PixelDataByRef(0)
    'vPixelDataDrain = pPixelBlockDrain.PixelDataByRef(0)
    
    'MsgBox "From TraceBMPDrainageArea 6"
    
    For iCol = 0 To pRasterPropBMP.Width - 1
        For iRow = 0 To pRasterPropBMP.Height - 1
            'MsgBox "From TraceBMPDrainageArea 6_1"
            pValueBMP = vPixelDataBMP(iCol, iRow)
            'MsgBox "From TraceBMPDrainageArea 6_2"
            If pValueBMP <> noDataForBMP Then
                'MsgBox "From TraceBMPDrainageArea 6_6_1"
                pPixelDataDrain(iCol, iRow) = pValueBMP
                'MsgBox "From TraceBMPDrainageArea 6_6_2"
            Else
                'MsgBox "From TraceBMPDrainageArea 6_6_3"
                pPixelDataDrain(iCol, iRow) = noDataForDrain
                'MsgBox "From TraceBMPDrainageArea 6_6_4"
            End If
        Next iRow
    Next iCol
    
    Dim ptArray()
    Dim pLocation As IPnt
    Dim nCount As Integer
    Dim bContinueNext As Boolean
    Dim i As Long
    
    'MsgBox "From TraceBMPDrainageArea 7"
    
    Dim noDataForFlowDir As Double
    noDataForFlowDir = pRasterPropFlowDir.NoDataValue(0)
    
    'MsgBox "From TraceBMPDrainageArea 7_1"
    
    pStepProgressor.Message = "Tracing drainage area ..."
    pStepProgressor.Show
    
    For iCol = 0 To pRasterPropBMP.Width - 1
        pStepProgressor.Position = iCol
        pStepProgressor.Step
        
        For iRow = 0 To pRasterPropBMP.Height - 1
            bContinueNext = False
            
            ' If this cell has been assigned a stream id, continue to process next cell
            pValueDrain = pPixelDataDrain(iCol, iRow)
            If pValueDrain <> noDataForDrain Then
                bContinueNext = True
            End If
            
            ' If this cell has no value for flow direction, continue to process next cell
            pValueFlowDir = vPixelDataFlowDir(iCol, iRow)
            If pValueFlowDir = noDataForFlowDir Then 'pRasterPropFlowDir.NoDataValue Then ' (0)
                bContinueNext = True
            End If
            
            ' Otherwise, search the flow path to the nearest cell with stream id assigned starting from this cell
            If Not bContinueNext Then
                cCol = iCol
                cRow = iRow
                nCount = 0
                
                Do
                    Set pLocation = New DblPnt
                    pLocation.SetCoords cCol, cRow
                    ReDim Preserve ptArray(nCount)
                    ptArray(nCount) = pLocation
                    nCount = nCount + 1
                    
                    Select Case pValueFlowDir
                        Case 1
                            cCol = cCol + 1
                        Case 2
                            cRow = cRow + 1
                            cCol = cCol + 1
                        Case 4
                            cRow = cRow + 1
                        Case 8
                            cRow = cRow + 1
                            cCol = cCol - 1
                        Case 16
                            cCol = cCol - 1
                        Case 32
                            cRow = cRow - 1
                            cCol = cCol - 1
                        Case 64
                            cRow = cRow - 1
                        Case 128
                            cRow = cRow - 1
                            cCol = cCol + 1
                        Case Else
                            pValueDrain = 0
                            Exit Do
                    End Select
                    
                    ' Searching is out of grid boundary
                    If cRow < 0 Or cRow >= pPixelBlockBMP.Height Or cCol < 0 Or cCol >= pPixelBlockBMP.Width Then
                        pValueDrain = 0
                        Exit Do
                    End If
                    
                    ' This cell has been assigned bmp id
                    pValueDrain = pPixelDataDrain(cCol, cRow)
                    If pValueDrain <> noDataForDrain Then
                        Exit Do
                    End If
                    
                    ' This cell has no value for flow direction
                    pValueFlowDir = vPixelDataFlowDir(cCol, cRow)
                    If pValueFlowDir = noDataForFlowDir Then 'pRasterPropFlowDir.NoDataValue Then  ' (0)
                        pValueDrain = 0
                        Exit Do
                    End If
                Loop
                
                For i = 0 To UBound(ptArray)
                    Set pLocation = ptArray(i)
                    ' Assign the bmp id to the whole flow path
                    pPixelDataDrain(pLocation.X, pLocation.Y) = pValueDrain
                Next
            End If
        Next iRow
    Next iCol
    
    'MsgBox "From TraceBMPDrainageArea 8"
    
    For iCol = 0 To pPixelBlockBMP.Width - 1
        For iRow = 0 To pPixelBlockBMP.Height - 1
            ' Set 0 to NoData
            pValueDrain = pPixelDataDrain(iCol, iRow)
            If pValueDrain = 0 Then
                pPixelDataDrain(iCol, iRow) = noDataForDrain
            End If
        Next iRow
    Next iCol
    
    ' Write the pixeldata back
    Dim pCache
    Set pCache = pRawpixelDrain.AcquireCache
    pRawpixelDrain.Write pOrigin, pPixelBlockDrain
    pRawpixelDrain.ReturnCache pCache
  
    Dim pBandCol As IRasterBandCollection
    Set pBandCol = pRasterDrain

    Dim pBand As IRasterBand
    Set pBand = pBandCol.Item(0)

    Dim pRawPixel As IRawPixels
    Set pRawPixel = pBand

    pStepProgressor.Message = "Saving drainage area raster ..."
    ' Write back the pixel block to flow accumulation grid
    pRawPixel.Write pOrigin, pPixelBlockDrain

    pBand.ComputeStatsAndHist
    pStepProgressor.Hide
    
    'MsgBox "From TraceBMPDrainageArea 9"
    ' ********************************
    ' Arun Raj
    DeleteLayerFromMap "SubWatershed"
    Delete_Raster gMapTempFolder, "iDrainage"
    Dim pRaster As IRaster

    pBandCol.SaveAs "iDrainage", pRWS, "GRID"
    Set pRaster = OpenRasterDatasetFromDisk("iDrainage")
    
    Set TraceBMPDrainageArea = pRaster
    GoTo CleanUp
    
ShowError:
    MsgBox "BMP drainage area calculation error: " & Err.description, vbExclamation
CleanUp:
    Set pStatusBar = Nothing
    Set pStepProgressor = Nothing
End Function



'Private Function DetermineBufferStrip(pFCBMP As IFeatureClass, pFCVFS As IFeatureClass) As IRaster
'On Error GoTo ShowError
'
'    Dim pRasterDSBMP As IRasterDataset
'    Dim pRasterBMP As IRaster
'    If (Not pFCBMP Is Nothing) Then
'        Set pRasterDSBMP = ConvertFeatureToRaster(pFCBMP, "POINTID", "bmpgrid", Nothing)
'        Set pRasterBMP = pRasterDSBMP.CreateDefaultRaster
'    End If
'
'
'    Dim pRasterDSVFS As IRasterDataset
'    Dim pRasterVFS As IRaster
'    If (Not pFCVFS Is Nothing) Then
'        Set pRasterDSVFS = ConvertFeatureToRaster(pFCVFS, "ID", "vfsgrid", Nothing)
'        Set pRasterVFS = pRasterDSVFS.CreateDefaultRaster
'    End If
'
'    '** Merge both rasters if both are present
'    If ((Not pRasterBMP Is Nothing) And (Not pRasterVFS Is Nothing)) Then
'        Dim pRasterMerge As IRaster
'        gAlgebraOp.BindRaster pRasterBMP, "BMP"
'        gAlgebraOp.BindRaster pRasterVFS, "VFS"
'        Set pRasterMerge = gAlgebraOp.Execute("Merge([BMP], [VFS])")
'        gAlgebraOp.UnbindRaster "BMP"
'        gAlgebraOp.UnbindRaster "VFS"
'    ElseIf (pRasterBMP Is Nothing) Then
'        gAlgebraOp.BindRaster pRasterVFS, "VFS"
'        Set pRasterMerge = gAlgebraOp.Execute("[VFS]")
'        gAlgebraOp.UnbindRaster "VFS"
'    ElseIf (pRasterVFS Is Nothing) Then
'        gAlgebraOp.BindRaster pRasterBMP, "BMP"
'        Set pRasterMerge = gAlgebraOp.Execute("[BMP]")
'        gAlgebraOp.UnbindRaster "BMP"
'    End If
'
'    AddRasterToMap pRasterMerge, "bmpras", True
'
'    Dim pStatusBar As esriSystem.IStatusBar
'    Set pStatusBar = gApplication.StatusBar
'    Dim pStepProgressor As IStepProgressor
'    Set pStepProgressor = pStatusBar.ProgressBar
'
'    Dim pDict As Object
'    Set pDict = CreateObject("Scripting.Dictionary")
'    Dim pPolyline As IPolyline
'    Dim pBMPPoint As IPoint
'    Dim BMPId As Long
'
'    Dim pFeatureCursor As IFeatureCursor
'    Dim pFeature As IFeature
'    Dim bmpIDFldIndex As Long
'    Dim vfsIDFldIndex As Long
'
'    If (Not pFCBMP Is Nothing) Then
'        Set pFeatureCursor = pFCBMP.Search(Nothing, False)
'        Set pFeature = pFeatureCursor.NextFeature
'        bmpIDFldIndex = pFCBMP.FindField("POINTID")
'        Do Until pFeature Is Nothing
'            Set pBMPPoint = pFeature.Shape
'            BMPId = CLng(pFeature.value(bmpIDFldIndex))
'            pDict.Add BMPId, pBMPPoint
'            Set pFeature = pFeatureCursor.NextFeature
'        Loop
'    End If
'
'    If (Not pFCVFS Is Nothing) Then
'        Set pFeatureCursor = pFCVFS.Search(Nothing, False)
'        Set pFeature = pFeatureCursor.NextFeature
'        vfsIDFldIndex = pFCVFS.FindField("ID")
'        Do Until pFeature Is Nothing
'            Set pPolyline = pFeature.Shape
'            BMPId = CLng(pFeature.value(vfsIDFldIndex))
'            pDict.Add BMPId, pPolyline
'            Set pFeature = pFeatureCursor.NextFeature
'        Loop
'    End If
'
'''    Dim pFCDesc As IFeatureClassDescriptor
'''    Set pFCDesc = New FeatureClassDescriptor
'''    pFCDesc.Create pFCBMP, Nothing, "ID"
'''
'''    Dim pRasterDS As IRasterDataset
'''    Set pRasterDS = ConvertFeatureToRaster(pFCBMP, "ID", "vfsgrid", Nothing)
'''
'''    Dim pRasterBMP As IRaster
'''    Set pRasterBMP = pRasterDS.CreateDefaultRaster
'
'
'    Dim pRasterBandBMP As IRasterBandCollection
'    Set pRasterBandBMP = pRasterMerge
'
'    Dim pRawpixelBMP As IRawPixels
'    Set pRawpixelBMP = pRasterBandBMP.Item(0)
'
'    Dim pRasterPropBMP As IRasterProps
'    Set pRasterPropBMP = pRawpixelBMP
'
'    ' get vb supported pixel type
'    'pRasterPropBMP.PixelType = GetVBSupportedPixelType(pRasterPropBMP.PixelType)
'
'    'Dim vPixelDataBMP As Variant
'
'    Dim pSize As IPnt
'    Set pSize = New DblPnt
'    pSize.SetCoords pRasterPropBMP.Width, pRasterPropBMP.Height
'
'    Dim pPixelBlockBMP As IPixelBlock3
'    Set pPixelBlockBMP = pRawpixelBMP.CreatePixelBlock(pSize)
'
'    Dim pOrigin As IPnt
'    Set pOrigin = New DblPnt
'    pOrigin.SetCoords 0, 0
'    ' Read pixelblock
'    pRawpixelBMP.Read pOrigin, pPixelBlockBMP
'
'    ' Get pixeldata array
'    Dim pPixelDataBMP
'    pPixelDataBMP = pPixelBlockBMP.PixelDataByRef(0)
'
'    Dim iCol As Integer
'    Dim iRow As Integer
'    Dim cCol As Integer
'    Dim cRow As Integer
'    Dim pValueBMP As Double
'
'    pStepProgressor.MinRange = 0
'    pStepProgressor.MaxRange = pRasterPropBMP.Width - 1
'    pStepProgressor.StepValue = pRasterPropBMP.Width / 100
'
'    Dim pPoint As IPoint
'    Set pPoint = New Point
'
'    Dim pNearPoint As IPoint
'    Set pNearPoint = New Point
'    Dim DistOnCurve As Double
'    Dim NearDist As Double
'    Dim bRight As Boolean
'
'    Dim noDataForBMP As Double
'    noDataForBMP = pRasterPropBMP.NoDataValue  ' .NoDataValue(0)
'    Dim bChange As Boolean
'    bChange = True
'
'    Dim bmpKey 'As Object
'    For Each bmpKey In pDict.keys
'        BMPId = CLng(bmpKey)
'        If (TypeOf pDict.Item(BMPId) Is IPolyline) Then
'                Set pPolyline = pDict.Item(BMPId)
'                For iCol = 0 To pRasterPropBMP.Width - 1
'                    For iRow = 0 To pRasterPropBMP.Height - 1
'                        If pPixelDataBMP(iCol, iRow) = BMPId Then
'                            If bChange Then
'                                'pPixelDataBMP(iCol, iRow) = 10 + BMPId   '-bmpID
'                            End If
'                            bChange = Not bChange
'
'                            For cCol = iCol - 1 To iCol + 1
'                                If cCol >= 0 And cCol < pRasterPropBMP.Width Then
'                                    pPoint.X = pRasterPropBMP.Extent.XMin + (cCol + 0.5) * pRasterPropBMP.MeanCellSize.X
'                                    For cRow = iRow - 1 To iRow + 1
'                                        If cRow >= 0 And cRow < pRasterPropBMP.Height Then
'                                            If pPixelDataBMP(cCol, cRow) = noDataForBMP Then
'                                                pPoint.Y = pRasterPropBMP.Extent.YMax - (cRow + 0.5) * pRasterPropBMP.MeanCellSize.Y
'                                                pPolyline.QueryPointAndDistance esriNoExtension, pPoint, True, pNearPoint, DistOnCurve, NearDist, bRight
'                                                If bRight Then
'                                                    'pPixelDataBMP(cCol, cRow) = 10 + BMPId   '-bmpID
'                                                End If
'                                            End If
'                                        End If
'                                    Next
'                                End If
'                            Next
'                        End If
'                    Next
'                Next
'        End If  '** extend only if it is a polyline
'    Next
'
' ' Write the pixeldata back
'  Dim pCache
'  Set pCache = pRawpixelBMP.AcquireCache
'  pRawpixelBMP.Write pOrigin, pPixelBlockBMP
'  pRawpixelBMP.ReturnCache pCache
'
'  pStepProgressor.Hide
'
'    Set DetermineBufferStrip = pRasterMerge
'    GoTo CleanUp
'
'ShowError:
'    MsgBox "Determine buffer strip error: " & Err.description, vbExclamation
'CleanUp:
'End Function


Private Function DetermineBufferStrip(pFCBMP As IFeatureClass, pFCVFS As IFeatureClass) As IRaster
    On Error GoTo ShowError

    Dim pPolyline As IPolyline
    Dim bmpId As Long
    Dim bmpIDFldIndex As Long
    Dim barrierID As Long
    Dim barrierIDFldIndex As Long
    Dim bufSide As String
    Dim sideIDFldIndex As Long
       
    If Not pFCVFS Is Nothing Then
               
        Dim pStatusBar As IStatusBar
        Set pStatusBar = gApplication.StatusBar
        Dim pStepProgressor As IStepProgressor
        Set pStepProgressor = pStatusBar.ProgressBar
    
        ' Find all relevant field indices
        bmpIDFldIndex = pFCVFS.FindField("ID")
        barrierIDFldIndex = pFCVFS.FindField("DSID")
        sideIDFldIndex = pFCVFS.FindField("TYPE2")
        
        ' Use the barrier ID to generate the template raster
        Dim pFCDesc As IFeatureClassDescriptor
        Set pFCDesc = New FeatureClassDescriptor
        pFCDesc.Create pFCVFS, Nothing, "ID" ' Need a Downstream ID...
               
        'Create a workspace
        Dim pWSF As IWorkspaceFactory
        Dim gRWS As IRasterWorkspace2
        Set pWSF = New RasterWorkspaceFactory
        Set gRWS = pWSF.OpenFromFile(gMapTempFolder, 0)
        Dim pConversionOp As IConversionOp
        ' Create the pConversionOp object
        Set pConversionOp = New RasterConversionOp
        
        Dim pRasterLayer As IRasterLayer
        Set pRasterLayer = GetInputRasterLayer("DEM")
        
        If pRasterLayer Is Nothing Then
            Set pRasterLayer = GetInputRasterLayer("Landuse")
        End If
        If pRasterLayer Is Nothing Then Err.Raise "DEM layer or Landuse layer is not defined"
               
        Dim pRasterProps As IRasterProps
        Set pRasterProps = pRasterLayer.Raster
            
        'Get the conversion environment
        Dim pEnv As IRasterAnalysisEnvironment
        Dim pCellSize As Double
        Set pEnv = pConversionOp
        
        Set pEnv.OutSpatialReference = pRasterProps.SpatialReference
        pEnv.SetExtent esriRasterEnvValue, pRasterProps.Extent
        pCellSize = pRasterProps.MeanCellSize.X
        pEnv.SetCellSize esriRasterEnvValue, pCellSize
        Set pEnv.Mask = pRasterLayer.Raster
        
        Delete_Raster gMapTempFolder, "bmpras"
        
        Dim pRasterDS As IRasterDataset
        Set pRasterDS = pConversionOp.ToRasterDataset(pFCDesc, "GRID", gRWS, "bmpras")
       
        Dim pRasterVFS As IRaster
        Set pRasterVFS = pRasterDS.CreateDefaultRaster
        
        Dim pRasterPropBMP As IRasterProps
        Set pRasterPropBMP = pRasterVFS
        
        ' get vb supported pixel type
        pRasterPropBMP.PixelType = GetVBSupportedPixelType(pRasterPropBMP.PixelType)
        
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        
        Dim pFeatureCursor As IFeatureCursor
        Set pFeatureCursor = pFCVFS.Search(Nothing, False)
        
        Dim pFeature As IFeature
        Set pFeature = pFeatureCursor.NextFeature
        
        Dim pPixelBlockBMP As IPixelBlock3
        Dim vPixelDataBMP As Variant
        
        Dim pSize As IPnt
        Set pSize = New DblPnt
        
        Dim pOrigin As IPnt
        Set pOrigin = New DblPnt
        
        pSize.SetCoords pRasterPropBMP.Width, pRasterPropBMP.Height
        Set pPixelBlockBMP = pRasterVFS.CreatePixelBlock(pSize)
        
        pOrigin.SetCoords 0, 0
        pRasterVFS.Read pOrigin, pPixelBlockBMP
        vPixelDataBMP = pPixelBlockBMP.PixelDataByRef(0)
    
        Dim iCol As Integer
        Dim iRow As Integer
        Dim cCol As Integer
        Dim cRow As Integer
        Dim pValueBMP As Double
        
        Do Until pFeature Is Nothing
            Set pPolyline = pFeature.Shape
            bmpId = CLng(pFeature.value(bmpIDFldIndex))
            barrierID = CLng(pFeature.value(barrierIDFldIndex))
            bufSide = CStr(pFeature.value(sideIDFldIndex))
            
''            pQueryFilter.WhereClause = "BMP_ID = " & bmpId
''            pFCDesc.Create pFCVFS, pQueryFilter, "BMP_ID"
            
            pQueryFilter.WhereClause = "ID = " & bmpId
            pFCDesc.Create pFCVFS, pQueryFilter, "ID"
            
            Delete_Raster gMapTempFolder, "bmpras_bs" & bmpId
            
            Dim pRasterDS_BS As IRasterDataset
            Set pRasterDS_BS = pConversionOp.ToRasterDataset(pFCDesc, "GRID", gRWS, "bmpras_bs" & bmpId)
            
            Dim pRasterBMP_BS As IRaster
            Set pRasterBMP_BS = pRasterDS_BS.CreateDefaultRaster
            
            Dim pRasterPropBMP_BS As IRasterProps
            Set pRasterPropBMP_BS = pRasterBMP_BS
            
            ' get vb supported pixel type
            pRasterPropBMP_BS.PixelType = GetVBSupportedPixelType(pRasterPropBMP_BS.PixelType)
            
            pStepProgressor.Message = "Processing Buffer Strip with ID " & bmpId & " ..."
            pStepProgressor.MinRange = 0
            pStepProgressor.MaxRange = pRasterPropBMP_BS.Width - 1
            pStepProgressor.StepValue = pRasterPropBMP_BS.Width / 100
        
            Dim pPixelBlockBMP_BS As IPixelBlock3
            Dim vPixelDataBMP_BS As Variant
        
            pSize.SetCoords pRasterPropBMP_BS.Width, pRasterPropBMP_BS.Height
            Set pPixelBlockBMP_BS = pRasterBMP_BS.CreatePixelBlock(pSize)
            
            pOrigin.SetCoords 0, 0
            pRasterBMP_BS.Read pOrigin, pPixelBlockBMP_BS
            vPixelDataBMP_BS = pPixelBlockBMP_BS.PixelDataByRef(0)
            
            Dim pPoint As IPoint
            Set pPoint = New Point
            
            Dim pNearPoint As IPoint
            Set pNearPoint = New Point
            Dim DistOnCurve As Double
            Dim NearDist As Double
            Dim bRight As Boolean
            
            Dim bChange As Boolean
            bChange = True
            
            For iCol = 0 To pRasterPropBMP_BS.Width - 1
                pStepProgressor.Position = iCol
                pStepProgressor.Step
                For iRow = 0 To pRasterPropBMP_BS.Height - 1
                    If vPixelDataBMP_BS(iCol, iRow) = bmpId Then
                        For cCol = iCol - 1 To iCol + 1
                            If cCol >= 0 And cCol < pRasterPropBMP_BS.Width Then
                                pPoint.X = pRasterPropBMP_BS.Extent.XMin + (cCol + 0.5) * pRasterPropBMP_BS.MeanCellSize.X
                                For cRow = iRow - 1 To iRow + 1
                                    If cRow >= 0 And cRow < pRasterPropBMP_BS.Height Then
                                        If pPixelBlockBMP.GetNoDataMaskVal(0, cCol, cRow) = 0 Then
                                            pPoint.Y = pRasterPropBMP_BS.Extent.YMax - (cRow + 0.5) * pRasterPropBMP_BS.MeanCellSize.Y
                                            pPolyline.QueryPointAndDistance esriNoExtension, pPoint, True, pNearPoint, DistOnCurve, NearDist, bRight
                                            If bufSide = "VFS_R" And bRight Then
                                                pOrigin.SetCoords 1, 1
                                                vPixelDataBMP(cCol, cRow) = bmpId
                                            End If
                                            If bufSide = "VFS_L" And Not bRight Then
                                                'MsgBox pOrigin.X & " " & pOrigin.Y
                                                pOrigin.SetCoords -1, -1
                                                vPixelDataBMP(cCol, cRow) = bmpId
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Next
        
            Set pFeature = pFeatureCursor.NextFeature
        Loop
        
        Dim pBandCol As IRasterBandCollection
        Set pBandCol = pRasterVFS
        
        Dim pBand As IRasterBand
        Set pBand = pBandCol.Item(0)
        
        Dim pRawPixel As IRawPixels
        Set pRawPixel = pBand
    
        ' Write back the pixel block to bmp template raster
        pStepProgressor.Message = "Saving buffer strip raster ..."
        pRawPixel.Write pOrigin, pPixelBlockBMP
        
        pBand.ComputeStatsAndHist
        pStepProgressor.Hide
        
    End If
    
    ' ********************************************
    ' Now Merge the POINT & LINE BMPs.....
    ' ********************************************
    
    Dim pRasterDSBMP As IRasterDataset
    Dim pRasterBMP As IRaster
    If (Not pFCBMP Is Nothing) Then
        Set pRasterDSBMP = ConvertFeatureToRaster(pFCBMP, "POINTID", "bmpgrid", Nothing)
        Set pRasterBMP = pRasterDSBMP.CreateDefaultRaster
    End If

    '** Merge both rasters if both are present
    If ((Not pRasterBMP Is Nothing) And (Not pRasterVFS Is Nothing)) Then
        Dim pRasterMerge As IRaster
        gAlgebraOp.BindRaster pRasterBMP, "BMP"
        gAlgebraOp.BindRaster pRasterVFS, "VFS"
        Set pRasterMerge = gAlgebraOp.Execute("Merge([BMP], [VFS])")
        gAlgebraOp.UnbindRaster "BMP"
        gAlgebraOp.UnbindRaster "VFS"
    ElseIf (pRasterBMP Is Nothing) Then
        gAlgebraOp.BindRaster pRasterVFS, "VFS"
        Set pRasterMerge = gAlgebraOp.Execute("[VFS]")
        gAlgebraOp.UnbindRaster "VFS"
    ElseIf (pRasterVFS Is Nothing) Then
        gAlgebraOp.BindRaster pRasterBMP, "BMP"
        Set pRasterMerge = gAlgebraOp.Execute("[BMP]")
        gAlgebraOp.UnbindRaster "BMP"
    End If
    DeleteLayerFromMap "bmpras"
        
    Delete_Raster gMapTempFolder, "bmpras_m"
    
    gAlgebraOp.BindRaster pRasterMerge, "BMP"
    Set pRasterMerge = gAlgebraOp.Execute("Float([BMP])")
    gAlgebraOp.UnbindRaster "BMP"
    
    WriteRasterDatasetToDisk pRasterMerge, "bmpras_m"
    Set pRasterMerge = Nothing
    Set pRasterMerge = OpenRasterDatasetFromDisk("bmpras_m")
    
    AddRasterToMap pRasterMerge, "bmpras", True
    
    Set DetermineBufferStrip = pRasterMerge
    GoTo CleanUp
    
ShowError:
    MsgBox "Determine buffer strip error: " & Err.description, vbExclamation
CleanUp:
End Function





