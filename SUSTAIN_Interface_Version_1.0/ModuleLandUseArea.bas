Attribute VB_Name = "ModuleLandUse"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleLandUse
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

Public gLuGroupIdDict As Scripting.Dictionary 'dictionary to store land use group name - group ID dictionary
Public gLuIdGroupDict As Scripting.Dictionary 'dictionary to store land use group ID - group Name dictionary

'******************************************************************************
'Subroutine: ComputeLanduseAreaForEachSubBasin
'Author:     Mira Chokshi
'Purpose:    Summarize Landuse Distribution for each Subbasin.
'            Create a dictionary of Landuse code --> Landuse Area
'            Create a dictionary of Subbasin --> (Dictionary of LU code --> area)
'            Save these public defined dictionary in memory
'******************************************************************************
Public Sub ComputeLanduseAreaForEachSubBasin()
    
On Error GoTo ShowError
    'Get the Landuse raster layer
    Dim pLandUseRLayer As IRasterLayer
    Set pLandUseRLayer = GetInputRasterLayer("Landuse")
    Dim pRasterLanduse As IRaster
    Set pRasterLanduse = pLandUseRLayer.Raster
    If pRasterLanduse Is Nothing Then
        MsgBox "Not found the Landuse raster layer."
        GoTo CleanUp
    End If
    'Get the subwatershed raster layer
    Dim pSubWaterRLayer As IRasterLayer
    Set pSubWaterRLayer = GetInputRasterLayer("SubWatershed")
    If pSubWaterRLayer Is Nothing Then
        MsgBox "Not found the SubWatershed raster layer."
        GoTo CleanUp
    End If
    Dim pRasterSubWater As IRaster
    Set pRasterSubWater = pSubWaterRLayer.Raster

    'Initialize storage dictionary
    Set gSubWaterLandUseDict = CreateObject("Scripting.Dictionary")
    Dim pLandUseAreaDict As Scripting.Dictionary
    Set pLandUseAreaDict = CreateObject("Scripting.Dictionary")
    
    'Define variables for cell based computing
    Dim pPixelBlockLanduse As IPixelBlock3
    Dim pPixelBlockSubWater As IPixelBlock3
    Dim pRasterPropLanduse As esriDataSourcesRaster.IRasterProps
    Dim pRasterPropSubWater As esriDataSourcesRaster.IRasterProps
    Dim vPixelDataLanduse As Variant
    Dim vPixelDataSubWater As Variant
    Dim pOrg As IPoint
    Dim pCellSize As Double
    Dim pOrigin As IPnt
    Dim pSize As IPnt
    Dim pLocation As IPnt
    Dim iCol As Integer
    Dim iRow As Integer
    Dim cCol As Integer
    Dim cRow As Integer
    Dim pValueLanduse As Single
    Dim pValueSubWater As Single
    ' get raster properties
    Set pRasterPropLanduse = pRasterLanduse
    Set pRasterPropSubWater = pRasterSubWater
    
    Dim pSubWaterNoDataValue As Single
    pSubWaterNoDataValue = pRasterPropSubWater.NoDataValue(0)
    Dim pLandUseNoDataValue As Single
    pLandUseNoDataValue = pRasterPropLanduse.NoDataValue(0)
    'compare raster properties
    If (pRasterPropLanduse.Width <> pRasterPropSubWater.Width) Or (pRasterPropLanduse.Height <> pRasterPropSubWater.Height) Then
        MsgBox "Row and/or column number doesn't match for Landuse and SubWatershed raster layers."
        GoTo CleanUp
    End If
    'get raster extent and cell size
    Set pOrg = New Point
    pOrg.X = pRasterPropSubWater.Extent.XMin
    pOrg.Y = pRasterPropSubWater.Extent.YMax
    pCellSize = (pRasterPropSubWater.MeanCellSize.X + pRasterPropSubWater.MeanCellSize.Y) / 2
    Dim pCellArea As Double
    pCellArea = pCellSize * pCellSize
    ' create a DblPnt to hold the PixelBlock size
    Set pSize = New DblPnt
    pSize.SetCoords pRasterPropLanduse.Width, pRasterPropLanduse.Height
    ' create pixelblock the size of the input raster
    Set pPixelBlockLanduse = pRasterLanduse.CreatePixelBlock(pSize)
    Set pPixelBlockSubWater = pRasterSubWater.CreatePixelBlock(pSize)
    ' get vb supported pixel type
    pRasterPropLanduse.PixelType = GetVBSupportedPixelType(pRasterPropLanduse.PixelType)
    pRasterPropSubWater.PixelType = GetVBSupportedPixelType(pRasterPropSubWater.PixelType)
    'define status bar
    Dim pStatusBar As esriSystem.IStatusBar
    Set pStatusBar = gApplication.StatusBar
    Dim pStepProgressor As IStepProgressor
    Set pStepProgressor = pStatusBar.ProgressBar
    pStepProgressor.Show
    ' get pixeldata
    Set pOrigin = New DblPnt
    pOrigin.SetCoords 0, 0
    'read the landuse raster into array
    pRasterLanduse.Read pOrigin, pPixelBlockLanduse
    vPixelDataLanduse = pPixelBlockLanduse.PixelDataByRef(0)
    'read the subwatershed raster into array
    pRasterSubWater.Read pOrigin, pPixelBlockSubWater
    vPixelDataSubWater = pPixelBlockSubWater.PixelDataByRef(0)
    'start processing
    pStepProgressor.MinRange = 1
    pStepProgressor.MaxRange = pPixelBlockLanduse.Width
    pStepProgressor.StepValue = pPixelBlockLanduse.Width / 100
    pStepProgressor.Message = "Reading Landuse grid ... "
    'for each col, row, read landuse type, add area for each lu code
    Dim pLandUse As Double
    Dim pSubwater As Double
    
    Dim pTotalLanduse As Double
    For iCol = 0 To pRasterPropLanduse.Width - 1
        pStepProgressor.Position = iCol
        pStepProgressor.Step
        For iRow = 0 To pRasterPropLanduse.Height - 1
            pLandUse = vPixelDataLanduse(iCol, iRow)
            pSubwater = vPixelDataSubWater(iCol, iRow)
            If (pLandUse <> pLandUseNoDataValue And pSubwater <> pSubWaterNoDataValue) Then
                pTotalLanduse = 0
                If Not gSubWaterLandUseDict.Exists(pSubwater) Then
                    Set pLandUseAreaDict = CreateObject("Scripting.Dictionary")
                    pLandUseAreaDict.Item(pLandUse) = pCellArea
                Else
                    Set pLandUseAreaDict = gSubWaterLandUseDict.Item(pSubwater)
                    pTotalLanduse = pLandUseAreaDict.Item(pLandUse) + pCellArea
                    pLandUseAreaDict.Item(pLandUse) = pTotalLanduse
                End If
                Set gSubWaterLandUseDict.Item(pSubwater) = pLandUseAreaDict
                Set pLandUseAreaDict = Nothing
            End If
        Next
    Next
        
    GoTo CleanUp

ShowError:
    MsgBox "ComputeLanduseAreaForEachSubBasin: " & Err.description
CleanUp:
    pStepProgressor.Hide
    Set pLandUseRLayer = Nothing
    Set pRasterLanduse = Nothing
    Set pSubWaterRLayer = Nothing
    Set pRasterSubWater = Nothing
    Set pLandUseAreaDict = Nothing
    Set pPixelBlockLanduse = Nothing
    Set pPixelBlockSubWater = Nothing
    Set pRasterPropLanduse = Nothing
    Set pRasterPropSubWater = Nothing
    Set vPixelDataLanduse = Nothing
    Set vPixelDataSubWater = Nothing
    Set pOrg = Nothing
    Set pOrigin = Nothing
    Set pSize = Nothing
    Set pLocation = Nothing
    Set pStatusBar = Nothing
    Set pStepProgressor = Nothing
     
End Sub

'******************************************************************************
'Subroutine: AddLanduseReclassification
'Author:     Mira Chokshi
'Purpose:    This subroutine calls a function to create a table LUReclass
'            The input parameter is an array containing values of landuse
'            reclassification. Add all items in the input array in LUReclass
'            table.
'******************************************************************************
Public Sub AddLanduseReclassification(LandUseTextFile() As String)
On Error GoTo ShowError
    
    'Get landuse reclassification table: LUReclass, Create new if not found
    Dim pLUReClasstable As iTable
    Set pLUReClasstable = GetInputDataTable("LUReclass")
    If (pLUReClasstable Is Nothing) Then
    'If the table is present delete and add new -- Sabu Paul, Aug 24, 2004
        Set pLUReClasstable = CreateLandUseReclassificationTable("LUReclass")
        AddTableToMap pLUReClasstable
        Set pLUReClasstable = GetInputDataTable("LUReclass")
    Else
        pLUReClasstable.DeleteSearchedRows Nothing    'delete all records
    End If
    
    
    Dim pLUGroupIDindex As Long
    pLUGroupIDindex = pLUReClasstable.FindField("LUGroupID")
    Dim pLUGroupindex As Long
    pLUGroupindex = pLUReClasstable.FindField("LUGroup")
    Dim pLUCodeindex As Long
    pLUCodeindex = pLUReClasstable.FindField("LUCode")
    Dim pLUDescIndex As Long
    pLUDescIndex = pLUReClasstable.FindField("LUDescrip")
    Dim pPercentageIndex As Long
    pPercentageIndex = pLUReClasstable.FindField("Percentage")
    Dim pLUTypeindex As Long
    pLUTypeindex = pLUReClasstable.FindField("Impervious")
    
    Dim pSandFracindex As Long
    pSandFracindex = pLUReClasstable.FindField("SandFrac")
    Dim pSiltFracindex As Long
    pSiltFracindex = pLUReClasstable.FindField("SiltFrac")
    Dim pClayFracindex As Long
    pClayFracindex = pLUReClasstable.FindField("ClayFrac")
    
    'Iterate over the entire array
    Dim pRow As iRow
    Dim pLUCode As Integer
    Dim pLUType As String
    Dim pLUDescription As String
    Dim pLUPercent As Double
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
    Dim pTimeSeries As String
               
    Dim sandFrac As Double
    Dim siltFrac As Double
    Dim clayFrac As Double
    
    Dim i As Integer
    For i = 2 To UBound(LandUseTextFile, 2)
        'Get values from array
        pLuGroupID = LandUseTextFile(1, i)
        pLuGroup = LandUseTextFile(2, i)
        pLUCode = CInt(LandUseTextFile(3, i))
        pLUDescription = LandUseTextFile(4, i)
        pLUType = LandUseTextFile(5, i)
        pLUPercent = 0
        If (LandUseTextFile(6, i) <> "") Then
            pLUPercent = CDbl(LandUseTextFile(6, i) / 100)
        End If
       
        sandFrac = LandUseTextFile(7, i)
        siltFrac = LandUseTextFile(8, i)
        clayFrac = LandUseTextFile(9, i)
        
        'add new row
        Set pRow = pLUReClasstable.CreateRow
        pRow.value(pLUGroupIDindex) = pLuGroupID
        pRow.value(pLUGroupindex) = pLuGroup
        pRow.value(pLUCodeindex) = pLUCode
        pRow.value(pLUTypeindex) = pLUType
        pRow.value(pLUDescIndex) = pLUDescription
        pRow.value(pPercentageIndex) = pLUPercent
       
        pRow.value(pSandFracindex) = sandFrac
        pRow.value(pSiltFracindex) = siltFrac
        pRow.value(pClayFracindex) = clayFrac
        
        pRow.Store
    Next
    
    GoTo CleanUp
ShowError:
    MsgBox "AddLanduseReclassification : " & Err.description
CleanUp:
    Set pLUReClasstable = Nothing
    Set pRow = Nothing
End Sub

'******************************************************************************
'Subroutine: AddTimeSeriesAssignments
'Author:     Mira Chokshi
'Purpose:    This subroutine calls a function to create a table LUReclass
'            The input parameter is an array containing values of landuse
'            reclassification. Add all items in the input array in LUReclass
'            table.
'******************************************************************************
Public Sub AddTimeSeriesAssignments(LandUseTextFile() As String)
On Error GoTo ShowError
    
    'Get landuse reclassification table: LUReclass, Create new if not found
    Dim pTSAssignTable As iTable
    Set pTSAssignTable = GetInputDataTable("TSAssigns")
    If (pTSAssignTable Is Nothing) Then
    'If the table is present delete and add new -- Sabu Paul, Aug 24, 2004
        Set pTSAssignTable = CreateTimeSeriesAssignTable("TSAssigns")
        AddTableToMap pTSAssignTable
        Set pTSAssignTable = GetInputDataTable("TSAssigns")
    Else
        pTSAssignTable.DeleteSearchedRows Nothing    'delete all records
    End If
    
    
    Dim pLUGroupIDindex As Long
    pLUGroupIDindex = pTSAssignTable.FindField("LUGroupID")
    Dim pLUGroupindex As Long
    pLUGroupindex = pTSAssignTable.FindField("LUGroup")
    Dim pLUCodeindex As Long
    pLUCodeindex = pTSAssignTable.FindField("LUCode")
    Dim pLUDescIndex As Long
    pLUDescIndex = pTSAssignTable.FindField("LUDescrip")
    Dim pPercentageIndex As Long
    pPercentageIndex = pTSAssignTable.FindField("Percentage")
    Dim pLUTypeindex As Long
    pLUTypeindex = pTSAssignTable.FindField("Impervious")
    Dim pTimeSeriesindex As Long
    pTimeSeriesindex = pTSAssignTable.FindField("TimeSeries")
    
    Dim pSandFracindex As Long
    pSandFracindex = pTSAssignTable.FindField("SandFrac")
    Dim pSiltFracindex As Long
    pSiltFracindex = pTSAssignTable.FindField("SiltFrac")
    Dim pClayFracindex As Long
    pClayFracindex = pTSAssignTable.FindField("ClayFrac")
    
    'Iterate over the entire array
    Dim pRow As iRow
    Dim pLUCode As Integer
    Dim pLUType As String
    Dim pLUDescription As String
    Dim pLUPercent As Double
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
    Dim pTimeSeries As String
               
    Dim sandFrac As Double
    Dim siltFrac As Double
    Dim clayFrac As Double
    
    Dim i As Integer
    For i = 2 To UBound(LandUseTextFile, 2)
        'Get values from array
        pLuGroupID = LandUseTextFile(1, i)
        pLuGroup = LandUseTextFile(2, i)
        pLUCode = CInt(LandUseTextFile(3, i))
        pLUDescription = LandUseTextFile(4, i)
        pLUType = LandUseTextFile(5, i)
        pLUPercent = 0
        If (LandUseTextFile(6, i) <> "") Then
            pLUPercent = CDbl(LandUseTextFile(6, i) / 100)
        End If
        pTimeSeries = LandUseTextFile(7, i)
        
        sandFrac = LandUseTextFile(8, i)
        siltFrac = LandUseTextFile(9, i)
        clayFrac = LandUseTextFile(10, i)
        
        'add new row
        Set pRow = pTSAssignTable.CreateRow
        pRow.value(pLUGroupIDindex) = pLuGroupID
        pRow.value(pLUGroupindex) = pLuGroup
        pRow.value(pLUCodeindex) = pLUCode
        pRow.value(pLUTypeindex) = pLUType
        pRow.value(pLUDescIndex) = pLUDescription
        pRow.value(pPercentageIndex) = pLUPercent
        pRow.value(pTimeSeriesindex) = pTimeSeries
        
        pRow.value(pSandFracindex) = sandFrac
        pRow.value(pSiltFracindex) = siltFrac
        pRow.value(pClayFracindex) = clayFrac
        
        pRow.Store
    Next
    
    GoTo CleanUp
ShowError:
    MsgBox "AddLanduseReclassification : " & Err.description
CleanUp:
    Set pTSAssignTable = Nothing
    Set pRow = Nothing
End Sub
'******************************************************************************
'Subroutine: CreateTimeSeriesAssignTable
'Author:     Mira Chokshi
'Purpose:    This function creates a table (.dbf) and adds required fields:
'            LUGroupID, LUGroup, LUCode, Impervious, LUDescrip, Percentage, TimeSeries
'            The name of the DBASE file should not contain the .dbf extension
'******************************************************************************
Public Function CreateTimeSeriesAssignTable(pFileName As String) As iTable
On Error GoTo ShowError

    'delete data table from temp folder
    DeleteDataTable gMapTempFolder, pFileName
    'ppen the workspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)
    'define variables to add New Fields
    Dim pFieldsEdit As IFieldsEdit
    Dim pFieldEdit As IFieldEdit
    Dim pField As esriGeoDatabase.IField
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 10
    'Create Landuse Group ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUGroupID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    'Create Landuse Group Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUGroup"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(1) = pField
    'Create Landuse Code Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUCode"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
    'Create Landuse Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Impervious"
        .Type = esriFieldTypeString
        .Length = 4
    End With
    Set pFieldsEdit.Field(3) = pField
    'Create Landuse Description Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUDescrip"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(4) = pField
    'Create Percentage Value
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Percentage"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(5) = pField
    'Create Landuse Time Series Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "TimeSeries"
        .Type = esriFieldTypeString
        .Length = 100
    End With
    Set pFieldsEdit.Field(6) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SandFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(7) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SiltFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(8) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ClayFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(9) = pField
       
  Set CreateTimeSeriesAssignTable = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateTimeSeriesAssignTable: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing
End Function

'******************************************************************************
'Subroutine: CreateLandUseReclassificationTable
'Author:     Mira Chokshi
'Purpose:    This function creates a table (.dbf) and adds required fields:
'            LUGroupID, LUGroup, LUCode, Impervious, LUDescrip, Percentage, TimeSeries
'            The name of the DBASE file should not contain the .dbf extension
'******************************************************************************
Public Function CreateLandUseReclassificationTable(pFileName As String) As iTable
On Error GoTo ShowError

    'delete data table from temp folder
    DeleteDataTable gMapTempFolder, pFileName
    'ppen the workspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)
    'define variables to add New Fields
    Dim pFieldsEdit As IFieldsEdit
    Dim pFieldEdit As IFieldEdit
    Dim pField As esriGeoDatabase.IField
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 9
    'Create Landuse Group ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUGroupID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    'Create Landuse Group Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUGroup"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(1) = pField
    'Create Landuse Code Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUCode"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
    'Create Landuse Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Impervious"
        .Type = esriFieldTypeString
        .Length = 4
    End With
    Set pFieldsEdit.Field(3) = pField
    'Create Landuse Description Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUDescrip"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(4) = pField
    'Create Percentage Value
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Percentage"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(5) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SandFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(6) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SiltFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(7) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ClayFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(8) = pField
       
  Set CreateLandUseReclassificationTable = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateLandUseReclassificationTable: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing
End Function




'******************************************************************************
'Subroutine: AddExternalTimeSeriesForBMP
'Author:     Mira Chokshi
'Purpose:    This subroutine calls a function to create a table LUReclass
'            and adds a new row for to associate a timeseries file for a
'            BMP.
'******************************************************************************
Public Sub AddExternalTimeSeriesForBMP(bmpId As Integer, description As String, _
    multiplier As Double, timeseriesfile As String, sandFrac As Double, siltFrac As Double, clayFrac As Double)
On Error GoTo ShowError
    
    'Get landuse reclassification table: LUReclass, Create new if not found
    Dim pExternalTSTable As iTable
    Set pExternalTSTable = GetInputDataTable("ExternalTS")
    If (pExternalTSTable Is Nothing) Then
        Set pExternalTSTable = CreateExternalTimeSeriesTable("ExternalTS")
        AddTableToMap pExternalTSTable
        Set pExternalTSTable = GetInputDataTable("ExternalTS")
    End If
   
    'Define table indexes
    Dim iBMPIndex As Long
    iBMPIndex = pExternalTSTable.FindField("BMPID")
    Dim iLUDescIndex As Long
    iLUDescIndex = pExternalTSTable.FindField("LUDescrip")
    Dim iMultiplierindex As Long
    iMultiplierindex = pExternalTSTable.FindField("Multiplier")
    Dim iTimeseriesindex As Long
    iTimeseriesindex = pExternalTSTable.FindField("TimeSeries")
    'Added three new fields - June 18, 2007
    Dim iSandFracIndex As Long
    iSandFracIndex = pExternalTSTable.FindField("SandFrac")
    Dim iSiltFracIndex As Long
    iSiltFracIndex = pExternalTSTable.FindField("SiltFrac")
    Dim iClayFracIndex As Long
    iClayFracIndex = pExternalTSTable.FindField("ClayFrac")
    
    'Iterate over the entire array
    Dim pRow As iRow
    'add new row
    Set pRow = pExternalTSTable.CreateRow
    pRow.value(iBMPIndex) = bmpId
    pRow.value(iLUDescIndex) = description
    pRow.value(iMultiplierindex) = multiplier
    pRow.value(iTimeseriesindex) = timeseriesfile
    'Add three new fields
    pRow.value(iSandFracIndex) = sandFrac
    pRow.value(iSiltFracIndex) = siltFrac
    pRow.value(iClayFracIndex) = clayFrac
    pRow.Store

    GoTo CleanUp
ShowError:
    MsgBox "AddExternalTimeSeriesForBMP : " & Err.description
CleanUp:
    Set pExternalTSTable = Nothing
    Set pRow = Nothing
End Sub

Public Function CreateExternalTimeSeriesTable(pFileName As String) As iTable
On Error GoTo ShowError

    'delete data table from temp folder
    DeleteDataTable gMapTempFolder, pFileName
    'ppen the workspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)
    'define variables to add New Fields
    Dim pFieldsEdit As IFieldsEdit
    Dim pFieldEdit As IFieldEdit
    Dim pField As esriGeoDatabase.IField
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 7
    'Create Landuse Group ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "BMPID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField

    'Create Landuse Description Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LUDescrip"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create Multiplier Value
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Multiplier"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create Landuse Time Series Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "TimeSeries"
        .Type = esriFieldTypeString
        .Length = 100
    End With
    Set pFieldsEdit.Field(3) = pField
    
    'Create Sand Fraction Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SandFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(4) = pField
   
    'Create Silt Fraction Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SiltFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(5) = pField
    
    'Create Clay Fraction Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ClayFrac"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Scale = 10
        .Precision = 3
    End With
    Set pFieldsEdit.Field(6) = pField
    
    
  Set CreateExternalTimeSeriesTable = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateExternalTimeSeriesTable: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing
End Function

Public Function SetLuGroupIDDict() As Boolean
On Error GoTo ShowError
    SetLuGroupIDDict = False
    Set gLuGroupIdDict = CreateObject("Scripting.Dictionary")
    Set gLuIdGroupDict = CreateObject("Scripting.Dictionary")
        
    Dim pTable As iTable
    Dim pTableName As String
    If gExternalSimulation Then
        pTableName = "TSAssigns"
    Else
        pTableName = "LUReclass"
    End If
    Set pTable = GetInputDataTable(pTableName) '"LUReclass")
    If (pTable Is Nothing) Then
        MsgBox pTableName & " table not found."
        Exit Function
    End If
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pLuGroupID As Integer
    Dim iLuGroupID As Long
    iLuGroupID = pTable.FindField("LUGroupID")
      
    Dim pImpervious As Integer
    Dim iImpervious  As Long
    iImpervious = pTable.FindField("Impervious")
    
    Dim pLuGroup As String
    Dim iLuGroup As Long
    iLuGroup = pTable.FindField("LUGroup")
    
    If iLuGroupID < 0 Or iImpervious < 0 Or iLuGroup < 0 Then
        Err.Raise 5002, , "Missing required fields (LUGroupID,Percentage,Impervious, or LuGroup ) in LuReclass"
    End If
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "LUGroupID > 0"
    
    Dim pTableSort As ITableSort
    Set pTableSort = New esriGeoDatabase.TableSort
    
    With pTableSort
      .Fields = "LUGroupID, LuGroup"
      .Ascending("LUGroupID") = True
      .Ascending("LuGroup ") = True
      Set .QueryFilter = pQueryFilter
      Set .Table = pTable
    End With

    pTableSort.Sort Nothing
       
    Set pCursor = pTableSort.Rows
    Set pRow = pCursor.NextRow
    
    Do Until pRow Is Nothing
        pLuGroupID = pRow.value(iLuGroupID)
        pImpervious = pRow.value(iImpervious)
        If pImpervious = 1 Then
            pLuGroup = pRow.value(iLuGroup) & "_Impervious"
        Else
            pLuGroup = pRow.value(iLuGroup) & "_Pervious"
        End If
        
        gLuGroupIdDict.Item(pLuGroup) = pLuGroupID
        gLuIdGroupDict.Item(pLuGroupID) = pLuGroup
            
'        If Not gLuGroupIdDict.Exists(pLuGroup) Then
'            pLUGroupId = pLUGroupId + 1
'            gLuGroupIdDict.Item(pLuGroup) = pLUGroupId
'            gLuIdGroupDict.Item(pLUGroupId) = pLuGroup
'        End If
        Set pRow = pCursor.NextRow
    Loop
    SetLuGroupIDDict = True
    GoTo CleanUp
ShowError:
   MsgBox "Error in SetLuGroupIDDict: " & Err.description
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pTableSort = Nothing
End Function

Public Function GetSubWsLuGroupDict(pSubWsID As Integer) As Scripting.Dictionary
On Error GoTo ShowError
    Dim pResultDict As Scripting.Dictionary
    Set pResultDict = New Scripting.Dictionary
        
    Dim pTable As iTable
    Dim pTableName As String
    
    If gExternalSimulation Then
        pTableName = "TSAssigns"
    Else
        pTableName = "LUReclass"
    End If
    Set pTable = GetInputDataTable(pTableName)  '"LUReclass")
    If (pTable Is Nothing) Then
        MsgBox pTableName & " table not found."
        Exit Function
    End If
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pLUCode As Integer
    Dim iLuCode As Long
    iLuCode = pTable.FindField("LUCode")
    
    Dim pPercentage As Double
    Dim iPercentageCode As Long
    iPercentageCode = pTable.FindField("Percentage")
   
    Dim pImpervious As Integer
    Dim iImpervious  As Long
    iImpervious = pTable.FindField("Impervious")
    
    Dim pLuGroup As String
    Dim iLuGroup As Long
    iLuGroup = pTable.FindField("LUGroup")
    
    If iLuCode < 0 Or iPercentageCode < 0 Or iImpervious < 0 Or iLuGroup < 0 Then
        Err.Raise 5002, , "Missing required fields (LuCode,Percentage,Impervious, or LuGroup ) in LuReclass"
    End If
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    'pQueryFilter.WhereClause = "LUGroupID > 0 ORDER BY LUGroupID, LuGroup"
    pQueryFilter.WhereClause = "LUGroupID > 0"
    
    Dim pTableSort As ITableSort
    Set pTableSort = New esriGeoDatabase.TableSort
    
    With pTableSort
      .Fields = "LUGroupID, LuGroup"
      .Ascending("LUGroupID") = True
      .Ascending("LuGroup ") = True
      Set .QueryFilter = pQueryFilter
      Set .Table = pTable
    End With

    pTableSort.Sort Nothing
    
    Dim pLandUseAreaDict As Scripting.Dictionary
    Set pLandUseAreaDict = gSubWaterLandUseDict.Item(pSubWsID)
    If pLandUseAreaDict Is Nothing Then Exit Function
    
    Dim pLandTypeKeys, iLandType As Integer
    Dim pLUAreaPerGroup As Double
    
    Dim pSQAcreFactor As Double
    If gMetersPerUnit = 0# Then GetMetersPerLinearUnit
    pSQAcreFactor = gMetersPerUnit * gMetersPerUnit * 0.0002471044       'sq meter to acre conversion
    
    'Set pCursor = pTable.Search(pQueryFilter, False)
    Set pCursor = pTableSort.Rows
    
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
        pLUCode = pRow.value(iLuCode)
        pImpervious = pRow.value(iImpervious)
        If pImpervious = 1 Then
            pLuGroup = pRow.value(iLuGroup) & "_Impervious"
        Else
            pLuGroup = pRow.value(iLuGroup) & "_Pervious"
        End If
        
        pPercentage = pRow.value(iPercentageCode)
        pLandTypeKeys = pLandUseAreaDict.keys

        'Make sure this logic is correct!!!!!!!!!!!
        If pPercentage = 0 Then pPercentage = 1 '100
        
'        For iLandType = 0 To pLandUseAreaDict.Count - 1
'            If (pLandTypeKeys(iLandType) = pLUCode) Then
'                pLUAreaPerGroup = 0
'                If (pLandUseAreaDict.Exists(pLandTypeKeys(iLandType))) Then
'                     pLUAreaPerGroup = pLandUseAreaDict.Item(pLandTypeKeys(iLandType))
'                End If
'                pResultDict.Item(pLuGroup) = pResultDict.Item(pLuGroup) + ((pLUAreaPerGroup * pPercentage) / 4046.856)
'                'iGroupArea = iGroupArea + ((pLUAreaPerGroup * pPercentage) / 4046.856)
'            End If
'        Next
        
        If pLandUseAreaDict.Exists(pLUCode) Then
            pLUAreaPerGroup = pLandUseAreaDict.Item(pLUCode)
            pResultDict.Item(pLuGroup) = pResultDict.Item(pLuGroup) + ((pLUAreaPerGroup * pPercentage) * pSQAcreFactor)
        End If
        
        Set pRow = pCursor.NextRow
    Loop

    Set GetSubWsLuGroupDict = pResultDict
    
    GoTo CleanUp
ShowError:
    MsgBox "Error in GetSubWsLuGroupDict: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
End Function

