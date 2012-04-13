Attribute VB_Name = "ModuleMenu"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleMenu
'   Purpose:     This module enables and disables various menu items.
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created: 08/23/2004 - mira chokshi
'
'******************************************************************************

Option Explicit
Option Base 0


Public Function EnableSustain() As Boolean
'    If Date > CDate("7/06/2009") Then
'        EnableSustain = False
'    Else
'        EnableSustain = True
'    End If
    EnableSustain = True
End Function
Public Function EnableDefineBMP() As Boolean
    If Not EnableSustain Then
        EnableDefineBMP = False
        Exit Function
    End If
    'Initialize Map Document
    InitializeMapDocument
    
    'Cannot define BMP, if datasources are not define
    EnableDefineBMP = ModuleUtility.ValidateDataSource
    If (EnableDefineBMP = False) Then
        Exit Function
    End If
    
    'if landuse is not reclassified define BMP cannot be enabled
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    
    If (pTable Is Nothing) Then
        EnableDefineBMP = False
        Exit Function
    End If
    
    ' Set the value for the Flag....
    EnableDefineBMP = (pTable.RowCount(Nothing) > 0)

End Function

Public Function EnableAddVFS() As Boolean
    
    'Cannot add bmp is BMP is not defined
    EnableAddVFS = EnableDefineBMP
    If (EnableAddVFS = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If (pTable Is Nothing) Then
        EnableAddVFS = False
        Exit Function
    End If
    EnableAddVFS = True
    Set pTable = Nothing
    
End Function

Public Function EnableAddBMPOnLand() As Boolean
    'Cannot add bmp is BMP is not defined
    EnableAddBMPOnLand = EnableDefineBMP
    If (EnableAddBMPOnLand = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPTypes")
    If (pTable Is Nothing) Then
        EnableAddBMPOnLand = False
        Exit Function
    End If
    Set pTable = Nothing
    Set pTable = GetInputDataTable("BMPDefaults")
    If (pTable Is Nothing) Then
        EnableAddBMPOnLand = False
        Exit Function
    End If
    EnableAddBMPOnLand = True
End Function


Public Function EnableIndividualBMPTool(pBMPType As String) As Boolean
    'Cannot add bmp is BMP is not defined
    If pBMPType = "Aggregate" And gInternalSimulation Then
        EnableIndividualBMPTool = False
        Exit Function
    End If
    EnableIndividualBMPTool = EnableAddBMPOnLand
    If (EnableIndividualBMPTool = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPTypes")
    If (pTable Is Nothing) Then
        EnableIndividualBMPTool = False
        Exit Function
    End If
    
    'find it the type is defined
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "Type = '" & pBMPType & "'"
    Dim pTypeCount As Integer
    pTypeCount = pTable.RowCount(pQueryFilter)
    If (pTypeCount = 0) Then
        EnableIndividualBMPTool = False
    Else
        EnableIndividualBMPTool = True
    End If
    Set pQueryFilter = Nothing
    Set pTable = Nothing
End Function


Public Function EnableAddBMPOnStream() As Boolean
    'Cannot add bmp is BMP is not defined
    EnableAddBMPOnStream = EnableAddBMPOnLand
    If (EnableAddBMPOnStream = False) Then
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("STREAM")
    If (pFeatureLayer Is Nothing) Then
        EnableAddBMPOnStream = False
        Exit Function
    End If
    EnableAddBMPOnStream = True
End Function


Public Function EnableBufferStrip() As Boolean
    '** check if data sources are defined
    EnableBufferStrip = ModuleUtility.ValidateDataSource
    If (EnableBufferStrip = False) Then
        Exit Function
    End If
    '** get the buffer strip defaults table is present
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BufferStripDefault")
    If (pTable Is Nothing) Then
        EnableBufferStrip = False
        Exit Function
    End If
    Set pTable = Nothing

End Function

Public Function EnableEditBMP() As Boolean
    'Cannot edit bmp unless templates defined and bmps added
    EnableEditBMP = EnableAddBMPOnLand
    If (EnableEditBMP = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPDetail")
    If (pTable Is Nothing) Then
        Set pTable = GetInputDataTable("AgBMPDetail")
        If pTable Is Nothing Then
            EnableEditBMP = False
            Exit Function
        End If
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        EnableEditBMP = False
        Exit Function
    End If
    EnableEditBMP = True
End Function

Public Function EnableBMPNetworkRouting() As Boolean
    EnableBMPNetworkRouting = EnableEditBMP
End Function

Public Function EnableAutoBMPNetworkRouting() As Boolean
    EnableAutoBMPNetworkRouting = EnableAutoDelineation
    If (EnableAutoBMPNetworkRouting = False) Then
        Exit Function
    End If
    
''    Dim pRasterFlowDir As IRaster
''    Set pRasterFlowDir = OpenRasterDatasetFromDisk("FlowDir")
''    If (pRasterFlowDir Is Nothing) Then
''         EnableAutoBMPNetworkRouting = False
''         Exit Function
''    End If
''    Set pRasterFlowDir = Nothing
''
''    Dim pRasterFlowAccu As IRaster
''    Set pRasterFlowAccu = OpenRasterDatasetFromDisk("FlowAccu")
''    If (pRasterFlowAccu Is Nothing) Then
''         EnableAutoBMPNetworkRouting = False
''         Exit Function
''    End If
''    Set pRasterFlowAccu = Nothing
    
    
End Function

Public Function EnableDrainageAreaBMPConnection() As Boolean
    EnableDrainageAreaBMPConnection = EnableBMPNetworkRouting
    If (EnableDrainageAreaBMPConnection = False) Then
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    If (pFeatureLayer Is Nothing) Then
        EnableDrainageAreaBMPConnection = False
    Else
        EnableDrainageAreaBMPConnection = True
    End If
    Set pFeatureLayer = Nothing
End Function

Public Function EnableDeleteBMP() As Boolean
    'Cannot Delete bmp unless templates defined and bmps added
    EnableDeleteBMP = EnableAddBMPOnLand
    If (EnableDeleteBMP = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPDetail")
    If (pTable Is Nothing) Then
        Set pTable = GetInputDataTable("AgBMPDetail")
        If pTable Is Nothing Then
            EnableDeleteBMP = False
            Exit Function
        End If
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        EnableDeleteBMP = False
        Exit Function
    End If
    EnableDeleteBMP = True
End Function

Public Function EnableEditAssessPoints() As Boolean
    'Cannot Edit assessment points unless BMPs can be editable
    EnableEditAssessPoints = EnableEditBMP
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    If (pTable Is Nothing) Then
        EnableEditAssessPoints = False
        Exit Function
    End If
    EnableEditAssessPoints = True
End Function

Public Function EnableDelineation() As Boolean
    If Not EnableSustain Then
        EnableDelineation = False
        Exit Function
    End If
    InitializeMapDocument
    'Delineation can be performed if bmps defined
    EnableDelineation = ValidateDataSource
    If (EnableDelineation = False) Then
        Exit Function
    End If
    
'    Dim pTable As IFeatureLayer
'    Set pTable = GetInputFeatureLayer("BMPs")
'
'    If (pTable Is Nothing) Then
'        EnableDelineation = False
'    Else
'        EnableDelineation = True
'    End If
End Function
    
Public Function EnableLUReclassification() As Boolean
    InitializeMapDocument
    'EnableLUReclassification = ModuleUtility.ValidateDataSource
    If gDefLayers And gInternalSimulation Then EnableLUReclassification = True
End Function

Public Function EnableTimeSeriesFactors() As Boolean
    If Not EnableSustain Then
        EnableTimeSeriesFactors = False
        Exit Function
    End If
    Dim pTable As IFeatureLayer
    Set pTable = GetInputFeatureLayer("BasinRouting")
    If (pTable Is Nothing) Then
        EnableTimeSeriesFactors = False
        Exit Function
    End If
    EnableTimeSeriesFactors = True
End Function
           
Public Function EnableSchematicLayer() As Boolean
    EnableSchematicLayer = EnableEditBMP
    If (EnableSchematicLayer = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPNetwork")
    If (pTable Is Nothing) Then
        EnableSchematicLayer = False
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        EnableSchematicLayer = False
        Exit Function
    End If
    Set pFeatureLayer = GetInputFeatureLayer("Conduits")
    If (pFeatureLayer Is Nothing) Then
        EnableSchematicLayer = False
    End If
    EnableSchematicLayer = True

End Function

Public Function EnableToggleSchematicLayer() As Boolean
    EnableToggleSchematicLayer = EnableSchematicLayer
    If (EnableToggleSchematicLayer = False) Then
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Schematic BMPs")
    If (pFeatureLayer Is Nothing) Then
        EnableToggleSchematicLayer = False
        Exit Function
    End If
    Set pFeatureLayer = GetInputFeatureLayer("Schematic Route")
    If (pFeatureLayer Is Nothing) Then
        EnableToggleSchematicLayer = False
        Exit Function
    End If
    EnableToggleSchematicLayer = True
    Set pFeatureLayer = Nothing
End Function

Public Function EnableDefineAssessPoints() As Boolean
    EnableDefineAssessPoints = EnableDelineation
    If (EnableDefineAssessPoints = False) Then
        Exit Function
    End If
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPNetwork")
    If (pTable Is Nothing) Then
        EnableDefineAssessPoints = False
        Exit Function
    End If
    Set pTable = GetInputDataTable("BMPDetail")
    If (pTable Is Nothing) Then
        EnableDefineAssessPoints = False
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        EnableDefineAssessPoints = False
        Exit Function
    End If
    Set pFeatureLayer = GetInputFeatureLayer("Conduits")
    If (pFeatureLayer Is Nothing) Then
        EnableDefineAssessPoints = False
        Exit Function
    End If
    Set pFeatureLayer = GetInputFeatureLayer("BasinRouting")
    If (pFeatureLayer Is Nothing) Then
        EnableDefineAssessPoints = False
        Exit Function
    End If
End Function

Public Function EnableCreateInputFile() As Boolean
    EnableCreateInputFile = False
    If Not EnableSustain Then Exit Function
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    If pTable Is Nothing Then
        EnableCreateInputFile = False
        Exit Function
    End If
        
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID > 0 "
    Dim pCount As Integer
    pCount = pTable.RowCount(pQueryFilter)
    If (pCount = 0) Then
        EnableCreateInputFile = False
        Exit Function
    End If
    'Disable if stop delta, etc values not entered for option 1 and 2
    pQueryFilter.WhereClause = "PropName = 'Option' AND PropValue <> '0'"
    pCount = pTable.RowCount(pQueryFilter)
    If (pCount > 0) Then
        pQueryFilter.WhereClause = "PropName = 'StopDelta' AND PropValue = '-99'"
        pCount = pTable.RowCount(pQueryFilter)
        If (pCount = 1) Then
            pQueryFilter.WhereClause = "PropName = 'NumBreak' AND PropValue > '0'"
            pCount = pTable.RowCount(pQueryFilter)
            If pCount <= 0 Then
                EnableCreateInputFile = False
                Exit Function
            End If
        End If
    End If
    Set pQueryFilter = Nothing
    Set pTable = Nothing
    EnableCreateInputFile = True

End Function

Public Function EnableEditInputFile() As Boolean
    EnableEditInputFile = EnableDefineAssessPoints
    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    If (pTable Is Nothing) Then
        EnableEditInputFile = False
        Exit Function
    End If
    EnableEditInputFile = True

End Function

Public Function EnableSimulation() As Boolean
    If Not EnableSustain Then Exit Function
    EnableSimulation = EnableEditInputFile
  
End Function


Public Function EnableOptimization() As Boolean
    EnableOptimization = EnableSimulation
    If (EnableOptimization = False) Then
        Exit Function
    End If
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    If (pTable Is Nothing) Then
        EnableOptimization = False
        Exit Function
    End If

End Function


Public Function EnableOptimizationForGivenScenario(poption As Integer) As Boolean
    EnableOptimizationForGivenScenario = EnableOptimization
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "PropName = 'Option' AND PropValue = '" & poption & "'"
    Dim pCount As Integer
    pCount = pTable.RowCount(pQueryFilter)
    If (pCount = 0) Then
       EnableOptimizationForGivenScenario = False
        Exit Function
    End If
    pQueryFilter.WhereClause = "ID > 0 "
    pCount = pTable.RowCount(pQueryFilter)
    If (pCount = 0) Then
       EnableOptimizationForGivenScenario = False
        Exit Function
    End If
    Set pQueryFilter = Nothing
    Set pTable = Nothing
    EnableOptimizationForGivenScenario = True
End Function


Public Function EnableAutoDelineation() As Boolean
    EnableAutoDelineation = EnableDelineation
'    If (gMap Is Nothing) Then
'        MsgBox "gmap is nothing"
'    End If
    Dim pTable As IFeatureLayer
    Set pTable = GetInputFeatureLayer("BMPs")

    If (pTable Is Nothing) Then
        EnableAutoDelineation = False
        Exit Function
    End If
    Dim pRasterDemLayer As IRasterLayer
    Set pRasterDemLayer = GetInputRasterLayer("DEM")
    If (pRasterDemLayer Is Nothing) Then
        EnableAutoDelineation = False
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("STREAM")
    If (pFeatureLayer Is Nothing) Then
        EnableAutoDelineation = False
        Exit Function
    End If
End Function

Public Function EnableManualDelineation() As Boolean
    EnableManualDelineation = EnableDelineation
End Function

Public Function EnablePostProcessor() As Boolean
    EnablePostProcessor = EnableSimulationOption
End Function



'***************************************************************
'   SWMM Menu Enable/Disable Functions
'***************************************************************

Public Function EnableSWMMMeterologicalData() As Boolean
    '** set it to true to begin with
    EnableSWMMMeterologicalData = gInternalSimulation
    If (EnableSWMMMeterologicalData = False) Then
        Exit Function
    End If
    EnableSWMMMeterologicalData = EnableSimulationOption
    If (EnableSWMMMeterologicalData = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LUReclass")
    If (pTable Is Nothing) Then
        EnableSWMMMeterologicalData = False
    End If
    '** cleanup
    Set pTable = Nothing
End Function


Public Function EnableSWMMPollutantProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMPollutantProperties = gInternalSimulation
    If EnableSWMMPollutantProperties = False Then Exit Function
    EnableSWMMPollutantProperties = EnableSWMMMeterologicalData
    If (EnableSWMMPollutantProperties = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDClimatology")
    If (pTable Is Nothing) Then
        EnableSWMMPollutantProperties = False
    End If
    '** cleanup
    Set pTable = Nothing
End Function


Public Function EnableSWMMLanduseReclassify() As Boolean
    '** set it to true to begin with
    EnableSWMMLanduseReclassify = gInternalSimulation
    If EnableSWMMLanduseReclassify = False Then Exit Function
    EnableSWMMLanduseReclassify = EnableLUReclassification
    If (EnableSWMMLanduseReclassify = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDClimatology")
    If (pTable Is Nothing) Then
        EnableSWMMLanduseReclassify = False
    End If
    '** cleanup
    Set pTable = Nothing
End Function


Public Function EnableSWMMLanduseProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMLanduseProperties = gInternalSimulation
    If EnableSWMMLanduseProperties = False Then Exit Function
    EnableSWMMLanduseProperties = EnableSWMMPollutantProperties
    If (EnableSWMMLanduseProperties = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDPollutants")
    If (pTable Is Nothing) Then
        EnableSWMMLanduseProperties = False
    End If
    '** cleanup
    Set pTable = Nothing
End Function


Public Function EnableSWMMRainGageProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMRainGageProperties = gInternalSimulation
    If EnableSWMMRainGageProperties = False Then Exit Function
    EnableSWMMRainGageProperties = EnableSWMMLanduseProperties
    If (EnableSWMMRainGageProperties = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDLanduses")
    If (pTable Is Nothing) Then
        EnableSWMMRainGageProperties = False
    End If
    '** cleanup
    Set pTable = Nothing
    
End Function

Public Function EnableSWMMAquiferProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMAquiferProperties = gInternalSimulation
    If EnableSWMMAquiferProperties = False Then Exit Function
    EnableSWMMAquiferProperties = EnableSWMMRainGageProperties
    If (EnableSWMMAquiferProperties = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDRainGages")
    If (pTable Is Nothing) Then
        EnableSWMMAquiferProperties = False
    End If
    '** cleanup
    Set pTable = Nothing
End Function


Public Function EnableSWMMSnowpackProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMSnowpackProperties = gInternalSimulation
    If EnableSWMMSnowpackProperties = False Then Exit Function
    EnableSWMMSnowpackProperties = EnableSWMMAquiferProperties
    If (EnableSWMMSnowpackProperties = False) Then
        Exit Function
    End If
    'Comment following check - SP, March 2009
    '** check if swmm options table is defined
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("LANDAquifers")
''    If (pTable Is Nothing) Then
''        EnableSWMMSnowpackProperties = False
''    End If
    '** cleanup
''    Set pTable = Nothing
End Function


Public Function EnableSWMMWatershedProperties() As Boolean
    '** set it to true to begin with
    EnableSWMMWatershedProperties = gInternalSimulation
    If EnableSWMMWatershedProperties = False Then Exit Function
    EnableSWMMWatershedProperties = EnableSWMMSnowpackProperties
    If (EnableSWMMWatershedProperties = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDRainGages")
    If (pTable Is Nothing) Then
        EnableSWMMWatershedProperties = False
    End If
    'Comment the following checks - SP, March 2009
    ' ** Check if Aquifers & Snow Packs.....
''    Set pTable = GetInputDataTable("LANDAquifers")
''    If (pTable Is Nothing) Then
''        EnableSWMMWatershedProperties = False
''    End If
''    Set pTable = GetInputDataTable("LANDSnowPacks")
''    If (pTable Is Nothing) Then
''        EnableSWMMWatershedProperties = False
''    End If
    
    '** check if Watershed feature layer is defined
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    If (pFeatureLayer Is Nothing) Then
        EnableSWMMWatershedProperties = False
    End If
        
    '** check if BMPs feature layer is defined
'''    Set pFeatureLayer = Nothing
'''    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
'''    If (pFeatureLayer Is Nothing) Then
'''        EnableSWMMWatershedProperties = False
'''    End If
'''
'''    '** check if BasinRouting feature layer is defined
'''    Set pFeatureLayer = Nothing
'''    Set pFeatureLayer = GetInputFeatureLayer("BasinRouting")
'''    If (pFeatureLayer Is Nothing) Then
'''        EnableSWMMWatershedProperties = False
'''    End If
    
    '** cleanup
    Set pTable = Nothing
    'Set pFeatureLayer = Nothing
    
End Function


Public Function EnableSWMMSimulationOptions() As Boolean
    '** set it to true to begin with
    EnableSWMMSimulationOptions = gInternalSimulation
    If EnableSWMMSimulationOptions = False Then Exit Function
    EnableSWMMSimulationOptions = EnableSWMMWatershedProperties
    If (EnableSWMMSimulationOptions = False) Then
        Exit Function
    End If
    '** check if swmm options table is defined
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDSubCatchments")
    If (pTable Is Nothing) Then
        EnableSWMMSimulationOptions = False
    End If
    '** cleanup
    Set pTable = Nothing
    
End Function

Public Function EnableSWMMViewInputFile() As Boolean
    EnableSWMMViewInputFile = gInternalSimulation
    If (EnableSWMMViewInputFile = False) Then
        Exit Function
    End If
    EnableSWMMViewInputFile = EnableLUReclassification
End Function

Public Function EnableSWMMRunSimulation() As Boolean
    EnableSWMMRunSimulation = gInternalSimulation
    If (EnableSWMMRunSimulation = False) Then
        Exit Function
    End If
    EnableSWMMRunSimulation = EnableLUReclassification
End Function


Public Function EnableLoadData() As Boolean
  On Error GoTo ErrorHandler
    EnableLoadData = False
    If Not EnableSustain Then Exit Function
    
    If gGDBFlag Then EnableLoadData = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
End Function

Public Function EnableDefPollutantsInternal() As Boolean
    On Error GoTo ErrorHandler

    EnableDefPollutantsInternal = False
    If Not EnableSustain Then Exit Function
    If gDefLayers And gInternalSimulation Then EnableDefPollutantsInternal = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
    
End Function
Public Function EnableDefPollutantsExternal() As Boolean
    On Error GoTo ErrorHandler

    EnableDefPollutantsExternal = False
    If Not EnableSustain Then Exit Function
    
    If gDefLayers And gExternalSimulation Then EnableDefPollutantsExternal = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
    
End Function

Public Function EnableAggLuDistribution() As Boolean
On Error GoTo ErrorHandler

    'only enable if it is external simulations
    EnableAggLuDistribution = EnableDefPollutantsExternal
''    Dim pTable As iTable
    
''    If EnableDefPollutantsExternal Then
''        Set pTable = GetInputDataTable("TSAssigns")
''    Else
''        Set pTable = GetInputDataTable("LUReclass")
''    End If

    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    
    'If Not (pTable Is Nothing Or pBMPFLayer Is Nothing) Then
    If Not pBMPFLayer Is Nothing Then
        Dim pBMPFClass As IFeatureClass
        Set pBMPFClass = pBMPFLayer.FeatureClass
        
        If Not pBMPFClass Is Nothing Then
            Dim pQueryFilter As IQueryFilter
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "TYPE = 'Aggregate'"
            
            If pBMPFClass.FeatureCount(pQueryFilter) > 0 Then
                EnableAggLuDistribution = True
            End If
       End If
    End If
    
  Exit Function
ErrorHandler:
  MsgBox "Error in EnableAggLuDistribution :" & Err.description
End Function
Public Function EnableSimulationOption() As Boolean
    On Error GoTo ErrorHandler

    EnableSimulationOption = False
    If Not EnableSustain Then Exit Function
    
    If gDefLayers Then EnableSimulationOption = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
    
End Function

Public Function EnableInternalSimulation() As Boolean
    On Error GoTo ErrorHandler

    EnableInternalSimulation = False
    If gInternalSimulation Then EnableInternalSimulation = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
    
End Function

Public Function EnableExternalSimulation() As Boolean
    On Error GoTo ErrorHandler

    EnableExternalSimulation = False
    If gExternalSimulation Then EnableExternalSimulation = True

  Exit Function
ErrorHandler:
  MsgBox Err.description
    
End Function
