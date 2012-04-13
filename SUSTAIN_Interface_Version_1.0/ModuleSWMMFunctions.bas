Attribute VB_Name = "ModuleSWMMFunctions"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleSWMMFunctions
'   Purpose:     Add BMPs on the map, open corresponding bmp dialog,
'                and prompt user to enter bmp parameters.
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  05/11/2005 - Mira Chokshi
'
'******************************************************************************
Option Explicit

Public gRainGaugeID As Integer
Public gSubCatchmentID As Integer
Public gRainGaugePoint As IPoint
Private m_pFile As TextStream

Public Sub ActivateRainGaugeTool()
    Dim pbars As ICommandBars
    Set pbars = gApplication.Document.CommandBars
       
    'Make the define assessment point tool as active tool
    Dim pUID As UID
    Set pUID = New UID
    'Below is a toolbar in this project (BMPTool..."
    pUID.value = "SUSTAIN.SWMMDefineRainGages"
    
    Dim pSelectTool As ICommandItem
    Set pSelectTool = pbars.Find(pUID)
     
    'Set the current tool of the application to be the Select Graphics Tool
    Set gApplication.CurrentTool = pSelectTool

End Sub


Public Sub ActivateSubCatchmentTool()
    Dim pbars As ICommandBars
    Set pbars = gApplication.Document.CommandBars
    
    'Make the define assessment point tool as active tool
    Dim pUID As UID
    Set pUID = New UID
    'Below is a toolbar in this project (Define Subcatchment tool..."
    pUID.value = "SUSTAIN.SWMMDefineSubBasinProperties"
    
    Dim pSelectTool As ICommandItem
    Set pSelectTool = pbars.Find(pUID)
     
    'Set the current tool of the application to be the Select Graphics Tool
    Set gApplication.CurrentTool = pSelectTool

End Sub

Public Sub AddRainGaugeStation(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

On Error GoTo ShowError

    'Get rain gauge feature layer, if not found, create feature class for it
    Dim pRainGaugeFeatureLayer As IFeatureLayer
    Set pRainGaugeFeatureLayer = GetInputFeatureLayer("Rain Gauges")

    'Get feature class, if not found, create it, create a feature layer from it
    Dim pRainGaugeFeatureClass As IFeatureClass
    If (pRainGaugeFeatureLayer Is Nothing) Then
        Set pRainGaugeFeatureClass = CreatePointFeatureClassForSchematics(gMapTempFolder, "raingauge")
        Set pRainGaugeFeatureLayer = New FeatureLayer
        Set pRainGaugeFeatureLayer.FeatureClass = pRainGaugeFeatureClass
        AddLayerToMap pRainGaugeFeatureLayer, "Rain Gauges"
    End If
    Set pRainGaugeFeatureClass = pRainGaugeFeatureLayer.FeatureClass

    Dim pDispTrans As IDisplayTransformation
    Dim pActView As IActiveView
    Set pActView = gMxDoc.ActiveView
    Dim pDisp As IScreenDisplay
    Set pDisp = pActView.ScreenDisplay
    Set pDispTrans = pDisp.DisplayTransformation
     
    'Set the global rain gauge point to update
     Set gRainGaugePoint = pDispTrans.ToMapPoint(X, Y)

    Dim pRGID As Integer
    pRGID = pRainGaugeFeatureClass.FeatureCount(Nothing) + 1
   
    '** cleanup
    Set pRainGaugeFeatureClass = Nothing
    Set pRainGaugeFeatureLayer = Nothing
    Set pDispTrans = Nothing
    Set pActView = Nothing
    Set pDisp = Nothing
   
    'Open the rain gauge form, to receive additional information
    FrmSWMMDefineRainGauges.txtName.Text = "RG" & pRGID
    FrmSWMMDefineRainGauges.txtRainGaugeID.Text = pRGID
    FrmSWMMDefineRainGauges.Show vbModal

    GoTo CleanUp
ShowError:
    MsgBox "AddRainGaugeStation: " & Err.description
CleanUp:
    Set pRainGaugeFeatureLayer = Nothing
    Set pRainGaugeFeatureClass = Nothing
    Set pDispTrans = Nothing
    Set pActView = Nothing
    Set pDisp = Nothing
    'Set pMapPoint = Nothing
    'Set pFeature = Nothing
End Sub


'** Render rain gauge stations
Public Sub RenderRainGaugeLayer(pFeatureLayer As IFeatureLayer)
On Error GoTo ShowError

    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    pColor.RGB = vbBlack
    
    Dim pPictureMarkerSymbol As IPictureMarkerSymbol
    Set pPictureMarkerSymbol = New PictureMarkerSymbol
    '** Create the Markers and assign their properties.
    With pPictureMarkerSymbol
       Set .Picture = LoadResPicture("RainGauge", vbResBitmap)
      .Angle = 0
      .Size = 15
    End With
   
    Dim pSimpleRenderer As ISimpleRenderer
    Set pSimpleRenderer = New SimpleRenderer
    Set pSimpleRenderer.Symbol = pPictureMarkerSymbol
   
    Dim pGeoFeatLyr As IGeoFeatureLayer
    Set pGeoFeatLyr = pFeatureLayer
    Set pGeoFeatLyr.Renderer = pSimpleRenderer
    
     ' setup LabelEngineProperties for the FeatureLayer
     ' get the AnnotateLayerPropertiesCollection for the FeatureLayer
     Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
     Set pAnnoLayerPropsColl = pGeoFeatLyr.AnnotationProperties
     pGeoFeatLyr.DisplayAnnotation = True
     pAnnoLayerPropsColl.Clear
     ' create a new LabelEngineLayerProperties object
     Dim aLELayerProps As ILabelEngineLayerProperties
     Set aLELayerProps = New LabelEngineLayerProperties
     aLELayerProps.IsExpressionSimple = True
     aLELayerProps.Expression = "[LABEL]"

     ' assign it to the layer's AnnotateLayerPropertiesCollection
     pAnnoLayerPropsColl.add aLELayerProps
    
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     '** Draw the map
     gMxDoc.ActiveView.Refresh
    
    GoTo CleanUp
    
ShowError:
    MsgBox "RenderRainGaugeLayer: " & Err.description
CleanUp:
    Set pColor = Nothing
    Set pPictureMarkerSymbol = Nothing
    Set pSimpleRenderer = Nothing
    Set pGeoFeatLyr = Nothing
End Sub



Public Sub SaveSWMMOptionsToTable(pPropertyDictionary As Scripting.Dictionary)

On Error GoTo ShowError

    'Find the SWMM option table
    Dim pSWMMOptionsTable As iTable
    Set pSWMMOptionsTable = GetInputDataTable("LANDOptions")
    
    'Create the table if not found, add it to the Map
    If (pSWMMOptionsTable Is Nothing) Then
        Set pSWMMOptionsTable = CreatePropertiesTableDBF("LANDOptions")
        AddTableToMap pSWMMOptionsTable
    Else
        pSWMMOptionsTable.DeleteSearchedRows Nothing
    End If

    Dim iPropName As Long
    Dim iPropValue As Long
    iPropName = pSWMMOptionsTable.FindField("PropName")
    iPropValue = pSWMMOptionsTable.FindField("PropValue")
    
    Dim pKeys
    pKeys = pPropertyDictionary.keys
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pRow As iRow
    Dim i As Integer
    
    'Iterate over the property dictionary, and save the values
    For i = 0 To (pPropertyDictionary.Count - 1)
        pPropertyName = pKeys(i)
        pPropertyValue = pPropertyDictionary.Item(pPropertyName)
        Set pRow = pSWMMOptionsTable.CreateRow
        pRow.value(iPropName) = pPropertyName
        pRow.value(iPropValue) = pPropertyValue
        pRow.Store
    Next
    GoTo CleanUp
    
ShowError:
    MsgBox "SaveSWMMOptionsToTable: " & Err.description
CleanUp:
    Set pSWMMOptionsTable = Nothing
    Set pRow = Nothing
    Set pPropertyDictionary = Nothing
End Sub


'*** Create a SWMM Property table and save the values
Public Sub SaveSWMMPropertiesTable(pTableName As String, pID As String, pPropertyDictionary As Scripting.Dictionary)

On Error GoTo ShowError

    'Find the SWMM option table
    Dim pSWMMPropertyTable As iTable
    Set pSWMMPropertyTable = GetInputDataTable(pTableName)
    
    'Create the table if not found, add it to the Map
    If (pSWMMPropertyTable Is Nothing) Then
        Set pSWMMPropertyTable = CreatePropertiesTableDBF(pTableName)
        AddTableToMap pSWMMPropertyTable
    End If

    Dim iPropName As Long
    Dim iPropValue As Long
    Dim iID As Long
    
    iPropName = pSWMMPropertyTable.FindField("PropName")
    iPropValue = pSWMMPropertyTable.FindField("PropValue")
    iID = pSWMMPropertyTable.FindField("ID")
    
    Dim pKeys
    pKeys = pPropertyDictionary.keys
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim i As Integer
    
    'Iterate over the property dictionary, and save the values
    For i = 0 To (pPropertyDictionary.Count - 1)
        pPropertyName = pKeys(i)
        pPropertyValue = pPropertyDictionary.Item(pPropertyName)
        pQueryFilter.WhereClause = "ID = " & pID & " AND PropName = '" & pPropertyName & "'"
        Set pCursor = pSWMMPropertyTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pSWMMPropertyTable.CreateRow
        End If
        pRow.value(iID) = pID
        pRow.value(iPropName) = pPropertyName
        pRow.value(iPropValue) = pPropertyValue
        pRow.Store
    Next
    GoTo CleanUp
    
ShowError:
    MsgBox "SaveSWMMPropertiesTable: " & Err.description
CleanUp:
    Set pQueryFilter = Nothing
    Set pSWMMPropertyTable = Nothing
    Set pRow = Nothing
    Set pPropertyDictionary = Nothing
End Sub


Public Function LoadPollutantNames() As Collection

    '** call the load property names function to get all pollutant names
   Set LoadPollutantNames = LoadPropertyNames("LANDPollutants")
   
End Function


Public Function LoadLanduseNames() As Collection
    
    '** call the load property names function to get all landuse names
   Set LoadLanduseNames = LoadPropertyNames("LANDLanduses")
   
End Function


Public Function LoadRainGaugeNames() As Collection
    
    '** call the load property names function to get all rain gauges names
   Set LoadRainGaugeNames = LoadPropertyNames("LANDRainGages")
   
End Function

Public Function LoadAquiferNames() As Collection
    
    '** call the load property names function to get all rain gauges names
   Set LoadAquiferNames = LoadPropertyNames("LANDAquifers")
   
End Function

Public Function LoadSnowPackNames() As Collection
    
    '** call the load property names function to get all rain gauges names
   Set LoadSnowPackNames = LoadPropertyNames("LANDSnowPacks")
   
End Function

Public Function LoadTransectNames() As Collection
    
    '** call the load property names function to get all rain gauges names
   Set LoadTransectNames = LoadPropertyNames("Transects")
   
End Function

Public Function LoadPropertyNames(pTableName As String) As Collection

    '* Load the list box with property names
    Dim pSWMMPropertyTable As iTable
    Set pSWMMPropertyTable = GetInputDataTable(pTableName)
    If (pSWMMPropertyTable Is Nothing) Then
        Exit Function
    End If
    
    Dim iID As Long
    iID = pSWMMPropertyTable.FindField("ID")
    Dim iPropValue As Long
    iPropValue = pSWMMPropertyTable.FindField("PropValue")
    
    'Define query filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    
    'Get the cursor to iterate over the table
    Dim pCursor As ICursor
    Set pCursor = pSWMMPropertyTable.Search(pQueryFilter, True)
    
    'Define a row variable to loop over the table
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim pCollection As Collection
    Set pCollection = New Collection
    Do While Not pRow Is Nothing
        pCollection.add Trim(pRow.value(iPropValue))
        Set pRow = pCursor.NextRow
    Loop
    
    
    '** Cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPropertyTable = Nothing
    Set pQueryFilter = Nothing
  
    '** Return the value
    Set LoadPropertyNames = pCollection
End Function


'** write swmm related project details in an input file
Public Sub WriteSWMMProjectDetails(pInputFileName As String)
On Error GoTo ErrorHandler

    If (pInputFileName = "") Then
        MsgBox "LAND Simulation Input file not defined"
        Exit Sub
    End If
    
    '** define file system object for creating file
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '** define variable to control file input
    Set m_pFile = fso.OpenTextFile(pInputFileName, ForWriting, True, TristateUseDefault)

    '** write TITLE, OPTIONS, FILES information
    Call WriteSWMMTitleOptionsAndFiles
    
    '** Write Climate details.......
    Call WriteSWMMClimateInfo
    
    '** write RAIN GAGE information
    Call WriteSWMMRainGaugeInfo

    '** write SUBCATCHMENT, SUBAREA, INFILTRATION information
    Call WriteSWMMSubCatchmentInfo
    
    ' ** Write the AQUIFERS information.....
    Call WriteSWMMAquifersinfo
    
    ' ** Write the GROUNDWATER information.....
    Call WriteSWMMGroundwaterinfo
    
    ' ** Write the SNOWPACKS information.....
    Call WriteSWMMSnowPackinfo
    
    '** write JUNCTIONS information
    Call WriteSWMMJunctionInfo
    
    '** write CONDUITS, XSECTIONS information
    'Call WriteSWMMConduitsInfo
    
    '** write POLLUTANTS information
    Call WriteSWMMPollutantInfo
    
    '** write LANDUSES information
    Call WriteSWMMLanduseInfo
    
    '** write subcatchment COVERAGES, LOADINGS information
    Call WriteSWMMCoveragesAndLoadingsInfo
        
    '** write BUILDUP, WASHOFF information
    Call WriteSWMMBuildupWashoffInfo
    
    '** write additional OPTIONS information
    'Call WriteSWMMAdditionalOptionsInfo
    
    '** close the input file
    m_pFile.Close
    
    '** cleanup
    Set m_pFile = Nothing
    Set fso = Nothing
    
    '** display the message
    MsgBox "Land Simulation Input file created !"
    Exit Sub
    
ErrorHandler:
    MsgBox "WriteSWMMProjectDetails: " & Err.description

End Sub

'** write swmm related project details in an input file
Public Sub WriteSWMMPredevelopedLanduseFile(pPredevInputFile As String, pPreDevLanduse As String, _
                                            pConductivity As Double, pSuctionHead As Double, pInitialDef As Double)
On Error GoTo ErrorHandler
    
    '** define file system object for creating file
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '** define variable to control file input
    Set m_pFile = fso.OpenTextFile(pPredevInputFile, ForWriting, True, TristateUseDefault)

    '** write TITLE, OPTIONS, FILES information
    '0: UPDATED FOR PREDEVELOPED LANDUSE CONDITION
    Call WriteSWMMPreDevelopedTitleOptionsAndFiles(pPredevInputFile)
    
    '** Write Climate details.......
    Call WriteSWMMClimateInfo
    
    '** write RAIN GAGE information
    Call WriteSWMMRainGaugeInfo

    '** write SUBCATCHMENT, SUBAREA, INFILTRATION information
    '1: UPDATED FOR PREDEVELOPED LANDUSE CONDITION
    Call WriteSWMMPreDevelopedSubCatchmentInfo(pConductivity, pSuctionHead, pInitialDef)
    
    ' ** Write the AQUIFERS information.....
    Call WriteSWMMAquifersinfo
    
    ' ** Write the GROUNDWATER information.....
    Call WriteSWMMGroundwaterinfo
    
    ' ** Write the SNOWPACKS information.....
    Call WriteSWMMSnowPackinfo
    
    '** write JUNCTIONS information
    Call WriteSWMMJunctionInfo
    
    '** write CONDUITS, XSECTIONS information
    'Call WriteSWMMConduitsInfo
    
    '** write POLLUTANTS information
    Call WriteSWMMPollutantInfo
    
    '** write LANDUSES information
    '2: UPDATED FOR PREDEVELOPED LANDUSE CONDITION
    Call WriteSWMMPreDevelopedLanduseInfo(pPreDevLanduse)
    
    '** write subcatchment COVERAGES, LOADINGS information
    '3: UPDATED FOR PREDEVELOPED LANDUSE CONDITION
    Call WriteSWMMPreDevelopedCoveragesAndLoadingsInfo(pPreDevLanduse)
        
    '** write BUILDUP, WASHOFF information
    '4: UPDATED FOR PREDEVELOPED LANDUSE CONDITION
    Call WriteSWMMPreDevelopedBuildupWashoffInfo(pPreDevLanduse)
    
    '** write additional OPTIONS information
    'Call WriteSWMMAdditionalOptionsInfo
    
    '** close the input file
    m_pFile.Close
    
    '** cleanup
    Set m_pFile = Nothing
    Set fso = Nothing
    
    '** display the message
    MsgBox "Predeveloped Landuse Input file created !"
    Exit Sub
    
ErrorHandler:
    MsgBox "WriteSWMMPredevelopedLanduseFile: " & Err.description

End Sub

''Private Function GetSimulationInputFile() As String
''
''    Dim pSWMMOptionsTable As iTable
''    Set pSWMMOptionsTable = GetInputDataTable("LANDOptions")
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''    pQueryFilter.WhereClause = "PropName = 'SWMM_INPUT_FILE'"
''    Dim pCursor As ICursor
''    Set pCursor = pSWMMOptionsTable.Search(pQueryFilter, True)
''    Dim iPropValue As Long
''    iPropValue = pSWMMOptionsTable.FindField("PropValue")
''    Dim pRow As iRow
''    Set pRow = pCursor.NextRow
''
''    Dim pSimulationInputFile As String
''    pSimulationInputFile = ""
''
''    If Not (pRow Is Nothing) Then
''          pSimulationInputFile = pRow.value(iPropValue)
''    End If
''
''    '** cleanup
''    Set pRow = Nothing
''    Set pCursor = Nothing
''    Set pQueryFilter = Nothing
''    Set pSWMMOptionsTable = Nothing
''
''    '** return the input file
''    GetSimulationInputFile = pSimulationInputFile
''End Function
'** Write project title, options, and file path
Private Sub WriteSWMMTitleOptionsAndFiles()

    Dim pSWMMOptionsTable As iTable
    Set pSWMMOptionsTable = GetInputDataTable("LANDOptions")
    Dim pCursor As ICursor
    Set pCursor = pSWMMOptionsTable.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim iPropName As Long
    iPropName = pSWMMOptionsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMOptionsTable.FindField("PropValue")
    Dim pOutflowFile As String
    Dim pPostoutFile As String
        
    '** iterate over the table to write OPTIONS value
    m_pFile.WriteLine "[OPTIONS]"
    Do While Not (pRow Is Nothing)
        If (pRow.value(iPropName) = "SAVE OUTFLOWS") Then
            pOutflowFile = pRow.value(iPropValue)
        ElseIf (pRow.value(iPropName) = "SAVE POST OUTFLOWS") Then
            pPostoutFile = pRow.value(iPropValue)
        ElseIf (pRow.value(iPropName) <> "REPORT CONTROL") Then
            m_pFile.WriteLine pRow.value(iPropName) & vbTab & pRow.value(iPropValue)
        End If
        Set pRow = pCursor.NextRow
    Loop
    m_pFile.WriteLine ""
    
    '** write the outflow file name
    m_pFile.WriteLine "[FILES]"
    m_pFile.WriteLine "SAVE OUTFLOWS" & "   " & pPostoutFile 'pOutflowFile
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMOptionsTable = Nothing

End Sub



'** Write rain SWMMClimateInfo
Private Sub WriteSWMMClimateInfo()

    '** get rain gauge info table
    Dim pSWMMRainGagesTable As iTable
    Set pSWMMRainGagesTable = GetInputDataTable("LANDClimatology")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMRainGagesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMRainGagesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMRainGagesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pOutflowFile As String
    Dim pPropertyDict As Scripting.Dictionary
    Dim pClimateString As String
    '** iterate over the table to write OPTIONS value
    Set pCursor = pSWMMRainGagesTable.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    Set pPropertyDict = CreateObject("Scripting.Dictionary")
    Do While Not pRow Is Nothing
        pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    Loop
    
    
    '** write header information for Rain Gauges
    m_pFile.WriteLine "[EVAPORATION]"
    m_pFile.WriteLine ";;"
    m_pFile.WriteLine ";;" & "Type" & vbTab & "Parameters"
    m_pFile.WriteLine ";;"
    '** write values from dictionary to string variable
    pClimateString = Replace(Replace(pPropertyDict.Item("Evaporation"), ",", "  "), ":", "")
    '** write string variable to the input file.
    m_pFile.WriteLine pClimateString
        
    m_pFile.WriteLine ""
    If pPropertyDict.Item("Temperature") = "NO DATA" Then Exit Sub
    
    '** write header information for Temperature.....
    pClimateString = Replace(Replace(pPropertyDict.Item("Windspeed"), ",", "  "), ":", "")
    If StringContains(pClimateString, "FILE") Then pClimateString = "FILE"
    m_pFile.WriteLine "[TEMPERATURE]"
    '** write values from dictionary to string variable
    pClimateString = Replace(pPropertyDict.Item("Temperature"), "FILE:", "FILE") & vbNewLine & _
                                "WINDSPEED" & vbTab & pClimateString & vbNewLine & _
                                "SNOWMELT" & vbTab & pPropertyDict.Item("SnowDividingTemp") & vbTab & _
                                pPropertyDict.Item("ATI Weight") & vbTab & _
                                pPropertyDict.Item("Negative Melt Ratio") & vbTab & _
                                pPropertyDict.Item("Elevation") & vbTab & _
                                pPropertyDict.Item("Latitude") & vbTab & _
                                pPropertyDict.Item("Longitude Correction") & vbNewLine & _
                                "ADC         IMPERVIOUS" & vbTab & Replace(pPropertyDict.Item("ADC: IMPERVIOUS"), ",", "  ") & vbNewLine & _
                                "ADC         PERVIOUS" & vbTab & Replace(pPropertyDict.Item("ADC: PERVIOUS"), ",", "  ")
                                

    '** write string variable to the input file.
    m_pFile.WriteLine pClimateString
        
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    'Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMRainGagesTable = Nothing

End Sub


'** Write rain gauge information
Private Sub WriteSWMMRainGaugeInfo()

    '** get rain gauge info table
    Dim pSWMMRainGagesTable As iTable
    Set pSWMMRainGagesTable = GetInputDataTable("LANDRainGages")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMRainGagesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMRainGagesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMRainGagesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMRainGagesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    
    Dim pOutflowFile As String
    
    '** write header information for Rain Gauges
    m_pFile.WriteLine "[RAINGAGES]"
    m_pFile.WriteLine ";;"
    m_pFile.WriteLine ";;" & "Name" & vbTab & "Rain Type" & vbTab & "Recd. Freq" & _
                        vbTab & "Snow Catch" & vbTab & "Data Source" & vbTab & _
                        "Source Name" & vbTab & "Station ID" & vbTab & "Rain Units"
    m_pFile.WriteLine ";;"
    
    Dim pPropertyDict As Scripting.Dictionary
    Dim pRainGageString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMRainGagesTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
    
        '** write values from dictionary to string variable
        pRainGageString = pPropertyDict.Item("Name") & vbTab & _
                          pPropertyDict.Item("Rain Type") & vbTab & _
                          pPropertyDict.Item("Recd. Freq") & vbTab & _
                          pPropertyDict.Item("Snow Catch") & vbTab & _
                          pPropertyDict.Item("Data Source") & vbTab & _
                          pPropertyDict.Item("Source Name") & vbTab & _
                          pPropertyDict.Item("Station ID") & vbTab & _
                          pPropertyDict.Item("Rain Units")
        '** write string variable to the input file.
        m_pFile.WriteLine pRainGageString
        
    Next
   
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMRainGagesTable = Nothing

End Sub



'** Write subcatchment information
Private Sub WriteSWMMSubCatchmentInfo()

    '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name' ORDER BY ID"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    Set pIDCollection = SortCollection(pIDCollection, True)
        
    Dim pOutflowFile As String
    
    '** write header information for SUBCATCHMENTS
    Dim pSubCatchmentHeader As String
    Dim pSubCatchmentBody As String
    pSubCatchmentHeader = "[SUBCATCHMENTS]" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & "Name" & vbTab & "Raingage" & vbTab & "Outlet" & _
                                                vbTab & "Total Area" & vbTab & "Pcnt. Imperv" & vbTab & _
                                                "Width" & vbTab & "Pcnt. Slope" & vbTab & "Curb Length" & _
                                                vbTab & "Snow Pack" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & vbNewLine
    
    
    '** write header information for SUBAREAS
    Dim pSubAreaHeader As String
    Dim pSubAreaBody As String
    pSubAreaHeader = "[SUBAREAS]" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & "Subcatchment" & vbTab & "N-Imperv" & vbTab & "N-Perv" & _
                                            vbTab & "S-Imperv" & vbTab & "S-Perv" & vbTab & _
                                            "PctZero" & vbTab & "RouteTo" & vbTab & "PctRouted" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & vbNewLine
    
    
    '** write header information for INFILTRATION
    Dim pInfiltrationHeader As String
    Dim pInfiltrationBody As String
    pInfiltrationHeader = "[INFILTRATION]" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & "Subcatchment" & vbTab & "Suction" & vbTab & "HydCon" & _
                                            vbTab & "IMDmax" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & vbNewLine
    
    
    
    Dim pPropertyDict As Scripting.Dictionary
    Dim pRainGageString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        pPropertyDict.add "ID", pIDCollection.Item(iCount)
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
    
        '** write values for SUBCATCHMENTS
        pSubCatchmentBody = pSubCatchmentBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("Rain Gauge") & vbTab & _
                          pPropertyDict.Item("Outlet") & vbTab & _
                          pPropertyDict.Item("Area") & vbTab & _
                          pPropertyDict.Item("%Impervious") & vbTab & _
                          pPropertyDict.Item("Width") & vbTab & _
                          pPropertyDict.Item("%Slope") & vbTab & _
                          pPropertyDict.Item("Curb Length") & vbTab & _
                          pPropertyDict.Item("Snow Packs") & vbNewLine
                          
                          
        '** write values for SUBAREAS
        pSubAreaBody = pSubAreaBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("NImpervious") & vbTab & _
                          pPropertyDict.Item("NPervious") & vbTab & _
                          pPropertyDict.Item("DImpervious") & vbTab & _
                          pPropertyDict.Item("DPervious") & vbTab & _
                          pPropertyDict.Item("%ZeroImpervious") & vbTab & _
                          pPropertyDict.Item("SubareaRouting") & vbTab & _
                          pPropertyDict.Item("%Routing") & vbNewLine
                          
                          
        '** write values for INFILTRATION
        pInfiltrationBody = pInfiltrationBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("Suction Head") & vbTab & _
                          pPropertyDict.Item("Conductivity") & vbTab & _
                          pPropertyDict.Item("Initial Deficit") & vbNewLine
                                  
    Next
    
    '** write SUBCATCHMENT values to the input file.
    m_pFile.Write pSubCatchmentHeader
    m_pFile.Write pSubCatchmentBody
    m_pFile.WriteLine ""
    
    '** write SUBAREAS values to the input file.
    m_pFile.Write pSubAreaHeader
    m_pFile.Write pSubAreaBody
    m_pFile.WriteLine ""
    
    '** write INFILTRATION values to the input file.
    m_pFile.Write pInfiltrationHeader
    m_pFile.Write pInfiltrationBody
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentsTable = Nothing
    pSubCatchmentHeader = ""
    pSubCatchmentBody = ""
    pSubAreaHeader = ""
    pSubAreaBody = ""
    pInfiltrationHeader = ""
    pInfiltrationBody = ""
End Sub


'** WriteSWMMAquifersinfo
Public Sub WriteSWMMAquifersinfo()
    
     '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDAquifers")
    
    If pSWMMSubCatchmentsTable Is Nothing Then Exit Sub
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    
    Dim strFields As String
    Dim strValues As String
    Dim pFLag As Boolean
    m_pFile.WriteLine "[AQUIFERS]"
    m_pFile.WriteLine ";;"
    
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        strValues = ""
        strFields = ";;"
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            strFields = strFields & vbTab & pRow.value(iPropName)
            strValues = strValues & vbTab & pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
        
        If Not pFLag Then
            m_pFile.WriteLine strFields
            m_pFile.WriteLine ";;"
            pFLag = True
        End If
        m_pFile.WriteLine strValues
        
    Next iCount
    
    m_pFile.WriteLine ""

End Sub

'** WriteSWMMGroundwater
Public Sub WriteSWMMGroundwaterinfo()
    
     '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    Set pIDCollection = SortCollection(pIDCollection, True)
    
    ' Check if Ground water table is selected.....
    pQueryFilter.WhereClause = "PropName='Ground Water' And PropValue='1'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, False)
    If pCursor.NextRow Is Nothing Then Exit Sub
    
    Dim strFields As String
    Dim strValues As String
    Dim pFLag As Boolean
    m_pFile.WriteLine "[GROUNDWATER]"
    m_pFile.WriteLine ";;"
    
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        strValues = ""
        strFields = ";;Subcatchment"
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount) & " And PropName LIKE 'GW%'"
        Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            strFields = strFields & vbTab & Replace(pRow.value(iPropName), "GW", "")
            strValues = strValues & vbTab & pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
        
        If Not pFLag Then
            m_pFile.WriteLine strFields
            m_pFile.WriteLine ";;"
            pFLag = True
        End If
        m_pFile.WriteLine pIDCollection.Item(iCount) & vbTab & strValues
        
    Next iCount
    
    m_pFile.WriteLine ""

End Sub

'** WriteSWMMSnowPacks
Public Sub WriteSWMMSnowPackinfo()
    
    Dim pPropertyDict As Scripting.Dictionary
     '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSnowPacks")
    
    If pSWMMSubCatchmentsTable Is Nothing Then Exit Sub
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    
    Dim strValues As String
    m_pFile.WriteLine "[SNOWPACKS]"
    m_pFile.WriteLine ";;"
    
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
        
        Dim pValues(0 To 6)
        pValues(0) = Split(pPropertyDict.Item("Min Coeff"), ";")
        pValues(1) = Split(pPropertyDict.Item("Max Coeff"), ";")
        pValues(2) = Split(pPropertyDict.Item("Base Temp"), ";")
        pValues(3) = Split(pPropertyDict.Item("Fraction Capacity"), ";")
        pValues(4) = Split(pPropertyDict.Item("Initial Snow Depth"), ";")
        pValues(5) = Split(pPropertyDict.Item("Initial Free Water"), ";")
        pValues(6) = Split(pPropertyDict.Item("Depth"), ";")
    
        '** write values for Snow Pack details......
        strValues = pPropertyDict.Item("Name") & vbTab & _
                          "PLOWABLE" & vbTab & _
                          pValues(0)(0) & vbTab & _
                          pValues(1)(0) & vbTab & _
                          pValues(2)(0) & vbTab & _
                          pValues(3)(0) & vbTab & _
                          pValues(4)(0) & vbTab & _
                          pValues(5)(0) & vbTab & _
                          pValues(6)(0) & vbTab
        m_pFile.WriteLine strValues
        strValues = pPropertyDict.Item("Name") & vbTab & _
                          "IMPERVIOUS" & vbTab & _
                          pValues(0)(1) & vbTab & _
                          pValues(1)(1) & vbTab & _
                          pValues(2)(1) & vbTab & _
                          pValues(3)(1) & vbTab & _
                          pValues(4)(1) & vbTab & _
                          pValues(5)(1) & vbTab & _
                          pValues(6)(1) & vbTab
        m_pFile.WriteLine strValues
        strValues = pPropertyDict.Item("Name") & vbTab & _
                          "PERVIOUS" & vbTab & _
                          pValues(0)(2) & vbTab & _
                          pValues(1)(2) & vbTab & _
                          pValues(2)(2) & vbTab & _
                          pValues(3)(2) & vbTab & _
                          pValues(4)(2) & vbTab & _
                          pValues(5)(2) & vbTab & _
                          pValues(6)(2) & vbTab
        m_pFile.WriteLine strValues
        strValues = pPropertyDict.Item("Name") & vbTab & _
                          "REMOVAL" & vbTab & _
                          pPropertyDict.Item("Snow Depth") & vbTab & _
                          pPropertyDict.Item("Fraction Watershed") & vbTab & _
                          pPropertyDict.Item("Fraction Impervious") & vbTab & _
                          pPropertyDict.Item("Fraction Pervious") & vbTab & _
                          pPropertyDict.Item("Fraction Melt") & vbTab & _
                          pPropertyDict.Item("Fraction subcatchment") & vbTab & _
                          pPropertyDict.Item("SnowPackName") & vbTab
        m_pFile.WriteLine strValues
                
    Next iCount
    
    m_pFile.WriteLine ""

End Sub

'** Write bmp and dummy junction information
'** Each subcatchment will have a dummy junction, which will flow
'** to an actual BMP juntion. The JUNCTIONS card will include
'** information for both of these nodes. However the dummy junction values
'** will be made up from SubCatchment layer
Private Sub WriteSWMMJunctionInfo()

    '** get swmm subcatchments info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query swmmsubcatchments table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Outlet'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pDummyJunctionCollection As Collection
    Set pDummyJunctionCollection = New Collection
    Do While Not pRow Is Nothing
        pDummyJunctionCollection.add pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    Loop
    
    '** get bmp attributes info table
    Dim pBMPFeatureLayer As IFeatureLayer
    Set pBMPFeatureLayer = GetInputFeatureLayer("BMPs")
    Dim pBMPFeatureClass As IFeatureClass
    Set pBMPFeatureClass = pBMPFeatureLayer.FeatureClass
    
    '** define field indexes
    Dim iID As Long
    iID = pBMPFeatureClass.FindField("ID")
    
    '** define iterator variables
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    
    '** query bmp attributes table to get all unique ids, put them in a collection
    Set pFeatureCursor = pBMPFeatureClass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    Dim pBMPJunctionCollection As Collection
    Set pBMPJunctionCollection = New Collection
    Do While Not pFeature Is Nothing
        pBMPJunctionCollection.add pFeature.value(iID)
        Set pFeature = pFeatureCursor.NextFeature
    Loop
   
    '** write header information for Rain Gauges
    m_pFile.WriteLine "[JUNCTIONS]"
    m_pFile.WriteLine ";;"
    m_pFile.WriteLine ";;" & "Name" & vbTab & "Invert Elev." & vbTab & "Max. Depth" & _
                        vbTab & "Init. Depth" & vbTab & "Surcharge Depth" & vbTab & _
                        "Ponded Area"
    m_pFile.WriteLine ";;"
    
    '** iterate over the bmp and dummy junction collection
    '** write dummy values for each junction
    Dim iCount As Integer
    Dim pID As String
    Dim pJunctionString As String
    
    '** write values for bmp junction
    For iCount = 1 To pBMPJunctionCollection.Count
        '** for each ID, add values to a dictionary
        pID = pBMPJunctionCollection.Item(iCount)
        '** write dummy values to string variable
        pJunctionString = pID & vbTab & _
                          "1000" & vbTab & _
                          "5" & vbTab & _
                          "0" & vbTab & _
                          "0" & vbTab & _
                          "0"
        '** write string variable to the input file.
        m_pFile.WriteLine pJunctionString
    Next
    
    '** write values for dummy junction
'    For iCount = 1 To pDummyJunctionCollection.Count
'        '** for each ID, add values to a dictionary
'        pID = pDummyJunctionCollection.Item(iCount)
'        '** write dummy values to string variable
'        pJunctionString = pID & vbTab & _
'                          "1000" & vbTab & _
'                          "5" & vbTab & _
'                          "0" & vbTab & _
'                          "0" & vbTab & _
'                          "0"
'        '** write string variable to the input file.
'        m_pFile.WriteLine pJunctionString
'    Next
   
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pBMPJunctionCollection = Nothing
    Set pBMPFeatureClass = Nothing
    Set pBMPFeatureLayer = Nothing
    Set pDummyJunctionCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentsTable = Nothing
End Sub

'** Write conduits dimensions, cross-sections info
'** All conduits are going to be dummy. Since each subcatchment has an
'** actual bmp it flows into. For swmm purposes, route each subcatchment
'** to dummy node and dummy node to actual bmp. the connection between
'** dummy node to actual bmp is the dummy conduit.
Private Sub WriteSWMMConduitsInfo()

    '** get watershed feature layer
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    
    '** define field indexes
    Dim iID As Long
    iID = pWatershedFClass.FindField("ID")
    Dim iBMPID As Long
    iBMPID = pWatershedFClass.FindField("BMPID")
    
    '** define variables to access the feature layer
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature

    '** define conduit info header and cross-section header
    Dim pConduitInfoHeader As String
    Dim pConduitInfoBody As String
    Dim pCrossSectionHeader As String
    Dim pCrossSectionBody As String
    
    '** define conduit info header
    pConduitInfoHeader = "[CONDUITS]" & vbNewLine
    pConduitInfoHeader = pConduitInfoHeader & ";;" & vbNewLine
    pConduitInfoHeader = pConduitInfoHeader & _
                        ";;" & "Name" & vbTab & "Inlet Node" & vbTab & "Outlet Node" & vbTab & _
                        "Length" & vbTab & "Manning N" & vbTab & "Inlet Height" & vbTab & _
                        "Outlet Height" & vbTab & "Init. Flow" & vbNewLine
     
    '** define cross-section header
    pCrossSectionHeader = "[XSECTIONS]" & vbNewLine
    pCrossSectionHeader = pCrossSectionHeader & ";;" & vbNewLine
    pCrossSectionHeader = pCrossSectionHeader & _
                        ";;" & "Link" & vbTab & "Type" & vbTab & "Geom1" & vbTab & _
                        "Geom2" & vbTab & "Geom3" & vbTab & "Geom4" & vbTab & _
                        "Barrels" & vbNewLine
    
    
    '** query watershed attributes - watershed id, and downstream bmp id
    Set pFeatureCursor = pWatershedFClass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    Dim pWID As Integer
    Dim pBMPID As Integer
    Dim pDummyCond As String
    Dim pDummyNode As String
    Do While Not pFeature Is Nothing
        pWID = pFeature.value(iID)
        pBMPID = pFeature.value(iBMPID)
        pDummyCond = "dc" & pWID
        pDummyNode = "dn" & pWID    ' since each subcatchment flows to its own dummy node
        
        '** add the properties - dummy values
        pConduitInfoBody = pConduitInfoBody & _
                           pDummyCond & vbTab & pDummyNode & vbTab & pBMPID & vbTab & _
                           "1" & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "0" & vbNewLine
                           
        '** add cross-section properties - dummy values
        pCrossSectionBody = pCrossSectionBody & _
                            pDummyCond & vbTab & "DUMMY" & vbTab & "0" & vbTab & "0" & vbTab & _
                            "0" & vbTab & "0" & vbTab & "1" & vbNewLine
        
        '* move to next feature
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    
   
    '** write header/body information for conduit info
    m_pFile.Write pConduitInfoHeader
    m_pFile.Write pConduitInfoBody
    m_pFile.WriteLine ""
    
    '** write header/body information for conduit cross-section
    m_pFile.Write pCrossSectionHeader
    m_pFile.Write pCrossSectionBody
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pFeature = Nothing
    Set pFeatureCursor = Nothing
    Set pWatershedFClass = Nothing
    Set pWatershedFLayer = Nothing
End Sub




'** Write pollutant information
Private Sub WriteSWMMPollutantInfo()

    '** get pollutant info table
    Dim pSWMMPollutantsTable As iTable
    Set pSWMMPollutantsTable = GetInputDataTable("LANDPollutants")
    Dim pPollutantTable As iTable
    Set pPollutantTable = GetInputDataTable("Pollutants")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMPollutantsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMPollutantsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMPollutantsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMPollutantsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    
   
    '** write header information for Rain Gauges
    m_pFile.WriteLine "[POLLUTANTS]"
    m_pFile.WriteLine ";;"
    m_pFile.WriteLine ";;" & "Name" & vbTab & "Mass Units" & vbTab & "Rain Conc." & _
                        vbTab & "GW Conc." & vbTab & "I&I Conc." & vbTab & _
                        "Decay Coeff." & vbTab & "Sediment Flag." & vbTab & "Snow Only" & vbTab & "Co-Pollutant"
    m_pFile.WriteLine ";;"
    
    Dim pPropertyDict As Scripting.Dictionary
    Dim pPollutantString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMPollutantsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
        
        ' Get the Sediment Flag......
        pQueryFilter.WhereClause = "Name='" & pPropertyDict.Item("Name") & "'"
        Set pCursor = pPollutantTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        
        '** write values from dictionary to string variable
        pPollutantString = pPropertyDict.Item("Name") & vbTab & _
                          pPropertyDict.Item("Mass Units") & vbTab & _
                          pPropertyDict.Item("Rain Conc.") & vbTab & _
                          pPropertyDict.Item("GW Conc.") & vbTab & _
                          pPropertyDict.Item("I&I Conc.") & vbTab & _
                          pPropertyDict.Item("Decay Coeff.") & vbTab & _
                          pRow.value(pPollutantTable.FindField("Sediment")) & vbTab & _
                          pPropertyDict.Item("Snow Only") & vbTab & _
                          pPropertyDict.Item("Co-Pollutant")
        '** write string variable to the input file.
        m_pFile.WriteLine pPollutantString
        
    Next
   
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPollutantsTable = Nothing

End Sub


'** Write landuse information
Private Sub WriteSWMMLanduseInfo()

    '** get pollutant info table
    Dim pSWMMLandusesTable As iTable
    Set pSWMMLandusesTable = GetInputDataTable("LANDLanduses")
    
    Dim pLUReClasstable As iTable
    Set pLUReClasstable = GetInputDataTable("LUReclass")
    Dim strLucode As String
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMLandusesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMLandusesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMLandusesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    
   
    '** write header information for Rain Gauges
    m_pFile.WriteLine "[LANDUSES]"
    m_pFile.WriteLine ";;"
    'm_pFile.WriteLine ";;" & "Name" & vbTab & "Cleaning Interval" & vbTab & "Fraction Available" & _
                        vbTab & "Last Cleaned"
    'm_pFile.WriteLine ";;"
    
    Dim strFields As String
    Dim strValues As String
    Dim pFLag As Boolean
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        strValues = ""
        strFields = ";;"
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount) & " And PropName NOT LIKE '%-%'"
        Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            strFields = strFields & vbTab & pRow.value(iPropName)
            strValues = strValues & vbTab & pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
        
        If Not pFLag Then
            m_pFile.WriteLine strFields & vbTab & "Sand    Silt    Clay"
            m_pFile.WriteLine ";;"
            pFLag = True
        End If
        
        ' Write the Sand/Silt/Clay props....
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount) & " And PropName = 'Name'"
        Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        strLucode = pRow.value(iPropValue)
        
        If Right(strLucode, 4) = "_imp" Then
            strLucode = Left(strLucode, Len(strLucode) - 4)
        ElseIf Right(strLucode, 5) = "_perv" Then
            strLucode = Left(strLucode, Len(strLucode) - 5)
        End If
        pQueryFilter.WhereClause = "LUGroup= '" & strLucode & "'"
        Set pCursor = pLUReClasstable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        strLucode = pRow.value(pLUReClasstable.FindField("SandFrac")) & vbTab & pRow.value(pLUReClasstable.FindField("SiltFrac")) & vbTab & pRow.value(pLUReClasstable.FindField("ClayFrac"))
        
        m_pFile.WriteLine strValues & vbTab & strLucode
        
    Next iCount
    
'    Dim pPropertyDict As Scripting.Dictionary
'    Dim pLanduseString As String
'    '** iterate over the table to write OPTIONS value
'    Dim iCount As Integer
'    For iCount = 1 To pIDCollection.Count
'        '** for each ID, add values to a dictionary
'        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
'        Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
'        Set pRow = pCursor.NextRow
'        Set pPropertyDict = CreateObject("Scripting.Dictionary")
'        Do While Not pRow Is Nothing
'            pPropertyDict.Add pRow.value(iPropName), pRow.value(iPropValue)
'            Set pRow = pCursor.NextRow
'        Loop
'
'        '** write values from dictionary to string variable
'        pLanduseString = pPropertyDict.Item("Name") & vbTab & _
'                          pPropertyDict.Item("Interval") & vbTab & _
'                          pPropertyDict.Item("Availibility") & vbTab & _
'                          pPropertyDict.Item("LastSwept")
'
'        '** write string variable to the input file.
'        m_pFile.WriteLine pLanduseString
'
'    Next
   
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    'Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMLandusesTable = Nothing

End Sub


'** Write landuse coverages and loadings for sub catchment
Private Sub WriteSWMMCoveragesAndLoadingsInfo()

    '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
     
    '** write header information for COVERAGES
    Dim pCoveragesHeader As String
    Dim pCoveragesBody As String
    pCoveragesHeader = "[COVERAGES]" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & "Subcatchment" & vbTab & "Landuse" & vbTab & "Percent" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & vbNewLine
    
    
    '** write header information for LOADINGS
    Dim pLoadingsHeader As String
    Dim pLoadingsBody As String
    pLoadingsHeader = "[LOADINGS]" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & "Subcatchment" & vbTab & "Pollutant" & vbTab & "Loading" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & vbNewLine
       
       
    '** iterate over the table to get coverage values
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName LIKE 'Landuse%'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pLanduseName As String
    Do While Not pRow Is Nothing
        pLanduseName = Trim(Replace(pRow.value(iPropName), "Landuse: ", ""))
        '** write values for COVERAGES
        pCoveragesBody = pCoveragesBody & pRow.value(iID) & vbTab & _
                         pLanduseName & vbTab & pRow.value(iPropValue) & vbNewLine
        Set pRow = pCursor.NextRow
    Loop
    
    '** cleanup to use again
    Set pCursor = Nothing
    Set pRow = Nothing
                          
    '** iterate over the table to get loadings values
    pQueryFilter.WhereClause = "PropName LIKE 'Pollutant%'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pPollutantName As String
    Do While Not pRow Is Nothing
        pPollutantName = Trim(Replace(pRow.value(iPropName), "Pollutant: ", ""))
        '** write values for LOADINGS
        pLoadingsBody = pLoadingsBody & pRow.value(iID) & vbTab & _
                         pPollutantName & vbTab & pRow.value(iPropValue) & vbNewLine
        Set pRow = pCursor.NextRow
    Loop
    
    '** write COVERAGES values to the input file.
    m_pFile.Write pCoveragesHeader
    m_pFile.Write pCoveragesBody
    m_pFile.WriteLine ""
    
    '** write LOADINGS values to the input file.
    m_pFile.Write pLoadingsHeader
    m_pFile.Write pLoadingsBody
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentsTable = Nothing
End Sub


'** Write BUILDUP, WASHOFF information for subcatchment
Private Sub WriteSWMMBuildupWashoffInfo()

    '** get landuses info table
    Dim pSWMMLandusesTable As iTable
    Set pSWMMLandusesTable = GetInputDataTable("LANDLanduses")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMLandusesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMLandusesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMLandusesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop

    '** write header information for BUILDUP
    Dim pBuildupHeader As String
    Dim pBuildupBody As String
    pBuildupHeader = "[BUILDUP]" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & "Landuse" & vbTab & "Pollutant" & vbTab & "Function" & _
                                                vbTab & "Coeff1" & vbTab & "Coeff2" & vbTab & _
                                                "Coeff3" & vbTab & "Normalizer" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & vbNewLine
    
    
    '** write header information for WASHOFF
    Dim pWashoffHeader As String
    Dim pWashoffBody As String
    pWashoffHeader = "[WASHOFF]" & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & "Landuse" & vbTab & "Pollutant" & vbTab & "Function" & _
                                                vbTab & "Coeff1" & vbTab & "Coeff2" & vbTab & _
                                                "Clean. Effic." & vbTab & "BMP Effic." & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & vbNewLine
    
    
    '** get the pollutant collection
    Dim pPollutantColl As Collection
    Set pPollutantColl = ModuleSWMMFunctions.LoadPollutantNames
    If (pPollutantColl Is Nothing) Then
        Exit Sub
    End If
    Dim pCount As Integer
    Dim pPollutant As String
    
    Dim pPropertyDict As Scripting.Dictionary
    Dim pRainGageString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        pPropertyDict.add "ID", pIDCollection.Item(iCount)
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
    
        '** iterate over each pollutant : BUILDUP values
        For pCount = 1 To pPollutantColl.Count
            '** get pollutant name
            pPollutant = pPollutantColl.Item(pCount)
            '** get BUILDUP values for this pollutant and write to the string variable
            pBuildupBody = pBuildupBody & _
                              pPropertyDict.Item("Name") & vbTab & _
                              pPollutant & vbTab & _
                              pPropertyDict.Item(pPollutant & "-BUILDUPFUNCTION") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-MAXBUILDUP") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-RATECONSTANT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-POWER") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-NORMALIZER") & vbNewLine
        
        Next
        
        '** iterate over each pollutant : WASHOFF values
        For pCount = 1 To pPollutantColl.Count
            '** get pollutant name
            pPollutant = pPollutantColl.Item(pCount)
            '** get WASHOFF values for this pollutant and write to the string variable
            pWashoffBody = pWashoffBody & _
                              pPropertyDict.Item("Name") & vbTab & _
                              pPollutant & vbTab & _
                              pPropertyDict.Item(pPollutant & "-WASHOFFFUNCTION") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-COEFFICIENT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-EXPONENT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-CLEANINGEFF") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-BMPEFF") & vbNewLine
        
        Next
                                
    Next
    
    '** write BUILDUP values to the input file.
    m_pFile.Write pBuildupHeader
    m_pFile.Write pBuildupBody
    m_pFile.WriteLine ""
    
    '** write WASHOFF values to the input file.
    m_pFile.Write pWashoffHeader
    m_pFile.Write pWashoffBody

    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pPollutantColl = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMLandusesTable = Nothing

End Sub




'** Write OPTIONS information
Private Sub WriteSWMMAdditionalOptionsInfo()

    '** get landuses info table
    Dim pSWMMOptionsTable As iTable
    Set pSWMMOptionsTable = GetInputDataTable("LANDLanduses")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMOptionsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMOptionsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMOptionsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'REPORT CONTROL'"
    Set pCursor = pSWMMOptionsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pReportControl As String
    If Not pRow Is Nothing Then
        pReportControl = pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    End If

    '** write information for REPORT CONTROL
    m_pFile.WriteLine "[REPORT]"
    m_pFile.WriteLine "CONTROLS" & vbTab & pReportControl
    m_pFile.WriteLine ""
    
    '** write information for TEMPORARY DIRECTORY
    m_pFile.WriteLine "[OPTIONS]"
    m_pFile.WriteLine "TEMPDIR" & vbTab & gMapTempFolder
              
    '** cleanup
    Set pQueryFilter = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMOptionsTable = Nothing

End Sub


Private Sub WriteSWMMPreDevelopedTitleOptionsAndFiles(pPreDevLandUseFileName As String)

    Dim pSWMMOptionsTable As iTable
    Set pSWMMOptionsTable = GetInputDataTable("LANDOptions")
    Dim pCursor As ICursor
    Set pCursor = pSWMMOptionsTable.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim iPropName As Long
    iPropName = pSWMMOptionsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMOptionsTable.FindField("PropValue")
    Dim pOutflowFile As String
        
    '** iterate over the table to write OPTIONS value
    m_pFile.WriteLine "[OPTIONS]"
    Do While Not (pRow Is Nothing)
        If (pRow.value(iPropName) = "SAVE OUTFLOWS") Then
            pOutflowFile = pRow.value(iPropValue)
'        ElseIf (pRow.value(iPropName) = "SAVE POST OUTFLOWS") Then
'            pOutflowFile = pRow.value(iPropValue)
        ElseIf (pRow.value(iPropName) <> "REPORT CONTROL" And pRow.value(iPropName) <> "SAVE POST OUTFLOWS") Then
            m_pFile.WriteLine pRow.value(iPropName) & vbTab & pRow.value(iPropValue)
        End If
        Set pRow = pCursor.NextRow
    Loop
    m_pFile.WriteLine ""
        
    '** write the outflow file name
    m_pFile.WriteLine "[FILES]"
    m_pFile.WriteLine "SAVE OUTFLOWS" & "   " & pOutflowFile
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMOptionsTable = Nothing

End Sub

'** Write subcatchment information for predeveloped
Private Sub WriteSWMMPreDevelopedSubCatchmentInfo(pConductivity As Double, pSuctionHead As Double, pInitialDeficit As Double)

    '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    Set pIDCollection = SortCollection(pIDCollection, True)
    
    Dim pOutflowFile As String
    
    '** write header information for SUBCATCHMENTS
    Dim pSubCatchmentHeader As String
    Dim pSubCatchmentBody As String
    pSubCatchmentHeader = "[SUBCATCHMENTS]" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & "Name" & vbTab & "Raingage" & vbTab & "Outlet" & _
                                                vbTab & "Total Area" & vbTab & "Pcnt. Imperv" & vbTab & _
                                                "Width" & vbTab & "Pcnt. Slope" & vbTab & "Curb Length" & _
                                                vbTab & "Snow Pack" & vbNewLine
    pSubCatchmentHeader = pSubCatchmentHeader & ";;" & vbNewLine
    
    
    '** write header information for SUBAREAS
    Dim pSubAreaHeader As String
    Dim pSubAreaBody As String
    pSubAreaHeader = "[SUBAREAS]" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & "Subcatchment" & vbTab & "N-Imperv" & vbTab & "N-Perv" & _
                                            vbTab & "S-Imperv" & vbTab & "S-Perv" & vbTab & _
                                            "PctZero" & vbTab & "RouteTo" & vbTab & "PctRouted" & vbNewLine
    pSubAreaHeader = pSubAreaHeader & ";;" & vbNewLine
    
    
    '** write header information for INFILTRATION
    Dim pInfiltrationHeader As String
    Dim pInfiltrationBody As String
    pInfiltrationHeader = "[INFILTRATION]" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & "Subcatchment" & vbTab & "Suction" & vbTab & "HydCon" & _
                                            vbTab & "IMDmax" & vbNewLine
    pInfiltrationHeader = pInfiltrationHeader & ";;" & vbNewLine
      
    Dim pPropertyDict As Scripting.Dictionary
    Dim pRainGageString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        pPropertyDict.add "ID", pIDCollection.Item(iCount)
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
    
        '** write values for SUBCATCHMENTS - The % impervious value for predeveloped condition is 0
        pSubCatchmentBody = pSubCatchmentBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("Rain Gauge") & vbTab & _
                          pPropertyDict.Item("Outlet") & vbTab & _
                          pPropertyDict.Item("Area") & vbTab & _
                          0 & vbTab & _
                          pPropertyDict.Item("Width") & vbTab & _
                          pPropertyDict.Item("%Slope") & vbTab & _
                          pPropertyDict.Item("Curb Length") & vbTab & _
                          pPropertyDict.Item("Snow Packs") & vbNewLine
                          
                          
        '** write values for SUBAREAS
        pSubAreaBody = pSubAreaBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("NImpervious") & vbTab & _
                          pPropertyDict.Item("NPervious") & vbTab & _
                          pPropertyDict.Item("DImpervious") & vbTab & _
                          pPropertyDict.Item("DPervious") & vbTab & _
                          pPropertyDict.Item("%ZeroImpervious") & vbTab & _
                          pPropertyDict.Item("SubareaRouting") & vbTab & _
                          pPropertyDict.Item("%Routing") & vbNewLine
                          
                          
        '** write values for INFILTRATION - FOR PREDEVELOPED CONDITIONS
        pInfiltrationBody = pInfiltrationBody & _
                          pPropertyDict.Item("ID") & vbTab & _
                          pPropertyDict.Item("Suction Head") & vbTab & _
                          pPropertyDict.Item("Conductivity") & vbTab & _
                          pPropertyDict.Item("Initial Deficit") & vbNewLine
                                  
    Next
    
    '** write SUBCATCHMENT values to the input file.
    m_pFile.Write pSubCatchmentHeader
    m_pFile.Write pSubCatchmentBody
    m_pFile.WriteLine ""
    
    '** write SUBAREAS values to the input file.
    m_pFile.Write pSubAreaHeader
    m_pFile.Write pSubAreaBody
    m_pFile.WriteLine ""
    
    '** write INFILTRATION values to the input file.
    m_pFile.Write pInfiltrationHeader
    m_pFile.Write pInfiltrationBody
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentsTable = Nothing
    pSubCatchmentHeader = ""
    pSubCatchmentBody = ""
    pSubAreaHeader = ""
    pSubAreaBody = ""
    pInfiltrationHeader = ""
    pInfiltrationBody = ""
    
End Sub


'** Write landuse information
Private Sub WriteSWMMPreDevelopedLanduseInfo(pPreDevLanduse As String)

    '** get pollutant info table
    Dim pSWMMLandusesTable As iTable
    Set pSWMMLandusesTable = GetInputDataTable("LANDLanduses")
    
    Dim pLUReClasstable As iTable
    Set pLUReClasstable = GetInputDataTable("LUReclass")
    Dim strLucode As String
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMLandusesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMLandusesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMLandusesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name' And PropValue = '" & pPreDevLanduse & "'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pPreDevIDValue As Integer
    If Not pRow Is Nothing Then
        pPreDevIDValue = pRow.value(iID)
    End If
    
   '** write header information for Rain Gauges
    m_pFile.WriteLine "[LANDUSES]"
    m_pFile.WriteLine ";;"
    'm_pFile.WriteLine ";;" & "Name" & vbTab & "Cleaning Interval" & vbTab & "Fraction Available" & _
                        vbTab & "Last Cleaned"
    'm_pFile.WriteLine ";;"
    
    Dim strFields As String
    Dim strValues As String
    Dim pFLag As Boolean
    Dim iCount As Integer
    strValues = ""
    strFields = ";;"
    '** for each ID, add values to a dictionary
    pQueryFilter.WhereClause = "ID = " & pPreDevIDValue & " And PropName NOT LIKE '%-%'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        strFields = strFields & vbTab & pRow.value(iPropName)
        strValues = strValues & vbTab & pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    Loop
    
    If Not pFLag Then
        m_pFile.WriteLine strFields & vbTab & "Sand    Silt    Clay"
        m_pFile.WriteLine ";;"
        pFLag = True
    End If
    
    ' Write the Sand/Silt/Clay props....
    pQueryFilter.WhereClause = "ID = " & pPreDevIDValue & " And PropName = 'Name'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    strLucode = pRow.value(iPropValue)
    
    If Right(strLucode, 4) = "_imp" Then
        strLucode = Left(strLucode, Len(strLucode) - 4)
    ElseIf Right(strLucode, 5) = "_perv" Then
        strLucode = Left(strLucode, Len(strLucode) - 5)
    End If
        
    pQueryFilter.WhereClause = "LUGroup= '" & strLucode & "'"
    Set pCursor = pLUReClasstable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    If Not pRow Is Nothing Then strLucode = pRow.value(pLUReClasstable.FindField("SandFrac")) & vbTab & pRow.value(pLUReClasstable.FindField("SiltFrac")) & vbTab & pRow.value(pLUReClasstable.FindField("ClayFrac"))

    m_pFile.WriteLine strValues & vbTab & strLucode
        
'    '** write header information for Rain Gauges
'    m_pFile.WriteLine "[LANDUSES]"
'    m_pFile.WriteLine ";;"
'    m_pFile.WriteLine ";;" & "Name" & vbTab & "Cleaning Interval" & vbTab & "Fraction Available" & _
'                        vbTab & "Last Cleaned"
'    m_pFile.WriteLine ";;"
'
'    Dim pPropertyDict As Scripting.Dictionary
'    Dim pLanduseString As String
'    '** for each ID, add values to a dictionary
'    pQueryFilter.WhereClause = "ID = " & pPreDevIDValue
'    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
'    Set pRow = pCursor.NextRow
'    Set pPropertyDict = CreateObject("Scripting.Dictionary")
'    Do While Not pRow Is Nothing
'        pPropertyDict.Add pRow.value(iPropName), pRow.value(iPropValue)
'        Set pRow = pCursor.NextRow
'    Loop
'
'    '** write values from dictionary to string variable
'    pLanduseString = pPropertyDict.Item("Name") & vbTab & _
'                      pPropertyDict.Item("Interval") & vbTab & _
'                      pPropertyDict.Item("Availibility") & vbTab & _
'                      pPropertyDict.Item("LastSwept")
'
'    '** write string variable to the input file.
'    m_pFile.WriteLine pLanduseString
    
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    'Set pPropertyDict = Nothing
    'Set pIDCollection = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMLandusesTable = Nothing

End Sub


'** Write landuse coverages and loadings for sub catchment
Private Sub WriteSWMMPreDevelopedCoveragesAndLoadingsInfo(pPreDevLanduse As String)

    '** get sub catchment info table
    Dim pSWMMSubCatchmentsTable As iTable
    Set pSWMMSubCatchmentsTable = GetInputDataTable("LANDSubCatchments")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMSubCatchmentsTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentsTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentsTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
     
    '** write header information for COVERAGES
    Dim pCoveragesHeader As String
    Dim pCoveragesBody As String
    pCoveragesHeader = "[COVERAGES]" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & "Subcatchment" & vbTab & "Landuse" & vbTab & "Percent" & vbNewLine
    pCoveragesHeader = pCoveragesHeader & ";;" & vbNewLine
        
    '** write header information for LOADINGS
    Dim pLoadingsHeader As String
    Dim pLoadingsBody As String
    pLoadingsHeader = "[LOADINGS]" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & "Subcatchment" & vbTab & "Pollutant" & vbTab & "Loading" & vbNewLine
    pLoadingsHeader = pLoadingsHeader & ";;" & vbNewLine
              
    '** iterate over the table to get coverage values
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName LIKE 'Name'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pLanduseName As String
    Do While Not pRow Is Nothing
        pLanduseName = pPreDevLanduse
        '** write values for COVERAGES  - For all watersheds, contributing
        'landuse is predeveloped with 100%
        pCoveragesBody = pCoveragesBody & pRow.value(iID) & vbTab & _
                         pLanduseName & vbTab & 100 & vbNewLine
        Set pRow = pCursor.NextRow
    Loop
    
    '** cleanup to use again
    Set pCursor = Nothing
    Set pRow = Nothing
                          
    '** iterate over the table to get loadings values
    pQueryFilter.WhereClause = "PropName LIKE 'Pollutant%'"
    Set pCursor = pSWMMSubCatchmentsTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pPollutantName As String
    Do While Not pRow Is Nothing
        pPollutantName = Trim(Replace(pRow.value(iPropName), "Pollutant: ", ""))
        '** write values for LOADINGS
        pLoadingsBody = pLoadingsBody & pRow.value(iID) & vbTab & _
                         pPollutantName & vbTab & pRow.value(iPropValue) & vbNewLine
        Set pRow = pCursor.NextRow
    Loop
    
    '** write COVERAGES values to the input file.
    m_pFile.Write pCoveragesHeader
    m_pFile.Write pCoveragesBody
    m_pFile.WriteLine ""
    
    '** write LOADINGS values to the input file.
    m_pFile.Write pLoadingsHeader
    m_pFile.Write pLoadingsBody
    m_pFile.WriteLine ""
    
    '** cleanup
    Set pQueryFilter = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentsTable = Nothing
End Sub


'** Write BUILDUP, WASHOFF information for subcatchment
'** Updated for predeveloped landuse condition
Private Sub WriteSWMMPreDevelopedBuildupWashoffInfo(pPreDevLanduse As String)

    '** get landuses info table
    Dim pSWMMLandusesTable As iTable
    Set pSWMMLandusesTable = GetInputDataTable("LANDLanduses")
    
    '** define field indexes
    Dim iID As Long
    iID = pSWMMLandusesTable.FindField("ID")
    Dim iPropName As Long
    iPropName = pSWMMLandusesTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMLandusesTable.FindField("PropValue")
    
    '** define variables to access the table
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    '** query the table to get all unique ids, put them in a collection
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name' And PropValue = '" & pPreDevLanduse & "'"
    Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDCollection As Collection
    Set pIDCollection = New Collection
    Do While Not pRow Is Nothing
        pIDCollection.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop

    '** write header information for BUILDUP
    Dim pBuildupHeader As String
    Dim pBuildupBody As String
    pBuildupHeader = "[BUILDUP]" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & "Landuse" & vbTab & "Pollutant" & vbTab & "Function" & _
                                                vbTab & "Coeff1" & vbTab & "Coeff2" & vbTab & _
                                                "Coeff3" & vbTab & "Normalizer" & vbNewLine
    pBuildupHeader = pBuildupHeader & ";;" & vbNewLine
    
    
    '** write header information for WASHOFF
    Dim pWashoffHeader As String
    Dim pWashoffBody As String
    pWashoffHeader = "[WASHOFF]" & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & "Landuse" & vbTab & "Pollutant" & vbTab & "Function" & _
                                                vbTab & "Coeff1" & vbTab & "Coeff2" & vbTab & _
                                                "Clean. Effic." & vbTab & "BMP Effic." & vbNewLine
    pWashoffHeader = pWashoffHeader & ";;" & vbNewLine
    
    
    '** get the pollutant collection
    Dim pPollutantColl As Collection
    Set pPollutantColl = ModuleSWMMFunctions.LoadPollutantNames
    If (pPollutantColl Is Nothing) Then
        Exit Sub
    End If
    Dim pCount As Integer
    Dim pPollutant As String
    
    Dim pPropertyDict As Scripting.Dictionary
    Dim pRainGageString As String
    '** iterate over the table to write OPTIONS value
    Dim iCount As Integer
    For iCount = 1 To pIDCollection.Count
        '** for each ID, add values to a dictionary
        pQueryFilter.WhereClause = "ID = " & pIDCollection.Item(iCount)
        Set pCursor = pSWMMLandusesTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
        pPropertyDict.add "ID", pIDCollection.Item(iCount)
        Do While Not pRow Is Nothing
            pPropertyDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow
        Loop
    
        '** iterate over each pollutant : BUILDUP values
        For pCount = 1 To pPollutantColl.Count
            '** get pollutant name
            pPollutant = pPollutantColl.Item(pCount)
            '** get BUILDUP values for this pollutant and write to the string variable
            pBuildupBody = pBuildupBody & _
                              pPropertyDict.Item("Name") & vbTab & _
                              pPollutant & vbTab & _
                              pPropertyDict.Item(pPollutant & "-BUILDUPFUNCTION") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-MAXBUILDUP") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-RATECONSTANT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-POWER") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-NORMALIZER") & vbNewLine
        
        Next
        
        '** iterate over each pollutant : WASHOFF values
        For pCount = 1 To pPollutantColl.Count
            '** get pollutant name
            pPollutant = pPollutantColl.Item(pCount)
            '** get WASHOFF values for this pollutant and write to the string variable
            pWashoffBody = pWashoffBody & _
                              pPropertyDict.Item("Name") & vbTab & _
                              pPollutant & vbTab & _
                              pPropertyDict.Item(pPollutant & "-WASHOFFFUNCTION") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-COEFFICIENT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-EXPONENT") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-CLEANINGEFF") & vbTab & _
                              pPropertyDict.Item(pPollutant & "-BMPEFF") & vbNewLine
        
        Next
                                
    Next
    
    '** write BUILDUP values to the input file.
    m_pFile.Write pBuildupHeader
    m_pFile.Write pBuildupBody
    m_pFile.WriteLine ""
    
    '** write WASHOFF values to the input file.
    m_pFile.Write pWashoffHeader
    m_pFile.Write pWashoffBody
    m_pFile.WriteLine ""
       
    '** cleanup
    Set pQueryFilter = Nothing
    Set pPropertyDict = Nothing
    Set pIDCollection = Nothing
    Set pPollutantColl = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMLandusesTable = Nothing

End Sub

'Which was pretty simple. It sorts a collection by a property.
'Ascending of Decending

Private Function SortCollection(oCol As Collection, pAscend As Boolean) As Collection

    On Error GoTo EH
    Dim z As Integer, X As Integer, Y As Integer
    Dim temp As Variant
    Dim oNewCol As Collection
    Set oNewCol = New Collection
    
    ReDim A(0 To oCol.Count - 1) As Integer
    X = 0
    Dim Val As Variant
    For Each Val In oCol
        A(X) = Val
        X = X + 1
    Next
    
    z = oCol.Count - 1
    For X = 1 To z
      For Y = 0 To z
       If Y <> z Then
        If pAscend Then
         If A(Y + 1) < A(Y) Then
          temp = A(Y)
          A(Y) = A(Y + 1)
          A(Y + 1) = temp
         End If
        Else
         If A(Y + 1) > A(Y) Then
          temp = A(Y)
          A(Y) = A(Y + 1)
          A(Y + 1) = temp
         End If
        End If
       End If
      Next Y
    Next X
    
    For z = 0 To UBound(A)
        oNewCol.add A(z)
    Next z
    
    Set SortCollection = oNewCol
         
     
  Exit Function
EH:

  MsgBox Err.Number & " " & Err.description & " in SortCollection"

End Function

Public Function LoadSWMMClimatologyDataToDictionary() As Scripting.Dictionary
   
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDClimatology")
    If (pTable Is Nothing) Then
        Set LoadSWMMClimatologyDataToDictionary = Nothing
        Exit Function
    End If
    
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iPropName As Long
    iPropName = pTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pOptionDictionary As Scripting.Dictionary
    Set pOptionDictionary = CreateObject("Scripting.Dictionary")
    Do While Not (pRow Is Nothing)
        pOptionDictionary.Item(pRow.value(iPropName)) = pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    Loop
    
    '** return the option dictionary back
    Set LoadSWMMClimatologyDataToDictionary = pOptionDictionary
    
    GoTo CleanUp

    
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
    
End Function
