Attribute VB_Name = "ModuleVFSFunctions"
Option Explicit

Public Enum SNAP_MODE
  SNAP_NEAREST_POINT = 1
  SNAP_NEAREST_NODE = 2
  SNAP_NEAREST_JUNCTION = 3
End Enum

Public Enum TRACE_MODE
  TRACE_DOWN = 1
  TRACE_JUNCTION = 2
End Enum
'*** Get properties from Vegetative Filter Strip Property table
Public Function GetVFSProperties(pTableName As String, pID As String) As Scripting.Dictionary
On Error GoTo ShowError
    'Find the VFS option table
    Dim pVFSPropertyTable As iTable
    Set pVFSPropertyTable = GetInputDataTable(pTableName)
    
    'Create the table if not found, add it to the Map
    If (pVFSPropertyTable Is Nothing) Then
        Set GetVFSProperties = Nothing
        Exit Function
    End If
    Dim iPropName As Long
    Dim iPropValue As Long
    
    iPropName = pVFSPropertyTable.FindField("PropName")
    iPropValue = pVFSPropertyTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pPropertyDictionary As Scripting.Dictionary
    Set pPropertyDictionary = New Scripting.Dictionary
    
    pQueryFilter.WhereClause = "ID = " & pID
    Set pCursor = pVFSPropertyTable.Search(pQueryFilter, False)
    'Iterate over the selected rows
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
        pPropertyName = pRow.value(iPropName)
        pPropertyValue = pRow.value(iPropValue)
        pPropertyDictionary.Item(pPropertyName) = pPropertyValue
        Set pRow = pCursor.NextRow
    Loop
    
    Set GetVFSProperties = pPropertyDictionary
    
    GoTo CleanUp
ShowError:
    MsgBox "Error in GetVFSProperties: " & Err.description
CleanUp:
    Set pQueryFilter = Nothing
    Set pVFSPropertyTable = Nothing
    Set pRow = Nothing
    Set pPropertyDictionary = Nothing
End Function
Public Function GetDefaultsForVFS(vfsId As Integer, vfsName As String) As Scripting.Dictionary
On Error GoTo ShowError

    Dim pPropDictionary As Scripting.Dictionary
    Set pPropDictionary = New Scripting.Dictionary
    
    pPropDictionary.Item("Name") = vfsName
    pPropDictionary.Item("ID") = vfsId

''    Call CreatePollutantList
''    Dim pTotalPollutants As Integer
''    pTotalPollutants = UBound(gPollutants) + 1
    
    'Check to see whether DBF file exists
''    Dim pMultiplierTable As iTable
''    Set pMultiplierTable = GetInputDataTable("TSMultipliers")
''
''    Dim sedPollInd As Integer
''
''    Dim pRow As iRow
''    Dim pCursor As esriGeoDatabase.ICursor
''    Dim iR As Integer
''
''    If pMultiplierTable Is Nothing Then
''        MsgBox "TSMultipliers table is missing. Can not identify sediment"
''        Exit Function
''    Else
''
''        Set pCursor = pMultiplierTable.Search(Nothing, False)
''        Dim pSedFlagInd As Integer
''        pSedFlagInd = pMultiplierTable.FindField("SedFlag")
''
''        Set pRow = pCursor.NextRow
''        iR = 1
''        Do While Not pRow Is Nothing
''            If pRow.value(pSedFlagInd) = 1 Then
''                sedPollInd = iR 'Set the pollutant index for sediment
''                Exit Do
''            End If
''            Set pRow = pCursor.NextRow
''            iR = iR + 1
''        Loop
''    End If
    
''    For iR = 1 To pTotalPollutants
''        If iR <> sedPollInd Then
''            pPropDictionary.Item("SedFrac" & iR) = 0.05
''            pPropDictionary.Item("SedDec" & iR) = 1.05
''            pPropDictionary.Item("SedCorr" & iR) = 0.05
''            pPropDictionary.Item("WatDec" & iR) = 1.05
''            pPropDictionary.Item("WatCorr" & iR) = 0.05
''        End If
''    Next
        
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        MsgBox "Missing pollutants table: Define pollutants first"
        Exit Function
    End If
    
    Dim iFlagFld As Integer
    iFlagFld = pTable.FindField("Sediment")
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pCount As Integer
    pCount = pTable.RowCount(Nothing)
    
    ReDim Preserve gPollutants(pCount - 1) As String
    
    Dim iR As Integer
    For iR = 1 To pCount
        pQueryFilter.WhereClause = " ID = " & iR
        Set pCursor = pTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        If UCase(pRow.value(iFlagFld)) = "NO" Then
            pPropDictionary.Item("SedFrac" & iR) = 0.05
            pPropDictionary.Item("SedDec" & iR) = 1.05
            pPropDictionary.Item("SedCorr" & iR) = 0.05
            pPropDictionary.Item("WatDec" & iR) = 1.05
            pPropDictionary.Item("WatCorr" & iR) = 0.05
        End If
    Next
    'pPropDictionary.Item("N") = 51
    'pPropDictionary.Item("THETAW") = 0.5
    'pPropDictionary.Item("CR") = 0.6
    'pPropDictionary.Item("MAXITER") = 100
    'pPropDictionary.Item("NPOL") = 3
    'pPropDictionary.Item("KPG") = 0
    pPropDictionary.Item("VKS") = 0.02 '0.5 * 3.3 'in/hr
    pPropDictionary.Item("Sav") = 1#   '0.5 * 3.3 'ft
    pPropDictionary.Item("OI") = 0.31
    pPropDictionary.Item("OS") = 0.125
    pPropDictionary.Item("SM") = 0.2 '0.5 * 3.3
    pPropDictionary.Item("SCHK") = 1#
    pPropDictionary.Item("SS") = 0.2 '0.5 / 2.54
    pPropDictionary.Item("H") = 6 ' in
    pPropDictionary.Item("VN") = 0.012
    pPropDictionary.Item("Vn2") = 0.04
    
    pPropDictionary.Item("PORSand") = 0.5
    pPropDictionary.Item("NPARTSand") = 5
    pPropDictionary.Item("COARSESand") = 0.9
    pPropDictionary.Item("DPSand") = 0.0078 '0.02 / 2.54
    pPropDictionary.Item("SGSand") = 165#  '2.65 * 62.4
    pPropDictionary.Item("PORSilt") = 0.5
    pPropDictionary.Item("NPARTSilt") = 6
    pPropDictionary.Item("COARSESilt") = 0.5
    pPropDictionary.Item("DPSilt") = 0.001 ' 0.003 / 2.54
    pPropDictionary.Item("SGSilt") = 165 '2.65 * 62.4
    pPropDictionary.Item("PORClay") = 0.5
    pPropDictionary.Item("NPARTClay") = 1
    pPropDictionary.Item("COARSEClay") = 0.001
    pPropDictionary.Item("DPClay") = 0.00001 '0.0002 / 2.54
    pPropDictionary.Item("SGClay") = 165 ' 2.65 * 62.4
    
    pPropDictionary.Item("NPROP") = 1
    Dim i As Integer
    For i = 1 To pPropDictionary.Item("NPROP")
        pPropDictionary.Item("SX" & i) = 1
        pPropDictionary.Item("RNA" & i) = 0.4
        pPropDictionary.Item("SOA" & i) = 0.05
    Next
    
            
    Set GetDefaultsForVFS = pPropDictionary
    Exit Function
ShowError:
    MsgBox "Error in GetDefaultsForVFS: " & Err.description
End Function
Public Function InitializeVFSPropertyForm(pPropertyDictionary As Scripting.Dictionary) As Boolean
    'Get the number of pollutants
    Call CreatePollutantList
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants) + 1
    
    'Check to see whether DBF file exists
''    Dim pMultiplierTable As iTable
''    Set pMultiplierTable = GetInputDataTable("TSMultipliers")
''
''    Dim sedPollInd As Integer
''
''    Dim pRow As iRow
''    Dim pCursor As esriGeoDatabase.ICursor
''    Dim iR As Integer
''
''    If pMultiplierTable Is Nothing Then
''        MsgBox "TSMultipliers table is missing. Can not identify sediment"
''        Exit Function
''    Else
''
''        Set pCursor = pMultiplierTable.Search(Nothing, False)
''        Dim pSedFlagInd As Integer
''        pSedFlagInd = pMultiplierTable.FindField("SedFlag")
''
''        Set pRow = pCursor.NextRow
''        iR = 1
''        Do While Not pRow Is Nothing
''            If pRow.value(pSedFlagInd) = 1 Then
''                sedPollInd = iR 'Set the pollutant index for sediment
''                Exit Do
''            End If
''            Set pRow = pCursor.NextRow
''            iR = iR + 1
''        Loop
''    End If
    
    'If DBF file exists and the number of records match with that of the pollutants
    'then initialize the DB grid with the values
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "ID", adVarChar, 50
    oRs.Fields.Append "Pollutant", adVarChar, 50
    oRs.Fields.Append "SedFrac", adDouble
    oRs.Fields.Append "Decay1", adDouble 'Adsorbed fraction decay
    oRs.Fields.Append "TempCorr1", adDouble 'Adsorbed fraction
    oRs.Fields.Append "Decay2", adDouble 'Dissolved fraction decay
    oRs.Fields.Append "TempCorr2", adDouble 'Adsorbed fraction
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    'Create a record set for grass segments
    Dim oRsSeg As ADODB.Recordset
    Set oRsSeg = New ADODB.Recordset
    oRsSeg.Fields.Append "Distance", adDouble
    oRsSeg.Fields.Append "Roughness", adDouble
    oRsSeg.Fields.Append "Slope", adDouble
    oRsSeg.CursorType = adOpenDynamic
    oRsSeg.Open
    
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        MsgBox "Missing pollutants table: Define pollutants first"
        Exit Function
    End If
    
    Dim iFlagFld As Integer
    iFlagFld = pTable.FindField("Sediment")
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pCount As Integer
    pCount = pTable.RowCount(Nothing)
    
    Dim pSedFlags() As String
    ReDim Preserve pSedFlags(pCount - 1) As String
    
    Dim iR As Integer
    For iR = 1 To pCount
        pQueryFilter.WhereClause = " ID = " & iR
        Set pCursor = pTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        If Not pRow Is Nothing Then pSedFlags(iR - 1) = UCase(pRow.value(iFlagFld))
    Next
    
     '* Set datagrid value, header caption and width
    With FrmVFSParams
        Set .DataGridSedDec.DataSource = oRs
        .DataGridSedDec.ColumnHeaders = True
        .DataGridSedDec.Columns(0).Caption = "ID"
        .DataGridSedDec.Columns(0).Locked = True
        .DataGridSedDec.Columns(0).Visible = False
        .DataGridSedDec.Columns(1).Caption = "Pollutant"
        .DataGridSedDec.Columns(1).Locked = True
        .DataGridSedDec.Columns(1).Width = 2400
        .DataGridSedDec.Columns(2).Caption = "Sediment fraction"
        .DataGridSedDec.Columns(3).Width = 1300
        .DataGridSedDec.Columns(3).Caption = "Decay Factor"
        .DataGridSedDec.Columns(3).Width = 1300
        .DataGridSedDec.Columns(4).Caption = "Temperature Correction Factor"
        .DataGridSedDec.Columns(4).Width = .DataGridSedDec.Width - 5100
        .DataGridSedDec.Columns(5).Visible = False
        .DataGridSedDec.Columns(6).Visible = False

        '* Set datagrid value, header caption and width
        Set .DataGridDissDec.DataSource = oRs
        .DataGridDissDec.ColumnHeaders = True
        .DataGridDissDec.Columns(0).Visible = False
        .DataGridDissDec.Columns(1).Caption = "Pollutant"
        .DataGridDissDec.Columns(1).Locked = True
        .DataGridDissDec.Columns(1).Width = 2400
        .DataGridDissDec.Columns(2).Visible = False
        .DataGridDissDec.Columns(3).Visible = False
        .DataGridDissDec.Columns(4).Visible = False
        .DataGridDissDec.Columns(5).Caption = "Decay Factor"
        .DataGridDissDec.Columns(5).Width = 1300
        .DataGridDissDec.Columns(6).Caption = "Temperature Correction Factor"
        .DataGridDissDec.Columns(6).Width = 2500
        
        If Not pPropertyDictionary Is Nothing Then 'Initiliaze from dictionary
            Dim pControl
            For Each pControl In .Controls
                If ((TypeOf pControl Is TextBox) And (pControl.Enabled)) Then
                   If pControl.name <> "txtName" And pControl.name <> "txtVFSID" Then
                        If pPropertyDictionary.Exists(pControl.name) Then _
                            pControl.Text = pPropertyDictionary.Item(pControl.name)
                   End If
                End If
            Next
            .txtName = pPropertyDictionary.Item("Name")
            .txtVFSID = pPropertyDictionary.Item("ID")
            
            'Set the pollutantdecay parameters
            For iR = 1 To pTotalPollutants
                'If iR <> sedPollInd Then
                If UCase(pSedFlags(iR - 1)) = "NO" Then
                    oRs.AddNew
                    oRs.Fields(0).value = iR
                    oRs.Fields(1).value = gPollutants(iR - 1)
                    If pPropertyDictionary.Exists("SedFrac" & iR) Then
                        If pPropertyDictionary.Exists("SedFrac" & iR) Then _
                            oRs.Fields(2).value = pPropertyDictionary.Item("SedFrac" & iR)
                        If pPropertyDictionary.Exists("SedDec" & iR) Then _
                            oRs.Fields(3).value = pPropertyDictionary.Item("SedDec" & iR)
                        If pPropertyDictionary.Exists("SedCorr" & iR) Then _
                            oRs.Fields(4).value = pPropertyDictionary.Item("SedCorr" & iR)
                        If pPropertyDictionary.Exists("WatDec" & iR) Then _
                            oRs.Fields(5).value = pPropertyDictionary.Item("WatDec" & iR)
                        If pPropertyDictionary.Exists("WatCorr" & iR) Then _
                            oRs.Fields(6).value = pPropertyDictionary.Item("WatCorr" & iR)
                    End If
                End If
            Next
            Set .DataGridSegments.DataSource = oRsSeg
            .DataGridSegments.ColumnHeaders = True
            .DataGridSegments.Columns(0).Caption = "X-Distance (ft)"
            .DataGridSegments.Columns(1).Width = 1500
            .DataGridSegments.Columns(1).Caption = "Mannings N"
            .DataGridSegments.Columns(1).Width = 1500
            .DataGridSegments.Columns(2).Caption = "Segment Slope"
            .DataGridSegments.Columns(2).Width = 1500
            
            Dim segCount As Integer
            segCount = 0
            If pPropertyDictionary.Exists("NPROP") Then _
                segCount = CInt(pPropertyDictionary.Item("NPROP"))
                
            If segCount > 0 Then
                For iR = 1 To segCount
                    oRsSeg.AddNew
                    If pPropertyDictionary.Exists("SX" & iR) Then
                        oRsSeg.Fields(0).value = pPropertyDictionary.Item("SX" & iR)
                    End If
                    If pPropertyDictionary.Exists("RNA" & iR) Then
                        oRsSeg.Fields(1).value = pPropertyDictionary.Item("RNA" & iR)
                    End If
                    If pPropertyDictionary.Exists("SOA" & iR) Then
                        oRsSeg.Fields(2).value = pPropertyDictionary.Item("SOA" & iR)
                    End If
                Next
            End If
            
            'populate values on cost tab
            FrmVFSParams.InitCostFromDB
            FrmVFSParams.Update_Component_List pPropertyDictionary
            
            InitializeVFSPropertyForm = True
        Else
            InitializeVFSPropertyForm = False
            Exit Function
        End If
    
    End With
   
    
    
    Exit Function
ShowError:
    MsgBox "Error loading VFSParameter Form: " & Err.description

End Function


'*** Create a Vegetative Filter Strip Property table and save the values
Public Sub SaveVFSPropertiesTable(pTableName As String, pID As String, pPropertyDictionary As Scripting.Dictionary)

On Error GoTo ShowError
    
    'Find the VFS option table
    Dim pVFSPropertyTable As iTable
    Set pVFSPropertyTable = GetInputDataTable(pTableName)
    
    'Create the table if not found, add it to the Map
    If (pVFSPropertyTable Is Nothing) Then
        Set pVFSPropertyTable = CreatePropertiesTableDBF(pTableName)
        AddTableToMap pVFSPropertyTable
    End If

    Dim iPropName As Long
    Dim iPropValue As Long
    Dim iID As Long
    
    iPropName = pVFSPropertyTable.FindField("PropName")
    iPropValue = pVFSPropertyTable.FindField("PropValue")
    iID = pVFSPropertyTable.FindField("ID")
    
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
        Set pCursor = pVFSPropertyTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pVFSPropertyTable.CreateRow
        End If
        pRow.value(iID) = pID
        pRow.value(iPropName) = pPropertyName
        pRow.value(iPropValue) = pPropertyValue
        pRow.Store
    Next
    GoTo CleanUp
    
ShowError:
    MsgBox "SaveVFSPropertiesTable: " & Err.description
CleanUp:
    Set pQueryFilter = Nothing
    Set pVFSPropertyTable = Nothing
    Set pRow = Nothing
    Set pPropertyDictionary = Nothing
End Sub

Public Function HitTestStream(ByRef pFCStream As IFeatureClass, ByVal X As Long, ByVal Y As Long)
  If pFCStream Is Nothing Then Exit Function
  
  Dim hitTestRes(2)
  Set hitTestRes(0) = Nothing
  Set hitTestRes(1) = Nothing
  hitTestRes(2) = 0
    
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFCStream

  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.FocusMap
    
  Dim pScreenDisplay As IScreenDisplay
  Set pScreenDisplay = pActiveView.ScreenDisplay
  
  Dim pDT As IDisplayTransformation
  Set pDT = pScreenDisplay.DisplayTransformation
    
  Dim pPt0 As IPoint
  Dim pPt1 As IPoint
  Dim pPt2 As IPoint

  Set pPt0 = pDT.ToMapPoint(X, Y)
  Set pPt1 = pDT.ToMapPoint(X - 4, Y - 4)
  Set pPt2 = pDT.ToMapPoint(X + 4, Y + 4)
  Set pPt0.SpatialReference = gMap.SpatialReference
  pPt0.Project pGeoDataset.SpatialReference
  Set pPt1.SpatialReference = gMap.SpatialReference
  pPt1.Project pGeoDataset.SpatialReference
  Set pPt2.SpatialReference = gMap.SpatialReference
  pPt2.Project pGeoDataset.SpatialReference
    
  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  pEnv.PutCoords pPt1.X, pPt1.Y, pPt2.X, pPt2.Y
  
  'Create the spatial filter
  Dim pSpatialFilter As ISpatialFilter
  Set pSpatialFilter = New SpatialFilter
  With pSpatialFilter
    Set .Geometry = pEnv
    .GeometryField = "Shape"
    .SpatialRel = esriSpatialRelIntersects
  End With
    
  Dim curDis As Double
  Dim tmpDis As Double
  Dim bRight As Boolean
  Dim minDis As Double
  minDis = 2 * Abs(pPt2.X - pPt1.X)
  
  'Use the spatial filter to select features
  Dim pFeatureCursor As IFeatureCursor
  Set pFeatureCursor = pFCStream.Search(pSpatialFilter, False)
    
  Dim pFeature As IFeature
  Set pFeature = pFeatureCursor.NextFeature
    
  Dim pPolyline As IPolyline
  Dim pNearPt As IPoint
  Dim pNearDis As Double
  Dim pNearFeature As IFeature
  Set pNearFeature = Nothing
  
  Do Until pFeature Is Nothing
    Set pPolyline = pFeature.Shape
    pPolyline.QueryPointAndDistance esriNoExtension, pPt0, True, pPt1, tmpDis, curDis, bRight
    
    If curDis < minDis Then
      minDis = curDis
      hitTestRes(0) = pFeature
      hitTestRes(1) = pPt1
      hitTestRes(2) = tmpDis
    End If
    Set pFeature = pFeatureCursor.NextFeature
  Loop

  HitTestStream = hitTestRes
End Function

Public Function TraceBufferStrip(ByRef pFCStream As IFeatureClass, ByRef pStartFeature As IFeature, ByVal posStart As Double, ByVal traceDis As Double, ByVal bTraceToEnd As Boolean) As IPolyline
  If pFCStream Is Nothing Then Exit Function
  
  Dim strIDName As String, strDSIDName As String
  strIDName = gSUBBASINFieldName
  strDSIDName = gSUBBASINRFieldName
  
  Dim pLayer As IFeatureLayer
  Set pLayer = New FeatureLayer
  Set pLayer.FeatureClass = pFCStream
  
  Dim pTable As IDisplayTable
  Set pTable = pLayer
  
  Dim lStreamIDFldIndex As Long, lDownStreamIDFldIndex As Long
  lStreamIDFldIndex = pTable.DisplayTable.FindField(strIDName)
  lDownStreamIDFldIndex = pTable.DisplayTable.FindField(strDSIDName)
  If lStreamIDFldIndex < 0 Or lDownStreamIDFldIndex < 0 Then
    MsgBox "Required field is missing in Stream layer", vbExclamation
    Exit Function
  End If
  
  Dim strStreamID As String, strFirstID As String
  Dim lineLength As Double
  Dim totalDis As Double
  Dim posFrom As Double
  Dim posTo As Double
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPolyline As IPolyline
  Dim pGeometry As IGeometry
  Dim pSubPolyline As IPolyline

  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter
  
  Dim pBufferStrip As IPolyline
  Set pBufferStrip = New Polyline
  pBufferStrip.SetEmpty
  
  Dim pPtCollection As IPointCollection
  Set pPtCollection = pBufferStrip
  
  Dim pTopoOp As ITopologicalOperator2
  Set pTopoOp = pBufferStrip
  
  totalDis = 0
  strStreamID = pStartFeature.value(lStreamIDFldIndex)
  strFirstID = strStreamID
  
  Dim streamCount As Long
  streamCount = 0
  
  Do
    If strStreamID = "" Or strStreamID = "0" Then
      MsgBox "Tracing operation stopped before reaching to the specified distance."
      Exit Do
    End If
    
    pQueryFilter.WhereClause = strIDName & " = " & strStreamID
    Set pFCursor = pTable.DisplayTable.Search(pQueryFilter, False)
    Set pFeature = pFCursor.NextFeature
    If pFeature Is Nothing Then
      MsgBox "Tracing operation stopped before reaching to the specified distance."
      Exit Do
    End If
    
    Set pPolyline = pFeature.ShapeCopy
    lineLength = pPolyline.Length
    
    If strStreamID = strFirstID Then
        'posFrom = posStart
        If (gIsStreamAlongFlowDir) Then
            lineLength = (1 - posStart) * lineLength
            posFrom = posStart
         Else
            lineLength = posStart * lineLength
            posFrom = 0
         End If
    Else
      posFrom = 0
    End If
    
    'testing
    posTo = 1
    If ((Not gIsStreamAlongFlowDir) And (streamCount = 0) And (posStart <> 1)) Then
        posTo = posStart
    End If
    

    If totalDis + lineLength >= traceDis Then
      If Not bTraceToEnd Then
        If (gIsStreamAlongFlowDir) Then
            posTo = posFrom + (traceDis - totalDis) / pPolyline.Length
        Else
            posFrom = posTo - (traceDis - totalDis) / pPolyline.Length
        End If
      End If
      
      If posFrom <> 0 Or posTo <> 1 Then
        pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
      Else
        Set pSubPolyline = pPolyline
      End If
      
      ConcatenatePolyline pPtCollection, pSubPolyline
      pTopoOp.IsKnownSimple = False
      pTopoOp.Simplify
      
      Set TraceBufferStrip = pBufferStrip
      Exit Do
    Else
      If posFrom <> 0 Or posTo <> 1 Then
        pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
      Else
        Set pSubPolyline = pPolyline
      End If
      
      ConcatenatePolyline pPtCollection, pSubPolyline
    End If
    
    totalDis = totalDis + lineLength
    streamCount = streamCount + 1
    strStreamID = pFeature.value(lDownStreamIDFldIndex)
  Loop
'''  If pFCStream Is Nothing Then Exit Function
'''
'''  Dim strIDName As String, strDSIDName As String
'''  strIDName = gSUBBASINFieldName
'''  strDSIDName = gSUBBASINRFieldName
'''
'''  Dim lStreamIDFldIndex As Long, lDownStreamIDFldIndex As Long
'''  lStreamIDFldIndex = pFCStream.FindField(strIDName)
'''  lDownStreamIDFldIndex = pFCStream.FindField(strDSIDName)
'''  If lStreamIDFldIndex < 0 Or lDownStreamIDFldIndex < 0 Then
'''    MsgBox "Required field is missing in Stream layer", vbExclamation
'''    Exit Function
'''  End If
'''
'''  Dim strStreamID As String, strFirstID As String
'''  Dim lineLength As Double
'''  Dim totalDis As Double
'''  Dim posFrom As Double
'''  Dim posTo As Double
'''  Dim pFCursor As IFeatureCursor
'''  Dim pFeature As IFeature
'''  Dim pPolyline As IPolyline
'''  Dim pGeometry As IGeometry
'''  Dim pSubPolyline As IPolyline
'''
'''  Dim pQueryFilter As IQueryFilter
'''  Set pQueryFilter = New QueryFilter
'''
'''  Dim pBufferStrip As IPolyline
'''  Set pBufferStrip = New Polyline
'''  pBufferStrip.SetEmpty
'''
'''  Dim pPtCollection As IPointCollection
'''  Set pPtCollection = pBufferStrip
'''
'''  Dim pTopoOp As ITopologicalOperator2
'''  Set pTopoOp = pBufferStrip
'''
'''  totalDis = 0
'''  strStreamID = pStartFeature.value(lStreamIDFldIndex)
'''  strFirstID = strStreamID
'''
'''  Do
'''    If strStreamID = "" Or strStreamID = "0" Then
'''      MsgBox "Tracing operation stopped before reaching to the specified distance."
'''      Exit Do
'''    End If
'''
'''    pQueryFilter.WhereClause = strIDName & " = " & strStreamID
'''    Set pFCursor = pFCStream.Search(pQueryFilter, False)
'''    Set pFeature = pFCursor.NextFeature
'''    If pFeature Is Nothing Then
'''      MsgBox "Tracing operation stopped before reaching to the specified distance."
'''      Exit Do
'''    End If
'''
'''    Set pPolyline = pFeature.ShapeCopy
'''    lineLength = pPolyline.Length
'''
'''''    If strStreamID = strFirstID Then
'''''      lineLength = (1 - posStart) * lineLength
'''''      posFrom = posStart
'''''    Else
'''''      posFrom = 0
'''''    End If
'''
'''    '**************to account for the stream direction
'''    If strStreamID = strFirstID Then
'''      If (gIsStreamAlongFlowDir) Then
'''        lineLength = (1 - posStart) * lineLength
'''        posFrom = posStart
'''       Else
'''        lineLength = posStart * lineLength
'''        posFrom = 0
'''       End If
'''    Else
'''      posFrom = 0
'''    End If
'''
'''    'testing
'''    If (gIsStreamAlongFlowDir) Then
'''        posTo = 1
'''    Else
'''        posTo = posStart
'''    End If
'''
'''    'End **************to account for the stream direction
'''    'posTo = 1
'''    If totalDis + lineLength >= traceDis Then
'''      If Not bTraceToEnd Then
'''        posTo = posFrom + (traceDis - totalDis) / pPolyline.Length
'''      End If
'''
'''      If posFrom <> 0 Or posTo <> 1 Then
'''        'pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
'''        If (gIsStreamAlongFlowDir) Then
'''
'''            pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
'''        Else
'''            pPolyline.GetSubcurve 1 - posFrom, 1 - posTo, True, pSubPolyline
'''        End If
'''      Else
'''        Set pSubPolyline = pPolyline
'''      End If
'''
'''      ConcatenatePolyline pPtCollection, pSubPolyline
'''      pTopoOp.IsKnownSimple = False
'''      pTopoOp.Simplify
'''
'''      Set TraceBufferStrip = pBufferStrip
'''      Exit Do
'''    Else
'''      If posFrom <> 0 Or posTo <> 1 Then
'''        pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
'''
'''      Else
'''        Set pSubPolyline = pPolyline
'''      End If
'''
'''      ConcatenatePolyline pPtCollection, pSubPolyline
'''    End If
'''
'''    totalDis = totalDis + lineLength
'''    strStreamID = pFeature.value(lDownStreamIDFldIndex)
'''  Loop
End Function

Public Sub ConcatenatePolyline(ByRef pPtColl As IPointCollection, ByRef pPolyline As IPolyline)
  If pPtColl Is Nothing Or pPolyline Is Nothing Then Exit Sub
  If pPtColl.PointCount = 0 Then
    pPtColl.AddPointCollection pPolyline
    Exit Sub
  End If
  
  Dim pPtStart As IPoint, pPtEnd As IPoint
  Set pPtStart = pPtColl.Point(0)
  Set pPtEnd = pPtColl.Point(-1)
  
  Dim pPtColl1 As IPointCollection
  Set pPtColl1 = pPolyline
  
  Dim pPtStart1 As IPoint, pPtEnd1 As IPoint
  Set pPtStart1 = pPtColl1.Point(0)
  Set pPtEnd1 = pPtColl1.Point(-1)
  
  If pPtStart.X = pPtStart1.X And pPtStart.Y = pPtStart1.Y Then
    pPolyline.ReverseOrientation
    pPtColl.InsertPointCollection 0, pPolyline
  ElseIf pPtStart.X = pPtEnd1.X And pPtStart.Y = pPtEnd1.Y Then
    pPtColl.InsertPointCollection 0, pPolyline
  ElseIf pPtEnd.X = pPtStart1.X And pPtEnd.Y = pPtStart1.Y Then
    pPtColl.AddPointCollection pPolyline
  ElseIf pPtEnd.X = pPtEnd1.X And pPtEnd.Y = pPtEnd1.Y Then
    pPolyline.ReverseOrientation
    pPtColl.AddPointCollection pPolyline
  End If
End Sub

Public Function TraceToInStreamBMP(ByRef pFCStream As IFeatureClass, ByRef pStartFeature As IFeature, ByVal posStart As Double, ByVal strBMPType As String) As IPolyline
  Dim pBMPFLayer As IFeatureLayer
  Set pBMPFLayer = GetInputFeatureLayer("BMPs")
  If pBMPFLayer Is Nothing Then Exit Function
  
  Dim pFCBMP As IFeatureClass
  Set pFCBMP = pBMPFLayer.FeatureClass
  
  Dim pLayer As IFeatureLayer
  Set pLayer = New FeatureLayer
  Set pLayer.FeatureClass = pFCStream
  
  Dim pTable As IDisplayTable
  Set pTable = pLayer

  If pFCStream Is Nothing Then Exit Function
  Dim strIDName As String, strDSIDName As String
  strIDName = gSUBBASINFieldName
  strDSIDName = gSUBBASINRFieldName
  
  Dim lStreamIDFldIndex As Long, lDownStreamIDFldIndex As Long
  lStreamIDFldIndex = pTable.DisplayTable.FindField(strIDName)
  lDownStreamIDFldIndex = pTable.DisplayTable.FindField(strDSIDName)
  If lStreamIDFldIndex < 0 Or lDownStreamIDFldIndex < 0 Then
    MsgBox "Required field is missing in Stream layer", vbExclamation
    Exit Function
  End If
  
  Dim strStreamID As String, strFirstID As String
  Dim posFrom As Double
  Dim posTo As Double
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPolyline As IPolyline
  Dim pGeometry As IGeometry
  Dim pSubPolyline As IPolyline

  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter
  
  Dim pBufferStrip As IPolyline
  Set pBufferStrip = New Polyline
  pBufferStrip.SetEmpty
  
  Dim pPtCollection As IPointCollection
  Set pPtCollection = pBufferStrip
  
  Dim pTopoOp As ITopologicalOperator2
  Set pTopoOp = pBufferStrip
  
  strStreamID = pStartFeature.value(lStreamIDFldIndex)
  strFirstID = strStreamID
  
  Do
    If strStreamID = "" Or strStreamID = "0" Then
      MsgBox "Tracing operation stopped before reaching to the next junction."
      Exit Do
    End If
    
    pQueryFilter.WhereClause = strIDName & " = " & strStreamID
    Set pFCursor = pTable.DisplayTable.Search(pQueryFilter, False)
    Set pFeature = pFCursor.NextFeature
    If pFeature Is Nothing Then
      MsgBox "Tracing operation stopped before reaching to the next junction."
      Exit Do
    End If
    
    Set pPolyline = pFeature.ShapeCopy
    
    If strStreamID = strFirstID Then
      posFrom = posStart
    Else
      posFrom = 0
    End If
    
    posTo = 1
    If FindNextInStreamBMP(strBMPType, strFirstID, strStreamID, pFCBMP, pPolyline, posFrom, posTo) Then
      If posFrom <> 0 Or posTo <> 1 Then
        pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
      Else
        Set pSubPolyline = pPolyline
      End If
      
      ConcatenatePolyline pPtCollection, pSubPolyline
      pTopoOp.IsKnownSimple = False
      pTopoOp.Simplify
      
      Set TraceToInStreamBMP = pBufferStrip
      Exit Do
    Else
      If posFrom <> 0 Or posTo <> 1 Then
        pPolyline.GetSubcurve posFrom, posTo, True, pSubPolyline
      Else
        Set pSubPolyline = pPolyline
      End If
      
      ConcatenatePolyline pPtCollection, pSubPolyline
    End If
    
    strStreamID = pFeature.value(lDownStreamIDFldIndex)
  Loop
End Function

Public Function SnapToInStreamBMP(ByRef pFCStream As IFeatureClass, ByVal X As Long, ByVal Y As Long, ByVal strBMPType As String)
  If pFCStream Is Nothing Then Exit Function
  Dim strIDName As String
  strIDName = gSUBBASINFieldName
  
  Dim pLayer As IFeatureLayer
  Set pLayer = New FeatureLayer
  Set pLayer.FeatureClass = pFCStream
  
  Dim pTable As IDisplayTable
  Set pTable = pLayer
  
  Dim lStreamIDFldIndex As Long, lBMPStreamIDFldIndex As Long
  lStreamIDFldIndex = pTable.DisplayTable.FindField(strIDName)
  If lStreamIDFldIndex < 0 Then
    MsgBox "Required field is missing in Stream layer", vbExclamation
    Exit Function
  End If
  
  Dim pBMPFLayer As IFeatureLayer
  Set pBMPFLayer = GetInputFeatureLayer("BMPs")
  If pBMPFLayer Is Nothing Then Exit Function
  
  Dim pFCBMP As IFeatureClass
  Set pFCBMP = pBMPFLayer.FeatureClass
  
  lBMPStreamIDFldIndex = pFCBMP.FindField("STREAMID")
  If lBMPStreamIDFldIndex < 0 Then
    MsgBox "Required field is missing in BMP layer", vbExclamation
    Exit Function
  End If
  
  Dim hitTestRes(2)
  Set hitTestRes(0) = Nothing
  Set hitTestRes(1) = Nothing
  hitTestRes(2) = 0
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.FocusMap
    
  Dim pScreenDisplay As IScreenDisplay
  Set pScreenDisplay = pActiveView.ScreenDisplay
  
  Dim pDT As IDisplayTransformation
  Set pDT = pScreenDisplay.DisplayTransformation
  
  Dim pPt As IPoint
  Dim pPt1 As IPoint
  Dim pPt2 As IPoint
    
  Set pPt = pDT.ToMapPoint(X, Y)
  Set pPt1 = pDT.ToMapPoint(X - 4, Y - 4)
  Set pPt2 = pDT.ToMapPoint(X + 4, Y + 4)
    
  Dim curDis As Double, minDis As Double, tmpDis As Double
  minDis = 2 * Abs(pPt2.X - pPt1.X)
  
  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter
  pQueryFilter.WhereClause = "TYPE = '" & strBMPType & "' AND STREAMID > 0"
  
  Dim pFCCursor As IFeatureCursor
  Set pFCCursor = pFCBMP.Search(pQueryFilter, False)
  
  Dim pFeature As IFeature
  Set pFeature = pFCCursor.NextFeature
  
  Dim pProximityOp As IProximityOperator
  Set pProximityOp = pPt
  
  Dim pBMPPt As IPoint
  Dim strStreamID As String
  
  Do Until pFeature Is Nothing
    curDis = pProximityOp.ReturnDistance(pFeature.Shape)
    If curDis < minDis Then
      minDis = curDis
      Set pBMPPt = pFeature.ShapeCopy
      strStreamID = pFeature.value(lBMPStreamIDFldIndex)
    End If
    Set pFeature = pFCCursor.NextFeature
  Loop
  
  If pBMPPt Is Nothing Then Exit Function
  
  pQueryFilter.WhereClause = strIDName & " = " & strStreamID
  Set pFCCursor = pTable.DisplayTable.Search(pQueryFilter, False)
  Set pFeature = pFCCursor.NextFeature
  If pFeature Is Nothing Then Exit Function
  
  Dim bRight As Boolean
  Dim pPolyline As IPolyline
  Set pPolyline = pFeature.Shape
  pPolyline.QueryPointAndDistance esriNoExtension, pBMPPt, True, pPt, tmpDis, curDis, bRight
  
  hitTestRes(0) = pFeature
  hitTestRes(1) = pPt
  hitTestRes(2) = tmpDis
  SnapToInStreamBMP = hitTestRes
End Function

Public Function FindNextInStreamBMP(ByVal strType As String, ByVal strFirstID As String, ByVal strStreamID As String, ByRef pFCBMP As IFeatureClass, ByRef pPolyline As IPolyline, ByVal posFrom As Double, ByRef posTo As Double) As Boolean
  If pFCBMP Is Nothing Then Exit Function
  If pPolyline Is Nothing Then Exit Function
  
  Dim pQueryFilter As IQueryFilter
  Set pQueryFilter = New QueryFilter
  pQueryFilter.WhereClause = "TYPE = '" & strType & "' AND STREAMID = " & strStreamID
  
  Dim pFCCursor As IFeatureCursor
  Set pFCCursor = pFCBMP.Search(pQueryFilter, False)
  
  Dim pFeature As IFeature
  Set pFeature = pFCCursor.NextFeature
  
  Dim tmpDis As Double, curDis As Double, minDis As Double
  Dim bRight As Boolean
  Dim pPt As IPoint
  Set pPt = New Point
  
  minDis = 1
  Do Until pFeature Is Nothing
    pPolyline.QueryPointAndDistance esriNoExtension, pFeature.Shape, True, pPt, tmpDis, curDis, bRight
    If strFirstID <> strStreamID Or posFrom <> tmpDis Then ' Prevent the starting Junction is picked up
      If tmpDis >= posFrom And (tmpDis - posFrom) < minDis Then
        minDis = tmpDis - posFrom
        posTo = tmpDis
        FindNextInStreamBMP = True
      End If
    End If
    
    Set pFeature = pFCCursor.NextFeature
  Loop
End Function

'Find if the streams follow the flow direction or against the flow direction
Public Sub SetSreamDirectionFlag()
  
    Dim pFeatureLayer As IFeatureLayer
    Dim pFeatureclass As IFeatureClass
    Dim pFeature As IFeature
    Dim pFeatureCursor As IFeatureCursor
    
    Dim subIdF As Long
    Dim dsIdF As Long
    
    Set pFeatureLayer = GetInputFeatureLayer("STREAM")
    If (pFeatureLayer Is Nothing) Then
        MsgBox "No stream available"
        Exit Sub
    End If
    Dim pTable As IDisplayTable
    Set pTable = pFeatureLayer
    
    Set pFeatureclass = pFeatureLayer.FeatureClass
    
    subIdF = pTable.DisplayTable.FindField(gSUBBASINFieldName)
    dsIdF = pTable.DisplayTable.FindField(gSUBBASINRFieldName)
    
    Set pFeatureCursor = pFeatureclass.Search(Nothing, False)
    
    Set pFeature = pFeatureCursor.NextFeature
    
    Dim subId As String
    Dim dsID As String
    
    Dim pTempFCursor As IFeatureCursor
    Dim ptempFeature As IFeature
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pPolylineWs As IPolyline
    Dim pPolylineDs As IPolyline
    
    Do While Not pFeature Is Nothing
        subId = pFeature.value(subIdF)
        dsID = pFeature.value(dsIdF)
        
        pQueryFilter.WhereClause = gSUBBASINFieldName & " = " & dsID & ""
        
        Set pTempFCursor = pTable.DisplayTable.Search(pQueryFilter, False)
        Set ptempFeature = pTempFCursor.NextFeature
        If (Not ptempFeature Is Nothing) Then
            
            Set pPolylineWs = pFeature.Shape
            Set pPolylineDs = ptempFeature.Shape
            
            If (pPolylineWs.ToPoint.Compare(pPolylineDs.FromPoint) = 0) Then
                gIsStreamAlongFlowDir = True
            ElseIf (pPolylineWs.FromPoint.Compare(pPolylineDs.ToPoint) = 0) Then
                gIsStreamAlongFlowDir = False
            End If
            Exit Do
        End If
        
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    
End Sub

'*******************************************************************************
'Subroutine : DefineVFSSimulationOptions
'Purpose    : Creates a DBASE file in the project temp directory and store the
'             VFSMOD Simulation options
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub DefineVFSSimulationOptions(N As Integer, THETAW As Double, NPOL As Integer, _
    KPG As Integer, CR As Double, MAXITER As Integer, IELOUT As Integer)
    
On Error GoTo ShowError
    Dim pVFSSimOptTable As iTable
    Set pVFSSimOptTable = GetInputDataTable("VFSSimOptions")
    
    If (pVFSSimOptTable Is Nothing) Then
        Set pVFSSimOptTable = CreatePropertiesTableDBF("VFSSimOptions")
        AddTableToMap pVFSSimOptTable
    End If
    Dim iIDindex As Long
    iIDindex = pVFSSimOptTable.FindField("ID")
    Dim iPropNameIndex As Long
    iPropNameIndex = pVFSSimOptTable.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pVFSSimOptTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    'Delete all rows from the table
    pVFSSimOptTable.DeleteSearchedRows Nothing

    Dim pRow As iRow
    
    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "N"
    pRow.value(iPropValueIndex) = N
    pRow.Store
    
    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "THETAW"
    pRow.value(iPropValueIndex) = THETAW
    pRow.Store
 
    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "NPOL"
    pRow.value(iPropValueIndex) = NPOL
    pRow.Store

    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "KPG"
    pRow.value(iPropValueIndex) = KPG
    pRow.Store

    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "CR"
    pRow.value(iPropValueIndex) = CR
    pRow.Store

    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "MAXITER"
    pRow.value(iPropValueIndex) = MAXITER
    pRow.Store


    Set pRow = pVFSSimOptTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "IELOUT"
    pRow.value(iPropValueIndex) = IELOUT
    pRow.Store
    
    Exit Sub
ShowError:
    MsgBox "Error in DefineVFSSimulationOptions :" & Err.description
End Sub


'*******************************************************************************
'Subroutine : GetVFSSimulationOptions
'Purpose    : Creates a DBASE file in the project temp directory and store the
'             VFSMOD Simulation options
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetVFSSimulationOptions() As String
    
On Error GoTo ShowError
    Dim result As String
    result = ""
    
    Dim pVFSSimOptTable As iTable
    Set pVFSSimOptTable = GetInputDataTable("VFSSimOptions")
    
    If (pVFSSimOptTable Is Nothing) Then
        MsgBox "VFSSimOptions table is missing "
        Exit Function
    End If
    
    Dim iIDindex As Long
    iIDindex = pVFSSimOptTable.FindField("ID")
    Dim iPropNameIndex As Long
    iPropNameIndex = pVFSSimOptTable.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pVFSSimOptTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    Dim pCursor As ICursor
    Set pCursor = pVFSSimOptTable.Search(Nothing, True)
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pDict As Scripting.Dictionary
    Set pDict = New Scripting.Dictionary
        
    Do Until pRow Is Nothing
        pDict.Item(pRow.value(iPropNameIndex)) = pRow.value(iPropValueIndex)
        Set pRow = pCursor.NextRow
    Loop
        
    Dim keys() As Variant
    keys() = Array("N", "THETAW", "CR", "MAXITER", "NPOL", "IELOUT", "KPG")

    Dim i As Integer
    For i = 0 To UBound(keys)
        If Not pDict.Exists(keys(i)) Then
            MsgBox "VFS Option " & keys(i) & " is missing "
            Exit Function
        End If
    Next
        
    For i = 0 To UBound(keys)
        result = result & pDict.Item(keys(i)) & " "
    Next
        
    GetVFSSimulationOptions = Trim(result) & vbNewLine
    Exit Function
ShowError:
    MsgBox "Error in GetVFSSimulationOptions :" & Err.description
End Function

'*******************************************************************************
'Subroutine : EditVFSDetails
'Purpose    : Edits the VFS details
'Note       :
'Arguments  :
'Author     : Ying Cao
'History    : 11/19/2008 - Ying Cao
'*******************************************************************************
Public Sub EditVFSDetails(pVFSID As Integer, pVFSType As String)

On Error GoTo ShowError
    
    Dim pVFSDetailDict As Scripting.Dictionary
    Dim pVFSDetailTable As iTable
    Dim pQueryFilter As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Set pVFSDetailTable = GetInputDataTable("VFSDetail")
    
    Dim pIDindex As Long
    pIDindex = pVFSDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pVFSDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pVFSDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim i As Integer
    Dim pVFSKeys
    Dim pSelRowCount As Long
    Dim pVFSName As String
    
    Set pVFSDetailDict = GetVFSDetailDict(pVFSID)
    gNewVFSType = pVFSDetailDict.Item("VFSType")
        
    pVFSName = pVFSDetailDict.Item("VFSName")
    
'    '** Show the FrmTraceDown window and populate the parameters
'    With FrmTraceDown
'        .optNearestPoint.value = (pVFSDetailDict.Item("SnapOptionIndex") = SNAP_NEAREST_POINT)
'        .optNearestNode.value = (pVFSDetailDict.Item("SnapOptionIndex") = SNAP_NEAREST_NODE)
'        .optNearestJunction.value = (pVFSDetailDict.Item("SnapOptionIndex") = SNAP_NEAREST_JUNCTION)
'        .optTraceDown.value = (pVFSDetailDict.Item("TraceOptionIndex") = TRACE_DOWN)
'        .optTraceJunction.value = (pVFSDetailDict.Item("TraceOptionIndex") = TRACE_JUNCTION)
'        .strSnapBMPType = pVFSDetailDict.Item("SnapOptionStr")
'        .strTraceBMPType = pVFSDetailDict.Item("TraceOptionStr")
'        .tbxDistance.Text = pVFSDetailDict.Item("BufferWidth")
'        .cmbVFSTypes.ListIndex = pVFSDetailDict.Item("VFSTypeIndex")
'        .optionLeft.value = (pVFSDetailDict.Item("Bank") = "Left")
'        .optionRight.value = (pVFSDetailDict.Item("Bank") = "Right")
'    End With
'
'    FrmTraceDown.Show vbModal
    
    'InitializeVFSPropertyForm pVFSDictionary
    'FrmVFSParams.BufferWidth.Text = lfBufWidth
    FrmVFSParams.Show vbModal
    
            
    'Remove records for vfs which are not in the gvfsdetaildictionary
'    If Not (gVFSDetailDict Is Nothing) Then
'        Set pQueryFilter = New QueryFilter
'        pQueryFilter.WhereClause = "ID = " & pVFSID
'        Set pCursor = pVFSDetailTable.Update(pQueryFilter, True)
'        Set pRow = pCursor.NextRow
'        Do While Not pRow Is Nothing
'            If (Not gVFSDetailDict.Exists(pRow.value(pPropNameIndex))) Then
'                pCursor.DeleteRow
'            End If
'            Set pRow = pCursor.NextRow
'        Loop
'        Set pRow = Nothing
'        Set pCursor = Nothing
'    End If
'
'    If Not (gVFSDetailDict Is Nothing) Then
'        pVFSKeys = gVFSDetailDict.keys
'        For i = 0 To (gVFSDetailDict.Count - 1)
'            pPropertyName = pVFSKeys(i)
'            pPropertyValue = gVFSDetailDict.Item(pPropertyName)
'            Set pQueryFilter = New QueryFilter
'            pQueryFilter.WhereClause = "ID = " & pVFSID & " AND PropName = '" & pPropertyName & "'"
'            Set pCursor = pVFSDetailTable.Search(pQueryFilter, False)
'            Set pRow = pCursor.NextRow
'            If Not pRow Is Nothing Then
'                pRow.value(pPropValueIndex) = pPropertyValue
'                pRow.Store
'            Else
'                'Create if the row is not already in the table
'                Set pRow = pVFSDetailTable.CreateRow
'                pRow.value(pIDindex) = pVFSID
'                pRow.value(pPropNameIndex) = pPropertyName
'                pRow.value(pPropValueIndex) = pPropertyValue
'                pRow.Store
'            End If
'        Next i
'    End If
        
    GoTo CleanUp
    
ShowError:
    MsgBox "Edit VFS Details :" & Err.description

CleanUp:
    Set pVFSDetailDict = Nothing
    Set pVFSDetailTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub


'*******************************************************************************
'Subroutine : GetVFSDetailDict
'Purpose    : Gets the VFS details for individual VFS from VFSDetail
'Note       :
'Arguments  : Id of the VFS
'Author     : Ying Cao
'History    : 11/19/2008 Ying Cao
'*******************************************************************************
Public Function GetVFSDetailDict(vfsId As Integer, Optional vfsTableName As String) As Dictionary

On Error GoTo ErrorHandler
    Dim pVFSDefaultTable As iTable
    
    If vfsTableName = "" Then
        Set pVFSDefaultTable = GetInputDataTable("VFSDetail")
    Else
        Set pVFSDefaultTable = GetInputDataTable(vfsTableName)
    End If
    
    If (pVFSDefaultTable Is Nothing) Then
         MsgBox "No VFSDetail table in the map: Add the table and continue"
         Exit Function
    End If
    
    Dim pIDindex As Long
    pIDindex = pVFSDefaultTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pVFSDefaultTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pVFSDefaultTable.FindField("PropValue")
                
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pTmpVFSName As String
    Dim pTmpVFSID As Integer
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "ID = " & vfsId

    Dim pCursor As ICursor
    Set pCursor = pVFSDefaultTable.Search(pQueryFilter, False)
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pVFSDetailDict As Scripting.Dictionary
    Set pVFSDetailDict = CreateObject("Scripting.Dictionary")
    
    Do While Not pRow Is Nothing
        pTmpVFSID = pRow.value(pIDindex)
        pPropertyName = pRow.value(pPropNameIndex)
        pPropertyValue = pRow.value(pPropValueIndex)
        If pPropertyName <> "ID" Then
            pVFSDetailDict.add pPropertyName, pPropertyValue
        End If
        Set pRow = pCursor.NextRow
    Loop
    Set GetVFSDetailDict = pVFSDetailDict

    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "GetVFSDetailDict :", Err.description
CleanUp:
    Set pVFSDefaultTable = Nothing
    Set pQueryFilter = Nothing
    Set pVFSDetailDict = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
End Function



'*******************************************************************************
'Subroutine : DeleteSelectedVFS
'Purpose    : Deletes the VFS and the corresponding details from the tables
'             Also modifies the VFS ids
'Note       :
'Arguments  :
'Author     : Ying Cao
'History    : 11/20/2008 - Ying Cao
'*******************************************************************************
Public Sub DeleteSelectedVFS(pSelectedVFSId As Integer)
On Error GoTo ShowError
    
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("VFS")
    Dim pFeatureclass As IFeatureClass
    If Not (pFeatureLayer Is Nothing) Then
        Set pFeatureclass = pFeatureLayer.FeatureClass
    Else
        MsgBox "VFS feature layer not found."
        Exit Sub
    End If
    Dim iVfsIdFld As Long
    iVfsIdFld = pFeatureclass.FindField("ID")
    Dim pVFSID As Integer
    Dim pType As String

    Dim pConduitsFLayer As IFeatureLayer
    Set pConduitsFLayer = GetInputFeatureLayer("Conduits")
    Dim pConduitsFClass As IFeatureClass
    Dim iConduitIDFld As Long
    Dim iConduitFROMFld As Long
    Dim iConduitTOFld As Long
    If Not (pConduitsFLayer Is Nothing) Then
        Set pConduitsFClass = pConduitsFLayer.FeatureClass
        iConduitIDFld = pConduitsFClass.FindField("ID")
        iConduitFROMFld = pConduitsFClass.FindField("CFROM")
        iConduitTOFld = pConduitsFClass.FindField("CTO")
    End If

    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    Dim pWatershedFClass As IFeatureClass
    If Not pWatershedFLayer Is Nothing Then
        Set pWatershedFClass = pWatershedFLayer.FeatureClass
    End If

    Dim pBasinRoutingFLayer As IFeatureLayer
    Set pBasinRoutingFLayer = GetInputFeatureLayer("BasinRouting")
    Dim pBasinRoutingFClass As IFeatureClass
    If Not (pBasinRoutingFLayer Is Nothing) Then
        Set pBasinRoutingFClass = pBasinRoutingFLayer.FeatureClass
    End If

    Dim pVFSDetailTable As iTable
    Set pVFSDetailTable = GetInputDataTable("VFSDetail")
    Dim pIDindex As Long
    pIDindex = pVFSDetailTable.FindField("ID")
    
    Dim pVFSNetworkTable As iTable
    Set pVFSNetworkTable = GetInputDataTable("VFSNetwork")
    Dim pNetIDindex As Long
    Dim pNetDSIDindex As Long
    'pNetIDindex = pVFSNetworkTable.FindField("ID")
    'pNetDSIDindex = pVFSNetworkTable.FindField("DSID")

    Dim pDecayFactTable As iTable
    Set pDecayFactTable = GetInputDataTable("DecayFact")

    Dim pPctRemovalTable As iTable
    Set pPctRemovalTable = GetInputDataTable("PctRemoval")

    Dim pAggDetailTable As iTable
    Set pAggDetailTable = GetInputDataTable("AgVFSDetail")

    Dim pTabCursor As ICursor
    Dim pRow As iRow
    
    Dim pIDArray() As Integer
    Dim pIdCount As Integer
    Dim pIdIncr As Integer
    pIdCount = 0

    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pFeatureCursor1 As IFeatureCursor
    Dim pFeature1 As IFeature
    Dim pFeatureCursor2 As IFeatureCursor
    Dim pFeature2 As IFeature
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pQueryFilter1 As IQueryFilter
    Set pQueryFilter1 = New QueryFilter
    Dim pQueryFilter2 As IQueryFilter
    Set pQueryFilter2 = New QueryFilter
    
    Dim pFeatureEdit As IFeatureEdit
    Dim pDeleteSet As ISet
    
    Dim aggIdField As Integer, aggVFSID As Integer
    
    pQueryFilter.WhereClause = "ID = " & pSelectedVFSId
    Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    
    If Not pFeature Is Nothing Then
        pVFSID = pFeature.value(iVfsIdFld)
        '********* DELETE RECORDS FROM ALL FEATURE LAYERS
        'Delete the features from VFS feature layer with ID = deleted VFS
        'call the subroutine to start editing
        Call StartEditingFeatureLayer("VFS")
        Set pDeleteSet = New esriSystem.Set
        pDeleteSet.add pFeature
        pDeleteSet.Reset
        Set pFeatureEdit = pDeleteSet.Next
        If Not pFeatureEdit Is Nothing Then
          pFeatureEdit.DeleteSet pDeleteSet
        End If
        Call StopEditingFeatureLayer
        
        'Delete all records from the VFSDETAIL table with ID = deleted VFS
        pQueryFilter1.WhereClause = "ID = " & pVFSID
        If Not (pVFSDetailTable Is Nothing) Then
            pVFSDetailTable.DeleteSearchedRows pQueryFilter1
        End If
        
        'Include the option to delete AgVFSDetail
        If Not pAggDetailTable Is Nothing Then
            pQueryFilter1.WhereClause = "PropName='VFSID' And PropValue = '" & pVFSID & "'"
            If pAggDetailTable.RowCount(pQueryFilter1) > 0 Then
                Set pTabCursor = pAggDetailTable.Search(pQueryFilter1, False)
                aggIdField = pAggDetailTable.FindField("ID")
                Set pRow = pTabCursor.NextRow
                Do While Not (pRow Is Nothing)
                    aggVFSID = pRow.value(aggIdField)
                    pQueryFilter1.WhereClause = "ID = " & aggVFSID
                    pAggDetailTable.DeleteSearchedRows pQueryFilter1
                    Set pRow = pTabCursor.NextRow
                Loop
            End If
            
            'Deduct one from all VFSID values greater than pVFSID
            pQueryFilter1.WhereClause = "PropName='VFSID' "
            If pAggDetailTable.RowCount(pQueryFilter1) > 0 Then
                Set pTabCursor = pAggDetailTable.Update(pQueryFilter1, False)
                Set pRow = pTabCursor.NextRow
                Do While Not (pRow Is Nothing)
                    aggVFSID = CInt(pRow.value(pAggDetailTable.FindField("PropValue")))
                    If aggVFSID > pVFSID Then
                        pRow.value(pAggDetailTable.FindField("PropValue")) = CStr(aggVFSID - 1)
                        pRow.Store
                    End If
                    Set pRow = pTabCursor.NextRow
                Loop
            End If
        End If

        'Delete the features from Conduits feature layer with ID = deleted VFS
        Call StartEditingFeatureLayer("VFS")
        If Not (pConduitsFClass Is Nothing) Then
            Set pFeatureCursor1 = Nothing
            Set pFeature1 = Nothing
            pQueryFilter1.WhereClause = "CFROM = " & pVFSID & " OR CTO = " & pVFSID
            Set pFeatureCursor1 = pConduitsFClass.Search(pQueryFilter1, False)
            Set pFeature1 = pFeatureCursor1.NextFeature
            Set pDeleteSet = New esriSystem.Set
            Do While Not (pFeature1 Is Nothing)
                'Delete all records from the VFSDetail table with ID = deleted conduit
                pQueryFilter1.WhereClause = "ID = " & pFeature1.value(iConduitIDFld)
                If Not (pVFSDetailTable Is Nothing) Then
                    pVFSDetailTable.DeleteSearchedRows pQueryFilter1
                End If
                'Add the feature to delete in the deleteset
                pDeleteSet.add pFeature1
                Set pFeature1 = pFeatureCursor1.NextFeature
            Loop
            pDeleteSet.Reset
            Set pFeatureEdit = pDeleteSet.Next
            Do While Not pFeatureEdit Is Nothing
              pFeatureEdit.DeleteSet pDeleteSet
              Set pFeatureEdit = pDeleteSet.Next
            Loop
        End If
        Call StopEditingFeatureLayer
        
        'Delete the features from BasinRouting feature layer with ID = deleted VFS
        If (Not pWatershedFClass Is Nothing) And (Not pBasinRoutingFClass Is Nothing) Then
            Set pFeatureCursor1 = Nothing
            Set pFeature1 = Nothing
            pQueryFilter1.WhereClause = "VFSID = " & pVFSID
            Set pFeatureCursor1 = pWatershedFClass.Search(pQueryFilter1, True)
            Dim iWaterFld As Long
            iWaterFld = pFeatureCursor1.FindField("ID")
            Set pFeature1 = pFeatureCursor1.NextFeature
            Do While Not (pFeature1 Is Nothing)
                pQueryFilter2.WhereClause = "ID = " & pFeature1.value(iWaterFld)
                Set pFeatureCursor2 = pBasinRoutingFClass.Search(pQueryFilter2, True)
                Set pFeature2 = pFeatureCursor2.NextFeature
                If (Not pFeature2 Is Nothing) Then
                    pFeature2.Delete
                End If
                Set pFeature1 = pFeatureCursor1.NextFeature
            Loop
            Set pFeature2 = Nothing
            Set pFeatureCursor2 = Nothing
            Set pFeature1 = Nothing
            Set pFeatureCursor1 = Nothing
        End If

        '********** DELETE RECORDS FROM ALL TABLES
        'Delete the details from the VFSNETWORK table with ID = deleted VFS
        pQueryFilter1.WhereClause = "ID = " & pVFSID
        If Not (pVFSNetworkTable Is Nothing) Then
            pVFSNetworkTable.DeleteSearchedRows pQueryFilter1
        End If

        pQueryFilter1.WhereClause = "DSID = " & pVFSID
        If Not (pVFSNetworkTable Is Nothing) Then
            Set pTabCursor = pVFSNetworkTable.Search(pQueryFilter1, False)
            Set pRow = pTabCursor.NextRow
            Do While Not (pRow Is Nothing)
                pRow.value(pNetDSIDindex) = 0
                pRow.Store
                Set pRow = pTabCursor.NextRow
            Loop
        End If

        'Delete the details from DECAYFACT table with ID = deleted VFS
        pQueryFilter1.WhereClause = "VFSID = " & pVFSID
        If Not (pDecayFactTable Is Nothing) Then
            pDecayFactTable.DeleteSearchedRows pQueryFilter1
        End If

        'Delete the details from PCTREMOVAL table with ID = deleted VFS
        pQueryFilter1.WhereClause = "VFSID = " & pVFSID
        If Not (pPctRemovalTable Is Nothing) Then
            pPctRemovalTable.DeleteSearchedRows pQueryFilter1
        End If
        
        Set pTabCursor = Nothing
        Set pRow = Nothing
    End If
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing

    gMxDoc.ActiveView.ContentsChanged
    gMxDoc.ActiveView.Refresh

    GoTo CleanUp
    
ShowError:
    MsgBox "DeleteSelectedVFS:  " & vbTab & Err.Number & vbTab & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pConduitsFLayer = Nothing
    Set pConduitsFClass = Nothing
    Set pWatershedFLayer = Nothing
    Set pWatershedFClass = Nothing
    Set pBasinRoutingFLayer = Nothing
    Set pBasinRoutingFClass = Nothing
    Set pVFSDetailTable = Nothing
    Set pVFSNetworkTable = Nothing
    Set pDecayFactTable = Nothing
    Set pPctRemovalTable = Nothing
    Set pTabCursor = Nothing
    Set pRow = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pFeatureCursor1 = Nothing
    Set pFeature1 = Nothing
    Set pFeatureCursor2 = Nothing
    Set pFeature2 = Nothing
    Set pQueryFilter = Nothing
    Set pQueryFilter1 = Nothing
    Set pQueryFilter2 = Nothing
    Set pFeatureEdit = Nothing
    Set pDeleteSet = Nothing
End Sub

Public Sub Define_Bufferstrip()
    FrmVFSParams.InitCostFromDB
    
    FrmVFSParams.txtName.Text = gNewBMPName
    FrmVFSParams.txtName.Enabled = True
    FrmVFSParams.BufferLength.Text = ""
    FrmVFSParams.txtName.Enabled = True
    FrmVFSParams.BufferWidth.Text = ""
    FrmVFSParams.BufferWidth.Enabled = True
    
    '** Open the VFS Defaults table to get default name
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If (pTable Is Nothing) Then
        '** open the form that defines the buffer strip params
'        FrmVFSData.txtVFSID.Text = 1
'        FrmVFSData.txtName.Text = "VFS1"
'        FrmVFSData.Show vbModal
        
        Dim pVFSDictionary As Scripting.Dictionary
        Set pVFSDictionary = GetDefaultsForVFS(1, "VFS1")
            
        InitializeVFSPropertyForm pVFSDictionary
        FrmVFSParams.Show vbModal
    Else
        FrmVFSTypes.Show vbModal
    End If
    Set pTable = Nothing
       
    
    If (FrmVFSParams.bContinue = True) Then
        Dim pIDValue As Integer
        pIDValue = FrmVFSParams.txtVFSID.Text
        
'        '** create the dictionary
'        Set gBufferStripDetailDict = CreateObject("Scripting.Dictionary")
'        gBufferStripDetailDict.Add "Name", FrmVFSData.txtName.Text
'        gBufferStripDetailDict.Add "BufferLength", FrmVFSData.txtBufferLength.Text
'        gBufferStripDetailDict.Add "BufferWidth", FrmVFSData.txtBufferWidth.Text
        
        '** call the generic function to create and add rows for values
        ModuleVFSFunctions.SaveVFSPropertiesTable "VFSDefaults", CStr(pIDValue), gBufferStripDetailDict
            
        '** set it to nothing
        Set gBufferStripDetailDict = Nothing
        Unload FrmVFSParams
    End If
    
End Sub
