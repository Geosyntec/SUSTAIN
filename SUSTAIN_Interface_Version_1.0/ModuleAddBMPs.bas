Attribute VB_Name = "ModuleAddBMPs"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleAddBMPs
'   Purpose:     Add BMPs on the map, open corresponding bmp dialog,
'                and prompt user to enter bmp parameters.
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  08/../2004 - Mira Chokshi
'                Modified: 08/19/2004 - Sabu Paul added comments to project
'
'******************************************************************************
Option Explicit
Option Base 0

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public bContinue As Variant
Public gBMPTypeToolbox As String
Public gDisplayBMPTemplate As Boolean
Public gSnapBMPToStream As Boolean

'*******************************************************************************
'Subroutine : AddBMPNetworkInformation
'Purpose    : Add record for Routing Network routing
'Arguments  : Id of the BMP site, Flag to decide its a splitter or regular BMP site
'Author     : Mira Chokshi
'History    :
'*******************************************************************************

Public Sub AddBMPNetworkInformation(iBMPID As Integer, boolSplitter As Boolean, boolRegulator As Boolean)
On Error GoTo ShowError
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPNetwork")
    If (pTable Is Nothing) Then
        Set pTable = CreateBMPRoutingDBF("BMPNetwork")
        AddTableToMap pTable
    End If
    Dim iID As Long
    Dim iType As Long
    Dim iDesc As Long
    iID = pTable.FindField("ID")
    iType = pTable.FindField("OutletType")
    iDesc = pTable.FindField("TypeDesc")
    Dim pRow As iRow
    If (boolSplitter = True) Then
        'Define Weir outlet
        Set pRow = pTable.CreateRow
        pRow.value(iID) = iBMPID
        pRow.value(iType) = 2
        pRow.value(iDesc) = "Weir"
        pRow.Store
        'Define orifice outlet
        Set pRow = pTable.CreateRow
        pRow.value(iID) = iBMPID
        pRow.value(iType) = 3
        pRow.value(iDesc) = "Orifice/Channel"
        pRow.Store
        If (boolRegulator = False) Then
            'Define underdrain outlet
            Set pRow = pTable.CreateRow
            pRow.value(iID) = iBMPID
            pRow.value(iType) = 4
            pRow.value(iDesc) = "Underdrain"
            pRow.Store
        End If
    Else
        'MsgBox iBMPID & " is inserted into bmpnetwork table"
        'Define total outlet
        Set pRow = pTable.CreateRow
        pRow.value(iID) = iBMPID
        pRow.value(iType) = 1
        pRow.value(iDesc) = "Total"
        pRow.Store
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "AddBMPNetworkInformation :", Err.description
    
CleanUp:
    Set pRow = Nothing
    Set pTable = Nothing
    
End Sub

'*******************************************************************************
'Subroutine : AddBMPOnMap
'Purpose    : Add a new BMP site. User needs to click at the location
'Author     : Mira Chokshi
'*******************************************************************************

Public Sub AddBMPOnLand(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long, bType As String)
On Error GoTo ShowError
    Dim pDispTrans As IDisplayTransformation
    Dim pActView As IActiveView
    Set pActView = gMxDoc.ActiveView
    Dim pDisp As IScreenDisplay
    Set pDisp = pActView.ScreenDisplay
    Set pDispTrans = pDisp.DisplayTransformation
    Dim pMapPoint As IPoint
    Set pMapPoint = pDispTrans.ToMapPoint(X, Y)
    AddBMPOnMap Button, Shift, pMapPoint, bType, 0
    GoTo CleanUp
ShowError:
    MsgBox "AddBMPOnLand: " & Err.description
CleanUp:
    Set pDispTrans = Nothing
    Set pActView = Nothing
    Set pDisp = Nothing
    Set pMapPoint = Nothing
End Sub

Public Sub AddBMPOnMap(ByVal Button As Long, ByVal Shift As Long, ByVal pMapPoint As IPoint, bType As String, pLngStreamID As Long)
        
On Error GoTo ShowError
   
   'Flash the watershed polygon thrice
   If (bType = "GreenRoof" Or bType = "PorousPavement") Then
        Dim pWatershedFLayer As IFeatureLayer
        Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
        If (pWatershedFLayer Is Nothing) Then
            MsgBox "Watershed feature layer not found. Watershed layer required for " & bType & " BMP."
            Exit Sub
        End If
        Dim pFeatureclass As IFeatureClass
        Set pFeatureclass = pWatershedFLayer.FeatureClass
        Dim pSpatialFilter As ISpatialFilter
        Set pSpatialFilter = New SpatialFilter
        Set pSpatialFilter.Geometry = pMapPoint
        pSpatialFilter.SpatialRel = esriSpatialRelWithin
        If (pFeatureclass.FeatureCount(pSpatialFilter) <> 1) Then
            MsgBox "BMP should be placed on a drainage area."
            Exit Sub
        End If
        Dim pFeatureCursor As IFeatureCursor
        Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, True)
        Dim pWatershedFeat As IFeature
        Set pWatershedFeat = pFeatureCursor.NextFeature
        If Not (pWatershedFeat Is Nothing) Then
            'Get the area
            Dim pAreaMeter As Double
            pAreaMeter = ReturnPolygonArea(pWatershedFeat)
            Dim pPercentDA As Double
            If (gBMPDetailDict.Item("PercentDA")) Then
                pPercentDA = gBMPDetailDict.Item("PercentDA")
            End If
    
            'Get meters per unit factor and convert Meter to feet, Mira Chokshi 03/03/2005
            Dim pDimension As Double
            pDimension = Sqr((pPercentDA / 100) * pAreaMeter) * gMetersPerUnit * 3.28
            pDimension = Format(pDimension, "#.0")
            'Update the length and width of the bmp
           
            gBMPDetailDict.Item("BMPLength") = pDimension
            gBMPDetailDict.Item("BMPWidth") = pDimension
            FlashWatershedFeature pWatershedFeat
        End If
   End If
   
    Dim pPourPointFLayer As IFeatureLayer
    Set pPourPointFLayer = GetInputFeatureLayer("BMPs")
    Dim pPourPointFClass As IFeatureClass
    If (pPourPointFLayer Is Nothing) Then
        Set pPourPointFClass = CreateFeatureClassForBMPOrVFS(gMapTempFolder, "bmp", "Point")
        'Delete Existing BMPDetails table
        DeleteDataTable gMapTempFolder, "BMPDetail"
        DeleteDataTable gMapTempFolder, "BMPNetwork"
'        DeleteDataTable gMapTempFolder, "DecayFact"
'        DeleteDataTable gMapTempFolder, "PctRemoval"
        DeleteDataTable gMapTempFolder, "ExternalTS"
        DeleteDataTable gMapTempFolder, "OptimizationDetail"
        DeleteDataTable gMapTempFolder, "AgBMPDetail"
        DeleteDataTable gMapTempFolder, "AgLuDistribution"
        DeleteLayerFromMap "BasinRouting"
        DeleteLayerFromMap "Conduits"
    Else
        Set pPourPointFClass = pPourPointFLayer.FeatureClass
    End If
  
    'Get total number of conduits on map
    Dim pConduitFLayer As IFeatureLayer
    Set pConduitFLayer = GetInputFeatureLayer("Conduits")
    Dim pConduitFClass As IFeatureClass
    Dim pConduitCount As Integer
    pConduitCount = 0
    If Not (pConduitFLayer Is Nothing) Then
        Set pConduitFClass = pConduitFLayer.FeatureClass
        pConduitCount = pConduitFClass.FeatureCount(Nothing)    'Get conduit feature count
    End If

    'Get total number of conduits on map
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    Dim pVFSCount As Integer
    pVFSCount = 0
    If Not (pVFSFLayer Is Nothing) Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
        pVFSCount = pVFSFClass.FeatureCount(Nothing)    ' Get vfs feature count
    End If
    
    'Add new bmp (point) feature
    Dim pBMPCount As Integer
    pBMPCount = pPourPointFClass.FeatureCount(Nothing)  'Get bmp feature count
    
    'Get new bmp feature ID
    Dim pBMPID As Integer
    pBMPID = pConduitCount + pBMPCount + pVFSCount + 1

    'Get a new label for bmp layer
    Dim pLabel As String
    Dim pLabelCount As Integer
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "TYPE = '" & gNewBMPType & "'"
    pLabelCount = pPourPointFClass.FeatureCount(pQueryFilter)
    pLabel = Left(gNewBMPType, 1) & CStr(pLabelCount + 1)
    
    'Add a new bmp feature
    Dim pFeature As IFeature
    If Button = 1 Then
        Set pFeature = pPourPointFClass.CreateFeature
        Set pFeature.Shape = pMapPoint
        pFeature.value(pPourPointFClass.FindField("ID")) = pBMPID
        pFeature.value(pPourPointFClass.FindField("TYPE")) = gNewBMPType
        pFeature.value(pPourPointFClass.FindField("TYPE2")) = gNewBMPType
        pFeature.value(pPourPointFClass.FindField("LABEL")) = pLabel
        pFeature.value(pPourPointFClass.FindField("STREAMID")) = pLngStreamID
  
        
        'Added to modify the rendering part -- Sabu Paul
        If gNewBMPType <> "AssessmentPoint" Then
           If gBMPDetailDict.Item("isAssessmentPoint") = "True" Then
                pFeature.value(pPourPointFClass.FindField("TYPE2")) = gNewBMPType & "X"
            End If
        End If
        pFeature.Store
    End If
    
    '** Add BMPs feature layer to the map
    If (pPourPointFLayer Is Nothing) Then
        Set pPourPointFLayer = New FeatureLayer
        Set pPourPointFLayer.FeatureClass = pPourPointFClass
        AddLayerToMap pPourPointFLayer, "BMPs"
        gMxDoc.ActiveView.Refresh
        gMxDoc.UpdateContents
    End If
        
    'Call subroutine to render the bmp layer
    RenderSchematicBMPLayer pPourPointFLayer
    
    gNewBMPId = pBMPID
    AddBMPInformation gBMPDetailDict
    'Third argument is for regulator, which should be false
    AddBMPNetworkInformation pBMPID, bSplitter, bRegulator
   
    '*** Call the subroutine to open the BMP Details form again
    If (gDisplayBMPTemplate = True) Then
        If (gNewBMPType <> "Junction" And gNewBMPType <> "VirtualOutlet" And gNewBMPType <> "Regulator") Then
            EditBmpDetails pBMPID, gNewBMPType
        End If
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "AddBMPOnMap :" & Err.description

CleanUp:
    Set pPourPointFLayer = Nothing
    Set pPourPointFClass = Nothing
    Set pFeature = Nothing

End Sub


Public Function ReturnPolygonArea(pFeature As IFeature) As Double
        'Get the area
        Dim pPolygon As IPolygon
        Set pPolygon = pFeature.Shape
        
        Dim pArea As IArea
        Set pArea = pPolygon
        Dim pAreaMeter As Double
        pAreaMeter = pArea.Area    'Return area in projection units

        Set pPolygon = Nothing
        Set pArea = Nothing
        ReturnPolygonArea = Abs(pAreaMeter) 'Always return absolute value of area

End Function

'Subroutine to flash the watershed feature
Public Sub FlashWatershedFeature(pFeature As IFeature)
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.ActiveView
  
  ' Start Drawing on screen
  pActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
  
  If (pFeature.Shape.GeometryType <> esriGeometryPolygon) Then
    Exit Sub
  End If
  
  Dim pFillSymbol As ISimpleFillSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pFillSymbol = New SimpleFillSymbol
  pFillSymbol.Outline = Nothing
  
  Set pRGBColor = New RgbColor
  pRGBColor.RGB = RGB(0, 0, 250)
 
  
  Set pSymbol = pFillSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pActiveView.ScreenDisplay.SetSymbol pFillSymbol
  pActiveView.ScreenDisplay.DrawPolygon pFeature.Shape
  Sleep 300
  pActiveView.ScreenDisplay.DrawPolygon pFeature.Shape
  
  ' Finish drawing on screen
  pActiveView.ScreenDisplay.FinishDrawing
End Sub


'*******************************************************************************
'Subroutine : AddBMPFromToolbox
'Purpose    : Common routine to different bmp types
'Arguments  : Properties of mouse click. i.e. X, Y, Button, etc and BMP Type
'Author     : Mira Chokshi
'*******************************************************************************
Public Sub AddBMPFromToolbox(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long, bType As String)

   
    bContinue = MsgBox("Do you want to add a BMP at this point ?" & _
                        " Click Yes to continue and No to cancel. ", vbYesNo, "Add BMP")
    If (bContinue = vbYes) Then
        
        ' ********************************************
        ' Now check for the Suitability from this BMP.......
        ' ********************************************
        Dim pFlayer As IFeatureLayer
        Set pFlayer = GetInputFeatureLayer("Composite")
        If Not pFlayer Is Nothing Then
            Dim pActiveView As IActiveView
            Set pActiveView = gMap
            Dim pPoint As esriGeometry.IPoint
            Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                        
            Dim pSelFeatCursor As IFeatureCursor
            Dim pSelFeature As IFeature
            Set pSelFeatCursor = SelectByLocationIN(gMap, pFlayer, pPoint, pFlayer.FeatureClass.ShapeFieldName)
            Set pSelFeature = pSelFeatCursor.NextFeature
            If Not pSelFeature Is Nothing Then
                If pSelFeature.Fields.FindField("BMP_Combin") > -1 Then
                    If Not StringContains(UCase(pSelFeature.value(pSelFeature.Fields.FindField("BMP_Combin"))), UCase(Get_BMP_MappingName(bType))) Then
                        If MsgBox("This is not a suitable area for this BMP. Do you want to continue?", vbInformation + vbYesNo, "SUSTAIN") = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
    
        'Delete subwatershed and snappoints and conduits layer
        DeleteLayerFromMap ("SubWatershed")
        DeleteLayerFromMap ("SnapPoints")
        DeleteLayerFromMap ("Schematic BMPs")
        DeleteLayerFromMap ("Schematic Route")
        gMxDoc.ActiveView.Refresh
        gMxDoc.UpdateContents
        If (gToggleLayer = "Schematic BMPs") Then
            Call ToggleSchematicLayer
            gToggleLayer = "BMPs"
        End If
        gBMPTypeToolbox = bType
        If (gBMPTypeToolbox = "Regulator") Then
            FrmRegulator.Show vbModal
            gNewBMPType = gBMPTypeToolbox
            bSplitter = True
            bRegulator = True
        Else
            FrmSelectBMP.Form_Initialize
            FrmSelectBMP.Show vbModal
            bRegulator = False
        End If
        If (bContinue = vbYes) Then
            If (gSnapBMPToStream = False) Then
                AddBMPOnLand Button, Shift, X, Y, bType
            Else
                SnapBMPToClosestStream Button, Shift, X, Y, bType
            End If
        End If
    End If
    
End Sub

'*******************************************************************************
'Subroutine : CreateFeatureClassForBMPOrVFS
'Purpose    : Creates a new feature class file to store the point shapes
'Arguments  : Destination directory, name of the feature class file
'Author     : Mira Chokshi
'*******************************************************************************
Public Function CreateFeatureClassForBMPOrVFS(DirName As String, FileName As String, pFClassType As String) As IFeatureClass

On Error GoTo ShowError

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
  Dim pFieldsEdit As IFieldsEdit
  Set pFields = New esriGeoDatabase.Fields
  Set pFieldsEdit = pFields

  Dim pFieldShape As esriGeoDatabase.IField
  Dim pFieldEditShape As IFieldEdit

  ' Make the shape field
  ' it will need a geometry definition, with a spatial reference
  Set pFieldShape = New esriGeoDatabase.Field
  Set pFieldEditShape = pFieldShape
  pFieldEditShape.name = strShapeFieldName
  pFieldEditShape.Type = esriFieldTypeGeometry

  Dim pFieldID As esriGeoDatabase.IField
  Dim pFieldEditID As IFieldEdit
  Set pFieldID = New esriGeoDatabase.Field
  Set pFieldEditID = pFieldID
  pFieldEditID.name = "ID"
  pFieldEditID.Type = esriFieldTypeInteger
  pFieldEditID.IsNullable = True

  Dim pFieldType As esriGeoDatabase.IField
  Dim pFieldEditType As IFieldEdit
  Set pFieldType = New esriGeoDatabase.Field
  Set pFieldEditType = pFieldType
  pFieldEditType.name = "TYPE"
  pFieldEditType.Type = esriFieldTypeString
  pFieldEditType.Length = 30
  pFieldEditType.IsNullable = True
  
  Dim pFieldType2 As esriGeoDatabase.IField
  Dim pFieldEditType2 As IFieldEdit
  Set pFieldType2 = New esriGeoDatabase.Field
  Set pFieldEditType2 = pFieldType2
  pFieldEditType2.name = "TYPE2"
  pFieldEditType2.Type = esriFieldTypeString
  pFieldEditType2.Length = 30
  pFieldEditType2.IsNullable = True
  
  Dim pFieldDSID As esriGeoDatabase.IField
  Dim pFieldEditDSID As IFieldEdit
  Set pFieldDSID = New esriGeoDatabase.Field
  Set pFieldEditDSID = pFieldDSID
  pFieldEditDSID.name = "DSID"
  pFieldEditDSID.Type = esriFieldTypeInteger
  pFieldEditDSID.IsNullable = True
  
  Dim pFieldLABEL As esriGeoDatabase.IField
  Dim pFieldEditLABEL As IFieldEdit
  Set pFieldLABEL = New esriGeoDatabase.Field
  Set pFieldEditLABEL = pFieldLABEL
  pFieldEditLABEL.name = "LABEL"
  pFieldEditLABEL.Type = esriFieldTypeString
  pFieldEditLABEL.IsNullable = True
  
  Dim pFieldStreamID As esriGeoDatabase.IField
  Dim pFieldEditStreamID As IFieldEdit
  Set pFieldStreamID = New esriGeoDatabase.Field
  Set pFieldEditStreamID = pFieldStreamID
  pFieldEditStreamID.name = "STREAMID"
  pFieldEditStreamID.Type = esriFieldTypeInteger
  pFieldEditStreamID.DefaultValue = 0
  pFieldEditStreamID.IsNullable = True
  
  Dim pRasterDEMProps As IRasterAnalysisProps
  If Not GetInputRasterLayer("DEM") Is Nothing Then
     Set pRasterDEMProps = GetInputRasterLayer("DEM").Raster
  Else
    Set pRasterDEMProps = Nothing
  End If

  'if DEM is optional use Land use
  Dim pRasterLUProps As IRasterAnalysisProps
  Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster

  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    Select Case pFClassType
        Case "Point":
            .GeometryType = esriGeometryPoint
        Case "Polyline":
            .GeometryType = esriGeometryPolyline
    End Select
    If Not pRasterDEMProps Is Nothing Then
        Set .SpatialReference = pRasterDEMProps.AnalysisExtent.SpatialReference
    Else
        Set .SpatialReference = pRasterLUProps.AnalysisExtent.SpatialReference
    End If
  End With
  Set pFieldEditShape.GeometryDef = pGeomDef
  pFieldsEdit.AddField pFieldShape
  pFieldsEdit.AddField pFieldID
  pFieldsEdit.AddField pFieldDSID
  pFieldsEdit.AddField pFieldType
  pFieldsEdit.AddField pFieldType2
  pFieldsEdit.AddField pFieldLABEL
  pFieldsEdit.AddField pFieldStreamID

  ' Create the shapefile some parameters apply to geodatabase options and can be defaulted as Nothing
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(pFileName, pFields, Nothing, Nothing, esriFTSimple, strShapeFieldName, "")

  ' Return the value
  Set CreateFeatureClassForBMPOrVFS = pFeatClass
    
  GoTo CleanUp
ShowError:
    MsgBox "CreateFeatureClassForBMPOrVFS: " & Err.description
CleanUp:

End Function


Public Function SnapBMPToClosestStream(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long, bType As String) As IPoint
On Error GoTo ShowError
    
    'Get the device units and convert them to map units
    Dim pActView As IActiveView
    Set pActView = gMxDoc.ActiveView
    Dim pDisp As IScreenDisplay
    Set pDisp = pActView.ScreenDisplay
    Dim pDispTrans As IDisplayTransformation
    Set pDispTrans = pDisp.DisplayTransformation
    Dim pPoint As IPoint
    Set pPoint = pDispTrans.ToMapPoint(X, Y)
    Dim pSnapPoint As IPoint
    Set pSnapPoint = New Point
    
    'Get the stream layer and get selected feature
    Dim pSTREAMLayer As IFeatureLayer
    Set pSTREAMLayer = GetInputFeatureLayer("STREAM")
    Dim pSTREAMFClass As IFeatureClass
    Set pSTREAMFClass = pSTREAMLayer.FeatureClass
    Dim pSTREAMSelection As IFeatureSelection
    Set pSTREAMSelection = pSTREAMLayer
    Dim pSelection As ISelectionSet
    Set pSelection = pSTREAMSelection.SelectionSet
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pProximityOp As IProximityOperator
    Dim pStreamID As Long
    Dim iSubBasinFld As Long
    iSubBasinFld = pSTREAMFClass.FindField(gSUBBASINFieldName)
    
    'Ask user about snapping choice
    Dim boolSelected
    Dim boolClosest
    boolSelected = vbNo
    boolClosest = vbNo
    If (pSelection.Count = 1) Then
        boolSelected = MsgBox("Do you want to snap BMP to selected stream ?", vbYesNo, "Snap BMP")
    ElseIf (pSelection.Count > 1 Or pSelection.Count = 0) Then
        boolClosest = MsgBox("You have more than 1 stream or no streams selected, do you want to snap BMP to closest stream ?", vbYesNo, "Snap BMP")
    End If
    
    'Find the closest point on selected stream and move the point there
    If (boolSelected = vbYes) Then
        pSelection.Search Nothing, False, pFeatureCursor
        Set pFeature = pFeatureCursor.NextFeature
        If Not (pFeature Is Nothing) Then
            pStreamID = pFeature.value(iSubBasinFld)
            Set pProximityOp = pFeature.Shape
            Set pSnapPoint = pProximityOp.ReturnNearestPoint(pPoint, esriNoExtension)
        End If
    End If
    
    'Find the closest point on all streams and move the point there
    Dim pClosestDistance As Double
    Dim pDistance As Double
    
  
      '* Get the current map extent
      Dim pMapEnvelope As IEnvelope
      Set pMapEnvelope = pActView.Extent
    
      Dim pSpatialFilter As ISpatialFilter
      Set pSpatialFilter = New SpatialFilter
      Set pSpatialFilter.Geometry = pMapEnvelope
      pSpatialFilter.SpatialRel = esriSpatialRelIntersects
  
    If (boolClosest = vbYes) Then
        pClosestDistance = 1E+32
        'Set the proximity operator to the stream feature
        Set pFeatureCursor = pSTREAMFClass.Search(pSpatialFilter, True)
        Set pFeature = pFeatureCursor.NextFeature
        Do While Not pFeature Is Nothing
            Set pProximityOp = pFeature.Shape
On Error GoTo ProximityError
            pDistance = pProximityOp.ReturnDistance(pPoint)
            If (pDistance < pClosestDistance) Then
                pClosestDistance = pDistance
                Set pSnapPoint = pProximityOp.ReturnNearestPoint(pPoint, esriNoExtension)
                pStreamID = pFeature.value(iSubBasinFld)
            End If
ProximityError:
     Err.Clear
     Resume Next
            Set pFeature = pFeatureCursor.NextFeature
        Loop
    End If
    
    ' If user wants to add a snapped bmp convert back to screen units
    ' Call the subroutine to add bmp on stream
    If (boolClosest = vbYes Or boolSelected = vbYes) Then
        AddBMPOnMap Button, Shift, pSnapPoint, bType, pStreamID
    End If
    GoTo CleanUp

ShowError:
    MsgBox "SnapBMPToClosestStream: " & Err.description
   
CleanUp:
    Set pActView = Nothing
    Set pDisp = Nothing
    Set pDispTrans = Nothing
    Set pPoint = Nothing
    Set pSnapPoint = Nothing
    Set pSTREAMLayer = Nothing
    Set pSTREAMFClass = Nothing
    Set pSTREAMSelection = Nothing
    Set pSelection = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pProximityOp = Nothing
End Function


