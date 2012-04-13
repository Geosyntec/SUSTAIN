Attribute VB_Name = "ModuleEditWatershed"
Option Explicit

Public g_pEditor As IEditor
Public gMapTopology As IMapTopology
  
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
  
  
Public Sub StartEditingFeatureLayer(pFeatureLayerName As String)

  Dim pID As New UID
  Dim TaskCount As Integer
  
  pID = "esriEditor.Editor"
  Set g_pEditor = gApplication.FindExtensionByCLSID(pID)
    
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer(pFeatureLayerName)
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  
  Dim pDataset As IDataset
  Set pDataset = pFeatureLayer.FeatureClass
  Dim pWorkspace As IWorkspace
  Set pWorkspace = pDataset.Workspace
  g_pEditor.StartEditing pWorkspace
  
End Sub

Public Sub StopEditingFeatureLayer()
    If Not (g_pEditor Is Nothing) Then
        g_pEditor.StopEditing True
    End If
End Sub

'Start editing the watershed edges
Public Sub EditWatershedEdges()

  'Loop through the edit tasks checking each one's name
  Dim pEditTask As IEditTask
  Dim TaskCount As Integer
  For TaskCount = 0 To g_pEditor.TaskCount - 1
    Set pEditTask = g_pEditor.Task(TaskCount)
    If pEditTask.name = "Modify Edge" Then
      Set g_pEditor.CurrentTask = pEditTask
    End If
  Next TaskCount
  
  'Make the topology tool active
  'Get the topology extension
  Dim pTopologyExtension As ITopologyExtension
  Dim pUID As UID
  Set pUID = New UID
  pUID.value = "esriEditorExt.topologyextension"
  Set pTopologyExtension = gApplication.FindExtensionByCLSID(pUID)
  If pTopologyExtension Is Nothing Then Exit Sub
    
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass
     
  'Set the map topology active
  Set gMapTopology = pTopologyExtension.MapTopology
  gMapTopology.ClearClasses
  gMapTopology.AddClass pFeatureclass
        
End Sub



'Start editing the watershed layer to add new feature
Public Sub AddNewWatershed()

  'Loop through the edit tasks checking each one's name
  Dim pEditTask As IEditTask
  Dim TaskCount As Integer
  For TaskCount = 0 To g_pEditor.TaskCount - 1
    Set pEditTask = g_pEditor.Task(TaskCount)
    If pEditTask.name = "Create New Feature" Then
      Set g_pEditor.CurrentTask = pEditTask
    End If
  Next TaskCount
  
  'Make the topology tool active
  'Get the topology extension
  Dim pTopologyExtension As ITopologyExtension
  Dim pUID As UID
  Set pUID = New UID
  pUID.value = "esriEditorExt.TopologyExtension"
  Set pTopologyExtension = gApplication.FindExtensionByCLSID(pUID)
  If pTopologyExtension Is Nothing Then Exit Sub
    
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass
     
  'Set the map topology active
  Set gMapTopology = pTopologyExtension.MapTopology
  gMapTopology.ClearClasses
  gMapTopology.AddClass pFeatureclass
        
End Sub

'Start editing the watershed layer to add new feature
Public Sub SplitWatershed()

  'Loop through the edit tasks checking each one's name
  Dim pEditTask As IEditTask
  Dim TaskCount As Integer
  For TaskCount = 0 To g_pEditor.TaskCount - 1
    Set pEditTask = g_pEditor.Task(TaskCount)
    If pEditTask.name = "Cut Polygon Features" Then
      Set g_pEditor.CurrentTask = pEditTask
    End If
  Next TaskCount
   
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass
     
        
End Sub

Public Sub DeleteWatershed()

  'Make the topology tool active
  'Get the topology extension
  Dim pTopologyExtension As ITopologyExtension
  Dim pUID As UID
  Set pUID = New UID
  pUID.value = "esriEditorExt.topologyextension"
  Set pTopologyExtension = gApplication.FindExtensionByCLSID(pUID)
  If pTopologyExtension Is Nothing Then Exit Sub
    
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass
     
  'Set the map topology active
  Set gMapTopology = pTopologyExtension.MapTopology
  gMapTopology.ClearClasses
  gMapTopology.AddClass pFeatureclass
  
End Sub


'Start editing the watershed edges
Public Sub CommitChangesToEditing()
  Dim boolCommit
  boolCommit = MsgBox("Do you want to save the editing ?", vbYesNo, "Delete Watershed")
  If (boolCommit = vbYes) Then
    'Find all features that intersect this feature and remove the intersecting part from existing features
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    If (pFeatureLayer Is Nothing) Then
        Exit Sub
    End If
    'Save feature workspace, dataset
    Dim pDataset As IDataset
    Set pDataset = pFeatureLayer
    Dim pWorkspaceEdit As IWorkspaceEdit
    Set pWorkspaceEdit = pDataset.Workspace
    
    'Iterate over all selected features and remove the overlapping ones
    Dim pEnumFeature As IEnumFeature
    Set pEnumFeature = gMap.FeatureSelection
    pEnumFeature.Reset
    Dim pFeature As IFeature
    Set pFeature = pEnumFeature.Next
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pSpatialFilter As ISpatialFilter
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature1 As IFeature
    Dim pOutputGeom As IGeometry
    Dim pTopoOperator As ITopologicalOperator
    
    Do While Not pFeature Is Nothing
        'Find overlapping features to the selected feature
        Set pSpatialFilter = New SpatialFilter
        Set pSpatialFilter.Geometry = pFeature.Shape
        pSpatialFilter.SpatialRel = esriSpatialRelOverlaps
        Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, True)
        Set pFeature1 = pFeatureCursor.NextFeature
        Do While Not pFeature1 Is Nothing
            Set pTopoOperator = pFeature1.Shape
            Set pOutputGeom = pTopoOperator.Difference(pFeature.Shape)
            pWorkspaceEdit.StartEditOperation
            Set pFeature1.Shape = pOutputGeom
            pFeature1.Store
            pWorkspaceEdit.StopEditOperation
            Set pFeature1 = pFeatureCursor.NextFeature
            Set pTopoOperator = Nothing
            Set pOutputGeom = Nothing
        Loop
        pWorkspaceEdit.StopEditing True
        'Clean memory
        Set pSpatialFilter = Nothing
        Set pFeatureCursor = Nothing
        Set pFeature1 = Nothing
        'Find next selected feature
        Set pFeature = pEnumFeature.Next
    Loop
            
    g_pEditor.StopEditing True  'Save changes done by editing
    RenumberWatershedFeatures
  Else
    g_pEditor.StopEditing False 'Discard changes done by editing
  End If
        
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  pActiveView.PartialRefresh esriDPGeography, pFeatureLayer, Nothing
  
End Sub


Public Sub MergeSelectedFeatures(pEnvelope As IEnvelope)
On Error GoTo ShowError
  'Get the edit_watershed feature class
  Dim pFeatureLayer As IFeatureLayer
  Set pFeatureLayer = GetInputFeatureLayer("Watershed")
  If (pFeatureLayer Is Nothing) Then
    Exit Sub
  End If
  
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pFeatureLayer.FeatureClass
  
  Dim pSpatialFilter As ISpatialFilter
  Set pSpatialFilter = New SpatialFilter
  pSpatialFilter.SpatialRel = esriSpatialRelIntersects
  Set pSpatialFilter.Geometry = pEnvelope
  
  Dim pSelectionCount As Integer
  pSelectionCount = pFeatureclass.FeatureCount(pSpatialFilter)
  
  If (pSelectionCount < 2) Then
     MsgBox "Please select atleast 2 features to merge.", vbExclamation
     Exit Sub
  End If
  
  Dim pFeatureCursor As IFeatureCursor
  Set pFeatureCursor = pFeatureclass.Update(pSpatialFilter, False)
  
  'Get a collection of all selected geometries
  Dim pFeature As IFeature
  Set pFeature = pFeatureCursor.NextFeature
  
  Dim pGeomCollection As IGeometryCollection
  Set pGeomCollection = New GeometryBag

  'Delete all selected features
  Call StartEditingFeatureLayer("Watershed")
  'Delete each selected feature
  Do While Not pFeature Is Nothing
    pGeomCollection.AddGeometry pFeature.ShapeCopy
    pFeatureCursor.DeleteFeature
    Set pFeature = pFeatureCursor.NextFeature
  Loop

  'Stop editing watershed feature layer
  Call StopEditingFeatureLayer
 
  'Release memory for variables
  Set pFeature = Nothing
  Set pFeatureCursor = Nothing
  Set pFeatureclass = Nothing
  
  'Get the feature class variable again
  Set pFeatureclass = pFeatureLayer.FeatureClass
 
   'start editing watershed feature layer again
  Call StartEditingFeatureLayer("Watershed")
  
  'Get the first geometry to start editing
  Dim pUnionGeometry As IGeometry
  Set pUnionGeometry = pGeomCollection.Geometry(0)

  'Define topological operator
  Dim pTopoOperator As ITopologicalOperator
  
  'Merge all other features
  Dim iT As Integer
  For iT = 1 To (pGeomCollection.GeometryCount - 1)
      Set pTopoOperator = pUnionGeometry
      Set pUnionGeometry = pTopoOperator.Union(pGeomCollection.Geometry(iT))
  Next
  
  'Create a new feature and update the merged shape
  Dim pMergedFeature As IFeature
  Set pMergedFeature = pFeatureclass.CreateFeature
  Set pMergedFeature.Shape = pUnionGeometry
  'simplify the shape and save it
  Dim pSimplifyFeature As IFeatureSimplify
  Set pSimplifyFeature = pMergedFeature
  pSimplifyFeature.SimplifyGeometry pUnionGeometry
  pMergedFeature.Store

  'Stop editing feature layer
  StopEditingFeatureLayer
  
GoTo CleanUp
ShowError:
    MsgBox "MergeSelectedFeatures: " & Err.description & vbTab & Err.Number
CleanUp:
  Set pFeatureLayer = Nothing
  Set pFeatureclass = Nothing
  Set pFeatureCursor = Nothing
  Set pFeature = Nothing
  Set pGeomCollection = Nothing
  Set pTopoOperator = Nothing
  Set pUnionGeometry = Nothing
  Set pSimplifyFeature = Nothing
  Set pMergedFeature = Nothing
End Sub


Public Function CreateFeatureClassForPolygonShapeFile(DirName As String, FileName As String) As IFeatureClass

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
  
  Dim pFieldBMPID As esriGeoDatabase.IField
  Dim pFieldEditBMPID As IFieldEdit
  Set pFieldBMPID = New esriGeoDatabase.Field
  Set pFieldEditBMPID = pFieldBMPID
  pFieldEditBMPID.name = "BMPID"
  pFieldEditBMPID.Type = esriFieldTypeInteger
  pFieldEditBMPID.IsNullable = True
  
  Dim pFieldAreaM As esriGeoDatabase.IField
  Dim pFieldEditAreaM As IFieldEdit
  Set pFieldAreaM = New esriGeoDatabase.Field
  Set pFieldEditAreaM = pFieldAreaM
  pFieldEditAreaM.name = "Area_SQM"
  pFieldEditAreaM.Type = esriFieldTypeDouble
  pFieldEditAreaM.IsNullable = True
  
  Dim pFieldAreaA As esriGeoDatabase.IField
  Dim pFieldEditAreaA As IFieldEdit
  Set pFieldAreaA = New esriGeoDatabase.Field
  Set pFieldEditAreaA = pFieldAreaA
  pFieldEditAreaA.name = "Area_Acre"
  pFieldEditAreaA.Type = esriFieldTypeDouble
  pFieldEditAreaA.IsNullable = True
  
  Dim pRasterDEMProps As IRasterAnalysisProps
  If Not GetInputRasterLayer("DEM") Is Nothing Then
    Set pRasterDEMProps = GetInputRasterLayer("DEM").Raster
  Else
    Set pRasterDEMProps = Nothing
  End If
  
  Dim pRasterLUProps As IRasterAnalysisProps
  Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster

  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = esriGeometryPolygon
    If Not GetInputRasterLayer("DEM") Is Nothing Then
        Set .SpatialReference = pRasterDEMProps.AnalysisExtent.SpatialReference
    Else
        Set .SpatialReference = pRasterLUProps.AnalysisExtent.SpatialReference
    End If
  End With
  Set pFieldEditShape.GeometryDef = pGeomDef
  pFieldsEdit.AddField pFieldShape
  pFieldsEdit.AddField pFieldID
  pFieldsEdit.AddField pFieldBMPID
  pFieldsEdit.AddField pFieldAreaM
  pFieldsEdit.AddField pFieldAreaA
  
  ' Create the shapefile some parameters apply to geodatabase options and can be defaulted as Nothing
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(pFileName, pFields, Nothing, Nothing, esriFTSimple, strShapeFieldName, "")

  ' Return the value
  Set CreateFeatureClassForPolygonShapeFile = pFeatClass
    
  GoTo CleanUp
ShowError:
    MsgBox "CreateFeatureClassForPolygonShapeFile: " & Err.description
CleanUp:
    Set pRasterLUProps = Nothing
    Set pRasterDEMProps = Nothing
End Function


'Subroutine to create new watershed feature
Public Sub CreateWatershedFeature(pGeom As IPolygon)
  On Error GoTo ErrorHandler

  If (Not pGeom Is Nothing) Then
    ' We have a valid geometry so we must create a feature and give it the geometry
   
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    Dim pFeatureclass As IFeatureClass
    If (pFeatureLayer Is Nothing) Then  'Create a new polygon layer called Watershed
        Set pFeatureclass = CreateFeatureClassForPolygonShapeFile(gMapTempFolder, "watershed")
        Set pFeatureLayer = New FeatureLayer
        Set pFeatureLayer.FeatureClass = pFeatureclass
        AddLayerToMap pFeatureLayer, "Watershed"
        
        'This will make the Watershed as dictionary name
        If gLayerNameDictionary.Exists("Watershed") Then gLayerNameDictionary.Remove "Watershed"
        gLayerNameDictionary.Item("Watershed") = "Watershed"
    Else
        Set pFeatureclass = pFeatureLayer.FeatureClass
    End If
 
    'Find all features that intersect this feature and remove the intersecting part from existing features
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pSpatialFilter As ISpatialFilter
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature1 As IFeature
    Dim pOutputGeom As IGeometry
    Dim pTopoOperator As ITopologicalOperator2
   
    'If the new geometry contains any existing geometry, clip the smaller one from larger one
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pGeom
    pSpatialFilter.SpatialRel = esriSpatialRelContains
    Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, False)
    Set pFeature1 = pFeatureCursor.NextFeature
    Do While Not pFeature1 Is Nothing
            Set pTopoOperator = pGeom
            Set pGeom = pTopoOperator.Difference(pFeature1.Shape)
        Set pFeature1 = pFeatureCursor.NextFeature
    Loop
    Set pFeatureCursor = Nothing
    Set pFeature1 = Nothing
    Set pTopoOperator = Nothing
    
    'Now just check if the geometry has any intersecting parts
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pGeom
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, False)
    Set pFeature1 = pFeatureCursor.NextFeature

    Do While Not pFeature1 Is Nothing
            Set pTopoOperator = pFeature1.Shape
            Set pOutputGeom = pTopoOperator.Difference(pGeom)
            Set pTopoOperator = pOutputGeom
            pTopoOperator.IsKnownSimple = False
            pTopoOperator.Simplify  'Simplify the geometry
            Set pFeature1.Shape = pOutputGeom
            pFeature1.Store
          
            Set pOutputGeom = Nothing
            Set pTopoOperator = Nothing
        Set pFeature1 = pFeatureCursor.NextFeature
    Loop
    
    'Add the new feature
    Dim pFeature As IFeature
    Set pFeature = pFeatureclass.CreateFeature
    Set pTopoOperator = pGeom
    pTopoOperator.IsKnownSimple = False
    pTopoOperator.Simplify
    Set pFeature.Shape = pGeom
    'pFeature.value(pFeatureClass.FindField("ID")) = pFeature.OID + 1
    pFeature.Store
    
    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    pActiveView.Refresh

    'Call the subroutine to render the watershed layer
    RenderWatershedLayer pFeatureLayer
    
  End If

  GoTo CleanUp
SubTypeError:
  Err.Clear
  Resume Next
  Exit Sub
ErrorHandler:
  MsgBox "CreateWatershedFeature: " & Err.description
CleanUp:

   Set pGeom = Nothing
   Set pFeatureLayer = Nothing
   Set pFeatureclass = Nothing
   'Set pDataset = Nothing
   'Set pWorkspaceEdit = Nothing
   Set pFeature = Nothing
   'Set pRowSubTypes = Nothing
   'Set pSimplifyFeature = Nothing
   Set pSpatialFilter = Nothing
   Set pFeatureCursor = Nothing
   Set pFeature1 = Nothing
   'Set pIntersectGeom = Nothing
   Set pOutputGeom = Nothing
   Set pTopoOperator = Nothing
   Set pActiveView = Nothing

End Sub




'Subroutine to split watershed feature
Public Sub SplitWatershedFeature(pPolyline As IPolyline)
On Error GoTo ErrorHandler:

  If (Not pPolyline Is Nothing) Then
    ' We have a valid geometry so we must create a feature and give it the geometry
   
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    Dim pFeatureclass As IFeatureClass
    If (pFeatureLayer Is Nothing) Then  'Create a new polygon layer called Watershed
       Exit Sub
    End If
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pDataset As IDataset
    Set pDataset = pFeatureclass
    Dim pWorkspaceEdit As IWorkspaceEdit
    Set pWorkspaceEdit = pDataset.Workspace
    pWorkspaceEdit.StartEditing False
   
    '** project the polyline in same projection as feature layer
    Dim pGeoDataset As IGeoDataset
    Set pGeoDataset = pFeatureLayer
    pPolyline.Project pGeoDataset.SpatialReference
    
    '** find all features that intersect this feature and remove the intersecting part from existing features
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pPolyline
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Dim pFeatureCount As Integer
    
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureclass.Update(pSpatialFilter, False)
    Dim pFeature1 As IFeature
    Set pFeature1 = pFeatureCursor.NextFeature
    Dim pNewFeature As IFeature
    Dim pLeftPoly As IPolygon
    Dim pRightPoly As IPolygon
    Dim pCutPoly As IPolygon
    Dim pTopoOperator As ITopologicalOperator
    Dim pTopoOperator2 As ITopologicalOperator2
        
    Set pTopoOperator2 = pPolyline
    pTopoOperator2.Simplify
    pTopoOperator2.IsKnownSimple = True
    Set pTopoOperator2 = Nothing
    
    Dim pRelationalOp As IRelationalOperator
    Dim bFromPtContains As Boolean
    Dim bToPtContains As Boolean
    'Check which feature has the line and cut it
    Do While Not pFeature1 Is Nothing
            Set pCutPoly = Nothing
            Set pCutPoly = pFeature1.ShapeCopy
            Set pTopoOperator = Nothing
            Set pTopoOperator = pCutPoly
            pTopoOperator.Simplify
            
            '** check if both ends of the line are within the polygon, then don't use it to cut it
            Set pRelationalOp = pCutPoly
            bFromPtContains = pRelationalOp.Contains(pPolyline.FromPoint)
            bToPtContains = pRelationalOp.Contains(pPolyline.ToPoint)
            
            If (bToPtContains = False Or bToPtContains = False) Then
            
                '** split the polygon feature into left and right
                Set pLeftPoly = Nothing
                Set pRightPoly = Nothing
On Error GoTo MoveToFeature
                    pTopoOperator.Cut pPolyline, pLeftPoly, pRightPoly
                    If ((Not pLeftPoly.IsEmpty) And (Not pRightPoly.IsEmpty)) Then
                        pWorkspaceEdit.StartEditOperation
                        Set pFeature1.Shape = pLeftPoly
                        pFeature1.value(pFeatureclass.FindField("ID")) = pFeature1.OID + 1
                        pFeature1.Store
                        'Create a new feature to store the new geometry
                        Set pNewFeature = Nothing
                        Set pNewFeature = pFeatureclass.CreateFeature
                        Set pNewFeature.Shape = pRightPoly
                        pNewFeature.value(pFeatureclass.FindField("ID")) = pNewFeature.OID + 1
                        pNewFeature.Store
                        pWorkspaceEdit.StopEditOperation
                    End If
            End If
MoveToFeature:
            Set pFeature1 = pFeatureCursor.NextFeature
    Loop
    
    GoTo CleanUp
  End If

  Exit Sub
ErrorHandler:
  MsgBox "SplitWatershedFeature: " & Err.description & " " & Err.Number
CleanUp:
If Not (pPolyline Is Nothing) Then
    pWorkspaceEdit.StopEditing True
    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    pActiveView.PartialRefresh esriViewGeography, pFeatureLayer, Nothing
    'Call the subroutine to render the watershed layer
    RenderWatershedLayer pFeatureLayer
End If

    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pDataset = Nothing
    Set pWorkspaceEdit = Nothing
    Set pGeoDataset = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature1 = Nothing
    Set pNewFeature = Nothing
    Set pLeftPoly = Nothing
    Set pRightPoly = Nothing
    Set pCutPoly = Nothing
    Set pTopoOperator = Nothing
    Set pTopoOperator2 = Nothing
End Sub


Public Sub CreateWatershedForEditing()
On Error GoTo ErrorHandler
    
    'Delete Watershed feature layer if present
    Dim pOutputFeatureLayer As IFeatureLayer
    Set pOutputFeatureLayer = GetInputFeatureLayer("Watershed")
    If Not (pOutputFeatureLayer Is Nothing) Then
        DeleteLayerFromMap "Watershed"
    End If
    
    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = GetInputRasterLayer("SubWatershed")
    If (pRasterLayer Is Nothing) Then
        Exit Sub
    End If
    Dim pRaster As IRaster
    Set pRaster = pRasterLayer.Raster
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = ConvertRasterToFeature(pRaster, "Value", "watershed.shp", "Polygon")
    Dim pInputDataset As IDataset
    Set pInputDataset = pFeatureLayer
     
    Dim pInputTable As iTable
    Set pInputTable = pFeatureLayer
        
    ' Get the feature class properties needed for the output
    Dim pInputFeatureCLass As IFeatureClass
    Set pInputFeatureCLass = pFeatureLayer.FeatureClass
    Dim pFeatureClassName As IFeatureClassName
    Set pFeatureClassName = New FeatureClassName
    With pFeatureClassName
        .FeatureType = esriFTSimple
        .ShapeFieldName = "Shape"
        .ShapeType = pFeatureLayer.FeatureClass.ShapeType
    End With
    
  ' Set output location and output feature class name
  Dim pNewWSName As IWorkspaceName
  Set pNewWSName = New WorkspaceName
  pNewWSName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
  pNewWSName.PathName = gMapTempFolder
  Dim pDatasetName As IDatasetName
  Set pDatasetName = pFeatureClassName
  Dim pFileName As String
  pFileName = CreateUniqueTableName(gMapTempFolder, "editwtrshed")
  pDatasetName.name = pFileName
  Set pDatasetName.WorkspaceName = pNewWSName
  
  ' Perform the dissolve.
  ' Since we are performing a spatial dissolve, we must use the operation code
  ' Dissolve on the Shape field.
  Dim iBGP As IBasicGeoprocessor
  Set iBGP = New BasicGeoprocessor
  Dim pOutputTable As iTable
  Set pOutputTable = iBGP.Dissolve(pInputTable, False, "ID", "Dissolve.Shape, Minimum.ID", pDatasetName)

  ' Add the output to the map
  Dim pOutputFeatureClass As IFeatureClass
  Set pOutputFeatureClass = pOutputTable
    
  ' Error checking
  If pOutputFeatureClass Is Nothing Then
      MsgBox "FeatureClass QI Failed"
      Exit Sub
  End If
  
  'Add layer to map
  Set pOutputFeatureLayer = New FeatureLayer
  Set pOutputFeatureLayer.FeatureClass = pOutputFeatureClass
  DeleteLayerFromMap "Watershed" ' Arun Raj
  AddLayerToMap pOutputFeatureLayer, "Watershed"
  RenderWatershedLayer pOutputFeatureLayer
  
 
  'Add ID;s and area
  Call RenumberWatershedFeatures
  Exit Sub
  
ErrorHandler:
    MsgBox "CreateWatershedForEditing: " & Err.description
End Sub


Public Sub RenderWatershedLayer(pWatershedLayer As IFeatureLayer)
On Error GoTo ErrorHandler

  'Check if ID field is present
  Dim pFeatureclass As IFeatureClass
  Set pFeatureclass = pWatershedLayer.FeatureClass
    
  Dim iIDFld As Long
  iIDFld = pFeatureclass.FindField("ID")
  If (iIDFld < 0) Then
    'Add ID field
    Dim pFieldEditID As IFieldEdit
    Set pFieldEditID = New esriGeoDatabase.Field
    pFieldEditID.name = "ID"
    pFieldEditID.Type = esriFieldTypeInteger
    pFieldEditID.IsNullable = True
    pFeatureclass.AddField pFieldEditID
    iIDFld = pFeatureclass.FindField("ID")
  End If


  Dim pRender As ISimpleRenderer
  Set pRender = New SimpleRenderer
  
  '** These properties should be set prior to adding values
  pRender.Label = "ID"
  
  '** Make the color ramp we will use for the symbols in the renderer
  Dim pRGBColor As IRgbColor
  Set pRGBColor = New RgbColor
  pRGBColor.Red = 0
  pRGBColor.Green = 169
  pRGBColor.Blue = 240
  Dim pSymbol As ISimpleFillSymbol
  Set pSymbol = New SimpleFillSymbol
  pSymbol.Style = esriSFSSolid
  Dim pLineSymbol As ILineSymbol
  Set pLineSymbol = New SimpleLineSymbol
  pLineSymbol.Color = pRGBColor
  pLineSymbol.Width = 1
  pSymbol.Outline = pLineSymbol

  Set pRGBColor = New RgbColor
  'Set as single blue color
  pRGBColor.Red = 202
  pRGBColor.Green = 254
  pRGBColor.Blue = 255
  pRGBColor.Transparency = 120  'Make it 50% transparent

  pSymbol.Color = pRGBColor
  Set pRender.Symbol = pSymbol
  
  Dim pLyr As IGeoFeatureLayer
  Set pLyr = pWatershedLayer
  Set pLyr.Renderer = pRender
  pLyr.DisplayField = "ID"
  
  Dim pLayerEffects As ILayerEffects
  Set pLayerEffects = pWatershedLayer
  pLayerEffects.Transparency = 20
    

    ' setup LabelEngineProperties for the FeatureLayer
    ' get the AnnotateLayerPropertiesCollection for the FeatureLayer
    Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
    Set pAnnoLayerPropsColl = pLyr.AnnotationProperties
    pLyr.DisplayAnnotation = True
    pAnnoLayerPropsColl.Clear
    ' create a new LabelEngineLayerProperties object
    Dim aLELayerProps As ILabelEngineLayerProperties
    Set aLELayerProps = New LabelEngineLayerProperties
    aLELayerProps.IsExpressionSimple = True
    aLELayerProps.Expression = "[ID]"
    Dim pTextSymbol As ITextSymbol
    Set pTextSymbol = New TextSymbol
    pTextSymbol.Size = 8
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    pColor.RGB = RGB(0, 169, 240)
    pTextSymbol.Color = pColor
    Set aLELayerProps.Symbol = pTextSymbol
    ' assign it to the layer's AnnotateLayerPropertiesCollection
    pAnnoLayerPropsColl.add aLELayerProps
    'get the BasicOverposterLayerProperties
    Dim pBasicOverposterLayerProps As IBasicOverposterLayerProperties
    Set pBasicOverposterLayerProps = aLELayerProps.BasicOverposterLayerProperties
    pBasicOverposterLayerProps.NumLabelsOption = esriOneLabelPerShape
          
  '** Refresh the TOC
  pWatershedLayer.Visible = True
  gMxDoc.ActiveView.ContentsChanged
  gMxDoc.UpdateContents

  '** Draw the map
  gMxDoc.ActiveView.Refresh
  
  
  Exit Sub
ErrorHandler:
    MsgBox "RenderWatershedLayer: " & Err.description
End Sub


'** Mira Chokshi - 11/05/2004 **'
'** Find Watershed feature layer, if not found, try find SubWatershed raster layer
Public Function FindAndConvertWatershedFeatureLayerToRaster() As Boolean
On Error GoTo ErrorHandler

    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    Dim pRasterLayer As IRasterLayer
    If (pFeatureLayer Is Nothing) Then
            MsgBox "Delineated Watersheds required. Use Delineation methods to create, modify Watershed layer. "
            FindAndConvertWatershedFeatureLayerToRaster = False
    Else
        'Feature layer found, delete existing SubWatershed layer
        DeleteLayerFromMap ("SubWatershed")
        
        'Set ID values of Watershed feature layer = FID + 1
        Dim pFeatureclass As IFeatureClass
        Set pFeatureclass = pFeatureLayer.FeatureClass
''        Dim pCalculator As ICalculator
''        Set pCalculator = New Calculator
''        Dim pCursor As ICursor
''        Set pCursor = pFeatureclass.Update(Nothing, True)
''        With pCalculator
''            Set .Cursor = pCursor
''            .Expression = "[FID] + 1"
''            .Field = "ID"
''        End With
''        pCalculator.Calculate
        
        Dim iIDField As Integer
        iIDField = pFeatureclass.FindField("ID")
        
        If iIDField < 0 Then
            MsgBox "Missing ID field in watershed feature layer. Use Delineation methods to create, modify Watershed layer. "
            FindAndConvertWatershedFeatureLayerToRaster = False
            Exit Function
        End If
        
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "ID >= 0"
                
        Dim pTableSort As ITableSort
        Set pTableSort = New esriGeoDatabase.TableSort
    
        With pTableSort
          .Fields = "ID"
          .Ascending("ID") = True
          Set .QueryFilter = pQueryFilter
          Set .Table = pFeatureclass
        End With
        
        pTableSort.Sort Nothing
           
        Dim pRow As iRow
        
        Dim pCursor As ICursor
        Set pCursor = pTableSort.Rows
        Set pRow = pCursor.NextRow
        
        Dim curID As Integer
        If Not pRow Is Nothing Then
            curID = pRow.value(iIDField)
            Set pRow = pCursor.NextRow
        End If
        Do Until pRow Is Nothing
            If pRow.value(iIDField) <> curID + 1 Then
                MsgBox "Watershed ID values are not in sequence. Use Delineation methods to create, modify Watershed layer. "
                FindAndConvertWatershedFeatureLayerToRaster = False
                Exit Function
            Else
                curID = pRow.value(iIDField)
            End If
            Set pRow = pCursor.NextRow
        Loop
    
        'Convert Watershed feature layer to raster layer
        Dim pRasterDataset As IRasterDataset
        Set pRasterDataset = ConvertFeatureToRaster(pFeatureclass, "ID", "subwatershd", Nothing)
        Dim pRaster As IRaster
        gAlgebraOp.BindRaster pRasterDataset, "SUB"
        Set pRaster = gAlgebraOp.Execute("Int([SUB])")
        gAlgebraOp.UnbindRaster "SUB"
        Set pRasterLayer = New RasterLayer
        pRasterLayer.CreateFromRaster pRaster
        AddLayerToMap pRasterLayer, "SubWatershed"
        pRasterLayer.Visible = False
        FindAndConvertWatershedFeatureLayerToRaster = True
    End If
    GoTo CleanUp

ErrorHandler:
    MsgBox "FindAndConvertWatershedFeatureLayerToRaster: " & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pRasterLayer = Nothing
    Set pFeatureclass = Nothing
    'Set pCalculator = Nothing
    Set pCursor = Nothing
    Set pRasterDataset = Nothing
    Set pRaster = Nothing
End Function


Public Sub RenumberWatershedFeatures()
On Error GoTo ShowError
'     Not sure why this is done like this - Sabu Paul - Dec 4, 2008
''    If gLayerNameDictionary Is Nothing Then Set gLayerNameDictionary = New Scripting.Dictionary
''    If gLayerNameDictionary.Exists("Watershed") Then gLayerNameDictionary.Remove "Watershed"
''    gLayerNameDictionary.Item("Watershed") = "Watershed"

    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Watershed")
    If pFeatureLayer Is Nothing Then Exit Sub
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    If pFeatureclass Is Nothing Then Exit Sub
    
    Dim iIDFld As Long
    iIDFld = pFeatureclass.FindField("ID")
    Dim iAreaSQMFld As Long
    iAreaSQMFld = pFeatureclass.FindField("Area_SQM")
    If (iAreaSQMFld < 0) Then   'Create the field
        Dim pFieldEditAreaSQM As IFieldEdit
        Set pFieldEditAreaSQM = New esriGeoDatabase.Field
        pFieldEditAreaSQM.name = "Area_SQM"
        pFieldEditAreaSQM.Type = esriFieldTypeDouble
        pFieldEditAreaSQM.IsNullable = True
        pFeatureclass.AddField pFieldEditAreaSQM
        iAreaSQMFld = pFeatureclass.FindField("Area_SQM")
    End If
    Dim iAreaSQAcreFld As Long
    iAreaSQAcreFld = pFeatureclass.FindField("Area_Acre")
    If (iAreaSQAcreFld < 0) Then   'Create the field
        Dim pFieldEditAreaSQAcre As IFieldEdit
        Set pFieldEditAreaSQAcre = New esriGeoDatabase.Field
        pFieldEditAreaSQAcre.name = "Area_Acre"
        pFieldEditAreaSQAcre.Type = esriFieldTypeDouble
        pFieldEditAreaSQAcre.IsNullable = True
        pFeatureclass.AddField pFieldEditAreaSQAcre
        iAreaSQAcreFld = pFeatureclass.FindField("Area_Acre")
    End If
    
    Dim pSqUnitArea As Double
    Dim pSQMeterFactor As Double
    pSQMeterFactor = gMetersPerUnit * gMetersPerUnit    'Per sq. unit area converted to sq. meter
    Dim pSQAcreFactor As Double
    pSQAcreFactor = pSQMeterFactor * 0.0002471044       'sq meter to acre conversion
        
    'Iterate over the entire feature class and set their ID's
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Do While Not pFeature Is Nothing
        pFeature.value(iIDFld) = pFeature.OID + 1
        pSqUnitArea = ReturnPolygonArea(pFeature)
        pFeature.value(iAreaSQMFld) = FormatNumber(pSqUnitArea * pSQMeterFactor, "#.##")
        pFeature.value(iAreaSQAcreFld) = FormatNumber(pSqUnitArea * pSQAcreFactor, "#.####")
        pFeature.Store
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    Set pFeature = Nothing
    Set pFeatureCursor = Nothing
    Set pFeatureclass = Nothing
    Set pFeatureLayer = Nothing
    gMxDoc.ActiveView.Refresh
    
    Exit Sub
ShowError:
    MsgBox "RenumberWatershedFeatures: " & Err.description
End Sub

'*** Helper function to format a number
Private Function FormatNumber(pValue As Double, pFormat As String) As Double
On Error GoTo ShowError

    Dim pFormattedString As String
    pFormattedString = Format(CStr(pValue), pFormat)
    
    Dim pNumber As Double
    If (IsNumeric(pFormattedString)) Then
        pNumber = CDbl(pFormattedString)
    Else
        pNumber = 0
    End If
    'Return the formatted number
    FormatNumber = pNumber
    Exit Function
    
ShowError:
    MsgBox "FormatNumber: " & Err.description
    

End Function
