Attribute VB_Name = "ModuleSchematics"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleSchematics
'   Purpose:     Add BMPs on the map, open corresponding bmp dialog,
'                and prompt user to enter bmp parameters.
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  08/26/2004 - Mira Chokshi
'
'******************************************************************************
Option Explicit
Option Base 0

Public gToggleLayer As String
Public gConduitIDValue As Integer
Private pDictionary As Scripting.Dictionary
Private pRouteDictionary As Scripting.Dictionary
Private pSchemClass As IFeatureClass
Private pSchemRouteFClass As IFeatureClass
Private pTotalOutlets As Integer
Private pMaxHorizontal As Double
Private pUpstreamBMPPresent As Scripting.Dictionary

'*******************************************************************************
'Subroutine : ToggleSchematicLayer
'Purpose    : Creates a new point feature class for bmp for schematic view
'Arguments  : Destination directory, name of the feature class file
'Author     : Mira Chokshi
'*******************************************************************************
Public Sub ToggleSchematicLayer()
On Error GoTo ShowError
    'Get the Active View
    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    
    'Get the schematic bmps feature layer
    Dim pSchematicBMPLayer As IFeatureLayer
    Set pSchematicBMPLayer = GetInputFeatureLayer("Schematic BMPs")
    If (pSchematicBMPLayer Is Nothing) Then
        Exit Sub
    End If

    'Get the BMP Feature layer
    Dim pBMPLayer As IFeatureLayer
    Set pBMPLayer = GetInputFeatureLayer("BMPs")
    Dim pLayer As ILayer
    
    
    'If toggle layer variable not defined, set it to DEM
    If (gToggleLayer = "") Then
        gToggleLayer = "BMPs"
    End If
    
    'If current layer focussed is DEM, set it to schematic bmps & vice versa
    If (gToggleLayer = "BMPs") Then
        Set pLayer = pSchematicBMPLayer
        gToggleLayer = "Schematic BMPs"
    Else
        Set pLayer = pBMPLayer
        gToggleLayer = "BMPs"
    End If
    
    'Get current view and refresh it
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pLayer.AreaOfInterest
    pEnvelope.Expand 1.5, 1.5, True
    pActiveView.Extent = pEnvelope
    pActiveView.Refresh

    
GoTo CleanUp
ShowError:
    MsgBox "ToggleSchematicLayer: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pSchematicBMPLayer = Nothing
    Set pBMPLayer = Nothing
    Set pLayer = Nothing
    Set pEnvelope = Nothing
End Sub


'*******************************************************************************
'Subroutine : ToggleBMPIconPointView
'Purpose    : Renders the BMP Feature layer with Icon view or point view
'Arguments  : none
'Author     : Mira Chokshi
'*******************************************************************************
Public Sub ToggleBMPIconPointView()
On Error GoTo ShowError
    
    'If current layer focussed is Schematic BMPs, toggle it to BMPs layer
    If (gToggleLayer = "") Then
        gToggleLayer = "BMPs"
    End If
    
    'Get the BMP Feature layer
    Dim pBMPLayer As IFeatureLayer
    Set pBMPLayer = GetInputFeatureLayer(gToggleLayer)
    If (pBMPLayer Is Nothing) Then
        Exit Sub
    End If
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pBMPLayer.FeatureClass
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    If (pFeature Is Nothing) Then
        Exit Sub
    End If
    
    Dim pLyr As IGeoFeatureLayer
    Set pLyr = pBMPLayer
    
    Dim pFeatureRenderer As IFeatureRenderer
    Set pFeatureRenderer = pLyr.Renderer
    
    Dim pSymbol As ISymbol
    Set pSymbol = pFeatureRenderer.SymbolByFeature(pFeature)
    If (TypeOf pSymbol Is IMultiLayerMarkerSymbol) Then
        RenderPointViewBMPLayer pBMPLayer
    ElseIf (TypeOf pSymbol Is ISimpleMarkerSymbol) Then
        RenderSchematicBMPLayer pBMPLayer
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "ToggleBMPIconPointView: " & Err.description
CleanUp:
    Set pBMPLayer = Nothing
    Set pFeatureclass = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pLyr = Nothing
    Set pFeatureRenderer = Nothing
    Set pSymbol = Nothing
End Sub


'*******************************************************************************
'Subroutine : CreateSchematicsForBMPs
'Purpose    : Creates a new point feature class for bmp for schematic view
'Arguments  : Destination directory, name of the feature class file
'Author     : Mira Chokshi
'History    : 11/26/2008 added VFS as elements of the schematic
'*******************************************************************************
Public Function CreateSchematicsForBMPs()
On Error GoTo ShowError
        
    'Delete schematic layers from map
    DeleteLayerFromMap "Schematic BMPs"
    DeleteLayerFromMap "Schematic Route"
    
    'Define a collection: it works like a stack
    'Very efficient in storing the routing information
    Dim pCollection As Collection
    Set pCollection = New Collection
    
    'Define a dictionary to hold the point location
    Set pDictionary = CreateObject("Scripting.Dictionary")
    Set pRouteDictionary = CreateObject("Scripting.Dictionary")
    Set pUpstreamBMPPresent = CreateObject("Scripting.Dictionary")
        
    'Create a feature class for the new schematic layer
    Set pSchemClass = Nothing
    Set pSchemClass = CreatePointFeatureClassForSchematics(gMapTempFolder, "schembmp")
    Dim pSchemBMPLayer As IFeatureLayer
    Set pSchemBMPLayer = New FeatureLayer
    Set pSchemBMPLayer.FeatureClass = pSchemClass
    
    'Create a feature class for the new schematic route layer
    Dim pSchemRouteLayer As IFeatureLayer
    Set pSchemRouteFClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, "schemroute")
    Set pSchemRouteLayer = New FeatureLayer
    Set pSchemRouteLayer.FeatureClass = pSchemRouteFClass
    
    'for BMP
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")     'replace it with BMPs
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    Dim iID As Long
    iID = pBMPFClass.FindField("ID")
    Dim iDSID As Long
    iDSID = pBMPFClass.FindField("DSID")
    Dim pID As Integer
    Dim pDSID As Integer
    Dim pType As String
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    'for VFS
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    Dim iVFSID As Long
    Dim iVFSDSID As Long
    Dim pVFSID As Integer
    Dim pVFSDSID As Integer
    Dim pVFSFeatureCursor As IFeatureCursor
    Dim pVFSFeature As IFeature
    
    '** Iterate over bmp feature class to find if each bmp has any upstream BMP's
    Set pFeatureCursor = pBMPFClass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    Do While Not pFeature Is Nothing
        pUpstreamBMPPresent.Item(pFeature.value(iDSID)) = True
        Set pFeature = pFeatureCursor.NextFeature
    Loop
 
    '** Iterate over VFS feature class to find if each VFS has any upstream BMP
    If Not pVFSFLayer Is Nothing Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
        iVFSID = pVFSFClass.FindField("ID")
        iVFSDSID = pVFSFClass.FindField("DSID")
        Set pVFSFeatureCursor = pVFSFClass.Search(Nothing, True)
        Set pVFSFeature = pVFSFeatureCursor.NextFeature
        Do While Not pVFSFeature Is Nothing
            pUpstreamBMPPresent.Item(pVFSFeature.value(iVFSDSID)) = True
            Set pVFSFeature = pVFSFeatureCursor.NextFeature
        Loop
    End If
    
    Dim pContinue As Boolean
    pContinue = True
    pCollection.add 0
    Dim pItem As Integer
       
    'Find outlets for BMP
    pTotalOutlets = 0
    Do While pContinue
        pItem = pCollection.Item(1)
        pCollection.Remove (1)
        pQueryFilter.WhereClause = "DSID = " & pItem
        Set pFeatureCursor = pBMPFClass.Search(pQueryFilter, True)
        Set pFeature = pFeatureCursor.NextFeature
        Do While Not pFeature Is Nothing
              pID = pFeature.value(iID)
              pDSID = pFeature.value(iDSID)
              pCollection.add pID
              pRouteDictionary.Item(pID) = pDSID
              Set pFeature = pFeatureCursor.NextFeature
        Loop
        Set pFeature = Nothing
        Set pFeatureCursor = Nothing
        
        
        If Not pVFSFLayer Is Nothing Then
            'pContinue = True
            'pCollection.Add 0
            'VFS
            If Not pVFSFClass Is Nothing Then
                'pQueryFilter.WhereClause = "DSID = " & pItem
                Set pVFSFeatureCursor = pVFSFClass.Search(pQueryFilter, True)
                Set pVFSFeature = pVFSFeatureCursor.NextFeature
                Do While Not pVFSFeature Is Nothing
                    pVFSID = pVFSFeature.value(iVFSID)
                    pVFSDSID = pVFSFeature.value(iVFSDSID)
                    pCollection.add pVFSID
                    pRouteDictionary.Item(pVFSID) = pVFSDSID
                    'Call AddSchematicBMP
                    Set pVFSFeature = pVFSFeatureCursor.NextFeature
                Loop
                Set pVFSFeature = Nothing
                Set pVFSFeatureCursor = Nothing
            End If
        End If
        If (pCollection.Count = 0) Then
            pContinue = False
        Else
            Call AddSchematicBMP
        End If
    Loop
    
    'Add schematic bmps
    pSchemBMPLayer.Visible = True
    RenderSchematicBMPLayer pSchemBMPLayer
    AddLayerToMap pSchemBMPLayer, "Schematic BMPs"
    gToggleLayer = "Schematic BMPs"
    
    'call subroutine to add route
    If (AddSchematicBMPRoute(pSchemRouteFClass, pSchemClass)) Then
        pSchemRouteLayer.Visible = True
        RenderSchematicRouteLayer pSchemRouteLayer
        AddLayerToMap pSchemRouteLayer, "Schematic Route"
    End If
    
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pSchemBMPLayer.AreaOfInterest
    pEnvelope.Expand 1.2, 1.2, True
    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    pActiveView.Extent = pEnvelope
    pActiveView.Refresh
    
    GoTo CleanUp
ShowError:
    MsgBox "CreateSchematicsForBMPs: " & Err.description
CleanUp:
    Set pCollection = Nothing
    Set pSchemBMPLayer = Nothing
    Set pSchemRouteLayer = Nothing
    Set pBMPFLayer = Nothing
    Set pBMPFClass = Nothing
    Set pQueryFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pEnvelope = Nothing
    Set pActiveView = Nothing
    Set pUpstreamBMPPresent = Nothing

End Function


Public Sub ModifyRouteLayer(pRouteLayerName As String, pBMPLayerName As String)
On Error GoTo ShowError
        
    'Get total number of VFS on map
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    Dim pVFSCount As Integer
    pVFSCount = 0
    If Not (pVFSFLayer Is Nothing) Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
        pVFSCount = pVFSFClass.FeatureCount(Nothing)    ' Get vfs feature count
    End If
    
    'call subroutine to add route
    'Create a feature class for the new schematic route layer
    Dim pSchemRouteLayer As IFeatureLayer
    Set pSchemRouteLayer = GetInputFeatureLayer(pRouteLayerName)
    If (pSchemRouteLayer Is Nothing) Then
        Dim pFClassName As String
        If (pBMPLayerName = "BMPs") Then
            pFClassName = "conduits"
        Else
            pFClassName = "schemroute"
        End If
        Dim pSchemRouteFClass As IFeatureClass
        Set pSchemRouteFClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, pFClassName)
        Set pSchemRouteLayer = New FeatureLayer
        Set pSchemRouteLayer.FeatureClass = pSchemRouteFClass
        AddLayerToMap pSchemRouteLayer, pRouteLayerName
    End If
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pSchemRouteLayer.FeatureClass
    Dim iConduitID As Long
    Dim iFROM As Long
    Dim iTO As Long
    Dim iOUTTYPE As Long
    Dim iTypeDesc As Long
    iConduitID = pFeatureclass.FindField("ID")
    iFROM = pFeatureclass.FindField("CFROM")
    iTO = pFeatureclass.FindField("CTO")
    iOUTTYPE = pFeatureclass.FindField("OUTLETTYPE")
    iTypeDesc = pFeatureclass.FindField("TYPEDESC")
    
    Dim pFROM, pTO, pOutType As Integer
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    
    Dim pSchemBMPLayer As IFeatureLayer
    Set pSchemBMPLayer = GetInputFeatureLayer(pBMPLayerName)
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pSchemBMPLayer.FeatureClass
    Dim pBMPFilter As IQueryFilter
    Set pBMPFilter = New QueryFilter
    Dim pBMPCursor As IFeatureCursor
    Dim pBMPFeat As IFeature
    
    Dim pPoint As IPoint
    Dim pPolyline As IPolyline
    Dim pRouteLine As IPointCollection
    
    Dim pNetworkTable As iTable
    Set pNetworkTable = GetInputDataTable("BMPNetwork")
    If (pNetworkTable Is Nothing) Then
        MsgBox "BMPNetwork table not found."
        GoTo CleanUp
    End If
    Dim iID As Long
    Dim iDSID As Long
    Dim iIDType As Long
    iID = pNetworkTable.FindField("ID")
    iDSID = pNetworkTable.FindField("DSID")
    iIDType = pNetworkTable.FindField("OutletType")
    Dim pID, pDSID, pIDType As Integer
    Dim pCursor As ICursor
    Dim pRow As iRow
    Set pCursor = pNetworkTable.Search(Nothing, True)
    Set pRow = pCursor.NextRow
    
    Dim pDataset As IDataset
    Set pDataset = pBMPFClass
    Dim pWorkspaceEdit As IWorkspaceEdit
    Set pWorkspaceEdit = pDataset.Workspace
    pWorkspaceEdit.StartEditing False
    Dim pXAdjust As Integer
    Do While Not pRow Is Nothing
        pID = pRow.value(iID)
        pDSID = pRow.value(iDSID)
        pIDType = pRow.value(iIDType)
        'If a line exists with downstream = 0, delete it
        pQueryFilter.WhereClause = "CFROM = " & pID & " And CTO = 0"
        Set pFeatureCursor = Nothing
        Set pFeatureCursor = pFeatureclass.Update(pQueryFilter, True)
        Set pFeature = Nothing
        Set pFeature = pFeatureCursor.NextFeature
        If Not (pFeature Is Nothing) Then
            pFeature.Delete
        End If
          
        pQueryFilter.WhereClause = "From = " & pID & " And OUTLETTYPE = " & pIDType
        Set pFeatureCursor = Nothing
        Set pFeatureCursor = pFeatureclass.Update(pQueryFilter, True)
        Set pFeature = Nothing
        Set pFeature = pFeatureCursor.NextFeature
        pTO = -1
        If Not (pFeature Is Nothing) Then
            pTO = pFeature.value(iTO)
        End If
        
        Dim pBMPCount As Integer
        pBMPCount = pBMPFClass.FeatureCount(Nothing)
        Dim pConduitCount As Integer
        pConduitCount = pFeatureclass.FeatureCount(Nothing)
        
        'If the Network table has a Downstream bmp and Route Layer does not, create a new feature
        If (pDSID > 0 And pTO = -1) Then
            Set pFeature = pFeatureclass.CreateFeature
            'Enter Conduit ID, if feature layer is 'Conduits'
            If (iConduitID > 0) Then
                gConduitIDValue = pBMPCount + pConduitCount + pVFSCount + 1
                pFeature.value(iConduitID) = gConduitIDValue
            End If
            pFeature.value(iFROM) = pID
            pFeature.value(iTO) = pDSID
            pFeature.value(iOUTTYPE) = pIDType
            pFeature.value(iTypeDesc) = pRow.value(pCursor.FindField("TypeDesc"))
            
            'Create a route line
            Set pRouteLine = New Polyline
            'Add the upstream point
            pBMPFilter.WhereClause = "ID = " & pID
            Set pBMPCursor = Nothing
            Set pBMPCursor = pBMPFClass.Search(pBMPFilter, True)
            Set pBMPFeat = Nothing
            Set pBMPFeat = pBMPCursor.NextFeature
            If Not (pBMPFeat Is Nothing) Then
                pRouteLine.AddPoint pBMPFeat.Shape
            End If
            'Add the downstream point
            pBMPFilter.WhereClause = "ID = " & pDSID
            Set pBMPCursor = Nothing
            Set pBMPCursor = pBMPFClass.Search(pBMPFilter, True)
            Set pBMPFeat = Nothing
            Set pBMPFeat = pBMPCursor.NextFeature
            If Not (pBMPFeat Is Nothing) Then
                pRouteLine.AddPoint pBMPFeat.Shape
            End If
            
            'update the line to the feature and save
            Set pFeature.Shape = pRouteLine
            pFeature.Store
        End If
        
        'User has modified the downstream id, get the bmp change the value
        If (pTO <> pDSID And pTO > 0) Then
            'Enter Conduit ID, if feature layer is 'Conduits', get the conduit id
            If (iConduitID > 0) Then
                gConduitIDValue = pFeature.value(iConduitID)
            End If
            pFeature.value(iTO) = pDSID
            Set pPolyline = pFeature.Shape
            'to change the shape of line, get the new downstream bmp
            pBMPFilter.WhereClause = "ID = " & pDSID
            Set pBMPCursor = Nothing
            Set pBMPCursor = pBMPFClass.Search(pBMPFilter, True)
            Set pBMPFeat = Nothing
            Set pBMPFeat = pBMPCursor.NextFeature
            If Not (pBMPFeat Is Nothing) Then
                Set pPoint = pBMPFeat.Shape
                Select Case pIDType
                    Case 2:     'Weir
                        pXAdjust = 0
                    Case 3:     'Orifice
                        pXAdjust = 1
                    Case 4:     'Underdrain
                        pXAdjust = 2
                    Case Else:
                        pXAdjust = 0
                End Select
                'Adjust X co-ords of both From & To points
                pPoint.X = pPoint.X - pXAdjust
                pPolyline.ToPoint = pPoint
                Set pPoint = pPolyline.FromPoint
                pPoint.X = pPoint.X - pXAdjust
                pPolyline.FromPoint = pPoint
                Set pFeature.Shape = pPolyline
            End If
            pFeature.Store  'update the feature
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    RenderSchematicRouteLayer pSchemRouteLayer
    pWorkspaceEdit.StopEditing True
    'Refresh the screen
    gMxDoc.ActiveView.Refresh
    
    GoTo CleanUp
ShowError:
    MsgBox "ModifyRouteLayer: " & Err.description
CleanUp:
    Set pSchemRouteLayer = Nothing
    Set pFeatureclass = Nothing
    Set pQueryFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pSchemBMPLayer = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFilter = Nothing
    Set pBMPCursor = Nothing
    Set pBMPFeat = Nothing
    Set pPoint = Nothing
    Set pPolyline = Nothing
    Set pNetworkTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pDataset = Nothing
    Set pWorkspaceEdit = Nothing

End Sub



Public Sub ModifySingleRoute(pRouteLayerName As String, pBMPLayerName As String, FromBMP As Integer, ToBMP As Integer, pRouteType As Integer)
On Error GoTo ShowError
   
    'Call subroutine to add route
    'Create a feature class for the new schematic route layer
    Dim pSchemRouteLayer As IFeatureLayer
    Set pSchemRouteLayer = GetInputFeatureLayer(pRouteLayerName)
    If (pSchemRouteLayer Is Nothing) Then
        Dim pFClassName As String
        If (pBMPLayerName = "BMPs") Then
            pFClassName = "conduits"
        Else
            pFClassName = "schemroute"
        End If
        Dim pSchemRouteFClass As IFeatureClass
        Set pSchemRouteFClass = CreateFeatureClassForLineShapeFile(gMapTempFolder, pFClassName)
        Set pSchemRouteLayer = New FeatureLayer
        Set pSchemRouteLayer.FeatureClass = pSchemRouteFClass
        AddLayerToMap pSchemRouteLayer, pRouteLayerName
        Set pSchemRouteFClass = Nothing
    End If
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pSchemRouteLayer.FeatureClass
    Dim iConduitID As Long
    Dim iFROM As Long
    Dim iTO As Long
    Dim iOUTTYPE As Long
    Dim iTypeDesc As Long
    Dim iLabel As Long
    iConduitID = pFeatureclass.FindField("ID")
    iFROM = pFeatureclass.FindField("CFROM")
    iTO = pFeatureclass.FindField("CTO")
    iOUTTYPE = pFeatureclass.FindField("OUTLETTYPE")
    iTypeDesc = pFeatureclass.FindField("TYPEDESC")
    iLabel = pFeatureclass.FindField("LABEL")
    Dim pConduitLength As Double
    
    Dim pFROM, pTO, pOutType As Integer
    Dim pSchemBMPLayer As IFeatureLayer
    Set pSchemBMPLayer = GetInputFeatureLayer(pBMPLayerName)
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pSchemBMPLayer.FeatureClass
    Dim pBMPFilter As IQueryFilter
    Set pBMPFilter = New QueryFilter
    Dim pBMPCursor As IFeatureCursor
    Dim pBMPFeat As IFeature
    Dim pLabel As String
    Dim pFromLabel As String
    Dim pToLabel As String
    Dim iBMPLabel As Long
    iBMPLabel = pBMPFClass.FindField("LABEL")
    
    Dim pPoint As IPoint
    Dim pPolyline As IPolyline
    Dim pRouteLine As IPointCollection
    Dim pXAdjust As Integer
    Dim pRouteTypeDesc As String
    pRouteTypeDesc = GetRouteDesc(pRouteType)

   'Check if user is trying to reverse route, the type is TOTAL, then 2 conduits cannot flow from 1 bmp
   'Check for reverse route, if found, reverse the line and exit sub
   Dim pQueryFilter As IQueryFilter
   Set pQueryFilter = New QueryFilter
   pQueryFilter.WhereClause = "CFROM = " & FromBMP & " AND OUTLETTYPE = 1" & " AND CTO <> " & FromBMP
   Dim pCountExisting As Integer
   pCountExisting = 0
   pCountExisting = pFeatureclass.FeatureCount(pQueryFilter)

   'Check for reverse route, if found, reverse the line and exit sub
   pQueryFilter.WhereClause = "CTO = " & FromBMP & " AND CFROM = " & ToBMP
   Dim pFeatureCursor As IFeatureCursor
   Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, False)
   Dim pFeature As IFeature
   Set pFeature = pFeatureCursor.NextFeature

   If Not (pFeature Is Nothing) Then
        If (pCountExisting = 1) Then
            MsgBox "Cannot reverse route direction. BMP " & FromBMP & " will have 2 downstream BMPs", vbExclamation
            GoTo CleanUp
        End If
        gConduitIDValue = pFeature.value(iConduitID)
        Set pPolyline = pFeature.Shape
        pPolyline.ReverseOrientation
        'Get the label, reverse it too
        pLabel = pFeature.value(iBMPLabel)
        pFromLabel = Left(pLabel, 2)
        pToLabel = Right(pLabel, 2)
        pLabel = pToLabel & "-" & pFromLabel
        
        pFeature.value(iFROM) = FromBMP
        pFeature.value(iTO) = ToBMP
        pFeature.value(iOUTTYPE) = pRouteType
        pFeature.value(iTypeDesc) = pRouteTypeDesc
        pFeature.value(iLabel) = pLabel
        Set pFeature.Shape = pPolyline
        pFeature.Store
        GoTo RefreshMap
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
    
    'Check if other route exist, if not create it and define the direction
    pQueryFilter.WhereClause = "CFROM = " & FromBMP & " AND OUTLETTYPE = " & pRouteType
    Set pFeatureCursor = Nothing
    Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, False)
    Set pFeature = pFeatureCursor.NextFeature
    If (pFeature Is Nothing) Then
        'Find new conduit id
        Dim pBMPCount As Integer
        pBMPCount = pBMPFClass.FeatureCount(Nothing)
        Dim pConduitCount As Integer
        pConduitCount = pFeatureclass.FeatureCount(Nothing)
        Set pFeature = pFeatureclass.CreateFeature
        'Enter Conduit ID, if feature layer is 'Conduits'
        gConduitIDValue = pBMPCount + pConduitCount + pVFSCount + 1
        pFeature.value(iConduitID) = gConduitIDValue
    Else
        gConduitIDValue = pFeature.value(iConduitID)
    End If

    
    'Create a route line
    Set pRouteLine = New Polyline
    'Add the upstream point
    pBMPFilter.WhereClause = "ID = " & FromBMP
    Set pBMPCursor = Nothing
    Set pBMPCursor = pBMPFClass.Search(pBMPFilter, False)
    Set pBMPFeat = Nothing
    Set pBMPFeat = pBMPCursor.NextFeature
    pFromLabel = ""
    If Not (pBMPFeat Is Nothing) Then
        pFromLabel = pBMPFeat.value(iBMPLabel)
        pRouteLine.AddPoint pBMPFeat.Shape
    End If
    
    '** If frompoint is not a bmp, it must be a vfs
    Dim pVFSLayer As IFeatureLayer
    Set pVFSLayer = GetInputFeatureLayer("VFS")
    Dim pVFSClass As IFeatureClass
    Dim pVFSPolyline As IPolyline
    If (pFromLabel = "" And (Not pVFSLayer Is Nothing)) Then
        Set pVFSClass = pVFSLayer.FeatureClass
        pBMPFilter.WhereClause = "ID = " & FromBMP
        Set pBMPCursor = Nothing
        Set pBMPCursor = pVFSClass.Search(pBMPFilter, False)
        Set pBMPFeat = Nothing
        Set pBMPFeat = pBMPCursor.NextFeature
        pFromLabel = ""
        If Not (pBMPFeat Is Nothing) Then
            Set pVFSPolyline = pBMPFeat.Shape
            pFromLabel = pBMPFeat.value(pVFSClass.FindField("LABEL"))
            pRouteLine.AddPoint pVFSPolyline.ToPoint
        End If
    End If
    
    
    'Add the downstream point
    pBMPFilter.WhereClause = "ID = " & ToBMP
    Set pBMPCursor = Nothing
    Set pBMPCursor = pBMPFClass.Search(pBMPFilter, False)
    Set pBMPFeat = Nothing
    Set pBMPFeat = pBMPCursor.NextFeature
    pToLabel = ""
    If Not (pBMPFeat Is Nothing) Then
        pToLabel = pBMPFeat.value(iBMPLabel)
        pRouteLine.AddPoint pBMPFeat.Shape
    End If
    
    '** If downstream point is not a bmp, it must be a vfs
    If (pToLabel = "" And (Not pVFSLayer Is Nothing)) Then
        Set pVFSClass = pVFSLayer.FeatureClass
        pBMPFilter.WhereClause = "ID = " & ToBMP
        Set pBMPCursor = Nothing
        Set pBMPCursor = pVFSClass.Search(pBMPFilter, False)
        Set pBMPFeat = Nothing
        Set pBMPFeat = pBMPCursor.NextFeature
        pToLabel = ""
        If Not (pBMPFeat Is Nothing) Then
            Set pVFSPolyline = pBMPFeat.Shape
            pToLabel = pBMPFeat.value(iBMPLabel)
            pRouteLine.AddPoint pVFSPolyline.FromPoint
        End If
    End If
       
       
    'Conduit label is from BMP - to BMP labels
    pLabel = pFromLabel & "-" & pToLabel
    
    'SET feature values
    pFeature.value(iFROM) = FromBMP
    pFeature.value(iTO) = ToBMP
    pFeature.value(iOUTTYPE) = pRouteType
    pFeature.value(iTypeDesc) = pRouteTypeDesc
    pFeature.value(iLabel) = pLabel
    'update the line to the feature and save
    Set pFeature.Shape = pRouteLine
    pFeature.Store
    
    Set pPolyline = pRouteLine
    pConduitLength = Format(pPolyline.Length * gMetersPerUnit * 3.28, "#.##")
    
RefreshMap:

''    '*** Initialize Pollutant Data
''    Call InitPollutantData

    '**** Mira Chokshi, Feb 4 2005, To input conduit cross-section
    FrmConduitCSection.txtLength.Text = CStr(pConduitLength)
    FrmConduitCSection.Show vbModal
    
    'Refresh the screen, and render the route layer
    RenderSchematicRouteLayer pSchemRouteLayer
    gMxDoc.ActiveView.Refresh
    
    GoTo CleanUp
ShowError:
    MsgBox "ModifySingleRoute: " & Err.description & vbTab & Err.Number
CleanUp:
    Set pSchemRouteLayer = Nothing
    Set pFeatureclass = Nothing
    Set pQueryFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pSchemBMPLayer = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFilter = Nothing
    Set pBMPCursor = Nothing
    Set pBMPFeat = Nothing
    Set pPoint = Nothing
    Set pPolyline = Nothing
    Set pVFSLayer = Nothing
    Set pVFSClass = Nothing
    Set pVFSPolyline = Nothing
End Sub


Public Function GetRouteDesc(outletType As Integer) As String
    Dim pRouteTypeDesc As String
    Select Case outletType
        Case 1:
            pRouteTypeDesc = "Total"
        Case 2:
            pRouteTypeDesc = "Weir"
        Case 3:
            pRouteTypeDesc = "Orifice/Channel"
        Case 4:
            pRouteTypeDesc = "Underdrain"
    End Select
    GetRouteDesc = pRouteTypeDesc
End Function

'changes: Ying Cao 11/28/2008 accepts VFS as input in addition to BMP
Private Sub AddSchematicBMP()
On Error GoTo ShowError

    Dim pToPnt As IPnt
    Set pToPnt = New DblPnt
    Dim pFromPnt As IPnt
    Set pFromPnt = New DblPnt
    
    Dim pX As Double
    Dim pY As Double
    Dim i As Integer
    Dim j As Integer
    Dim iBMP As Integer
    Dim ToBMP As Integer
    Dim FromBMP As Integer
    Dim typeBMP As String
    Dim typeBMP2 As String
    Dim labelBMP As String
    Dim pAddXY As Double
    
    Dim pBMPFeature As IFeature
    Dim pPoint As IPoint
    Dim pBMPKeys
    pBMPKeys = pRouteDictionary.keys
    Dim pConstAngle As Double
    pConstAngle = 3.1428 / CDbl((pRouteDictionary.Count + 1))
    Dim pConstDist As Double
    pConstDist = 100
    Dim pAngle As Double
    Dim pDistance As Double

    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    If (pBMPFLayer Is Nothing And pVFSFLayer Is Nothing) Then
        MsgBox "BMPs feature layer or VFS layer not found."
        Exit Sub
    End If
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pBMPCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pType As Long
    pType = pBMPFClass.FindField("TYPE")
    
    'for VFS
    Dim pVFSFClass As IFeatureClass
    Dim pVFSCursor As IFeatureCursor
    Dim pFeature2 As IFeature
    If Not pVFSFClass Is Nothing Then   'vfs layer may not exist
        Set pVFSFClass = pVFSFLayer.FeatureClass
    End If
    'For schematic
    Dim pType2 As Long 'Sabu paul - Sept 20, 2004
    Dim pLabel As Long                      'Mira Chokshi - 03/23/2005
                                            'VFS shares the same index for Type2 and Label
    pType2 = pBMPFClass.FindField("TYPE2")  'Sabu paul - Sept 20, 2004
    pLabel = pBMPFClass.FindField("LABEL")

    'Fields of schematic bmp class
    Dim iID As Long
    Dim iDSID As Long
    Dim iType As Long
    iID = pSchemClass.FindField("ID")
    iDSID = pSchemClass.FindField("DSID")
    iType = pSchemClass.FindField("TYPE")

    Dim iType2 As Long                  'Sabu paul - Sept 20, 2004
    iType2 = pSchemClass.FindField("TYPE2") 'Sabu paul - Sept 20, 2004
    
    Dim iLabel As Long                      'Mira Chokshi - 03/23/2005
    iLabel = pSchemClass.FindField("LABEL")
    
    For i = 1 To pRouteDictionary.Count
        pAngle = pConstAngle * i    'Get angle from 0 degree
        FromBMP = pBMPKeys(i - 1)
        ToBMP = pRouteDictionary.Item(FromBMP)
        If (ToBMP = 0) Then    'Output BMP
            If (pTotalOutlets = 0) Then
                pX = 1000    'Initialize some value
                pY = 1000
            Else
                pX = pMaxHorizontal + 200
                pY = 1000
            End If
            pTotalOutlets = pTotalOutlets + 1
        Else
            'Get the x,y co-ordinates of the downstream (pToBMP)
            Set pToPnt = Nothing
            Set pToPnt = New DblPnt
            Set pToPnt = pDictionary.Item(ToBMP)
            
            '*** Check if this fromBMP is a downstream BMP to other bmps
            pAddXY = 1
            If (pUpstreamBMPPresent.Exists(FromBMP)) Then
                pAddXY = 2.5
            End If
            
            pX = pToPnt.X - (pConstDist * Cos(pAngle) * pAddXY)  'Point on left
            pY = pToPnt.Y + (pConstDist * Sin(pConstAngle) * pAddXY)     'Point on top
        End If
        
        If ((pX + Abs(pConstDist * Cos(pAngle))) > pMaxHorizontal) Then
            pMaxHorizontal = pX + Abs(pConstDist * Cos(pAngle))
        End If
        'Save Point
        Set pFromPnt = Nothing
        Set pFromPnt = New DblPnt
        
        pFromPnt.SetCoords pX, pY
        Set pPoint = New Point
        pPoint.X = pX
        pPoint.Y = pY
        Set pBMPFeature = Nothing
        Set pBMPFeature = pSchemClass.CreateFeature
        Set pBMPFeature.Shape = pPoint
        
        pQueryFilter.WhereClause = "ID = " & FromBMP
        Set pBMPCursor = Nothing
        Set pBMPCursor = pBMPFClass.Search(pQueryFilter, True)
        Set pFeature = Nothing
        Set pFeature = pBMPCursor.NextFeature
        
        If Not (pFeature Is Nothing) Then       'BMP
            typeBMP = pFeature.value(pType)
            typeBMP2 = pFeature.value(pType2) '--Sabu Paul
            labelBMP = pFeature.value(pLabel)
        Else                                    'VFS
            Set pVFSCursor = Nothing
            Set pVFSCursor = pVFSFClass.Search(pQueryFilter, True)
            Set pFeature2 = Nothing
            Set pFeature2 = pVFSCursor.NextFeature
            typeBMP = pFeature2.value(pType)
            typeBMP2 = pFeature2.value(pType2)
            labelBMP = pFeature2.value(pLabel)
        End If
        
        pBMPFeature.value(iID) = FromBMP
        pBMPFeature.value(iDSID) = ToBMP
        pBMPFeature.value(iType) = typeBMP
        pBMPFeature.value(iType2) = typeBMP2 '--Sabu Paul
        pBMPFeature.value(iLabel) = labelBMP
        pBMPFeature.Store
                                       
        'Add the point to dictionary
        pDictionary.Item(FromBMP) = pFromPnt
    
    Next
    'Remove all mapping
    pRouteDictionary.RemoveAll
    Exit Sub
    
ShowError:
    MsgBox "AddSchematicBMP: " & Err.description
   

End Sub

Public Function AddSchematicBMPRoute(pRouteFeatureClass As IFeatureClass, pBMPFeatureClass As IFeatureClass) As Boolean

    Dim pNetworkTable As iTable
    Set pNetworkTable = GetInputDataTable("BMPNetwork")
    If (pNetworkTable Is Nothing) Then
        MsgBox "BMPNetwork table not found."
        Exit Function
    End If
    Dim pConduitFLayer As IFeatureLayer
    Set pConduitFLayer = GetInputFeatureLayer("Conduits")
    If (pConduitFLayer Is Nothing) Then
        MsgBox "Conduits feature layer not found."
        Exit Function
    End If
    
    Dim iID As Long
    Dim iDSID As Long
    Dim iOUTTYPE As Long
    Dim iTypeDesc As Long
    iID = pNetworkTable.FindField("ID")
    iDSID = pNetworkTable.FindField("DSID")
    iOUTTYPE = pNetworkTable.FindField("OUTLETTYPE")
    iTypeDesc = pNetworkTable.FindField("TYPEDESC")
    
    
    Dim pIDValue As Integer
    Dim FromBMP As Integer
    Dim ToBMP As Integer
    Dim outletType As Integer
    Dim typeDESC As String
    Dim pCursor As ICursor
    Dim pRow As iRow
    Set pCursor = pNetworkTable.Search(Nothing, True)
    Set pRow = pCursor.NextRow
    
    Dim pQueryFilter As IQueryFilter
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pFromPoint As IPoint
    Dim pToPoint As IPoint
    Dim pRouteLine As IPointCollection

    'Fields of schematic route class
    Dim iRteID As Long
    iRteID = pRouteFeatureClass.FindField("ID")
    Dim iFROM As Long
    iFROM = pRouteFeatureClass.FindField("CFROM")
    Dim pOutType As Long
    pOutType = pRouteFeatureClass.FindField("OUTLETTYPE")
    Dim iTO As Long
    iTO = pRouteFeatureClass.FindField("CTO")
    Dim pTypeDesc As Long
    pTypeDesc = pRouteFeatureClass.FindField("TYPEDESC")
    Dim pSchLabel As Long
    pSchLabel = pRouteFeatureClass.FindField("LABEL")
    
    Dim iBMPLabel As Long
    iBMPLabel = pBMPFeatureClass.FindField("LABEL")
    Dim pLabelVal As String
    pLabelVal = ""
    
    Dim pLineCount As Long
    Dim pXAdjust As Double
    Do While Not pRow Is Nothing
        FromBMP = pRow.value(iID)
        ToBMP = pRow.value(iDSID)
        outletType = pRow.value(iOUTTYPE)
        typeDESC = pRow.value(iTypeDesc)
        Set pRouteLine = New Polyline
        Select Case outletType
            Case 2:     'Weir
                pXAdjust = 0
            Case 3:     'Orifice
                pXAdjust = 1
            Case 4:     'Underdrain
                pXAdjust = 2
            Case Else:
                pXAdjust = 0
        End Select
                
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "ID = " & FromBMP
        Set pFeatureCursor = pBMPFeatureClass.Search(pQueryFilter, True)
        Set pFeature = pFeatureCursor.NextFeature
        If Not pFeature Is Nothing Then
              Set pFromPoint = pFeature.Shape
              pLabelVal = pFeature.value(iBMPLabel)
              pFromPoint.X = pFromPoint.X - pXAdjust
              pRouteLine.AddPoint pFromPoint
        End If
        'Add - connector
        pLabelVal = pLabelVal & "-"
        pQueryFilter.WhereClause = "ID = " & ToBMP
        Set pFeatureCursor = pBMPFeatureClass.Search(pQueryFilter, True)
        Set pFeature = pFeatureCursor.NextFeature
        If Not pFeature Is Nothing Then
              Set pToPoint = pFeature.Shape
                pLabelVal = pLabelVal & pFeature.value(iBMPLabel)
              pToPoint.X = pToPoint.X - pXAdjust
              pRouteLine.AddPoint pToPoint
        End If
        
        'Find the ID of this feature in Conduits Feature Class
        Dim pConduitFClass As IFeatureClass
        Set pConduitFClass = pConduitFLayer.FeatureClass
        pQueryFilter.WhereClause = "CFROM = " & FromBMP & " AND CTO = " & ToBMP
        Set pFeatureCursor = pConduitFClass.Search(pQueryFilter, True)
        Set pFeature = pFeatureCursor.NextFeature
        If Not pFeature Is Nothing Then
            pIDValue = pFeature.value(pFeatureCursor.FindField("ID"))
        End If
        
        'If downstream is not an outlet
        If (ToBMP <> 0) Then
            'Add Line to Layer
            Dim pLineFeature As IFeature
            Set pLineFeature = pRouteFeatureClass.CreateFeature
            Set pLineFeature.Shape = pRouteLine
            pLineFeature.value(iRteID) = pIDValue
            pLineFeature.value(iFROM) = FromBMP
            pLineFeature.value(iTO) = ToBMP
            pLineFeature.value(pOutType) = outletType
            pLineFeature.value(pTypeDesc) = typeDESC
            pLineFeature.value(pSchLabel) = pLabelVal
            pLineFeature.Store
        End If
        
        'Next row in network table
        Set pRow = pCursor.NextRow
    Loop
    
    If (pRouteFeatureClass.FeatureCount(Nothing) = 0) Then
        AddSchematicBMPRoute = False
    Else
        AddSchematicBMPRoute = True
    End If
    
End Function
'*******************************************************************************
'Subroutine : CreatePointFeatureClassForSchematics
'Purpose    : Creates a new feature class file to store the point shapes
'Arguments  : Destination directory, name of the feature class file
'Author     : Mira Chokshi - 08/26/04
'*******************************************************************************
Public Function CreatePointFeatureClassForSchematics(DirName As String, FileName As String) As IFeatureClass

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

  Dim pFieldDSID As esriGeoDatabase.IField
  Dim pFieldEditDSID As IFieldEdit
  Set pFieldDSID = New esriGeoDatabase.Field
  Set pFieldEditDSID = pFieldDSID
  pFieldEditDSID.name = "DSID"
  pFieldEditDSID.Type = esriFieldTypeInteger
  pFieldEditDSID.IsNullable = True
  
  Dim pFieldType As esriGeoDatabase.IField
  Dim pFieldEditType As IFieldEdit
  Set pFieldType = New esriGeoDatabase.Field
  Set pFieldEditType = pFieldType
  pFieldEditType.name = "TYPE"
  pFieldEditType.Type = esriFieldTypeString
  pFieldEditType.Length = 30
  pFieldEditType.IsNullable = True
  
  Dim pFieldType2 As esriGeoDatabase.IField                 ' Sabu Paul
  Dim pFieldEditType2 As IFieldEdit
  Set pFieldType2 = New esriGeoDatabase.Field
  Set pFieldEditType2 = pFieldType2
  pFieldEditType2.name = "TYPE2"
  pFieldEditType2.Type = esriFieldTypeString
  pFieldEditType2.Length = 30
  pFieldEditType2.IsNullable = True
  
  Dim pFieldLABEL As esriGeoDatabase.IField
  Dim pFieldEditLABEL As IFieldEdit
  Set pFieldLABEL = New esriGeoDatabase.Field
  Set pFieldEditLABEL = pFieldLABEL
  pFieldEditLABEL.name = "LABEL"
  pFieldEditLABEL.Type = esriFieldTypeString
  pFieldEditLABEL.IsNullable = True
  
  
  'Create a SpatialReferenceFactory
  Dim pSpatialRefFact As ISpatialReferenceFactory2
  Set pSpatialRefFact = New SpatialReferenceEnvironment
    
  'Create the two coordinate systems
  Dim pGeographic As IGeographicCoordinateSystem
  Set pGeographic = pSpatialRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_WGS1984)
  
  Dim pRasterDEMProps As IRasterAnalysisProps
  If Not GetInputRasterLayer("DEM") Is Nothing Then
    Set pRasterDEMProps = GetInputRasterLayer("DEM").Raster
  Else
    Set pRasterDEMProps = Nothing
  End If
  
  'In case DEM is optional get the properties from Landuse -- Sabu Paul
  Dim pRasterLUProps As IRasterAnalysisProps
  Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster
  
  Dim pGeomDef As IGeometryDef
  Dim pGeomDefEdit As IGeometryDefEdit
  Set pGeomDef = New GeometryDef
  Set pGeomDefEdit = pGeomDef
  With pGeomDefEdit
    .GeometryType = esriGeometryPoint
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
  pFieldsEdit.AddField pFieldType2 ' Sabu Paul
  pFieldsEdit.AddField pFieldLABEL ' Mira Chokshi 03/23/2005
  
  ' Create the shapefile some parameters apply to geodatabase options and can be defaulted as Nothing
  Dim pFeatClass As IFeatureClass
  Set pFeatClass = pFWS.CreateFeatureClass(pFileName, pFields, Nothing, Nothing, esriFTSimple, strShapeFieldName, "")

  ' Return the value
  Set CreatePointFeatureClassForSchematics = pFeatClass
    
  GoTo CleanUp
ShowError:
    MsgBox "CreatePointFeatureClassForSchematics: " & Err.description
CleanUp:

End Function




'''******************************************************************************
'''Subroutine: CreateFeatureClassForSchematicRoute
'''Author:     Mira Chokshi
'''Purpose:    Creates line shape feature class for conduit network class.
'''            Adds integer fields: ID, CFROM, CTO. Sets the spatial reference of
'''            the feature class same as dem's spatial reference
'''******************************************************************************
''Public Function CreateFeatureClassForSchematicRoute(DirName As String, FileName As String) As IFeatureClass
''
''On Error GoTo ShowError
''    'Create a unique file name for feature class
''    Dim pFileName As String
''    pFileName = CreateUniqueTableName(DirName, FileName)
''    Dim strFolder As String
''    strFolder = DirName
''    Dim strShapeFieldName As String
''    strShapeFieldName = "Shape"
''    ' Open the folder to contain the shapefile as a workspace
''    Dim pFWS As IFeatureWorkspace
''    Dim pWorkspaceFactory As IWorkspaceFactory
''    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
''    Set pFWS = pWorkspaceFactory.OpenFromFile(strFolder, 0)
''    ' Set up a simple fields collection
''    Dim pFields As IFields
''    Dim pFieldsEdit As IFieldsEdit
''    Set pFields = New Fields
''    Set pFieldsEdit = pFields
''    Dim pFieldShape As IField
''    Dim pFieldEditShape As IFieldEdit
''
''    ' Make the shape field
''    ' it will need a geometry definition, with a spatial reference
''    Set pFieldShape = New esriGeoDatabase.Field
''    Set pFieldEditShape = pFieldShape
''    pFieldEditShape.Name = strShapeFieldName
''    pFieldEditShape.Type = esriFieldTypeGeometry
''    'Define FROM field
''    Dim pFieldFrom As IField
''    Dim pFieldEditFrom As IFieldEdit
''    Set pFieldFrom = New esriGeoDatabase.Field
''    Set pFieldEditFrom = pFieldFrom
''    pFieldEditFrom.Name = "CFROM"
''    pFieldEditFrom.Type = esriFieldTypeInteger
''    pFieldEditFrom.IsNullable = True
''    'Define OUTLETTYPE field
''    Dim pFieldOUTTYPE As IField
''    Dim pFieldEditOUTTYPE As IFieldEdit
''    Set pFieldOUTTYPE = New esriGeoDatabase.Field
''    Set pFieldEditOUTTYPE = pFieldOUTTYPE
''    pFieldEditOUTTYPE.Name = "OUTLETTYPE"
''    pFieldEditOUTTYPE.Type = esriFieldTypeInteger
''    pFieldEditOUTTYPE.IsNullable = True
''    'Define OUTLET Description field
''    Dim pFieldOUTDESC As IField
''    Dim pFieldEditOUTDESC As IFieldEdit
''    Set pFieldOUTDESC = New esriGeoDatabase.Field
''    Set pFieldEditOUTDESC = pFieldOUTDESC
''    pFieldEditOUTDESC.Name = "TYPEDESC"
''    pFieldEditOUTDESC.Type = esriFieldTypeString
''    pFieldEditOUTDESC.IsNullable = True
''
''    'Define CTO field
''    Dim pFieldTo As IField
''    Dim pFieldEditTo As IFieldEdit
''    Set pFieldTo = New esriGeoDatabase.Field
''    Set pFieldEditTo = pFieldTo
''    pFieldEditTo.Name = "CTO"
''    pFieldEditTo.Type = esriFieldTypeInteger
''    pFieldEditTo.IsNullable = True
''
''    'Get DEM raster properties
''    Dim pRasterDEMProps As IRasterAnalysisProps
''    If Not GetInputRasterLayer("DEM") Is Nothing Then
''        Set pRasterDEMProps = GetInputRasterLayer("DEM").Raster
''    Else
''        Set pRasterDEMProps = Nothing
''    End If
''
''    'Get Landuse raster properties in case DEM is optional -- Sabu Paul
''    Dim pRasterLUProps As IRasterAnalysisProps
''    Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster
''
''    'Get spatial reference properties
''    Dim pGeomDef As IGeometryDef
''    Dim pGeomDefEdit As IGeometryDefEdit
''    Set pGeomDef = New GeometryDef
''    Set pGeomDefEdit = pGeomDef
''    With pGeomDefEdit
''      .GeometryType = esriGeometryPolyline
''      .HasM = False
''      .HasZ = False
''      If Not pRasterDEMProps Is Nothing Then
''        Set .SpatialReference = pRasterDEMProps.AnalysisExtent.SpatialReference
''      Else
''        Set .SpatialReference = pRasterLUProps.AnalysisExtent.SpatialReference
''      End If
''    End With
''    Set pFieldEditShape.GeometryDef = pGeomDef
''    pFieldsEdit.AddField pFieldShape
''    pFieldsEdit.AddField pFieldFrom
''    pFieldsEdit.AddField pFieldOUTTYPE
''    pFieldsEdit.AddField pFieldTo
''    pFieldsEdit.AddField pFieldOUTDESC
''
''    ' Create the shapefile some parameters apply to geodatabase options and can be defaulted as Nothing
''    Dim pFeatClass As IFeatureClass
''    Set pFeatClass = pFWS.CreateFeatureClass(pFileName, pFields, Nothing, Nothing, esriFTSimple, strShapeFieldName, "")
''    ' Return the value
''    Set CreateFeatureClassForSchematicRoute = pFeatClass
''
''  GoTo CleanUp
''ShowError:
''    MsgBox "CreateFeatureClassForSchematicRoute: " & Err.Description
''CleanUp:
''    Set pFWS = Nothing
''    Set pWorkspaceFactory = Nothing
''    Set pFields = Nothing
''    Set pFieldsEdit = Nothing
''    Set pFieldShape = Nothing
''    Set pFieldEditShape = Nothing
''    Set pFieldEditOUTTYPE = Nothing
''    Set pFieldOUTTYPE = Nothing
''    Set pFieldFrom = Nothing
''    Set pFieldEditFrom = Nothing
''    Set pFieldTo = Nothing
''    Set pFieldEditTo = Nothing
''    Set pRasterDEMProps = Nothing
''    Set pRasterLUProps = Nothing
''    Set pGeomDef = Nothing
''    Set pGeomDefEdit = Nothing
''    Set pFeatClass = Nothing
''End Function


Public Sub RenderSchematicRouteLayer(pFeatureLayer As IFeatureLayer)

     '** Make the renderer
     Dim pRender As IUniqueValueRenderer
     Set pRender = New UniqueValueRenderer
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "OUTLETTYPE"
     pRender.UseDefaultSymbol = False
     
     Dim pFeatureclass As IFeatureClass
     Set pFeatureclass = pFeatureLayer.FeatureClass
     Dim pFeatureCursor As IFeatureCursor
     Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
     Dim pFeature As IFeature
     Set pFeature = pFeatureCursor.NextFeature
     Dim iID As Long
     iID = pFeatureclass.FindField("ID")
     Dim iType As Long
     iType = pFeatureclass.FindField("OUTLETTYPE")
     Dim iTypeDesc As Long
     iTypeDesc = pFeatureclass.FindField("TYPEDESC")
     Dim pBMPType As Integer
     Dim pBMPTypeDesc As String
     Dim ValFound As Boolean
     Dim pRouteLineSymbol As ICartographicLineSymbol
     Dim uh As Integer
     
     Do While Not pFeature Is Nothing
         'get values for feature
         pBMPType = pFeature.value(iType)
         pBMPTypeDesc = pFeature.value(iTypeDesc)
         
         Set pRouteLineSymbol = ReturnBMPtoBMPRouteSymbol(pBMPType)
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.value(uh) = pBMPType Then
             ValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             pRender.AddValue pBMPType, "OUTLETTYPE", pRouteLineSymbol
             pRender.Label(pBMPType) = pBMPTypeDesc
             pRender.Symbol(pBMPType) = pRouteLineSymbol
         End If
        Set pFeature = pFeatureCursor.NextFeature
     Loop
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     Dim pLyr As IGeoFeatureLayer
     Set pLyr = pFeatureLayer
     pRender.ColorScheme = "Custom"
     pRender.fieldType(0) = False
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = "TYPEDESC"
 
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
     aLELayerProps.Expression = "[LABEL]"
     Dim pTextSymbol As ITextSymbol
     Set pTextSymbol = New TextSymbol
     pTextSymbol.Size = 8
     Dim pColor As IRgbColor
     Set pColor = New RgbColor
     pColor.RGB = vbBlack
     pTextSymbol.Color = pColor
     Set aLELayerProps.Symbol = pTextSymbol
     ' assign it to the layer's AnnotateLayerPropertiesCollection
     pAnnoLayerPropsColl.add aLELayerProps
     'get the BasicOverposterLayerProperties
     Dim pBasicOverposterLayerProps As IBasicOverposterLayerProperties
     Set pBasicOverposterLayerProps = aLELayerProps.BasicOverposterLayerProperties
     pBasicOverposterLayerProps.NumLabelsOption = esriOneLabelPerShape
     
     '** This makes the layer properties symbology tab show the correct interface.
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     '** Draw the map
     gMxDoc.ActiveView.Refresh

End Sub

'* Subroutine to define conduit network for BMPs
Public Function ReturnBMPtoBMPRouteSymbol(pRouteType As Integer) As ICartographicLineSymbol

On Error GoTo ShowError
    
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    Select Case pRouteType
        Case "1":   'TOTAL
            pColor.RGB = vbBlack
        Case "2":   'Weir
            pColor.RGB = vbRed
        Case "3":   'Orifice
            pColor.RGB = vbGreen
        Case "4":   'UnderDrain
            pColor.RGB = vbBlue
        Case Else
            pColor.RGB = vbBlack
    End Select
    
    'This is setting up the marker symbols for the lines
    Dim pMarker As IArrowMarkerSymbol
    Set pMarker = New ArrowMarkerSymbol
    pMarker.Style = esriAMSPlain
    
    pMarker.Angle = 0
    pMarker.Color = pColor
    pMarker.Length = 8#
    pMarker.Width = 6#
    pMarker.Size = 8#
    pMarker.XOffset = 0
    pMarker.YOffset = 0
        
    Dim thelinedec As ILineDecoration
    Set thelinedec = New LineDecoration
    
    Dim ptsymbol As ISimpleLineDecorationElement
    Set ptsymbol = New SimpleLineDecorationElement
    ptsymbol.Rotate = True
    ptsymbol.MarkerSymbol = pMarker
    ptsymbol.ClearPositions
    ptsymbol.AddPosition (0.5)
    ptsymbol.FlipFirst = True
    thelinedec.AddElement ptsymbol
     
    Dim theproperties As ILineProperties
    Dim theline As ICartographicLineSymbol
    Set theline = New CartographicLineSymbol
    theline.Color = pColor
    theline.Width = 1
    Set theproperties = theline
      
    Set theproperties.LineDecoration = thelinedec
    theproperties.Offset = 0
        
    'Return the symbol
    Set ReturnBMPtoBMPRouteSymbol = theline
    
    GoTo CleanUp
ShowError:
    MsgBox "ReturnBMPtoBMPRouteSymbol: " & Err.description
CleanUp:
    Set pColor = Nothing
    Set pMarker = Nothing
    Set thelinedec = Nothing
    Set ptsymbol = Nothing
    Set theproperties = Nothing
    Set theline = Nothing
End Function


'* Subroutine to define conduit network for Basins to BMPs
Public Function ReturnBasintoBMPRouteSymbol() As ICartographicLineSymbol

On Error GoTo ShowError
    
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    pColor.RGB = RGB(0, 38, 155)
    
    'This is setting up the marker symbols for the lines
    Dim pMarker As ISimpleMarkerSymbol
    Set pMarker = New SimpleMarkerSymbol
    pMarker.Style = esriSMSSquare
    
    pMarker.Angle = 0
    pMarker.Color = pColor
    pMarker.Size = 4#
    pMarker.XOffset = 0
    pMarker.YOffset = 0
        
    Dim thelinedec As ILineDecoration
    Set thelinedec = New LineDecoration
    
    Dim ptsymbol As ISimpleLineDecorationElement
    Set ptsymbol = New SimpleLineDecorationElement
    ptsymbol.Rotate = True
    ptsymbol.MarkerSymbol = pMarker
    ptsymbol.ClearPositions
    ptsymbol.AddPosition (0)
    ptsymbol.FlipFirst = True
    thelinedec.AddElement ptsymbol
            
    Dim theproperties As ILineProperties
    Dim theline As ICartographicLineSymbol
    Set theline = New CartographicLineSymbol
    Dim theTemplate As ITemplate
    Set theproperties = theline
    theline.Color = pColor
    theline.Width = 1
    Set theTemplate = New Template
    theTemplate.Interval = 1
    theTemplate.AddPatternElement 4, 2
    Set theproperties.Template = theTemplate
      
    Set theproperties.LineDecoration = thelinedec
    theproperties.Offset = 0
        
    'Return the symbol
    Set ReturnBasintoBMPRouteSymbol = theline
    
    GoTo CleanUp
ShowError:
    MsgBox "ReturnBasintoBMPRouteSymbol: " & Err.description
CleanUp:

End Function


Public Sub RenderSchematicBMPLayer(pFeatureLayer As IFeatureLayer)
On Error GoTo ShowError
     '** Make the renderer
     Dim pRender As IUniqueValueRenderer
     Set pRender = New UniqueValueRenderer
    
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "TYPE2"
     pRender.UseDefaultSymbol = False
     
     Dim pFeatureclass As IFeatureClass
     Set pFeatureclass = pFeatureLayer.FeatureClass
     Dim pFeatureCursor As IFeatureCursor
     Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
     Dim pFeature As IFeature
     Set pFeature = pFeatureCursor.NextFeature
     Dim iID As Long
     iID = pFeatureclass.FindField("ID")
     Dim iType As Long
     iType = pFeatureclass.FindField("TYPE2")   ' "TYPE")
     Dim pBMPType As String
     Dim ValFound As Boolean
     Dim pBMPMarkerSymbol As IMultiLayerMarkerSymbol    ' IPictureMarkerSymbol
     Dim uh As Integer
     
     Dim pExternalTS As iTable
     Set pExternalTS = GetInputDataTable("ExternalTS")
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter
     Dim pTSFlag As Boolean
     
     Do While Not pFeature Is Nothing
         pBMPType = pFeature.value(iType)
         pTSFlag = False
         'Check if this bmp has a external time series
         If Not (pExternalTS Is Nothing) Then
            pQueryFilter.WhereClause = "BMPID = " & pFeature.value(iID)
            If (pExternalTS.RowCount(pQueryFilter) = 1) Then
                pTSFlag = True
            End If
         End If
         
         Set pBMPMarkerSymbol = ReturnBMPSymbol(pBMPType, pTSFlag)
                 
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.value(uh) = pBMPType Then
             ValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             pRender.AddValue pBMPType, "Type", pBMPMarkerSymbol
             pRender.Label(pBMPType) = pBMPType
             pRender.Symbol(pBMPType) = pBMPMarkerSymbol
         End If
        Set pFeature = pFeatureCursor.NextFeature
     Loop
    
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     Dim pLyr As IGeoFeatureLayer
     Set pLyr = pFeatureLayer
     pRender.ColorScheme = "Custom"
     pRender.fieldType(0) = True
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = "Type2"
 
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
     aLELayerProps.Expression = "[LABEL]"
     Dim pTextSymbol As ISimpleTextSymbol
     Set pTextSymbol = New TextSymbol
     
     'Create a font symbol and grab hold of the stdole.stdFont interface
     Dim pFont As stdole.StdFont
     Set pFont = New stdole.StdFont
     'Set font and text symbol properties
     pFont.name = "Arial"
     pFont.Bold = True
     pFont.Size = 9
     pTextSymbol.Font = pFont
     
     Dim pColor As IRgbColor
     Set pColor = New RgbColor
     pColor.RGB = RGB(56, 168, 0) 'greenish
     pTextSymbol.Color = pColor
     
     'Set x, y offset
     pTextSymbol.XOffset = 0
     pTextSymbol.YOffset = 0
     
     Set aLELayerProps.Symbol = pTextSymbol
     ' assign it to the layer's AnnotateLayerPropertiesCollection
     pAnnoLayerPropsColl.add aLELayerProps
    
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     '** Draw the map
     gMxDoc.ActiveView.Refresh
     GoTo CleanUp
     
ShowError:
    MsgBox "RenderSchematicBMPLayer: " & Err.description
CleanUp:
     Set pRender = Nothing
     Set pFeatureclass = Nothing
     Set pFeatureCursor = Nothing
     Set pFeature = Nothing
     Set pBMPMarkerSymbol = Nothing
     Set pLyr = Nothing
     Set pAnnoLayerPropsColl = Nothing
     Set aLELayerProps = Nothing
     Set pTextSymbol = Nothing
     Set pColor = Nothing
End Sub


Public Sub RenderPointViewBMPLayer(pFeatureLayer As IFeatureLayer)

    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    pColor.RGB = vbBlack
    
    Dim pMarkerSymbol As ISimpleMarkerSymbol
    Set pMarkerSymbol = New SimpleMarkerSymbol
    pMarkerSymbol.Size = 5
    pMarkerSymbol.Color = pColor
    pMarkerSymbol.Style = esriSMSCircle
   
    Dim pSimpleRenderer As ISimpleRenderer
    Set pSimpleRenderer = New SimpleRenderer
    Set pSimpleRenderer.Symbol = pMarkerSymbol
   
    Dim pGeoFeatLyr As IGeoFeatureLayer
    Set pGeoFeatLyr = pFeatureLayer
    Set pGeoFeatLyr.Renderer = pSimpleRenderer
    
    gMxDoc.ActiveView.Refresh
    gMxDoc.UpdateContents

End Sub


Public Function ReturnBMPSymbol(pBMPType As String, Optional TSFlag As Boolean) As IMultiLayerMarkerSymbol  ' IPictureMarkerSymbol
    
  
    '** define the necessary variables
    Dim mpMrkSym As IMultiLayerMarkerSymbol
    Set mpMrkSym = New MultiLayerMarkerSymbol
    
    '** Define base layer
    Dim pictBMPMrkSym1 As IPictureMarkerSymbol
    Set pictBMPMrkSym1 = New PictureMarkerSymbol
    '** Create the Markers and assign their properties.
    With pictBMPMrkSym1
       Set .Picture = ReturnBMPPicture(pBMPType)
      .Angle = 0
      .Size = 18
      .XOffset = 0
      .YOffset = 0
    End With

    '** Add the symbols in the order of bottommost to topmost
    mpMrkSym.AddLayer pictBMPMrkSym1
    
    Dim pRGBColor As IRgbColor
    Set pRGBColor = New RgbColor
    pRGBColor.RGB = RGB(255, 255, 255)
    
    '*** ASSESSMENT POINT LAYER
    If (Right(Trim(pBMPType), 1) = "X") Then
        '** Define layer for assessment pont
        Dim pictBMPMrkSym2 As IPictureMarkerSymbol
        Set pictBMPMrkSym2 = New PictureMarkerSymbol
        '** Create the Markers and assign their properties.
        With pictBMPMrkSym2
           Set .Picture = LoadResPicture("EditAssess", vbResBitmap)
          .Angle = 0
          .Size = 15
          .XOffset = 7
          .YOffset = -5
          .BitmapTransparencyColor = pRGBColor
        End With
        '*** Add the second layer
        mpMrkSym.AddLayer pictBMPMrkSym2
    End If
    
    '*** TIME SERIES LAYER
    If (TSFlag = True) Then
        '** Define layer for assessment pont
        Dim pictBMPMrkSym3 As IPictureMarkerSymbol
        Set pictBMPMrkSym3 = New PictureMarkerSymbol
        '** Create the Markers and assign their properties.
        With pictBMPMrkSym3
           Set .Picture = LoadResPicture("TS", vbResBitmap)
          .Angle = 0
          .Size = 12
          .XOffset = 9
          .YOffset = 8
          .BitmapTransparencyColor = pRGBColor
        End With
        '*** Add the second layer
        mpMrkSym.AddLayer pictBMPMrkSym3
    End If

    'return the symbol
    Set ReturnBMPSymbol = mpMrkSym
    
End Function

Public Function ReturnBMPPicture(pBMPType As String) As IPictureDisp
   
   Dim pPicture As IPictureDisp
    Select Case pBMPType:
        Case "BioRetentionBasin":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpBR", vbResBitmap)
        Case "WetPond":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpWP", vbResBitmap)
        Case "Cistern":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpC", vbResBitmap)
        Case "DryPond":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpDP", vbResBitmap)
        Case "InfiltrationTrench":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpIT", vbResBitmap)
        Case "RainBarrel":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpRB", vbResBitmap)
        Case "GreenRoof":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpGR", vbResBitmap)
        Case "PorousPavement":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpPP", vbResBitmap)
        Case "VegetativeSwale":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpVS", vbResBitmap)
        Case "Regulator":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpREG", vbResBitmap)
        Case "VirtualOutlet":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("VirtualOutlet", vbResBitmap)
        Case "Junction":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("Junction", vbResBitmap)
        Case "Aggregate":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("AggregateBMP", vbResBitmap)
            
        'New Types if an assessment point
        Case "BioRetentionBasinX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpBRX", vbResBitmap)
        Case "WetPondX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpWPX", vbResBitmap)
        Case "CisternX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpCX", vbResBitmap)
        Case "DryPondX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpDPX", vbResBitmap)
        Case "InfiltrationTrenchX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpITX", vbResBitmap)
        Case "RainBarrelX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpRBX", vbResBitmap)
        Case "GreenRoofX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpGRX", vbResBitmap)
        Case "PorousPavementX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpPPX", vbResBitmap)
        Case "VegetativeSwaleX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpVSX", vbResBitmap)
        Case "RegulatorX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("bmpREGX", vbResBitmap)
        Case "VirtualOutletX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("VirtualOutletX", vbResBitmap)
        Case "JunctionX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("JunctionX", vbResBitmap)
        Case "AggregateX":
            '** Create the Markers and assign their properties.
            Set pPicture = LoadResPicture("AggregateBMPX", vbResBitmap)
       Case Else:
            Set pPicture = LoadResPicture("bmpX", vbResBitmap)
    End Select
    'Return the BMP Picture
    Set ReturnBMPPicture = pPicture
End Function



'*******************************************************************************
'Subroutine : AddVirtualPointsForSchematic
'Purpose    : Adds virtual outlets to snap points for creating the schematic
'Arguments  :
'Author     : Sabu Paul
'*******************************************************************************
Public Sub AddVirtualPointsForSchematic()

    'Get SnapPoints feature layer
    Dim pSnapPointsFLayer As IFeatureLayer
    Set pSnapPointsFLayer = GetInputFeatureLayer("SnapPoints")
    If (pSnapPointsFLayer Is Nothing) Then
        MsgBox "SnapPoints feature layer not found."
        GoTo CleanUp
    End If

    Dim pAssessmentFClass As IFeatureClass
    Set pAssessmentFClass = pSnapPointsFLayer.FeatureClass
    
    'Add the Virtual Outlet to the Snappoint
    'if the point is not alread present in the layer
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    If (pBMPFLayer Is Nothing) Then
        MsgBox "BMPs layer required to continue !. "
        Exit Sub
    End If
    'Select the Virtual Outlets and add it to snappoints -- Sabu Paul: Aug 15, 2004
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "TYPE = 'VirtualOutlet'"
    
    'Select the Virtual Outlets and add it to snappoints -- Sabu Paul: Aug 15, 2004
    Dim pQueryFilter2 As IQueryFilter
    Set pQueryFilter2 = New QueryFilter
    
    
    Dim pBMPFeatureClass As IFeatureClass
    Set pBMPFeatureClass = pBMPFLayer.FeatureClass
    Dim pBMPFeature As IFeature
    Dim pFeature As IFeature
    Dim pCurId As Integer
    Dim pCurPID As Integer
    Dim pCurGCode As Integer
    
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pBMPFeatureClass.Search(pQueryFilter, False)

    Set pBMPFeature = pFeatureCursor.NextFeature
    Do Until pBMPFeature Is Nothing
        pCurId = pBMPFeature.value(pBMPFeatureClass.FindField("ID"))
        
        pQueryFilter2.WhereClause = "GRID_CODE = " & pCurId

        If pAssessmentFClass.FeatureCount(pQueryFilter2) = 0 Then
            Set pFeature = pAssessmentFClass.CreateFeature
            Set pFeature.Shape = pBMPFeature.Shape
            
            pCurPID = pAssessmentFClass.FeatureCount(Nothing)
            pCurGCode = pCurId
            pFeature.value(pAssessmentFClass.FindField("PointID")) = pCurPID
            pFeature.value(pAssessmentFClass.FindField("GRID_CODE")) = pCurGCode
            pFeature.value(pAssessmentFClass.FindField("ID")) = pCurId
            pFeature.value(pAssessmentFClass.FindField("DSID")) = 0
            pFeature.Store
        End If
        Set pBMPFeature = pFeatureCursor.NextFeature
    Loop

CleanUp:
    Set pBMPFeature = Nothing
    Set pFeature = Nothing
    Set pQueryFilter = Nothing
    Set pQueryFilter2 = Nothing
    Set pBMPFeatureClass = Nothing
    Set pAssessmentFClass = Nothing
    Set pSnapPointsFLayer = Nothing
    Set pBMPFLayer = Nothing
End Sub



Public Sub RenderVFSFeatureLayer()
On Error GoTo ShowError

     '** Make the renderer
     Dim pRender As IUniqueValueRenderer
     Set pRender = New UniqueValueRenderer
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "TYPE2"
     pRender.UseDefaultSymbol = False
     
     Dim pFeatureLayer As IFeatureLayer
     Set pFeatureLayer = GetInputFeatureLayer("VFS")
     If (pFeatureLayer Is Nothing) Then
        Exit Sub
     End If
     Dim pFeatureclass As IFeatureClass
     Set pFeatureclass = pFeatureLayer.FeatureClass
     Dim pFeatureCursor As IFeatureCursor
     Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
     Dim pFeature As IFeature
     Set pFeature = pFeatureCursor.NextFeature
     Dim iID As Long
     iID = pFeatureclass.FindField("ID")
     Dim iType As Long
     iType = pFeatureclass.FindField("TYPE2")
     Dim iLabel As Long
     iLabel = pFeatureclass.FindField("LABEL")
     Dim pVFSType As String
     Dim pVFSLabel As String
     Dim ValFound As Boolean
     Dim pRouteLineSymbol As ICartographicLineSymbol
     Dim uh As Integer
     
     Do While Not pFeature Is Nothing
         'get values for feature
         pVFSType = pFeature.value(iType)
         pVFSLabel = pFeature.value(iLabel)
         
         Set pRouteLineSymbol = ReturnVFSRouteSymbol(pVFSType)
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.value(uh) = pVFSType Then
             ValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             pRender.AddValue pVFSType, "TYPE2", pRouteLineSymbol
             pRender.Label(pVFSType) = pVFSLabel
             pRender.Symbol(pVFSType) = pRouteLineSymbol
         End If
        Set pFeature = pFeatureCursor.NextFeature
     Loop
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     Dim pLyr As IGeoFeatureLayer
     Set pLyr = pFeatureLayer
     pRender.ColorScheme = "Custom"
     pRender.fieldType(0) = False
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = "TYPE2"
 
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
     aLELayerProps.Expression = "[LABEL]"
     Dim pTextSymbol As ITextSymbol
     Set pTextSymbol = New TextSymbol
     pTextSymbol.Size = 8
     Dim pColor As IRgbColor
     Set pColor = New RgbColor
     pColor.RGB = vbBlack
     pTextSymbol.Color = pColor
     Set aLELayerProps.Symbol = pTextSymbol
     ' assign it to the layer's AnnotateLayerPropertiesCollection
     pAnnoLayerPropsColl.add aLELayerProps
     'get the BasicOverposterLayerProperties
     Dim pBasicOverposterLayerProps As IBasicOverposterLayerProperties
     Set pBasicOverposterLayerProps = aLELayerProps.BasicOverposterLayerProperties
     pBasicOverposterLayerProps.NumLabelsOption = esriOneLabelPerShape
     
     '** This makes the layer properties symbology tab show the correct interface.
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     '** Draw the map
     gMxDoc.ActiveView.Refresh
     
     GoTo CleanUp
ShowError:
    MsgBox "RenderVFSFeatureLayer: " & Err.description
CleanUp:

End Sub


'* Subroutine to define conduit network for BMPs
Public Function ReturnVFSRouteSymbol(pVFSBankType As String) As ICartographicLineSymbol

On Error GoTo ShowError
    
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    Dim pOffset As Integer
    Select Case pVFSBankType
        Case "VFS_L":   'left bank
            pColor.RGB = RGB(56, 168, 0)
            pOffset = 4
        Case "VFS_R":   'right bank
            pColor.RGB = RGB(168, 168, 0)
            pOffset = -4
        Case Else
            pColor.RGB = vbBlack
    End Select
    
    'Define a 5 pt wide line with 4 pt offset
    Dim theproperties As ILineProperties
    Dim theline As ICartographicLineSymbol
    Set theline = New CartographicLineSymbol
    theline.Color = pColor
    theline.Width = 5
    Set theproperties = theline
    theproperties.Offset = pOffset
        
    'Return the symbol
    Set ReturnVFSRouteSymbol = theline
    
    GoTo CleanUp
ShowError:
    MsgBox "ReturnVFSRouteSymbol: " & Err.description
CleanUp:
    Set pColor = Nothing
    Set theproperties = Nothing
    Set theline = Nothing
End Function
