Attribute VB_Name = "ModuleBMPDetails"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleBMPDetails
'   Purpose:     Functions and subroutine to store and retrieve details for BMP templates
'                and specific BMP sites
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  08/../2004 - Mira Chokshi and Sabu Paul
'                Modified: 08/19/2004 - Sabu Paul added comments to project
'
'******************************************************************************

Option Explicit
Option Base 0
Public gAggBMPFlagDict As Scripting.Dictionary

'*******************************************************************************
'Subroutine : AddBMPInformation
'Purpose    : Add the detailed information of the new BMP site
'             Into the BMPDetail table
'Arguments  : Data dictionary containing the BMP properties
'Author     : Mira Chokshi
'History    : 08/10/2004 - Sabu Paul
'*******************************************************************************


Public Sub AddBMPInformation(BMPDetailDict As Dictionary)

On Error GoTo ErrorHandler:
    
    If BMPDetailDict Is Nothing Then
        'MsgBox "BMPDetailDict is nothing"
        Exit Sub
    End If
    
    Dim pBMPDetailTable As iTable
    
    If gBMPTypeToolbox = "Aggregate" Then
        Set pBMPDetailTable = GetInputDataTable("AgBMPDetail")
        If (pBMPDetailTable Is Nothing) Then
            Set pBMPDetailTable = CreatePropertiesTableDBF("AgBMPDetail")
            AddTableToMap pBMPDetailTable
        End If
    Else
        Set pBMPDetailTable = GetInputDataTable("BMPDetail")
        If (pBMPDetailTable Is Nothing) Then
            Set pBMPDetailTable = CreatePropertiesTableDBF("BMPDetail")
            AddTableToMap pBMPDetailTable
        End If
    End If
    
    Dim pIDindex As Long
    pIDindex = pBMPDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    'Iterate over the entire dictionary
    Dim pRow As iRow
    Dim pBMPKeys
    Dim i As Integer
    'Changed to include BMP Type - Jan 2009
'    Dim pCategories
'    pCategories = Array("On-Site Interception", "On-Site Treatment", "Routing Attenuation", "Regional Storage/Treatment")
        
    If gBMPTypeToolbox = "Aggregate" Then
        Dim catId As Integer
        Dim catKey As String
        Dim bmpId As Integer
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "PropName = 'Type'"
        bmpId = pBMPDetailTable.RowCount(pQueryFilter)
        
        
        Dim catDict As Scripting.Dictionary
        'For catId = 0 To UBound(pCategories)
        For catId = 0 To BMPDetailDict.Count - 1
            'If BMPDetailDict.Exists(pCategories(catId)) Then
            catKey = BMPDetailDict.keys(catId)
            If BMPDetailDict.Exists(catKey) Then
                If TypeOf BMPDetailDict.Item(catKey) Is Scripting.Dictionary Then
                    bmpId = bmpId + 1
                    'Set catDict = BMPDetailDict.Item(pCategories(catId))
                    Set catDict = BMPDetailDict.Item(catKey)
                    If Not catDict Is Nothing Then
                        catDict.Item("BMPID") = gNewBMPId
                        pBMPKeys = catDict.keys
                        For i = 0 To (catDict.Count - 1)
                            pPropertyName = pBMPKeys(i)
                            pPropertyValue = catDict.Item(pPropertyName)
                            Set pRow = pBMPDetailTable.CreateRow
                            pRow.value(pIDindex) = bmpId 'catId + 1
                            pRow.value(pPropNameIndex) = pPropertyName
                            pRow.value(pPropValueIndex) = pPropertyValue
                            pRow.Store
                        Next
                    End If
                End If
            End If
        Next
    Else
        pBMPKeys = BMPDetailDict.keys
        For i = 0 To (BMPDetailDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            pPropertyValue = BMPDetailDict.Item(pPropertyName)
            Set pRow = pBMPDetailTable.CreateRow
            pRow.value(pIDindex) = gNewBMPId
            pRow.value(pPropNameIndex) = pPropertyName
            pRow.value(pPropValueIndex) = pPropertyValue
            pRow.Store
        Next
    End If
    GoTo CleanUp
ErrorHandler:
    MsgBox "AddBMPInformation :", Err.description
CleanUp:
    Set pBMPDetailTable = Nothing
    Set pRow = Nothing
End Sub

'*******************************************************************************
'Subroutine : CreateBMPDefaultsDBF
'Purpose    : Creates a DBASE file in the project temp directory to store the
'             details of BMP templates
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Sabu Paul
'History    :
'*******************************************************************************

Public Function CreateBMPDefaultsDBF(pFileName As String) As iTable

On Error GoTo ShowError

  'Delete data table from temp folder
  DeleteDataTable gMapTempFolder, pFileName
  
  ' Open the Workspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  Dim pFWS As IFeatureWorkspace
  Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As esriGeoDatabase.IField
  Dim pFields As esriGeoDatabase.IFields

    ' create the fields used by our object
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 3

    'Create ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    

    'Create PropName Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "PropName"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create PropValue Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "PropValue"
        .Type = esriFieldTypeString
        .Length = 250
    End With
    Set pFieldsEdit.Field(2) = pField

  Set CreateBMPDefaultsDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateBMPDefaultsDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function


'*******************************************************************************
'Subroutine : CreatePropertiesTableDBF
'Purpose    : Creates a DBASE file to store the BMP details in the project temp directory
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Mira Chokshi
'History    : 08/10/2004 - Sabu Paul - Modified the field width
'*******************************************************************************
Public Function CreatePropertiesTableDBF(pFileName As String) As iTable
On Error GoTo ShowError

  'Delete data table from temp folder
  DeleteDataTable gMapTempFolder, pFileName
  
  ' Open the Workspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  Dim pFWS As IFeatureWorkspace
  Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As esriGeoDatabase.IField
  Dim pFields As esriGeoDatabase.IFields

  ' if a fields collection is not passed in then create one
    ' create the fields used by our object
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 3

    'Create BRANCH Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField

    'Create SEGMENTNAME Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "PropName"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create FILE Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "PropValue"
        .Type = esriFieldTypeString
        .Length = 250
    End With
    Set pFieldsEdit.Field(2) = pField

  Set CreatePropertiesTableDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreatePropertiesTableDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function

'*******************************************************************************
'Subroutine : CreatePropertiesTableDBF
'Purpose    : Creates a DBASE file to store the BMP details in the project temp directory
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Mira Chokshi
'History    : 08/10/2004 - Sabu Paul - Modified the field width
'*******************************************************************************
Public Function CreatePollutantsTableDBF(pFileName As String) As iTable
On Error GoTo ShowError

  'Delete data table from temp folder
  DeleteDataTable gMapTempFolder, pFileName
  
  ' Open the Workspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  Dim pFWS As IFeatureWorkspace
  Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As esriGeoDatabase.IField
  Dim pFields As esriGeoDatabase.IFields

  ' if a fields collection is not passed in then create one
    ' create the fields used by our object
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 8

    'Create BRANCH Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField

    'Create SEGMENTNAME Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Name"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create FILE Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Multiplier"
        .Type = esriFieldTypeDouble
        .Length = 250
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create FILE Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Sediment"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(3) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SedAssoc"
        .Type = esriFieldTypeInteger
        .Length = 5
    End With
    Set pFieldsEdit.Field(4) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SandFrac"
        .Type = esriFieldTypeSingle
        .Length = 6
        .Precision = 5
        .Scale = 3
    End With
    Set pFieldsEdit.Field(5) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SiltFrac"
        .Type = esriFieldTypeSingle
        .Length = 6
        .Precision = 5
        .Scale = 3
    End With
    Set pFieldsEdit.Field(6) = pField
    
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ClayFrac"
        .Type = esriFieldTypeSingle
        .Length = 6
        .Precision = 5
        .Scale = 3
    End With
    Set pFieldsEdit.Field(7) = pField

  Set CreatePollutantsTableDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreatePollutantsTableDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function




'*******************************************************************************
'Subroutine : CreateBMPRoutingDBF
'Purpose    : Creates a DBASE file in the project temp directory to store the
'             routing information
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Mira Chokshi
'History    :
'*******************************************************************************
Public Function CreateBMPRoutingDBF(pFileName As String) As iTable

On Error GoTo ShowError

  'Delete data table from temp folder
  DeleteDataTable gMapTempFolder, pFileName
  
  ' Open the Workspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  Dim pFWS As IFeatureWorkspace
  Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As esriGeoDatabase.IField
  Dim pFields As esriGeoDatabase.IFields

  ' if a fields collection is not passed in then create one
    ' create the fields used by our object
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 4

    'Create ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    
    'Create BRANCH Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "OutletType"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(1) = pField

    'Create SEGMENTNAME Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "DSID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create FILE Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "TypeDesc"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(3) = pField

  Set CreateBMPRoutingDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateBMPRoutingDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function
'*******************************************************************************
'Subroutine : CreateBMPTypesDBF
'Purpose    : Creates a DBASE file in the project temp directory to store the
'             different types of BMP templates
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function CreateBMPTypesDBF(pFileName As String) As iTable

On Error GoTo ShowError

  'Delete data table from temp folder
  DeleteDataTable gMapTempFolder, pFileName
  
  ' Open the Workspace
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  Dim pFWS As IFeatureWorkspace
  Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

  Dim pFieldsEdit As IFieldsEdit
  Dim pFieldEdit As IFieldEdit
  Dim pField As esriGeoDatabase.IField
  Dim pFields As esriGeoDatabase.IFields

  ' if a fields collection is not passed in then create one
    ' create the fields used by our object
    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 4

    'Create ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    
    'Create BRANCH Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Name"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(1) = pField

    'Create SEGMENTNAME Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Type"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create FILE Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Class"
        .Type = esriFieldTypeString
        .Length = 30
    End With
    Set pFieldsEdit.Field(3) = pField

  Set CreateBMPTypesDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateBMPTypesDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function

'*******************************************************************************
'Subroutine : DeleteSelectedBMP
'Purpose    : Deletes the BMP sites and the corresponding details from the tables
'             Also modifies the BMP ids
'Note       :
'Arguments  :
'Author     : Sabu Paul
'History    : 08/10/2004 - Sabu Paul
'*******************************************************************************
Public Sub DeleteSelectedBMP(pSelectedBMPId As Integer)
On Error GoTo ShowError
    
    Dim pPourPointFLayer As IFeatureLayer
    Set pPourPointFLayer = GetInputFeatureLayer("BMPs")
    Dim pPourPointFClass As IFeatureClass
    If Not (pPourPointFLayer Is Nothing) Then
        Set pPourPointFClass = pPourPointFLayer.FeatureClass
    Else
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If
    Dim iBmpIdFld As Long
    iBmpIdFld = pPourPointFClass.FindField("ID")
    Dim pBMPID As Integer
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
       
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
    Dim pIDindex As Long
    If Not pBMPDetailTable Is Nothing Then pIDindex = pBMPDetailTable.FindField("ID")
    
    Dim pBMPNetworkTable As iTable
    Set pBMPNetworkTable = GetInputDataTable("BMPNetwork")
    
    Dim pNetIDindex As Long
    Dim pNetDSIDindex As Long
    
    If Not pBMPNetworkTable Is Nothing Then
        pNetIDindex = pBMPNetworkTable.FindField("ID")
        pNetDSIDindex = pBMPNetworkTable.FindField("DSID")
    End If
    
'    Dim pDecayFactTable As iTable
'    Set pDecayFactTable = GetInputDataTable("DecayFact")
'
'    Dim pPctRemovalTable As iTable
'    Set pPctRemovalTable = GetInputDataTable("PctRemoval")
        
    Dim pAggDetailTable As iTable
    Set pAggDetailTable = GetInputDataTable("AgBMPDetail")
    
    Dim pAggLuTable As iTable
    Set pAggLuTable = GetInputDataTable("AgLuDistribution")
    
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
    
    Dim aggIdField As Integer, aggBmpID As Integer
    
    pQueryFilter.WhereClause = "ID = " & pSelectedBMPId
    Set pFeatureCursor = pPourPointFClass.Search(pQueryFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    If Not pFeature Is Nothing Then
        pBMPID = pFeature.value(iBmpIdFld)
        '********* DELETE RECORDS FROM ALL FEATURE LAYERS
        'Delete the features from BMPs feature layer with ID = deleted BMP
        'call the subroutine to start editing watershed related layer editing
        Call StartEditingFeatureLayer("BMPs")
        Set pDeleteSet = New esriSystem.Set
        pDeleteSet.add pFeature
        pDeleteSet.Reset
        Set pFeatureEdit = pDeleteSet.Next
        If Not pFeatureEdit Is Nothing Then
          pFeatureEdit.DeleteSet pDeleteSet
        End If
        Call StopEditingFeatureLayer
        
        'Delete all records from the BMPDETAIL table with ID = deleted BMP
        pQueryFilter1.WhereClause = "ID = " & pBMPID
        If Not (pBMPDetailTable Is Nothing) Then
            pBMPDetailTable.DeleteSearchedRows pQueryFilter1
        End If
        
        'Include the option to delete AgBMPDetail
        If Not pAggDetailTable Is Nothing Then
            pQueryFilter1.WhereClause = "PropName='BMPID' And PropValue = '" & pBMPID & "'"
            If pAggDetailTable.RowCount(pQueryFilter1) > 0 Then
                Set pTabCursor = pAggDetailTable.Search(pQueryFilter1, False)
                aggIdField = pAggDetailTable.FindField("ID")
                Set pRow = pTabCursor.NextRow
                Do While Not (pRow Is Nothing)
                    aggBmpID = pRow.value(aggIdField)
                    pQueryFilter1.WhereClause = "ID = " & aggBmpID
                    pAggDetailTable.DeleteSearchedRows pQueryFilter1
                    Set pRow = pTabCursor.NextRow
                Loop
            End If
            
'            'Deduct one from all BMPID values greater than pBMPID
'            pQueryFilter1.WhereClause = "PropName='BMPID' "
'            If pAggDetailTable.RowCount(pQueryFilter1) > 0 Then
'                Set pTabCursor = pAggDetailTable.Update(pQueryFilter1, False)
'                Set pRow = pTabCursor.NextRow
'                Do While Not (pRow Is Nothing)
'                    aggBmpID = CInt(pRow.value(pAggDetailTable.FindField("PropValue")))
'                    If aggBmpID > pBMPID Then
'                        pRow.value(pAggDetailTable.FindField("PropValue")) = CStr(aggBmpID - 1)
'                        pRow.Store
'                    End If
'                    Set pRow = pTabCursor.NextRow
'                Loop
'            End If

            'Delete records from AggLu
            pQueryFilter1.WhereClause = "BMPID = " & pBMPID
            If Not (pAggLuTable Is Nothing) Then
                If pAggLuTable.RowCount(pQueryFilter1) > 0 Then
                    pAggLuTable.DeleteSearchedRows pQueryFilter1
                End If
                
'                pQueryFilter1.WhereClause = "BMPID > " & pBMPID
'                Set pTabCursor = pAggLuTable.Update(pQueryFilter1, False)
'                Set pRow = pTabCursor.NextRow
'                Do While Not (pRow Is Nothing)
'                    aggBmpID = pRow.value(pAggLuTable.FindField("BMPID"))
'                    pRow.value(pAggLuTable.FindField("BMPID")) = aggBmpID - 1
'                    pRow.Store
'                Loop

            End If
            
        End If
                
        
        If Not (pConduitsFClass Is Nothing) Then
            'Delete the features from Conduits feature layer with ID = deleted BMP
            Call StartEditingFeatureLayer("Conduits")
            Set pFeatureCursor1 = Nothing
            Set pFeature1 = Nothing
            pQueryFilter1.WhereClause = "CFROM = " & pBMPID & " OR CTO = " & pBMPID
            Set pFeatureCursor1 = pConduitsFClass.Search(pQueryFilter1, False)
            Set pFeature1 = pFeatureCursor1.NextFeature
            Set pDeleteSet = New esriSystem.Set
            Do While Not (pFeature1 Is Nothing)
                'Delete all records from the BMPDetail table with ID = deleted conduit
                pQueryFilter1.WhereClause = "ID = " & pFeature1.value(iConduitIDFld)
                If Not (pBMPDetailTable Is Nothing) Then
                    pBMPDetailTable.DeleteSearchedRows pQueryFilter1
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
            
            Call StopEditingFeatureLayer
        End If
        
        
        'Delete the features from BasinRouting feature layer with ID = deleted BMP
        If (Not pWatershedFClass Is Nothing) And (Not pBasinRoutingFClass Is Nothing) Then
            Set pFeatureCursor1 = Nothing
            Set pFeature1 = Nothing
            pQueryFilter1.WhereClause = "BMPID = " & pBMPID
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
        'Delete the details from the BMPNETWORK table with ID = deleted BMP
        pQueryFilter1.WhereClause = "ID = " & pBMPID
        If Not (pBMPNetworkTable Is Nothing) Then
            pBMPNetworkTable.DeleteSearchedRows pQueryFilter1
        End If
        pQueryFilter1.WhereClause = "DSID = " & pBMPID
        Set pTabCursor = pBMPNetworkTable.Update(pQueryFilter1, False)
        Set pRow = pTabCursor.NextRow
        Do While Not (pRow Is Nothing)
            pRow.value(pNetDSIDindex) = 0
            pRow.Store
            Set pRow = pTabCursor.NextRow
        Loop
        'Delete the details from DECAYFACT table with ID = deleted BMP
''        pQueryFilter1.WhereClause = "BMPID = " & pBMPID
''        If Not (pDecayFactTable Is Nothing) Then
''            pDecayFactTable.DeleteSearchedRows pQueryFilter1
''        End If
''        'Delete the details from PCTREMOVAL table with ID = deleted BMP
''        pQueryFilter1.WhereClause = "BMPID = " & pBMPID
''        If Not (pPctRemovalTable Is Nothing) Then
''            pPctRemovalTable.DeleteSearchedRows pQueryFilter1
''        End If
        
        Set pTabCursor = Nothing
        Set pRow = Nothing
    End If
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing

    gMxDoc.ActiveView.ContentsChanged
    gMxDoc.ActiveView.Refresh
    
    GoTo CleanUp
    
ShowError:
    MsgBox "DeleteSelectedBMP:  " & vbTab & Err.Number & vbTab & Err.description
CleanUp:
    Set pPourPointFLayer = Nothing
    Set pPourPointFClass = Nothing
    Set pConduitsFLayer = Nothing
    Set pConduitsFClass = Nothing
    Set pWatershedFLayer = Nothing
    Set pWatershedFClass = Nothing
    Set pBasinRoutingFLayer = Nothing
    Set pBasinRoutingFClass = Nothing
    Set pBMPDetailTable = Nothing
    Set pBMPNetworkTable = Nothing
'    Set pDecayFactTable = Nothing
'    Set pPctRemovalTable = Nothing
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

Public Sub UpdateBMPRelatedDatasets()
On Error GoTo ShowError
    
    Dim pPourPointFLayer As IFeatureLayer
    Set pPourPointFLayer = GetInputFeatureLayer("BMPs")
    Dim pPourPointFClass As IFeatureClass
    If Not (pPourPointFLayer Is Nothing) Then
        Set pPourPointFClass = pPourPointFLayer.FeatureClass
    Else
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If
    Dim iBmpIdFld As Long
    iBmpIdFld = pPourPointFClass.FindField("ID")
    Dim pBMPID As Integer
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
       
    Dim pVFSLayer As IFeatureLayer
    Set pVFSLayer = GetInputFeatureLayer("VFS")
    Dim pVFSClass As IFeatureClass
    If (Not pVFSLayer Is Nothing) Then
        Set pVFSClass = pVFSLayer.FeatureClass
    End If
    Dim iVfsIdFld As Long
    Dim pVFSID As Integer
    Dim pVFSType As String
    
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
    Dim pIDindex As Long
    If Not pBMPDetailTable Is Nothing Then pIDindex = pBMPDetailTable.FindField("ID")
    
    Dim pVFSDetailTable As iTable
    Set pVFSDetailTable = GetInputDataTable("VFSDetail")
    
    Dim pBMPNetworkTable As iTable
    Set pBMPNetworkTable = GetInputDataTable("BMPNetwork")
    Dim pNetIDindex As Long
    Dim pNetDSIDindex As Long
    If Not pBMPNetworkTable Is Nothing Then
        pNetIDindex = pBMPNetworkTable.FindField("ID")
        pNetDSIDindex = pBMPNetworkTable.FindField("DSID")
    End If
                        
    Dim pOptimizationDetail As iTable
    Set pOptimizationDetail = GetInputDataTable("OptimizationDetail")
    
    Dim pTabCursor As ICursor
    Dim pRow As iRow
    
    Dim pIDArray() As Integer
    Dim pIdCount As Integer
    Dim pIdIncr As Integer
    pIdCount = 0

    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter

    '******* Update the IDs (renumber) in feature layers and tables
   'sort the elements in the descending order - BMPs
    pIdCount = 0
    Set pFeatureCursor = pPourPointFClass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    iBmpIdFld = pFeatureCursor.FindField("ID")
    Do While Not (pFeature Is Nothing)
        ReDim Preserve pIDArray(pIdCount)
        pIDArray(pIdCount) = pFeature.value(iBmpIdFld)
        pIdCount = pIdCount + 1
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    
    'sort the elements in the descending order - Conduits
    If Not (pConduitsFClass Is Nothing) Then
        Set pFeatureCursor = pConduitsFClass.Search(Nothing, True)
        Set pFeature = pFeatureCursor.NextFeature
        Do While Not (pFeature Is Nothing)
            ReDim Preserve pIDArray(pIdCount)
            pIDArray(pIdCount) = pFeature.value(iConduitIDFld)
            pIdCount = pIdCount + 1
            Set pFeature = pFeatureCursor.NextFeature
        Loop
    End If
   Set pFeatureCursor = Nothing
   Set pFeature = Nothing
    
    'sort the elements in the descending order - VFS
    If Not (pVFSClass Is Nothing) Then
        iVfsIdFld = pVFSClass.FindField("ID")
        Set pFeatureCursor = pVFSClass.Search(Nothing, True)
        Set pFeature = pFeatureCursor.NextFeature
        Do While Not (pFeature Is Nothing)
            ReDim Preserve pIDArray(pIdCount)
            pIDArray(pIdCount) = pFeature.value(iVfsIdFld)
            pIdCount = pIdCount + 1
            Set pFeature = pFeatureCursor.NextFeature
        Loop
    End If
    
    Dim pAggDetailTable As iTable
    Set pAggDetailTable = GetInputDataTable("AgBMPDetail")
    
    Dim pAggLuTable As iTable
    Set pAggLuTable = GetInputDataTable("AgLuDistribution")

    If (pIdCount > 0) Then
        'Sort the bmp/conduit id array
        Call BubbleSort(pIDArray, , False)
    
        Dim pSequenceBMPId As Integer
        Dim pCurBMPId As Integer
        Dim pBMPIdInArray As Integer
    
        'All bmps/conduits must be re-numbered starting from 1
        For pSequenceBMPId = LBound(pIDArray) To UBound(pIDArray)
            pBMPIdInArray = pIDArray(pSequenceBMPId)
            pCurBMPId = pSequenceBMPId + 1 'current bmp id should be the sequence bmp id
            'If the next value in the sorted bmp id array is not the same as next sequence id,
            'change the value in bmp/conduits feature layer and related tables
    
            If (pBMPIdInArray <> pCurBMPId) Then
                    'Modify the IDs in the BMP Feature layer
                    pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
                    Set pFeatureCursor = pPourPointFClass.Update(pQueryFilter, False)
                    Set pFeature = pFeatureCursor.NextFeature
                    Do While Not (pFeature Is Nothing)
                        pFeature.value(iBmpIdFld) = pCurBMPId
                        pFeature.Store
                        Set pFeature = pFeatureCursor.NextFeature
                    Loop
                    'Modify the IDs in the Conduits Feature layer
                    If Not (pConduitsFClass Is Nothing) Then
                        'Update the ID of conduits feature layer
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
                        Set pFeatureCursor = pConduitsFClass.Update(pQueryFilter, False)
                        Set pFeature = pFeatureCursor.NextFeature
                        Do While Not (pFeature Is Nothing)
                            pFeature.value(iConduitIDFld) = pCurBMPId
                            pFeature.Store
                            Set pFeature = pFeatureCursor.NextFeature
                        Loop
                        'Update the FROM value of the conduits feature layer
                        pQueryFilter.WhereClause = "CFROM = " & pBMPIdInArray
                        Set pFeatureCursor = pConduitsFClass.Update(pQueryFilter, False)
                        Set pFeature = pFeatureCursor.NextFeature
                        Do While Not (pFeature Is Nothing)
                            pFeature.value(iConduitFROMFld) = pCurBMPId
                            pFeature.Store
                            Set pFeature = pFeatureCursor.NextFeature
                        Loop
                        'Update the CTO value of the conduits feature layer
                        pQueryFilter.WhereClause = "CTO = " & pBMPIdInArray
                        Set pFeatureCursor = pConduitsFClass.Update(pQueryFilter, False)
                        Set pFeature = pFeatureCursor.NextFeature
                        Do While Not (pFeature Is Nothing)
                            pFeature.value(iConduitTOFld) = pCurBMPId
                            pFeature.Store
                            Set pFeature = pFeatureCursor.NextFeature
                        Loop
                    End If
                    '* Modify the IDs in VFS feature layer
                    If (Not pVFSClass Is Nothing) Then
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
                        Set pFeatureCursor = pVFSClass.Search(pQueryFilter, False)
                        Set pFeature = pFeatureCursor.NextFeature
                        Do While Not (pFeature Is Nothing)
                            pFeature.value(iVfsIdFld) = pCurBMPId
                            pFeature.Store
                            Set pFeature = pFeatureCursor.NextFeature
                        Loop
                    End If
                    'Modify the IDs in the BMP Detail Table
                    If Not pBMPDetailTable Is Nothing Then
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
    '                    Dim iBdIDFld As Long
    '                    iBdIDFld = pBMPDetailTable.FindField("ID")
                        Set pTabCursor = pBMPDetailTable.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            pRow.value(pIDindex) = pCurBMPId  'iBdIDFld
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                    'Modify the IDs in the VFS Detail Table
                    If (Not pVFSDetailTable Is Nothing) Then
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
                        Dim iVdIDFld As Long
                        iVdIDFld = pVFSDetailTable.FindField("ID")
                        Set pTabCursor = pVFSDetailTable.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            pRow.value(iVdIDFld) = pCurBMPId
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                    If Not (pBMPNetworkTable Is Nothing) Then
                        'Modify the IDs in the Routing Network Table
'                        Dim iBnIDFld As Long
'                        iBnIDFld = pBMPNetworkTable.FindField("ID")
'                        Dim iBnDSIDFld As Long
'                        iBnDSIDFld = pBMPNetworkTable.FindField("DSID")
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray & "OR DSID = " & pBMPIdInArray
                    
                        Set pTabCursor = pBMPNetworkTable.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            If (pRow.value(pNetIDindex) = pBMPIdInArray) Then
                                pRow.value(pNetIDindex) = pCurBMPId
                            End If
                            If (pRow.value(pNetDSIDindex) = pBMPIdInArray) Then
                                pRow.value(pNetDSIDindex) = pCurBMPId
                            End If
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                    
                    'Modify the IDs in the Optimization Detail Table
                    If Not (pOptimizationDetail Is Nothing) Then
                        pQueryFilter.WhereClause = "ID = " & pBMPIdInArray
                        Dim iOdIDFld As Long
                        iOdIDFld = pOptimizationDetail.FindField("ID")
                        Set pTabCursor = pOptimizationDetail.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            pRow.value(iOdIDFld) = pCurBMPId
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                    
                    'Modify the IDs in the AgBMPDetail
                    If Not pAggDetailTable Is Nothing Then
                        pQueryFilter.WhereClause = "PropName='BMPID' AND PropValue = '" & pBMPIdInArray & "'"
                        Set pTabCursor = pAggDetailTable.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            pRow.value(pAggDetailTable.FindField("PropValue")) = CStr(pCurBMPId)
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                    
                    'Modify the IDs in the AgLuDistribution
                    If Not pAggLuTable Is Nothing Then
                        pQueryFilter.WhereClause = "BMPID = " & pBMPIdInArray
                        Set pTabCursor = pAggLuTable.Update(pQueryFilter, False)
                        Set pRow = pTabCursor.NextRow
                        Do While Not (pRow Is Nothing)
                            pRow.value(pAggLuTable.FindField("BMPID")) = pCurBMPId
                            pRow.Store
                            Set pRow = pTabCursor.NextRow
                        Loop
                    End If
                                        
            'End of the condition for comparing sequence id and bmp id in sorted array
            End If
        Next pSequenceBMPId
    End If
    
    'Call subroutine to render the bmp layer
    gMxDoc.ActiveView.Refresh
    GoTo CleanUp
    
ShowError:
    MsgBox "UpdateBMPRelatedDatasets:  " & vbTab & Err.Number & vbTab & Err.description
CleanUp:
    Set pPourPointFLayer = Nothing
    Set pPourPointFClass = Nothing
    Set pConduitsFLayer = Nothing
    Set pConduitsFClass = Nothing
    Set pWatershedFLayer = Nothing
    Set pWatershedFClass = Nothing
    Set pBasinRoutingFLayer = Nothing
    Set pBasinRoutingFClass = Nothing
    Set pBMPDetailTable = Nothing
    Set pBMPNetworkTable = Nothing
    Set pOptimizationDetail = Nothing
    Set pTabCursor = Nothing
    Set pRow = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pQueryFilter = Nothing
End Sub

'*******************************************************************************
'Subroutine : EditBmpDetails
'Purpose    : Edits the BMP details
'Note       :
'Arguments  :
'Author     : Sabu Paul
'History    : 08/10/2004 - Sabu Paul
'*******************************************************************************
Public Sub EditBmpDetails(pBMPID As Integer, pBMPType As String)

On Error GoTo ShowError
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Dim pBMPDetailTable As iTable
    Dim pQueryFilter As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    gBMPTypeTag = "BMPOnMap"
    
    'Modify to handler aggregate BMPs
    If UCase(pBMPType) = "AGGREGATE" Then
        Load FrmAggBmpSelection
        FrmAggBmpSelection.Initialize_Form (pBMPID)
        FrmAggBmpSelection.Show vbModal
    
    ElseIf UCase(pBMPType) = "REGULATOR" Then
        Set pBmpDetailDict = GetBMPDetailDict(pBMPID)
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        ModuleBMPData.CallInitRoutines pBMPType, pBmpDetailDict

        Set pBMPDetailTable = GetInputDataTable("BMPDetail")

        Dim pIDindex As Long
        pIDindex = pBMPDetailTable.FindField("ID")
        Dim pPropNameIndex As Long
        pPropNameIndex = pBMPDetailTable.FindField("PropName")
        Dim pPropValueIndex As Long
        pPropValueIndex = pBMPDetailTable.FindField("PropValue")

        Dim pPropertyName As String
        Dim pPropertyValue As String
        Dim i As Integer
        Dim pBMPKeys
        Dim pSelRowCount As Long
        Dim pBMPName As String


        'gNewBMPType = pBmpDetailDict.Item("BMPType")
        'pBMPName = pBmpDetailDict.Item("BMPName")

        
        'Remove records for bmp which are not in the gbmpdetaildictionary
        If Not (gBMPDetailDict Is Nothing) Then
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "ID = " & pBMPID
            Set pCursor = pBMPDetailTable.Update(pQueryFilter, True)
            Set pRow = pCursor.NextRow
            Do While Not pRow Is Nothing
                If (Not gBMPDetailDict.Exists(pRow.value(pPropNameIndex))) Then
                    pCursor.DeleteRow
                End If
                Set pRow = pCursor.NextRow
            Loop
            Set pRow = Nothing
            Set pCursor = Nothing
        End If

        If Not (gBMPDetailDict Is Nothing) Then
            pBMPKeys = gBMPDetailDict.keys
            For i = 0 To (gBMPDetailDict.Count - 1)
                pPropertyName = pBMPKeys(i)
                pPropertyValue = gBMPDetailDict.Item(pPropertyName)
                Set pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = "ID = " & pBMPID & " AND PropName = '" & pPropertyName & "'"
                Set pCursor = pBMPDetailTable.Search(pQueryFilter, False)
                Set pRow = pCursor.NextRow
                If Not pRow Is Nothing Then
                    pRow.value(pPropValueIndex) = pPropertyValue
                    pRow.Store
                Else
                    'Create if the row is not already in the table
                    Set pRow = pBMPDetailTable.CreateRow
                    pRow.value(pIDindex) = pBMPID
                    pRow.value(pPropNameIndex) = pPropertyName
                    pRow.value(pPropValueIndex) = pPropertyValue
                    pRow.Store
                End If
            Next i
        End If
    Else
        Call InitBmpTypeCatDict
        gNewBMPId = pBMPID
        gBMPEditMode = True
        Set pBmpDetailDict = GetBMPDetailDict(pBMPID)
        
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        
        Load frmBMPDef
        frmBMPDef.Form_Initialize
        gBMPTypeTag = "BMPOnMap"
        'frmBMPDef.Initialize_Form ("BMPOnMap")
        SetupFrmBmpTypeDef pBMPType, pBmpDetailDict
        frmBMPDef.Show vbModal
                    
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "Edit BMP Details :" & Err.description

CleanUp:
    Set pBmpDetailDict = Nothing
    Set pBMPDetailTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub


'*******************************************************************************
'Subroutine : GetBMPDetailDict
'Purpose    : Gets the BMP details for individual BMPSite from BMPDetail
'Note       :
'Arguments  : Id of the BMP site
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetBMPDetailDict(bmpId As Integer, Optional bmpTableName As String) As Dictionary

On Error GoTo ErrorHandler
    Dim pBMPDefaultTable As iTable
    'bmpTableName is added to handle the aggbmpdetail table
    If bmpTableName = "" Then
        Set pBMPDefaultTable = GetInputDataTable("BMPDetail")
    Else
        Set pBMPDefaultTable = GetInputDataTable(bmpTableName)
    End If
    
    If (pBMPDefaultTable Is Nothing) Then
         MsgBox "No BMPDetail table in the map: Add the table and continue"
         Exit Function
    End If
    
    Dim pIDindex As Long
    pIDindex = pBMPDefaultTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDefaultTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDefaultTable.FindField("PropValue")
                
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pTmpBMPName As String
    Dim pTmpBMPID As Integer
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "ID = " & bmpId

    Dim pCursor As ICursor
    Set pCursor = pBMPDefaultTable.Search(pQueryFilter, False)
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = CreateObject("Scripting.Dictionary")
    
    Do While Not pRow Is Nothing
        pTmpBMPID = pRow.value(pIDindex)
        pPropertyName = pRow.value(pPropNameIndex)
        pPropertyValue = pRow.value(pPropValueIndex)
        If pPropertyName <> "ID" Then
            pBmpDetailDict.add pPropertyName, pPropertyValue
        End If
        Set pRow = pCursor.NextRow
    Loop
    Set GetBMPDetailDict = pBmpDetailDict

    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "GetBMPDetailDict :", Err.description
CleanUp:
    Set pBMPDefaultTable = Nothing
    Set pQueryFilter = Nothing
    Set pBmpDetailDict = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
End Function

'''*******************************************************************************
'''Subroutine : EditBmpEvaluationFactors
'''Purpose    : Edits the BMP Evaluation Factors for an assessment point
'''Note       :
'''Arguments  :
'''Author     : Sabu Paul
'''History    : 09/10/2004 - Sabu Paul
'''*******************************************************************************
''Public Sub EditBmpEvaluationFactors()
''
''On Error GoTo ShowError
''
''    InitializeMapDocument
''
''    Dim pPourPointFLayer As IFeatureLayer
''    Set pPourPointFLayer = GetInputFeatureLayer("BMPs")
''    Dim pPourPointFClass As IFeatureClass
''    If Not (pPourPointFLayer Is Nothing) Then
''        Set pPourPointFClass = pPourPointFLayer.FeatureClass
''    Else
''        Exit Sub
''    End If
''
''    Dim pFldID As Long
''    pFldID = pPourPointFClass.FindField("ID")
''
''    Dim pFldType As Long
''    pFldType = pPourPointFClass.FindField("Type")
''
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''
''    Dim pFeatureCursor As IFeatureCursor
''    Dim pFeature As IFeature
''
''    Dim pID As Integer
''    Dim pType As String
''
''    Dim pFeatureSelection As IFeatureSelection
''    Set pFeatureSelection = pPourPointFLayer 'curFeatureClass
''
''    Dim pSelectionSet As ISelectionSet
''    Set pSelectionSet = pFeatureSelection.SelectionSet
''
''    Dim pBMPDetailDict As Scripting.Dictionary
''
''    Dim pBMPDetailTable As iTable
''    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
''
''    Dim pPropertyName As String
''    Dim pPropertyValue As String
''    Dim i As Integer
''    'Iterate over the entire dictionary
''    Dim pRow As iRow
''    Dim pBMPKeys
''
''    Dim pIDindex As Long
''    pIDindex = pBMPDetailTable.FindField("ID")
''    Dim pPropNameIndex As Long
''    pPropNameIndex = pBMPDetailTable.FindField("PropName")
''    Dim pPropValueIndex As Long
''    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
''
''    Dim pBMPName As String
''    Dim pIsAssessPoint As Boolean
''    Dim pContinue As Variant
''
''    Dim pControl As Control
''
''    pContinue = vbNo
''    Dim pCursor As ICursor
''    Dim pSelRowCount As Long
''
''    If pSelectionSet.Count > 0 Then
''        pSelectionSet.Search Nothing, False, pFeatureCursor
''        Set pFeature = pFeatureCursor.NextFeature
''        Do Until pFeature Is Nothing
''            pID = pFeature.value(pFldID)
''            pType = pFeature.value(pFldType)
''            Set pBMPDetailDict = GetBMPDetailDict(pID)
''            pBMPKeys = pBMPDetailDict.Keys
''
''            If pBMPDetailDict.Exists("isAssessmentPoint") Then
''                pIsAssessPoint = CBool(pBMPDetailDict.Item("isAssessmentPoint"))
''            Else
''                pIsAssessPoint = False
''            End If
''            If pIsAssessPoint = False Then
''                pContinue = MsgBox("Site at BMPID = " & pID & " is not currently an assessment point! " & vbNewLine & _
''                        " Click Yes to add assessment point option or No to cancel. ", vbYesNo, "Make Assessment Point")
''            End If
''
''            Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
''
''            If pIsAssessPoint = True Then
''                'Get the current info from table and initialize form
''                ModuleBMPData.LoadEvaluationFactors pBMPDetailDict
''                frmAssessPt.Show vbModal
''            ElseIf pContinue = vbYes Then
''                'just open the form and get the info
''                gBMPDetailDict.Add "isAssessmentPoint", "True"
''                ModuleBMPData.InitEvaluationFactors
''                frmAssessPt.Show vbModal
''                'Modify the Type2 field in the BMP layer
''                pFeature.value(pPourPointFClass.FindField("Type2")) = pType & "X"
''                pFeature.Store
''                'Call subroutine to render the bmp layer
''                RenderSchematicBMPLayer pPourPointFLayer
''            End If
''            If Not gBMPDetailDict Is Nothing Then
''                pBMPKeys = gBMPDetailDict.Keys
''                For i = 0 To (gBMPDetailDict.Count - 1)
''                    pPropertyName = pBMPKeys(i)
''                    pPropertyValue = gBMPDetailDict.Item(pPropertyName)
''                    'Modify if the row present in the table
''                    'pQueryFilter.WhereClause = "ID = " & pID
''                    pQueryFilter.WhereClause = "ID = " & pID & " AND PropName  = '" & pPropertyName & "'"
''                    Set pCursor = pBMPDetailTable.Search(pQueryFilter, False)
''                    pSelRowCount = pBMPDetailTable.RowCount(pQueryFilter)
''                    If pSelRowCount > 0 Then
''                        Set pRow = pCursor.NextRow
''                        Do Until pRow Is Nothing
''                            pRow.value(pPropValueIndex) = pPropertyValue
''                            pRow.Store
''                            Set pRow = pCursor.NextRow
''                        Loop
''                    Else
''                        'Create if the row is not already in the table
''                        Set pRow = pBMPDetailTable.CreateRow
''                        pRow.value(pIDindex) = pID
''                        pRow.value(pPropNameIndex) = pPropertyName
''                        pRow.value(pPropValueIndex) = pPropertyValue
''                        pRow.Store
''                    End If
''
''                Next i
''            End If
''            Set pFeature = pFeatureCursor.NextFeature
''        Loop
''
''    Else
''        MsgBox "Select BMP(s) for editing.", vbExclamation
''    End If
''
''    GoTo CleanUp
''
''ShowError:
''    MsgBox "Edit BMP Details :" & Err.description
''
''CleanUp:
''    Set pPourPointFLayer = Nothing
''    Set pPourPointFClass = Nothing
''    Set pFeature = Nothing
''    Set pRow = Nothing
''    Set pBMPDetailDict = Nothing
''    Set pBMPDetailTable = Nothing
''End Sub

'*******************************************************************************
'Subroutine : EditAggBMP
'Purpose    : Subroutine to modify the existing Aggregate BMPs
'Note       :
'Arguments  : Id of the BMP site
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub EditAggBMP(bmpId As Integer)
On Error GoTo ShowError
    Dim pBMPDetailsTable As iTable
    Set pBMPDetailsTable = GetInputDataTable("AgBMPDetail")
    
    Dim pIDindex As Long
    pIDindex = pBMPDetailsTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPDetailsTable.FindField("PropName")
    Dim pValueIndex As Long
    pValueIndex = pBMPDetailsTable.FindField("PropValue")
    
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName='BMPID' And PropValue = %" & bmpId & " ORDER BY ID"
    
    Dim pCursor As ICursor
    Set pCursor = pBMPDetailsTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPDetailsTable.RowCount(pQueryFilter)
    
    Dim existCatList() As String
    Dim bmpNameList() As String
    
    ReDim existCatList(pSelRowCount - 1)
    ReDim bmpNameList(pSelRowCount - 1)
    
    Dim pRow As iRow
        
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Dim nameInd As Integer
    Dim catInd As Integer
    nameInd = 0
    catInd = 0
    
    If pSelRowCount > 0 Then
        Set pRow = pCursor.NextRow
        Do Until (pRow Is Nothing)
            'BMPName
            pQueryFilter.WhereClause = "PropName='BMPName' And ID = %" & pRow.value(pIDindex)
            Set pCursor2 = pBMPDetailsTable.Search(pQueryFilter, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then
                bmpNameList(nameInd) = pRow2.value(pValueIndex)
                nameInd = nameInd + 1
            End If
            'BMPCategory
            pQueryFilter.WhereClause = "PropName='Category' And ID = %" & pRow.value(pIDindex)
            Set pCursor2 = pBMPDetailsTable.Search(pQueryFilter, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then
                existCatList(nameInd) = pRow2.value(pValueIndex)
                catInd = catInd + 1
            End If
        Loop
    End If
        
    If nameInd > 0 Then
        gBMPTypeTag = "BMPOnMap"
        Load frmAggTypes
        With frmAggTypes
            .cmbExistBMPs.Clear
            .cmbBMPs.Clear
            For nameInd = 0 To UBound(bmpNameList)
                .cmbBMPs.AddItem bmpNameList(nameInd)
            Next
            '.Tag = "BMPOnMap"
            .cmdNew(0).Enabled = False
            .cmbExistBMPs.Enabled = False
        End With
        frmAggTypes.Show vbModal
    End If
    GoTo CleanUp
ShowError:
    MsgBox "Error in EditAggBMP: " & Err.description
    
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    
End Sub


'*******************************************************************************
'Subroutine : CreateAgBmpLuDistTableDBF
'Purpose    : Creates a DBASE file to store the Aggregate BMP Landuse distribution
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Sabu Paul
'*******************************************************************************
Public Function CreateAgBmpLuDistTableDBF(pFileName As String) As iTable
On Error GoTo ShowError
    'Delete data table from temp folder
    DeleteDataTable gMapTempFolder, pFileName
    
    ' Open the Workspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
    Dim pFWS As IFeatureWorkspace
    Set pFWS = pWorkspaceFactory.OpenFromFile(gMapTempFolder, 0)

    Dim pFieldsEdit As IFieldsEdit
    Dim pFieldEdit As IFieldEdit
    Dim pField As esriGeoDatabase.IField
    Dim pFields As esriGeoDatabase.IFields

    Set pFields = New esriGeoDatabase.Fields
    Set pFieldsEdit = pFields
    pFieldsEdit.FieldCount = 5 '9

    'Create BMPID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "BMPID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField

    'Create Lugroup field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LuGroup"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create BMPID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "LuGrpID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create TotalArea Field - Total landuse area
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "TotalArea"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(3) = pField
    
    
    'Create Lugroup field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "AreaDist"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(4) = pField
    
    'Create IntCpn Field - Interception
'    Set pField = New Field
'    Set pFieldEdit = pField
'    With pFieldEdit
'        .name = "Intcpn"
'        .Type = esriFieldTypeDouble
'        .Length = 12
'        .Precision = 11
'        .Scale = 6
'    End With
'    Set pFieldsEdit.Field(4) = pField
'
'    'Create Trtmnt Field - Treatment
'    Set pField = New Field
'    Set pFieldEdit = pField
'    With pFieldEdit
'        .name = "Trtmnt"
'        .Type = esriFieldTypeDouble
'        .Length = 12
'        .Precision = 11
'        .Scale = 6
'    End With
'    Set pFieldsEdit.Field(5) = pField
'
'    'Create Routing Field - Routing
'    Set pField = New Field
'    Set pFieldEdit = pField
'    With pFieldEdit
'        .name = "Routing"
'        .Type = esriFieldTypeDouble
'        .Length = 12
'        .Precision = 11
'        .Scale = 6
'    End With
'    Set pFieldsEdit.Field(6) = pField
'
'    'Create Storage Field - Storage
'    Set pField = New Field
'    Set pFieldEdit = pField
'    With pFieldEdit
'        .name = "Storage"
'        .Type = esriFieldTypeDouble
'        .Length = 12
'        .Precision = 11
'        .Scale = 6
'    End With
'    Set pFieldsEdit.Field(7) = pField
'
'    'Create Outlet Field
'    Set pField = New Field
'    Set pFieldEdit = pField
'    With pFieldEdit
'        .name = "Outlet"
'        .Type = esriFieldTypeDouble
'        .Length = 12
'        .Precision = 11
'        .Scale = 6
'    End With
'    Set pFieldsEdit.Field(8) = pField
    
    Dim pAgBmpLuDTable As iTable
    Set CreateAgBmpLuDistTableDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
    
'    Import_Shape_To_GDB gMapTempFolder, pFileName & ".dbf", gGDBpath, esriDTTable
'    Set CreateAgBmpLuDistTableDBF = GetTable(gGDBpath, pFileName)
    GoTo CleanUp

    
ShowError:
    MsgBox "CreateAgBmpLuDistTableDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing
    Set pAgBmpLuDTable = Nothing
End Function

Public Function Get_OnMap_BMP_Categories(ByVal strBMPID As Integer) As Scripting.Dictionary
    
    Set Get_OnMap_BMP_Categories = CreateObject("Scripting.Dictionary")
        
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("AgBMPDetail")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
        
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    pQueryFilter.WhereClause = "PropName='BMPID' And PropValue='" & strBMPID & "'"
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow

    Dim pQueryFilter2 As IQueryFilter
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Do While Not pRow Is Nothing
        Set pQueryFilter2 = New QueryFilter
        pQueryFilter2.WhereClause = "PropName='Category' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then Get_OnMap_BMP_Categories.add pRow2.value(pNameIndex), pRow2.value(pIDindex)
        Set pRow = pCursor.NextRow
    Loop

End Function

Public Function Get_Agg_BMP_Lu_Distrib(pBMPID As Integer) As Scripting.Dictionary
On Error GoTo ShowError
    Dim pAgBmpLuDTable As iTable
    Set pAgBmpLuDTable = GetInputDataTable("AgLuDistribution")
    
    Dim luDistDict As Scripting.Dictionary
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pAgCursor As ICursor
    Dim pAgRow As iRow
        
    Dim iFldLuGroupID As Integer
    Dim iFldArea As Integer
    Dim iFldAreaDist As Integer
    
    Dim luGroupId As Integer
    Dim luArea As Double
    Dim areaDist As String
    

    If Not pAgBmpLuDTable Is Nothing Then
        iFldLuGroupID = pAgBmpLuDTable.FindField("LuGrpID")
        iFldArea = pAgBmpLuDTable.FindField("TotalArea")
        iFldAreaDist = pAgBmpLuDTable.FindField("AreaDist")
        
        pQueryFilter.WhereClause = "BMPID = " & pBMPID
        If pAgBmpLuDTable.RowCount(pQueryFilter) > 0 Then
            Set luDistDict = New Scripting.Dictionary
            Set pAgCursor = pAgBmpLuDTable.Search(pQueryFilter, False)
            Set pAgRow = pAgCursor.NextRow
            Do Until pAgRow Is Nothing
                luGroupId = pAgRow.value(iFldLuGroupID)
                luArea = pAgRow.value(iFldArea)
                areaDist = pAgRow.value(iFldAreaDist)
                luDistDict.Item(luGroupId) = Array(luArea, areaDist)
                Set pAgRow = pAgCursor.NextRow
            Loop
        End If
    End If
    
    Set Get_Agg_BMP_Lu_Distrib = luDistDict
    GoTo CleanUp
    
ShowError:
    Set Get_Agg_BMP_Lu_Distrib = Nothing
    MsgBox "Error in Get_Agg_BMP_Lu_Distrib:" & Err.description
CleanUp:
    Set pAgBmpLuDTable = Nothing
    Set luDistDict = Nothing
    Set pQueryFilter = Nothing
    Set pAgRow = Nothing
    Set pAgCursor = Nothing
End Function

'return Success or error message
Public Function Check_Completeness_Agg_Bmp_Lu() As String
On Error GoTo ShowError
    Set gAggBMPFlagDict = New Scripting.Dictionary
    
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    
    Dim pAgBmpLuDTable As iTable
    Set pAgBmpLuDTable = GetInputDataTable("AgLuDistribution")
    
    If pBMPFLayer Is Nothing Then Check_Completeness_Agg_Bmp_Lu = "Success": Exit Function
    
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "TYPE = 'Aggregate'"
    
    If pBMPFClass.FeatureCount(pQueryFilter) = 0 Then Check_Completeness_Agg_Bmp_Lu = "Success": Exit Function
    
    If pAgBmpLuDTable Is Nothing Then Check_Completeness_Agg_Bmp_Lu = "Missing Aggregage BMP Land use distribution": Exit Function
    
    Dim pBMPCursor As IFeatureCursor
    Set pBMPCursor = pBMPFClass.Search(pQueryFilter, False)
    
    Dim pBMPFeature As IFeature
    Dim pBMPID As Integer
    Dim pBmpIdFld As Integer
    
    Dim errList As String
    errList = ""
    pBmpIdFld = pBMPFClass.FindField("ID")
    
    Set pBMPFeature = pBMPCursor.NextFeature
    Do Until pBMPFeature Is Nothing
        pBMPID = pBMPFeature.value(pBmpIdFld)
        gAggBMPFlagDict.Item(pBMPID) = True
        pQueryFilter.WhereClause = "BMPID = " & pBMPID
        If pAgBmpLuDTable.RowCount(pQueryFilter) = 0 Then errList = errList & vbNewLine & pBMPID
        Set pBMPFeature = pBMPCursor.NextFeature
    Loop
    If errList = "" Then
        Check_Completeness_Agg_Bmp_Lu = "Success"
    Else
        Check_Completeness_Agg_Bmp_Lu = "Missing land use distribution for aggregate BMPs" & errList
    End If
    GoTo CleanUp
ShowError:
    MsgBox "Error in Check_Completeness_Agg_Bmp_Lu:" & Err.description
CleanUp:
    Set pBMPFLayer = Nothing
    Set pBMPCursor = Nothing
    Set pQueryFilter = Nothing
    Set pAgBmpLuDTable = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFLayer = Nothing
End Function

Public Function GetAggBMPTypes(bmpId As Integer) As Scripting.Dictionary
On Error GoTo ShowError
    Dim resDict As Scripting.Dictionary
    Set resDict = New Scripting.Dictionary
    
    Dim pBMPDetailsTable As iTable
    Set pBMPDetailsTable = GetInputDataTable("AgBMPDetail")
    
    Dim pIDindex As Long
    pIDindex = pBMPDetailsTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPDetailsTable.FindField("PropName")
    Dim pValueIndex As Long
    pValueIndex = pBMPDetailsTable.FindField("PropValue")
    
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName='BMPID' And PropValue = '" & bmpId & "'"
    
    Dim pCursor As ICursor
    Set pCursor = pBMPDetailsTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPDetailsTable.RowCount(pQueryFilter)
    
    If pSelRowCount = 0 Then Set GetAggBMPTypes = Nothing: Exit Function
    
    Dim pRow As iRow
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Dim pBMPName As String
    Dim pBMPType As String
    Dim pID As Integer
    Dim pBMPCat As String
    
    If pSelRowCount > 0 Then
        Set pRow = pCursor.NextRow
        Do Until (pRow Is Nothing)
            pID = pRow.value(pIDindex)
            'BMPName
            pQueryFilter.WhereClause = "PropName='BMPName' And ID = " & pID
            Set pCursor2 = pBMPDetailsTable.Search(pQueryFilter, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then
                pBMPName = pRow2.value(pValueIndex)
            End If
            'BMPCategory
            pQueryFilter.WhereClause = "PropName='Category' And ID = " & pID
            Set pCursor2 = pBMPDetailsTable.Search(pQueryFilter, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then
                pBMPCat = pRow2.value(pValueIndex)
            End If
            'BMPType
            pQueryFilter.WhereClause = "PropName='BMPType' And ID = " & pID
            Set pCursor2 = pBMPDetailsTable.Search(pQueryFilter, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then
                pBMPType = pRow2.value(pValueIndex)
            End If
            resDict.Item(pID) = Array(pBMPName, pBMPCat, pBMPType)
            Set pRow = pCursor.NextRow
        Loop
    End If
    
    Set GetAggBMPTypes = resDict
    
    GoTo CleanUp
ShowError:
    MsgBox "Error in GetAggBMPTypes: " & Err.description
    
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    
End Function
'*******************************************************************************
'Subroutine : ModifyBmpDetails
'Purpose    : Modify the detail information of an existing BMP type in BMPDefaults table
'Note       :
'Arguments  : Id and name of BMP Type
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub ModifyBmpDetails(pBMPID As Integer, pBMPName As String)
On Error GoTo ShowError
    
    gNewBMPId = pBMPID
    gNewBMPName = pBMPName
        
    '***creating the dictionary to hold control names and corresponging values
    Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = GetBMPDetailDict(gNewBMPId)
    'Set the decay rates and underdrain percentage removals
    LoadPollutantData pBmpDetailDict

    Call CallInitRoutines(gNewBMPType, pBmpDetailDict)
       
    'First delete the rows with specific ID
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "ID = " & pBMPID
    If Not (gBMPDetailDict Is Nothing) Then
          If Not (pBMPDetailTable Is Nothing) Then
              pBMPDetailTable.DeleteSearchedRows pQueryFilter
          Else
              MsgBox "Missing 'BMPDetail' Table"
          End If
        
          'Then insert the new rows into the table
          Call InsertNewBmpDetails(gBMPDetailDict)
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "ModifyBmpDetails :", Err.description
    
CleanUp:
    Set pBMPDetailTable = Nothing
    Set pQueryFilter = Nothing
End Sub



'*******************************************************************************
'Subroutine : insertNewBmpDetails
'Purpose    : Store the detail information of a BMP type in BMPDetails table
'Note       :
'Arguments  : Dictionary containing the BMP details
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InsertNewBmpDetails(BMPDetailDict As Dictionary)
On Error GoTo ShowError

    InitializeMapDocument
    '**** This inserts the dimensions of a new BMP type into BMP table
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
    
    If (pBMPDetailTable Is Nothing) Then
        Set pBMPDetailTable = CreatePropertiesTableDBF("BMPDetail")
        AddTableToMap pBMPDetailTable
    End If
    Dim pIDindex As Long
    pIDindex = pBMPDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
  
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    'Iterate over the entire dictionary
    Dim pRow As iRow

    Dim i As Integer
'    If BMPDetailDict.Exists("BMPName") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPName"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
'    If BMPDetailDict.Exists("BMPType") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPType"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
'    If BMPDetailDict.Exists("BMPClass") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPClass"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
    
    Dim pBMPKeys
    pBMPKeys = BMPDetailDict.keys
        
    For i = 0 To (BMPDetailDict.Count - 1)
        pPropertyName = pBMPKeys(i)
        pPropertyValue = BMPDetailDict.Item(pPropertyName)
        Set pRow = pBMPDetailTable.CreateRow
        pRow.value(pIDindex) = gNewBMPId
        pRow.value(pPropNameIndex) = pPropertyName
        pRow.value(pPropValueIndex) = pPropertyValue
        pRow.Store
    Next
    
    ' ***************************************
    ' Now Insert the BMP option Parameters....
    ' ***************************************
    pBMPKeys = gBMPOptionsDict.keys
    For i = 0 To (gBMPOptionsDict.Count - 1)
        pPropertyName = pBMPKeys(i)
        'Modified on Feb 23, 2009 - Sabu Paul - No need to insert the entry if it is already in the BMPDetailDict.
        'If pPropertyName <> "BMPType" Then
        If Not BMPDetailDict.Exists(pPropertyName) Then
            pPropertyValue = gBMPOptionsDict.Item(pPropertyName)
            Set pRow = pBMPDetailTable.CreateRow
            pRow.value(pIDindex) = gNewBMPId
            pRow.value(pPropNameIndex) = pPropertyName
            pRow.value(pPropValueIndex) = pPropertyValue
            pRow.Store
        End If
    Next
    ' Store the Placed BMP to dict.....
    'If gBMPOptionsDict.Exists("Category") Then gBMPPlacedDict.add gBMPOptionsDict.Item("Category"), gBMPOptionsDict.Item("Category")
    'If gBMPOptionsDict.Exists("BMPType") Then gBMPPlacedDict.add gBMPOptionsDict.Item("BMPType"), gBMPOptionsDict.Item("BMPType")

    GoTo CleanUp
    
ShowError:
    MsgBox "insertNewBmpDetails :", Err.description
    
CleanUp:
    Set pBMPDetailTable = Nothing
    Set pRow = Nothing
    Set pBMPKeys = Nothing
End Sub
