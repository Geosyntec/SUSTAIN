Attribute VB_Name = "ModuleDecayFact"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleBMPTypes
'   Purpose:     Functions and subroutine to
'                and specific BMP sites
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  08/../2004 - Mira Chokshi
'                Modified: 08/19/2004 - Sabu Paul added comments to project
'
'******************************************************************************

Option Explicit
Option Base 0

'*******************************************************************************
'Subroutine : AddRecordInInputDBFTable
'Purpose    : Add a new record into the decay table
'Note       :
'Arguments  : Table, name of the files, Id of the BMP, Pollutant index,Field name
'Author     : Mira Chokshi
'History    :
'*******************************************************************************

Private Sub AddRecordInInputDBFTable(pTable As iTable, pFileName As String, pBMPID As Integer, pQualCount As Integer, pFieldName As String)
    
On Error GoTo ErrorHandler:
    Dim iBMP As Long
    iBMP = pTable.FindField("BMPID")
    Dim pRow As iRow
    'Insert a new row in the table
    Set pRow = pTable.CreateRow
    pRow.value(iBMP) = pBMPID
    Dim pQualField As String
    Dim iQualField As Long
    Dim iqual As Integer
    For iqual = 1 To pQualCount
        pQualField = pFieldName & iqual
        iQualField = pTable.FindField(pQualField)
        If (iQualField >= 0) Then
            'set the field value
            pRow.value(iQualField) = 0
        End If
    Next
    'Save the record
    pRow.Store
    GoTo CleanUp

ErrorHandler:
    MsgBox "AddRecordInInputDBFTable :", Err.description
CleanUp:
    Set pRow = Nothing
    
End Sub

'*******************************************************************************
'Subroutine : CreateTEMPDBFTable
'Purpose    : Create a DBASE file to hold decay factor or percent removal factor
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the files, No of pollutants, Field name (decay/precent removal)
'Author     : Mira Chokshi
'History    :
'*******************************************************************************

Public Function CreateTEMPDBFTable(pFileName As String) As iTable

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
    pFieldsEdit.FieldCount = 5

    'Create Landuse Group Code Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "POLLUTANT"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(0) = pField

    'Create PARAM Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "DECAY"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create PARAM Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "K"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(2) = pField
    
    'Create PARAM Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "C"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(3) = pField
    
    'Create PARAM Type Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "REMOVAL"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(4) = pField
    

    
  Set CreateTEMPDBFTable = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateTEMPDBFTable: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function

'*******************************************************************************
'Subroutine : DefineDecayFactorForBMPs
'Purpose    : Based on the number of polluants present in the timeseries files
'             creates two tables (1)to hold the decay factors and (2) for percent removal rates
'             then let the user to set the decay factors and percent removal rates
'             for each BMP, pollutant combination
'Note       :
'Arguments  :
'Author     : Mira Chokshi
'History    :
'*******************************************************************************
'''Public Sub DefineDecayFactorForBMPs()
'''On Error GoTo ShowError
'''
'''    'Delete decayfact and pctremoval
'''    DeleteDataTable gMapTempFolder, "DecayFact"
'''    DeleteDataTable gMapTempFolder, "PctRemoval"
'''
'''    'Read the timeseries file and find the total count of pollutants
'''    Dim pTotalQUAL As Integer
'''    pTotalQUAL = ReadTimeSeriesForPollutantCount
'''    'Create a new table with field names
'''    Dim boolAddDF As Boolean
'''    boolAddDF = False
'''    Dim pTableDF As iTable
'''    Set pTableDF = GetInputDataTable("DecayFact")
'''    If (pTableDF Is Nothing) Then
'''        boolAddDF = True
'''        Set pTableDF = CreateTEMPDBFTable("DecayFact", pTotalQUAL, "QUALDECAY")
'''    End If
'''
'''    Dim boolAddUD As Boolean
'''    boolAddUD = False
'''    Dim pTablePR As iTable
'''    Set pTablePR = GetInputDataTable("PctRemoval")
'''    If (pTablePR Is Nothing) Then
'''        boolAddUD = True
'''        Set pTablePR = CreateTEMPDBFTable("PctRemoval", pTotalQUAL, "QUALPCT")
'''    End If
'''
'''    'Iterate over each bmp to get the id and add a new record in decayfact table
'''    Dim pFeatureLayer As IFeatureLayer
'''    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
'''    If (pFeatureLayer Is Nothing) Then
'''        MsgBox "BMPs layer not found."
'''        Exit Sub
'''    End If
'''    Dim pFeatureClass As IFeatureClass
'''    Set pFeatureClass = pFeatureLayer.FeatureClass
'''    Dim pFeatureCursor As IFeatureCursor
'''    Set pFeatureCursor = pFeatureClass.Search(Nothing, True)
'''    Dim iBMP As Long
'''    iBMP = pFeatureCursor.FindField("ID")
'''    Dim pBMP As Integer
'''    Dim pFeature As IFeature
'''    Set pFeature = pFeatureCursor.NextFeature
'''    Do While Not pFeature Is Nothing
'''        pBMP = pFeature.value(iBMP)
'''        If (boolAddDF) Then
'''            AddRecordInInputDBFTable pTableDF, "DecayFact", pBMP, pTotalQUAL, "QUALDECAY"
'''        End If
'''        If (boolAddUD) Then
'''            AddRecordInInputDBFTable pTablePR, "PctRemoval", pBMP, pTotalQUAL, "QUALPCTREM"
'''        End If
'''        Set pFeature = pFeatureCursor.NextFeature
'''    Loop
'''
'''    'Add data table to the map
'''    If (boolAddDF) Then
'''        AddTableToMap pTableDF
'''    End If
'''    If (boolAddUD) Then
'''        AddTableToMap pTablePR
'''    End If
'''    'Set the decay and percent removal based on the
'''    'BMP Templates
'''    SetDecayFactorForBMPs
'''    'Open FrmDataGrid Form
'''    'FrmDataGrid.Show vbModal
'''
'''    GoTo CleanUp
'''ShowError:
'''    MsgBox "DefineDecayFactors: " & Err.description
'''CleanUp:
'''    Set pTableDF = Nothing
'''    Set pTablePR = Nothing
'''    Set pFeatureClass = Nothing
'''    Set pFeatureCursor = Nothing
'''    Set pFeature = Nothing
'''
'''End Sub



'*******************************************************************************
'Subroutine : ReadTimeSeriesForPollutantCount
'Purpose    : Read the timeseries files for set during the land use reclassification
'             step and set the number of pollutant
'Note       :
'Arguments  :
'Author     : Mira Chokshi
'History    :
'*******************************************************************************
Private Function ReadTimeSeriesForPollutantCount() As Integer
On Error GoTo ShowError
    Dim pTable As iTable
    Set pTable = GetInputDataTable("TSAssigns")
    If (pTable Is Nothing) Then
        MsgBox "TSAssigns table not found."
        Exit Function
    End If
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iTimeSeries As Long
    iTimeSeries = pCursor.FindField("TimeSeries")
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim pTimeSeriesFile As String
    If Not (pRow Is Nothing) Then
        pTimeSeriesFile = pRow.value(iTimeSeries)
    End If
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pTS As TextStream
    Set pTS = fso.OpenTextFile(pTimeSeriesFile, ForReading, False, TristateUseDefault)
    Dim doContinue As Boolean
    doContinue = True
    Dim pString As String
    Dim pSubStrings() As String
    Dim pTotalQUAL As Integer
    pTotalQUAL = 0
    Dim pSecondToken As String
    Do While doContinue
        pString = pTS.ReadLine

        If (StringContains(pString, "SOQUAL") Or StringContains(pString, "SLDS") Or StringContains(pString, "WSSD")) Then ' Modified to include WSSD (Total suspended solids) -- Sabu Paul
            pTotalQUAL = pTotalQUAL + 1
        ElseIf (StringContains(pString, "Date/time")) Then
            doContinue = False
        End If
    Loop
    pTS.Close
    
    ReadTimeSeriesForPollutantCount = pTotalQUAL
    GoTo CleanUp

ShowError:
    MsgBox "Read TimeSeries File: " & Err.description
    
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set fso = Nothing
    Set pTS = Nothing
End Function
'*******************************************************************************
'Subroutine : StringContains
'Purpose    : Checks where a given string is contained in another one
'Note       :
'Arguments  :
'Author     : Mira Chokshi
'History    :
'*******************************************************************************
Public Function StringContains(FindString As String, SearchString As String) As Boolean
    Dim TempString As String
    TempString = Replace(FindString, SearchString, "")
    If (FindString <> TempString) Then
        StringContains = True
    Else
        StringContains = False
    End If
End Function

'''Public Sub SetDecayFactorForBMPs()
'''On Error GoTo ShowError
'''
'''
'''    'Read the timeseries file and find the total count of pollutants
'''    Dim pTotalQUAL As Integer
'''    pTotalQUAL = ReadTimeSeriesForPollutantCount
'''    'Create a new table with field names
'''    Dim pTableDF As iTable
'''    Set pTableDF = GetInputDataTable("DecayFact")
'''    If (pTableDF Is Nothing) Then
'''        Exit Sub
'''    End If
'''
'''    Dim pTablePR As iTable
'''    Set pTablePR = GetInputDataTable("PctRemoval")
'''    If (pTablePR Is Nothing) Then
'''        Exit Sub
'''    End If
'''    'Iterate over each bmp to get the id and add a new record in decayfact table
'''    Dim pFeatureLayer As IFeatureLayer
'''    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
'''    If (pFeatureLayer Is Nothing) Then
'''        MsgBox "BMPs layer not found."
'''        Exit Sub
'''    End If
'''    Dim pFeatureClass As IFeatureClass
'''    Set pFeatureClass = pFeatureLayer.FeatureClass
'''    Dim pFeatureCursor As IFeatureCursor
'''    Set pFeatureCursor = pFeatureClass.Search(Nothing, True)
'''    Dim iBMP As Long
'''    iBMP = pFeatureCursor.FindField("ID")
'''    Dim pBMP As Integer
'''    Dim pFeature As IFeature
'''
'''    Dim pBMPDetailDict As Scripting.Dictionary
'''    Set pBMPDetailDict = CreateObject("Scripting.Dictionary")
'''
'''    Dim pCursor As ICursor
'''    Dim pQueryFilter As IQueryFilter
'''    Set pQueryFilter = New QueryFilter
'''
'''    Dim pRow As iRow
'''    Dim pCols As Long
'''    Dim pC As Long
'''
'''    Dim pCurFieldIndex As Long
'''    Dim pCurFieldValue As Double
'''
'''    Set pFeature = pFeatureCursor.NextFeature
'''    Do While Not pFeature Is Nothing
'''        pBMP = pFeature.value(iBMP)
'''        Set pBMPDetailDict = ModuleBMPDetails.GetBMPDetailDict(pBMP)
'''        'Set pQueryFilter
'''        pQueryFilter.WhereClause = "BMPID = " & pBMP
'''        'Modify the Decay table
'''        Set pCursor = pTableDF.Search(pQueryFilter, True)
'''        Set pRow = pCursor.NextRow
'''        pCols = pTableDF.Fields.FieldCount - 1
'''        Do While Not pRow Is Nothing
'''            For pC = 2 To pCols
'''                pCurFieldIndex = pCursor.FindField("QualDecay" & pC - 1)
'''                If pBMPDetailDict.Exists("Decay" & pC - 1) Then
'''                    pCurFieldValue = pBMPDetailDict.Item("Decay" & pC - 1)
'''                    pRow.value(pCurFieldIndex) = pCurFieldValue
'''                End If
'''            Next
'''            pRow.Store
'''            Set pRow = pCursor.NextRow
'''        Loop
'''        'Modify the Pecent removal table
'''        Set pCursor = pTablePR.Search(pQueryFilter, True)
'''        Set pRow = pCursor.NextRow
'''        pCols = pTablePR.Fields.FieldCount - 1
'''        Do While Not pRow Is Nothing
'''            For pC = 2 To pCols
'''                pCurFieldIndex = pCursor.FindField("QualPCT" & pC - 1)
'''                If pBMPDetailDict.Exists("Decay" & pC - 1) Then
'''                    pCurFieldValue = pBMPDetailDict.Item("PctRem" & pC - 1)
'''                    pRow.value(pCurFieldIndex) = pCurFieldValue
'''                End If
'''            Next
'''            pRow.Store
'''            Set pRow = pCursor.NextRow
'''        Loop
'''
'''        Set pFeature = pFeatureCursor.NextFeature
'''    Loop
'''    GoTo CleanUp
'''ShowError:
'''    MsgBox "SetDecayFactorForBMPs: " & Err.description
'''CleanUp:
'''    Set pTableDF = Nothing
'''    Set pTablePR = Nothing
'''    Set pFeatureClass = Nothing
'''    Set pFeatureCursor = Nothing
'''    Set pFeature = Nothing
'''
'''End Sub



'*******************************************************************************
'Subroutine : CreatePollutantList
'Purpose    : Read the timeseries files and get the pollutant list-> Update the global variable
'             gPollutants
'Note       :
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub CreatePollutantList()
On Error GoTo ShowError
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("LUReclass")
''    If (Not pTable Is Nothing) Then
''        Call CreatePollutantListFromLUReclass
''    End If
''
''    Dim pTable2 As iTable
''    Set pTable2 = GetInputDataTable("LANDLUReclass")
''    If (Not pTable2 Is Nothing) Then
''        Call CreatePollutantListFromSWMMLUReclass
''    End If
''
''    If (pTable Is Nothing And pTable2 Is Nothing) Then
''        MsgBox "Landuse Reclassification required."
''        Exit Sub
''    End If
''
''    '** cleanup
''    Set pTable = Nothing
''    Set pTable2 = Nothing


    'Pollutant list is created using pollutants table
    'Sabu Paul - June 14, 2007
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        MsgBox "Missing pollutants table: Define pollutants first"
        Exit Sub
    End If
    
    Dim iNameFld As Integer
    iNameFld = pTable.FindField("Name")
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pCount As Integer
    pCount = pTable.RowCount(Nothing)
    
    ReDim Preserve gPollutants(pCount - 1) As String
    
    Dim i As Integer
    For i = 1 To pCount
        pQueryFilter.WhereClause = " ID = " & i
        Set pCursor = pTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        gPollutants(i - 1) = pRow.value(iNameFld)
    Next
    
    GoTo CleanUp
ShowError:
    MsgBox "Error in CreatePollutantList : " & Err.description
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pQueryFilter = Nothing
End Sub


Private Sub CreatePollutantListFromLUReclass()
On Error GoTo ShowError
    Dim pTable As iTable
    Set pTable = GetInputDataTable("TSAssigns")
    If (pTable Is Nothing) Then
        MsgBox "TSAssigns table not found."
        Exit Sub
    End If
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iTimeSeries As Long
    iTimeSeries = pCursor.FindField("TimeSeries")
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim pTimeSeriesFile As String
    If Not (pRow Is Nothing) Then
        pTimeSeriesFile = pRow.value(iTimeSeries)
    End If
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim pTS As TextStream
    Set pTS = fso.OpenTextFile(pTimeSeriesFile, ForReading, False, TristateUseDefault)
    Dim doContinue As Boolean
    doContinue = True
    Dim pString As String
    Dim pSubStrings
    Dim pTotalQUAL As Integer
    pTotalQUAL = 0
    
    
    Dim pPollutName As String
    
    Do While doContinue
        pString = pTS.ReadLine
        If StringContains(pString, "SOQUAL") Then
            ReDim Preserve gPollutants(pTotalQUAL) As String
            pPollutName = Mid(pString, 13, 9)
            If StringContains(pPollutName, "(") Then
                pPollutName = Replace(pPollutName, "(", "")
            ElseIf StringContains(pPollutName, ")") Then
                pPollutName = Replace(pPollutName, ")", "")
            End If
            gPollutants(pTotalQUAL) = pPollutName
            pTotalQUAL = pTotalQUAL + 1
        ElseIf StringContains(pString, "SLDS") Or StringContains(pString, "WSSD") Then ' Modified to include WSSD (Total suspended solids) -- Sabu Paul
            ReDim Preserve gPollutants(pTotalQUAL) As String
            gPollutants(pTotalQUAL) = "Total Suspended Solids (TSS)"
            pTotalQUAL = pTotalQUAL + 1
        ElseIf (StringContains(pString, "Date/time")) Then
        doContinue = False
        End If
    Loop
    pTS.Close
    GoTo CleanUp

ShowError:
    MsgBox "CreatePollutantListFromLUReclass", Err.description
CleanUp:
End Sub


Private Sub CreatePollutantListFromSWMMLUReclass()
On Error GoTo ShowError
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = LoadPropertyNames("LANDPollutants")
    If (pPollutantCollection Is Nothing) Then
        MsgBox "No pollutants defined for SWMM option."
        Exit Sub
    End If
    
    Dim pTotalQUAL As Integer
    pTotalQUAL = pPollutantCollection.Count
    ReDim Preserve gPollutants(pTotalQUAL - 1)
    
    Dim iCount As Integer
    For iCount = 1 To pPollutantCollection.Count
        gPollutants(iCount - 1) = pPollutantCollection.Item(iCount)
    Next
    GoTo CleanUp
    
ShowError:
    MsgBox "CreatePollutantListFromSWMMLUReclass", Err.description
CleanUp:
    Set pPollutantCollection = Nothing
End Sub


'*******************************************************************************
'Subroutine : CreateTsMultipliersDBF
'Purpose    : Creates a DBASE file in the project temp directory to store the
'             time series multipliers
'Note       : Name of the DBASE file should not contain the .dbf extension
'Arguments  : Name of the DBASE file
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function CreateTsMultipliersDBF(pFileName As String) As iTable

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

    'Create ID Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "ID"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(0) = pField
    
    'Create multiplier Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Multiplier"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create sediment flag Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "SedFlag"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
 

  Set CreateTsMultipliersDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
  GoTo CleanUp

ShowError:
    MsgBox "CreateTsMultipliersDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function


'*******************************************************************************
'Subroutine : CreateSoilFractionsDBF
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function CreateSoilFractionsDBF(pFileName As String) As iTable

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
        .name = "TimeSeries"
        .Type = esriFieldTypeString
        .Length = 50
    End With
    Set pFieldsEdit.Field(0) = pField
    
    'Create sand fraction Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Sand"
        .Type = esriFieldTypeDouble
        .Length = 12
        .Precision = 11
        .Scale = 6
    End With
    Set pFieldsEdit.Field(1) = pField
    
    'Create silt fraction Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Silt"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(2) = pField
 
    'Create clay Field
    Set pField = New esriGeoDatabase.Field
    Set pFieldEdit = pField
    With pFieldEdit
        .name = "Clay"
        .Type = esriFieldTypeInteger
        .Length = 10
    End With
    Set pFieldsEdit.Field(3) = pField
    
    Set CreateSoilFractionsDBF = pFWS.CreateTable(pFileName, pFields, Nothing, Nothing, "")
    GoTo CleanUp

ShowError:
    MsgBox "CreateSoilFractionsDBF: " & Err.description
CleanUp:
    Set pWorkspaceFactory = Nothing
    Set pFWS = Nothing
    Set pFieldsEdit = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pFields = Nothing

End Function
