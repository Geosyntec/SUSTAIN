VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSWMMAssignLanduses 
   Caption         =   "Reclassify Landuses"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   Icon            =   "FrmSWMMAssignLanduses.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      Picture         =   "FrmSWMMAssignLanduses.frx":08CA
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   3360
      Width           =   255
   End
   Begin VB.ComboBox cmbLUGroup 
      Height          =   315
      ItemData        =   "FrmSWMMAssignLanduses.frx":0C0C
      Left            =   2760
      List            =   "FrmSWMMAssignLanduses.frx":0C0E
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2280
      Width           =   3720
   End
   Begin VB.TextBox txtPercentage 
      Height          =   360
      Left            =   2760
      TabIndex        =   9
      Top             =   2760
      Width           =   840
   End
   Begin VB.CommandButton cmdAddLandUseReclass 
      Height          =   405
      Left            =   2760
      Picture         =   "FrmSWMMAssignLanduses.frx":0C10
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1560
   End
   Begin VB.CommandButton cmdRemoveLandUseReclass 
      Height          =   400
      Left            =   4920
      Picture         =   "FrmSWMMAssignLanduses.frx":2692
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   6720
      TabIndex        =   6
      Top             =   120
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   6720
      TabIndex        =   5
      Top             =   720
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      Picture         =   "FrmSWMMAssignLanduses.frx":47A4
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      Picture         =   "FrmSWMMAssignLanduses.frx":4AE6
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      Picture         =   "FrmSWMMAssignLanduses.frx":4E28
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   2760
      Width           =   255
   End
   Begin MSComctlLib.ListView ListTimeSeries 
      Height          =   2400
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView listLUType 
      Height          =   1905
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Percentage Imperviousness"
      Height          =   240
      Left            =   600
      TabIndex        =   14
      Top             =   2775
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Landuse Group"
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Select Input Landuse Types"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Add New or Remove Existing Landuse Groups"
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSWMMAssignLanduses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LandUseTextFiles() As String
Private LandUseDictionary As Scripting.Dictionary
Private pGroupIndex As Integer
Private pGroupIdDict As Scripting.Dictionary

Private Sub cmbLUGroup_Click()
    Dim pLanduseType As String
    pLanduseType = cmbLUGroup.Text
       
    Select Case pLanduseType
        Case "Forest":
            txtPercentage.Text = "0"
        Case "Agriculture":
            txtPercentage.Text = "0"
        Case "High-Density-Residential":
            txtPercentage.Text = "80"
        Case "High-Density-Residential-PERVIOUS":
            txtPercentage.Text = "0"
        Case "High-Density-Residential-IMPERVIOUS":
            txtPercentage.Text = "100"
        Case "Medium-Density-Residential":
            txtPercentage.Text = "60"
        Case "Medium-Density-Residential-PERVIOUS":
            txtPercentage.Text = "0"
        Case "Medium-Density-Residential-IMPERVIOUS":
            txtPercentage.Text = "100"
        Case "Low-Density-Residential":
            txtPercentage.Text = "30"
        Case "Low-Density-Residential-PERVIOUS":
            txtPercentage.Text = "0"
        Case "Low-Density-Residential-IMPERVIOUS":
            txtPercentage.Text = "100"
        Case "Commercial":
            txtPercentage.Text = "90"
        Case "Commercial-PERVIOUS":
            txtPercentage.Text = "0"
        Case "Commercial-IMPERVIOUS":
            txtPercentage.Text = "100"
        Case "Road":
            txtPercentage.Text = "100"
        Case "Rooftop":
            txtPercentage.Text = "100"
    End Select
    
    '*** Validate input text boxes for percent impervious values
    ValidatePercentImperviousValues
    
End Sub

Private Sub cmdAddLandUseReclass_Click()

    Dim pGroupName As String
    Dim pImpKey As String
    
    pGroupName = cmbLUGroup.Text
    pImpKey = pGroupName
    
    Dim curIndex As Integer
    
    'On August 31, 20004 - Sabu Paul
    Dim pImpPerc As Double
    pImpPerc = 0
    
    If Not (IsNumeric(Trim(txtPercentage.Text))) Then
        MsgBox "Percentage imperviousness number should be a valid number.", vbExclamation
        Exit Sub
    End If
    If (CDbl(txtPercentage.Text) < 0 Or CDbl(txtPercentage.Text) > 100) Then
        MsgBox "Percentage imperviousness number should be within (0-100) range.", vbExclamation
        Exit Sub
    End If
    '** define the impervious percentage
    pImpPerc = CDbl(Trim(txtPercentage.Text))

    Dim bSelected As Boolean
    bSelected = False
    Dim pIndex As Integer
    Dim pSelectedList As ListItems
    
    For pIndex = listLUType.ListItems.Count To 1 Step -1
        Dim pRow As Integer
        Dim pLUPresent As Boolean
        Dim iCnt As Integer
        Dim luCode As Integer
        Dim luDesc As String
        If listLUType.ListItems.Item(pIndex).Selected Then
            luCode = listLUType.ListItems.Item(pIndex)
            luDesc = listLUType.ListItems.Item(pIndex).SubItems(1)
            '** check if the impervious key exists
            If pGroupIdDict.Exists(pImpKey) Then
                curIndex = pGroupIdDict.Item(pImpKey)
            Else
                pGroupIdDict.Item(pImpKey) = pGroupIndex
                curIndex = pGroupIdDict.Item(pImpKey)
                pGroupIndex = pGroupIndex + 1
            End If
            Dim itmX As ListItem
            Set itmX = ListTimeSeries.ListItems.Add(, , curIndex)   ' LuCode.
            itmX.SubItems(1) = cmbLUGroup.Text ' Lu Description
            itmX.SubItems(2) = luCode
            itmX.SubItems(3) = luDesc
            itmX.SubItems(4) = "1"
            itmX.SubItems(5) = pImpPerc
           
            listLUType.ListItems.Remove (pIndex)
        End If
    Next pIndex
    
    '*** Clear the textbox values for percent pervious and impervious values
    cmbLUGroup.ListIndex = 0
    txtPercentage.Text = 0

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If listLUType.ListItems.Count > 0 Then
        MsgBox "Please reclassify all landuse types to continue.", vbExclamation
        Exit Sub
    End If
    Dim pIndex As Integer
    Dim isImp As Integer
    Dim pGroupId As String
    Dim pkey As String
    
    For pIndex = 1 To ListTimeSeries.ListItems.Count
        pGroupId = ListTimeSeries.ListItems.Item(pIndex).SubItems(1)
        isImp = ListTimeSeries.ListItems.Item(pIndex).SubItems(4)
        If isImp = 1 Then
            pkey = "imp" & pGroupId
        Else
            pkey = "perv" & pGroupId
        End If
    Next pIndex
    
    ReDim LandUseTextFiles(1 To 6, 1 To 1)
    LandUseTextFiles(1, 1) = "LU Group Code"
    LandUseTextFiles(2, 1) = "LU Group"
    LandUseTextFiles(3, 1) = "Landuse Code"
    LandUseTextFiles(4, 1) = "Landuse Description"
    LandUseTextFiles(5, 1) = "Impervious ?"
    LandUseTextFiles(6, 1) = "Percentage"
    
    For pIndex = 1 To ListTimeSeries.ListItems.Count
        ReDim Preserve LandUseTextFiles(1 To 6, 1 To pIndex + 1)
        LandUseTextFiles(1, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex)
        LandUseTextFiles(2, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(1)
        LandUseTextFiles(3, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(2)
        LandUseTextFiles(4, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(3)
        LandUseTextFiles(5, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(4)
        LandUseTextFiles(6, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(5)
    Next pIndex
    
    AddSWMMLanduseReclassification LandUseTextFiles
    Unload Me
End Sub

Private Sub cmdRemoveLandUseReclass_Click()
    
    Dim pIndex As Integer
    Dim luCode As Integer
    Dim luDesc As String
    
    Dim pLuIndex As Integer
    Dim isLuFound As Boolean
    
    Dim pIndex2 As Integer
    Dim removedCodeArray() As Integer
    Dim numCodeRemoved As Integer
    numCodeRemoved = 0
    
    For pIndex = ListTimeSeries.ListItems.Count To 1 Step -1
        If ListTimeSeries.ListItems.Item(pIndex).Selected Then
            luCode = ListTimeSeries.ListItems.Item(pIndex).SubItems(2)
            
            '** modified the code to remove all the rows with same luCode
            ReDim Preserve removedCodeArray(numCodeRemoved)
            removedCodeArray(numCodeRemoved) = luCode
            numCodeRemoved = numCodeRemoved + 1
            
            luDesc = ListTimeSeries.ListItems.Item(pIndex).SubItems(3)
            isLuFound = False
            For pLuIndex = listLUType.ListItems.Count To 1 Step -1
                If listLUType.ListItems.Item(pLuIndex) = luCode Then
                    isLuFound = True
                End If
            Next pLuIndex
            If isLuFound = False Then
                Dim itmX As ListItem
                Set itmX = listLUType.ListItems.Add(, , luCode)   ' LuCode.
                itmX.SubItems(1) = luDesc ' Lu Description
            End If
            '** modified the code to remove all the rows with same luCode
            'ListTimeSeries.ListItems.Remove (pIndex)
            
        End If
    Next pIndex
    
    '** modified the code to remove all the rows with same luCode
    For numCodeRemoved = 0 To UBound(removedCodeArray)
        For pIndex2 = ListTimeSeries.ListItems.Count To 1 Step -1
            If (removedCodeArray(numCodeRemoved) = ListTimeSeries.ListItems.Item(pIndex2).SubItems(2)) Then
                ListTimeSeries.ListItems.Remove (pIndex2)
            End If
        Next pIndex2
    Next numCodeRemoved
    
    Dim pGroupIdDictKeys
    pGroupIdDictKeys = pGroupIdDict.keys
    Dim pkey As String
    Dim pIndexId As Integer
    Dim i As Integer
    Dim isIdFound As Boolean
    'Iterate through the dictionary and Timeseries list items
    'If any of the dictionary key is not found in the timeseries list item
    'Remove that item from the dictionary and also subtract the index by one
    For i = 0 To (pGroupIdDict.Count - 1)
        pkey = pGroupIdDictKeys(i)
        pIndexId = pGroupIdDict.Item(pkey)
        isIdFound = False
        For pIndex = 1 To ListTimeSeries.ListItems.Count
            If pIndexId = ListTimeSeries.ListItems.Item(pIndex) Then
                isIdFound = True
            End If
        Next pIndex
        If isIdFound = False Then 'then remove item from dictionary
            For pIndex = 1 To ListTimeSeries.ListItems.Count
                If ListTimeSeries.ListItems.Item(pIndex) > pIndexId Then
                    ListTimeSeries.ListItems.Item(pIndex) = ListTimeSeries.ListItems.Item(pIndex) - 1
                End If
            Next pIndex
            pGroupIdDict.Remove (pkey)
            pGroupIndex = pGroupIndex - 1
        End If
    Next
           
    
End Sub

Private Sub Form_Load()
    
On Error GoTo ShowError:
    'Create the dictionary to store group ids
    Set pGroupIdDict = CreateObject("Scripting.Dictionary")
     
    Dim pLandUseRLayer As IRasterLayer
    Set pLandUseRLayer = GetInputRasterLayer("Landuse")
    Dim pTable As iTable
    Set pTable = pLandUseRLayer
    Dim pLULookup As iTable
    Set pLULookup = GetInputDataTable("lulookup")
    Dim pLanduseDict As Scripting.Dictionary
    Set pLanduseDict = CreateObject("Scripting.Dictionary")
    Dim pCursor As ICursor
    Set pCursor = pLULookup.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        pLanduseDict.Add pRow.value(pCursor.FindField("LUCODE")), pRow.value(pCursor.FindField("LUNAME"))
        Set pRow = pCursor.NextRow
    Loop
    Set pRow = Nothing
    Set pCursor = Nothing
    
    Dim pRowCount As Integer
    pRowCount = pTable.RowCount(Nothing)
    Dim LandUseArray() As String
    ReDim LandUseArray(1 To 2, 1 To pRowCount)
    
    Set pCursor = pTable.Search(Nothing, True)
    Set pRow = pCursor.NextRow
    Dim pLandUse As Integer
    Dim StrLandUse As String
    pRowCount = 1
    
    Dim pQueryFilter As IQueryFilter
    Dim pSelRowCount As Long
    
    Dim pSWMMLUReclassTable As iTable
    Set pSWMMLUReclassTable = GetInputDataTable("LANDLUReclass")
    listLUType.ColumnHeaders.Add , , "Landuse Code", listLUType.Width * 0.4
    listLUType.ColumnHeaders.Add , , "Landuse Description", listLUType.Width * 0.6
    Dim itmX As ListItem
        
    pGroupIndex = 0
    
    Do While Not pRow Is Nothing
        pLandUse = pRow.value(pCursor.FindField("Value"))
        StrLandUse = pLandUse & "   " & pLanduseDict.Item(pLandUse)
        LandUseArray(1, pRowCount) = pLandUse
        LandUseArray(2, pRowCount) = pLanduseDict.Item(pLandUse)
        If (pSWMMLUReclassTable Is Nothing) Then
            Set itmX = listLUType.ListItems.Add(, , pLandUse)   ' LuCode.
            itmX.SubItems(1) = pLanduseDict.Item(pLandUse) ' Lu Description
            pRowCount = pRowCount + 1
        Else
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "LUCode = " & pLandUse
            pSelRowCount = pSWMMLUReclassTable.RowCount(pQueryFilter)
            If Not pSelRowCount >= 1 Then
                Set itmX = listLUType.ListItems.Add(, , pLandUse)   ' LuCode.
                itmX.SubItems(1) = pLanduseDict.Item(pLandUse) ' Lu Description
                pRowCount = pRowCount + 1
            End If
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    
    ListTimeSeries.ColumnHeaders.Add , , "LU Group Code", 0.001
    ListTimeSeries.ColumnHeaders.Add , , "Landuse Group", ListTimeSeries.Width * 50 / 250
    ListTimeSeries.ColumnHeaders.Add , , "Landuse Code", 0.001
    ListTimeSeries.ColumnHeaders.Add , , "Landuse Description", ListTimeSeries.Width * 95 / 250
    ListTimeSeries.ColumnHeaders.Add , , "Impervious ?", ListTimeSeries.Width * 50 / 250
    ListTimeSeries.ColumnHeaders.Add , , "Percentage", ListTimeSeries.Width * 50 / 250
        
    
    'Add Landuse Groups to the combo box
    cmbLUGroup.AddItem "Forest"
    cmbLUGroup.AddItem "Agriculture"
    cmbLUGroup.AddItem "High-Density-Residential"
    cmbLUGroup.AddItem "High-Density-Residential-PERVIOUS"
    cmbLUGroup.AddItem "High-Density-Residential-IMPERVIOUS"
    cmbLUGroup.AddItem "Medium-Density-Residential"
    cmbLUGroup.AddItem "Medium-Density-Residential-PERVIOUS"
    cmbLUGroup.AddItem "Medium-Density-Residential-IMPERVIOUS"
    cmbLUGroup.AddItem "Low-Density-Residential"
    cmbLUGroup.AddItem "Low-Density-Residential-PERVIOUS"
    cmbLUGroup.AddItem "Low-Density-Residential-IMPERVIOUS"
    cmbLUGroup.AddItem "Commercial"
    cmbLUGroup.AddItem "Commercial-PERVIOUS"
    cmbLUGroup.AddItem "Commercial-IMPERVIOUS"
    cmbLUGroup.AddItem "Road"
    cmbLUGroup.AddItem "Rooftop"
    cmbLUGroup.ListIndex = 0
    
    'Add Timeseries file list from the table : LuReClass -- Sabu Paul; Aug 24 2004
    
    If (pSWMMLUReclassTable Is Nothing) Then
        pGroupIndex = 1
        Exit Sub
    End If
    
    Dim pLUGroupIDindex As Long
    pLUGroupIDindex = pSWMMLUReclassTable.FindField("LUGroupID")
    Dim pLUGroupindex As Long
    pLUGroupindex = pSWMMLUReclassTable.FindField("LUGroup")
    Dim pLUCodeindex As Long
    pLUCodeindex = pSWMMLUReclassTable.FindField("LUCode")
    Dim pLUDescIndex As Long
    pLUDescIndex = pSWMMLUReclassTable.FindField("LUDescrip")
    Dim pPercentageIndex As Long
    pPercentageIndex = pSWMMLUReclassTable.FindField("Percentage")
    Dim pLUTypeindex As Long
    pLUTypeindex = pSWMMLUReclassTable.FindField("Impervious")
    
    Dim pLUCode As Integer
    Dim pLUType As String
    Dim pLUDescription As String
    Dim pLUPercent As Double
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
    Dim pTimeSeries As String
    Set pCursor = pSWMMLUReclassTable.Search(Nothing, True)
    Set pRow = pCursor.NextRow
    pRowCount = 1
    
    Dim pkey As String
    
    Do While Not pRow Is Nothing
        pLuGroupID = pRow.value(pLUGroupIDindex)
        pLUType = pRow.value(pLUTypeindex)
        pLUCode = pRow.value(pLUCodeindex)
        pLUDescription = pRow.value(pLUDescIndex)
        pLUPercent = pRow.value(pPercentageIndex) * 100 ' The table value is in fraction -- Sabu Paul
        pLuGroup = pRow.value(pLUGroupindex)
        Set itmX = ListTimeSeries.ListItems.Add(, , pLuGroupID)   ' LuCode.
        itmX.SubItems(1) = pLuGroup ' Lu Description
        itmX.SubItems(2) = pLUCode
        itmX.SubItems(3) = pLUDescription
        itmX.SubItems(4) = pLUType
        itmX.SubItems(5) = pLUPercent
        pRowCount = pRowCount + 1
        
        pkey = pLuGroup
        If Not pGroupIdDict.Exists(pkey) Then
            pGroupIdDict.Item(pkey) = pLuGroupID
            If pGroupIndex < pLuGroupID Then
                pGroupIndex = pLuGroupID
            End If
        End If
        
        Set pRow = pCursor.NextRow
    Loop
    pGroupIndex = pGroupIndex + 1
    GoTo Cleanup:
ShowError:
    MsgBox "Error loading the landuse reclass form", Err.description
Cleanup:
    Set pRow = Nothing
    Set itmX = Nothing
    Set pCursor = Nothing
    Set pSWMMLUReclassTable = Nothing
    Set pLanduseDict = Nothing
    Set pTable = Nothing
    Set pLULookup = Nothing
        
End Sub

Private Sub txtPercentage_Change()
    ValidatePercentImperviousValues
End Sub


Private Sub ValidatePercentImperviousValues()

    If Not (IsNumeric(Trim(txtPercentage.Text))) Then
        MsgBox "Percentage imperviousness number should be a valid number", vbExclamation
        Exit Sub
    End If
    If (CDbl(txtPercentage.Text) < 0 Or CDbl(txtPercentage.Text) > 100) Then
        MsgBox "Percentage imperviousness number should be within (0-100) range.", vbExclamation
        Exit Sub
    End If
  

End Sub


'******************************************************************************
'Subroutine: AddSWMMLanduseReclassification
'Author:     Mira Chokshi
'******************************************************************************
Public Sub AddSWMMLanduseReclassification(LandUseTextFile() As String)
On Error GoTo ShowError
    
    'Get landuse reclassification table: LUReclass, Create new if not found
    Dim pSWMMLUReclassTable As iTable
    Set pSWMMLUReclassTable = GetInputDataTable("LANDLUReclass")
    If (pSWMMLUReclassTable Is Nothing) Then
    'If the table is present delete and add new -- Sabu Paul, Aug 24, 2004
        Set pSWMMLUReclassTable = CreateLandUseReclassificationTable("LANDLUReclass")
        AddTableToMap pSWMMLUReclassTable
        Set pSWMMLUReclassTable = GetInputDataTable("LANDLUReclass")
    Else
        pSWMMLUReclassTable.DeleteSearchedRows Nothing    'delete all records
    End If
    
    
    Dim pLUGroupIDindex As Long
    pLUGroupIDindex = pSWMMLUReclassTable.FindField("LUGroupID")
    Dim pLUGroupindex As Long
    pLUGroupindex = pSWMMLUReclassTable.FindField("LUGroup")
    Dim pLUCodeindex As Long
    pLUCodeindex = pSWMMLUReclassTable.FindField("LUCode")
    Dim pLUDescIndex As Long
    pLUDescIndex = pSWMMLUReclassTable.FindField("LUDescrip")
    Dim pPercentageIndex As Long
    pPercentageIndex = pSWMMLUReclassTable.FindField("Percentage")
    Dim pLUTypeindex As Long
    pLUTypeindex = pSWMMLUReclassTable.FindField("Impervious")
    
    'Iterate over the entire array
    Dim pRow As iRow
    Dim pLUCode As Integer
    Dim pLUType As String
    Dim pLUDescription As String
    Dim pLUPercent As Double
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
           
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
        
        'add new row
        Set pRow = pSWMMLUReclassTable.CreateRow
        pRow.value(pLUGroupIDindex) = pLuGroupID
        pRow.value(pLUGroupindex) = pLuGroup
        pRow.value(pLUCodeindex) = pLUCode
        pRow.value(pLUTypeindex) = pLUType
        pRow.value(pLUDescIndex) = pLUDescription
        pRow.value(pPercentageIndex) = pLUPercent
        pRow.Store
    Next
    
    GoTo Cleanup:
ShowError:
    MsgBox "AddSWMMLanduseReclassification : " & Err.description
Cleanup:
    Set pSWMMLUReclassTable = Nothing
    Set pRow = Nothing
End Sub


