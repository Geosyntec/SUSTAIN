VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTSAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Series Assignment for Landuse"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTSAssign.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameSoilFracs 
      Caption         =   "Define Soil Fractions in Sediment (0-1)"
      Height          =   735
      Left            =   360
      TabIndex        =   25
      Top             =   4200
      Width           =   7455
      Begin VB.TextBox txtClay 
         Height          =   285
         Left            =   5400
         TabIndex        =   31
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSilt 
         Height          =   285
         Left            =   3000
         TabIndex        =   29
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSand 
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Clay"
         Height          =   255
         Left            =   5040
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Silt"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Sand"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "FrmTSAssign.frx":08CA
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   23
      Top             =   4200
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "FrmTSAssign.frx":0C0C
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton cmdFindPervTimeSeries 
      Caption         =   "..."
      Height          =   360
      Left            =   6360
      TabIndex        =   20
      Top             =   3120
      Width           =   600
   End
   Begin VB.TextBox txtPervTimeSeriesFile 
      Height          =   360
      Left            =   2520
      TabIndex        =   19
      Top             =   3120
      Width           =   3720
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "FrmTSAssign.frx":0F4E
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "FrmTSAssign.frx":1290
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      Picture         =   "FrmTSAssign.frx":15D2
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      Picture         =   "FrmTSAssign.frx":1914
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   360
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7440
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListTimeSeries 
      Height          =   2160
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   3810
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
      Height          =   1545
      Left            =   2640
      TabIndex        =   12
      Top             =   360
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   2725
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   6960
      TabIndex        =   0
      Top             =   1080
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   6960
      TabIndex        =   1
      Top             =   360
      Width           =   945
   End
   Begin VB.CommandButton cmdRemoveLandUseReclass 
      Height          =   400
      Left            =   4680
      Picture         =   "FrmTSAssign.frx":1C56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1560
   End
   Begin VB.CommandButton cmdAddLandUseReclass 
      Height          =   400
      Left            =   2520
      Picture         =   "FrmTSAssign.frx":3D68
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1560
   End
   Begin VB.TextBox txtPercentage 
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   840
   End
   Begin VB.ComboBox cmbLUGroup 
      Height          =   315
      ItemData        =   "FrmTSAssign.frx":57EA
      Left            =   2520
      List            =   "FrmTSAssign.frx":57EC
      TabIndex        =   8
      Top             =   2040
      Width           =   3720
   End
   Begin VB.CommandButton cmdFindImpTimeSeries 
      Caption         =   "..."
      Height          =   360
      Left            =   6360
      TabIndex        =   9
      Top             =   3720
      Width           =   600
   End
   Begin VB.TextBox txtImpTimeSeriesFile 
      Height          =   360
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   3720
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "FrmTSAssign.frx":57EE
      Top             =   5040
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Add New or Remove Existing Landuse Groups"
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Pervious Time Series File"
      Height          =   360
      Left            =   360
      TabIndex        =   21
      Top             =   3120
      Width           =   2040
   End
   Begin VB.Label Label6 
      Caption         =   "Select Input Landuse Types"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "% Pervious"
      Height          =   240
      Left            =   5880
      TabIndex        =   4
      Top             =   -240
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Landuse Group"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Percentage Imperviousness"
      Height          =   240
      Left            =   360
      TabIndex        =   7
      Top             =   2535
      Width           =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Impervious Time Series File"
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   2160
   End
End
Attribute VB_Name = "FrmTSAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
        'Enable & Disable the Time series controls.....
    If gInternalSimulation Then
        txtPervTimeSeriesFile.Enabled = False
        cmdFindPervTimeSeries.Enabled = False
        txtImpTimeSeriesFile.Enabled = False
        cmdFindImpTimeSeries.Enabled = False
    Else
        txtPervTimeSeriesFile.Enabled = True
        cmdFindPervTimeSeries.Enabled = True
        txtImpTimeSeriesFile.Enabled = True
        cmdFindImpTimeSeries.Enabled = True
    End If
    
End Sub

Private Sub cmdAddLandUseReclass_Click()

    Dim pGroupName As String
    Dim pImpKey As String
    Dim pPervKey As String
    
    pGroupName = Replace(cmbLUGroup.Text, " ", "_")
    pImpKey = "imp" & pGroupName
    pPervKey = "perv" & pGroupName
    
    Dim curIndex As Integer
    
    'On August 31, 20004 - Sabu Paul
    Dim pPervPerc As Double
    Dim pImpPerc As Double
    pPervPerc = 0
    pImpPerc = 0
    
    Dim pPervTSFile As String
    Dim pImpTSFile As String
    pPervTSFile = ""
    pImpTSFile = ""
    If Not (IsNumeric(Trim(txtPercentage.Text))) Then
        MsgBox "Percentage imperviousness number should be a valid number.", vbExclamation
        Exit Sub
    End If
    If (CDbl(txtPercentage.Text) < 0 Or CDbl(txtPercentage.Text) > 100) Then
        MsgBox "Percentage imperviousness number should be within (0-100) range.", vbExclamation
        Exit Sub
    End If
    
    If FrameSoilFracs.Enabled Then
        If Not (IsNumeric(Trim(txtSand.Text)) And IsNumeric(Trim(txtSilt.Text)) And IsNumeric(Trim(txtClay.Text))) Then
            MsgBox "Soil fractions should be valid numbers", vbExclamation
            Exit Sub
        End If
        
'        If (CDbl(txtSand.Text) + CDbl(txtSilt.Text) + CDbl(txtClay.Text) < 0 Or CDbl(txtSand.Text) + CDbl(txtSilt.Text) + CDbl(txtClay.Text) > 1) Then
'            MsgBox "Sum of soil fractions number should be equal to 1.0", vbExclamation
'            Exit Sub
'        End If
        If CDbl(txtSand.Text) + CDbl(txtSilt.Text) + CDbl(txtClay.Text) <> 1 Then
            MsgBox "Sum of soil fractions number should be equal to 1.0", vbExclamation
            Exit Sub
        End If
    End If
    
    pImpPerc = CDbl(Trim(txtPercentage.Text))
    pPervPerc = 100 - pImpPerc
    
    'If gExternalSimulation Then ' Now this form is only used for external simulation
    'If CDbl(Trim(txtPercentage.Text)) = 0 Then
''    If pPervPerc = 100 Then
''            If Trim(txtPervTimeSeriesFile.Text) = "" Then
''                MsgBox "Select the timeseries file for pervious land types", vbExclamation
''                Exit Sub
''            End If
''            pPervTSFile = txtPervTimeSeriesFile.Text
''    'ElseIf CDbl(Trim(txtPercentage.Text)) = 100 Then
''    ElseIf pImpPerc = 100 Then
''            If Trim(txtImpTimeSeriesFile.Text) = "" Then
''                MsgBox "Select the timeseries file for impervious land types", vbExclamation
''                Exit Sub
''            End If
''            pImpTSFile = txtImpTimeSeriesFile.Text
''    Else
''            If Trim(txtPervTimeSeriesFile.Text) = "" Or Trim(txtImpTimeSeriesFile.Text) = "" Then
''                MsgBox "Select the timeseries file for pervious and impervious land types", vbExclamation
''                Exit Sub
''            End If
''            pPervTSFile = txtPervTimeSeriesFile.Text
''            pImpTSFile = txtImpTimeSeriesFile.Text
''    End If
    'End If
    
    If pPervPerc > 0 Then
        If Trim(txtPervTimeSeriesFile.Text) = "" Then
            MsgBox "Select the timeseries file for pervious land types", vbExclamation
            Exit Sub
        End If
        pPervTSFile = txtPervTimeSeriesFile.Text
    End If
    If pImpPerc > 0 Then
        If Trim(txtImpTimeSeriesFile.Text) = "" Then
            MsgBox "Select the timeseries file for impervious land types", vbExclamation
            Exit Sub
        End If
        pImpTSFile = txtImpTimeSeriesFile.Text
    End If
    
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
            Dim itmX As ListItem
            'If pImpPerc <> 0 And pImpTSFile <> "" Then
            If pImpPerc <> 0 Then
                If pGroupIdDict.Exists(pImpKey) Then
                    curIndex = pGroupIdDict.Item(pImpKey)
                Else
                    pGroupIdDict.Item(pImpKey) = pGroupIndex
                    curIndex = pGroupIdDict.Item(pImpKey)
                    pGroupIndex = pGroupIndex + 1
                End If
                
                'Set itmX = ListTimeSeries.ListItems.Add(, , pGroupIndex)   ' LuCode.
                Set itmX = ListTimeSeries.ListItems.add(, , curIndex)   ' LuCode.
                itmX.SubItems(1) = pGroupName ' Lu Description
                itmX.SubItems(2) = luCode
                itmX.SubItems(3) = luDesc
                itmX.SubItems(4) = "1"
                itmX.SubItems(5) = pImpPerc
                itmX.SubItems(6) = pImpTSFile
                'Sand,Silt and Clay fraction
                itmX.SubItems(7) = txtSand.Text
                itmX.SubItems(8) = txtSilt.Text
                itmX.SubItems(9) = txtClay.Text
            End If
            'If pPervPerc <> 0 And pPervTSFile <> "" Then
            If pPervPerc <> 0 Then
                If pGroupIdDict.Exists(pPervKey) Then
                    curIndex = pGroupIdDict.Item(pPervKey)
                Else
                    pGroupIdDict.Item(pPervKey) = pGroupIndex
                    curIndex = pGroupIdDict.Item(pPervKey)
                    pGroupIndex = pGroupIndex + 1
                End If
                'Set itmX = ListTimeSeries.ListItems.Add(, , pGroupIndex)   ' LuCode.
                Set itmX = ListTimeSeries.ListItems.add(, , curIndex)   ' LuCode.
                itmX.SubItems(1) = pGroupName ' Lu Description
                itmX.SubItems(2) = luCode
                itmX.SubItems(3) = luDesc
                itmX.SubItems(4) = "0"
                itmX.SubItems(5) = pPervPerc
                itmX.SubItems(6) = pPervTSFile
                'Sand,Silt and Clay fraction
                itmX.SubItems(7) = txtSand.Text
                itmX.SubItems(8) = txtSilt.Text
                itmX.SubItems(9) = txtClay.Text
            End If
            listLUType.ListItems.Remove (pIndex)
        End If
        'luCode = pSelectedList.Item(pIndex).SubItems(0)
        
    Next pIndex
    
    '*** Clear the textbox values for percent pervious and impervious values
    txtImpTimeSeriesFile.Text = ""
    txtPervTimeSeriesFile.Text = ""
    cmbLUGroup.ListIndex = 0
    txtPercentage.Text = 0

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindImpTimeSeries_Click()
    'FrmLUReclass.CommonDialog.ShowOpen
    Me.CommonDialog.ShowOpen
    txtImpTimeSeriesFile.Text = Me.CommonDialog.FileName 'FrmLUReclass.CommonDialog.FileName
End Sub

Private Sub cmdFindPervTimeSeries_Click()
    Me.CommonDialog.ShowOpen 'FrmLUReclass.CommonDialog.ShowOpen
    txtPervTimeSeriesFile.Text = Me.CommonDialog.FileName 'FrmLUReclass.CommonDialog.FileName
End Sub

Private Sub cmdOk_Click()
    If listLUType.ListItems.Count > 0 Then
        MsgBox "Assign timeseries files for all landuse types", vbExclamation
        Exit Sub
    End If
    Dim pLuGroupTSDict As Scripting.Dictionary
    Set pLuGroupTSDict = CreateObject("Scripting.Dictionary")
    Dim pIndex As Integer
    Dim isImp As Integer
    Dim pGroupId As String
    Dim pTsFile As String
    Dim pkey As String
    
    For pIndex = 1 To ListTimeSeries.ListItems.Count
        pGroupId = ListTimeSeries.ListItems.Item(pIndex).SubItems(1)
        isImp = ListTimeSeries.ListItems.Item(pIndex).SubItems(4)
        pTsFile = ListTimeSeries.ListItems.Item(pIndex).SubItems(6)
        If isImp = 1 Then
            pkey = "imp" & pGroupId
        Else
            pkey = "perv" & pGroupId
        End If
        If pLuGroupTSDict.Exists(pkey) Then
            If pLuGroupTSDict.Item(pkey) <> pTsFile Then
                MsgBox "Landuse group timeseries file combination should be same for " & pGroupId
                Exit Sub
            End If
        Else
            pLuGroupTSDict.add pkey, pTsFile
        End If
    Next pIndex
    
    ReDim LandUseTextFiles(1 To 10, 1 To 1)
    LandUseTextFiles(1, 1) = "LU Group Code"
    LandUseTextFiles(2, 1) = "LU Group"
    LandUseTextFiles(3, 1) = "Landuse Code"
    LandUseTextFiles(4, 1) = "Landuse Description"
    LandUseTextFiles(5, 1) = "Impervious ?"
    LandUseTextFiles(6, 1) = "Percentage"
    LandUseTextFiles(7, 1) = "Time Series File"
    'Sand,Silt and Clay fraction
    LandUseTextFiles(8, 1) = "Sand Fraction"
    LandUseTextFiles(9, 1) = "Silt Fraction"
    LandUseTextFiles(10, 1) = "Clay Fraction"
    
    Dim i As Integer
    For pIndex = 1 To ListTimeSeries.ListItems.Count
        ReDim Preserve LandUseTextFiles(1 To 10, 1 To pIndex + 1)
        LandUseTextFiles(1, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex)
'        LandUseTextFiles(2, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(1)
'        LandUseTextFiles(3, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(2)
'        LandUseTextFiles(4, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(3)
'        LandUseTextFiles(5, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(4)
'        LandUseTextFiles(6, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(5)
'        LandUseTextFiles(7, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(6)
        
        'Sand,Silt and Clay fraction
        For i = 1 To 9
            LandUseTextFiles(i + 1, pIndex + 1) = ListTimeSeries.ListItems.Item(pIndex).SubItems(i)
        Next
    Next pIndex
    
    AddTimeSeriesAssignments LandUseTextFiles
    Unload Me
    
    'FrmTSSoilFractions.Show vbModal
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
                Set itmX = listLUType.ListItems.add(, , luCode)   ' LuCode.
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
    
On Error GoTo ShowError
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    'Create the dictionary to store group ids
    Set pGroupIdDict = CreateObject("Scripting.Dictionary")
    
    Dim pPollTable As iTable
    Set pPollTable = GetInputDataTable("Pollutants")
        
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = " Sediment = 'SEDIMENT'"
    
    'Only if the sediment is defined as one of the pollutants, then enable SoilFraction frame
    If pPollTable.RowCount(pQueryFilter) > 0 Then
        FrameSoilFracs.Enabled = True
    Else
        FrameSoilFracs.Enabled = False
    End If
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
        pLanduseDict.add pRow.value(pCursor.FindField("LUCODE")), pRow.value(pCursor.FindField("LUNAME"))
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
    
    Dim pSelRowCount As Long
    
    Dim pLUReClasstable As iTable
    Set pLUReClasstable = GetInputDataTable("TSAssigns")
    listLUType.ColumnHeaders.add , , "Landuse Code", listLUType.Width * 0.4
    listLUType.ColumnHeaders.add , , "Landuse Description", listLUType.Width * 0.6
    Dim itmX As ListItem
        
    pGroupIndex = 0
    
    Do While Not pRow Is Nothing
        pLandUse = pRow.value(pCursor.FindField("Value"))
        StrLandUse = pLandUse & "   " & pLanduseDict.Item(pLandUse)
        LandUseArray(1, pRowCount) = pLandUse
        LandUseArray(2, pRowCount) = pLanduseDict.Item(pLandUse)
        If (pLUReClasstable Is Nothing) Then
            Set itmX = listLUType.ListItems.add(, , pLandUse)   ' LuCode.
            itmX.SubItems(1) = pLanduseDict.Item(pLandUse) ' Lu Description
            pRowCount = pRowCount + 1
        Else
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "LUCode = " & pLandUse
            pSelRowCount = pLUReClasstable.RowCount(pQueryFilter)
            If Not pSelRowCount >= 1 Then
                Set itmX = listLUType.ListItems.add(, , pLandUse)   ' LuCode.
                itmX.SubItems(1) = pLanduseDict.Item(pLandUse) ' Lu Description
                pRowCount = pRowCount + 1
            End If
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    'listLUType.Column = LandUseArray
    'listLUType.ListIndex = 0

    
    ListTimeSeries.ColumnHeaders.add , , "LU Group Code", 0.001
    ListTimeSeries.ColumnHeaders.add , , "Landuse Group", ListTimeSeries.Width * 45 / 395
    ListTimeSeries.ColumnHeaders.add , , "Landuse Code", 0.001
    ListTimeSeries.ColumnHeaders.add , , "Landuse Description", ListTimeSeries.Width * 100 / 395
    ListTimeSeries.ColumnHeaders.add , , "Impervious ?", ListTimeSeries.Width * 55 / 395
    ListTimeSeries.ColumnHeaders.add , , "Percentage", ListTimeSeries.Width * 50 / 395
    ListTimeSeries.ColumnHeaders.add , , "Time Series File", ListTimeSeries.Width * 145 / 395
    
    'Sand, Silt and Clay fraction
    ListTimeSeries.ColumnHeaders.add , , "Sand Fraction", ListTimeSeries.Width * 145 / 395
    ListTimeSeries.ColumnHeaders.add , , "Silt Fraction", ListTimeSeries.Width * 145 / 395
    ListTimeSeries.ColumnHeaders.add , , "Clay Fraction", ListTimeSeries.Width * 145 / 395
    
    
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
    
    If (pLUReClasstable Is Nothing) Then
        pGroupIndex = 1
        Exit Sub
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
    Dim pTimeSeriesindex As Long
    pTimeSeriesindex = pLUReClasstable.FindField("TimeSeries")
    
    'Sand, Silt and Clay fraction
    Dim pSandFracindex As Long
    pSandFracindex = pLUReClasstable.FindField("SandFrac")
    Dim pSiltFracindex As Long
    pSiltFracindex = pLUReClasstable.FindField("SiltFrac")
    Dim pClayFracindex As Long
    pClayFracindex = pLUReClasstable.FindField("ClayFrac")
    
    Dim pLUCode As Integer
    Dim pLUType As String
    Dim pLUDescription As String
    Dim pLUPercent As Double
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
    Dim pTimeSeries As String
    'Sand, Silt and Clay fraction
    Dim pSandFrac As Double
    Dim pSiltFrac As Double
    Dim pClayFrac As Double
    
    Set pCursor = pLUReClasstable.Search(Nothing, True)
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
        pTimeSeries = pRow.value(pTimeSeriesindex)
        'Sand, Silt and Clay
        pSandFrac = pRow.value(pSandFracindex)
        pSiltFrac = pRow.value(pSiltFracindex)
        pClayFrac = pRow.value(pClayFracindex)
        
        Set itmX = ListTimeSeries.ListItems.add(, , pLuGroupID)   ' LuCode.
        itmX.SubItems(1) = pLuGroup ' Lu Description
        itmX.SubItems(2) = pLUCode
        itmX.SubItems(3) = pLUDescription
        itmX.SubItems(4) = pLUType
        itmX.SubItems(5) = pLUPercent
        itmX.SubItems(6) = pTimeSeries
        'Sand, Silt and Clay
        itmX.SubItems(7) = pSandFrac
        itmX.SubItems(8) = pSiltFrac
        itmX.SubItems(9) = pClayFrac
        pRowCount = pRowCount + 1
        
        If pLUType = 1 Then
            pkey = "imp" & pLuGroup
        Else
            pkey = "perv" & pLuGroup
        End If
        If Not pGroupIdDict.Exists(pkey) Then
            pGroupIdDict.Item(pkey) = pLuGroupID
            If pGroupIndex < pLuGroupID Then
                pGroupIndex = pLuGroupID
            End If
        End If
        
        Set pRow = pCursor.NextRow
    Loop
    pGroupIndex = pGroupIndex + 1
    
    GoTo CleanUp
ShowError:
    MsgBox "Error loading the landuse reclass form", Err.description
CleanUp:
    Set pRow = Nothing
    Set itmX = Nothing
    Set pCursor = Nothing
    Set pLUReClasstable = Nothing
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
    
    If CDbl(Trim(txtPercentage.Text)) = 0 Then
        cmdFindImpTimeSeries.Enabled = False
        cmdFindPervTimeSeries.Enabled = True
        txtPervTimeSeriesFile.Enabled = True
        txtPervTimeSeriesFile.BackColor = vbWhite
        txtImpTimeSeriesFile.Enabled = False
        txtImpTimeSeriesFile.BackColor = &H80000016
    ElseIf CDbl(Trim(txtPercentage.Text)) = 100 Then
        cmdFindImpTimeSeries.Enabled = True
        cmdFindPervTimeSeries.Enabled = False
        txtImpTimeSeriesFile.Enabled = True
        txtImpTimeSeriesFile.BackColor = vbWhite
        txtPervTimeSeriesFile.Enabled = False
        txtPervTimeSeriesFile.BackColor = &H80000016
    Else
        cmdFindImpTimeSeries.Enabled = True
        cmdFindPervTimeSeries.Enabled = True
        txtImpTimeSeriesFile.Enabled = True
        txtImpTimeSeriesFile.BackColor = vbWhite
        txtPervTimeSeriesFile.Enabled = True
        txtPervTimeSeriesFile.BackColor = vbWhite
    End If

End Sub

