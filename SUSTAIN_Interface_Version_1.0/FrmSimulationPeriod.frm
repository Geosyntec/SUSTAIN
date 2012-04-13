VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmSimulationPeriod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Simulation Options"
   ClientHeight    =   7410
   ClientLeft      =   4860
   ClientTop       =   3585
   ClientWidth     =   8400
   Icon            =   "FrmSimulationPeriod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameMultiTier 
      Caption         =   "For Multi-tier Simulation"
      Height          =   615
      Left            =   240
      TabIndex        =   36
      Top             =   3120
      Width           =   6615
      Begin VB.OptionButton OptTier2 
         Caption         =   "Tier 2"
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OptTier1 
         Caption         =   "Tier 1"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame FrameOptimizationOptions 
      Caption         =   "Optimization Options"
      Height          =   615
      Left            =   240
      TabIndex        =   32
      Top             =   2520
      Width           =   6615
      Begin VB.OptionButton optNSGAII 
         Caption         =   "NSGAII"
         Height          =   255
         Left            =   4920
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optScatter 
         Caption         =   "Scatter Search"
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Select Optimization Technique"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kinematic Wave Numerical Solution Parameters (for VFSMOD)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   6615
      Begin VB.Frame Frame7 
         Caption         =   "Solution Method (KPG)"
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   6375
         Begin VB.OptionButton optKPG1 
            Caption         =   "Petrov-Galerkin solution"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton optKPG0 
            Caption         =   "Regular Finite Element"
            Height          =   195
            Left            =   3120
            TabIndex        =   25
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.TextBox MAXITER 
         Height          =   285
         Left            =   4800
         TabIndex        =   23
         Text            =   "150"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.TextBox CR 
         Height          =   285
         Left            =   4800
         TabIndex        =   22
         Text            =   "0.6"
         ToolTipText     =   "Between 0.5 - 0.8"
         Top             =   2520
         Width           =   1000
      End
      Begin VB.TextBox NPOL 
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Text            =   "3"
         Top             =   1335
         Width           =   1000
      End
      Begin VB.TextBox THETAW 
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Text            =   "0.5"
         ToolTipText     =   "0.5 is recommended"
         Top             =   840
         Width           =   1000
      End
      Begin VB.TextBox N 
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Text            =   "99"
         Top             =   360
         Width           =   1000
      End
      Begin VB.Label Label19 
         Caption         =   "Maximum Iterations (MAXITER)"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "Courant Number (CR)"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Number of Element Nodal Points (NPOL)"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1335
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "Time Weight factor (THETAW)"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label14 
         Caption         =   "Number of Nodes in Solution Domain (N)"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.ComboBox cmbPreDevLanduse 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "..."
      Height          =   315
      Left            =   6240
      TabIndex        =   15
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtOutputFolder 
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   315
      Left            =   6240
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtSimulation 
      Height          =   315
      Left            =   2760
      TabIndex        =   11
      Top             =   1080
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7080
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbOutTS 
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      Text            =   "Hourly"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtBMPTS 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Text            =   "5"
      Top             =   600
      Width           =   975
   End
   Begin MSComCtl2.DTPicker endDate 
      Height          =   315
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   75235331
      CurrentDate     =   36891.9583333333
   End
   Begin MSComCtl2.DTPicker startDate 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   75235331
      CurrentDate     =   32874
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblPreDevLanduse 
      Caption         =   "Define Pre-development Landuse Type (only for external land simulation option) :"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "Define Input File:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1110
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Define Output Folder:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1590
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Output Time Step"
      Height          =   195
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "BMP Simulation Time Step (Minutes, 1- 60)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmSimulationPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pLanduseDict As Scripting.Dictionary
Private bVFSModeled As Boolean

Private Sub cmdCancel_Click()
    Unload Me
    Set pLanduseDict = Nothing
    pInputFileName = ""
End Sub

Private Sub cmdFile_Click()
    CommonDialog.Filter = "Input File (*.inp)|*.inp"
    CommonDialog.ShowSave
    txtSimulation = CommonDialog.FileName
End Sub

Private Sub cmdOk_Click()
    gHasInFileError = True
    If startDate.value >= endDate.value Then
        MsgBox "End date should be after start date"
        Exit Sub
    End If
    If CInt(txtBMPTS.Text) < 1 Or CInt(txtBMPTS.Text) > 60 Then
        MsgBox "BMP time step should be between 1 and 60"
        Exit Sub
    End If
    If (Trim(txtSimulation.Text) = "") Then
        MsgBox "Please specify input template file."
        Exit Sub
    End If
    If (Trim(txtOutputFolder.Text) = "") Then
        MsgBox "Please specify output folder for simulation."
        Exit Sub
    End If
    
    '** Get Values for SWMM Simulation
''    If (optionExternalLanduse.value = True) Then
''        pLanduseSimulationOption = 0
''    ElseIf (optionInternalLanduse.value = True) Then
''        pSWMMLanduseOutflowFile = Trim(txtSWMMLanduseOutflow.Text)
''        pSWMMPreDevOutflowFile = Trim(txtSWMMPreDevOutflow.Text)
''        '** check if landuse and predeveloped outflow file are specified.
''        If (pSWMMLanduseOutflowFile = "") Then
''            MsgBox "Please select SWMM Landuse outflow file to continue."
''            Exit Sub
''        End If
''        If (pSWMMPreDevOutflowFile = "") Then
''            MsgBox "Please select SWMM predeveloped landuse outflow file to continue."
''            Exit Sub
''        End If
''        '** assign the file header option
''        pLanduseSimulationOption = 1
''    End If


    '** Get Values for SWMM Simulation
    If (gExternalSimulation) Then
        pLanduseSimulationOption = 0
    Else
        'Get the following information from other source
        'pSWMMLanduseOutflowFile = "dummy" 'Trim(txtSWMMLanduseOutflow.Text)
        'pSWMMPreDevOutflowFile = "dummy" 'Trim(txtSWMMPreDevOutflow.Text)
        
        Dim pSWMMOptionsTable As iTable
        Set pSWMMOptionsTable = GetInputDataTable("LANDOptions")
        
        If pSWMMOptionsTable Is Nothing Then
            MsgBox "Missing LANDOptions table"
            Exit Sub
        End If
        Dim iPropName As Long
        Dim iPropValue As Long
        iPropName = pSWMMOptionsTable.FindField("PropName")
        iPropValue = pSWMMOptionsTable.FindField("PropValue")
        
        Dim pPropertyName As String
        Dim pPropertyValue As String
        
        Dim pCursor As ICursor
        Dim pRow As iRow
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        
        pQueryFilter.WhereClause = "PropName = 'SAVE OUTFLOWS'"
        Set pCursor = pSWMMOptionsTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        pSWMMPreDevOutflowFile = Trim(pRow.value(iPropValue))
        
        pQueryFilter.WhereClause = "PropName = 'SAVE POST OUTFLOWS'"
        Set pCursor = pSWMMOptionsTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        pSWMMLanduseOutflowFile = Trim(pRow.value(iPropValue))
        
        '** assign the file header option
        pLanduseSimulationOption = 1
    End If
    
    'Get the input file name
    pInputFileName = txtSimulation.Text
    'Define the output folder name
    pOutputFolder = txtOutputFolder.Text

    'Write values back to remember
    txtSimulation.Text = pInputFileName
    txtOutputFolder.Text = pOutputFolder

    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim bContinue
    If (fso.FileExists(pInputFileName)) Then
        bContinue = MsgBox("Do you want to overwrite existing file ?", vbYesNo)
        If (bContinue = vbNo) Then
            Exit Sub
        End If
    End If

    'Date should be printed as YYYY MM DD
    gStrStartDate = Year(startDate.value) & vbTab & Month(startDate.value) & vbTab & Day(startDate.value)
    gStrEndDate = Year(endDate.value) & vbTab & Month(endDate.value) & vbTab & Day(endDate.value)
    
    'Following line was modified to fit to the new card - Sabu Paul, June 13,2007
    'strTimeStepLine = pLanduseSimulationOption & vbTab & txtBMPTS.Text
    strTimeStepLine = txtBMPTS.Text
    If cmbOutTS.Text = "Hourly" Then
        strTimeStepLine = strTimeStepLine & vbTab & "1"
    Else
        strTimeStepLine = strTimeStepLine & vbTab & "0" 'if its daily
    End If
    'Following line was added to fit to the new card - Sabu Paul, June 13,2007
    strTimeStepLine = strTimeStepLine & vbTab & pOutputFolder
    
    ' get the predeveloped landuse type
    If pLanduseSimulationOption = 0 Then _
        pPredevelopedLanduse = pLanduseDict.Item(cmbPredevLanduse.Text)

    'Following is moved below
'    'Save input file and output folder name
'    UpdateSimulationParameters pInputFileName, pOutputFolder
    
    If FrameOptimizationOptions.Enabled Then
        Dim techniqueValue As Integer
        If optScatter.value Then
            techniqueValue = 1
        Else
            techniqueValue = 2
        End If
'        Dim numBreaks As Integer
'         If (Trim(txtNumBreaks.Text) = "" Or Not IsNumeric(txtNumBreaks.Text)) Then
'            MsgBox "Please specify integer value for number of breaks."
'            Exit Sub
'        End If
'        numBreaks = CInt(txtNumBreaks.Text)
        'Save input file and output folder name
        
        UpdateSimulationParameters pInputFileName, pOutputFolder, techniqueValue, , Format(startDate.value, "MM/DD/YYYY"), Format(endDate.value, "MM/DD/YYYY") ', numBreaks
    Else
        'Save input file and output folder name
        UpdateSimulationParameters pInputFileName, pOutputFolder, , , Format(startDate.value, "MM/DD/YYYY"), Format(endDate.value, "MM/DD/YYYY")
    End If
    If bVFSModeled Then
        Dim N_num As Integer
        Dim THETAW_num As Double
        Dim NPOL_num As Integer
        Dim KPG_num As Integer
        Dim CR_num As Double
        Dim MAXITER_num As Integer
        
        Dim IELOUT As Integer
        IELOUT = 0
        
        If (Trim(N.Text) = "" Or Not IsNumeric(N.Text)) Then
            MsgBox "Please specify integer value for N."
            Exit Sub
        End If
        If (Trim(THETAW.Text) = "" Or Not IsNumeric(THETAW.Text)) Then
            MsgBox "Please specify real value for THETAW."
            Exit Sub
        End If
        If (Trim(NPOL.Text) = "" Or Not IsNumeric(NPOL.Text)) Then
            MsgBox "Please specify integer value for NPOL."
            Exit Sub
        End If
        If (Trim(CR.Text) = "" Or Not IsNumeric(CR.Text)) Then
            MsgBox "Please specify real value (05. to 0.8) for CR."
            Exit Sub
        End If

        If (Trim(MAXITER.Text Or Not IsNumeric(MAXITER.Text)) = "") Then
            MsgBox "Please specify integer value for MAXITER."
            Exit Sub
        End If
        
        N_num = CInt(Trim(N.Text))
        THETAW_num = CDbl(Trim(THETAW.Text))
        NPOL_num = CInt(Trim(NPOL.Text))
        CR_num = CDbl(Trim(CR.Text))
        MAXITER_num = CInt(Trim(MAXITER.Text))
        
        If optKPG0.value = True Then
            KPG_num = 0
        Else
            KPG_num = 1
        End If
        
        DefineVFSSimulationOptions N_num, THETAW_num, NPOL_num, KPG_num, CR_num, MAXITER_num, IELOUT
    End If
    
    'set dictionary to nothing
    Set pLanduseDict = Nothing
    
    gHasInFileError = False
    'Close the form
    Unload Me

End Sub

Private Sub cmdOutput_Click()
On Error GoTo ShowError
    Dim strTmpDir As String
   'now fill the strPath with the choice by user
    'strTmpDir = BrowseForFolder(0, "Select the folder to save output files")
    strTmpDir = BrowseForSpecificFolder("Select the folder to save output files", gApplicationPath)
    
    If (Trim(strTmpDir) <> "") Then
        txtOutputFolder.Text = strTmpDir
    End If
    Exit Sub
ShowError:
    MsgBox "cmdOutput_Click :" & Err.description
End Sub

''Private Sub cmdSWMMLanduse_Click()
''    CommonDialog.Filter = "Outflow File (*.txt)|*.txt"
''    CommonDialog.ShowSave
''    txtSWMMLanduseOutflow = CommonDialog.FileName
''
''    '** update the dates
''    UserSelectedInternalSelection
''End Sub
''
''Private Sub cmdSWMMPreDevOutflow_Click()
''    CommonDialog.Filter = "Outflow File (*.txt)|*.txt"
''    CommonDialog.ShowSave
''    txtSWMMPreDevOutflow = CommonDialog.FileName
''End Sub

Private Sub Form_Load()
On Error GoTo ShowError
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    Dim pVFSFClass As IFeatureClass
    Dim pVFSCount As Integer
    pVFSCount = 0
    If Not (pVFSFLayer Is Nothing) Then
        Set pVFSFClass = pVFSFLayer.FeatureClass
        pVFSCount = pVFSFClass.FeatureCount(Nothing)    ' Get vfs feature count
    End If
    bVFSModeled = False
    If pVFSCount > 0 Then
        bVFSModeled = True
        Me.Height = 7800 '8400
    Else
        Me.Height = 4200
    End If

    Dim pOptTable As iTable
    Set pOptTable = GetInputDataTable("OptimizationDetail")
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'Option'"
    Dim pCursor As ICursor
    Set pCursor = pOptTable.Search(pQueryFilter, False)
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
 
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Dim optmizationOption As Integer
    optmizationOption = pRow.value(iPropValue)
    
    FrameMultiTier.Visible = False
    If optmizationOption = 3 Then FrameMultiTier.Visible = True
    
    If optmizationOption <> 0 Then
        FrameOptimizationOptions.Enabled = True
''        optScatter.value = True
''        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'Technique'"
''        Set pCursor2 = pOptTable.Search(pQueryFilter, False)
''        Set pRow2 = pCursor2.NextRow
''        If Not pRow2 Is Nothing Then
''            If pRow2.value(iPropValue) = 2 Then
''                optNSGAII.value = True
''            End If
''        End If
        
        If optmizationOption = 1 Then
            optScatter.value = True
            optNSGAII.Enabled = False
        Else
            optNSGAII.value = True
            optScatter.Enabled = False
        End If
        
'        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'NumBreak'"
'        Set pCursor2 = pOptTable.Search(pQueryFilter, False)
'        Set pRow2 = pCursor2.NextRow
'        If Not pRow2 Is Nothing Then
'            txtNumBreaks.Text = pRow2.value(iPropValue)
'        End If
    Else
        FrameOptimizationOptions.Enabled = False
        optNSGAII.Enabled = False
        optScatter.Enabled = False
    End If
    
    Dim pSWMMTable As iTable
    Dim pExternalTable As iTable
    
    'Modified the following
''    If ((pExternalTable Is Nothing) And (pSWMMTable Is Nothing)) Then
''        MsgBox "LUReclass/SWMMLUReclass table not found."
''        Exit Sub
''    End If
    
    
    '** check if the tables are present
''    If (pSWMMTable Is Nothing) Then
''        '** disable the radio button for internal simulation
''        optionInternalLanduse.value = False
''        optionInternalLanduse.Enabled = False
''        optionExternalLanduse.value = True
''        optionExternalLanduse.Enabled = True
''        lblPreDevLanduse.Enabled = True
''        cmbPredevLanduse.Enabled = True
''    End If
''
''    If (pExternalTable Is Nothing) Then
''        '** disable the radio button for external simulation
''        optionInternalLanduse.value = True
''        optionInternalLanduse.Enabled = True
''        optionExternalLanduse.value = False
''        optionExternalLanduse.Enabled = False
''        lblPreDevLanduse.Enabled = False
''        cmbPredevLanduse.Enabled = False
''    End If
''
''    If ((Not pSWMMTable Is Nothing) And (Not pExternalTable Is Nothing)) Then
''        optionExternalLanduse.Enabled = True
''        optionExternalLanduse.value = True
''        Call UserSelectedExternalSelection
''
''    End If

    If gExternalSimulation Then
        Set pExternalTable = GetInputDataTable("TSAssigns")
        If pExternalTable Is Nothing Then
            MsgBox "TSAssigns table not found."
            Exit Sub
        End If
        lblPreDevLanduse.Enabled = True
        cmbPredevLanduse.Enabled = True
        Call UserSelectedExternalSelection
    Else
        Set pSWMMTable = GetInputDataTable("LUReclass")
        If pSWMMTable Is Nothing Then
            MsgBox "LUReclass table not found."
            Exit Sub
        End If
        lblPreDevLanduse.Enabled = False
        cmbPredevLanduse.Enabled = False
        Call UserSelectedInternalSelection
    End If
   
    GoTo CleanUp
    
ShowError:
    gHasInFileError = True
    MsgBox "FrmSimulationPeriod load failed.", Err.description

CleanUp:
    Set pVFSFLayer = Nothing
    Set pVFSFClass = Nothing
    Set pOptTable = Nothing
    Set pQueryFilter = Nothing
    Set pSWMMTable = Nothing
    Set pExternalTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub


Private Sub UserSelectedExternalSelection()
    Dim pTable As iTable
    Set pTable = GetInputDataTable("TSAssigns")

    '** variables to iterate the table
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iTimeSeries As Long
    iTimeSeries = pCursor.FindField("TimeSeries")
    Dim iLuGroup As Long
    iLuGroup = pCursor.FindField("LUGroup")
    Dim iLuGroupID As Long
    iLuGroupID = pCursor.FindField("LUGroupID")
    Set pLanduseDict = CreateObject("Scripting.Dictionary")

    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim pTimeSeriesFile As String
    Dim pLuGroupID As Integer
    Dim pLuGroup As String

    Do While Not (pRow Is Nothing)
        pTimeSeriesFile = pRow.value(iTimeSeries)
        pLuGroupID = pRow.value(iLuGroupID)
        pLuGroup = pRow.value(iLuGroup)
        If (Not pLanduseDict.Exists(pLuGroup)) Then
            pLanduseDict.add pLuGroup, pLuGroupID
            'Add to the predeveloped landuse combo control
            FrmSimulationPeriod.cmbPredevLanduse.AddItem pLuGroup
        End If
        'End If
        Set pRow = pCursor.NextRow
    Loop
    FrmSimulationPeriod.cmbPredevLanduse.ListIndex = 0

    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pTS As TextStream
    Set pTS = fso.OpenTextFile(pTimeSeriesFile, ForReading, False, TristateUseDefault)

    Dim fileLine
    Dim fileLines
    Dim startLine As String
    Dim endLine As String

    startLine = ""
    fileLine = pTS.ReadLine
    Do While Not pTS.AtEndOfStream
        If (InStr(1, fileLine, "Date/time", vbTextCompare) > 1) Then
            fileLine = pTS.ReadLine
            startLine = fileLine
            Exit Do
        End If
        fileLine = pTS.ReadLine
    Loop

    Dim startArray
    Dim endArray
    startArray = CustomSplit(startLine)
    
    If UBound(startArray) = 0 Then
        Do While Not pTS.AtEndOfStream
            If UBound(startArray) > 0 Then
                Exit Do
            End If
            fileLine = pTS.ReadLine
            startLine = fileLine
            startArray = CustomSplit(startLine)
        Loop
    End If
    
    fileString = pTS.ReadAll
    fileLines = Split(fileString, vbNewLine)
    endLine = fileLines(UBound(fileLines) - 1)
    endArray = CustomSplit(endLine)

    Dim strStartDate As String
    Dim strEndDate As String
    '2: 1998 12 31  9  0
    ' strStartDate = startArray(2) & "/" & startArray(3) & "/" & startArray(1) & " 12:00:00"
    ' strEndDate = endArray(2) & "/" & endArray(3) & "/" & endArray(1) & " 12:00:00"
    strStartDate = startArray(2) & "/" & startArray(3) & "/" & startArray(1)
    strEndDate = endArray(2) & "/" & endArray(3) & "/" & endArray(1)

    startDate.value = strStartDate
    endDate.value = strEndDate
    startDate.MinDate = strStartDate    'Set date limitation
    endDate.MinDate = strStartDate
    startDate.MaxDate = strEndDate     'Set date limitation
    endDate.MaxDate = strEndDate
    txtWSTS = 1
    txtBMPTS = 5
    cmbOutTS.AddItem "Hourly", 0
    cmbOutTS.AddItem "Daily", 1

    Set fso = Nothing
    Set pTS = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing

End Sub
Private Sub UserSelectedInternalSelection()
On Error GoTo ErrorHandler
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDOptions")
    
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    
    Dim strStartDate As String
    Dim strEndDate As String
    
    Dim pQueryFilter As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'START_DATE'"
    Set pCursor = pTable.Search(pQueryFilter, True)
    
    Set pRow = pCursor.NextRow
    If Not pRow Is Nothing Then
        strStartDate = pRow.value(iPropValue)
    End If
    
    pQueryFilter.WhereClause = "PropName = 'END_DATE'"
    Set pCursor = pTable.Search(pQueryFilter, True)
    
    Set pRow = pCursor.NextRow
    If Not pRow Is Nothing Then
        strEndDate = pRow.value(iPropValue)
    End If
    
    If strStartDate <> "" And strEndDate <> "" Then
        startDate.value = strStartDate
        endDate.value = strEndDate
        startDate.MinDate = strStartDate    'Set date limitation
        endDate.MinDate = strStartDate
        startDate.MaxDate = strEndDate     'Set date limitation
        endDate.MaxDate = strEndDate
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in Updating the date: " & Err.description
End Sub

''Private Sub UserSelectedInternalSelection()
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("LANDLUReclass")
''
''    '** variables to iterate the table
''    Dim pCursor As ICursor
''    Set pCursor = pTable.Search(Nothing, True)
''    Dim iLuGroup As Long
''    iLuGroup = pCursor.FindField("LUGroup")
''    Dim iLuGroupID As Long
''    iLuGroupID = pCursor.FindField("LUGroupID")
''    Set pLanduseDict = CreateObject("Scripting.Dictionary")
''
''    Dim pRow As iRow
''    Set pRow = pCursor.NextRow
''    Dim pTimeSeriesFile As String
''    Dim pLuGroupID As Integer
''    Dim pLuGroup As String
''
''    Do While Not (pRow Is Nothing)
''        pLuGroupID = pRow.value(iLuGroupID)
''        pLuGroup = pRow.value(iLuGroup)
''        If (Not pLanduseDict.Exists(pLuGroup)) Then
''            pLanduseDict.Add pLuGroup, pLuGroupID
''            'Add to the predeveloped landuse combo control
''            FrmSimulationPeriod.cmbPredevLanduse.AddItem pLuGroup
''        End If
''        'End If
''        Set pRow = pCursor.NextRow
''    Loop
''    FrmSimulationPeriod.cmbPredevLanduse.ListIndex = 0
''
''    '** check if outflow file is defined
''    Dim pLanduseOutflowFile As String
''    pLanduseOutflowFile = Trim(txtSWMMLanduseOutflow.Text)
''    If (pLanduseOutflowFile = "") Then
''        GoTo CleanUp
''    End If
''
''    '** open the outflow file, read values and load the dates
''    Dim fso As Scripting.FileSystemObject
''    Set fso = CreateObject("Scripting.FileSystemObject")
''    If (Not fso.FileExists(pLanduseOutflowFile)) Then
''        MsgBox "LAND Landuse Outflow file does not exist."
''        GoTo CleanUp
''    End If
''
''    Dim pTextStream As TextStream
''    Set pTextStream = fso.OpenTextFile(pLanduseOutflowFile, ForReading, False, TristateUseDefault)
''
''    '** parse the file to read values
''    Dim pString As String
''    Dim pColumns
''    Dim pYear, pMonth, pDay
''    Dim pStrStartDate, pStrEndDate
''    Dim doContinue As Boolean
''    doContinue = True
''    Do While doContinue
''        pString = pTextStream.ReadLine
''        '*** Get the START DATE
''        If (StringContains(pString, "Year")) Then
''            pString = pTextStream.ReadLine  'read next line, this has the date
''            pYear = Mid(pString, 18, 4)
''            pMonth = Mid(pString, 23, 2)
''            pDay = Mid(pString, 27, 2)
''            pStrStartDate = pYear & "/" & pMonth & "/" & pDay
''        End If
''        If (pTextStream.AtEndOfLine = True) Then
''            '*** Get the END DATE
''            pYear = Mid(pString, 18, 4)
''            pMonth = Mid(pString, 23, 2)
''            pDay = Mid(pString, 27, 2)
''            pStrEndDate = pYear & "/" & pMonth & "/" & pDay
''            doContinue = False
''        End If
''    Loop
''    '** cleanup
''    pTextStream.Close
''    Set pTextStream = Nothing
''    Set fso = Nothing
''
''    startDate.value = pStrStartDate
''    endDate.value = pStrEndDate
''    startDate.MinDate = pStrStartDate    'Set date limitation
''    endDate.MinDate = pStrStartDate
''    startDate.MaxDate = pStrEndDate     'Set date limitation
''    endDate.MaxDate = pStrEndDate
''    txtWSTS = 1
''    txtBMPTS = 5
''    cmbOutTS.AddItem "Hourly", 0
''    cmbOutTS.AddItem "Daily", 1
''
''
''CleanUp:
''
''     '** cleanup
''    Set pRow = Nothing
''    Set pCursor = Nothing
''    Set pTable = Nothing
''
''End Sub


Private Sub UpdateSimulationParameters(pInputFilePath As String, pOutputFolderName As String, Optional technique As Integer, Optional numBreaks As Integer, Optional startDate As String, Optional endDate As String)
On Error GoTo ShowError

    'Write the input file folder and output folder to optimization detail
    Dim pOptimizationDetail As iTable
    Set pOptimizationDetail = GetInputDataTable("OptimizationDetail")
    If (pOptimizationDetail Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Sub
    End If

    'Query the table to find existing values, if found, update it, else insert it
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'InputFile'"

    Dim pCursor As ICursor
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)

    Dim iPropName As Long
    iPropName = pCursor.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")

    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    If (pRow Is Nothing) Then
        Set pRow = pOptimizationDetail.CreateRow
    End If
    pRow.value(iPropName) = "InputFile"
    pRow.value(iPropValue) = pInputFilePath
    pRow.Store

    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'OutputFolder'"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    If (pRow Is Nothing) Then
        Set pRow = pOptimizationDetail.CreateRow
    End If
    pRow.value(iPropName) = "OutputFolder"
    pRow.value(iPropValue) = pOutputFolderName
    pRow.Store

    If technique <> 0 Then
        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'Technique'"
        Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pOptimizationDetail.CreateRow
        End If
        pRow.value(iPropName) = "Technique"
        pRow.value(iPropValue) = technique
        pRow.Store
   End If
   If numBreaks <> 0 Then
        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'NumBreak'"
        Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pOptimizationDetail.CreateRow
        End If
        pRow.value(iPropName) = "NumBreak"
        pRow.value(iPropValue) = numBreaks
        pRow.Store
    End If
    
    If startDate <> "" Then
        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'StartDate'"
        Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pOptimizationDetail.CreateRow
        End If
        pRow.value(iPropName) = "StartDate"
        pRow.value(iPropValue) = startDate
        pRow.Store
    End If
    
    If endDate <> "" Then
        pQueryFilter.WhereClause = "ID = 0 AND PropName = 'EndDate'"
        Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pOptimizationDetail.CreateRow
        End If
        pRow.value(iPropName) = "EndDate"
        pRow.value(iPropValue) = endDate
        pRow.Store
    End If
    GoTo CleanUp

ShowError:
    MsgBox "UpdateSimulationParameters: " & Err.description
CleanUp:
    Set pOptimizationDetail = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub


'Commented the following - Sabu Paul: June 14, 2007
''Private Sub optionExternalLanduse_Click()
''    '** controls for external landuse assignment
''    cmbPreDevLanduse.Enabled = True
''    lblPreDevLanduse.Enabled = True
''
''    '** controls for internal landuse assignment
''    lblSWMMLanduseOutflow.Enabled = False
''    lblSWMMPreDevOutflow.Enabled = False
''    txtSWMMLanduseOutflow.Enabled = False
''    txtSWMMPreDevOutflow.Enabled = False
''    cmdSWMMLanduse.Enabled = False
''    cmdSWMMPreDevOutflow.Enabled = False
''
''    '** function to read dates dynamically
''    Call UserSelectedExternalSelection
''End Sub
''
''Private Sub optionInternalLanduse_Click()
''    '** controls for external landuse assignment
''    cmbPreDevLanduse.Enabled = False
''    lblPreDevLanduse.Enabled = False
''
''    '** controls for internal landuse assignment
''    lblSWMMLanduseOutflow.Enabled = True
''    lblSWMMPreDevOutflow.Enabled = True
''    txtSWMMLanduseOutflow.Enabled = True
''    txtSWMMPreDevOutflow.Enabled = True
''    cmdSWMMLanduse.Enabled = True
''    cmdSWMMPreDevOutflow.Enabled = True
''
''    '** function to read dates dynamically
''    Call UserSelectedInternalSelection
''End Sub

