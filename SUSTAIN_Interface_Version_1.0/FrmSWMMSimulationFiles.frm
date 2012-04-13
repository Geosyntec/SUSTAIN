VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSWMMSimulationFiles 
   Caption         =   "Define Simulation Files"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "FrmSWMMSimulationFiles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Predeveloped Landuse Properties"
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   6015
      Begin VB.ComboBox cmbPredevLanduse 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtPreDevInputFile 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdPreDevBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Predeveloped Landuse:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Predeveloped Landuse File:"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSWMMInputFile 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdBrowseSFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "Input File:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmSWMMSimulationFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseSFile_Click()
    CommonDialog.Filter = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtSWMMInputFile.Text = CommonDialog.FileName
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '** get all input values and validate them
    Dim pInputFileName As String
    Dim pPreDevFileName As String
    Dim pInfilConductivity As Double
    Dim pSuctionHead As Double
    Dim pInitialDef As Double
    Dim pPreDevLanduse As String
    
    '** Validate swmm input file
    pInputFileName = txtSWMMInputFile.Text
    If (Trim(pInputFileName) = "") Then
        MsgBox "Please specify SWMM input file to continue."
        Exit Sub
    End If
    
    '** Validate predeveloped landuse file
    pPreDevFileName = txtPreDevInputFile.Text
    If (Trim(pPreDevFileName) = "") Then
        MsgBox "Please specify SWMM predeveloped landuse file to continue."
        Exit Sub
    End If
        
    '** Get predeveloped landuse name
    pPreDevLanduse = cmbPredevLanduse.Text
    
    ' store to the Globals.....
    gPostDevfile = txtSWMMInputFile.Text
    gPreDevfile = txtPreDevInputFile.Text
    
    '** Close the form
    Unload Me
    
    '** write the SWMM output file
    ModuleSWMMFunctions.WriteSWMMProjectDetails pInputFileName
    
    '** write the SWMM output file
    ModuleSWMMFunctions.WriteSWMMPredevelopedLanduseFile pPreDevFileName, pPreDevLanduse, pInfilConductivity, pSuctionHead, pInitialDef
    
End Sub

Private Sub cmdPreDevBrowse_Click()
    CommonDialog.Filter = "LAND Predeveloped Input File (*.inp)|*.inp"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtPreDevInputFile.Text = CommonDialog.FileName
End Sub

Private Sub Form_Load()
On Error GoTo ShowError:
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDLUReclass")
    If (pTable Is Nothing) Then
        MsgBox "LANDLUReclass table not found."
        Exit Sub
    End If
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iTimeSeries As Long
    iTimeSeries = pCursor.FindField("TimeSeries")
    Dim iLuGroup As Long
    iLuGroup = pCursor.FindField("LUGroup")
    Dim iLuGroupID As Long
    iLuGroupID = pCursor.FindField("LUGroupID")
    Dim pLanduseDict As Scripting.Dictionary
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
            pLanduseDict.Add pLuGroup, pLuGroupID
            'Add to the predeveloped landuse combo control
            FrmSWMMSimulationFiles.cmbPredevLanduse.AddItem pLuGroup
        End If
        'End If
        Set pRow = pCursor.NextRow
    Loop
    FrmSWMMSimulationFiles.cmbPredevLanduse.ListIndex = 0
        
    GoTo CleanUp
ShowError:
    MsgBox "Error loading form: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
    Set pLanduseDict = Nothing
    
End Sub
