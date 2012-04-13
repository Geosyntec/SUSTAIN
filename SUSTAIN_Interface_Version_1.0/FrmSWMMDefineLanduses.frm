VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmSWMMDefineLanduses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Landuse Properties"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "FrmSWMMDefineLanduses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save All"
      Height          =   375
      Left            =   7200
      TabIndex        =   66
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   43
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtLanduseID 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabLanduse 
      Height          =   3615
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSWMMDefineLanduses.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtImp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Buildup"
      TabPicture(1)   =   "FrmSWMMDefineLanduses.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBUILDUPUpdate"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmbBUILDUPPollutant"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Washoff"
      TabPicture(2)   =   "FrmSWMMDefineLanduses.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdWASHOFFUpdate"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmbWASHOFFPollutant"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Sediments"
      TabPicture(3)   =   "FrmSWMMDefineLanduses.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   4095
         Begin VB.TextBox txtJg 
            Height          =   285
            Left            =   3000
            TabIndex        =   65
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtKj 
            Height          =   285
            Left            =   3000
            TabIndex        =   63
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtJs 
            Height          =   285
            Left            =   3000
            TabIndex        =   61
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtKs 
            Height          =   285
            Left            =   960
            TabIndex        =   59
            Text            =   "0"
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtCr 
            Height          =   285
            Left            =   960
            TabIndex        =   57
            Text            =   "0"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtCa 
            Height          =   285
            Left            =   960
            TabIndex        =   55
            Text            =   "0"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtJr 
            Height          =   285
            Left            =   960
            TabIndex        =   53
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtKr 
            Height          =   285
            Left            =   960
            TabIndex        =   51
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtP 
            Height          =   285
            Left            =   960
            TabIndex        =   49
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Jg"
            Height          =   255
            Left            =   2160
            TabIndex        =   64
            Top             =   1005
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Kg"
            Height          =   255
            Left            =   2160
            TabIndex        =   62
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Js"
            Height          =   255
            Left            =   2160
            TabIndex        =   60
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Ks"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   2085
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Cr "
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1725
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "Ca (1/day)"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1365
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Jr"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1005
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Kr"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "P"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   280
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdWASHOFFUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   -71640
         TabIndex        =   46
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdBUILDUPUpdate 
         Caption         =   "Update"
         Height          =   315
         Left            =   -71640
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbBUILDUPPollutant 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   480
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   3975
         Begin VB.TextBox txtBMPEff 
            Height          =   375
            Left            =   2040
            TabIndex        =   42
            Text            =   "0.0"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtCoefficient 
            Height          =   330
            Left            =   2040
            TabIndex        =   36
            Text            =   "0.0"
            Top             =   708
            Width           =   1095
         End
         Begin VB.TextBox txtExponent 
            Height          =   330
            Left            =   2040
            TabIndex        =   35
            Text            =   "0.0"
            Top             =   1191
            Width           =   1095
         End
         Begin VB.TextBox txtCleaningEff 
            Height          =   330
            Left            =   2040
            TabIndex        =   34
            Text            =   "0.0"
            Top             =   1674
            Width           =   1095
         End
         Begin VB.ComboBox cmbWASHOFFFunction 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "Coefficient:"
            Height          =   210
            Left            =   240
            TabIndex        =   41
            Top             =   705
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Exponent:"
            Height          =   210
            Left            =   240
            TabIndex        =   40
            Top             =   1185
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Cleaning Efficiency:"
            Height          =   210
            Left            =   240
            TabIndex        =   39
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "BMP Efficiency:"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Function:"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbWASHOFFPollutant 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   480
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   19
         Top             =   840
         Width           =   3975
         Begin VB.ComboBox cmbBUILDUPFunction 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbNormalizer 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtPower 
            Height          =   330
            Left            =   2040
            TabIndex        =   25
            Text            =   "0.0"
            ToolTipText     =   "Time exponent for power buildup or half saturation constant (days) for saturation buildup"
            Top             =   1674
            Width           =   1095
         End
         Begin VB.TextBox txtRateConstant 
            Height          =   330
            Left            =   2040
            TabIndex        =   23
            Text            =   "0.0"
            ToolTipText     =   "lbs per normailizer per day for power buildup or per days for exponential buildup"
            Top             =   1191
            Width           =   1095
         End
         Begin VB.TextBox txtMaxBuildup 
            Height          =   330
            Left            =   2040
            TabIndex        =   21
            Text            =   "0.0"
            ToolTipText     =   "lbs per unit of normalizer"
            Top             =   708
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Function:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Normalizer:"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Power/Sat. Constant:"
            Height          =   210
            Left            =   240
            TabIndex        =   24
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Rate Constant:"
            Height          =   450
            Left            =   240
            TabIndex        =   22
            Top             =   1185
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Max. Buildup:"
            Height          =   450
            Left            =   240
            TabIndex        =   20
            Top             =   705
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Street Sweeping"
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   4095
         Begin VB.TextBox txtLastSwept 
            Height          =   330
            Left            =   1800
            TabIndex        =   17
            Text            =   "0"
            ToolTipText     =   "Number of days since land use was last swept at the start of the simulation"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtAvailability 
            Height          =   330
            Left            =   1800
            TabIndex        =   16
            Text            =   "0"
            ToolTipText     =   "Fraction of pollutant buildup that is available for removal by sweeping"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtInterval 
            Height          =   330
            Left            =   1800
            TabIndex        =   15
            Text            =   "0"
            ToolTipText     =   "Days between street sweeping within the landuse ( 0 for no sweeping )"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Last Swept (days):"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Availability (fraction):"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Interval (days):"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtImp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Optional comments or description for land use."
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   330
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "User assigned name of land use"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "Pollutant:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Pollutant:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "% Imperviousness"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Land use Name"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Landuse Properties"
      Height          =   4215
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.ListBox listLanduses 
      Height          =   3570
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Landuse to View/Edit Properties"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmSWMMDefineLanduses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pLanduseDictionary As Scripting.Dictionary
Private pBuildupDictionary As Scripting.Dictionary
Private pWashoffDictionary As Scripting.Dictionary
Private pSedimentDictionary As Scripting.Dictionary
Private pTempDictionary As Scripting.Dictionary
Private pPollutantCollection As Collection
Private m_pImpCol As Collection

Private Sub cmbBUILDUPPollutant_Click()

   '** save the values just entered
    Dim pPollutantName As String
    pPollutantName = Trim(cmbBUILDUPPollutant.Text)
    If (pPollutantName = "") Then
        Exit Sub
    End If
    
    '** LOAD Default values
    cmbBUILDUPFunction.ListIndex = 0
    cmbNormalizer.ListIndex = 0
    Dim strVal As String
    
    '** get values for new pollutant, reset all input controls
    Dim pObjectDictionary As Scripting.Dictionary
    If (pBuildupDictionary.Exists(pPollutantName)) Then
        Set pObjectDictionary = pBuildupDictionary.Item(pPollutantName)
        strVal = pObjectDictionary.Item(pPollutantName & "-" & "BUILDUPFUNCTION")
        If Trim(strVal) = "" Then
            cmbBUILDUPFunction.ListIndex = 0
        Else
            cmbBUILDUPFunction.Text = pObjectDictionary.Item(pPollutantName & "-" & "BUILDUPFUNCTION")
        End If
        txtMaxBuildup.Text = pObjectDictionary.Item(pPollutantName & "-" & "MAXBUILDUP")
        txtRateConstant.Text = pObjectDictionary.Item(pPollutantName & "-" & "RATECONSTANT")
        txtPower.Text = pObjectDictionary.Item(pPollutantName & "-" & "POWER")
        strVal = pObjectDictionary.Item(pPollutantName & "-" & "NORMALIZER")
        If Trim(strVal) = "" Then
            cmbNormalizer.ListIndex = 0
        Else
            cmbNormalizer.Text = pObjectDictionary.Item(pPollutantName & "-" & "NORMALIZER")
        End If
    Else
        cmbBUILDUPFunction.ListIndex = 0
        txtMaxBuildup.Text = "0.0"
        txtRateConstant.Text = "0.0"
        txtPower.Text = "0.0"
        cmbNormalizer.ListIndex = 0
    End If
    
End Sub


Private Sub cmdBUILDUPUpdate_Click()
On Error GoTo ShowError

    If (pBuildupDictionary Is Nothing) Then
        GoTo CleanUp
    End If

    '** save the values just entered
    Dim pPollutantName As String
    pPollutantName = Trim(cmbBUILDUPPollutant.Text)
    If (pPollutantName = "") Then
        GoTo CleanUp
    End If
    
    '** create an object dictionary
    Dim pObjectDictionary As Scripting.Dictionary
    Set pObjectDictionary = CreateObject("Scripting.Dictionary")    'pBuildupDictionary.Item(pPollutantName)

    '** save the values just entered
    pObjectDictionary.Item(pPollutantName & "-" & "BUILDUPFUNCTION") = cmbBUILDUPFunction.Text
    pObjectDictionary.Item(pPollutantName & "-" & "MAXBUILDUP") = txtMaxBuildup.Text
    pObjectDictionary.Item(pPollutantName & "-" & "RATECONSTANT") = txtRateConstant.Text
    pObjectDictionary.Item(pPollutantName & "-" & "POWER") = txtPower.Text
    pObjectDictionary.Item(pPollutantName & "-" & "NORMALIZER") = cmbNormalizer.Text
    
    '** update the values back in buildup dictionary
    Set pBuildupDictionary.Item(pPollutantName) = pObjectDictionary

    '** cleanup
    Set pObjectDictionary = Nothing
    GoTo CleanUp
ShowError:
    MsgBox "error: " & Err.description
CleanUp:

End Sub

Private Sub cmbWASHOFFPollutant_Click()

   '** save the values just entered
    Dim pPollutantName As String
    pPollutantName = Trim(cmbWASHOFFPollutant.Text)
    If (pPollutantName = "") Then
        Exit Sub
    End If
  
      '** LOAD Default values
    cmbWASHOFFFunction.ListIndex = 0
    Dim strVal As String
    
    '** get values for new pollutant, reset all input controls
    Dim pObjectDictionary As Scripting.Dictionary
    If (pWashoffDictionary.Exists(pPollutantName)) Then
        Set pObjectDictionary = pWashoffDictionary.Item(pPollutantName)
        strVal = pObjectDictionary.Item(pPollutantName & "-" & "WASHOFFFUNCTION")
        If Trim(strVal) = "" Then
            cmbWASHOFFFunction.ListIndex = 0
        Else
            cmbWASHOFFFunction.Text = pObjectDictionary.Item(pPollutantName & "-" & "WASHOFFFUNCTION")
        End If
        txtCoefficient.Text = pObjectDictionary.Item(pPollutantName & "-" & "COEFFICIENT")
        txtExponent.Text = pObjectDictionary.Item(pPollutantName & "-" & "EXPONENT")
        txtCleaningEff.Text = pObjectDictionary.Item(pPollutantName & "-" & "CLEANINGEFF")
        txtBMPEff.Text = pObjectDictionary.Item(pPollutantName & "-" & "BMPEFF")
    Else
        cmbWASHOFFFunction.ListIndex = 0
        txtCoefficient.Text = "0.0"
        txtExponent.Text = "0.0"
        txtCleaningEff.Text = "0.0"
        txtBMPEff.Text = "0.0"
    End If
        
     '** cleanup
    Set pObjectDictionary = Nothing
End Sub

''Private Sub cmdAdd_Click()
''    SSTabLanduse.Enabled = True
''    SSTabLanduse.Tab = 0
''
''    '** clear the txtLanduseID value
''    txtLanduseID.Text = ""
''
''    '** load pollutant names from SWMMPollutant table
''    Set pPollutantCollection = ModuleSWMMFunctions.LoadPollutantNames
''
''    '** load pollutant names for buildup
''    Call LoadPollutantNamesForBuildup
''
''    '** load pollutant names for washoff
''    Call LoadPollutantNamesForWashoff
''
''    '** pre-populate the pollutant dictionary
''    Call PrepareDictionaryForUpdatingBuildupAndWashoff
''
''End Sub

Private Sub cmdCancel_Click()
    '** cleanup
    Set pLanduseDictionary = Nothing
    Set pBuildupDictionary = Nothing
    Set pWashoffDictionary = Nothing
    Set pPollutantCollection = Nothing
    
    '** close the form
    Unload Me
End Sub

''Private Sub cmdDelete_Click()
''    '** Confirm the deletion
''    Dim boolDelete
''    boolDelete = MsgBox("Are you sure you want to delete this landuse information ?", vbYesNo)
''    If (boolDelete = vbNo) Then
''        Exit Sub
''    End If
''
''    '** get landuse id
''    Dim pLanduseID As Integer
''    pLanduseID = listLanduses.ListIndex + 1
''
''    '** get the table to delete records
''    Dim pSWMMLanduseTable As iTable
''    Set pSWMMLanduseTable = GetInputDataTable("LANDLanduses")
''
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''    pQueryFilter.WhereClause = "ID = " & pLanduseID
''
''    '** delete records
''    pSWMMLanduseTable.DeleteSearchedRows pQueryFilter
''
''    '*** Increment the id's by 1 number for all records after deleted id
''    Dim pFromID As Integer
''    pFromID = pLanduseID + 1
''    Dim bContinue As Boolean
''    bContinue = True
''    Dim pCursor As ICursor
''    Dim pRow As iRow
''    Dim iID As Long
''    iID = pSWMMLanduseTable.FindField("ID")
''
''    Do While bContinue
''        pQueryFilter.WhereClause = "ID = " & pFromID
''        Set pCursor = Nothing
''        Set pRow = Nothing
''        Set pCursor = pSWMMLanduseTable.Search(pQueryFilter, False)
''        Set pRow = pCursor.NextRow
''        If (pRow Is Nothing) Then
''            bContinue = False
''        End If
''        Do While Not pRow Is Nothing
''            pRow.value(iID) = pFromID - 1
''            pRow.Store
''            Set pRow = pCursor.NextRow
''        Loop
''        pFromID = pFromID + 1
''    Loop
''
''    '** clean up
''    Set pQueryFilter = Nothing
''    Set pSWMMLanduseTable = Nothing
''
''    '** clear all control values
''    ClearLanduseRelatedControls
''
''    '** load landuse names
''    LoadLanduseNamesForLanduseForm
''End Sub

Private Sub cmdEdit_Click()

    If (listLanduses.ListIndex > -1) Then
    
        '** update the value from listbox
        txtName.Text = listLanduses.List(listLanduses.ListIndex)
        
        ' ** Update the Imperviousness....
        txtImp.Text = Format(m_pImpCol.Item(listLanduses.ListIndex + 1) * 100, "00.00")
        
        '*** Update the hidden landuse id value
        txtLanduseID.Text = listLanduses.ListIndex + 1
    
        cmdSave.Enabled = True
        SSTabLanduse.Enabled = True
        SSTabLanduse.Tab = 0
        
        '** load pollutant names from SWMMPollutant table
        Set pPollutantCollection = ModuleSWMMFunctions.LoadPollutantNames
    
        '** load pollutant names for buildup
        Call LoadPollutantNamesForBuildup
    
        '** load pollutant names for washoff
        Call LoadPollutantNamesForWashoff
            
        '** pre-populate the pollutant dictionary
        Call ReadLandusePropertyValuesFromTable(listLanduses.ListIndex + 1)
        
        '** populate values from the dictionary to the form
        If pLanduseDictionary.Count > 0 Then
            ''txtName.Text = pLanduseDictionary.Item("Name")
            txtInterval.Text = Trim(pLanduseDictionary.Item("Interval"))
            txtAvailability.Text = Trim(pLanduseDictionary.Item("Availibility"))
            txtLastSwept.Text = Trim(pLanduseDictionary.Item("LastSwept"))
            txtP.Text = Trim(pLanduseDictionary.Item("P"))
            txtKr.Text = Trim(pLanduseDictionary.Item("Kr"))
            txtJr.Text = Trim(pLanduseDictionary.Item("Jr"))
            txtCa.Text = Trim(pLanduseDictionary.Item("Ca"))
            txtCr.Text = Trim(pLanduseDictionary.Item("Cr"))
            txtKs.Text = Trim(pLanduseDictionary.Item("Ks"))
            txtJs.Text = Trim(pLanduseDictionary.Item("Js"))
            txtKj.Text = Trim(pLanduseDictionary.Item("Kg"))
            txtJg.Text = Trim(pLanduseDictionary.Item("Jg"))
        Else
            txtInterval.Text = "0"
            txtAvailability.Text = "0"
            txtLastSwept.Text = "0"
        End If
    End If

End Sub


Private Sub cmdSave_Click()

    '** input validation for all pollutant - buildup and washoff values

    '** save the landuse name, description and other properties
    pLanduseDictionary.RemoveAll
    pLanduseDictionary.add "Name", txtName.Text
    pLanduseDictionary.add "Interval", txtInterval.Text
    pLanduseDictionary.add "Availibility", txtAvailability.Text
    pLanduseDictionary.add "LastSwept", txtLastSwept.Text
    pLanduseDictionary.add "Imperviousness", txtImp.Text
    '** save the Sediment values......
    pLanduseDictionary.add "P", txtP.Text
    pLanduseDictionary.add "Kr", txtKr.Text
    pLanduseDictionary.add "Jr", txtJr.Text
    pLanduseDictionary.add "Ca", txtCa.Text
    pLanduseDictionary.add "Cr", txtCr.Text
    pLanduseDictionary.add "Ks", txtKs.Text
    pLanduseDictionary.add "Js", txtJs.Text
    pLanduseDictionary.add "Kg", txtKj.Text
    pLanduseDictionary.add "Jg", txtJg.Text
    
    
    '** get the landuse ID
    Dim pLanduseID As String
    If (Trim(txtLanduseID.Text) = "") Then
        pLanduseID = listLanduses.ListCount + 1
    Else
        pLanduseID = txtLanduseID.Text
    End If
    
    '** call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDLanduses", pLanduseID, pLanduseDictionary
     
    '*** define a dictionary object
    Dim pObjectDictionary As Scripting.Dictionary
    Set pObjectDictionary = CreateObject("Scripting.Dictionary")
    
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim pPollutantName As String
    Dim iCount As Integer
    
    '** buildup values for all pollutants
    For iCount = 1 To pCount
        pPollutantName = pPollutantCollection.Item(iCount)

        If pBuildupDictionary.Exists(pPollutantName) Then
          '** get the builtup dictionary
          Set pObjectDictionary = pBuildupDictionary.Item(pPollutantName)
    
          '** call the module to create table and add rows for these values
          ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDLanduses", pLanduseID, pObjectDictionary
        End If
    Next

    '** washoff values for all pollutants
    For iCount = 1 To pCount
        pPollutantName = pPollutantCollection.Item(iCount)
        
        If pWashoffDictionary.Exists(pPollutantName) Then
            '** get the builtup dictionary
            Set pObjectDictionary = pWashoffDictionary.Item(pPollutantName)
    
            '** call the module to create table and add rows for these values
            ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDLanduses", pLanduseID, pObjectDictionary
        End If
    Next
    
    '** clean up memory
    Set pObjectDictionary = Nothing
    cmdSave.Enabled = False

    '** clear all control values
    Call ClearLanduseRelatedControls
    
    '** refresh the landuse names
    Call LoadLanduseNamesForLanduseForm
End Sub


Private Sub cmdSaveAll_Click()
    
    Dim iCnt As Integer
    For iCnt = 0 To listLanduses.ListCount - 1
        listLanduses.Selected(iCnt) = True
        Call cmdEdit_Click
        Call cmdSave_Click
    Next iCnt
    
    
End Sub
Private Sub cmdWASHOFFUpdate_Click()
On Error GoTo ShowError

    If (pWashoffDictionary Is Nothing) Then
        GoTo CleanUp
    End If

    '** save the values just entered
    Dim pPollutantName As String
    pPollutantName = Trim(cmbWASHOFFPollutant.Text)
    If (pPollutantName = "") Then
        GoTo CleanUp
    End If
    
    '** create an object dictionary
    Dim pObjectDictionary As Scripting.Dictionary
    Set pObjectDictionary = CreateObject("Scripting.Dictionary") ' pWashoffDictionary.Item(pPollutantName)

    '** save the values just entered
    pObjectDictionary.Item(pPollutantName & "-" & "WASHOFFFUNCTION") = cmbWASHOFFFunction.Text
    pObjectDictionary.Item(pPollutantName & "-" & "COEFFICIENT") = txtCoefficient.Text
    pObjectDictionary.Item(pPollutantName & "-" & "EXPONENT") = txtExponent.Text
    pObjectDictionary.Item(pPollutantName & "-" & "CLEANINGEFF") = txtCleaningEff.Text
    pObjectDictionary.Item(pPollutantName & "-" & "BMPEFF") = txtBMPEff.Text
    
    '** update the values back in buildup dictionary
    Set pWashoffDictionary.Item(pPollutantName) = pObjectDictionary

    '** cleanup
    Set pObjectDictionary = Nothing
    GoTo CleanUp
ShowError:
    MsgBox "error: " & Err.description
CleanUp:

End Sub

Private Sub Form_Load()
On Error GoTo ShowError
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '** load buildup function values
    cmbBUILDUPFunction.AddItem "NONE"
    cmbBUILDUPFunction.AddItem "POW"
    cmbBUILDUPFunction.AddItem "EXP"
    cmbBUILDUPFunction.AddItem "SAT"
    
    '** load normalizer
    cmbNormalizer.AddItem "AREA"
    cmbNormalizer.AddItem "CURB"
    
    '** load washoff function values
    cmbWASHOFFFunction.AddItem "NONE"
    cmbWASHOFFFunction.AddItem "EXP"
    cmbWASHOFFFunction.AddItem "RC"
    cmbWASHOFFFunction.AddItem "EMC"
    
    ' ** Laod teh sediment defaults......
    txtP.Text = "0.17"
    txtKr.Text = "0.294"
    txtJr.Text = "1.81"
    txtCa.Text = "0.1"
    txtCr.Text = "0.78"
    txtKs.Text = "0.5"
    txtJs.Text = "1.67"
    txtKj.Text = "0.0"
    txtJg.Text = "2.0"
    

    '** load landuse names for editing/viewing purpose
    Call LoadLanduseNamesForLanduseForm
     
    ' ** Load Imperviousnes.....
    Set m_pImpCol = LoadImperviousForLanduseForm
   
    Exit Sub
ShowError:
    MsgBox "Error Initializing FrmSWMMDefineLanduses: " & Err.description

End Sub

Public Sub LoadPollutantNamesForWashoff()

    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    
    '** clear the combo box
    cmbWASHOFFPollutant.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        cmbWASHOFFPollutant.AddItem pPollutantCollection.Item(iCount)
    Next
    
End Sub




Public Sub LoadPollutantNamesForBuildup()

    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    
    '** clear the combo box
    cmbBUILDUPPollutant.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        cmbBUILDUPPollutant.AddItem pPollutantCollection.Item(iCount)
    Next
      
End Sub


Private Sub SSTabLanduse_GotFocus()
    If (SSTabLanduse.Tab > 0 And txtName.Text = "") Then
        MsgBox "Enter landuse name to continue."
        SSTabLanduse.Tab = 0
        Exit Sub
    End If
       
End Sub


Public Sub LoadLanduseNamesForLanduseForm()

    '** Refresh landuse names
    Dim pLanduseCollection As Collection
    Set pLanduseCollection = LoadLanduseReclassifiedCategories
    
    If pLanduseCollection Is Nothing Then
        Exit Sub
    End If
    
    listLanduses.Clear
    Dim pCount As Integer
    pCount = pLanduseCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        listLanduses.AddItem pLanduseCollection.Item(iCount)
    Next
    
    Set pLanduseCollection = Nothing
End Sub

Public Function LoadImperviousForLanduseForm() As Collection

    Dim pTable As iTable
    Set pTable = GetInputDataTable("LUReclass")
    
    If (pTable Is Nothing) Then
        MsgBox "LUReclass (Landuse reclassification) table not found."
        Exit Function
    End If
    
    Dim iGroupName As Long
    iGroupName = pTable.FindField("LUGroup")
    Dim iGroupCode As Long
    iGroupCode = pTable.FindField("Percentage")
    Dim iImpervious As Long
    iImpervious = pTable.FindField("Impervious")
        
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pDictionary As Scripting.Dictionary
    Set pDictionary = CreateObject("Scripting.Dictionary")
    
    Dim pCollection As Collection
    Set pCollection = New Collection
    
    Dim groupName As String
    Dim impValue As Double
    
    Do While Not pRow Is Nothing
        If CInt(pRow.value(iImpervious)) = 1 Then
            groupName = pRow.value(iGroupName) & "_imp"
            impValue = 1#   'pRow.value(iGroupCode)
        Else
            groupName = pRow.value(iGroupName) & "_perv"
            impValue = 0#   ' 1 - pRow.value(iGroupCode)
        End If
'        If Not (pDictionary.Exists(pRow.value(iGroupName) & pRow.value(iGroupCode))) Then
'            pDictionary.add pRow.value(iGroupName) & pRow.value(iGroupCode), pRow.value(iGroupCode)
'            pCollection.add pRow.value(iGroupCode)
'        End If
        If Not pDictionary.Exists(groupName) Then
            pDictionary.add groupName, impValue
            pCollection.add impValue
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    '** return the collection back
    Set LoadImperviousForLanduseForm = pCollection
    
    '** cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pDictionary = Nothing
    Set pTable = Nothing
    
End Function


Private Sub PrepareDictionaryForUpdatingBuildupAndWashoff()
On Error GoTo ShowError

    '** If a name is entered, create the dictionary object
    '** and store values
    Set pLanduseDictionary = CreateObject("Scripting.Dictionary")
    Set pBuildupDictionary = CreateObject("Scripting.Dictionary")
    Set pWashoffDictionary = CreateObject("Scripting.Dictionary")
  
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim pPollutantName As String
    Dim iCount As Integer
    
    '** define a temporary dictionary object
    Dim pObjectDictionary As Scripting.Dictionary
    
    '** buildup values for all pollutants
    For iCount = 1 To pCount
        pPollutantName = pPollutantCollection.Item(iCount)
                
        '** add variables in the dictionary for buildup
        Set pObjectDictionary = CreateObject("Scripting.Dictionary")
        pObjectDictionary.add pPollutantName & "-" & "BUILDUPFUNCTION", "NONE"
        pObjectDictionary.add pPollutantName & "-" & "MAXBUILDUP", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "RATECONSTANT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "POWER", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "NORMALIZER", "AREA"
        
        '** add object in the buildup dictionary
        pBuildupDictionary.add pPollutantName, pObjectDictionary
        
        '** cleanup
        Set pObjectDictionary = Nothing
        
        '** add variables in the dictionary for washoff
        Set pObjectDictionary = CreateObject("Scripting.Dictionary")
        pObjectDictionary.add pPollutantName & "-" & "WASHOFFFUNCTION", "NONE"
        pObjectDictionary.add pPollutantName & "-" & "COEFFICIENT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "EXPONENT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "CLEANINGEFF", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "BMPEFF", "0.0"
        
        '** add object in the washoff dictionary
        pWashoffDictionary.add pPollutantName, pObjectDictionary
        
        '** cleanup
        Set pObjectDictionary = Nothing
    Next

    GoTo CleanUp
    
ShowError:
    MsgBox "PrepareDictionaryForUpdatingBuildupAndWashoff: " & Err.description
CleanUp:

End Sub


'** This subroutine reads all landuse values from the swmmlanduses table
'** and will be used to populate the forms
Private Sub ReadLandusePropertyValuesFromTable(pLanduseID As Integer)
On Error GoTo ShowError
   
    '** initialize the landuse dictionary
    Set pLanduseDictionary = CreateObject("Scripting.Dictionary")
    Set pBuildupDictionary = CreateObject("Scripting.Dictionary")
    Set pWashoffDictionary = CreateObject("Scripting.Dictionary")
    
    ' ** Initialize the Buildup & Washoff Dict.....
    Call Init_BuildupWashoff_Dict
    
    '** find the table, exit if not found
    Dim pLanduseTable As iTable
    Set pLanduseTable = GetInputDataTable("LANDLanduses")
    If (pLanduseTable Is Nothing) Then
        Exit Sub
    End If
    
    '** define index to access fields
    Dim iPropName As Long
    iPropName = pLanduseTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pLanduseTable.FindField("PropValue")
    
    '** define the query and cursor to iterate
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pLanduseID
    
    Dim pCursor As ICursor
    Set pCursor = pLanduseTable.Search(pQueryFilter, True)
    
    '** define the row to access values
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    '** define variables to store values
    Dim pPropName As String
    Dim pPropValue As String
    Do While Not pRow Is Nothing
        pPropName = pRow.value(iPropName)
        pPropValue = pRow.value(iPropValue)
        pLanduseDictionary.add pPropName, pPropValue
        '** go to next row
        Set pRow = pCursor.NextRow
    Loop
    
    ' ** Flag....
    If pLanduseDictionary.Count = 0 Then Exit Sub
           
    '** initialize both dictionaries with values
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim pPollutantName As String
    Dim iCount As Integer
    Dim pBuilupPropertyValue As String
    For iCount = 1 To pCount
        '** get the pollutant name
        pPollutantName = Trim(pPollutantCollection.Item(iCount))
        
        '** create a dictionary for all buildup values, and add it
        '** add variables in the dictionary for buildup
        Set pTempDictionary = CreateObject("Scripting.Dictionary")
        
        '** add values for builup for a pollutant
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "BUILDUPFUNCTION"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "MAXBUILDUP"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "RATECONSTANT"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "POWER"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "NORMALIZER"
 
        '** add object in the buildup dictionary
        Set pBuildupDictionary.Item(pPollutantName) = pTempDictionary
        
        '** reinitialize the object dictionary
        Set pTempDictionary = Nothing
        Set pTempDictionary = CreateObject("Scripting.Dictionary")

        '** add values for washoff for a pollutant
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "WASHOFFFUNCTION"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "COEFFICIENT"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "EXPONENT"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "CLEANINGEFF"
        UpdateBuildUpWashoffDictionary pPollutantName & "-" & "BMPEFF"

        '** add object in the washoff dictionary
        Set pWashoffDictionary.Item(pPollutantName) = pTempDictionary
                        
    Next
           
    GoTo CleanUp
ShowError:
    MsgBox "ReadLandusePropertyValuesFromTable: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pLanduseTable = Nothing
    Set pTempDictionary = Nothing
End Sub



Private Sub UpdateBuildUpWashoffDictionary(pPropertyValue As String)

    '** add it to the object dictionary
    pTempDictionary.add pPropertyValue, pLanduseDictionary.Item(pPropertyValue)
    
    '** remove the value from landuse dictionary
    pLanduseDictionary.Remove (pPropertyValue)

End Sub

Private Sub Init_BuildupWashoff_Dict()
        
    '** create an object dictionary
    Dim pObjectDictionary As Scripting.Dictionary
    
    '** buildup values for all pollutants
    For iCount = 1 To pPollutantCollection.Count
        pPollutantName = pPollutantCollection.Item(iCount)
        
        Set pObjectDictionary = CreateObject("Scripting.Dictionary")    'pBuildupDictionary.Item(pPollutantName)
        '** add values for builup for a pollutant
        pObjectDictionary.add pPollutantName & "-" & "BUILDUPFUNCTION", "NONE"
        pObjectDictionary.add pPollutantName & "-" & "MAXBUILDUP", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "RATECONSTANT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "POWER", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "NORMALIZER", "AREA"
        '** update the values back in buildup dictionary
        pBuildupDictionary.add pPollutantName, pObjectDictionary
    
        Set pObjectDictionary = CreateObject("Scripting.Dictionary")    'pBuildupDictionary.Item(pPollutantName)
        '** add values for washoff for a pollutant
        pObjectDictionary.add pPollutantName & "-" & "WASHOFFFUNCTION", "NONE"
        pObjectDictionary.add pPollutantName & "-" & "COEFFICIENT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "EXPONENT", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "CLEANINGEFF", "0.0"
        pObjectDictionary.add pPollutantName & "-" & "BMPEFF", "0.0"
        '** update the values back in buildup dictionary
        pWashoffDictionary.add pPollutantName, pObjectDictionary
        
    Next


End Sub

Private Sub ClearLanduseRelatedControls()

    '** clear landuse controls
    txtName.Text = ""
    txtImp.Text = ""
    txtInterval.Text = "0"
    txtAvailability.Text = "0"
    txtLastSwept.Text = "0"
    
    '** clear buildup controls
    cmbBUILDUPPollutant.ListIndex = 0
    cmbBUILDUPFunction.ListIndex = 0
    txtMaxBuildup.Text = "0.0"
    txtRateConstant.Text = "0.0"
    txtPower.Text = "0.0"
    cmbNormalizer.ListIndex = 0
    
    '** clear washoff controls
    cmbBUILDUPFunction.ListIndex = 0
    txtCoefficient.Text = "0.0"
    txtExponent.Text = "0.0"
    txtCleaningEff.Text = "0.0"
    txtBMPEff.Text = "0.0"
    
        
    ' ** Laod teh sediment defaults......
    txtP.Text = "0.17"
    txtKr.Text = "0.294"
    txtJr.Text = "1.81"
    txtCa.Text = "0.1"
    txtCr.Text = "0.78"
    txtKr.Text = "0.5"
    txtJr.Text = "1.67"
    txtKj.Text = "0.0"
    txtJg.Text = "2.0"
    
End Sub


Private Function LoadLanduseReclassifiedCategories() As Collection

    Dim pTable As iTable
    Set pTable = GetInputDataTable("LUReclass")
    
    If (pTable Is Nothing) Then
        MsgBox "LUReclass (Landuse reclassification) table not found."
        Exit Function
    End If
    
    Dim iGroupName As Long
    iGroupName = pTable.FindField("LUGroup")
    
    Dim iImpervious As Long
    iImpervious = pTable.FindField("Impervious")
    
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pDictionary As Scripting.Dictionary
    Set pDictionary = CreateObject("Scripting.Dictionary")
    
    Dim pCollection As Collection
    Set pCollection = New Collection
    
    Dim groupName As String
    
    Do While Not pRow Is Nothing
        If CInt(pRow.value(iImpervious)) = 1 Then
            groupName = pRow.value(iGroupName) & "_imp"
        Else
            groupName = pRow.value(iGroupName) & "_perv"
        End If
        'If Not (pDictionary.Exists(pRow.value(iGroupName))) Then
        If Not (pDictionary.Exists(groupName)) Then
'            pDictionary.add pRow.value(iGroupName), pRow.value(iGroupName)
'            pCollection.add pRow.value(iGroupName)
            pDictionary.add groupName, groupName
            pCollection.add groupName
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    '** return the collection back
    Set LoadLanduseReclassifiedCategories = pCollection
    
    '** cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pDictionary = Nothing
    Set pTable = Nothing
    
End Function
