VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMaximizeBenefit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   4395
   ClientTop       =   2055
   ClientWidth     =   6150
   Icon            =   "FrmMaximizeBenefit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue defining assessment point"
      Height          =   400
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCreateInput 
      Caption         =   "Done, create input file"
      Height          =   400
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   400
      Left            =   3960
      TabIndex        =   6
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   5040
      TabIndex        =   5
      Top             =   3840
      Width           =   800
   End
   Begin VB.TextBox txtBMPID 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab TABParams 
      Height          =   3000
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5292
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Flow"
      TabPicture(0)   =   "FrmMaximizeBenefit.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Pollutant"
      TabPicture(1)   =   "FrmMaximizeBenefit.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
   End
   Begin VB.Label Label01 
      Caption         =   "    1 = least important"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "* 10 = most important"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Select evaluation factor and input priority factor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "FrmMaximizeBenefit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pValueDictionary As Scripting.Dictionary

Private Sub cmdCancel_Click()
    Unload Me
    'Make current tool inactive
    DeactivateCurrentTool
End Sub

''Private Sub cmdCreateInput_Click()
''    'Save all input params
''    If (SaveEvaluationFactorsForMaximizeBenefit = False) Then
''        Exit Sub
''    End If
''
''    'Close the form
''    Unload Me
''
''    'Make current tool inactive
''    DeactivateCurrentTool
''
''    'Show the optimize cost button
''    FrmOptimizeBenefit.Show vbModal
''
''    'Call the subroutine to check Watershed/SubWatershed layer to continue.
''    Dim boolWatershed As Boolean
''    boolWatershed = FindAndConvertWatershedFeatureLayerToRaster()
''
''    Dim boolInputFile As Boolean
''    If (boolWatershed = True) Then
''        boolInputFile = ModuleFile.WriteInputTextFile
''    End If
''
''    If (boolInputFile = True) Then
''        'Show the simulation dialog box
''        FrmRunDLL.Show vbModal
''    End If
''End Sub

Private Sub cmdDone_Click()
    'Save all input params
    If (SaveEvaluationFactorsForMaximizeBenefit = False) Then
        Exit Sub
    End If
    
    'Close the form
    Unload Me
    
    'Make current tool inactive
    DeactivateCurrentTool
   
End Sub

''Private Sub cmdContinue_Click()
''    'save all input params
''    If (SaveEvaluationFactorsForMaximizeBenefit = False) Then
''        Exit Sub
''    End If
''
''    'Close the form
''    Unload Me
''
''End Sub

Private Function SaveEvaluationFactorsForMaximizeBenefit() As Boolean

    SaveEvaluationFactorsForMaximizeBenefit = False
    
    'Find total pollutants
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants) + 1
    
    'Collect all checked parameters and enter in a dictionary
    Dim pBMPID As Integer
    pBMPID = CInt(txtBMPID.Text)
    
    'delete previously defined values
    DeletePreviouslyDefinedOptions pBMPID
    
    Dim pCalcDays As Integer
    Dim pStrCalcDays As String
    Dim pStrTargetValue As String
    Dim pTargetValue As Double
    Dim pFactorName As String
    'Retrieve values for pollutants
    Dim iP As Integer
    Dim pControl As Control
    
     'Iterate over all controls and store their values in a dictionary
    Dim pControlDictionary As Scripting.Dictionary
    Set pControlDictionary = CreateObject("Scripting.Dictionary")
    For Each pControl In FrmMaximizeBenefit.Controls
        If (TypeOf pControl Is CheckBox) Then
            pControlDictionary.add pControl.name, CBool(pControl.value)
        ElseIf (TypeOf pControl Is TextBox) Then
            pControlDictionary.add pControl.name, Trim(pControl.Text)
        End If
    Next
    
    'Collection for all input arrays
    Dim pBMPOptimizationColl As Collection
    Set pBMPOptimizationColl = New Collection
    
    TABParams.Tab = 0
    'Retrieve values for flow parameters - average annual flow volume
    If (pControlDictionary.Item("flowannual") = True) Then
        pStrTargetValue = pControlDictionary.Item("flowannualPrty")
        '*** Input Validation
        If (Not IsNumeric(pStrTargetValue)) Then
            MsgBox "Priority Factor for Average Annual Flow Volume must be a valid number "
            Exit Function
        End If
        pTargetValue = CDbl(pStrTargetValue)
        If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
            MsgBox "Priority Factor for Average Annual Flow Volume must be positive number between 1-10"
            Exit Function
        End If
        '*** Input Validation
        pBMPOptimizationColl.add Array(-1, -1, 0, 1, pTargetValue, "AAFV")
    End If
    
    'Retrieve values for flow parameters - peak discharge flow
    If (pControlDictionary.Item("flowstormpeak") = True) Then
        pStrTargetValue = pControlDictionary.Item("flowstormpeakPrty")
        '*** Input Validation
        If (Not IsNumeric(pStrTargetValue)) Then
            MsgBox "Priority Factor for Peak Discharge Flow must be a valid number "
            Exit Function
        End If
        pTargetValue = CDbl(pStrTargetValue)
        If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
            MsgBox "Priority Factor for Peak Discharge Flow must be positive number between 1-10"
            Exit Function
        End If
        '*** Input Validation
        pBMPOptimizationColl.add Array(-1, -2, 0, 1, pTargetValue, "PDF")
    End If
    
    'Retrieve values for flow parameters - peak discharge flow
    If (pControlDictionary.Item("flowfrequency") = True) Then
        '*** Input Validation
        pStrTargetValue = pControlDictionary.Item("flowfrequencyPrty")
        If (Not IsNumeric(pStrTargetValue)) Then
            MsgBox "Priority Factor for Exceeding frequency must be a valid number "
            Exit Function
        End If
        pTargetValue = CDbl(pStrTargetValue)
        If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
            MsgBox "Priority Factor for Exceeding frequency must be positive number between 1-10"
            Exit Function
        End If
        pStrCalcDays = pControlDictionary.Item("flowCFS")
        If (Not IsNumeric(pStrCalcDays)) Then
            MsgBox "Threshold value for Exceeding frequency must be a valid number "
            Exit Function
        End If
        pCalcDays = CDbl(pStrCalcDays)
        If (pCalcDays < 0) Then
            MsgBox "Threshold value for Exceeding frequency must be valid positive number."
            Exit Function
        End If
        '*** Input Validation
        pBMPOptimizationColl.add Array(-1, -3, pCalcDays, 1, pTargetValue, "FEF")
    End If
    
    Dim pLoadName As String
    Dim pConcName As String
    Dim pMaxDailyName As String
    Dim pDaysText As String
    Dim pPollutantName As String
    
    'Find parameters for each pollutant
    For iP = 1 To pTotalPollutants
        pLoadName = "Pollutant" & iP & "Load"
        pConcName = "Pollutant" & iP & "Concentration"
        pMaxDailyName = "Pollutant" & iP & "MaxDaily"
        pDaysText = "Pollutant" & iP & "Days"
        TABParams.Tab = iP
        pPollutantName = Replace(Replace(gPollutants(iP - 1), ",", ""), " ", "")
        If (pControlDictionary.Item(pLoadName) = True) Then 'Annual Average Load: AAL
            pFactorName = pPollutantName & "_AAL"
            '*** Input Validation
            pStrTargetValue = pControlDictionary.Item(pLoadName & "Prty")
            If (Not IsNumeric(pStrTargetValue)) Then
                MsgBox "Priority Factor for Average Annual Load (" & gPollutants(iP - 1) & ") must be a valid number "
                Exit Function
            End If
            pTargetValue = CDbl(pStrTargetValue)
            If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
                MsgBox "Priority Factor for Average Annual Load (" & gPollutants(iP - 1) & ") must be positive number between 1-10"
                Exit Function
            End If
            '*** Input Validation
            pBMPOptimizationColl.add Array(iP, 1, 0, 1, pTargetValue, pFactorName)
        ElseIf (pControlDictionary.Item(pConcName) = True) Then 'Annual Average Concentration: AAC
            pFactorName = pPollutantName & "_AAC"
            pStrTargetValue = pControlDictionary.Item(pConcName & "Prty")
            '*** Input Validation
            If (Not IsNumeric(pStrTargetValue)) Then
                MsgBox "Priority Factor for Average Annual Concentration (" & gPollutants(iP - 1) & ") must be a valid number "
                Exit Function
            End If
            pTargetValue = CDbl(pStrTargetValue)
            If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
                MsgBox "Priority Factor for Average Annual Concentration (" & gPollutants(iP - 1) & ") must be positive number between 1-10"
                Exit Function
            End If
            '*** Input Validation
            pBMPOptimizationColl.add Array(iP, 2, 0, 1, pTargetValue, pFactorName)
        ElseIf (pControlDictionary.Item(pMaxDailyName) = True) Then 'Maximum #Days Average Concentration: MAC
            pFactorName = pPollutantName & "_MAC"
            pStrTargetValue = pControlDictionary.Item(pMaxDailyName & "Prty")
            pStrCalcDays = pControlDictionary.Item(pDaysText)
             '*** Input Validation
            If (Not IsNumeric(pStrTargetValue)) Then
                MsgBox "Priority Factor for Maximum Days Average Concentration must be a valid number "
                Exit Function
            End If
            pTargetValue = CDbl(pStrTargetValue)
            If (CDbl(pTargetValue) < 1 Or CDbl(pTargetValue) > 10) Then
                MsgBox "Priority Factor for Maximum Days Average Concentration must be positive number between 1-10"
                Exit Function
            End If
            If (Not IsNumeric(pStrCalcDays)) Then
                MsgBox "Maximum Days Average Concentration value must be a valid number "
                Exit Function
            End If
            pCalcDays = CInt(pStrCalcDays)
            If (pCalcDays < 0) Then
                MsgBox "Maximum Days Average Concentration value must be valid positive number."
                Exit Function
            End If
            '*** Input Validation
            pBMPOptimizationColl.add Array(iP, 3, pCalcDays, 1, pTargetValue, pFactorName)
        End If
    Next
    
    If (pBMPOptimizationColl.Count > 0) Then
        UpdateOptimizationParamsForBMP pBMPID, pBMPOptimizationColl
    End If
    
    Set pBMPOptimizationColl = Nothing
    
    'return true value
    SaveEvaluationFactorsForMaximizeBenefit = True
    
End Function


Private Sub Form_Activate()
        
    'Call the subroutine to read any timeseries file and find total pollutants
    ModuleDecayFact.CreatePollutantList
    
    'Get values from table
    Set pValueDictionary = ReadOptimizationParametersForMaximizeBenefit(txtBMPID.Text)
    
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants)
        
    Dim pFormWidth As Double
    'pFormWidth = 2000 * (1 + pTotalPollutants)
    pFormWidth = 9000
'    If (pFormWidth > 9000) Then
'        pFormWidth = 9000
'    End If
    FrmMaximizeBenefit.Width = pFormWidth
    
    'Find total tabs
    Dim pTotalTabs As Integer
    pTotalTabs = 1 + pTotalPollutants
    'Define total tabs
    TABParams.Tabs = pTotalTabs   'Flow + pollutants
    TABParams.Width = pFormWidth - 1000
    If (pTotalTabs > 6) Then
        TABParams.TabsPerRow = 6
    Else
        TABParams.TabsPerRow = pTotalTabs
    End If
    
    Dim iP As Integer
    'Define tab values for each pollutant
    For iP = 1 To pTotalPollutants
        'Define caption and make the tab active
        TABParams.TabCaption(iP) = gPollutants(iP - 1)
        TABParams.Tab = iP
        '*** Add controls on the form for pollutants
        Dim pFrame1
        Set pFrame1 = FrmMaximizeBenefit.Controls.add("VB.Frame", "Frame" & iP)
        With pFrame1
            .Height = 2400
            .Left = 250
            .Top = 500
            .Width = 6200
            .Caption = gPollutants(iP - 1)
            .Visible = True
        End With
        Set pFrame1.Container = TABParams

        Dim pCheckBox1
        Set pCheckBox1 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "Pollutant" & iP & "Load")
        pCheckBox1.Top = 440
        pCheckBox1.Height = 375
        pCheckBox1.Left = 240
        pCheckBox1.Width = 2900
        pCheckBox1.Visible = True
        pCheckBox1.Caption = "Average annual load"
        If (pValueDictionary.Exists("Pollutant" & iP & "Load")) Then
            pCheckBox1.value = pValueDictionary.Item("Pollutant" & iP & "Load")
        End If
        Set pCheckBox1.Container = pFrame1

        Dim pCheckBox2
        Set pCheckBox2 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "Pollutant" & iP & "Concentration")
        pCheckBox2.Top = 1040
        pCheckBox2.Height = 375
        pCheckBox2.Left = 240
        pCheckBox2.Width = 2900
        pCheckBox2.Visible = True
        pCheckBox2.Caption = "Average annual concentration"
        If (pValueDictionary.Exists("Pollutant" & iP & "Concentration")) Then
            pCheckBox2.value = pValueDictionary.Item("Pollutant" & iP & "Concentration")
        End If
        Set pCheckBox2.Container = pFrame1
        
        Dim pCheckBox3
        Set pCheckBox3 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "Pollutant" & iP & "MaxDaily")
        pCheckBox3.Top = 1640
        pCheckBox3.Height = 375
        pCheckBox3.Left = 240
        pCheckBox3.Width = 975
        pCheckBox3.Visible = True
        pCheckBox3.Caption = "Maximum"
        If (pValueDictionary.Exists("Pollutant" & iP & "MaxDaily")) Then
            pCheckBox3.value = pValueDictionary.Item("Pollutant" & iP & "MaxDaily")
        End If
        Set pCheckBox3.Container = pFrame1
        
        Dim pTextBox4
        Set pTextBox4 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "Pollutant" & iP & "Days")
        pTextBox4.Top = 1640
        pTextBox4.Height = 375
        pTextBox4.Left = 1320
        pTextBox4.Width = 500
        pTextBox4.Visible = True
        pTextBox4.Text = ""
        If (pValueDictionary.Exists("Pollutant" & iP & "Days")) Then
            pTextBox4.Text = pValueDictionary.Item("Pollutant" & iP & "Days")
        End If
        Set pTextBox4.Container = pFrame1
        
        Dim pLabel5
        Set pLabel5 = FrmMaximizeBenefit.Controls.add("VB.Label", "Pollutant" & iP & "Label")
        pLabel5.Top = 1640
        pLabel5.Height = 375
        pLabel5.Left = 1920
        pLabel5.Width = 2100
        pLabel5.Visible = True
        pLabel5.Caption = "days average concentration"
        Set pLabel5.Container = pFrame1
        
        Dim pTextBox6
        Set pTextBox6 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "Pollutant" & iP & "LoadPrty")
        pTextBox6.Top = 440
        pTextBox6.Height = 375
        pTextBox6.Left = 4700
        pTextBox6.Width = 500
        pTextBox6.Visible = True
        pTextBox6.Text = ""
        If (pValueDictionary.Exists("Pollutant" & iP & "LoadPrty")) Then
            pTextBox6.Text = pValueDictionary.Item("Pollutant" & iP & "LoadPrty")
        End If
        Set pTextBox6.Container = pFrame1
        
        Dim pTextBox7
        Set pTextBox7 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "Pollutant" & iP & "ConcentrationPrty")
        pTextBox7.Top = 1040
        pTextBox7.Height = 375
        pTextBox7.Left = 4700
        pTextBox7.Width = 500
        pTextBox7.Visible = True
        pTextBox7.Text = ""
        If (pValueDictionary.Exists("Pollutant" & iP & "ConcentrationPrty")) Then
            pTextBox7.Text = pValueDictionary.Item("Pollutant" & iP & "ConcentrationPrty")
        End If
        Set pTextBox7.Container = pFrame1
        
        Dim pTextBox8
        Set pTextBox8 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "Pollutant" & iP & "MaxDailyPrty")
        pTextBox8.Top = 1640
        pTextBox8.Height = 375
        pTextBox8.Left = 4700
        pTextBox8.Width = 500
        pTextBox8.Visible = True
        pTextBox8.Text = ""
        If (pValueDictionary.Exists("Pollutant" & iP & "MaxDailyPrty")) Then
            pTextBox8.Text = pValueDictionary.Item("Pollutant" & iP & "MaxDailyPrty")
        End If
        Set pTextBox8.Container = pFrame1
        
        Dim pLabel9
        Set pLabel9 = FrmMaximizeBenefit.Controls.add("VB.Label", "Pollutant" & iP & "PFLabel")
        pLabel9.Top = 240
        pLabel9.Height = 375
        pLabel9.Left = 4500
        pLabel9.Width = 1500
        pLabel9.Visible = True
        pLabel9.Caption = "Priority Factor(1-10) *"
        Set pLabel9.Container = pFrame1
    Next
    
    'Make first tab active
    TABParams.Tab = 0
    '*** Add controls on the form for flow
    '*** Add controls on the form
    Set pFrame1 = FrmMaximizeBenefit.Controls.add("VB.Frame", "Frame0")
    With pFrame1
        .Height = 2400
        .Left = 250
        .Top = 500
        .Width = 6200
        .Caption = "Flow"
        .Visible = True
    End With
    Set pFrame1.Container = TABParams
   
    Set pCheckBox1 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "flowannual")
    pCheckBox1.Top = 440
    pCheckBox1.Height = 375
    pCheckBox1.Left = 240
    pCheckBox1.Width = 2900
    pCheckBox1.Visible = True
    pCheckBox1.Caption = "Average annual flow volume"
    If (pValueDictionary.Exists("flowannual")) Then
        pCheckBox1.value = pValueDictionary.Item("flowannual")
    End If
    Set pCheckBox1.Container = pFrame1
    
    Set pCheckBox2 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "flowstormpeak")
    pCheckBox2.Top = 1040
    pCheckBox2.Height = 375
    pCheckBox2.Left = 240
    pCheckBox2.Width = 2900
    pCheckBox2.Visible = True
    pCheckBox2.Caption = "2-yr storm peak discharge flow"
    If (pValueDictionary.Exists("flowstormpeak")) Then
        pCheckBox2.value = pValueDictionary.Item("flowstormpeak")
    End If
    Set pCheckBox2.Container = pFrame1

    Set pCheckBox3 = FrmMaximizeBenefit.Controls.add("VB.Checkbox", "flowfrequency")
    pCheckBox3.Top = 1640
    pCheckBox3.Height = 375
    pCheckBox3.Left = 240
    pCheckBox3.Width = 3700
    pCheckBox3.Visible = True
    pCheckBox3.Caption = "Exceeding frequency (times per yr). Threshold ="
    If (pValueDictionary.Exists("flowfrequency")) Then
        pCheckBox3.value = pValueDictionary.Item("flowfrequency")
    End If
    Set pCheckBox3.Container = pFrame1
    
    Set pTextBox4 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "flowCFS")
    pTextBox4.Top = 1640
    pTextBox4.Height = 375
    pTextBox4.Left = 4000
    pTextBox4.Width = 500
    pTextBox4.Visible = True
    pTextBox4.Text = ""
    If (pValueDictionary.Exists("flowCFS")) Then
        pTextBox4.Text = pValueDictionary.Item("flowCFS")
    End If
    Set pTextBox4.Container = pFrame1
    
    Set pLabel5 = FrmMaximizeBenefit.Controls.add("VB.Label", "flowLabel")
    pLabel5.Top = 1640
    pLabel5.Height = 375
    pLabel5.Left = 4600
    pLabel5.Width = 2100
    pLabel5.Visible = True
    pLabel5.Caption = "(cfs)"
    Set pLabel5.Container = pFrame1
   
    Set pTextBox6 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "flowannualPrty")
    pTextBox6.Top = 440
    pTextBox6.Height = 375
    pTextBox6.Left = 5100
    pTextBox6.Width = 500
    pTextBox6.Visible = True
    pTextBox6.Text = ""
    If (pValueDictionary.Exists("flowannualPrty")) Then
        pTextBox6.Text = pValueDictionary.Item("flowannualPrty")
    End If
    Set pTextBox6.Container = pFrame1
    
    Set pTextBox7 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "flowstormpeakPrty")
    pTextBox7.Top = 1040
    pTextBox7.Height = 375
    pTextBox7.Left = 5100
    pTextBox7.Width = 500
    pTextBox7.Visible = True
    pTextBox7.Text = ""
    If (pValueDictionary.Exists("flowstormpeakPrty")) Then
        pTextBox7.Text = pValueDictionary.Item("flowstormpeakPrty")
    End If
    Set pTextBox7.Container = pFrame1
    
    Set pTextBox8 = FrmMaximizeBenefit.Controls.add("VB.TextBox", "flowfrequencyPrty")
    pTextBox8.Top = 1640
    pTextBox8.Height = 375
    pTextBox8.Left = 5100
    pTextBox8.Width = 500
    pTextBox8.Visible = True
    pTextBox8.Text = ""
    If (pValueDictionary.Exists("flowfrequencyPrty")) Then
        pTextBox8.Text = pValueDictionary.Item("flowfrequencyPrty")
    End If
    Set pTextBox8.Container = pFrame1
    
    Set pLabel9 = FrmMaximizeBenefit.Controls.add("VB.Label", "flowPFLabel")
    pLabel9.Top = 240
    pLabel9.Height = 375
    pLabel9.Left = 4500
    pLabel9.Width = 1500
    pLabel9.Visible = True
    pLabel9.Caption = "Priority Factor(1-10)"
    Set pLabel9.Container = pFrame1
        
    'Set controls to nothing, to release memory
    Set pFrame1 = Nothing
    Set pCheckBox1 = Nothing
    Set pCheckBox2 = Nothing
    Set pCheckBox3 = Nothing
    Set pTextBox4 = Nothing
    Set pLabel5 = Nothing
    Set pTextBox6 = Nothing
    Set pTextBox7 = Nothing
    Set pTextBox8 = Nothing
    Set pLabel9 = Nothing
    Set pValueDictionary = Nothing
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
