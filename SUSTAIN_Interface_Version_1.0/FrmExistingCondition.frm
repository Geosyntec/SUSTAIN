VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmExistingCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Existing Condition"
   ClientHeight    =   4365
   ClientLeft      =   4920
   ClientTop       =   3600
   ClientWidth     =   6225
   Icon            =   "FrmExistingCondition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   340
      Left            =   4200
      TabIndex        =   6
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCreateInput 
      Caption         =   "Done, create input file"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtBMPID 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   340
      Left            =   5280
      TabIndex        =   2
      Top             =   3840
      Width           =   800
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue defining assessment point"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin TabDlg.SSTab TABParams 
      Height          =   3000
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5292
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Flow"
      TabPicture(0)   =   "FrmExistingCondition.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Pollutant I"
      TabPicture(1)   =   "FrmExistingCondition.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.Label Label1 
      Caption         =   "Select the evaluation factor."
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
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "FrmExistingCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    'Make current tool inactive
    DeactivateCurrentTool
End Sub


''Private Sub cmdCreateInput_Click()
''    'Save all input params
''    If (SaveEvaluationFactorExistingCond = False) Then
''        Exit Sub
''    End If
''
''    'Close the form
''    Unload Me
''
''    'Make current tool inactive
''    DeactivateCurrentTool
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
    If (SaveEvaluationFactorExistingCond = False) Then
        Exit Sub
    End If
    
    'Close the form
    Unload Me
    
    'Make current tool inactive
    DeactivateCurrentTool
    
End Sub

''Private Sub cmdContinue_Click()
''    'save all input params
''    If (SaveEvaluationFactorExistingCond = False) Then
''        Exit Sub
''    End If
''
''    'Close the form
''    Unload Me
''
''End Sub


Private Function SaveEvaluationFactorExistingCond() As Boolean
    SaveEvaluationFactorExistingCond = False
    
  'Collect all checked parameters and enter in a dictionary
    Dim pBMPID As Integer
    pBMPID = CInt(txtBMPID.Text)
    
    'Delete existing parameters
    DeletePreviouslyDefinedOptions pBMPID
    
    'Find total pollutants
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants) + 1
    Dim pPollutantName As String
    
    Dim pBMPSITE As Integer
    Dim pCalcDays As String
    Dim pFactorName As String
    'Retrieve values for pollutants
    Dim iP As Integer
    Dim pControl As Control
   
     'Iterate over all controls and store their values in a dictionary
    Dim pControlDictionary As Scripting.Dictionary
    Set pControlDictionary = CreateObject("Scripting.Dictionary")
    For Each pControl In FrmExistingCondition.Controls
        If (TypeOf pControl Is CheckBox) Then
            pControlDictionary.add pControl.name, CBool(pControl.value)
        ElseIf (TypeOf pControl Is TextBox) Then
            pControlDictionary.add pControl.name, Trim(pControl.Text)
        End If
    Next
    
    'Collection for all input arrays
    Dim pBMPOptimizationColl As Collection
    Set pBMPOptimizationColl = New Collection
    
    'Retrieve values for flow parameters - average annual flow volume
    If (pControlDictionary.Item("flowannual") = True) Then
        pBMPOptimizationColl.add Array(-1, -1, 0, 3, -99, "AAFV")
    End If
    
    'Retrieve values for flow parameters - peak discharge flow
    If (pControlDictionary.Item("flowstormpeak") = True) Then
        pBMPOptimizationColl.add Array(-1, -2, 0, 3, -99, "PDF")
    End If
    
    'Retrieve values for flow parameters - peak discharge flow
    If (pControlDictionary.Item("flowfrequency") = True) Then
        pCalcDays = pControlDictionary.Item("flowCFS")
        If (Not IsNumeric(pCalcDays)) Then
            MsgBox "The exceeding frequency threshold for FLOW should be a valid number."
            TABParams.Tab = 0
            Exit Function
        End If
        If (CDbl(pCalcDays) < 0) Then
            MsgBox "The exceeding frequency threshold for FLOW should be a positive number."
            TABParams.Tab = 0
            Exit Function
        End If
        pBMPOptimizationColl.add Array(-1, -3, CDbl(pCalcDays), 3, -99, "FEF")
    End If
    
    Dim pLoadName As String
    Dim pConcName As String
    Dim pMaxDailyName As String
    Dim pDaysText As String
    
    'Find parameters for each pollutant
    For iP = 1 To pTotalPollutants
        pLoadName = "pollutant" & iP & "Load"
        pConcName = "pollutant" & iP & "Concentration"
        pMaxDailyName = "pollutant" & iP & "MaxDaily"
        pDaysText = "pollutant" & iP & "Days"
        pPollutantName = Replace(Replace(gPollutants(iP - 1), ",", ""), " ", "")
        
        If (pControlDictionary.Item(pLoadName) = True) Then 'Annual Average Load: AAL
            pFactorName = pPollutantName & "_AAL"
            pBMPOptimizationColl.add Array(iP, 1, 0, 3, -99, pFactorName)
        End If
        If (pControlDictionary.Item(pConcName) = True) Then 'Annual Average Concentration: AAC
            pFactorName = pPollutantName & "_AAC"
            pBMPOptimizationColl.add Array(iP, 2, 0, 3, -99, pFactorName)
        End If
        If (pControlDictionary.Item(pMaxDailyName) = True) Then 'Maximum #Days Average Concentration: MAC
            pFactorName = pPollutantName & "_MAC"
            pCalcDays = pControlDictionary.Item(pDaysText)
            If (Not IsNumeric(pCalcDays)) Then
                MsgBox "The average days concentration for " & UCase(gPollutants(iP - 1)) & " should be a valid number."
                TABParams.Tab = iP
                Exit Function
            End If
            If (CDbl(pCalcDays) < 0) Then
                MsgBox "The average days concentration for " & UCase(gPollutants(iP - 1)) & " should be a positive number."
                TABParams.Tab = iP
                Exit Function
            End If
            pBMPOptimizationColl.add Array(iP, 3, CInt(pCalcDays), 3, -99, pFactorName)
        End If
    Next
    
    If (pBMPOptimizationColl.Count > 0) Then
        UpdateOptimizationParamsForBMP pBMPID, pBMPOptimizationColl
    End If
    
    Set pBMPOptimizationColl = Nothing
    
    'return true
    SaveEvaluationFactorExistingCond = True
    
End Function


Private Sub Form_Activate()
    
    'Call the subroutine to read any timeseries file and find total pollutants
    ModuleDecayFact.CreatePollutantList
    
    Dim pValueDictionary As Scripting.Dictionary
     
    'Call subroutine to read values from optimization table
    Set pValueDictionary = ReadOptimizationParametersForExistingCond(txtBMPID.Text)
    
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants) + 1
        
    Dim pFormWidth As Double
'    pFormWidth = 2000 * (1 + pTotalPollutants)
'    If (pFormWidth > 9000) Then
'        pFormWidth = 9000
'    End If

    pFormWidth = 9000
    FrmExistingCondition.Width = pFormWidth
    
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
 
   'Move the OK and Cancel controls
   cmdContinue.Left = pFormWidth - 6000
   cmdCreateInput.Left = pFormWidth - 4000
   cmdDone.Left = pFormWidth - 2500
   cmdCancel.Left = pFormWidth - 1000
    
    Dim iP As Integer
    'Define tab values for each pollutant
    For iP = 1 To pTotalPollutants
        'Define caption and make the tab active
        TABParams.TabCaption(iP) = gPollutants(iP - 1)
        TABParams.Tab = iP
        '*** Add controls on the form for pollutants
        Dim pFrame1
        Set pFrame1 = FrmExistingCondition.Controls.add("VB.Frame", "Frame" & iP)
        With pFrame1
            .Height = 2000
            .Left = 250
            .Top = 700
            .Width = 4200
            .Caption = gPollutants(iP - 1)
            .Visible = True
        End With
        Set pFrame1.Container = TABParams

        Dim pCheckBox1
        Set pCheckBox1 = FrmExistingCondition.Controls.add("VB.Checkbox", "pollutant" & iP & "Load")
        pCheckBox1.Top = 240
        pCheckBox1.Height = 375
        pCheckBox1.Left = 240
        pCheckBox1.Width = 2900
        pCheckBox1.Visible = True
        pCheckBox1.Caption = "Average annual load"
        If (pValueDictionary.Exists("AAL_Pollutant" & iP)) Then
            pCheckBox1.value = pValueDictionary.Item("AAL_Pollutant" & iP)
        End If
        Set pCheckBox1.Container = pFrame1

        Dim pCheckBox2
        Set pCheckBox2 = FrmExistingCondition.Controls.add("VB.Checkbox", "pollutant" & iP & "Concentration")
        pCheckBox2.Top = 840
        pCheckBox2.Height = 375
        pCheckBox2.Left = 240
        pCheckBox2.Width = 2900
        pCheckBox2.Visible = True
        pCheckBox2.Caption = "Average annual concentration"
        If (pValueDictionary.Exists("AAC_Pollutant" & iP)) Then
            pCheckBox2.value = pValueDictionary.Item("AAC_Pollutant" & iP)
        End If
        Set pCheckBox2.Container = pFrame1
        
        Dim pCheckBox3
        Set pCheckBox3 = FrmExistingCondition.Controls.add("VB.Checkbox", "pollutant" & iP & "MaxDaily")
        pCheckBox3.Top = 1440
        pCheckBox3.Height = 300
        pCheckBox3.Left = 240
        pCheckBox3.Width = 975
        pCheckBox3.Visible = True
        pCheckBox3.Caption = "Maximum"
        If (pValueDictionary.Exists("MAC_Pollutant" & iP)) Then
            pCheckBox3.value = pValueDictionary.Item("MAC_Pollutant" & iP)
        End If
        Set pCheckBox3.Container = pFrame1
        
        Dim pTextBox4
        Set pTextBox4 = FrmExistingCondition.Controls.add("VB.TextBox", "pollutant" & iP & "Days")
        pTextBox4.Top = 1440
        pTextBox4.Height = 300
        pTextBox4.Left = 1320
        pTextBox4.Width = 500
        pTextBox4.Visible = True
        pTextBox4.Text = ""
        If (pValueDictionary.Exists("MAC_CalcDays" & iP)) Then
            pTextBox4.Text = pValueDictionary.Item("MAC_CalcDays" & iP)
        End If
        Set pTextBox4.Container = pFrame1
        
        Dim pLabel5
        Set pLabel5 = FrmExistingCondition.Controls.add("VB.Label", "pollutant" & iP & "Label")
        pLabel5.Top = 1440
        pLabel5.Height = 300
        pLabel5.Left = 1920
        pLabel5.Width = 2100
        pLabel5.Visible = True
        pLabel5.Caption = "days average concentration"
        Set pLabel5.Container = pFrame1
    Next
    
    'Make first tab active
    TABParams.Tab = 0
    '*** Add controls on the form for flow
    '*** Add controls on the form
    Set pFrame1 = FrmExistingCondition.Controls.add("VB.Frame", "Frame0")
    With pFrame1
        .Height = 2000
        .Left = 250
        .Top = 700
        .Width = 6200
        .Caption = "Flow"
        .Visible = True
    End With
    Set pFrame1.Container = TABParams
   
    Set pCheckBox1 = FrmExistingCondition.Controls.add("VB.Checkbox", "flowannual")
    pCheckBox1.Top = 240
    pCheckBox1.Height = 375
    pCheckBox1.Left = 240
    pCheckBox1.Width = 2900
    pCheckBox1.Visible = True
    pCheckBox1.Caption = "Average annual flow volume"
    If (pValueDictionary.Exists("AAFV")) Then
        pCheckBox1.value = CInt(pValueDictionary.Item("AAFV"))
    End If
    Set pCheckBox1.Container = pFrame1
    
    Set pCheckBox2 = FrmExistingCondition.Controls.add("VB.Checkbox", "flowstormpeak")
    pCheckBox2.Top = 840
    pCheckBox2.Height = 375
    pCheckBox2.Left = 240
    pCheckBox2.Width = 2900
    pCheckBox2.Visible = True
    pCheckBox2.Caption = "Peak discharge flow"
    If (pValueDictionary.Exists("PDF")) Then
        pCheckBox2.value = CInt(pValueDictionary.Item("PDF"))
    End If
    Set pCheckBox2.Container = pFrame1

    Set pCheckBox3 = FrmExistingCondition.Controls.add("VB.Checkbox", "flowfrequency")
    pCheckBox3.Top = 1440
    pCheckBox3.Height = 300
    pCheckBox3.Left = 240
    pCheckBox3.Width = 3975
    pCheckBox3.Visible = True
    pCheckBox3.Caption = "Exceeding frequency (times per yr). Threshold = "
    If (pValueDictionary.Exists("FEF")) Then
        pCheckBox3.value = CInt(pValueDictionary.Item("FEF"))
    End If
    Set pCheckBox3.Container = pFrame1
    
    Set pTextBox4 = FrmExistingCondition.Controls.add("VB.TextBox", "flowCFS")
    pTextBox4.Top = 1440
    pTextBox4.Height = 300
    pTextBox4.Left = 4320
    pTextBox4.Width = 400
    pTextBox4.Visible = True
    pTextBox4.Text = ""
    If (pValueDictionary.Exists("FEF_CalcDays")) Then
        pTextBox4.Text = CDbl(pValueDictionary.Item("FEF_CalcDays"))
    End If
    Set pTextBox4.Container = pFrame1
    
    Set pLabel5 = FrmExistingCondition.Controls.add("VB.Label", "flowLabel")
    pLabel5.Top = 1440
    pLabel5.Height = 300
    pLabel5.Left = 4920
    pLabel5.Width = 2100
    pLabel5.Visible = True
    pLabel5.Caption = "(cfs)"
    Set pLabel5.Container = pFrame1
        
    'Set controls to nothing, to release memory
    Set pFrame1 = Nothing
    Set pCheckBox1 = Nothing
    Set pCheckBox2 = Nothing
    Set pCheckBox3 = Nothing
    Set pTextBox4 = Nothing
    Set pLabel5 = Nothing
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
