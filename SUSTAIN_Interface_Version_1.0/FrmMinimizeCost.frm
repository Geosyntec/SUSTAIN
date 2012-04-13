VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMinimizeCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minimize Cost"
   ClientHeight    =   8745
   ClientLeft      =   3660
   ClientTop       =   1515
   ClientWidth     =   8520
   Icon            =   "FrmMinimizeCost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue defining assessment point"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   7440
      TabIndex        =   5
      Top             =   8160
      Width           =   800
   End
   Begin VB.CommandButton cmdCreateInput 
      Caption         =   "Done, create input file"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   400
      Left            =   6240
      TabIndex        =   3
      Top             =   8160
      Width           =   975
   End
   Begin VB.TextBox txtBMPID 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "BMPID"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab TABParams 
      Height          =   7275
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Flow"
      TabPicture(0)   =   "FrmMinimizeCost.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Pollutant I"
      TabPicture(1)   =   "FrmMinimizeCost.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.Label Label1 
      Caption         =   "Select the evaluation factor and input control target."
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
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmMinimizeCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pTotalPollutants As Integer
Private pValueDictionary As Scripting.Dictionary

Private Sub cmdCancel_Click()
    Unload Me
    'Make current tool inactive
    DeactivateCurrentTool
End Sub


''Private Sub cmdCreateInput_Click()
''    'Save all input params
''    If (SaveEvaluationFactorsMinimizeCost = False) Then
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
''    FrmOptimizeCost.Show vbModal
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
    If (SaveEvaluationFactorsMinimizeCost = False) Then
        Exit Sub
    End If
    
    'Close the form
    Unload Me
    
    'Make current tool inactive
    DeactivateCurrentTool
        
    
End Sub

''Private Sub cmdContinue_Click()
''    'save all input params
''    If (SaveEvaluationFactorsMinimizeCost = False) Then
''        Exit Sub
''    End If
''
''    'Close the form
''    Unload Me
''
''End Sub


Private Function SaveEvaluationFactorsMinimizeCost() As Boolean
    
    SaveEvaluationFactorsMinimizeCost = False
    
    'Collect all checked parameters and enter in a dictionary
    Dim pBMPID As Integer
    pBMPID = CInt(txtBMPID.Text)

    'Delete existing parameters
    DeletePreviouslyDefinedOptions pBMPID
    
    Dim pControl As Control
     'Iterate over all controls and store their values in a dictionary
    Dim pOptionButtonDictionary As Scripting.Dictionary
    Set pOptionButtonDictionary = CreateObject("Scripting.Dictionary")
    Dim pTextBoxDictionary As Scripting.Dictionary
    Set pTextBoxDictionary = CreateObject("Scripting.Dictionary")
    For Each pControl In FrmMinimizeCost.Controls
        If (TypeOf pControl Is OptionButton) Then
            pOptionButtonDictionary.add pControl.name, CBool(pControl.value)
        ElseIf (TypeOf pControl Is TextBox) Then
            pTextBoxDictionary.add pControl.name, Trim(pControl.Text)
        End If
    Next
    
    Dim pOBControlName As String
    Dim pOBControlValue As Boolean
    
    'Collection for all input arrays
    Dim pBMPOptimizationColl As Collection
    Set pBMPOptimizationColl = New Collection
    
    Dim pTargetValue As Double
    Dim pStrTargetValue As String
    Dim pStrCalcDays As String
    Dim pCalcDays As Double
    Dim pCalcMode As Integer
    Dim pTextBoxName As String
    Dim iT As Integer

    Dim pP As Integer
    Dim pParameterName As String
    Dim pParameterNameArray
    pParameterNameArray = Array("_Flow_AAFV", "_Flow_PDF", "_Flow_FEF")
    Dim pParameterSuffixArray
    pParameterSuffixArray = Array("", "_%", "_S", "")
    Dim pFactorName As String
    Dim pFactorType As Integer
    Dim pErrorMessage As String
    'For each parameter(AAFV, PDF, FEF), get the the three options
    TABParams.Tab = 0
    For pP = 0 To 2
        pFactorType = -1 * (pP + 1)
        pParameterName = pParameterNameArray(pP)
        For iT = 1 To 3
            'Select case iT
            Select Case iT
                Case 1:
                    pErrorMessage = "Percent of value under existing condition for "
                Case 2:
                    pErrorMessage = "Between pre-development and existing condition for "
                Case 3:
                    pErrorMessage = "Specified value for "
            End Select
            
            'Select case pP for appropriate message
            Select Case pP
                Case 0:
                    pErrorMessage = pErrorMessage & "Average Annual Flow Volume "
                Case 1:
                    pErrorMessage = pErrorMessage & "Peak Discharge Flow "
                Case 2:
                    pErrorMessage = pErrorMessage & "Exceeding frequency "
            End Select

            'Read values from parameter dictionary
            pFactorName = Trim(Replace(pParameterName, "_Flow_", "")) & pParameterSuffixArray(iT)
            pTextBoxName = "TextBox" & iT & pParameterName
            pCalcMode = iT
            pOBControlName = Replace(pTextBoxName, "TextBox", "Option")
            pOBControlValue = pOptionButtonDictionary.Item(pOBControlName)

            If (pOBControlValue = True) Then
                pStrTargetValue = Trim(pTextBoxDictionary.Item(pTextBoxName))
                If (Not IsNumeric(pStrTargetValue)) Then
                    pErrorMessage = pErrorMessage & "must be valid number."
                    MsgBox pErrorMessage
                    Exit Function
                End If
                Select Case iT
                    Case 1:
                        If (CDbl(pStrTargetValue) < 0 Or CDbl(pStrTargetValue) > 100) Then
                            pErrorMessage = pErrorMessage & "must be within (0-100) range."
                            MsgBox pErrorMessage
                            Exit Function
                        End If
                    Case 2:
                        If (CDbl(pStrTargetValue) < 0 Or CDbl(pStrTargetValue) > 1) Then
                            pErrorMessage = pErrorMessage & "must be within (0-1) range."
                            MsgBox pErrorMessage
                            Exit Function
                        End If
                    Case 3:
                        If (CDbl(pStrTargetValue) < 0) Then
                            pErrorMessage = pErrorMessage & "must be positive number."
                            MsgBox pErrorMessage
                            Exit Function
                        End If
                End Select
                
                '*** IF its Exceeding frequency/ maximum days, check if textbox is empty
                pCalcDays = 0
                If (pFactorType = -3) Then
                    pStrCalcDays = Trim(pTextBoxDictionary.Item("CalcDays" & "_Flow_FEF"))
                    If (Not IsNumeric(pStrCalcDays)) Then
                        MsgBox "Threshold value must be a valid number."
                        Exit Function
                    End If
                    pCalcDays = CDbl(pStrCalcDays)
                    If (pCalcDays < 0) Then
                        MsgBox "Threshold value must be a positive number."
                        Exit Function
                    End If
                End If

                'No errors found, save it to collection
                pTargetValue = CDbl(pStrTargetValue)
                pBMPOptimizationColl.add Array(-1, pFactorType, pCalcDays, pCalcMode, pTargetValue, pFactorName)

            End If
        Next
    Next
    
   
    Dim iP As Integer   'counter for pollutants
    pParameterNameArray = Array("_AAL", "_AAC", "_MAC")
    
    Dim pPollutantName As String
    For iP = 1 To pTotalPollutants
        'Get the pollutant name
        pPollutantName = Replace(Replace(gPollutants(iP - 1), ",", ""), " ", "")
        'Make the tab active
        TABParams.Tab = iP
        For pP = 0 To 2
            pFactorType = 1 * (pP + 1)
            pParameterName = pParameterNameArray(pP)
            For iT = 1 To 3
                'Select case iT
                Select Case iT
                    Case 1:
                        pErrorMessage = "Percent of value under existing condition for "
                    Case 2:
                        pErrorMessage = "Between pre-development and existing condition for "
                    Case 3:
                        pErrorMessage = "Specified value for "
                End Select
                'Select case pP for appropriate message
                Select Case pP
                    Case 0:
                        pErrorMessage = pErrorMessage & "Average Annual Load "
                    Case 1:
                        pErrorMessage = pErrorMessage & "Average Annual Concentration "
                    Case 2:
                        pErrorMessage = pErrorMessage & "Maximum Average Concentration "
                End Select
                pErrorMessage = pErrorMessage & "(" & gPollutants(iP - 1) & ") "
                pFactorName = pPollutantName & Trim(Replace(pParameterName, "_", "")) & pParameterSuffixArray(iT)
                pTextBoxName = "TextBox" & iT & "_Pollutant" & iP & pParameterName
                pCalcMode = iT
                                
                '*** Error checking - Input Validation
                pOBControlName = Replace(pTextBoxName, "TextBox", "Option")
                pOBControlValue = pOptionButtonDictionary.Item(pOBControlName)
                If (pOBControlValue = True) Then
                    pStrTargetValue = Trim(pTextBoxDictionary.Item(pTextBoxName))
                    If (Not IsNumeric(pStrTargetValue)) Then
                        pErrorMessage = pErrorMessage & "must be valid number."
                        MsgBox pErrorMessage
                        Exit Function
                    End If
                    Select Case iT
                        Case 1:
                            If (CDbl(pStrTargetValue) < 0 Or CDbl(pStrTargetValue) > 100) Then
                                pErrorMessage = pErrorMessage & "must be within (0-100) range."
                                MsgBox pErrorMessage
                                Exit Function
                            End If
                        Case 2:
                            If (CDbl(pStrTargetValue) < 0 Or CDbl(pStrTargetValue) > 1) Then
                                pErrorMessage = pErrorMessage & "must be within (0-1) range."
                                MsgBox pErrorMessage
                                Exit Function
                            End If
                        Case 3:
                            If (CDbl(pStrTargetValue) < 0) Then
                                pErrorMessage = pErrorMessage & "must be positive number."
                                MsgBox pErrorMessage
                                Exit Function
                            End If
                    End Select
                    
                    '*** Check for maximum days concentration value
                    pCalcDays = 0
                    If (pFactorType = 3) Then
                        pStrCalcDays = Trim(pTextBoxDictionary.Item("CalcDays" & "_Pollutant" & iP & "_MAC"))
                        If (Not IsNumeric(pStrCalcDays)) Then
                            MsgBox "Maximum days average concentration value must be a valid number."
                            Exit Function
                        End If
                        pCalcDays = CDbl(pStrCalcDays)
                        If (pCalcDays < 0) Then
                            MsgBox "Maximum days average concentration value must be a positive number."
                            Exit Function
                        End If
                    End If
                
                    'Update value in table
                    pTargetValue = CDbl(pStrTargetValue)
                    pBMPOptimizationColl.add Array(iP, pFactorType, pCalcDays, pCalcMode, pTargetValue, pFactorName)
                End If
            Next
        Next
    Next
    
    '*** Update the collection values in the optimization detail table
    If (pBMPOptimizationColl.Count > 0) Then
        UpdateOptimizationParamsForBMP pBMPID, pBMPOptimizationColl
    End If
    Set pBMPOptimizationColl = Nothing
    Set pOptionButtonDictionary = Nothing
    Set pTextBoxDictionary = Nothing
    
    SaveEvaluationFactorsMinimizeCost = True
End Function

Private Sub Form_Activate()
      
    Dim pFrameText As String
    Dim pOption1Text As String
    Dim pOption2Text As String
    Dim pOption3Text As String
    Dim pLabel3Text As String
    pOption1Text = "Percent of the value under existing condition (0-100)"
    pOption2Text = "Between pre-development and existing condition (0-1)"
    pOption3Text = "Specified value"
   
    'Read values from table
    Set pValueDictionary = ReadOptimizationParametersForMinimizeCost(txtBMPID.Text)
    
    Dim iP As Integer
    For iP = 1 To pTotalPollutants
        'Make the tab active
         TABParams.Tab = iP
         TABParams.TabCaption(iP) = gPollutants(iP - 1)
        
        'Call subroutine to generate top label
        GenerateCommonLabels "Pollutant" & iP
    
        'Call subroutine to generate controls and affix them to FLOW tab
         pFrameText = "Annual Average Load"
         pLabel3Text = " kg/yr"
         GenerateControlsAndContainerFrame 700, 240, gPollutants(iP - 1), "_Pollutant" & iP & "_AAL", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, ""
         
        'Call subroutine to generate controls and affix them to FLOW tab
         pFrameText = "Annual Average Concentration"
         pLabel3Text = " mg/L"
         GenerateControlsAndContainerFrame 2800, 240, gPollutants(iP - 1), "_Pollutant" & iP & "_AAC", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, ""
         
        'Call subroutine to generate controls and affix them to FLOW tab
         pFrameText = "Maximum days average concentration"
         pLabel3Text = " mg/L"
         GenerateControlsAndContainerFrame 4900, 240, gPollutants(iP - 1), "_Pollutant" & iP & "_MAC", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, "Maximum days: "
    Next
   
    TABParams.Tab = 0
    
    'Call subroutine to generate top label
    GenerateCommonLabels "FlowLabel"
    
    'Call subroutine to generate controls and affix them to FLOW tab
    pFrameText = "Annual Average Flow Volume"
    pLabel3Text = " ft3/yr"
    GenerateControlsAndContainerFrame 700, 240, "Flow", "_Flow_AAFV", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, ""
    
    'Call subroutine to generate controls and affix them to FLOW tab
    pFrameText = "Peak Discharge Flow"
    pLabel3Text = " cfs"
    GenerateControlsAndContainerFrame 2800, 240, "Flow", "_Flow_PDF", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, ""
    
    'Call subroutine to generate controls and affix them to FLOW tab
    pFrameText = "Exceeding frequency (times per yr)"
    pLabel3Text = "times per yr"
    GenerateControlsAndContainerFrame 4900, 240, "Flow", "_Flow_FEF", pFrameText, pOption1Text, pOption2Text, pOption3Text, pLabel3Text, "Threshold(cfs)"
  
    Set pValueDictionary = Nothing
   
End Sub

Private Sub GenerateCommonLabels(pLabelName As String)
        Dim pLabel1
        Set pLabel1 = FrmMinimizeCost.Controls.add("VB.Label", pLabelName)
        pLabel1.Top = 500
        pLabel1.Height = 375
        pLabel1.Left = 5000
        pLabel1.Width = 1500
        pLabel1.Visible = True
        pLabel1.Caption = "Control Target"
        pLabel1.Font.Bold = True
'        pLabel1.Font.Color = vbRed
        Set pLabel1.Container = TABParams
        Set pLabel1 = Nothing
End Sub


Private Sub GenerateControlsAndContainerFrame(fTop As Long, fLeft As Long, _
                                              frameName, paramName As String, _
                                              frameText As String, option1Text As String, _
                                              option2Text As String, option3Text As String, _
                                              label3Text As String, Optional pLabelCaption As String)
                                              
    Dim pAdditionalHt As Integer
    pAdditionalHt = 0
    If (pLabelCaption <> "") Then
        pAdditionalHt = 500
    End If
                                              
    '*** Add controls on the form for pollutants
     Dim pFrame1
     Set pFrame1 = FrmMinimizeCost.Controls.add("VB.Frame", "Frame" & paramName & paramOption)
     With pFrame1
         .Height = 1700 + pAdditionalHt
         .Left = fLeft
         .Top = fTop
         .Width = 6200
         .Caption = frameText
         .Visible = True
     End With
     Set pFrame1.Container = TABParams
     
     If (pLabelCaption <> "") Then
        Dim pLabel0
        Set pLabel0 = FrmMinimizeCost.Controls.add("VB.Label", "Threshold" & paramName)
        pLabel0.Top = 200
        pLabel0.Height = 300
        pLabel0.Left = fLeft + 3000
        pLabel0.Width = 1400
        pLabel0.Visible = True
        pLabel0.Caption = pLabelCaption
        Set pLabel0.Container = pFrame1
        Set pLabel0 = Nothing
        
        Dim pTextBox0
        Set pTextBox0 = FrmMinimizeCost.Controls.add("VB.TextBox", "CalcDays" & paramName)
        pTextBox0.Top = 200
        pTextBox0.Height = 300
        pTextBox0.Left = fLeft + 4500
        pTextBox0.Width = 500
        pTextBox0.Visible = True
        pTextBox0.Text = ""
        If (pValueDictionary.Exists("CalcDays" & paramName)) Then
           pTextBox0.Text = pValueDictionary.Item("CalcDays" & paramName)
        End If
        Set pTextBox0.Container = pFrame1
    End If
     
     Dim pRadioBtn1
     Set pRadioBtn1 = FrmMinimizeCost.Controls.add("VB.OptionButton", "Option1" & paramName)
     pRadioBtn1.Top = 200 + pAdditionalHt
     pRadioBtn1.Height = 375
     pRadioBtn1.Left = fLeft + 500
     pRadioBtn1.Width = 4500
     pRadioBtn1.Visible = True
     pRadioBtn1.Caption = option1Text
     If (pValueDictionary.Exists("Option1" & paramName)) Then
        pRadioBtn1.value = pValueDictionary.Item("Option1" & paramName)
     End If
     Set pRadioBtn1.Container = pFrame1
 
     Dim pTextBox1
     Set pTextBox1 = FrmMinimizeCost.Controls.add("VB.TextBox", "TextBox1" & paramName)
     pTextBox1.Top = 200 + pAdditionalHt
     pTextBox1.Height = 375
     pTextBox1.Left = fLeft + 4800
     pTextBox1.Width = 500
     pTextBox1.Visible = True
     pTextBox1.Text = ""
     If (pValueDictionary.Exists("TextBox1" & paramName)) Then
        pTextBox1.Text = pValueDictionary.Item("TextBox1" & paramName)
     End If
     Set pTextBox1.Container = pFrame1
     
     Dim pRadioBtn2
     Set pRadioBtn2 = FrmMinimizeCost.Controls.add("VB.OptionButton", "Option2" & paramName)
     pRadioBtn2.Top = 700 + pAdditionalHt
     pRadioBtn2.Height = 375
     pRadioBtn2.Left = fLeft + 500
     pRadioBtn2.Width = 4500
     pRadioBtn2.Visible = True
     pRadioBtn2.Caption = option2Text
     If (pValueDictionary.Exists("Option2" & paramName)) Then
        pRadioBtn2.value = pValueDictionary.Item("Option2" & paramName)
     End If
     Set pRadioBtn2.Container = pFrame1
     
     Dim pTextBox2
     Set pTextBox2 = FrmMinimizeCost.Controls.add("VB.TextBox", "TextBox2" & paramName)
     pTextBox2.Top = 700 + pAdditionalHt
     pTextBox2.Height = 375
     pTextBox2.Left = fLeft + 4800
     pTextBox2.Width = 500
     pTextBox2.Visible = True
     pTextBox2.Text = ""
     If (pValueDictionary.Exists("TextBox2" & paramName)) Then
        pTextBox2.Text = pValueDictionary.Item("TextBox2" & paramName)
     End If
     Set pTextBox2.Container = pFrame1
     
     Dim pRadioBtn3
     Set pRadioBtn3 = FrmMinimizeCost.Controls.add("VB.OptionButton", "Option3" & paramName)
     pRadioBtn3.Top = 1200 + pAdditionalHt
     pRadioBtn3.Height = 375
     pRadioBtn3.Left = fLeft + 500
     pRadioBtn3.Width = 4500
     pRadioBtn3.Visible = True
     pRadioBtn3.Caption = option3Text & "     (" & label3Text & ")"
     If (pValueDictionary.Exists("Option3" & paramName)) Then
        pRadioBtn3.value = pValueDictionary.Item("Option3" & paramName)
     End If
     Set pRadioBtn3.Container = pFrame1
     
     Dim pTextBox3
     Set pTextBox3 = FrmMinimizeCost.Controls.add("VB.TextBox", "TextBox3" & paramName)
     pTextBox3.Top = 1200 + pAdditionalHt
     pTextBox3.Height = 375
     pTextBox3.Left = fLeft + 4800
     pTextBox3.Width = 500
     pTextBox3.Visible = True
     pTextBox3.Text = ""
     If (pValueDictionary.Exists("TextBox3" & paramName)) Then
        pTextBox3.Text = pValueDictionary.Item("TextBox3" & paramName)
     End If
     Set pTextBox3.Container = pFrame1
     
     Set pFrame1 = Nothing
     Set pRadioBtn1 = Nothing
     Set pTextBox1 = Nothing
     Set pRadioBtn2 = Nothing
     Set pTextBox2 = Nothing
     Set pRadioBtn3 = Nothing
     Set pTextBox3 = Nothing
     Set pLabel3 = Nothing
     Set pTextBox0 = Nothing
     
End Sub


 
Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    'Call the subroutine to read any timeseries file and find total pollutants
    ModuleDecayFact.CreatePollutantList
    
    'Find total pollutants and adjust form control dimensions and positions
    pTotalPollutants = UBound(gPollutants) + 1
    Dim pFormWidth As Double
    'pFormWidth = 2000 * (1 + pTotalPollutants)
    pFormWidth = 9000
    If (pFormWidth > 9000) Then
        pFormWidth = 9000
    End If
    FrmMinimizeCost.Width = pFormWidth
    
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
 
End Sub




