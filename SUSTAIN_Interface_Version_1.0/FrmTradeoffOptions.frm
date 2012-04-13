VERSION 5.00
Begin VB.Form FrmTradeoffOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Effectiveness Curve Options"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "FrmTradeoffOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Evaluation Factor Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   7215
      Begin VB.TextBox txtTargetMin 
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   750
         Width           =   1335
      End
      Begin VB.TextBox txtTargetMax 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   1215
         Width           =   1335
      End
      Begin VB.TextBox txtThreshold 
         Height          =   285
         Left            =   3600
         TabIndex        =   6
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Lower Target Value"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Upper Target Value"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label LabelThreshold 
         Caption         =   "Threshold"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label LabelUnit 
         Caption         =   "days average concentration"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set Search Stopping Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   7215
      Begin VB.TextBox MaxRunTime 
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum search time allowed"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "hour"
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   345
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Evaluation Factor Option"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   7215
      Begin VB.OptionButton optSpecified 
         Caption         =   "Specified value"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   4215
      End
      Begin VB.OptionButton optFraction 
         Caption         =   "Between pre-development and existing condition (0-1)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton optPercent 
         Caption         =   "Percent of value under existing condition (0-100)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   5535
      End
   End
   Begin VB.ComboBox cbxEvalType 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtBMPID 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "bmpId"
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3773
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2453
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox cbxEvalFact 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   978
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Select Evalution Factor Type (Pollutant/Flow)"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Select evaluation factor"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select the evaluation factor and input control targets."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "FrmTradeoffOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0


Private Sub cbxEvalFact_click()
    Dim evalFact As String
    evalFact = cbxEvalFact.List(cbxEvalFact.ListIndex)
    If evalFact = "Exceedance frequency (per year)" Or _
       evalFact = "Exceeding average concentration" Then
        LabelThreshold.Visible = True
        txtThreshold.Visible = True
        txtThreshold.Enabled = True
        LabelUnit.Visible = True
    Else
        LabelThreshold.Visible = False
        txtThreshold.Visible = False
        txtThreshold.Enabled = False
        LabelUnit.Visible = False
    End If
End Sub

'Private Sub cbxEvalType_Change()
'    Dim evalType As String
'    evalType = cbxEvalType.List(cbxEvalType.ListIndex)
'    cbxEvalFact.Clear
'    If evalType = "Flow" Then
'        cbxEvalFact.AddItem "Average annual flow volume"
'        cbxEvalFact.AddItem "Peak discharge flow"
'        cbxEvalFact.AddItem "Exceedance frequency (per year)"
'        LabelThreshold.Caption = "Threshold"
'        LabelUnit.Caption = "cfs"
'    Else
'        cbxEvalFact.AddItem "Average annual load"
'        cbxEvalFact.AddItem "Average annual concentration"
'        cbxEvalFact.AddItem "Exceeding average concentration"
'        LabelThreshold.Caption = "Maximum"
'        LabelUnit.Caption = "days average concentration"
'    End If
'    cbxEvalFact.ListIndex = 0
'    cbxEvalFact.Refresh
'
'End Sub


Private Sub cbxEvalType_Click()
    Dim evalType As String
    evalType = cbxEvalType.List(cbxEvalType.ListIndex)
    cbxEvalFact.Clear
    If evalType = "Flow" Then
        cbxEvalFact.AddItem "Average annual flow volume"
        cbxEvalFact.AddItem "Peak discharge flow"
        cbxEvalFact.AddItem "Exceedance frequency (per year)"
        LabelThreshold.Caption = "Threshold"
        LabelUnit.Caption = "cfs"
    Else
        cbxEvalFact.AddItem "Average annual load"
        cbxEvalFact.AddItem "Average annual concentration"
        cbxEvalFact.AddItem "Exceeding average concentration"
        LabelThreshold.Caption = "Maximum"
        LabelUnit.Caption = "days average concentration"
    End If
    cbxEvalFact.ListIndex = 0
    cbxEvalFact.Refresh
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    'Make current tool inactive
    DeactivateCurrentTool
End Sub

Private Sub cmdDone_Click()
On Error GoTo ErrorHandler
'    Dim numBreaks As Integer
'     If (Trim(txtNumBreaks.Text) = "" Or Not IsNumeric(txtNumBreaks.Text)) Then
'        MsgBox "Please specify integer value for number of breaks."
'        Exit Sub
'    End If
    
    If (Trim(txtTargetMin.Text) = "" Or Not IsNumeric(txtTargetMin.Text)) Then
        MsgBox "Please specify a number for Lower Target Value."
        Exit Sub
    End If
    If (Trim(txtTargetMax.Text) = "" Or Not IsNumeric(txtTargetMax.Text)) Then
        MsgBox "Please specify a number for Upper Target Value."
        Exit Sub
    End If

'    If Not IsNumeric(CostLimit.Text) Then
'        MsgBox "Please specify a number for stopping cost limit."
'        Exit Sub
'    End If
    If Not IsNumeric(MaxRunTime.Text) Then
        MsgBox "Please specify a number for maximum run time."
        Exit Sub
    End If
    
'    numBreaks = CInt(txtNumBreaks.Text)
    
    Dim pBMPID As Integer
    pBMPID = CInt(txtBMPID.Text)
    
    Dim pOptimizationDetail As iTable
    Set pOptimizationDetail = GetInputDataTable("OptimizationDetail")
    If (pOptimizationDetail Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iIDindex As Long
    iIDindex = pOptimizationDetail.FindField("ID")
    Dim iPropNameIndex As Long
    iPropNameIndex = pOptimizationDetail.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pOptimizationDetail.FindField("PropValue")
    
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID > 0 "
    pOptimizationDetail.DeleteSearchedRows pQueryFilter

''    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'NumBreak'"
''    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
''    Set pRow = pCursor.NextRow
''    If (pRow Is Nothing) Then
''        Set pRow = pOptimizationDetail.CreateRow
''    End If
''    pRow.value(iPropNameIndex) = "NumBreak"
''    pRow.value(iPropValueIndex) = numBreaks
''    pRow.Store
       
    Dim evalType As String
    Dim evalFact As String
    
    Dim pFactorName As String
    Dim FactorGroup As Integer
    FactorGroup = -1
    
    Dim FactorType As Integer
    Dim CalcDays As Integer
    CalcDays = 0
        
    Dim CalcMode As Integer
    CalcMode = 1
    If optFraction Then CalcMode = 2
    If optSpecified Then CalcMode = 3
    
    evalType = cbxEvalType.List(cbxEvalType.ListIndex)
        
    If evalType <> "Flow" Then
         pFactorName = evalType
         FactorGroup = cbxEvalType.ListIndex - 1
    End If
   
    
    Select Case cbxEvalFact.List(cbxEvalFact.ListIndex)
        Case "Average annual flow volume":
            pFactorName = "AAFV"
            FactorType = -1
        Case "Peak discharge flow":
            pFactorName = "PDF"
            FactorType = -2
        Case "Exceedance frequency (per year)":
             pFactorName = "FEF"
             FactorType = -3
             If IsNumeric(txtThreshold.Text) Then
                CalcDays = CDbl(txtThreshold.Text)
             Else
                MsgBox LabelThreshold.Caption & " needs to be real number"
                Exit Sub
             End If
        Case "Average annual load":
            pFactorName = evalType & "_AAL"
            FactorType = 1
        Case "Average annual concentration":
            pFactorName = evalType & "_AAC"
            FactorType = 2
        Case "Exceeding average concentration":
            pFactorName = evalType & "_MAC"
            FactorType = 3
            If IsNumeric(txtThreshold.Text) Then
               CalcDays = CInt(txtThreshold.Text)
            Else
                MsgBox LabelThreshold.Caption & " needs to be an integer number"
                Exit Sub
            End If
        Case Else
    End Select
    
    Dim targetMin As Double
    Dim targetMax As Double
    targetMin = CDbl(txtTargetMin.Text)
    targetMax = CDbl(txtTargetMax.Text)
''    Dim strOptions As String
''    strOptions = FactorGroup & "," & FactorType & "," & _
''        CalcDays & "," & CalcMode & "," & targetMin & "," & targetMax & "," & pFactorName
    'Remove all BMP specific optimization options from the table
'''    Set pQueryFilter = New QueryFilter
'''    pQueryFilter.WhereClause = "ID > 0 "
'''    pOptimizationDetail.DeleteSearchedRows pQueryFilter
'''
'''    Set pRow = pOptimizationDetail.CreateRow
'''    pRow.value(iIDindex) = CInt(txtBMPID.Text)
'''    pRow.value(iPropNameIndex) = "Parameters"
'''    pRow.value(iPropValueIndex) = strOptions
'''    pRow.Store
'''
'''    Dim pBMPFLayer As IFeatureLayer
'''    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
'''    Dim pBMPFClass As IFeatureClass
'''    Set pBMPFClass = pBMPFLayer.FeatureClass
'''    Set pQueryFilter = New QueryFilter
'''    pQueryFilter.WhereClause = "ID = " & CInt(txtBMPID.Text)
'''    Dim pFeatureCursor As IFeatureCursor
'''    Set pFeatureCursor = pBMPFClass.Search(pQueryFilter, True)
'''    Dim pFeature As IFeature
'''    Set pFeature = pFeatureCursor.NextFeature
'''    Dim pBMPType2Val As String
'''    If Not (pFeature Is Nothing) Then
'''        'Set the Type2 parameter to designate assessment point
'''        pBMPType2Val = Trim(pFeature.value(pFeatureCursor.FindField("TYPE2")))
'''        If (Right(pBMPType2Val, 1) <> "X") Then
'''            pFeature.value(pFeatureCursor.FindField("TYPE2")) = pFeature.value(pFeatureCursor.FindField("TYPE2")) & "X"
'''            pFeature.Store
'''        End If
'''    End If
    
    'DefineOptimizationMethod 2, 2000000, 0, 0.5, 2
    
'    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'CostLimit'"
'    Set pCursor = pOptimizationDetail.Update(pQueryFilter, False)
'    Set pRow = pCursor.NextRow
'    If pRow Is Nothing Then
'        Set pRow = pOptimizationDetail.CreateRow
'    End If
'    pRow.value(iPropValueIndex) = CostLimit.Text
'    pRow.Store
    
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'MaxRunTime'"
    Set pCursor = pOptimizationDetail.Update(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    If pRow Is Nothing Then
        Set pRow = pOptimizationDetail.CreateRow
    End If
    pRow.value(iPropValueIndex) = MaxRunTime.Text
    pRow.Store
    
    'pOptimizationDetail.DeleteSearchedRows Nothing
    Call DeleteAllAssessmentPointsDetails

    Dim pBMPOptimizationColl As Collection
    Set pBMPOptimizationColl = New Collection
    pBMPOptimizationColl.add Array(FactorGroup, FactorType, CalcDays, CalcMode, targetMin, targetMax, pFactorName)

    If (pBMPOptimizationColl.Count > 0) Then
        UpdateOptimizationParamsForBMP pBMPID, pBMPOptimizationColl
    End If
    
   
    Unload Me
    DeactivateCurrentTool
    GoTo CleanUp
ErrorHandler:
    MsgBox "Error saving optimization option :" & Err.description
CleanUp:
    Set pRow = Nothing
    Set pOptimizationDetail = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
'    Set pBMPFLayer = Nothing
'    Set pBMPFClass = Nothing
'    Set pFeatureCursor = Nothing
'    Set pFeature = Nothing
    Set pBMPOptimizationColl = Nothing
End Sub

Private Sub Form_Activate()
On Error GoTo ShowError
    Dim lineNum As Integer
    lineNum = 1
    cbxEvalType.AddItem "Flow", 0
    Call CreatePollutantList
    If UBound(gPollutants) < 0 Then Exit Sub
    Dim i As Integer
    For i = 0 To UBound(gPollutants)
        cbxEvalType.AddItem gPollutants(i), i + 1
    Next
    cbxEvalType.ListIndex = 0
    cbxEvalType.Refresh
    
    lineNum = 2
    Dim pOptimizationTable As iTable
    Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
    If (pOptimizationTable Is Nothing) Then Exit Sub

    'Query in the table for existing records
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
       
    lineNum = 3
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iPropValue As Long
    iPropValue = pOptimizationTable.FindField("PropValue")
    
    lineNum = 4
'    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'NumBreak'"
'    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
    
    Dim pRowString As String
'    Dim numBreaks As Integer
'
'    Set pRow = pCursor.NextRow
'    If Not pRow Is Nothing Then
'        pRowString = pRow.value(iPropValue)
'        numBreaks = CInt(pRowString)
'    End If
    
    lineNum = 5
    pQueryFilter.WhereClause = "ID = " & txtBMPID.Text
    lineNum = 51
    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
    lineNum = 52
    Set pRow = pCursor.NextRow
    
    Dim pSplittedString
    Dim pFactorName As String
    Dim pFactorGroup As String
    Dim pFactorType As String
    Dim pCalcDays As Double
    Dim pCalcMode As String
    Dim targetMin As Double
    Dim targetMax As Double
    
    lineNum = 6
    If Not pRow Is Nothing Then
        pRowString = pRow.value(iPropValue)
        pSplittedString = Split(pRowString, ",")
        pFactorGroup = pSplittedString(0)  'FACTOR-GROUP: 1,2,3 for pollutants, -1 for flow
        pFactorType = pSplittedString(1)   'FACTOR-TYPE: -1, -2, -3 for AAFV, PDF, FEF, 1,2,3 for AAL, AAC, MAC
        pCalcDays = CInt(pSplittedString(2))
        pCalcMode = pSplittedString(3)
        lineNum = 7
        If UBound(pSplittedString) = 6 Then
            targetMin = CDbl(pSplittedString(4))
            targetMax = CDbl(pSplittedString(5))
            txtTargetMin.Text = targetMin
            txtTargetMax.Text = targetMax
'            txtNumBreaks.Text = numBreaks
        End If
        lineNum = 8
        '-1 flow others in the order of pollutants
        If CInt(pFactorGroup) = -1 Then
            cbxEvalType.ListIndex = 0
            If (-1 * CInt(pFactorType)) - 1 > 0 And (-1 * CInt(pFactorType)) - 1 < cbxEvalFact.ListCount Then cbxEvalFact.ListIndex = (-1 * CInt(pFactorType)) - 1
        Else
            If CInt(pFactorGroup) + 1 < cbxEvalType.ListCount Then cbxEvalType.ListIndex = CInt(pFactorGroup) + 1
            If CInt(pFactorType) - 1 > 0 And CInt(pFactorType) - 1 < cbxEvalFact.ListCount Then cbxEvalFact.ListIndex = CInt(pFactorType) - 1
        End If
        
        lineNum = 9
        Select Case CInt(pCalcMode)
        Case 1:
            optPercent.value = True
        Case 2:
            optFraction.value = True
        Case 3:
            optSpecified.value = True
        End Select
        If CInt(pFactorType) = 3 Then txtThreshold.Text = pCalcDays
    End If
         
'    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'CostLimit'"
'    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
'    Set pRow = pCursor.NextRow
'    pRowString = pRow.value(iPropValue)
'    If pRowString = "" Then
'        CostLimit.Text = pRowString
'    Else
'        CostLimit.Text = "200000"
'    End If
    
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'MaxRunTime'"
    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    pRowString = pRow.value(iPropValue)
    If pRowString = "" Then
        MaxRunTime.Text = pRowString
    Else
        MaxRunTime.Text = "0.5"
    End If
    Me.Refresh
    Exit Sub
ShowError:
    MsgBox "Error loading TradeOffOptions form :" & Err.description & vbNewLine & "LINE NUM" & lineNum
End Sub

''Private Sub Form_Load()
''On Error GoTo ShowError
''    Dim lineNum As Integer
''    lineNum = 1
''    cbxEvalType.AddItem "Flow", 0
''    Call CreatePollutantList
''    Dim i As Integer
''    For i = 0 To UBound(gPollutants)
''        cbxEvalType.AddItem gPollutants(i), i + 1
''    Next
''    cbxEvalType.ListIndex = 0
''    cbxEvalType.Refresh
''
''    lineNum = 2
''    Dim pOptimizationTable As iTable
''    Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
''    If (pOptimizationTable Is Nothing) Then Exit Sub
''
''    'Query in the table for existing records
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''
''    lineNum = 3
''    Dim pCursor As ICursor
''    Dim pRow As iRow
''    Dim iPropValue As Long
''    iPropValue = pOptimizationTable.FindField("PropValue")
''
''    lineNum = 4
''    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'NumBreak'"
''    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
''
''    Dim pRowString As String
''    Dim numBreaks As Integer
''
''    Set pRow = pCursor.NextRow
''    If Not pRow Is Nothing Then
''        pRowString = pRow.value(iPropValue)
''        numBreaks = CInt(pRowString)
''    End If
''
''    lineNum = 5
''    pQueryFilter.WhereClause = "ID = " & txtBMPID.Text
''    lineNum = 51
''    Set pCursor = pOptimizationTable.Search(pQueryFilter, False)
''    lineNum = 52
''    Set pRow = pCursor.NextRow
''
''    Dim pSplittedString
''    Dim pFactorName As String
''    Dim pFactorGroup As String
''    Dim pFactorType As String
''    Dim pCalcDays As Double
''    Dim pCalcMode As String
''    Dim targetMin As Double
''    Dim targetMax As Double
''
''    lineNum = 6
''    If Not pRow Is Nothing Then
''        pRowString = pRow.value(iPropValue)
''        pSplittedString = Split(pRowString, ",")
''        pFactorGroup = pSplittedString(0)  'FACTOR-GROUP: 1,2,3 for pollutants, -1 for flow
''        pFactorType = pSplittedString(1)   'FACTOR-TYPE: -1, -2, -3 for AAFV, PDF, FEF, 1,2,3 for AAL, AAC, MAC
''        pCalcDays = CDbl(pSplittedString(2))
''        pCalcMode = pSplittedString(3)
''        lineNum = 7
''        If UBound(pSplittedString) = 6 Then
''            targetMin = CDbl(pSplittedString(4))
''            targetMax = CDbl(pSplittedString(5))
''            txtTargetMin.Text = targetMin
''            txtTargetMax.Text = targetMax
''            txtNumBreaks.Text = numBreaks
''        End If
''        lineNum = 8
''        '-1 flow others in the order of pollutants
''        If CInt(pFactorGroup) = -1 Then
''            cbxEvalType.ListIndex = 0
''            cbxEvalFact.ListIndex = (-1 * CInt(pFactorType)) - 1
''        Else
''            cbxEvalType.ListIndex = CInt(pFactorGroup)
''            cbxEvalFact.ListIndex = CInt(pFactorType) - 1
''        End If
''
''        lineNum = 9
''        Select Case CInt(pCalcMode)
''        Case 1:
''            optPercent.value = True
''        Case 2:
''            optFraction.value = True
''        Case 3:
''            optSpecified.value = True
''        End Select
''    End If
''
''    Me.Refresh
''    Exit Sub
''ShowError:
''    MsgBox "Error loading TradeOffOptions form :" & Err.description & vbNewLine & "LINE NUM" & lineNum
''End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
