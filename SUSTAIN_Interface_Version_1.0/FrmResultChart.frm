VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form FrmResultChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assessment Point Evaluation Functions"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "FrmResultChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "To view optimization results from any other directory, browse the output folder containing optimization results."
      Top             =   240
      Width           =   495
   End
   Begin MSChart20Lib.MSChart DUMMYCHART 
      Height          =   495
      Left            =   8280
      OleObjectBlob   =   "FrmResultChart.frx":08CA
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab TabCharts 
      Height          =   5400
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   7600
      _ExtentX        =   13414
      _ExtentY        =   9525
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmResultChart.frx":2C20
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Plots"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox OutputDir 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox BMPId 
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Cost Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label CostLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "COST"
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   8160
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Output Directory"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmResultChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
On Error GoTo ErrorHandler:

    Dim strTmpDir As String
    'strTmpDir = BrowseForFolder(0, "Select the output folder containing simulation results.")
    strTmpDir = BrowseForSpecificFolder("Select the output folder containing simulation results.", gApplicationPath)
    
    'strTmpDir = Left(strTmpDir, Len(strTmpDir) - 1)
    If (Trim(strTmpDir) <> "") Then
        OutputDir.Text = strTmpDir
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error while reading the input file: " & Err.description
CleanUp:
    Set fso = Nothing
    Set pFile = Nothing
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub


Private Sub cmdRefresh_Click()
    Dim pBMPID As Integer
    pBMPID = CInt(bmpId.Text)
    Dim pOutDir As String
    pOutDir = OutputDir.Text
    
    Dim evalFactorCodeArray
    Dim evalFactorCodeCount As Integer
    evalFactorCodeArray = GetEvalFactorList(pBMPID, pOutDir)
    If (evalFactorCodeArray(0) = "") Then
        MsgBox "Required result files missing in " & pOutputFolder
        Exit Sub
    End If
    evalFactorCodeCount = UBound(evalFactorCodeArray) + 1
    'Open the result chart
    FrmResultChart.InitForm evalFactorCodeCount
        
    Dim isSuccess As Boolean
    isSuccess = InitEvaluationChart(pBMPID, pOutDir)
    If (isSuccess = False) Then
        MsgBox "Cannot refresh chart. Errors in generating charts for data in " & pOutDir & " folder.", vbExclamation
    End If
End Sub

Public Sub PlotEvaluationChart(evalFactorDict As Scripting.Dictionary, ByVal chartNum As Integer, Optional ByVal Target As Double, Optional ByVal costs, Optional yAxisLabel As String)

On Error GoTo ErrorHandler:
    
    Dim numCols As Long
    numCols = evalFactorDict.Count
    
    Dim evalScens
    evalScens = evalFactorDict.keys

    Dim curChart As MSChart
    Set curChart = Nothing
    
    Dim pControl As Control
    Dim pTotalRows As Integer
    Dim numRows As Long
    
    For Each pControl In Controls
        If ((TypeOf pControl Is MSChart)) Then
            If (pControl.name = "Chart" & chartNum) Then
                Set curChart = pControl
                Exit For
            End If
        End If
    Next pControl

    'Need to add Error handling for no chart found
    If curChart Is Nothing Then
        Exit Sub
    End If

    pTotalRows = 0
    If (Target <> -99) Then
        pTotalRows = 1
    End If
    
    curChart.RowCount = pTotalRows + 1  'Add one row for other conditions (pre, post, existing)
    curChart.ColumnCount = numCols
    
    Dim numCol As Long
    numCol = 0
    
    Dim yAxis1Max As Double
    Dim yAxis2Max As Double
    yAxis1Max = 0
    yAxis2Max = 0
    
    Dim pkey
    Dim curEvalValue As Double

    numRows = 0
    'first plot target value
    If (Target <> -99) Then
        numRows = 1   'Add one for existing conditions
        For numCol = 1 To numCols
            curChart.DataGrid.SetData numRows, numCol, Target, nullflag
        Next numCol
        
        If (yAxis1Max < Target) Then
            yAxis1Max = Target
        End If
        
        'Define color/style for first row
        curChart.Plot.SeriesCollection(numRows).SeriesType = VtChSeriesType2dLine
        curChart.Plot.SeriesCollection(numRows).Pen.VtColor.Set 0, 0, 255
        curChart.Plot.SeriesCollection(numRows).LegendText = "Target"
        curChart.Plot.SeriesCollection(numRows).SecondaryAxis = False
    End If
   
    numRows = numRows + 1
    numCol = 0
    For Each pkey In evalScens
        numCol = numCol + 1
        curEvalValue = evalFactorDict.Item(pkey)
        curChart.DataGrid.SetData numRows, numCol, curEvalValue, nullflag
        
        curChart.Column = numCol
        curChart.ColumnLabel = pkey
        
        If (yAxis1Max < curEvalValue) Then
            yAxis1Max = curEvalValue
        End If
              
    Next pkey
    
    'Define color/style for third column
    curChart.Plot.SeriesCollection(numRows).SeriesType = VtChSeriesType2dBar
    curChart.Plot.SeriesCollection.Item(numRows).DataPoints.Item(-1).Brush.FillColor.Set 0, 200, 200
    curChart.Plot.SeriesCollection(numRows).LegendText = "Factor"
    curChart.Plot.SeriesCollection(numRows).SecondaryAxis = False
    
    'Set the y axis title
    If (Not IsMissing(yAxisLabel)) Then
        curChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle = yAxisLabel
    End If
    curChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.VtFont.Size = 14
    curChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.VtFont.name = "Times New Roman"
    curChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.TextLayout.Orientation = VtOrientationUp
    
    With curChart.Plot.Axis(VtChAxisIdY)
      .CategoryScale.Auto = True
    End With
    
    GoTo CleanUp
ErrorHandler:
    MsgBox "Error in PlotEvaluationChart" & Err.description
CleanUp:
    Set curChart = Nothing
    Set pControl = Nothing
End Sub


Public Sub DisplayCostValuesOnChartFrame(costs)
    'If no costs available
    If (costs(0) = -9999.9) Then
        CostLabel.Caption = "No cost values available."
        Exit Sub
    End If
    
    Dim strCostLabel As String
    strCostLabel = ""
    
    Dim iC As Integer
    For iC = 0 To UBound(costs)
        strCostLabel = strCostLabel + "Best " & (iC + 1) & ": $" & Format(costs(iC), "##,##0.0") & vbNewLine
    Next
    'Set the cost value label
    CostLabel.Caption = strCostLabel
    
End Sub


Public Sub InitForm(pTabCount As Integer) 'Public Sub Form_Activate()
On Error GoTo ErrorHandler:
    
    Dim pControlNameDict As Scripting.Dictionary
    Set pControlNameDict = CreateObject("Scripting.Dictionary")
    
    'Remove all controls from form
    For Each pControl In FrmResultChart.Controls
        If (TypeOf pControl Is MSChart) Then
            pControlNameDict.add pControl.name, True
        End If
    Next pControl
    
    
    'Sets the number of tabs
    TabCharts.Tabs = pTabCount
    If (pTabCount > 4) Then
        TabCharts.TabsPerRow = 4
    Else
        TabCharts.TabsPerRow = pTabCount
    End If
    
    Dim pTabHtDbl As Double
    pTabHtDbl = Format(CDbl(pTabCount / 4), "0")
    Dim pTabHeight As Integer
    pTabHeight = CInt(pTabHtDbl) * 300
    Dim pTabWidth As Integer
    pTabWidth = CInt(pTabHtDbl) * 100
    
    FrmResultChart.Height = 7300 + pTabHeight
    TabCharts.Height = 5400 + pTabHeight
    
    Dim pChart As MSChart
    
    Dim incr As Integer
    For incr = 1 To pTabCount
        TabCharts.Tab = incr - 1
        TabCharts.TabCaption(incr - 1) = "Tabular " & incr
       
        If Not (pControlNameDict.Exists("Chart" & incr)) Then   'If exists, go to next control
                    
            Set pChart = FrmResultChart.Controls.add("MSChart20Lib.MSChart", "Chart" & incr)
            Set pChart.Container = TabCharts
            pChart.Visible = True
            pChart.Top = 600 + pTabHeight
            pChart.Left = 400
            pChart.Width = 7000 - pTabWidth
            pChart.Height = 4300
            pChart.chartType = VtChChartType2dCombination
            pChart.Plot.DataSeriesInRow = True
            pChart.RowCount = 1  'Only one row for now
    
            pChart.Backdrop.Fill.Style = VtFillStyleBrush
            ' Sets chart fill color to red.
            With pChart.Backdrop.Fill.Brush.FillColor
               .Red = 255   ' Use properties to set color.
               .Green = 255
               .Blue = 255
            End With
            
            
            'Set Chart title
            pChart.Title.Text = "Plot of Evaluation Functions"
            pChart.Title.VtFont.Style = VtFontStyleBold
            pChart.Title.VtFont.Size = 12
            pChart.Title.VtFont.name = "Times New Roman"
            
            'Set the x axis title
            pChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.Text = "Scenarios"
            pChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.VtFont.Size = 12
            pChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.VtFont.name = "Times New Roman"
           
            pChart.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
            
            pChart.ShowLegend = True
            Set pChart = Nothing
        
        End If
    Next incr
        
    TabCharts.Tab = 0
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error in InitForm " & Err.description
CleanUp:
    Set pChart = Nothing
    
End Sub

Private Function GetMaxScale(maxValue As Double)
    GetMaxScale = CInt(maxValue * 1.1)
End Function

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
