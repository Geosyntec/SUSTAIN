VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form FrmTradeOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimization Results -Cost Effectiveness Curve"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   Icon            =   "FrmTradeOff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4605
      TabIndex        =   1
      Top             =   7440
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart MSChartTradeOff 
      Height          =   6975
      Left            =   240
      OleObjectBlob   =   "FrmTradeOff.frx":08CA
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "FrmTradeOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub InitializeTradeOffCurve(xArray() As Double, yArray() As Double, series1Label As String, yAxisTitle As String)
    Dim myChart As MSChart
    Set myChart = MSChartTradeOff
    
    Dim OldRowCount As Long
    
    Dim i As Integer
    Dim minCost As Double
    Dim xAxisTitle As String, divisor As Double
    xAxisTitle = "BMP Cost ($) "
    divisor = 1
    minCost = xArray(0)
    For i = 1 To UBound(xArray)
        If minCost > xArray(i) Then minCost = xArray(i)
    Next
    
    If minCost > 1000000000 Then
        xAxisTitle = "BMP Cost ($ Billion) "
        divisor = 1000000000
    ElseIf minCost > 1000000 Then
        xAxisTitle = "BMP Cost ($ Million) "
        divisor = 1000000
    ElseIf minCost > 1000 Then
        xAxisTitle = "BMP Cost ($ 1000) "
        divisor = 1000
    End If
    
    With myChart
        'Chart type is XY Scatter
        .chartType = VtChChartType2dXY
        
        'Show the chart legend
        .ShowLegend = True
        
        'Set the back ground color to white
        .Backdrop.Fill.Style = VtFillStyleBrush
        .Backdrop.Fill.Brush.FillColor.Set 255, 255, 255
        
        'Put the legend at the bottom
        .Legend.Location.LocationType = VtChLocationTypeBottom
        
        
        'Use differnt scale for x & y axis
        .Plot.UniformAxis = False
        
        .ColumnCount = 2
        .ColumnLabelCount = 2
        
        'Set the series data
        .RowCount = UBound(xArray) + 1
        Dim lRow As Long
        For lRow = 1 To UBound(xArray) + 1
             .DataGrid.SetData lRow, 1, xArray(lRow - 1) / divisor, False
             .DataGrid.SetData lRow, 2, yArray(lRow - 1), False
        Next
        
        Dim lRow2 As Long
        For lRow2 = lRow To OldRowCount&
            .DataGrid.SetData lRow2, 1, 0, True
            .DataGrid.SetData lRow2, 2, 0, True
        Next
        
        'Set the chart scale to linear
        .Plot.Axis(VtChAxisIdY).AxisScale.Type = VtChScaleTypeLinear
        .Plot.Axis(VtChAxisIdY).AxisScale.Type = VtChScaleTypeLinear
        
        .Plot.Axis(VtChAxisIdY).Tick.Style = VtChAxisTickStyleOutside
        .Plot.Axis(VtChAxisIdX).Tick.Style = VtChAxisTickStyleInside
        .Plot.Axis(VtChAxisIdY).Intersection.LabelsInsidePlot = True
        .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
        'Store the current RowCount
        OldRowCount& = .RowCount
    
        .Column = 1
        .ColumnLabel = series1Label
        
'        .Plot.SeriesCollection(1).SeriesMarker.Show = True
'        .Plot.SeriesCollection(1).Pen.Width = 6
'        .Plot.SeriesCollection(1).Pen.Cap = VtPenCapRound
        
        .Plot.SeriesCollection(1).DataPoints(-1).Marker.Style = VtMarkerStyleX
        .Plot.SeriesCollection(1).DataPoints(-1).Marker.Size = 10

        .Plot.Axis(VtChAxisIdY).AxisTitle.Text = yAxisTitle
        .Plot.Axis(VtChAxisIdX).AxisTitle.Text = xAxisTitle
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
