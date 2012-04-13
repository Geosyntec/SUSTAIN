VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewTradeoffCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 03/29/2005 - mira chokshi added comments to project
'
'******************************************************************************


Option Explicit
Implements ICommand

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Show Cost Effectiveness Curve"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = EnableSustain  'True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Show Cost Effectiveness Curve"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Show Cost Effectiveness Curve"
End Property

Private Sub ICommand_OnClick()
    InitializeMapDocument
    
    Dim X() As Double, Y() As Double
    Dim pOutputDirName As String
    
    Dim pOptimizationDetail As iTable
    Set pOptimizationDetail = GetInputDataTable("OptimizationDetail")
    If (pOptimizationDetail Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'OutputFolder'"
    Dim pCursor As ICursor
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    Dim pOutputFolder As String
    pOutputFolder = ""
    If Not (pRow Is Nothing) Then
        pOutputFolder = Trim(pRow.value(iPropValue))
    End If
    
    If (pOutputFolder = "") Then
        MsgBox "Output folder not found."
        Exit Sub
    End If
    
    'Find the BMPID
    pQueryFilter.WhereClause = "ID <> 0"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    If pRow Is Nothing Then MsgBox "No assessment point defined ": Exit Sub
    Dim bmpId As Integer
    bmpId = pRow.value(pCursor.FindField("ID"))
    
    Dim strCalMode As String
    strCalMode = Trim(pRow.value(iPropValue))
    Dim listCalMode
    listCalMode = Split(strCalMode, ",")
    
    'pOutputDirName = "C:\Temp\sustain\out\GA_TradeOff"
    
    Call SetGASolution(pOutputFolder, X, Y, bmpId, CInt(listCalMode(3)))
    
    Dim yAxisTitle As String
    
    Select Case CInt(listCalMode(1))
        Case -1
            yAxisTitle = "Flow Volume (ft3/yr)"
        Case -2
            yAxisTitle = "Flow Rate (cfs)"
        Case -3
            yAxisTitle = "Flow Exceedence Frequency (%)"
        Case 1
            yAxisTitle = "Load (kg/yr)"
        Case 2
            yAxisTitle = "Concentration (mg/L)"
        Case 3
            yAxisTitle = "Concentration (mg/L)"
    End Select
    
    Call CreatePollutantList
    If CInt(listCalMode(0)) > 0 Then
        If CInt(listCalMode(0)) - 1 < UBound(gPollutants) Then
            yAxisTitle = gPollutants(CInt(listCalMode(0)) - 1) & " " & yAxisTitle
        End If
    End If
    yAxisTitle = yAxisTitle & " % of Existing"
    FrmTradeOff.InitializeTradeOffCurve X, Y, "NSGAII", yAxisTitle
    FrmTradeOff.Show vbModal
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Show Cost Effectiveness Curve"
End Property









