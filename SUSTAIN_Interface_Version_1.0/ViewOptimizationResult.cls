VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "viewOptimizationResult"
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
Implements ITool
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("CHART", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "View Results On Chart"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = EnableSustain 'True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "View Results On Chart"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "View Optimization Result"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "View Results On Chart"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
   
End Property

Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
  
End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError

    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    
    Dim pPoint As IPoint
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
      
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        Exit Sub
    End If
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim iIDFld As Long
    iIDFld = pFeatureclass.FindField("ID")
    Dim iTypeFld As Long
    iTypeFld = pFeatureclass.FindField("TYPE2")
    Dim pBMPID As Integer
    Dim pBMPType As String
    
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pPoint.Envelope
    ExpandPointEnvelope pEnvelope
    
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
            
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, True)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    pBMPID = 0
    If (Not pFeature Is Nothing) Then
        FlashSelectedFeature pFeature
        pBMPID = pFeature.value(iIDFld)
        pBMPType = pFeature.value(iTypeFld)
    End If
    
    'Continue only if it is a assessment point
    If (Right(Trim(pBMPType), 1) <> "X") Then
        MsgBox "Selected BMP is not an assessment point. Cannot display chart."
        Exit Sub
    End If
    
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
    
    'Dim bContinue
    'bContinue = MsgBox("Do you want to view chart for this BMP ?", vbYesNo)
    'If (bContinue = vbYes) Then
        Load FrmResultChart
        Dim isInitSuccess As Boolean
        Dim evalFactorCodeArray
        Dim evalFactorCodeCount As Integer
        evalFactorCodeArray = GetEvalFactorList(pBMPID, pOutputFolder)
        If (evalFactorCodeArray(0) = "") Then
            MsgBox "Required result files missing in " & pOutputFolder
            Exit Sub
        End If
        evalFactorCodeCount = UBound(evalFactorCodeArray) + 1
        'Open the result chart
        FrmResultChart.InitForm evalFactorCodeCount

        isInitSuccess = InitEvaluationChart(pBMPID, pOutputFolder)
        If (isInitSuccess) Then
            FrmResultChart.Show vbModal
        Else
            MsgBox "Required result files missing in " & pOutputFolder
        End If

   'End If
    
    GoTo CleanUp
ShowError:
    MsgBox "ViewOptimizationResult_ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pOptimizationDetail = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub


Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub


