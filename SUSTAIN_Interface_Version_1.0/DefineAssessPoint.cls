VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefineAssessPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    Set m_Bitmap = LoadResPicture("EditAssess", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Assessment Points"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableDefineAssessPoints
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Assessment Points"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "EditAssessPoints"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
'    Call InitializeOperators
'    If (ModuleUtility.CheckSpatialAnalystLicense = False) Then
'        MsgBox "Spatial Analyst Not Available."
'        Exit Sub
'    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Assessment Points"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
   ITool_Cursor = LoadResPicture("Assess", vbResCursor)
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

'*** Check for optimization type and open respective evaluation factor form
'*** Update property of the bmp as isAssessment = TRUE
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
    Dim pBMPID As Integer
    
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
    End If
    
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    
    'Define the option value = -1
    Dim pOptionVal As Integer
    pOptionVal = -1
    'Check for existing parameters for same BMP
    If (pBMPID <= 0) Then
        Exit Sub
    End If
    
    'Get landuse reclassification table: LUReclass, Create new if not found
     Dim pPreviouslyDefined As Boolean
     pPreviouslyDefined = False
     Dim pOptimizationTable As iTable
     Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
     If Not (pOptimizationTable Is Nothing) Then
          'Query in the table for existing records
          Dim pQueryFilter As IQueryFilter
          Set pQueryFilter = New QueryFilter
          pQueryFilter.WhereClause = "ID = " & pBMPID
          If (pOptimizationTable.RowCount(pQueryFilter) > 0) Then
              pPreviouslyDefined = True
          End If
    
          'Get the option from table
          pQueryFilter.WhereClause = "ID = 0 and PropName = 'Option'"
          Dim pCursor As ICursor
          Set pCursor = pOptimizationTable.Search(pQueryFilter, True)
          Dim pRow As iRow
          Set pRow = pCursor.NextRow
          If Not pRow Is Nothing Then
              pOptionVal = pRow.value(pCursor.FindField("PropValue"))
          End If
        
      End If
      
      Dim bContinue1
      
      If (pPreviouslyDefined = True) Then   'If it was previously defined, ask for DELETE or OVERWRITE option
              
            bContinue1 = MsgBox("Optimization parameters are already defined for this Assessment Point. Do you want to overwrite ?", vbYesNo)
            If (bContinue1 = vbNo) Then
                  Exit Sub
            End If
       Else
            bContinue1 = MsgBox("Do you want to define selected BMP as Assessment point ?", vbYesNo)
            If (bContinue1 = vbNo) Then
                Exit Sub
            End If
       End If
    
               
    'Sleep for 100 milliseconds, to avoid subscript out of array error
    Sleep 100
    
    'Select option and open respective form
     Select Case pOptionVal
         Case 0:     'Existing condition
             'Show the form
             FrmExistingCondition.txtBMPID.Text = Str(pBMPID)
             FrmExistingCondition.Show vbModal
         Case 1:     'Minimize Cost
             'Show the form
             FrmMinimizeCost.txtBMPID.Text = Str(pBMPID)
             FrmMinimizeCost.Show vbModal
         Case 2:     'Maximize Benefit Control
             'Show the form
             FrmTradeoffOptions.txtBMPID.Text = Str(pBMPID)
             FrmTradeoffOptions.Show vbModal
         Case 3:     'Create trade off curve
             'Show the form
             FrmMaximizeBenefit.txtBMPID.Text = Str(pBMPID)
             FrmMaximizeBenefit.Show vbModal
     End Select
    
    GoTo CleanUp
    
ShowError:
    MsgBox "DefineAssessPoints_ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pOptimizationTable = Nothing
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


