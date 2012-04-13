VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeleteAssessPoint"
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
    Set m_Bitmap = LoadResPicture("DeleteAssess", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Delete Assessment Point"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableEditAssessPoints
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Delete Assessment Point"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "DeleteAssessPoint"
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
    ICommand_Tooltip = "Delete Assessment Point"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
   ITool_Cursor = LoadResPicture("DIGITIZE", vbResCursor)
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

'*** Delete assessment point properties for bmp from all relevant layers & tables.
'*** Re-render bmp feature layer

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
          
      End If
      
     Dim bContinue1
     If (pPreviouslyDefined = True) Then   'If it was previously defined, ask for DELETE or OVERWRITE option
            bContinue1 = MsgBox("Do you want to delete Assessment Point for this BMP?", vbYesNo)
            If (bContinue1 = vbNo) Then
                  Exit Sub
            End If
     End If
    
         '*** Delete from optimization table
     DeletePreviouslyDefinedOptions pBMPID

     '*** Update the BMP Detail table
     Dim pBMPDetail As iTable
     Set pBMPDetail = GetInputDataTable("BMPDetail")
     Dim iPropValue As Long
     iPropValue = pBMPDetail.FindField("PropValue")
     pQueryFilter.WhereClause = "ID = " & pBMPID & " AND PropName = 'isAssessmentPoint'"
     Dim pCursor As ICursor
     Dim pRow As iRow
     Set pCursor = pBMPDetail.Search(pQueryFilter, True)
     Set pRow = pCursor.NextRow
     If Not (pRow Is Nothing) Then
        pRow.value(iPropValue) = "False"
        pRow.Store
     End If

    '*** Update the BMPs feature layer
    pQueryFilter.WhereClause = "ID = " & pBMPID
    Dim iType As Long
    iType = pFeatureclass.FindField("TYPE")
    Dim iType2 As Long
    iType2 = pFeatureclass.FindField("TYPE2")
    Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    If Not (pFeature Is Nothing) Then
        pFeature.value(iType2) = pFeature.value(iType)
        pFeature.Store
    End If
    
    '*** Update the BMP rendering
    RenderSchematicBMPLayer pFeatureLayer
    
    GoTo CleanUp
ShowError:
    MsgBox "DeleteAssessPoint_ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pOptimizationTable = Nothing
    Set pBMPDetail = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub



