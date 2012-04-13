VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefineAggBmpLanduse"
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
    Set m_Bitmap = LoadResPicture("AggLuDistribution", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Aggregate BMP Landuse Distribution"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableAggLuDistribution  'ModuleMenu.EnableIndividualBMPTool("Aggregate")
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Aggregate BMP Landuse Distribution"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Define Aggregate BMP Landuse Distribution"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Call InitializeOperators(True)
    If (ModuleUtility.CheckSpatialAnalystLicense = False) Then
        MsgBox "Spatial Analyst Not Available."
        Exit Sub
    End If
    
    Dim boolWatershed As Boolean
    If gSubWaterLandUseDict Is Nothing Then
        boolWatershed = FindAndConvertWatershedFeatureLayerToRaster()
        If (boolWatershed = True) Then
            Call ComputeLanduseAreaForEachSubBasin
        End If
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Aggregate BMP Landuse Distribution"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
'   ITool_Cursor = LoadResPicture("SUBCATCH", vbResCursor)
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
    '** get the watershed on which the mouse is clicked
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If (pWatershedFLayer Is Nothing) Then
        MsgBox "Watershed feature layer not found."
        Exit Sub
    End If
        
    '** define the feature class and ID field index
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    Dim pSubWsID As Integer, selSubWsID As Integer
    Dim iID As Long
    iID = pWatershedFClass.FindField("ID")

    '** get the point on map and transform it
    Dim pDisp As IScreenDisplay
    Set pDisp = gMxDoc.ActiveView.ScreenDisplay
    Dim pDispTrans As IDisplayTransformation
    Set pDispTrans = pDisp.DisplayTransformation
    Dim pMapPoint As IPoint
    Set pMapPoint = pDispTrans.ToMapPoint(X, Y)

    '** expand the point
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pMapPoint.Envelope
    Set pEnvelope = ExpandPointEnvelope(pEnvelope)

    '** define a spatial filter
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects

    '** define the feature cursor and feature for it
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Set pFeatureCursor = pWatershedFClass.Search(pSpatialFilter, False)
    Set pFeature = pFeatureCursor.NextFeature
    
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    
    If pBMPFLayer Is Nothing Then
        MsgBox "Cannot find BMP layer!", vbExclamation
        Exit Sub
    End If
    
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
       
    If Not (pFeature Is Nothing) Then
        'Check to see of any aggregate BMP within subwatershed feature
        Set pSpatialFilter.Geometry = pFeature.Shape
        pSpatialFilter.SpatialRel = esriSpatialRelIntersects
        pSpatialFilter.WhereClause = "TYPE = 'Aggregate'"
        
        If pBMPFClass.FeatureCount(pSpatialFilter) = 0 Then
            MsgBox "There is no aggregate BMP within the selected subwatershed. Try again", vbInformation
            Exit Sub
        End If
        selSubWsID = pFeature.value(iID)
        FlashWatershedFeature pFeature
    End If
    
    'Reset featurecursor and feature
    'Select aggregate BMP features
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    
    Dim pBMPCursor As IFeatureCursor
    Dim pBMPFeature As IFeature
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "TYPE = 'Aggregate'"
    
    Dim pBMPID As Integer
    Dim selIndex As Integer
    selIndex = -1
    
    Dim iBmpIdFld As Integer
    iBmpIdFld = pBMPFClass.FindField("ID")
    
    If iBmpIdFld < 0 Then Err.Raise 5002, , "ID field missing in BMPs feature layer"
    
    
'    Dim treatmentDict As Scripting.Dictionary
'    Set treatmentDict = New Scripting.Dictionary
'
'    Dim treatmentArray
'    treatmentArray = Array(0, 0, 0, 0, 0, 0)
'
'    Dim pSubWsLuGroupDict As Scripting.Dictionary
'    Set pSubWsLuGroupDict = GetSubWsLuGroupDict(pSubWsID)
'
'    Dim luKey
'    For Each luKey In pSubWsLuGroupDict
'        treatmentArray(0) = pSubWsLuGroupDict.Item(luKey)
'        treatmentDict.Item(CStr(luKey)) = treatmentArray
'    Next
    
    If gLuGroupIdDict Is Nothing Then Call SetLuGroupIDDict
    Load FrmAggBmpLuDist
    FrmAggBmpLuDist.cbxSws.Clear
    
    Set pBMPCursor = pBMPFClass.Search(pQueryFilter, False)
    Set pBMPFeature = pBMPCursor.NextFeature
    Set pSpatialFilter = New SpatialFilter
    Do Until pBMPFeature Is Nothing
        pBMPID = pBMPFeature.value(iBmpIdFld)
        Set pSpatialFilter.Geometry = pBMPFeature.Shape
        pSpatialFilter.SpatialRel = esriSpatialRelIntersects
        If pWatershedFClass.FeatureCount(pSpatialFilter) <> 0 Then
            Set pFeatureCursor = pWatershedFClass.Search(pSpatialFilter, False)
            Set pFeature = pFeatureCursor.NextFeature
            pSubWsID = pFeature.value(iID)
            With FrmAggBmpLuDist.cbxSws
                .AddItem pSubWsID
                .ItemData(.NewIndex) = pBMPID
                If pSubWsID = selSubWsID Then selIndex = .NewIndex
            End With
        End If
        Set pBMPFeature = pBMPCursor.NextFeature
    Loop
    If selIndex <> -1 Then FrmAggBmpLuDist.cbxSws.ListIndex = selIndex
'    FrmAggBmpLuDist.InitializeDataGrid treatmentDict
'    FrmAggBmpLuDist.watershedID.Text = pSubWsID
'    FrmAggBmpLuDist.Show vbModal
    
    FrmAggBmpLuDist.Show vbModal
    
    GoTo CleanUp

ShowError:
    MsgBox Err.description
CleanUp:
    '** cleanup
    Set pFeature = Nothing
    Set pFeatureCursor = Nothing
    Set pMapPoint = Nothing
    Set pEnvelope = Nothing
    Set pDisp = Nothing
    Set pDispTrans = Nothing
    Set pSpatialFilter = Nothing
    Set pWatershedFClass = Nothing
    Set pWatershedFLayer = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub


