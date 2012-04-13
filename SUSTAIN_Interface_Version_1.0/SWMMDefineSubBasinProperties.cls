VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SWMMDefineSubBasinProperties"
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
    Set m_Bitmap = LoadResPicture("SUBCATCH", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Subwatershed Properties"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableSWMMWatershedProperties
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Subwatershed Properties"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Define Subwatershed Properties"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Call InitializeOperators(True)
    If (ModuleUtility.CheckSpatialAnalystLicense = False) Then
        MsgBox "Spatial Analyst Not Available."
        Exit Sub
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Subwatershed Properties"
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
    Set pFeatureCursor = pWatershedFClass.Search(pSpatialFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    Dim pX As Double
    Dim pY As Double
    If Not (pFeature Is Nothing) Then
        gSubCatchmentID = pFeature.value(iID)
        pX = pMapPoint.X
        pY = pMapPoint.Y
        '** flash the polygon to confirm
        FlashWatershedFeature pFeature
    End If
   
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
    
    '** open the form to enter properties
    FrmSWMMSubCatchmentProperties.txtName.Text = "SubWtr" & gSubCatchmentID
    FrmSWMMSubCatchmentProperties.txtSubCatchment.Text = gSubCatchmentID
    FrmSWMMSubCatchmentProperties.Show vbModal
    
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
   
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub

