VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeleteBMPs"
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
    Set m_Bitmap = LoadResPicture("DeleteBMP", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub


Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Delete Existing BMPs or VFSs"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableDeleteBMP
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Delete Existing BMPs or VFSs"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "DeleteBMPs"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
'    Call InitializeOperators
'    If (ModuleUtility.CheckSpatialAnalystLicense = False) Then
'        MsgBox "Spatial Analyst Not Available."
'        Exit Sub
'    End If
'    Call DeleteSelectedBMP(8)
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Delete Existing BMPs or VFSs"
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

'*** Delete BMP from relevant feature layer and tables, refresh the feature layer
' Ying Cao: Delete VFS if it is selected  11/20/2008
Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError

    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    
    Dim pPoint As IPoint
    Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
      
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    
    Dim pVFSFeatureLayer As IFeatureLayer
    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")
    
    If (pFeatureLayer Is Nothing And pVFSFeatureLayer Is Nothing) Then
        Exit Sub
    End If
    
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim iIDFld As Long
    iIDFld = pFeatureclass.FindField("ID")
    Dim pBMPID As Integer
    Dim pVFSID As Integer

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
    
    If pFeature Is Nothing Then
        MsgBox "No BMP is found."
        Exit Sub
    End If
    
    If (Not pFeature Is Nothing) Then
        FlashSelectedFeature pFeature
        pBMPID = pFeature.value(iIDFld)
    End If
    
    '**** PLACEHOLDER
    '** If pBMPID = 0, may be the user wants to delete the VFS
    '**** PLACEHOLDER
        
    '** Cleanup is required before passing to other functions
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureLayer = Nothing
    Set pFeatureclass = Nothing
    'Set pEnvelope = Nothing
    'Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    
    If (pBMPID > 0) Then
        Dim bContinue
        bContinue = MsgBox("Do you want to delete selected BMP ?", vbYesNo)
        If (bContinue = vbYes) Then
            DeleteSelectedBMP (pBMPID)
            Call UpdateBMPRelatedDatasets
        End If
    Else
        Set pFeatureclass = pVFSFeatureLayer.FeatureClass
        Set pSpatialFilter.Geometry = pEnvelope
        pSpatialFilter.SpatialRel = esriSpatialRelIntersects
        Set pFeatureCursor = pFeatureclass.Search(pSpatialFilter, False)
        Set pFeature = pFeatureCursor.NextFeature
        pVFSID = 0
        If (Not pFeature Is Nothing) Then
            FlashSelectedFeature pFeature
            pVFSID = pFeature.value(iIDFld)
        End If
        If (pVFSID > 0) Then 'pVFSType
            Dim vContinue
            vContinue = MsgBox("Do you want to delete this VFS ?", vbYesNo)
            If (vContinue = vbYes) Then
                '*** subroutine to delete vfs
                DeleteSelectedVFS pVFSID
                Call UpdateBMPRelatedDatasets
            End If
        End If
    End If
    
    Exit Sub
ShowError:
    MsgBox "DeleteBMPs_ITool_OnMouseDown: " & Err.description

End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub
