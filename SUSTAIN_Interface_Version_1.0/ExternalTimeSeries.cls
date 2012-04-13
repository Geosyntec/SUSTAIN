VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExternalTimeSeries"
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
    Set m_Bitmap = LoadResPicture("TS", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Assign External Time Series"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableTimeSeriesFactors
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Assign External Time Series"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "ExternalTS"
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
    ICommand_Tooltip = "Assign External Time Series"
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
    
    If (pBMPID > 0) Then
        'Get landuse reclassification table: LUReclass, Create new if not found
        Dim pExternalTSTable As iTable
        Set pExternalTSTable = GetInputDataTable("ExternalTS")
        If Not (pExternalTSTable Is Nothing) Then
            'Query in the table for existing records
            Dim pQueryFilter As IQueryFilter
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "BMPID = " & pBMPID
            If (pExternalTSTable.RowCount(pQueryFilter) > 0) Then
                Dim bContinue
                bContinue = MsgBox("External Timeseries is already defined for this BMP. Do you want to overwrite ?", vbYesNo)
                If (bContinue = vbYes) Then
                    pExternalTSTable.DeleteSearchedRows pQueryFilter
                Else
                    Exit Sub
                End If
            End If
        End If

        'Display the external time series dialog box
        frmExternalTS.txtBMP.Text = Str(pBMPID)
        frmExternalTS.Show vbModal
        
        'Render the BMP Feature Layer
        RenderSchematicBMPLayer pFeatureLayer

    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "ExternalTimeSeries_ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pFeatureclass = Nothing
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pFeatureLayer = Nothing
    Set pQueryFilter = Nothing
    Set pExternalTSTable = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub
