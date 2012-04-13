VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddVFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created: 06/20/2005 - Mira Chokshi create this module.
'
'******************************************************************************

Option Explicit

Implements ICommand
Implements ITool
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("BufferStrip", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Add Buffer Strip"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableAddVFS
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Add Buffer Strip"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Add Buffer Strip"
End Property

Private Sub ICommand_OnClick()

    Call InitializeMapDocument
    
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Add Buffer Strip"
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

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler
 
 If (gSUBBASINFieldName = "" And gSUBBASINRFieldName = "") Then
    MsgBox "STREAM ID AND DOWNSTREAM ID information for STREAMS feature layer not defined."
    Exit Sub
  End If
  
  'Find the stream direction - SP, Nov 2005
  Call SetSreamDirectionFlag
  
  Dim pLayerStream As IFeatureLayer
  Set pLayerStream = GetInputFeatureLayer("STREAM")
  If pLayerStream Is Nothing Then
    MsgBox "Stream layer is missing.", vbExclamation
    Exit Sub
  End If
  
  Dim pFCStream As IFeatureClass
  Set pFCStream = pLayerStream.FeatureClass
  
  Dim strIDName As String, strDSIDName As String
  strIDName = gSUBBASINFieldName
  strDSIDName = gSUBBASINRFieldName
  
  Dim lStreamIDFldIndex As Long
  lStreamIDFldIndex = pFCStream.FindField(strIDName)
  If lStreamIDFldIndex < 0 Then
    MsgBox "Required field is missing in Stream layer", vbExclamation
    Exit Sub
  End If
  
  FrmTraceDown.Show vbModal
  If Not FrmTraceDown.bContinue Then Exit Sub
  
  Dim pNearFeature As IFeature
  Dim pNearPt As IPoint
  Dim pNearDis As Double
  
  Dim hitTestRes
  If FrmTraceDown.nSnapOption = SNAP_NEAREST_JUNCTION Then
    hitTestRes = SnapToInStreamBMP(pFCStream, X, Y, FrmTraceDown.strSnapBMPType)
  Else
    hitTestRes = HitTestStream(pFCStream, X, Y)
  End If
  
  If IsEmpty(hitTestRes) Then
        MsgBox "BMPs layer not found.", vbInformation
        Exit Sub
  End If
  
  Set pNearFeature = hitTestRes(0)
  Set pNearPt = hitTestRes(1)
  pNearDis = hitTestRes(2)
  If pNearFeature Is Nothing Then
    MsgBox "Cannot snap to a stream at the specified location.", vbExclamation
    Exit Sub
  End If
  
  Dim strStreamID As String
  Dim pPolyline As IPolyline
  Dim pBufferStrip As IPolyline
    
  strStreamID = pNearFeature.value(lStreamIDFldIndex)
  Set pPolyline = pNearFeature.Shape
    
  If FrmTraceDown.nSnapOption = SNAP_NEAREST_NODE Then
    If pNearDis < 0.5 Then
      pNearDis = 0
      Set pNearPt = pPolyline.FromPoint
   Else
      pNearDis = 1
      Set pNearPt = pPolyline.ToPoint
    End If
  End If
  
  If FrmTraceDown.nTraceOption = TRACE_DOWN Then
    Set pBufferStrip = TraceBufferStrip(pFCStream, pNearFeature, pNearDis, FrmTraceDown.fTraceDistance, FrmTraceDown.bTraceToEnd)
  Else
    Set pBufferStrip = TraceToInStreamBMP(pFCStream, pNearFeature, pNearDis, FrmTraceDown.strTraceBMPType)
  End If
  
  If pBufferStrip Is Nothing Then
    MsgBox "Failed to trace Bufferstrip.", vbCritical
    Exit Sub
  End If
  
  Dim lfBufWidth As Double
  lfBufWidth = Format(pBufferStrip.Length, "0.00")
        
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.ActiveView
  
  Dim pTable As iTable
  Set pTable = GetInputDataTable("VFSDefaults")
  
  Dim pSelectedVFSId As Integer
  pSelectedVFSId = FrmTraceDown.cmbVFSTypes.ItemData(FrmTraceDown.bVFSDefaultID)
    
    '** Get the bank of the buffer - Left/Right
  Dim pBufferBank As String
  pBufferBank = FrmTraceDown.bBankSide
  
  '** Get VFS Name
  Dim pBufferName As String
  pBufferName = FrmTraceDown.pVFSName
  
  '** Close the FrmTraceDown form
  Unload FrmTraceDown
  
  '** Get the buffer length
  Dim pBufferLength As Double
  If (Not pTable Is Nothing) Then
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pSelectedVFSId & " AND PropName = 'BufferLength'"
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    If (Not pRow Is Nothing) Then
        pBufferLength = CDbl(pRow.value(pTable.FindField("PropValue")))
    End If
  End If
  Set pRow = Nothing
  Set pCursor = Nothing
  Set pQueryFilter = Nothing
  Set pTable = Nothing
  
  '** If user prompts to redefine parameters, do so
  Dim pVFSDictionary As Scripting.Dictionary
  Set pVFSDictionary = GetVFSProperties("VFSDefaults", CStr(pSelectedVFSId))
  
  If (FrmTraceDown.bRedefineParams = True) Then
'    FrmVFSData.txtName.Text = pBufferName
'    FrmVFSData.txtName.Enabled = False
'    FrmVFSData.txtBufferLength.Text = pBufferLength
'    FrmVFSData.txtBufferWidth.Text = lfBufWidth
'    FrmVFSData.txtBufferWidth.Enabled = False
'    FrmVFSData.Show vbModal
'    '** Get the entered buffer length value
'    pBufferLength = CDbl(FrmVFSData.txtBufferLength.Text)

    InitializeVFSPropertyForm pVFSDictionary
    FrmVFSParams.BufferWidth.Text = lfBufWidth
    FrmVFSParams.Show vbModal
 Else
    Set gBufferStripDetailDict = pVFSDictionary
    gBufferStripDetailDict.Item("BufferWidth") = lfBufWidth
 End If
  
  
  'Get total number of BMPs on map
  Dim pBMPFLayer As IFeatureLayer
  Set pBMPFLayer = GetInputFeatureLayer("BMPs")
  Dim pBMPFClass As IFeatureClass
  Dim pBMPCount As Integer
  pBMPCount = 0
  If Not (pBMPFLayer Is Nothing) Then
        Set pBMPFClass = pBMPFLayer.FeatureClass
        pBMPCount = pBMPFClass.FeatureCount(Nothing)    ' Get BMP feature count
  End If
    
  'Get total number of Conduits on map
  Dim pConduitsFLayer As IFeatureLayer
  Set pConduitsFLayer = GetInputFeatureLayer("Conduits")
  Dim pConduitsFClass As IFeatureClass
  Dim pConduitsCount As Integer
  pConduitsCount = 0
  If Not (pConduitsFLayer Is Nothing) Then
        Set pConduitsFClass = pConduitsFLayer.FeatureClass
        pConduitsCount = pConduitsFClass.FeatureCount(Nothing)    ' Get Conduits feature count
  End If
  
  Dim pLayerVFS As IFeatureLayer
  Set pLayerVFS = GetInputFeatureLayer("VFS")
  Dim pFCVFS As IFeatureClass
  If pLayerVFS Is Nothing Then
    '** Create a VFS layer
    Set pFCVFS = CreateFeatureClassForBMPOrVFS(gMapTempFolder, "vfs", "Polyline")
    Set pLayerVFS = New FeatureLayer
    Set pLayerVFS.FeatureClass = pFCVFS
    AddLayerToMap pLayerVFS, "VFS"
    Set pLayerVFS = GetInputFeatureLayer("VFS")
  End If
    
  Set pFCVFS = pLayerVFS.FeatureClass
  Dim lFldIDIndex As Long, lFldDSIDIndex As Long
  Dim lFldTypeIndex As Long, lFldType2Index As Long, lFldLabelIndex As Long
  
  lFldIDIndex = pFCVFS.FindField("ID")
  lFldDSIDIndex = pFCVFS.FindField("DSID")
  lFldTypeIndex = pFCVFS.FindField("TYPE")
  lFldType2Index = pFCVFS.FindField("TYPE2")
  lFldLabelIndex = pFCVFS.FindField("LABEL")
       
  Dim pVFSCount As Integer
  pVFSCount = pFCVFS.FeatureCount(Nothing) + 1
      
  Dim pVFSID As Integer
  pVFSID = pBMPCount + pConduitsCount + pVFSCount
    
  '** Make the pbuffer strip length short
  Dim pVFSPolyline As IPolyline
  Set pVFSPolyline = pBufferStrip
  Dim pVFSLen As Double
  pVFSLen = pBufferStrip.Length
  pBufferStrip.GetSubcurve 0.005 * pVFSLen, 0.995 * pVFSLen, False, pVFSPolyline
   
  Dim pFeatureBuffer As IFeatureBuffer
  Set pFeatureBuffer = pFCVFS.CreateFeatureBuffer
  Set pFeatureBuffer.Shape = pVFSPolyline
  pFeatureBuffer.value(lFldIDIndex) = pVFSID
  pFeatureBuffer.value(lFldDSIDIndex) = 0
  pFeatureBuffer.value(lFldTypeIndex) = "VFS"
  pFeatureBuffer.value(lFldType2Index) = "VFS_" & Left(pBufferBank, 1) 'TYPE2
  pFeatureBuffer.value(lFldLabelIndex) = "VFS" & CStr(pVFSCount)
  
  Dim pFeatureCursor As IFeatureCursor
  Set pFeatureCursor = pFCVFS.Insert(False)
  pFeatureCursor.InsertFeature pFeatureBuffer
  pFeatureCursor.Flush
  pActiveView.PartialRefresh esriViewGeography, Nothing, Nothing
    
  '** Add details in VFSDetail table
'  Dim pVFSDetailDict As Scripting.Dictionary
'  Set pVFSDetailDict = CreateObject("Scripting.Dictionary")
'  pVFSDetailDict.Add "ID", pVFSID
'  pVFSDetailDict.Add "Name", pBufferName
'  pVFSDetailDict.Add "Length", pBufferLength
'  pVFSDetailDict.Add "Width", lfBufWidth
'  pVFSDetailDict.Add "Bank", pBufferBank
  Dim pVFSDetailDict As Scripting.Dictionary
  Set pVFSDetailDict = gBufferStripDetailDict
  pVFSDetailDict.Item("ID") = pVFSID
'  pVFSDetailDict.Add "Name", pBufferName
'  pVFSDetailDict.Add "Length", pBufferLength
'  pVFSDetailDict.Add "Width", lfBufWidth
  pVFSDetailDict.add "Bank", pBufferBank
  
  '11/20/2008  Ying Cao: Add user selection on the form
  'updated 11/21/2008 FrmTraceDown parameters are not editable for now
'  pVFSDetailDict.Add "SnapOptionIndex", FrmTraceDown.nSnapOption
'  pVFSDetailDict.Add "SnapOptionStr", FrmTraceDown.strSnapBMPType
'  pVFSDetailDict.Add "TraceOptionIndex", FrmTraceDown.nTraceOption
'  pVFSDetailDict.Add "TraceOptionStr", FrmTraceDown.strTraceBMPType
'  pVFSDetailDict.Add "VFSTypeIndex", FrmTraceDown.cmbVFSTypes.ListIndex
  
  '** Close the FrmTraceDown form
  Unload FrmTraceDown
  
  '** Call function to add these values
  SaveVFSPropertiesTable "VFSDetail", CStr(pVFSID), pVFSDetailDict
    
  '** Render the feature layer
  RenderVFSFeatureLayer
  
  
  'Add the VFS related info to BMPNetwork table
  AddBMPNetworkInformation pVFSID, False, False
  
  GoTo CleanUp

ErrorHandler:
    MsgBox "Error in AddVFS module: " & Err.description
    Resume
    
CleanUp:
  Set pLayerStream = Nothing
  Set pFCStream = Nothing
  Set pNearFeature = Nothing
  Set pNearPt = Nothing
  Set pPolyline = Nothing
  Set pBufferStrip = Nothing
  Set pActiveView = Nothing
  Set pTable = Nothing
  Set pQueryFilter = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pBMPFLayer = Nothing
  Set pBMPFClass = Nothing
  Set pConduitsFLayer = Nothing
  Set pConduitsFClass = Nothing
  Set pLayerVFS = Nothing
  Set pFCVFS = Nothing
  Set pFeatureBuffer = Nothing
  Set pFeatureCursor = Nothing
  Set pVFSDetailDict = Nothing
  

End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub

