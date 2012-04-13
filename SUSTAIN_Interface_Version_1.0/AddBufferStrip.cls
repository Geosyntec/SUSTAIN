VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddBufferStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson
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
    ICommand_Caption = "Buffer Strip"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = True
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Buffer Strip"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "BufferStrip"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Buffer Strip"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("BufferStrip", vbResCursor)
End Property

Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean
  
End Function

Private Sub ITool_OnDblClick()

End Sub

Private Sub ITool_OnKeyDown(ByVal keyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnKeyUp(ByVal keyCode As Long, ByVal Shift As Long)
     
End Sub

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
  On Error GoTo ErrorHandler
  
  Dim pLayerStream As IFeatureLayer
  Set pLayerStream = GetInputFeatureLayer("STREAM")
  If pLayerStream Is Nothing Then
    MsgBox "Layer Stream is missing.", vbExclamation
    Exit Sub
  End If
  
  Dim pFCStream As IFeatureClass
  Set pFCStream = pLayerStream.FeatureClass
  
  Dim strIDName As String, strDSIDName As String
  strIDName = "SUBBASIN"
  strDSIDName = "SUBBASINR"
  
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
    MsgBox "Cannot snap to a stream at the specified location.", vbExclamation
    Exit Sub
  End If
    
  Set pNearFeature = hitTestRes(0)
  Set pNearPt = hitTestRes(1)
  pNearDis = hitTestRes(2)
  
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
  If pBufferStrip Is Nothing Then Exit Sub
  
  Dim lfBufWidth As Double
  lfBufWidth = pBufferStrip.Length
  MsgBox "Buffer Width is " & lfBufWidth
      
''''''''''''''''''''''''''''''''''
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.ActiveView

'  Dim pElement As IElement
'
'  Dim pGraphicsContainer As IGraphicsContainer
'  Set pGraphicsContainer = pActiveView.GraphicsContainer
'
'  Dim pGCSelect As IGraphicsContainerSelect
'  Set pGCSelect = pActiveView
'
'  pGCSelect.SelectAllElements
'  Dim pEnumElement As IEnumElement
'  Set pEnumElement = pGCSelect.SelectedElements
'
'  Set pElement = pEnumElement.Next
'
'  Do Until pElement Is Nothing
'    pGraphicsContainer.DeleteElement pElement
'    Set pElement = pEnumElement.Next
'  Loop
'
'  Set pElement = New LineElement
'  pElement.Geometry = pBufferStrip
'
'  Dim aColor As IRgbColor
'  Set aColor = New RgbColor
'
'  aColor.Red = 255
'  aColor.Green = 135
'  aColor.Blue = 43
'  aColor.Transparency = 50
'
'  Dim pLineSymbol As ISimpleLineSymbol
'  Set pLineSymbol = New SimpleLineSymbol
'  pLineSymbol.Style = esriSLSSolid
'  pLineSymbol.Width = 3
'  pLineSymbol.Color = aColor
'
'  Dim pLineElement As ILineElement
'  Set pLineElement = pElement
'  pLineElement.Symbol = pLineSymbol
'
'  pGraphicsContainer.AddElement pElement, 0
'  pElement.Activate gMxDoc.ActiveView.ScreenDisplay
'  pActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  frmBufferStripTemplate.Show vbModal
  
  Dim pLayerVFS As IFeatureLayer
  Set pLayerVFS = GetInputFeatureLayer("VFS")
  If pLayerVFS Is Nothing Then
    MsgBox "Layer VFS is missing.", vbExclamation
    Exit Sub
  End If
  
  Dim pFCVFS As IFeatureClass
  Set pFCVFS = pLayerVFS.FeatureClass
  
  Dim lFldIDIndex As Long, lFldDSIDIndex As Long
  Dim lFldTypeIndex As Long, lFldType2Index As Long, lFldLabelIndex As Long
  
  lFldIDIndex = pFCVFS.FindField("ID")
  lFldDSIDIndex = pFCVFS.FindField("DSID")
  lFldTypeIndex = pFCVFS.FindField("TYPE1")
  lFldType2Index = pFCVFS.FindField("TYPE2")
  lFldLabelIndex = pFCVFS.FindField("LABEL")
  
  Dim pFeatureBuffer As IFeatureBuffer
  Set pFeatureBuffer = pFCVFS.CreateFeatureBuffer
  Set pFeatureBuffer.Shape = pBufferStrip
  pFeatureBuffer.value(lFldIDIndex) = 2
  pFeatureBuffer.value(lFldDSIDIndex) = 0
  pFeatureBuffer.value(lFldTypeIndex) = "VFS"
  pFeatureBuffer.value(lFldType2Index) = "VFS"
  pFeatureBuffer.value(lFldLabelIndex) = "VFS1"
  
  Dim pFeatureCursor As IFeatureCursor
  Set pFeatureCursor = pFCVFS.Insert(False)
  pFeatureCursor.InsertFeature pFeatureBuffer
  pFeatureCursor.Flush
  pActiveView.PartialRefresh esriViewGeography, Nothing, Nothing

CleanUp:

  Exit Sub

ErrorHandler:
    
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub

