VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bmpToolMove"
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
Private m_bmpID As Integer
Private m_bmpType As String
Private m_BMPFLayer As IFeatureLayer
Private m_MovePoint As IPoint
Private m_pMovePointFeedback As IMoveImageFeedback

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("bmpMove", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
    Set m_BMPFLayer = Nothing
    Set m_MovePoint = Nothing
    Set m_pMovePointFeedback = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Move BMP"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableAddBMPOnLand
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Move BMP"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "MoveBMP"
End Property

Private Sub ICommand_OnClick()
    Set m_BMPFLayer = GetInputFeatureLayer("BMPs")
    If (m_BMPFLayer Is Nothing) Then
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Move BMP"
End Property

Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("EditMove", vbResCursor)
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

  ' Get the ScreenDisplay for the the ActiveView
  Dim pDisp As IDisplay
  Set pDisp = gMxDoc.ActiveView.ScreenDisplay
  
  ' Get the current mouse location in Map Units
  Dim pSearchPoint As IPoint
  Set pSearchPoint = pDisp.DisplayTransformation.ToMapPoint(X, Y)

  ' Check that user is not using an existing feedback
  If m_pMovePointFeedback Is Nothing Then
    ' Search for bmp's
    Set m_MovePoint = SearchSelectedBMP(pSearchPoint)
  End If
  
  If ((m_pMovePointFeedback Is Nothing) And (Not m_MovePoint Is Nothing)) Then
    ' Create a new Feedback
    Set m_pMovePointFeedback = New MoveImageFeedback
    Dim pDI As IDisplayFeedback
    Set pDI = m_pMovePointFeedback
    Set pDI.Display = pDisp
    
    Dim pPictureMarkerSymbol As IMultiLayerMarkerSymbol     ' IPictureMarkerSymbol
    Set pPictureMarkerSymbol = ReturnBMPSymbol(m_bmpType)
       
    Dim pSymbol As ISymbol
    Set pSymbol = pPictureMarkerSymbol
    Dim pBoundary As IPolygon
    Set pBoundary = New Polygon
    
    With m_pMovePointFeedback.Display
        .StartDrawing .hDC, esriNoScreenCache
        'Query boundary
        pSymbol.QueryBoundary .hDC, .DisplayTransformation, m_MovePoint, pBoundary
        .SetSymbol pPictureMarkerSymbol
        .DrawPoint m_MovePoint
        .FinishDrawing
    End With
    m_pMovePointFeedback.Bounds = pBoundary.Envelope
    
    'Start the Feedback at the current mouse location
    m_pMovePointFeedback.Start m_MovePoint
    
  ElseIf (Not m_pMovePointFeedback Is Nothing) Then  ' Otherwise use the current mouse location to add a vertex to the
    ' current feedback
    m_pMovePointFeedback.MoveTo pSearchPoint
    'move the selected point
    MoveSelectedBMP pSearchPoint
    m_bmpID = -1
    'm_pMovePointFeedback.Stop
    Set m_pMovePointFeedback = Nothing
  End If
  
  GoTo CleanUp

ShowError:
    MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp:
    Set pDisp = Nothing
    Set pSearchPoint = Nothing
    Set pDI = Nothing
    Set pPictureMarkerSymbol = Nothing
    Set pSymbol = Nothing
    Set pBoundary = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

  ' Get the ScreenDisplay for the the ActiveView
  Dim pDisp As IDisplay
  Set pDisp = gMxDoc.ActiveView.ScreenDisplay
  
  ' Check if the user is currently using the feedback
  If Not m_pMovePointFeedback Is Nothing Then
    ' Get the current mouse location in map units
    Dim pPnt As IPoint
    Set pPnt = pDisp.DisplayTransformation.ToMapPoint(X, Y)
    
    'Move the Feedback to the current mouse location (IDisplayFeedback::MoveTo)
    m_pMovePointFeedback.MoveTo pPnt
  End If

  Set pDisp = Nothing
  Set pPnt = Nothing
End Sub

Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
    
End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub


'*** Subroutine to search for bmps in a selected radius
Private Function SearchSelectedBMP(pPoint As IPoint) As IPoint
On Error GoTo ShowError

    'Enlarge the point envelope 5 times
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pPoint.Envelope
    ExpandPointEnvelope pEnvelope
    
    'Initialize spatial filter
    Dim pSpatialFilter As ISpatialFilter
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pEnvelope
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = m_BMPFLayer.FeatureClass
    Dim pBMPFCursor As IFeatureCursor
    Set pBMPFCursor = pBMPFClass.Search(pSpatialFilter, True)
    Dim pBMPFeature As IFeature
    Set pBMPFeature = pBMPFCursor.NextFeature
    If Not (pBMPFeature Is Nothing) Then
        m_bmpID = pBMPFeature.value(pBMPFCursor.FindField("ID"))
        m_bmpType = pBMPFeature.value(pBMPFCursor.FindField("TYPE2"))
        Set SearchSelectedBMP = pBMPFeature.Shape
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "SearchSelectedBMP: " & Err.description
CleanUp:
    Set pEnvelope = Nothing
    Set pSpatialFilter = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFCursor = Nothing
    Set pBMPFeature = Nothing
End Function

'*** Move selected bmp to the moved location
Private Sub MoveSelectedBMP(pPoint As IPoint)
On Error GoTo ShowError

    'Enlarge the point envelope 5 times
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pPoint.Envelope
    ExpandPointEnvelope pEnvelope
    
    'Initialize spatial filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & m_bmpID
    
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = m_BMPFLayer.FeatureClass
    Dim pBMPFCursor As IFeatureCursor
    Set pBMPFCursor = pBMPFClass.Search(pQueryFilter, True)
    Dim pBMPFeature As IFeature
    Set pBMPFeature = pBMPFCursor.NextFeature
    If Not (pBMPFeature Is Nothing) Then
        Set pBMPFeature.Shape = pPoint
        pBMPFeature.Store
        'Call function to update conduits feature layer
        MoveConduitsForSelectedBMP pPoint
        'Call function to update basinrouting feature layer
        MoveBasinRoutesForSelectedBMP pPoint
        'Refresh map document
        gMxDoc.ActiveView.Refresh
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "MoveSelectedBMP: " & Err.description
CleanUp:
    Set pEnvelope = Nothing
    Set pQueryFilter = Nothing
    Set pBMPFClass = Nothing
    Set pBMPFCursor = Nothing
    Set pBMPFeature = Nothing
End Sub


'*** Resize/Reshare conduits feature layer for moving bmps
Private Sub MoveConduitsForSelectedBMP(pPoint As IPoint)
On Error GoTo ShowError
    
    Dim pConduitsFLayer As IFeatureLayer
    Set pConduitsFLayer = GetInputFeatureLayer("Conduits")
    If (pConduitsFLayer Is Nothing) Then
        Exit Sub
    End If
    Dim pConduitsFClass As IFeatureClass
    Set pConduitsFClass = pConduitsFLayer.FeatureClass
    
    'Initialize spatial filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "CFROM = " & m_bmpID & " OR CTO = " & m_bmpID
    
    Dim bContinue
    bContinue = vbNo
    If (pConduitsFClass.FeatureCount(pQueryFilter) > 0) Then
        bContinue = MsgBox("Do you want to adjust conduit lengths automatically ?", vbYesNo)
        If (bContinue = vbNo) Then
            MsgBox "Update the conduit length for each modified conduit.", vbExclamation
        End If
    End If
    Dim pBMPDetail As iTable
    Set pBMPDetail = GetInputDataTable("BMPDetail")
    Dim pQueryFilter1 As IQueryFilter
    Set pQueryFilter1 = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iPropValue As Long
    iPropValue = pBMPDetail.FindField("PropValue")
    
    Dim pConduitsFCursor As IFeatureCursor
    Set pConduitsFCursor = pConduitsFClass.Search(pQueryFilter, True)
    Dim pConduitsFeature As IFeature
    Set pConduitsFeature = pConduitsFCursor.NextFeature
    Dim iFROM As Long
    iFROM = pConduitsFCursor.FindField("CFROM")
    Dim iTO As Long
    iTO = pConduitsFCursor.FindField("CTO")
    Dim iID As Long
    iID = pConduitsFCursor.FindField("ID")
    
    Dim pConduitLength As Double
    Dim pPolyline As IPolyline
    Do While Not pConduitsFeature Is Nothing
        Set pPolyline = pConduitsFeature.Shape
        If (pConduitsFeature.value(iFROM) = m_bmpID) Then
            pPolyline.FromPoint = pPoint
        ElseIf (pConduitsFeature.value(iTO) = m_bmpID) Then
            pPolyline.ToPoint = pPoint
        End If
        Set pConduitsFeature.Shape = pPolyline
        pConduitLength = Format(pPolyline.Length * gMetersPerUnit * 3.28, "#.##")   'Get meters per unit, convert meters to feet
        pConduitsFeature.Store
        
        'if bcontinue = vbYes, update the pbmpdetail table
        If (bContinue = vbYes) Then
            pQueryFilter1.WhereClause = "ID = " & pConduitsFeature.value(iID) & " AND PropName = 'LENGTH'"
            Set pCursor = pBMPDetail.Search(pQueryFilter1, True)
            Set pRow = pCursor.NextRow
            If Not (pRow Is Nothing) Then
                pRow.value(iPropValue) = pConduitLength
                pRow.Store
                Set pRow = pCursor.NextRow
            End If
        End If
        
        'Move to next conduit feature
        Set pConduitsFeature = pConduitsFCursor.NextFeature
    Loop
    
    GoTo CleanUp
    
ShowError:
    MsgBox "MoveSelectedBMP: " & Err.description
CleanUp:
    Set pConduitsFLayer = Nothing
    Set pConduitsFClass = Nothing
    Set pQueryFilter = Nothing
    Set pConduitsFCursor = Nothing
    Set pConduitsFeature = Nothing
    Set pPolyline = Nothing
    Set pBMPDetail = Nothing
    Set pQueryFilter1 = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub


'*** Resize/reshape basinrouting features due to moving bmps
Private Sub MoveBasinRoutesForSelectedBMP(pPoint As IPoint)
On Error GoTo ShowError
    
    Dim pBasinRoutesFLayer As IFeatureLayer
    Set pBasinRoutesFLayer = GetInputFeatureLayer("BasinRouting")
    If (pBasinRoutesFLayer Is Nothing) Then
        Exit Sub
    End If
    Dim pBasinRoutesFClass As IFeatureClass
    Set pBasinRoutesFClass = pBasinRoutesFLayer.FeatureClass
    
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If (pWatershedFLayer Is Nothing) Then
        Exit Sub
    End If
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "BMPID = " & m_bmpID
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pWatershedFClass.Search(pQueryFilter, True)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Dim pWaterID As Integer
    
    'Initialize query filter
    Dim pQueryFilter1 As IQueryFilter
    Set pQueryFilter1 = New QueryFilter
    Dim pBasinRoutesFCursor As IFeatureCursor
    Dim pBasinRoutesFeature As IFeature
    Dim iID As Long
    Dim pPolyline As IPolyline
    Do While Not pFeature Is Nothing
        pWaterID = pFeature.value(pFeatureCursor.FindField("ID"))
        pQueryFilter1.WhereClause = "ID = " & pWaterID
        Set pBasinRoutesFCursor = pBasinRoutesFClass.Search(pQueryFilter1, True)
        Set pBasinRoutesFeature = pBasinRoutesFCursor.NextFeature
        iID = pBasinRoutesFCursor.FindField("ID")
        Do While Not pBasinRoutesFeature Is Nothing
             Set pPolyline = pBasinRoutesFeature.Shape
             pPolyline.ToPoint = pPoint
             Set pBasinRoutesFeature.Shape = pPolyline
             pBasinRoutesFeature.Store
             Set pBasinRoutesFeature = pBasinRoutesFCursor.NextFeature
        Loop
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    
    GoTo CleanUp
    
ShowError:
    MsgBox "MoveSelectedBMP: " & Err.description
CleanUp:
    Set pWatershedFLayer = Nothing
    Set pWatershedFClass = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
    Set pBasinRoutesFLayer = Nothing
    Set pQueryFilter = Nothing
    Set pBasinRoutesFClass = Nothing
    Set pBasinRoutesFCursor = Nothing
    Set pBasinRoutesFeature = Nothing
    Set pPolyline = Nothing
    Set pQueryFilter1 = Nothing
End Sub

