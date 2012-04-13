VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddBMPNetwork"
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

Private m_pEditor As IEditor
Private m_pEditSketch As IEditSketch
Private m_Bitmap As IPictureDisp
Private m_bInUse As Boolean
Private m_pPoint As IPoint
Private m_pFeedback As IDisplayFeedback

Private pOutletType1 As Integer
Private pOutletType2 As Integer
Private FromPointID As Integer
Private ToPointID As Integer

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("ROUTE", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_pEditor = Nothing
    Set m_pEditSketch = Nothing
    Set m_Bitmap = Nothing
    Set m_pPoint = Nothing
    Set m_pFeedback = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Routing Network"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableBMPNetworkRouting
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Routing Network"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Define Routing Network"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Routing Network"
End Property


Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("DIGITIZE", vbResCursor)
End Property

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler


  Dim pPoint As IPoint
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  Set pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Dim pLineFeed As INewLineFeedback
  Dim pLayer As ILayer

  If (Not m_bInUse) Then
        m_bInUse = True
        Set m_pFeedback = New NewLineFeedback
        Set pLineFeed = m_pFeedback
        If (Not m_pFeedback Is Nothing) Then
            'Get the color and type of route
            pOutletType1 = GetBMPRouteType(pPoint, FromPointID, "BMPs", True)
            
            If (pOutletType1 > 1) Then     'If its Weir, Orifice,Underdrain
                'If it returns 2, disable the underdrain option
                FrmOutlet.UnderDrain.Enabled = True
                If (pOutletType1 = 2) Then
                    FrmOutlet.UnderDrain.Enabled = False
                End If
                FrmOutlet.Show vbModal
                pOutletType1 = gBMPOutletType
            End If
            
            '*****************************************************
            Dim pSLnSym As ISimpleLineSymbol
            Set pSLnSym = m_pFeedback.Symbol
            ' Create a new RGBColor and set it up
            Dim pRGB As IColor
            Set pRGB = New RgbColor
            Select Case pOutletType1
                Case 1:
                    pRGB.RGB = RGB(0, 0, 0)   'TOTAL
                Case 2:
                    pRGB.RGB = RGB(255, 0, 0) 'WEIR
                Case 3:
                    pRGB.RGB = RGB(0, 255, 0) 'ORIFICE
                Case 4:
                    pRGB.RGB = RGB(0, 0, 255) 'UNDERDRAIN
                Case Else:
                    pRGB.RGB = RGB(0, 0, 0)   'DEFAULT
            End Select

            ' Set the Color and Style for the Feedback's Symbol
            pSLnSym.Color = pRGB
            pSLnSym.Style = esriSLSSolid
            '*****************************************************
            Set m_pFeedback.Display = pActiveView.ScreenDisplay
            pLineFeed.Start pPoint

        End If
  Else

      Set pLineFeed = m_pFeedback
      pOutletType2 = GetBMPRouteType(pPoint, ToPointID, "BMPs", False)
       
       'Add the shifted point
      pLineFeed.AddPoint pPoint
      m_bInUse = False
      'Got a line with two points, stop it get the line
      Dim pPolyline As IPolyline
      Set pPolyline = pLineFeed.Stop
      Dim pGraphicsContainer As IGraphicsContainer
      Set pGraphicsContainer = gMap

      Dim pElement As IElement
      Set pElement = New LineElement
      pElement.Geometry = pPolyline
      Dim pLineElement As ILineElement
      Set pLineElement = pElement
      Dim pLineSymbol As ILineSymbol
      Set pLineSymbol = ReturnBMPtoBMPRouteSymbol(pOutletType1) 'Total = 1
      pLineElement.Symbol = pLineSymbol
      pGraphicsContainer.AddElement pElement, 0
      gMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

      Dim boolChange
      If (pOutletType1 > 0 And pOutletType2 > 0 And (FromPointID <> ToPointID)) Then
        boolChange = MsgBox("Do you want to save this route ?", vbYesNo, "Routing Network")
        If (boolChange = vbYes) Then
            'MsgBox FromPointID & " : " & ToPointID
            'Call subroutine to find the start and end points of line on subwatershed, bmps
            UpdateBMPNetworkTableRoute FromPointID, ToPointID, pOutletType1, "Conduits", "BMPs"
            UpdateBMPFeatureClassInformation FromPointID, ToPointID
            ModifySingleRoute "Conduits", "BMPs", FromPointID, ToPointID, pOutletType1
          End If
      End If
      pGraphicsContainer.DeleteAllElements
      gMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
  End If


  GoTo CleanUp
ErrorHandler:
  MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp:  ' Sabu Paul, Jan 17, 2005 -- Cleanup section was added
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pPolyline = Nothing
    Set pGraphicsContainer = Nothing
    Set pElement = Nothing
    Set pLineElement = Nothing
    Set pLineSymbol = Nothing
    Set pLineFeed = Nothing
    Set pSLnSym = Nothing
    Set pSLnSym = Nothing
    Set pRGB = Nothing
End Sub


Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrorHandler

  If (Not m_bInUse) Then Exit Sub
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  m_pFeedback.MoveTo pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Set m_pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

  GoTo CleanUp
ErrorHandler:
  MsgBox "ITool_OnMouseMove: " & Err.description
CleanUp:
    Set pActiveView = Nothing
End Sub


Private Sub ITool_OnDblClick()

End Sub


Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
End Function

Private Function ITool_OnContextMenu(ByVal X As Long, ByVal Y As Long) As Boolean

End Function


Private Sub ITool_OnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub

Private Sub ITool_OnKeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

End Sub


Private Sub ITool_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub ITool_Refresh(ByVal hDC As esriSystem.OLE_HANDLE)

End Sub




