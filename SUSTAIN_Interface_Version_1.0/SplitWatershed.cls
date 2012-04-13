VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitWatershed"
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

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("SPLIT", vbResBitmap)
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
    ICommand_Caption = "Split Watershed"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
       ICommand_Enabled = ModuleMenu.EnableDelineation
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Split Watershed"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Split Watershed"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Call StartEditingFeatureLayer("Watershed")
    Call SplitWatershed
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Split Watershed"
End Property


Private Property Get ITool_Cursor() As esriSystem.OLE_HANDLE
    ITool_Cursor = LoadResPicture("DIGITIZE", vbResCursor)
End Property

Private Sub ITool_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError

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
            '*****************************************************
            Dim pSLnSym As ISimpleLineSymbol
            Set pSLnSym = m_pFeedback.Symbol
            ' Create a new RGBColor and set it up
            Dim pRGB As IRgbColor
            Set pRGB = New RgbColor
            With pRGB
               .Red = 255
               .Green = 0
               .Blue = 0
            End With
            ' Set the Color and Style for the Feedback's Symbol
            pSLnSym.Color = pRGB
            pSLnSym.Style = esriSLSSolid
            '*****************************************************
            Set m_pFeedback.Display = pActiveView.ScreenDisplay
            pLineFeed.Start pPoint
        End If
  Else
      Set pLineFeed = m_pFeedback
      pLineFeed.AddPoint pPoint
  End If
 
   
  'Exit Sub ' Sabu Paul Jan 17, 2005
  GoTo CleanUp
ShowError:
  MsgBox "ITool_OnMouseDown: " & Err.description
CleanUp: ' Sabu Paul Jan 17, 2005 -- added CleanUp section
    Set pActiveView = Nothing
    Set pPoint = Nothing
    Set pSLnSym = Nothing
    Set pLineFeed = Nothing
    Set pRGB = Nothing
    Set pLayer = Nothing
End Sub

Private Sub ITool_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long)
On Error GoTo ShowError

  If (Not m_bInUse) Then Exit Sub
  
  Dim pActiveView As IActiveView
  Set pActiveView = gMap
  m_pFeedback.MoveTo pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
  Set m_pPoint = pActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)

  'Exit Sub ' Sabu Paul Jan 17, 2005 -- added CleanUp section
  GoTo CleanUp
ShowError:
  MsgBox "ITool_OnMouseMove: " & Err.description
CleanUp: ' Sabu Paul Jan 17, 2005 -- added CleanUp section
  Set pActiveView = Nothing
End Sub


Private Sub ITool_OnDblClick()
On Error GoTo ShowError

  If (Not m_bInUse) Then Exit Sub
 
  Dim pLineFeed As INewLineFeedback
  Set pLineFeed = m_pFeedback
  
  pLineFeed.AddPoint m_pPoint
  Dim pPolyline As IPolyline
  Set pPolyline = pLineFeed.Stop
  
  Dim boolCreate
  boolCreate = vbNo
  boolCreate = MsgBox("Are you sure you want to split this watershed ?", vbYesNo, "Split Watershed")
    
  Set m_pFeedback = Nothing
  m_bInUse = False
  
  If (boolCreate = vbYes) Then
    SplitWatershedFeature pPolyline
    RenumberWatershedFeatures
    StopEditingFeatureLayer
  End If

  GoTo CleanUp
ShowError:
  MsgBox "ITool_OnDblClick: " & Err.description
CleanUp: ' Sabu Paul Jan 17, 2005 -- added CleanUp section
  Set pLineFeed = Nothing
  Set pPolyline = Nothing
End Sub


Private Function ITool_Deactivate() As Boolean
    ITool_Deactivate = True
    StopEditingFeatureLayer
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




