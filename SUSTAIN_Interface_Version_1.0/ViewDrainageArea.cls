VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViewDrainageArea"
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
'                04/22/2005- mira chokshi added feature to view drainage area in acres
'
'******************************************************************************

Option Explicit
Implements ICommand
Private m_Bitmap As IPictureDisp

Private Sub Class_Initialize()
    Set m_Bitmap = LoadResPicture("AREA", vbResBitmap)
End Sub

Private Sub Class_Terminate()
    Set m_Bitmap = Nothing
End Sub

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE
    ICommand_Bitmap = m_Bitmap
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "View Drainage Area Values"
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
    ICommand_Message = "View Drainage Area Values"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "View Drainage Area Values"
End Property

Private Sub ICommand_OnClick()
    Call InitializeMapDocument
    Call ToggleAreaWatershedView
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "View Drainage Area Values"
End Property


Private Sub ToggleAreaWatershedView()

    Dim pWatershedLayer As IFeatureLayer
    Set pWatershedLayer = GetInputFeatureLayer("Watershed")
    
    '** These properties should be set prior to adding values
    Dim pLyr As IGeoFeatureLayer
    Set pLyr = pWatershedLayer
    Dim pDisplayField As String
    Dim pExpString As String
    If (pLyr.DisplayField = "ID") Then
        pDisplayField = "Area_Acre"
        pExpString = Chr(34) & "Area: " & Chr(34) & "& [" & pDisplayField & "] & " & Chr(34) & "(acres)" & Chr(34)
    Else
        pDisplayField = "ID"
        pExpString = "[" & pDisplayField & "]"
    End If
    pLyr.DisplayField = pDisplayField
     
    
    ' setup LabelEngineProperties for the FeatureLayer
    ' get the AnnotateLayerPropertiesCollection for the FeatureLayer
    Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
    Set pAnnoLayerPropsColl = pLyr.AnnotationProperties
    pLyr.DisplayAnnotation = True
    pAnnoLayerPropsColl.Clear
    ' create a new LabelEngineLayerProperties object
    Dim aLELayerProps As ILabelEngineLayerProperties
    Set aLELayerProps = New LabelEngineLayerProperties
    aLELayerProps.IsExpressionSimple = True
    aLELayerProps.Expression = pExpString
    Dim pTextSymbol As ITextSymbol
    Set pTextSymbol = New TextSymbol
    pTextSymbol.Size = 9
    Dim pColor As IRgbColor
    Set pColor = New RgbColor
    If (pLyr.DisplayField = "ID") Then
        pColor.RGB = RGB(0, 169, 240)
    Else
        pColor.RGB = RGB(0, 0, 0)
    End If
    pTextSymbol.Color = pColor
    Set aLELayerProps.Symbol = pTextSymbol
    ' assign it to the layer's AnnotateLayerPropertiesCollection
    pAnnoLayerPropsColl.add aLELayerProps
    'get the BasicOverposterLayerProperties
    Dim pBasicOverposterLayerProps As IBasicOverposterLayerProperties
    Set pBasicOverposterLayerProps = aLELayerProps.BasicOverposterLayerProperties
    pBasicOverposterLayerProps.NumLabelsOption = esriOneLabelPerShape
          
   '** Refresh the TOC
   gMxDoc.ActiveView.ContentsChanged
   gMxDoc.UpdateContents

   '** Draw the map
   gMxDoc.ActiveView.Refresh
  
  Exit Sub
ErrorHandler:
    MsgBox "ToggleAreaWatershedView: " & Err.description
     
End Sub






