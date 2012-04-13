VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BMPExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Application: Sustain - BMP Siting Tool
'   Company:     Tetra Tech, Inc
'   History:     Created:
'******************************************************************************

Option Explicit
Implements IExtension
Implements IExtensionConfig

Private m_ExtensionState As esriExtensionState
Private WithEvents pDocumentEvents As DocumentEvents
Attribute pDocumentEvents.VB_VarHelpID = -1

Private Property Get IExtension_Name() As String
  ' Internal name of the extension
  IExtension_Name = "BMP Extension"
End Property

Private Sub IExtension_Startup(ByRef initializationData As Variant)
  ' Set the ArcMap application interface
  If (TypeOf initializationData Is IMxApplication) Then
    Set gApplication = initializationData  ' Is ArcMap
  End If
End Sub

Private Sub IExtension_Shutdown()
    ' Release forms
    
    
    ' Release interface pointers
    Set pDocumentEvents = Nothing
    CleanUpMemory
    
End Sub

Private Property Get IExtensionConfig_ProductName() As String
  ' Name in Extension Manager Dialog
  If (gApplication Is Nothing) Then Exit Property
  IExtensionConfig_ProductName = "BMP Extension"
End Property
 
Private Property Get IExtensionConfig_Description() As String
  ' Description in Extension Manager Dialog
  If (gApplication Is Nothing) Then Exit Property
  
    IExtensionConfig_Description = "BMP Extension Version 1.0 " & vbCrLf & _
                               "Copywrite/Tetra Tech Inc./November 2006" & vbCrLf & vbCrLf & _
                               "Controls the enabled property of BMP commands."
End Property
 
Private Property Get IExtensionConfig_State() As esriExtensionState
  ' Get the extension state
  If (gApplication Is Nothing) Then Exit Property
  IExtensionConfig_State = m_ExtensionState
End Property
 
Private Property Let IExtensionConfig_State(ByVal ExtensionState As esriExtensionState)
  ' Set the extension state according to the checkbox in the ExtensionManager Dialog
  If (gApplication Is Nothing) Then Exit Property
  m_ExtensionState = ExtensionState
End Property






