VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SUSTAINExtension"
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
Implements IExtension
Implements IExtensionConfig

Private m_ExtensionState As esriExtensionState

Private Property Get IExtension_Name() As String
  ' Internal name of the extension
  IExtension_Name = "SUSTAIN Extension"
End Property

Private Sub IExtension_Startup(ByRef initializationData As Variant)
  ' Set the ArcMap application interface
  If (TypeOf initializationData Is IMxApplication) Then
    Set gApplication = initializationData  ' Is ArcMap
  End If
End Sub

Private Sub IExtension_Shutdown()
    CleanUpMemory
  ' Release interface pointers
    Set gApplication = Nothing
End Sub

Private Property Get IExtensionConfig_ProductName() As String
  ' Name in Extension Manager Dialog
  If (gApplication Is Nothing) Then Exit Property
  IExtensionConfig_ProductName = "SUSTAIN Extension"
End Property
 
Private Property Get IExtensionConfig_Description() As String
  ' Description in Extension Manager Dialog
  If (gApplication Is Nothing) Then Exit Property
  
    IExtensionConfig_Description = "SUSTAIN Extension Version 1.0 " & vbCrLf & _
                               "Copywrite/Tetra Tech Inc./April 2005" & vbCrLf & vbCrLf & _
                               "Controls the enabled property of SUSTAIN commands."
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

