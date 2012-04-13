VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CreateInputFile"
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
'                Modified: 03/29/2005 - Mira Chokshi added comments to project
'
'******************************************************************************
Option Explicit
Implements ICommand

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Create Input File"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableCreateInputFile
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Create Input File"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "Create Input File"
End Property

Private Sub ICommand_OnClick()
    
    Call InitializeMapDocument
    Call InitializeOperators(True) ' Set the operators based on landuse grid
    If (ModuleUtility.CheckSpatialAnalystLicense = False) Then
        MsgBox "Spatial Analyst Not Available."
        Exit Sub
    End If
    
    'Call the subroutine to check Watershed/SubWatershed layer to continue.
    Dim boolWatershed As Boolean
    boolWatershed = FindAndConvertWatershedFeatureLayerToRaster()
    'boolWatershed = True
    'If found, continue writing input file
    If (boolWatershed = True) Then
        Dim aggBmpCheck As String
        aggBmpCheck = Check_Completeness_Agg_Bmp_Lu
        If aggBmpCheck <> "Success" Then
            MsgBox aggBmpCheck, vbExclamation
        Else
            ModuleFile.WriteInputTextFile
        End If
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Create Input File"
End Property


