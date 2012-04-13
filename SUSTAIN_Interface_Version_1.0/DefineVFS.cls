VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefineVFS"
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
'                Modified: 03/29/2005 - mira chokshi added comments to project
'
'******************************************************************************

Option Explicit
Implements ICommand

Private Property Get ICommand_Bitmap() As esriSystem.OLE_HANDLE

End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Define Buffer Strip Template"
End Property

Private Property Get ICommand_Category() As String
    ICommand_Category = "SUSTAIN"
End Property

Private Property Get ICommand_Checked() As Boolean

End Property

Private Property Get ICommand_Enabled() As Boolean
    ICommand_Enabled = ModuleMenu.EnableDefineBMP
End Property

Private Property Get ICommand_HelpContextID() As Long

End Property

Private Property Get ICommand_HelpFile() As String

End Property

Private Property Get ICommand_Message() As String
    ICommand_Message = "Define Buffer Strip Template"
End Property

Private Property Get ICommand_Name() As String
    ICommand_Name = "DefineBufferStripTemplate"
End Property

Private Sub ICommand_OnClick()
    '** call functions to initialize map operators
    Call InitializeMapDocument
       
    '** Initialize all values
'    FrmVFSData.txtName.Text = ""
'    FrmVFSData.txtName.Enabled = True
'    FrmVFSData.txtBufferLength.Text = ""
'    FrmVFSData.txtName.Enabled = True
'    FrmVFSData.txtBufferWidth.Text = ""
'    FrmVFSData.txtBufferWidth.Enabled = True
    
    FrmVFSParams.txtName.Text = ""
    FrmVFSParams.txtName.Enabled = True
    FrmVFSParams.BufferLength.Text = ""
    FrmVFSParams.txtName.Enabled = True
    FrmVFSParams.BufferWidth.Text = ""
    FrmVFSParams.BufferWidth.Enabled = True
    
    '** Open the VFS Defaults table to get default name
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If (pTable Is Nothing) Then
        '** open the form that defines the buffer strip params
'        FrmVFSData.txtVFSID.Text = 1
'        FrmVFSData.txtName.Text = "VFS1"
'        FrmVFSData.Show vbModal
        
        Dim pVFSDictionary As Scripting.Dictionary
        Set pVFSDictionary = GetDefaultsForVFS(1, "VFS1")
            
        InitializeVFSPropertyForm pVFSDictionary
        FrmVFSParams.Show vbModal
    Else
        FrmVFSTypes.Show vbModal
    End If
    Set pTable = Nothing
       
    
    If (FrmVFSParams.bContinue = True) Then
        Dim pIDValue As Integer
        pIDValue = FrmVFSParams.txtVFSID.Text
        
'        '** create the dictionary
'        Set gBufferStripDetailDict = CreateObject("Scripting.Dictionary")
'        gBufferStripDetailDict.Add "Name", FrmVFSData.txtName.Text
'        gBufferStripDetailDict.Add "BufferLength", FrmVFSData.txtBufferLength.Text
'        gBufferStripDetailDict.Add "BufferWidth", FrmVFSData.txtBufferWidth.Text
        
        '** call the generic function to create and add rows for values
        ModuleVFSFunctions.SaveVFSPropertiesTable "VFSDefaults", CStr(pIDValue), gBufferStripDetailDict
            
        '** set it to nothing
        Set gBufferStripDetailDict = Nothing
        Unload FrmVFSParams
    End If
End Sub

Private Sub ICommand_OnCreate(ByVal hook As Object)
    Set gApplication = hook
End Sub

Private Property Get ICommand_Tooltip() As String
    ICommand_Tooltip = "Define Buffer Strip Template"
End Property

