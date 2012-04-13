VERSION 5.00
Begin VB.Form FrmVFSTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select VFS Template"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "FrmVFSTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   480
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   960
   End
   Begin VB.ComboBox cmbExistVFS 
      Height          =   315
      ItemData        =   "FrmVFSTypes.frx":08CA
      Left            =   2160
      List            =   "FrmVFSTypes.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   960
   End
   Begin VB.CommandButton cmdRedefine 
      Caption         =   "Redefine"
      Height          =   480
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Select an Existing VFS"
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "FrmVFSTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    '** Get the total count of entries in the drop down
    '** Add 1 to it and populate
    Dim pTotalVFS
    pTotalVFS = cmbExistVFS.ListCount
    
    '** Close the form
    Unload Me
    
    '** open the vfsdata form
'    FrmVFSData.txtVFSID.Text = CStr(pTotalVFS + 1)
'    FrmVFSData.txtName.Text = "VFS" & CStr(pTotalVFS + 1)
'    FrmVFSData.Show vbModal

    Dim pVFSDictionary As Scripting.Dictionary
    
    Set pVFSDictionary = GetDefaultsForVFS(pTotalVFS + 1, "VFS" & CStr(pTotalVFS + 1))

    InitializeVFSPropertyForm pVFSDictionary
    FrmVFSParams.Show vbModal
    
    If (FrmVFSParams.bContinue = True) Then
        Dim pIDValue As Integer
        pIDValue = FrmVFSParams.txtVFSID.Text
        
        '** call the generic function to create and add rows for values
        ModuleVFSFunctions.SaveVFSPropertiesTable "VFSDefaults", CStr(pIDValue), gBufferStripDetailDict
            
        '** set it to nothing
        Set gBufferStripDetailDict = Nothing
        Unload FrmVFSParams
    End If

End Sub

Private Sub cmdRedefine_Click()
    Dim pVFSTemplateID As Integer
    Dim pVFSTemplateName As String
    pVFSTemplateName = cmbExistVFS.Text
    
    '** Close the form
    Unload Me
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If (pTable Is Nothing) Then
        MsgBox "VFSDefaults table not defined."
        Exit Sub
    End If

    Dim iIDFld As Long
    iIDFld = pTable.FindField("ID")
    Dim iNameFld As Long
    iNameFld = pTable.FindField("PropName")
    Dim iValueFld As Long
    iValueFld = pTable.FindField("PropValue")

    '** Get the ID value of that VFS
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropValue = '" & pVFSTemplateName & "'"
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    If Not pRow Is Nothing Then
        pVFSTemplateID = pRow.value(iIDFld)
    End If
    Set pCursor = Nothing
    Set pRow = Nothing
    
    '** Iterate over all values of that VFS
''    pQueryFilter.WhereClause = "ID = " & pVFSTemplateID
''    Set pCursor = pTable.Search(pQueryFilter, True)
''    Set pRow = pCursor.NextRow
''
''    Dim pIDValue As Integer
''    Dim pVFSName As String
''    Dim pVFSDictionary As Scripting.Dictionary
''    Set pVFSDictionary = CreateObject("Scripting.Dictionary")
''    Dim pName, pValue As String
''    Do While Not pRow Is Nothing
''        pName = pRow.value(iNameFld)
''        pValue = pRow.value(iValueFld)
''        pVFSDictionary.Add pName, pValue
''        Set pRow = pCursor.NextRow
''    Loop
''
''    '** Load the values in the frmvfsdata with these values
''    If (pVFSDictionary.Count > 0) Then
''        FrmVFSData.txtVFSID.Text = pVFSTemplateID
''        FrmVFSData.txtName.Text = pVFSDictionary.Item("Name")
''        FrmVFSData.txtBufferLength.Text = pVFSDictionary.Item("BufferLength")
''        FrmVFSData.txtBufferWidth.Text = pVFSDictionary.Item("BufferWidth")
''    End If
''
''    '** cleanup
''    Set pRow = Nothing
''    Set pCursor = Nothing
''    Set pQueryFilter = Nothing
''    Set pTable = Nothing
''    Set pVFSDictionary = Nothing
''
''    '** open the vfs form
''    FrmVFSData.Show vbModal
    
    'Changes on April 24, 2007 - Sabu Paul
    Dim pVFSDictionary As Scripting.Dictionary
    Set pVFSDictionary = GetVFSProperties("VFSDefaults", CStr(pVFSTemplateID))
    InitializeVFSPropertyForm pVFSDictionary
    FrmVFSParams.Show vbModal
    
    If (FrmVFSParams.bContinue = True) Then
        Dim pIDValue As Integer
        pIDValue = FrmVFSParams.txtVFSID.Text
        
        '** call the generic function to create and add rows for values
        ModuleVFSFunctions.SaveVFSPropertiesTable "VFSDefaults", CStr(pIDValue), gBufferStripDetailDict
            
        '** set it to nothing
        Set gBufferStripDetailDict = Nothing
        Unload FrmVFSParams
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If (pTable Is Nothing) Then
        MsgBox "VFSDefaults table not defined."
        Exit Sub
    End If

    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'Name'"
        
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim iIDFld As Long
    iIDFld = pTable.FindField("ID")
    Dim iValueFld As Long
    iValueFld = pTable.FindField("PropValue")
    Dim pIDValue As Integer
    Dim pVFSName As String
    Do While Not pRow Is Nothing
        pIDValue = CInt(pRow.value(iIDFld))
        pVFSName = pRow.value(iValueFld)
        '** Load values in combo box
        FrmVFSTypes.cmbExistVFS.AddItem pVFSName
        Set pRow = pCursor.NextRow
    Loop
    FrmVFSTypes.cmbExistVFS.ListIndex = 0
    
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pTable = Nothing
    
End Sub
