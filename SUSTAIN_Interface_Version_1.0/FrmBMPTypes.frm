VERSION 5.00
Begin VB.Form FrmBMPTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select BMP Type"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBMPTypes.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRedefine 
      Caption         =   "Redefine"
      Height          =   480
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   960
   End
   Begin VB.ComboBox cmbExistBMPs 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2040
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   480
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Select an Existing BMP"
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "FrmBMPTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public pBmpIdDictionary As Scripting.Dictionary
Public m_BmpName As String
Public m_iTab As Integer



Private Sub cmdNew_Click()
    '***Add code to create a new BMP of the selected type
    Unload Me
    Call DefineBMP(gNewBMPType)
End Sub

Private Sub cmdRedefine_Click()
    '***Add code to modify the parameters for an existing BMP
    Dim pBMPName As String
    'pBmpName = cmbExistBMPs.Value
    pBMPName = cmbExistBMPs.Text
    Unload Me
    'MsgBox pBmpName & " From dialog"
    Dim pBMPID As Integer
    pBMPID = pBmpIdDictionary.Item(pBMPName)
    
    gNewBMPId = pBMPID
    gNewBMPName = pBMPName
    gBMPEditMode = True
    Load frmBMPDef
    frmBMPDef.EditType.Text = ""
    frmBMPDef.cmbBMPCategory.Enabled = False
    frmBMPDef.cmbBmpType.Enabled = False
    frmBMPDef.BMPNameA.Enabled = False
    frmBMPDef.TabBMPType.Enabled = False
    frmBMPDef.Form_Initialize
    frmBMPDef.TabBMPType.Tab = Get_Tab_Index(gBMPTypeDict.Item(m_BmpName))
    frmBMPDef.cmbBMPCategory.ListIndex = m_iTab
    frmBMPDef.Update_BMP_Types
    frmBMPDef.cmbBmpType.Text = m_BmpName
    frmBMPDef.BMPNameA.Text = gNewBMPName
    frmBMPDef.Show vbModal
        
End Sub

Private Sub cmdCancel_Click()
    bContinue = vbNo
    Unload Me
End Sub


Public Sub Form_Initialize()
    
    Dim curBMPType As String
    curBMPType = gNewBMPType
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("Name")
    
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "Type = '" & curBMPType & "'"
    Dim pCursor As ICursor
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    
    Dim pRow As iRow
    Dim pBmpIdCount As Integer
    pBmpIdCount = 0
    
    cmbExistBMPs.Clear
    
    If pSelRowCount > 0 Then
        Do
            Set pRow = pCursor.NextRow
            If Not (pRow Is Nothing) Then
                cmbExistBMPs.AddItem pRow.value(pNameIndex), pBmpIdCount
                pBmpIdCount = pBmpIdCount + 1
                pBmpIdDictionary.add pRow.value(pNameIndex), pRow.value(pIDindex)
            End If
        Loop Until (pRow Is Nothing)
        cmbExistBMPs.ListIndex = 0
    Else
        MsgBox "no bmp of this type"
    End If
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
End Sub


Private Sub DefineBMP(BMPType As String)
    
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
        
    Dim pNewID As Integer
    Dim pNewName As String
    
    pNewID = pBMPTypesTable.RowCount(Nothing) + 1
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "Type = '" & BMPType & "'"
 
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
  
    'Set the name of the BMP to bmpType
    Dim pNameCount As Long
    pNameCount = pSelRowCount + 1
    pNewName = BMPType & pNameCount
    
    gBMPEditMode = False
    gNewBMPId = pNewID
    gNewBMPType = BMPType
    gNewBMPName = pNewName
    
    Load frmBMPDef
    frmBMPDef.EditType.Text = ""
    frmBMPDef.Form_Initialize
    frmBMPDef.TabBMPType.Tab = Get_Tab_Index(gBMPTypeDict.Item(m_BmpName))
    frmBMPDef.cmbBMPCategory.ListIndex = m_iTab
    frmBMPDef.Update_BMP_Types
    frmBMPDef.cmbBmpType.Text = m_BmpName
    frmBMPDef.BMPNameA.Text = gNewBMPName
    frmBMPDef.Show vbModal


End Sub



Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
