VERSION 5.00
Begin VB.Form frmAggTypes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Aggregate BMP Type"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   Icon            =   "frmAggTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add"
      Height          =   480
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   960
   End
   Begin VB.ComboBox cmbBMPs 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   645
      Width           =   2400
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   480
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   960
   End
   Begin VB.ComboBox cmbExistBMPs 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2400
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   480
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   960
   End
   Begin VB.CommandButton cmdRedefine 
      Caption         =   "Redefine"
      Height          =   480
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Select BMP"
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   690
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Select Aggregate"
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   285
      Width           =   1320
   End
End
Attribute VB_Name = "frmAggTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public pBmpIdDictionary As Scripting.Dictionary
Public pBmpCatDictionary As Scripting.Dictionary
Public pBmpTypeDictionary As Scripting.Dictionary
Public m_BmpName As String
Public m_iTab As Integer


Private Sub cmbExistBMPs_Click()
    
    cmbBMPs.Clear
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    Set pBmpCatDictionary = CreateObject("Scripting.Dictionary")
    Set pBmpTypeDictionary = CreateObject("Scripting.Dictionary")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    pQueryFilter.WhereClause = "PropName='Type' And PropValue='" & cmbExistBMPs.Text & "'"
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow

    Dim pQueryFilter2 As IQueryFilter
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Do While Not pRow Is Nothing
        Set pQueryFilter2 = New QueryFilter
        pQueryFilter2.WhereClause = "PropName='BMPName' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then cmbBMPs.AddItem pRow2.value(pNameIndex): pBmpIdDictionary.add pRow2.value(pNameIndex), pRow2.value(pIDindex)
        Set pCursor2 = Nothing
        'Set the category dictionary
        pQueryFilter2.WhereClause = "PropName='Category' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then pBmpCatDictionary.add pRow2.value(pIDindex), pRow2.value(pNameIndex)
        'BMP Type dictionary
        pQueryFilter2.WhereClause = "PropName='BMPType' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then pBmpTypeDictionary.add pRow2.value(pIDindex), pRow2.value(pNameIndex)
        
        Set pRow = pCursor.NextRow
    Loop
    
    cmbBMPs.ListIndex = 0
    
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow2 = Nothing
    Set pQueryFilter2 = Nothing
    Set pCursor2 = Nothing

End Sub



Private Sub cmdNew_Click(Index As Integer)
    
    ' Initialize the Dict.....
    Set gBMPPlacedDict = CreateObject("Scripting.Dictionary")
    gNewBMPName = cmbExistBMPs.Text
    gNewBMPType = "Aggregate"
    
    If Index = 0 Then
        gBMPEditMode = False
        '***Add code to create a new BMP of the selected type
        Call DefineBMP(gNewBMPType)
    ElseIf Index = 1 Then
        Call DefineBMP(gNewBMPType, gNewBMPName)
    End If
    
End Sub

Private Sub cmdRedefine_Click()
    '***Add code to modify the parameters for an existing BMP
    Dim pBMPName As String
    'pBmpName = cmbExistBMPs.Value
    pBMPName = cmbBMPs.Text
    Dim pBMPType As String
    pBMPType = cmbExistBMPs.Text
    
    'MsgBox pBmpName & " From dialog"
    Dim pBMPID As Integer
    pBMPID = pBmpIdDictionary.Item(pBMPName)
    Dim pBMPCat As String
    pBMPCat = pBmpCatDictionary.Item(pBMPID)
    
    Dim pBMPIndType As String
    pBMPIndType = pBmpTypeDictionary.Item(pBMPID)
    ' *******************************
    ' Now set the parametes to the Form.....
    ' *******************************
    gNewBMPId = pBMPID
    gNewBMPName = pBMPName
    gBMPEditMode = True

    frmAggBMPDef.Form_Initialize
    frmAggBMPDef.cmbBMPCategory.Enabled = False
    frmAggBMPDef.cmbBmpType.Enabled = False
    frmAggBMPDef.BMPNameA.Enabled = False
    frmAggBMPDef.BMPType.Enabled = False
    frmAggBMPDef.TabBMPType.Tab = 3
    frmAggBMPDef.cmbBMPCategory.Text = pBMPCat
    frmAggBMPDef.cmbBmpType.Text = Get_BMP_Name(pBMPIndType)
    'frmAggBMPDef.cmbBmpType.Text = pBMPIndType
    frmAggBMPDef.BMPNameA.Text = pBMPName
    frmAggBMPDef.BMPType.Text = pBMPType
    frmAggBMPDef.Show vbModal
        
End Sub

Private Sub cmdCancel_Click()
    bContinue = vbNo
    Unload Me
End Sub


Private Sub DefineBMP(BMPType As String, Optional BMPName As String)
    
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
        
    Dim pNewID As Integer
    Dim pNewName As String
    pNewID = pBMPTypesTable.RowCount(Nothing) + 1
    
    If BMPName = "" Then
     
       'Set the name of the BMP to bmpType
       Dim pNameCount As Long
       pNameCount = cmbExistBMPs.ListCount + 1
       pNewName = BMPType & pNameCount
          
       gNewBMPId = pNewID
       gNewBMPType = BMPType
       gNewBMPName = pNewName
       
    Else
        Dim iCnt As Integer
        For iCnt = 0 To cmbBMPs.ListCount - 1
            'gBMPPlacedDict.add pBmpCatDictionary.Item(pBmpIdDictionary.Item(cmbBMPs.List(iCnt))), pBmpCatDictionary.Item(pBmpIdDictionary.Item(cmbBMPs.List(iCnt)))
            gBMPPlacedDict.add pBmpTypeDictionary.Item(pBmpIdDictionary.Item(cmbBMPs.List(iCnt))), pBmpCatDictionary.Item(pBmpIdDictionary.Item(cmbBMPs.List(iCnt)))
        Next
    End If
    
    Unload Me
    frmAggBMPDef.Form_Initialize
    frmAggBMPDef.TabBMPType.Tab = 3
    frmAggBMPDef.BMPType.Text = gNewBMPName
    frmAggBMPDef.Show vbModal

End Sub



Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    gBMPTypeTag = "BMPTemplate"
    
    Dim curBMPType As String
    curBMPType = gNewBMPType
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
    
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "PropName='Type' And PropValue LIKE '%" & curBMPType & "%' ORDER BY PropValue"
    Dim pCursor As ICursor
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    
    Dim pRow As iRow
    cmbExistBMPs.Clear
    
    If pSelRowCount > 0 Then
        Do
            Set pRow = pCursor.NextRow
            If Not (pRow Is Nothing) Then
                If Not pBmpIdDictionary.Exists(pRow.value(pNameIndex)) Then
                    cmbExistBMPs.AddItem pRow.value(pNameIndex)
                    pBmpIdDictionary.add pRow.value(pNameIndex), pRow.value(pIDindex)
                End If
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
