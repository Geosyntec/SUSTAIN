VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefPollutants 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Pollutants"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "frmDefPollutants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameSedAssoc 
      Caption         =   "Sediment Association"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Width           =   3975
      Begin VB.Frame frameSedAssocFrac 
         Caption         =   "Sediment Association Fractions"
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3735
         Begin VB.TextBox txtClayFrac 
            Height          =   285
            Left            =   1320
            TabIndex        =   25
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtSiltFrac 
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtSandFrac 
            Height          =   285
            Left            =   1320
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Clay"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Silt"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Sand"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox chkSedAssoc 
         Caption         =   "Sediment Associated Pollutant?"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox txtTip 
      BackColor       =   &H00D5FFFF&
      Height          =   1005
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2520
      Width           =   2835
   End
   Begin MSComctlLib.ListView lstPollutants 
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pollutant Name"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Multiplier"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sediment Flag"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sediment Associated?"
         Object.Width           =   3176
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sand Fraction"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Silt Fraction"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Clay Fraction"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6405
      TabIndex        =   12
      Top             =   480
      Width           =   750
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   6405
      TabIndex        =   11
      Top             =   1130
      Width           =   750
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   6405
      TabIndex        =   10
      Top             =   1800
      Width           =   750
   End
   Begin VB.Frame framePollutants 
      Caption         =   "Pollutant Properties"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   320
         Width           =   2415
      End
      Begin VB.TextBox txtMultiplier 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Text            =   "1.0"
         Top             =   770
         Width           =   1215
      End
      Begin VB.ComboBox cmbFlag 
         Height          =   315
         ItemData        =   "frmDefPollutants.frx":08CA
         Left            =   1440
         List            =   "frmDefPollutants.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Multiplier"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Sediment Flag"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1260
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   3650
      Width           =   855
   End
   Begin VB.TextBox txtPollutantID 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   3650
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Tips"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Pollutant to View/Edit Properties"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmDefPollutants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmDefPollutants.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
Private m_addFlag As Boolean
Private m_PollutantDict As Scripting.Dictionary


Private Sub chkSedAssoc_Click()
    txtTip.Text = "Select this box if the pollutant is a sediment associated pollutant"
    If chkSedAssoc.value = 0 Then
        frameSedAssocFrac.Enabled = False
    Else
        frameSedAssocFrac.Enabled = True
        If txtSandFrac.Text = "" Then txtSandFrac.Text = 0
        If txtSiltFrac.Text = "" Then txtSiltFrac.Text = 0
        If txtClayFrac.Text = "" Then txtClayFrac.Text = 0
    End If
End Sub



Private Sub cmbFlag_Click()
    If cmbFlag.Text <> "NO" Then
        chkSedAssoc.value = 0
        frameSedAssoc.Enabled = False
    Else
        If Not m_PollutantDict Is Nothing Then
            If m_PollutantDict.Exists("SEDIMENT") Or m_PollutantDict.Exists("SAND") Or m_PollutantDict.Exists("SILT") Or m_PollutantDict.Exists("CLAY") Then
                frameSedAssoc.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmbFlag_GotFocus()
    'txtTip.Text = "Select sediment flag. " & vbNewLine & _
                            "0 - No sediment" & vbTab & _
                            "1 - Sand" & vbNewLine & _
                            "2 - Silt" & vbTab & _
                            "3 - Clay" & vbNewLine & _
                            "4 - SEDIMENT"
End Sub

Private Sub cmdAdd_Click()
    m_addFlag = True
    txtName.Text = ""
    txtMultiplier.Text = "1.0"
    cmbFlag.ListIndex = 0
    
    If lstPollutants.ListItems.Count > 0 Then
        txtPollutantID.Text = lstPollutants.ListItems(lstPollutants.ListItems.Count).Text + 1
    Else
        txtPollutantID.Text = 1
    End If
    framePollutants.Enabled = True
    If m_PollutantDict.Exists("SEDIMENT") Or m_PollutantDict.Exists("SAND") Or m_PollutantDict.Exists("SILT") Or m_PollutantDict.Exists("CLAY") Then
        frameSedAssoc.Enabled = True
    Else
        frameSedAssoc.Enabled = False
    End If
    cmdSave.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    If lstPollutants.ListItems.Count = 0 Then Exit Sub
    m_addFlag = False
    framePollutants.Enabled = True
    cmdSave.Enabled = True
    txtPollutantID = lstPollutants.SelectedItem.Text
    txtName.Text = lstPollutants.SelectedItem.ListSubItems(1).Text
    txtMultiplier.Text = lstPollutants.SelectedItem.ListSubItems(2).Text
    cmbFlag.Text = lstPollutants.SelectedItem.ListSubItems(3).Text
    
    If cmbFlag.Text = "NO" And (m_PollutantDict.Exists("SEDIMENT") Or m_PollutantDict.Exists("SAND") Or m_PollutantDict.Exists("SILT") Or m_PollutantDict.Exists("CLAY")) Then
        frameSedAssoc.Enabled = True
        If CInt(lstPollutants.SelectedItem.ListSubItems(4).Text) = 1 Then
            chkSedAssoc.value = 1
            txtSandFrac.Text = lstPollutants.SelectedItem.ListSubItems(5).Text
            txtSiltFrac.Text = lstPollutants.SelectedItem.ListSubItems(6).Text
            txtClayFrac.Text = lstPollutants.SelectedItem.ListSubItems(7).Text
        End If
    Else
        frameSedAssoc.Enabled = False
    End If
End Sub

Private Sub cmdRemove_Click()
    
    On Error GoTo ErrorHandler
    m_addFlag = False
    
    If lstPollutants.SelectedItem Is Nothing Then Exit Sub
    
    '** Confirm the deletion
    Dim boolDelete
    boolDelete = MsgBox("Are you sure you want to delete this pollutant information ?", vbYesNo)
    If (boolDelete = vbNo) Then
        Exit Sub
    End If
    
    '** get pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstPollutants.SelectedItem.Text
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("Pollutants")
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pPollutantID
    
    '** delete records
    pSWMMPollutantTable.DeleteSearchedRows pQueryFilter
    
    '*** Increment the id's by 1 number for all records after deleted id
    Dim pFromID As Integer
    pFromID = pPollutantID + 1
    Dim bContinue As Boolean
    bContinue = True
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iID As Long
    iID = pSWMMPollutantTable.FindField("ID")
    
    Do While bContinue
        pQueryFilter.WhereClause = "ID = " & pFromID
        Set pCursor = Nothing
        Set pRow = Nothing
        Set pCursor = pSWMMPollutantTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            bContinue = False
        End If
        Do While Not pRow Is Nothing
            pRow.value(iID) = pFromID - 1
            pRow.Store
            Set pRow = pCursor.NextRow
        Loop
        pFromID = pFromID + 1
    Loop
    
    '** clean up
    Set pQueryFilter = Nothing
    Set pSWMMPollutantTable = Nothing
    
    '** load pollutant names
    Call LoadPollutants
    
    Exit Sub
ErrorHandler:
  HandleError True, "cmdRemove_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo ErrorHandler
    If Trim(txtName.Text) = "" Then
        MsgBox "Pollutant name cannot be blank.", vbInformation, "SUSTAIN"
        Exit Sub
    End If
    If Not lstPollutants.FindItem(txtName.Text, 1) Is Nothing And m_addFlag Then
        MsgBox "Pollutant name already exists.", vbInformation, "SUSTAIN"
        Exit Sub
    End If
    If Not m_PollutantDict Is Nothing And lstPollutants.ListItems.Count > 0 Then
        If cmbFlag.Text <> "NO" And (m_PollutantDict.Exists("SEDIMENT") Or m_PollutantDict.Exists(cmbFlag.Text)) And cmbFlag.Text <> lstPollutants.SelectedItem.ListSubItems(3).Text Then
            MsgBox "Cannot add this sediment.", vbInformation, "SUSTAIN"
            Exit Sub
        End If
    End If
    'Not sure why this condition was set - Sabu Paul, Dec 5, 2008
'    If cmbFlag.Text = "SEDIMENT" And m_PollutantDict.Count > 0 Then
'        MsgBox "Cannot add this sediment.", vbInformation, "SUSTAIN"
'        Exit Sub
'    End If
    
    'If it sediment make sure that sand, silt, or clay are not already defined
    If cmbFlag.Text = "SEDIMENT" And (m_PollutantDict.Exists("SAND") Or m_PollutantDict.Exists("SILT") Or m_PollutantDict.Exists("CLAY")) Then
        MsgBox "Cannot add this sediment as at least one of sand, silt, or clay is already in the list.", vbInformation, "SUSTAIN"
        Exit Sub
    End If
    
    If chkSedAssoc.value = 1 Then
        If Not (IsNumeric(txtSandFrac.Text) And IsNumeric(txtSiltFrac.Text) And IsNumeric(txtClayFrac.Text)) Then
            MsgBox "Enter valid numbers for sand, silt, and clay fractions (0-1) ", vbInformation
            Exit Sub
        End If
        If CDbl(txtSandFrac.Text) + CDbl(txtSiltFrac.Text) + CDbl(txtClayFrac.Text) <> 1# Then
            MsgBox "Sum of sand, silt, and clay fractions should be 1.0 ", vbInformation
            Exit Sub
        End If
    End If
    ' Update the Listview........
    If Not m_addFlag Then
        If lstPollutants.ListItems.Count > 0 Then
            lstPollutants.SelectedItem.ListSubItems(1).Text = txtName.Text
            lstPollutants.SelectedItem.ListSubItems(2).Text = txtMultiplier.Text
            lstPollutants.SelectedItem.ListSubItems(3).Text = cmbFlag.Text
            If chkSedAssoc.value = 1 Then
                lstPollutants.SelectedItem.ListSubItems(4).Text = 1
                lstPollutants.SelectedItem.ListSubItems(5).Text = CDbl(txtSandFrac.Text)
                lstPollutants.SelectedItem.ListSubItems(6).Text = CDbl(txtSiltFrac.Text)
                lstPollutants.SelectedItem.ListSubItems(7).Text = CDbl(txtClayFrac.Text)
            Else
                lstPollutants.SelectedItem.ListSubItems(4).Text = 0
                lstPollutants.SelectedItem.ListSubItems(5).Text = 0
                lstPollutants.SelectedItem.ListSubItems(6).Text = 0
                lstPollutants.SelectedItem.ListSubItems(7).Text = 0
            End If
        End If
    End If
    
    'Find the SWMM option table
    Dim pFLag As Boolean
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    
    'Create the table if not found, add it to the Map
    If (pTable Is Nothing) Then
        pFLag = True
        Set pTable = CreatePollutantsTableDBF("Pollutants")
    End If
        
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim i As Integer, pID As Integer
    
    'Iterate over the property dictionary, and save the values
    For i = 1 To lstPollutants.ListItems.Count
        pID = lstPollutants.ListItems(i).Text
        pQueryFilter.WhereClause = "ID = " & pID
        Set pCursor = pTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            Set pRow = pTable.CreateRow
        End If
        pRow.value(pTable.FindField("ID")) = pID
        pRow.value(pTable.FindField("Name")) = lstPollutants.ListItems(i).ListSubItems(1).Text
        pRow.value(pTable.FindField("Multiplier")) = lstPollutants.ListItems(i).ListSubItems(2).Text
        pRow.value(pTable.FindField("Sediment")) = lstPollutants.ListItems(i).ListSubItems(3).Text

        pRow.value(pTable.FindField("SedAssoc")) = lstPollutants.ListItems(i).ListSubItems(4).Text
        pRow.value(pTable.FindField("SandFrac")) = lstPollutants.ListItems(i).ListSubItems(5).Text
        pRow.value(pTable.FindField("SiltFrac")) = lstPollutants.ListItems(i).ListSubItems(6).Text
        pRow.value(pTable.FindField("ClayFrac")) = lstPollutants.ListItems(i).ListSubItems(7).Text
        
        pRow.Store
    Next
    
    ' Add the new record.....
    If m_addFlag Then
        Set pRow = pTable.CreateRow
        pRow.value(pTable.FindField("ID")) = txtPollutantID.Text
        pRow.value(pTable.FindField("Name")) = txtName.Text
        pRow.value(pTable.FindField("Multiplier")) = txtMultiplier.Text
        pRow.value(pTable.FindField("Sediment")) = cmbFlag.Text
        If chkSedAssoc.value = 1 Then
            pRow.value(pTable.FindField("SedAssoc")) = 1
            pRow.value(pTable.FindField("SandFrac")) = CDbl(txtSandFrac.Text)
            pRow.value(pTable.FindField("SiltFrac")) = CDbl(txtSiltFrac.Text)
            pRow.value(pTable.FindField("ClayFrac")) = CDbl(txtClayFrac.Text)
        Else
            pRow.value(pTable.FindField("SedAssoc")) = 0
            pRow.value(pTable.FindField("SandFrac")) = 0
            pRow.value(pTable.FindField("SiltFrac")) = 0
            pRow.value(pTable.FindField("ClayFrac")) = 0
        End If
        pRow.Store
    End If
    
    ' Now Load the data into the Geodatabase......
    If pFLag Then
        Import_Shape_To_GDB gMapTempFolder, "Pollutants.dbf", gGDBpath, esriDTTable
        Set pTable = GetTable(gGDBpath, "Pollutants")
    End If
    
    ' Finally disable controls.....
    framePollutants.Enabled = False
    
    txtSandFrac.Text = 0
    txtSiltFrac.Text = 0
    txtClayFrac.Text = 0
    chkSedAssoc.value = 0
    frameSedAssoc.Enabled = False
    cmdSave.Enabled = False
    Call LoadPollutants
    
    If Not pTable Is Nothing And lstPollutants.ListItems.Count > 0 Then gDefPollutants = True

    Exit Sub
ErrorHandler:
  HandleError True, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If gExternalSimulation Then
        Label1.Caption = "Select the Pollutant to View/Edit Properties - Pollutant Order should match External Timeseries"
        frameSedAssoc.Enabled = False
        chkSedAssoc.value = 0
        txtTip.Height = 2970
        Me.Height = 6675
        cmdSave.Top = 5640
        cmdCancel.Top = 5640
        lstPollutants.ColumnHeaders(1).Width = 400
        lstPollutants.ColumnHeaders(2).Width = 1000
        lstPollutants.ColumnHeaders(3).Width = 1200
        lstPollutants.ColumnHeaders(4).Width = 1200
        lstPollutants.ColumnHeaders(5).Width = 1800
        lstPollutants.ColumnHeaders(6).Width = 1500
        lstPollutants.ColumnHeaders(7).Width = 1500
        lstPollutants.ColumnHeaders(8).Width = 1500
    Else
        Label1.Caption = "Select the Pollutant to View/Edit Properties"
        frameSedAssoc.Enabled = False
        
        txtTip.Height = 1005
        Me.Height = 4575
        cmdSave.Top = 3650
        cmdCancel.Top = 3650
        lstPollutants.ColumnHeaders(1).Width = 600
        lstPollutants.ColumnHeaders(2).Width = 2500
        lstPollutants.ColumnHeaders(3).Width = 1800
        lstPollutants.ColumnHeaders(4).Width = 2000
        lstPollutants.ColumnHeaders(5).Width = 0.01
        lstPollutants.ColumnHeaders(6).Width = 0.01
        lstPollutants.ColumnHeaders(7).Width = 0.01
        lstPollutants.ColumnHeaders(8).Width = 0.01
    End If
    
    cmbFlag.AddItem "NO"
    cmbFlag.AddItem "SAND"
    cmbFlag.AddItem "SILT"
    cmbFlag.AddItem "CLAY"
    cmbFlag.AddItem "SEDIMENT"
    
    Call LoadPollutants
  
    If m_PollutantDict.Exists("SEDIMENT") Or m_PollutantDict.Exists("SAND") Or m_PollutantDict.Exists("SILT") Or m_PollutantDict.Exists("CLAY") Then frameSedAssoc.Enabled = True
        
    lstPollutants.HideSelection = False
    If lstPollutants.ListItems.Count > 0 Then gDefPollutants = True

CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub LoadPollutants()
On Error GoTo ErrorHandler
    
    lstPollutants.ListItems.Clear
    txtName.Text = ""
    txtMultiplier.Text = "1.0"
    cmbFlag.ListIndex = 0
    Set m_PollutantDict = New Scripting.Dictionary
    
     '* Load the list box with property names
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID > 0 Order by ID"
    Dim lstItem As ListItem
    Dim pCursor As ICursor
    Dim pRow As iRow
    Set pCursor = pTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        Set lstItem = lstPollutants.ListItems.add(, , pRow.value(pTable.FindField("ID")))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("Name"))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("Multiplier"))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("Sediment"))
        
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("SedAssoc"))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("SandFrac"))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("SiltFrac"))
        lstItem.ListSubItems.add , , pRow.value(pTable.FindField("ClayFrac"))
        If Not m_PollutantDict.Exists(pRow.value(pTable.FindField("Sediment"))) Then m_PollutantDict.add pRow.value(pTable.FindField("Sediment")), pRow.value(pTable.FindField("Sediment"))
        Set pRow = pCursor.NextRow
    Loop

Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub



Private Sub txtClayFrac_GotFocus()
    txtTip.Text = "The sediment-associated qual-fraction on clay (0-1), only required for sediment associated pollutant"
End Sub

Private Sub txtMultiplier_Change()
    If Not IsNumeric(txtMultiplier.Text) Then SendKeys "{BS}"
End Sub

Private Sub txtMultiplier_GotFocus()
    txtTip.Text = "Enter muliplier factor. A non-negative number"
End Sub

Private Sub txtName_GotFocus()
    txtTip.Text = "Enter Pollutant name without spaces"
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then SendKeys "{BS}"
End Sub

Private Sub txtSandFrac_GotFocus()
    txtTip.Text = "The sediment-associated qual-fraction on sand (0-1), only required for sediment associated pollutant"
End Sub

Private Sub txtSiltFrac_GotFocus()
    txtTip.Text = "The sediment-associated qual-fraction on silt (0-1), only required for sediment associated pollutant"
End Sub
