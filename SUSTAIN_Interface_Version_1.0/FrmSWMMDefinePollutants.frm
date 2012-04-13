VERSION 5.00
Begin VB.Form FrmSWMMDefinePollutants 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Pollutant Properties"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "FrmSWMMDefinePollutants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save All"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6840
      TabIndex        =   26
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtPollutantID 
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame framePollutants 
      Caption         =   "Pollutant Properties"
      Height          =   4815
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   4095
      Begin VB.TextBox txtCoFraction 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtCoPollutant 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   3705
         Width           =   2175
      End
      Begin VB.ComboBox cmbSnowOnly 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3270
         Width           =   1215
      End
      Begin VB.TextBox txtDecayCoeff 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Text            =   "0.0"
         Top             =   2775
         Width           =   1215
      End
      Begin VB.TextBox txtIIConc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Text            =   "0.0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtGWConc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Text            =   "0.0"
         Top             =   1785
         Width           =   1215
      End
      Begin VB.TextBox txtRainConc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Text            =   "0.0"
         Top             =   1290
         Width           =   1215
      End
      Begin VB.ComboBox cmbUnits 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   855
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Co-Fraction"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Co-Pollutant"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3705
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Snow Only"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Decay Coeff"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2775
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "I && I Conc."
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "GW Conc."
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1785
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Rain Conc."
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Units"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "-"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox listPollutants 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Pollutant to View/Edit Properties"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "FrmSWMMDefinePollutants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    txtCoFraction.Enabled = True
    txtCoPollutant.Enabled = True
    txtDecayCoeff.Enabled = True
    txtGWConc.Enabled = True
    txtIIConc.Enabled = True
    'txtName.Enabled = True
    txtRainConc.Enabled = True
    cmbSnowOnly.Enabled = True
    cmbUnits.Enabled = True
    
    '** Clear all input fields
    txtPollutantID.Text = ""
    txtName.Text = ""
    cmbUnits.ListIndex = 0
    txtRainConc.Text = "0.0"
    txtGWConc.Text = "0.0"
    txtIIConc.Text = "0.0"
    txtDecayCoeff.Text = "0.0"
    cmbSnowOnly.ListIndex = 0
    txtCoPollutant.Text = ""
    txtCoFraction.Text = ""
    
    
End Sub

Private Sub cmdCancel_Click()
    
    ' ** check & delete the pollutants from SWMM pollutants table.....
    Dim iCnt As Integer
    Dim strPoll As String
    For iCnt = 0 To listPollutants.ListCount - 1
        strPoll = strPoll & "','" & listPollutants.List(iCnt)
    Next iCnt
    strPoll = Mid(strPoll & "'", 3)
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDPollutants")
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iID As Long
    Dim pCol As Collection
    Set pCol = New Collection
    iID = pSWMMPollutantTable.FindField("ID")
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName='Name' And PropValue NOT IN (" & strPoll & ")"
    Set pCursor = pSWMMPollutantTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        pCol.add pRow.value(iID)
        Set pRow = pCursor.NextRow
    Loop
    Set pRow = Nothing
    Set pCursor = Nothing
    
    For iCnt = 1 To pCol.Count
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "ID = " & pCol.Item(iCnt)
        '** delete records
        pSWMMPollutantTable.DeleteSearchedRows pQueryFilter
    Next iCnt
    
    '** clean up
    Set pQueryFilter = Nothing
    Set pSWMMPollutantTable = Nothing

    Unload Me
End Sub

'** Edit a pollutant properties
Private Sub cmdEdit_Click()

    '** Enable them
    Call cmdAdd_Click
    
    txtName.Text = listPollutants.Text
    '*** Update the hidden pollutant id value
    txtPollutantID.Text = listPollutants.ListIndex + 1
    
    '** pollutant id
    Dim pPollutantID As Integer
    pPollutantID = listPollutants.ListIndex + 1
    
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = LoadPollutantDetails(pPollutantID)
    
    '** return if nothing found
    If pValueDictionary Is Nothing Then
        Exit Sub
    End If
    If pValueDictionary.Count = 0 Then
        Exit Sub
    End If
    
    txtName.Text = pValueDictionary.Item("Name")
    cmbUnits.Text = pValueDictionary.Item("Mass Units")
    txtRainConc.Text = pValueDictionary.Item("Rain Conc.")
    txtGWConc.Text = pValueDictionary.Item("GW Conc.")
    txtIIConc.Text = pValueDictionary.Item("I&I Conc.")
    txtDecayCoeff.Text = pValueDictionary.Item("Decay Coeff.")
    cmbSnowOnly.Text = pValueDictionary.Item("Snow Only")
    Dim pCoPollutantVal
    pCoPollutantVal = Split(pValueDictionary.Item("Co-Pollutant"), " ")
    If (UBound(pCoPollutantVal) >= 0) Then
        txtCoPollutant.Text = pCoPollutantVal(0)
    End If
    If (UBound(pCoPollutantVal) = 1) Then
        txtCoFraction.Text = pCoPollutantVal(1)
    End If
    
    
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    '** Confirm the deletion
    Dim boolDelete
    boolDelete = MsgBox("Are you sure you want to delete this pollutant information ?", vbYesNo)
    If (boolDelete = vbNo) Then
        Exit Sub
    End If
    
    '** get pollutant id
    Dim pPollutantID As Integer
    pPollutantID = listPollutants.ListIndex + 1
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDPollutants")
    
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
    LoadPollutantNamesForPollutantForm
End Sub

Public Sub LoadPollutantNamesForPollutantForm()
On Error GoTo ShowError

    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadPollutantNames
    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    listPollutants.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        listPollutants.AddItem pPollutantCollection.Item(iCount)
    Next
    
    Set pPollutantCollection = Nothing
    Exit Sub
    
ShowError:
    MsgBox "LoadPollutantNamesForPollutantForm :" & Err.description
End Sub
Private Sub cmdSave_Click()
On Error GoTo ShowError
    '** Input data-validation
    '* pollutant name
    If (txtName.Text = "") Then
        MsgBox "Please specify the pollutant name to continue.", vbExclamation
        If txtName.Enabled Then txtName.SetFocus
        Exit Sub
    End If
    '* rain water concentration
    If (Not IsNumeric(txtRainConc.Text)) Then
        MsgBox "Please specify concentration of pollutant in rain water as a valid number.", vbExclamation
        txtRainConc.SetFocus
        Exit Sub
    End If
    '* ground water concentration
    If (Not IsNumeric(txtGWConc.Text)) Then
        MsgBox "Please specify concentration of pollutant in ground water as a valid number.", vbExclamation
        txtGWConc.SetFocus
        Exit Sub
    End If
    '* I & I concentration
    If (Not IsNumeric(txtIIConc.Text)) Then
        MsgBox "Please specify concentration of pollutant in I&I flow as a valid number.", vbExclamation
        txtIIConc.SetFocus
        Exit Sub
    End If
    '* decay co-efficient
    If (Not IsNumeric(txtDecayCoeff.Text)) Then
        MsgBox "Please specify first order decay co-efficient of pollutant (1/days) as a valid number.", vbExclamation
        txtDecayCoeff.SetFocus
        Exit Sub
    End If
    
    '** All values are entered, save it a dictionary, and call a routine
    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = CreateObject("Scripting.Dictionary")
    
    '** Add all values to the dictionary
    pOptionProperty.add "Name", txtName.Text
    pOptionProperty.add "Mass Units", cmbUnits.Text
    pOptionProperty.add "Rain Conc.", txtRainConc.Text
    pOptionProperty.add "GW Conc.", txtGWConc.Text
    pOptionProperty.add "I&I Conc.", txtIIConc.Text
    pOptionProperty.add "Decay Coeff.", txtDecayCoeff.Text
    pOptionProperty.add "Snow Only", cmbSnowOnly.Text
    pOptionProperty.add "Co-Pollutant", txtCoPollutant.Text & " " & txtCoFraction.Text
            
    '** get the pollutant ID
    Dim pPollutantID As String
    If (Trim(txtPollutantID.Text) = "") Then
        pPollutantID = listPollutants.ListCount + 1
    Else
        pPollutantID = txtPollutantID.Text
    End If
    
    '** Call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDPollutants", pPollutantID, pOptionProperty
        
    '** Clean up
    Set pOptionProperty = Nothing
    
    '** Load the pollutant names from the table
    'Call LoadPollutantNamesForPollutantForm
    
    '** Clear all input fields
    txtPollutantID.Text = ""
    txtName.Text = ""
    cmbUnits.ListIndex = 0
    txtRainConc.Text = "0.0"
    txtGWConc.Text = "0.0"
    txtIIConc.Text = "0.0"
    txtDecayCoeff.Text = "0.0"
    cmbSnowOnly.ListIndex = 0
    txtCoPollutant.Text = ""
    txtCoFraction.Text = ""
    
    '** Set them to false
    txtCoFraction.Enabled = False
    txtCoPollutant.Enabled = False
    txtDecayCoeff.Enabled = False
    txtGWConc.Enabled = False
    txtIIConc.Enabled = False
    txtName.Enabled = False
    txtRainConc.Enabled = False
    cmbSnowOnly.Enabled = False
    cmbUnits.Enabled = False
    Exit Sub
ShowError:
    MsgBox "Error in Saving :" & Err.description
End Sub

Private Sub cmdSaveAll_Click()
    
    Dim iCnt As Integer
    For iCnt = 0 To listPollutants.ListCount - 1
        listPollutants.Selected(iCnt) = True
        Call cmdEdit_Click
        Call cmdSave_Click
    Next iCnt
    
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '* clear the hidden pollutant id
    txtPollutantID.Text = ""
    
    '* Load the combo box for Snow Only parameter
    cmbSnowOnly.AddItem "NO"
    cmbSnowOnly.AddItem "YES"
    cmbSnowOnly.ListIndex = 0
    
    '* Load the combo box for Units parameter
    cmbUnits.AddItem "MG/L"
    'cmbUnits.AddItem "UG/L"
    'cmbUnits.AddItem "#/L"
    cmbUnits.ListIndex = 0
    
    '* Load the pollutant names from the table
    Call LoadPollutants

End Sub

Private Sub LoadPollutants()
    
    On Error GoTo ErrorHandler
    listPollutants.Clear
        
     '* Load the list box with property names
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID > -1 Order by ID"
    Dim lstItem As ListItem
    Dim pCursor As ICursor
    Dim pRow As iRow
    Set pCursor = pTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        listPollutants.AddItem pRow.value(pTable.FindField("Name"))
        Set pRow = pCursor.NextRow
    Loop
    
Exit Sub
ErrorHandler:
     MsgBox "LoadPollutants: " & Err.description
End Sub

Private Function LoadPollutantDetails(pID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    '* Load the list box with pollutant names
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDPollutants")
    If (pSWMMPollutantTable Is Nothing) Then
        Exit Function
    End If
    Dim iPropName As Long
    iPropName = pSWMMPollutantTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMPollutantTable.FindField("PropValue")
    
    'Define query filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pID
    
    'Get the cursor to iterate over the table
    Dim pCursor As ICursor
    Set pCursor = pSWMMPollutantTable.Search(pQueryFilter, True)
    
    'Define a dictionary to store the values
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = CreateObject("Scripting.Dictionary")
    
    'Define a row variable to loop over the table
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        pValueDictionary.add Trim(pRow.value(iPropName)), Trim(pRow.value(iPropValue))
        Set pRow = pCursor.NextRow
    Loop
    
    '** Return the dictionary
    Set LoadPollutantDetails = pValueDictionary
    
    '** Cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPollutantTable = Nothing
    Set pQueryFilter = Nothing
    Exit Function
ShowError:
    MsgBox "LoadPollutantDetails: " & Err.description
    
End Function

