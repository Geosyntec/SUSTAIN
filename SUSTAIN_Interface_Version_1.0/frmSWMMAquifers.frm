VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSWMMAquifers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aquifer Editor"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmSWMMAquifers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPollutantID 
      Height          =   375
      Left            =   720
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView lstAquifers 
      Height          =   5775
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   495
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtTip 
      BackColor       =   &H80000004&
      Height          =   570
      Left            =   2295
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6570
      Width           =   3915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   7245
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   7245
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aquifer Properties"
      Enabled         =   0   'False
      Height          =   6480
      Left            =   2280
      TabIndex        =   6
      Top             =   20
      Width           =   3975
      Begin VB.TextBox txtPorosity 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":08CA
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Text            =   "0.5"
         Top             =   5970
         Width           =   1400
      End
      Begin MSComctlLib.ListView lstviewHead 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Property"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtAqName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":0BD4
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Top             =   540
         Width           =   1400
      End
      Begin VB.TextBox txtPor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":0EDE
         Height          =   345
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   18
         Text            =   "0.5"
         Top             =   960
         Width           =   1400
      End
      Begin VB.TextBox txtWilting 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":11E8
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   17
         Text            =   "0.15"
         Top             =   1380
         Width           =   1400
      End
      Begin VB.TextBox txtCapacity 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":14F2
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Text            =   "0.30"
         Top             =   1800
         Width           =   1400
      End
      Begin VB.TextBox txtConduct 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":17FC
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   15
         Text            =   "5.0"
         Top             =   2220
         Width           =   1400
      End
      Begin VB.TextBox txtSlope 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":1B06
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         Text            =   "10.0"
         Top             =   2640
         Width           =   1400
      End
      Begin VB.TextBox txtTension 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":1E10
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Text            =   "15.0"
         Top             =   3060
         Width           =   1400
      End
      Begin VB.TextBox txtUpperEvap 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":211A
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Text            =   "0.35"
         Top             =   3480
         Width           =   1400
      End
      Begin VB.TextBox txtLowerEvap 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":2424
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Text            =   "14.0"
         Top             =   3885
         Width           =   1400
      End
      Begin VB.TextBox txtLowerGW 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":272E
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Text            =   "0.002"
         Top             =   4290
         Width           =   1400
      End
      Begin VB.TextBox txtBottom 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":2A38
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Text            =   "0.0"
         Top             =   4695
         Width           =   1400
      End
      Begin VB.TextBox txtWaterTable 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":2D42
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Text            =   "10.0"
         Top             =   5115
         Width           =   1400
      End
      Begin VB.TextBox txtUnsat 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         DragIcon        =   "frmSWMMAquifers.frx":304C
         Height          =   350
         Left            =   2475
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Text            =   "0.30"
         Top             =   5520
         Width           =   1400
      End
      Begin VB.Label lblPores 
         Caption         =   "Macropores Porosity"
         Height          =   330
         Left            =   240
         TabIndex        =   36
         Top             =   6030
         Width           =   2175
      End
      Begin VB.Line Line16 
         X1              =   120
         X2              =   3840
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label lblAqName 
         Caption         =   "Aquifer Name"
         Height          =   330
         Left            =   240
         TabIndex        =   33
         Top             =   585
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   2445
         X2              =   2445
         Y1              =   240
         Y2              =   6360
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3840
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3840
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label lblPorosity 
         Caption         =   "Total Porosity"
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   3840
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Label lblWiltingPoint 
         Caption         =   "Wilting Point"
         Height          =   330
         Left            =   240
         TabIndex        =   31
         Top             =   1425
         Width           =   2175
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   3840
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Label lblFieldCapacity 
         Caption         =   "Field Capacity"
         Height          =   330
         Left            =   240
         TabIndex        =   30
         Top             =   1815
         Width           =   2175
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   3840
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Label lblConductivity 
         Caption         =   "Conductivity"
         Height          =   330
         Left            =   240
         TabIndex        =   29
         Top             =   2265
         Width           =   2175
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   3840
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Label lblConductSlope 
         Caption         =   "Conduct. Slope"
         Height          =   330
         Left            =   240
         TabIndex        =   28
         Top             =   2670
         Width           =   2175
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   3840
         Y1              =   3435
         Y2              =   3435
      End
      Begin VB.Label lblTensionSlope 
         Caption         =   "Tension Slope"
         Height          =   330
         Left            =   240
         TabIndex        =   27
         Top             =   3090
         Width           =   2175
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   3840
         Y1              =   3855
         Y2              =   3855
      End
      Begin VB.Label lblUpperEvapFraction 
         Caption         =   "Upper Evap. Fraction"
         Height          =   330
         Left            =   240
         TabIndex        =   26
         Top             =   3510
         Width           =   2175
      End
      Begin VB.Line Line10 
         X1              =   120
         X2              =   3840
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label lblLowerEvapDepth 
         Caption         =   "Lower Evap. Depth"
         Height          =   330
         Left            =   225
         TabIndex        =   25
         Top             =   3915
         Width           =   2175
      End
      Begin VB.Label lblLowerGWLossRate 
         Caption         =   "Lower GW Loss Rate"
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Line Line11 
         X1              =   120
         X2              =   3840
         Y1              =   4665
         Y2              =   4665
      End
      Begin VB.Label lblBottomElevation 
         Caption         =   "Bottom Elevation"
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   4725
         Width           =   2175
      End
      Begin VB.Line Line13 
         X1              =   120
         X2              =   3840
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Label lblWaterTableElevation 
         Caption         =   "Water Table Elevation"
         Height          =   330
         Left            =   240
         TabIndex        =   22
         Top             =   5145
         Width           =   2175
      End
      Begin VB.Line Line15 
         X1              =   120
         X2              =   3840
         Y1              =   5490
         Y2              =   5490
      End
      Begin VB.Label lblUnsatZoneMoisture 
         Caption         =   "Unsat. Zone Moisture"
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   5550
         Width           =   2175
      End
      Begin VB.Line Line17 
         X1              =   120
         X2              =   3840
         Y1              =   5895
         Y2              =   5895
      End
      Begin VB.Line Line12 
         X1              =   3870
         X2              =   3870
         Y1              =   240
         Y2              =   6360
      End
      Begin VB.Line Line14 
         X1              =   120
         X2              =   120
         Y1              =   240
         Y2              =   6360
      End
   End
End
Attribute VB_Name = "frmSWMMAquifers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_txtBackColor
Private m_pAquiferID As String

Private Sub cmdAdd_Click()
    
    Call LoadAquifers
    Frame1.Enabled = True
    cmdOK.Enabled = True
    txtPollutantID.Text = ""
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
        
    If lstAquifers.ListItems.Count = 0 Then Exit Sub
    Frame1.Enabled = True
    cmdOK.Enabled = True
    txtPollutantID.Text = lstAquifers.SelectedItem.Index
    
    '** pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstAquifers.SelectedItem.Index
    
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = LoadAquiferDetails(pPollutantID)
    
    '** return if nothing found
    If (pValueDictionary.Count = 0) Then
        Exit Sub
    End If
    
    Frame1.Enabled = True
    txtAqName.Text = pValueDictionary.Item("Name")
    txtPor.Text = pValueDictionary.Item("Porosity")
    txtWilting.Text = pValueDictionary.Item("Wilting Point")
    txtCapacity.Text = pValueDictionary.Item("Field Capacity")
    txtConduct.Text = pValueDictionary.Item("Conductivity")
    txtSlope.Text = pValueDictionary.Item("Conduct Slope")
    txtTension.Text = pValueDictionary.Item("Tension Slope")
    txtUpperEvap.Text = pValueDictionary.Item("Upper Evap Fraction")
    txtLowerEvap.Text = pValueDictionary.Item("Lower Evap Depth")
    txtLowerGW.Text = pValueDictionary.Item("Lower GW Loss Rate")
    txtBottom.Text = pValueDictionary.Item("Bottom Elevation")
    txtWaterTable.Text = pValueDictionary.Item("Water Table Elevation")
    txtUnsat.Text = pValueDictionary.Item("Unsat Zone Moisture")
    txtPorosity.Text = pValueDictionary.Item("Macropores Porosity")
    
End Sub

Private Sub cmdOk_Click()
        
    ' check if Duplicate exists.....
    Dim lstItem As ListItem
    Set lstItem = lstAquifers.FindItem(txtAqName.Text)
    If Not lstItem Is Nothing And txtPollutantID.Text = "" Then Exit Sub
    
    '** get the pollutant ID
    If (Trim(txtPollutantID.Text) = "") Then
        m_pAquiferID = lstAquifers.ListItems.Count + 1
    Else
        m_pAquiferID = txtPollutantID.Text
    End If
 
    '** All values are entered, save it a dictionary, and call a routine
    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = CreateObject("Scripting.Dictionary")
    
    '** Add all values to the dictionary
    pOptionProperty.add "Name", txtAqName.Text
    pOptionProperty.add "Porosity", txtPor.Text
    pOptionProperty.add "Wilting Point", txtWilting.Text
    pOptionProperty.add "Field Capacity", txtCapacity.Text
    pOptionProperty.add "Conductivity", txtConduct.Text
    pOptionProperty.add "Conduct Slope", txtSlope.Text
    pOptionProperty.add "Tension Slope", txtTension.Text
    pOptionProperty.add "Upper Evap Fraction", txtUpperEvap.Text
    pOptionProperty.add "Lower Evap Depth", txtLowerEvap.Text
    pOptionProperty.add "Lower GW Loss Rate", txtLowerGW.Text
    pOptionProperty.add "Bottom Elevation", txtBottom.Text
    pOptionProperty.add "Water Table Elevation", txtWaterTable.Text
    pOptionProperty.add "Unsat Zone Moisture", txtUnsat.Text
    pOptionProperty.add "Macropores Porosity", txtPorosity.Text
            
    
    '** Call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDAquifers", m_pAquiferID, pOptionProperty
        
    '** Clean up
    Set pOptionProperty = Nothing
    
    '** Load the pollutant names from the table
    Call LoadAquifers
    cmdOK.Enabled = False
        
End Sub

Private Sub cmdRemove_Click()
    
    '** Confirm the deletion
    If lstAquifers.ListItems.Count = 0 Then Exit Sub
    Dim boolDelete
    boolDelete = MsgBox("Are you sure you want to delete this aquifer information ?", vbYesNo)
    If (boolDelete = vbNo) Then
        Exit Sub
    End If
    
    '** get pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstAquifers.SelectedItem.Index
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDAquifers")
    
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
    Call LoadAquifers
    
End Sub

Private Sub Form_Load()
    On Error GoTo ShowError
    
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    lstviewHead.Height = 275
    lstviewHead.ColumnHeaders.Item(1).Width = lstviewHead.Width
    lstviewHead.FlatScrollBar = False
    lstviewHead.Enabled = False
    
    m_txtBackColor = txtAqName.BackColor ' &H579977
    txtAqName.BackColor = Me.BackColor
    
    Call LoadAquifers

Exit Sub
    
ShowError:
    MsgBox "Form_Load :" & Err.description
End Sub

Public Sub LoadAquifers()
On Error GoTo ShowError

    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadAquiferNames
    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    lstAquifers.ListItems.Clear
    Dim lstItem As ListItem
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        Set lstItem = lstAquifers.ListItems.add(, , pPollutantCollection.Item(iCount))
    Next
    
    Set pPollutantCollection = Nothing
    
    ' Load the Default values.....
    txtAqName.Text = ""
    txtPorosity.Text = "0.5"
    txtWilting.Text = "0.15"
    txtCapacity.Text = "0.30"
    txtConduct.Text = "5.0"
    txtSlope.Text = "10.0"
    txtTension.Text = "15.0"
    txtUpperEvap.Text = "0.35"
    txtLowerEvap.Text = "14.0"
    txtLowerGW.Text = "0.002"
    txtBottom.Text = "0.0"
    txtWaterTable.Text = "10.0"
    txtUnsat.Text = "0.30"
    txtPorosity.Text = "0.5"
    
    Exit Sub
    
ShowError:
    MsgBox "LoadAquifers :" & Err.description
End Sub

Private Function LoadAquiferDetails(pID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    '* Load the list box with pollutant names
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDAquifers")
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
    Set LoadAquiferDetails = pValueDictionary
    
    '** Cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPollutantTable = Nothing
    Set pQueryFilter = Nothing
    Exit Function
ShowError:
    MsgBox "LoadAquiferDetails: " & Err.description
    
End Function


' **************************************
Private Sub txtPorosity_GotFocus()
    txtPorosity.BackColor = m_txtBackColor
    txtPorosity.MousePointer = 0
End Sub

Private Sub txtPorosity_LostFocus()
    txtPorosity.BackColor = Me.BackColor
    txtPorosity.MousePointer = 1
End Sub
Private Sub lblPores_Click()
    txtPorosity.SetFocus
    txtPorosity.SelStart = 0
    txtPorosity.SelLength = Len(txtPorosity.Text)
End Sub


' **************************************
Private Sub txtAqName_GotFocus()
    txtAqName.BackColor = m_txtBackColor
    txtAqName.MousePointer = 0
End Sub

Private Sub txtAqName_LostFocus()
    txtAqName.BackColor = Me.BackColor
    txtAqName.MousePointer = 1
End Sub
Private Sub lblAqName_Click()
    txtAqName.SetFocus
    txtAqName.SelStart = 0
    txtAqName.SelLength = Len(txtAqName.Text)
End Sub



' **************************************
Private Sub txtBottom_GotFocus()
    txtBottom.BackColor = m_txtBackColor
    txtBottom.MousePointer = 0
End Sub

Private Sub txtBottom_LostFocus()
    txtBottom.BackColor = Me.BackColor
    txtBottom.MousePointer = 1
End Sub
Private Sub lblBottomElevation_Click()
    txtBottom.SetFocus
    txtBottom.SelStart = 0
    txtBottom.SelLength = Len(txtBottom.Text)
End Sub

' **************************************
Private Sub txtCapacity_GotFocus()
    txtCapacity.BackColor = m_txtBackColor
    txtCapacity.MousePointer = 0
End Sub

Private Sub txtCapacity_LostFocus()
    txtCapacity.BackColor = Me.BackColor
    txtCapacity.MousePointer = 1
End Sub
Private Sub lblFieldCapacity_Click()
    txtCapacity.SetFocus
    txtCapacity.SelStart = 0
    txtCapacity.SelLength = Len(txtCapacity.Text)
End Sub


' **************************************
Private Sub txtConduct_GotFocus()
    txtConduct.BackColor = m_txtBackColor
    txtConduct.MousePointer = 0
End Sub

Private Sub txtConduct_LostFocus()
    txtConduct.BackColor = Me.BackColor
    txtConduct.MousePointer = 1
End Sub
Private Sub lblConductivity_Click()
    txtConduct.SetFocus
    txtConduct.SelStart = 0
    txtConduct.SelLength = Len(txtConduct.Text)
End Sub


' **************************************
Private Sub txtLowerEvap_GotFocus()
    txtLowerEvap.BackColor = m_txtBackColor
    txtLowerEvap.MousePointer = 0
End Sub

Private Sub txtLowerEvap_LostFocus()
    txtLowerEvap.BackColor = Me.BackColor
    txtLowerEvap.MousePointer = 1
End Sub
Private Sub lblLowerEvapDepth_Click()
    txtLowerEvap.SetFocus
    txtLowerEvap.SelStart = 0
    txtLowerEvap.SelLength = Len(txtLowerEvap.Text)
End Sub


' **************************************
Private Sub txtLowerGW_GotFocus()
    txtLowerGW.BackColor = m_txtBackColor
    txtLowerGW.MousePointer = 0
End Sub

Private Sub txtLowerGW_LostFocus()
    txtLowerGW.BackColor = Me.BackColor
    txtLowerGW.MousePointer = 1
End Sub
Private Sub lblLowerGWLossRate_Click()
    txtLowerGW.SetFocus
    txtLowerGW.SelStart = 0
    txtLowerGW.SelLength = Len(txtLowerGW.Text)
End Sub

' **************************************
Private Sub txtPor_Gotfocus()
    txtPor.BackColor = m_txtBackColor
    txtPor.MousePointer = 0
End Sub

Private Sub txtPor_LostFocus()
    txtPor.BackColor = Me.BackColor
    txtPor.MousePointer = 1
End Sub
Private Sub lblPorosity_Click()
    txtPor.SetFocus
    txtPor.SelStart = 0
    txtPor.SelLength = Len(txtPor.Text)
End Sub


' **************************************
Private Sub txtSlope_GotFocus()
    txtSlope.BackColor = m_txtBackColor
    txtSlope.MousePointer = 0
End Sub

Private Sub txtSlope_LostFocus()
    txtSlope.BackColor = Me.BackColor
    txtSlope.MousePointer = 1
End Sub
Private Sub lblConductSlope_Click()
    txtSlope.SetFocus
    txtSlope.SelStart = 0
    txtSlope.SelLength = Len(txtSlope.Text)
End Sub


' **************************************
Private Sub txtTension_GotFocus()
    txtTension.BackColor = m_txtBackColor
    txtTension.MousePointer = 0
End Sub

Private Sub txtTension_LostFocus()
    txtTension.BackColor = Me.BackColor
    txtTension.MousePointer = 1
End Sub
Private Sub lblTensionSlope_Click()
    txtTension.SetFocus
    txtTension.SelStart = 0
    txtTension.SelLength = Len(txtTension.Text)
End Sub



' **************************************
Private Sub txtUnsat_GotFocus()
    txtUnsat.BackColor = m_txtBackColor
    txtUnsat.MousePointer = 0
End Sub

Private Sub txtUnsat_LostFocus()
    txtUnsat.BackColor = Me.BackColor
    txtUnsat.MousePointer = 1
End Sub
Private Sub lblUnsatZoneMoisture_Click()
    txtUnsat.SetFocus
    txtUnsat.SelStart = 0
    txtUnsat.SelLength = Len(txtUnsat.Text)
End Sub


' **************************************
Private Sub txtUpperEvap_GotFocus()
    txtUpperEvap.BackColor = m_txtBackColor
    txtUpperEvap.MousePointer = 0
End Sub

Private Sub txtUpperEvap_LostFocus()
    txtUpperEvap.BackColor = Me.BackColor
    txtUpperEvap.MousePointer = 1
End Sub
Private Sub lblUpperEvapFraction_Click()
    txtUpperEvap.SetFocus
    txtUpperEvap.SelStart = 0
    txtUpperEvap.SelLength = Len(txtUpperEvap.Text)
End Sub


' **************************************
Private Sub txtWaterTable_GotFocus()
    txtWaterTable.BackColor = m_txtBackColor
    txtWaterTable.MousePointer = 0
End Sub

Private Sub txtWaterTable_LostFocus()
    txtWaterTable.BackColor = Me.BackColor
    txtWaterTable.MousePointer = 1
End Sub
Private Sub lblWaterTableElevation_Click()
    txtWaterTable.SetFocus
    txtWaterTable.SelStart = 0
    txtWaterTable.SelLength = Len(txtWaterTable.Text)
End Sub

' **************************************
Private Sub txtWilting_Gotfocus()
    txtWilting.BackColor = m_txtBackColor
    txtWilting.MousePointer = 0
End Sub

Private Sub txtWilting_LostFocus()
    txtWilting.BackColor = Me.BackColor
    txtWilting.MousePointer = 1
End Sub
Private Sub lblWiltingPoint_Click()
    txtWilting.SetFocus
    txtWilting.SelStart = 0
    txtWilting.SelLength = Len(txtWilting.Text)
End Sub
