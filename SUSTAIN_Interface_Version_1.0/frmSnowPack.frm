VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmSnowPack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snow Pack Editor"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   Icon            =   "frmSnowPack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   840
      TabIndex        =   26
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   6360
      Width           =   615
   End
   Begin VB.ListBox lstSnowPack 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtSnowPackID 
      Height          =   375
      Left            =   840
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSnowPackName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin TabDlg.SSTab TABParams 
      Height          =   5475
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Snap Pack Parameters"
      TabPicture(0)   =   "frmSnowPack.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGridSnowPack"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtfractionImp"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Snow Removal Parameters"
      TabPicture(1)   =   "frmSnowPack.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtSnowDepth"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtFracWS"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtFracImp"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtFracPer"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtFracmelt"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtFracSub"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtName"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.TextBox txtName 
         Height          =   405
         Left            =   -72960
         TabIndex        =   19
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtFracSub 
         Height          =   405
         Left            =   -69720
         TabIndex        =   17
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtFracmelt 
         Height          =   405
         Left            =   -69720
         TabIndex        =   15
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtFracPer 
         Height          =   405
         Left            =   -69720
         TabIndex        =   13
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtFracImp 
         Height          =   405
         Left            =   -69720
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtFracWS 
         Height          =   405
         Left            =   -69720
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSnowDepth 
         Height          =   405
         Left            =   -69720
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtfractionImp 
         Height          =   405
         Left            =   4920
         TabIndex        =   5
         Top             =   4800
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGridSnowPack 
         Height          =   4095
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   27
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Note: all fractions must be either zero or sum to 1.0."
         Height          =   255
         Left            =   -73800
         TabIndex        =   20
         Top             =   5040
         Width           =   4935
      End
      Begin VB.Label Label9 
         Caption         =   "(Name)"
         Height          =   255
         Left            =   -73800
         TabIndex        =   18
         Top             =   4305
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Fraction moved to another subcatchment"
         Height          =   255
         Left            =   -73800
         TabIndex        =   16
         Top             =   3705
         Width           =   3375
      End
      Begin VB.Label Label7 
         Caption         =   "Fraction converted into immediate melt"
         Height          =   255
         Left            =   -73800
         TabIndex        =   14
         Top             =   3105
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Fraction transferred to the pervious area"
         Height          =   255
         Left            =   -73800
         TabIndex        =   12
         Top             =   2505
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Fraction transferred to the impervious area"
         Height          =   255
         Left            =   -73800
         TabIndex        =   10
         Top             =   1905
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Fraction transferred out of the watershed"
         Height          =   255
         Left            =   -73800
         TabIndex        =   8
         Top             =   1305
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Depth at which snow removal begins (in)"
         Height          =   255
         Left            =   -73800
         TabIndex        =   6
         Top             =   705
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Fraction of Impervious Area That is Plowable"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   4905
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Snow Pack Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   285
      Width           =   1335
   End
End
Attribute VB_Name = "frmSnowPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    
     '** Load the pollutant names from the table
    Call LoadSnowPackNamesforForm
    Call LoadGridDetails
    
    txtSnowPackName.Enabled = True
    TABParams.Enabled = True
    cmdOK.Enabled = True
    txtSnowPackID.Text = ""
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
        
    On Error GoTo ShowError
    
    txtSnowPackName.Enabled = True
    TABParams.Enabled = True
    cmdOK.Enabled = True
    txtSnowPackID.Text = lstSnowPack.ListIndex + 1
    
    '** pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstSnowPack.ListIndex + 1
        
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = LoadSnowPackDetails(pPollutantID)
    
    '** return if nothing found
    If (pValueDictionary.Count = 0) Then
        Exit Sub
    End If
    
    txtSnowPackName.Enabled = True
    TABParams.Enabled = True
    txtSnowPackName.Text = pValueDictionary.Item("Name")
    txtfractionImp.Text = pValueDictionary.Item("Fraction Plowable")
    txtSnowDepth.Text = pValueDictionary.Item("Snow Depth")
    txtFracWS.Text = pValueDictionary.Item("Fraction Watershed")
    txtFracImp.Text = pValueDictionary.Item("Fraction Impervious")
    txtFracPer.Text = pValueDictionary.Item("Fraction Pervious")
    txtFracmelt.Text = pValueDictionary.Item("Fraction Melt")
    txtFracSub.Text = pValueDictionary.Item("Fraction subcatchment")
    txtName.Text = pValueDictionary.Item("SnowPackName")
    
    ' Now load the Grid details...........
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSnowPack.DataSource
    If oRs.State = adStateClosed Then
        oRs.CursorType = adOpenDynamic
        oRs.Open
    End If
    
    Dim pValues
    
    pValues = Split(pValueDictionary.Item("Min Coeff"), ";")
    oRs.MoveFirst
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Max Coeff"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Base Temp"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Fraction Capacity"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Initial Snow Depth"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Initial Free Water"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.MoveNext
    pValues = Split(pValueDictionary.Item("Depth"), ";")
    oRs.Fields(1).value = pValues(0)
    oRs.Fields(2).value = pValues(1)
    oRs.Fields(3).value = pValues(2)
    oRs.Update
    Set DataGridSnowPack.DataSource = oRs
    DataGridSnowPack.ColumnHeaders = True
    DataGridSnowPack.Columns(0).Caption = "Subcatchment Surface Type"
    DataGridSnowPack.Columns(0).Locked = True
    DataGridSnowPack.Columns(0).Width = 2400
    DataGridSnowPack.Columns(1).Caption = "Plowable"
    DataGridSnowPack.Columns(1).Width = 1300
    DataGridSnowPack.Columns(2).Caption = "Impervious"
    DataGridSnowPack.Columns(2).Width = 1300
    DataGridSnowPack.Columns(2).Caption = "Pervious"
    DataGridSnowPack.Columns(2).Width = 1300
    DataGridSnowPack.Refresh
    
    Exit Sub
ShowError:
    MsgBox "cmdEdit :" & Err.description
    
End Sub

Private Function LoadSnowPackDetails(pID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    '* Load the list box with pollutant names
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDSnowPacks")
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
    Set LoadSnowPackDetails = pValueDictionary
    
    '** Cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPollutantTable = Nothing
    Set pQueryFilter = Nothing
    Exit Function
ShowError:
    MsgBox "LoadSnowPackDetails: " & Err.description
    
End Function

Private Sub cmdOk_Click()
    
    On Error GoTo ShowError
    
    ' check if Duplicate exists.....
    If GetListBoxIndex(lstSnowPack, txtSnowPackName.Text) > -1 And txtSnowPackID.Text = "" Then Exit Sub

    '** All values are entered, save it a dictionary, and call a routine
    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = CreateObject("Scripting.Dictionary")
    
    '** get the pollutant ID
    Dim pPollutantID As String
    If (Trim(txtSnowPackID.Text) = "") Then
        pPollutantID = lstSnowPack.ListCount + 1
    Else
        pPollutantID = txtSnowPackID.Text
    End If
    
    pOptionProperty.add "Name", txtSnowPackName.Text
    pOptionProperty.add "Fraction Plowable", txtfractionImp.Text
    pOptionProperty.add "Snow Depth", txtSnowDepth.Text
    pOptionProperty.add "Fraction Watershed", txtFracWS.Text
    pOptionProperty.add "Fraction Impervious", txtFracImp.Text
    pOptionProperty.add "Fraction Pervious", txtFracPer.Text
    pOptionProperty.add "Fraction Melt", txtFracmelt.Text
    pOptionProperty.add "Fraction subcatchment", txtFracSub.Text
    pOptionProperty.add "SnowPackName", txtName.Text
    
    ' Now load the Grid details...........
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSnowPack.DataSource
    If oRs.State = adStateClosed Then
        oRs.CursorType = adOpenDynamic
        oRs.Open
    End If
    
    ' Now add the Grid details to the Table...
    oRs.MoveFirst
    pOptionProperty.add "Min Coeff", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Max Coeff", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Base Temp", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Fraction Capacity", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Initial Snow Depth", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Initial Free Water", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    oRs.MoveNext
    pOptionProperty.add "Depth", oRs.Fields(1).value & ";" & oRs.Fields(2).value & ";" & oRs.Fields(3).value
    
    '** Call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDSnowPacks", pPollutantID, pOptionProperty
        
    '** Clean up
    Set pOptionProperty = Nothing
    
    '** Load the pollutant names from the table
    Call LoadSnowPackNamesforForm
    Call LoadGridDetails
    
    txtSnowPackName.Enabled = False
    TABParams.Enabled = False
    cmdOK.Enabled = False
    
    Exit Sub
ShowError:
    MsgBox "cmdOk:" & Err.description
End Sub

Private Sub cmdRemove_Click()
    
    On Error GoTo ShowError
    '** Confirm the deletion
    Dim boolDelete
    boolDelete = MsgBox("Are you sure you want to delete this pollutant information ?", vbYesNo)
    If (boolDelete = vbNo) Then
        Exit Sub
    End If
    
    '** get pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstSnowPack.ListIndex + 1
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDSnowPacks")
    
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
    Call LoadSnowPackNamesforForm
    Call LoadGridDetails
    
    Exit Sub
ShowError:
    MsgBox "cmdRemove :" & Err.description
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Call LoadSnowPackNamesforForm
       
    Call LoadGridDetails
    
End Sub

Private Sub LoadGridDetails()
    
    On Error GoTo ShowError
    
    'If DBF file exists and the number of records match with that of the pollutants
    'then initialize the DB grid with the values
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "SubcatchmentType", adVarChar, 100
    oRs.Fields.Append "Plowable", adSingle
    oRs.Fields.Append "Impervious", adSingle
    oRs.Fields.Append "Pervious", adSingle
    oRs.CursorType = adOpenDynamic
    oRs.Open
                
    oRs.AddNew
    oRs.Fields(0).value = "Min. Melt Coeff. (in/hr/deg F)"
    oRs.Fields(1).value = 0.001
    oRs.Fields(2).value = 0.001
    oRs.Fields(3).value = 0.001
    oRs.AddNew
    oRs.Fields(0).value = "Max. Melt Coeff. (in/hr/deg F)"
    oRs.Fields(1).value = 0.001
    oRs.Fields(2).value = 0.001
    oRs.Fields(3).value = 0.001
    oRs.AddNew
    oRs.Fields(0).value = "Base Temperature (deg F)"
    oRs.Fields(1).value = 32#
    oRs.Fields(2).value = 32#
    oRs.Fields(3).value = 32#
    oRs.AddNew
    oRs.Fields(0).value = "Fraction Free Water Capacity"
    oRs.Fields(1).value = 0.1
    oRs.Fields(2).value = 0.1
    oRs.Fields(3).value = 0.1
    oRs.AddNew
    oRs.Fields(0).value = "Initial Snow Depth (in)"
    oRs.Fields(1).value = 0#
    oRs.Fields(2).value = 0#
    oRs.Fields(3).value = 0#
    oRs.AddNew
    oRs.Fields(0).value = "Initial Free Water (in)"
    oRs.Fields(1).value = 0#
    oRs.Fields(2).value = 0#
    oRs.Fields(3).value = 0#
    oRs.AddNew
    oRs.Fields(0).value = "Depth at 100% Cover (in)"
    oRs.Fields(1).value = 0#
    oRs.Fields(2).value = 0#
    oRs.Fields(3).value = 0#
        
    txtfractionImp.Text = "0.0"
    txtSnowDepth.Text = "1.0"
    txtFracWS.Text = "0.0"
    txtFracImp.Text = "0.0"
    txtFracPer.Text = "0.0"
    txtFracmelt.Text = "0.0"
    txtFracSub.Text = "0.0"
        
    '* Set datagrid value, header caption and width
    DataGridSnowPack.ClearFields
    Set DataGridSnowPack.DataSource = oRs
    DataGridSnowPack.ColumnHeaders = True
    DataGridSnowPack.Columns(0).Caption = "Subcatchment Surface Type"
    DataGridSnowPack.Columns(0).Locked = True
    DataGridSnowPack.Columns(0).Width = 2400
    DataGridSnowPack.Columns(1).Caption = "Plowable"
    DataGridSnowPack.Columns(1).Width = 1300
    DataGridSnowPack.Columns(2).Caption = "Impervious"
    DataGridSnowPack.Columns(2).Width = 1300
    DataGridSnowPack.Columns(2).Caption = "Pervious"
    DataGridSnowPack.Columns(2).Width = 1300
    
    Exit Sub
ShowError:
    MsgBox "Load Grid Details :" & Err.description
    
End Sub

Private Sub LoadSnowPackNamesforForm()
On Error GoTo ShowError

    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadSnowPackNames
    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    lstSnowPack.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        lstSnowPack.AddItem pPollutantCollection.Item(iCount)
    Next
    
    Set pPollutantCollection = Nothing
    Exit Sub
    
ShowError:
    MsgBox "LoadSnowPackNames :" & Err.description
End Sub
