VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transect Editor"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   Icon            =   "frmTransect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGridSnowPack 
      Height          =   3975
      Left            =   2400
      TabIndex        =   27
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtElev 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":08CA
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      Text            =   "1000"
      Top             =   4320
      Width           =   1520
   End
   Begin VB.TextBox txtStations 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":0BD4
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   22
      Text            =   "0"
      Top             =   3960
      Width           =   1520
   End
   Begin VB.TextBox txtRight 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":0EDE
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Text            =   "10"
      Top             =   3240
      Width           =   1520
   End
   Begin VB.TextBox txtLeft 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":11E8
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      Text            =   "1"
      Top             =   2880
      Width           =   1520
   End
   Begin VB.TextBox txtChannel 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":14F2
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Text            =   "0.01"
      Top             =   2160
      Width           =   1520
   End
   Begin VB.TextBox txtRightBank 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":17FC
      Height          =   350
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Text            =   "0.01"
      Top             =   1800
      Width           =   1520
   End
   Begin VB.TextBox txtLeftBank 
      Appearance      =   0  'Flat
      DragIcon        =   "frmTransect.frx":1B06
      Height          =   345
      Left            =   7395
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Text            =   "0.01"
      Top             =   1440
      Width           =   1520
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.ListBox lstSnowPack 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtSnowPackID 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSnowPackName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin MSComctlLib.ListView lstviewHead 
      Height          =   300
      Left            =   5880
      TabIndex        =   9
      Top             =   780
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   529
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   7380
      X2              =   7380
      Y1              =   960
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   5880
      X2              =   5880
      Y1              =   960
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   8910
      X2              =   8910
      Y1              =   960
      Y2              =   4680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   8880
      X2              =   5880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label8 
      Caption         =   "Elevations"
      Height          =   330
      Left            =   6000
      TabIndex        =   25
      Top             =   4380
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Stations"
      Height          =   330
      Left            =   6000
      TabIndex        =   23
      Top             =   4005
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Modifiers:"
      Height          =   330
      Left            =   5880
      TabIndex        =   21
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
      Height          =   330
      Left            =   6000
      TabIndex        =   20
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
      Height          =   330
      Left            =   6000
      TabIndex        =   18
      Top             =   2925
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Bank Stations:"
      Height          =   330
      Left            =   5880
      TabIndex        =   16
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Channel"
      Height          =   330
      Left            =   6000
      TabIndex        =   15
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label lblWiltingPoint 
      Caption         =   "Right Bank"
      Height          =   330
      Left            =   6000
      TabIndex        =   13
      Top             =   1825
      Width           =   1335
   End
   Begin VB.Label lblPorosity 
      Caption         =   "Left Bank"
      Height          =   315
      Left            =   6000
      TabIndex        =   12
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Transect Name"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   285
      Width           =   1335
   End
   Begin VB.Label lblAqName 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Roughness:"
      Height          =   330
      Left            =   5880
      TabIndex        =   28
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "frmTransect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    
     '** Load the pollutant names from the table
    Call LoadSnowPackNamesforForm
    Call LoadDefaults
    
    txtSnowPackName.Enabled = True
    DataGridSnowPack.Enabled = True
    cmdOK.Enabled = True
    txtSnowPackID.Text = ""
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
        
    On Error GoTo ShowError
    
    txtSnowPackName.Enabled = True
    DataGridSnowPack.Enabled = True
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
    txtSnowPackName.Text = pValueDictionary.Item("Name")
    txtLeftBank.Text = pValueDictionary.Item("Left Bank")
    txtRightBank.Text = pValueDictionary.Item("Right Bank")
    txtChannel.Text = pValueDictionary.Item("Channel")
    txtLeft.Text = pValueDictionary.Item("Left")
    txtRight.Text = pValueDictionary.Item("Right")
    txtStations.Text = pValueDictionary.Item("Stations")
    txtElev.Text = pValueDictionary.Item("Elevations")
    
    ' Now load the Grid details...........
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "Station", adInteger
    oRs.Fields.Append "Elevation", adInteger
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    Dim pValues(0 To 1), iCnt As Integer
    pValues(0) = Split(pValueDictionary.Item("Stations_grid"), ";")
    pValues(1) = Split(pValueDictionary.Item("Elevations_grid"), ";")
    For iCnt = 0 To UBound(pValues(0))
        oRs.AddNew
        oRs.Fields(0).value = pValues(0)(iCnt)
        oRs.Fields(1).value = pValues(1)(iCnt)
    Next
      
    
    Set DataGridSnowPack.DataSource = oRs
    DataGridSnowPack.ColumnHeaders = True
    DataGridSnowPack.Columns(0).Caption = "Station (ft)"
    DataGridSnowPack.Columns(0).Width = DataGridSnowPack.Width / 2.2
    DataGridSnowPack.Columns(1).Caption = "Elevation (ft)"
    DataGridSnowPack.Columns(1).Width = DataGridSnowPack.Width / 2.2
    DataGridSnowPack.Refresh
    
    Exit Sub
ShowError:
    MsgBox "cmdEdit :" & Err.description
    
End Sub

Private Function LoadSnowPackDetails(pID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    '* Load the list box with pollutant names
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("Transects")
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
    
    ' Check for Transect Values....
    ' Now load the Grid details...........
    Dim strStations As String, strElevations As String
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSnowPack.DataSource
    If oRs.State = adStateClosed Then
        oRs.CursorType = adOpenDynamic
        oRs.Open
    End If
    If oRs.RecordCount = 0 Then Exit Sub
    
    ' Now add the Grid details to the Table...
    oRs.MoveFirst
    Do While Not oRs.EOF
        strStations = strStations & ";" & oRs.Fields(0).value
        strElevations = strElevations & ";" & oRs.Fields(1).value
        oRs.MoveNext
    Loop
    
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
    pOptionProperty.add "Left Bank", txtLeftBank.Text
    pOptionProperty.add "Right Bank", txtRightBank.Text
    pOptionProperty.add "Channel", txtChannel.Text
    pOptionProperty.add "Left", txtLeft.Text
    pOptionProperty.add "Right", txtRight.Text
    pOptionProperty.add "Stations", txtStations.Text
    pOptionProperty.add "Elevations", txtElev.Text
    pOptionProperty.add "Stations_grid", Mid(strStations, 2)
    pOptionProperty.add "Elevations_grid", Mid(strElevations, 2)
    
    '** Call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "Transects", pPollutantID, pOptionProperty
        
    '** Clean up
    Set pOptionProperty = Nothing
    
    '** Load the pollutant names from the table
    Call LoadSnowPackNamesforForm
    Call LoadDefaults
    
    txtSnowPackName.Enabled = False
    DataGridSnowPack.Enabled = False
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
    Set pSWMMPollutantTable = GetInputDataTable("Transects")
    
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
    Call LoadDefaults
    
    Exit Sub
ShowError:
    MsgBox "cmdRemove :" & Err.description
    
End Sub

Private Sub cmdView_Click()
    
    Dim oValDict1 As Scripting.Dictionary
    Dim oValDict2 As Scripting.Dictionary
    Set oValDict1 = New Scripting.Dictionary
    Set oValDict2 = New Scripting.Dictionary
    
    ' Now load the Grid details...........
    Dim strStations As String, strElevations As String
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSnowPack.DataSource
    If oRs.State = adStateClosed Then
        oRs.CursorType = adOpenDynamic
        oRs.Open
    End If
    If oRs.RecordCount = 0 Then Exit Sub
    
    ' Now add the Grid details to the Table...
    Dim iCnt As Integer: iCnt = 1
    oRs.MoveFirst
    Do While Not oRs.EOF
        oValDict2.add iCnt, oRs.Fields(0).value
        oValDict1.add iCnt, oRs.Fields(1).value
        iCnt = iCnt + 1
        oRs.MoveNext
    Loop
    
    Dim f As frmChart
    Set f = New frmChart
    Call InitializeBarCharts(f, oValDict1, oValDict2, txtSnowPackName.Text, "Station number", "Elevation (ft)", "", "")
    f.Show vbModal
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Call LoadSnowPackNamesforForm
       
    Call LoadDefaults
    
End Sub

Private Sub LoadDefaults()
    
    On Error GoTo ShowError
        
    txtLeftBank.Text = "0.01"
    txtRightBank.Text = "0.01"
    txtChannel.Text = "0.01"
    txtLeft.Text = "1"
    txtRight.Text = "10"
    txtStations.Text = "0"
    txtElev.Text = "1000"
    
    ' Now load the Grid details...........
    Dim strStations As String, strElevations As String
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "Station", adInteger
    oRs.Fields.Append "Elevation", adInteger
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    '* Set datagrid value, header caption and width
    Set DataGridSnowPack.DataSource = oRs
    DataGridSnowPack.ColumnHeaders = True
    DataGridSnowPack.Columns(0).Caption = "Station (ft)"
    DataGridSnowPack.Columns(0).Width = DataGridSnowPack.Width / 2.2
    DataGridSnowPack.Columns(1).Caption = "Elevation (ft)"
    DataGridSnowPack.Columns(1).Width = DataGridSnowPack.Width / 2.2
    DataGridSnowPack.Refresh
   
    
    Exit Sub
ShowError:
    MsgBox "Load Grid Details :" & Err.description
    
End Sub

Private Sub LoadSnowPackNamesforForm()
On Error GoTo ShowError

    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadTransectNames
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


Public Sub InitializeBarCharts(ByVal f As Form, ByVal valueDict1 As Scripting.Dictionary, ByVal valueDict2 As Scripting.Dictionary, chartTitle As String, xAxisTitle As String, yAxisTitle As String, chart1Legend As String, chart2Legend As String)
  On Error GoTo ErrorHandler

    Dim myChart As MSChart
    Set myChart = f.MSChartBar
    
    Dim valueKey
    myChart.ColumnCount = 1
    myChart.RowCount = valueDict1.Count
    Dim nullflag As Integer
    Dim rowNum As Integer
    For rowNum = 1 To valueDict1.Count
        valueKey = valueDict1.keys(rowNum - 1)
        myChart.DataGrid.SetData rowNum, 1, valueDict1.Item(valueKey), nullflag
        myChart.Row = rowNum
        myChart.RowLabel = valueKey
    Next
   
    myChart.Plot.SeriesCollection(1).LegendText = chart1Legend
    
    'Set the bar gap
    myChart.Plot.BarGap = 0
    
    'Set chart legend properties
    myChart.Legend.VtFont.name = "Times New Roman"
    myChart.Legend.VtFont.Size = 7
    myChart.Legend.Location.LocationType = VtChLocationTypeTopRight
    
    
    'Set Chart title
    myChart.Title.Text = chartTitle
    myChart.Title.VtFont.Style = VtFontStyleBold
    myChart.Title.VtFont.Size = 9
    myChart.Title.VtFont.name = "Arial"
    
    'Set the x axis title
    myChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.Text = xAxisTitle
    myChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.VtFont.Size = 10
    myChart.Plot.Axis(VtChAxisIdX, 0).AxisTitle.VtFont.name = "Arial"
    'Set the y axis title
    myChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.Text = yAxisTitle
    myChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.VtFont.Size = 10
    myChart.Plot.Axis(VtChAxisIdY, 0).AxisTitle.VtFont.name = "Arial"
    myChart.Refresh

  Exit Sub
ErrorHandler:
  HandleError True, "InitializeBarCharts " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
End Sub

