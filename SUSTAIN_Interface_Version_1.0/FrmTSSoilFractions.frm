VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmTSSoilFractions 
   Caption         =   "Soil Fractions for Time Series "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "FrmTSSoilFractions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGridSoilFracs 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Define soil fractions for pollutant time series"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmTSSoilFractions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ShowError
    
    Dim iR As Integer
    'Check the values.
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSoilFracs.DataSource
    
    Dim pRow As iRow
    oRs.MoveFirst

    Do Until oRs.EOF
        For iR = 1 To 3
            If oRs.Fields(iR).value < 0 Or oRs.Fields(iR).value > 1 Then
                MsgBox "Soil fractions should be between 0 and 1"
                Exit Sub
            End If
        Next
        
        oRs.MoveNext
    Loop

    'Create a DBF table to store the multipliers
    Dim pSoilFracTable As iTable
    Set pSoilFracTable = GetInputDataTable("TSSoilFractions")
        
    If pSoilFracTable Is Nothing Then
        Set pSoilFracTable = CreateTsMultipliersDBF("TSSoilFractions")
        AddTableToMap pSoilFracTable
    Else
        'Delete all the records from the table
        pSoilFracTable.DeleteSearchedRows Nothing
    End If
    
   
    Dim pTSindex As Long
    pTSindex = pSoilFracTable.FindField("TimeSeries")
    Dim pSandindex As Long
    pSandindex = pSoilFracTable.FindField("Sand")
    Dim pSiltindex As Long
    pSiltindex = pSoilFracTable.FindField("Silt")
    Dim pClayindex As Long
    pClayindex = pSoilFracTable.FindField("Clay")
        
    If pTSindex < 0 Or pSandindex < 0 Or pSiltindex < 0 Or pClayindex < 0 Then
        MsgBox "Required fields are missing in TSSoilFractions Table"
        Exit Sub
    End If

    
    oRs.MoveFirst
    Do Until oRs.EOF
        Set pRow = pSoilFracTable.CreateRow
        pRow.value(pTSindex) = oRs.Fields(1).value
        pRow.value(pSandindex) = oRs.Fields(1).value
        pRow.value(pSiltindex) = oRs.Fields(2).value
        pRow.value(pClayindex) = oRs.Fields(3).value
        pRow.Store
        oRs.MoveNext
    Loop
    Unload Me
    GoTo CleanUp

ShowError:
    MsgBox "Error updating soil fractions:" & Err.description
CleanUp:
    oRs.Close
    Set oRs = Nothing
    Set pSoilFracTable = Nothing
End Sub

Private Sub Form_Load()
    'Get the landuse reclassificatio table
    Dim pLUReclassTable As iTable
    Set pLUReclassTable = GetInputDataTable("LUReclass")
    If (pLUReclassTable Is Nothing) Then
        MsgBox "LUReclass table is missing."
        Exit Sub
    End If
    
    'Initialize the DB grid with the values
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "TSFILE", adVarChar, 50
    oRs.Fields.Append "Sand", adDouble
    oRs.Fields.Append "Silt", adDouble
    oRs.Fields.Append "Clay", adDouble
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    Dim iR As Integer
    
    'Check to
    Dim pSoilFracTable As iTable
    Set pSoilFracTable = GetInputDataTable("TSSoilFractions")
       
    Dim soilFDict As Scripting.Dictionary
    Set soilFDict = New Scripting.Dictionary
    
    Dim pRow As iRow
    
    'Add all pollutants and their multipliers
    If Not pSoilFracTable Is Nothing Then
        Dim pCursor As esriGeoDatabase.ICursor
        Set pCursor = pSoilFracTable.Search(Nothing, False)
        
        Dim pTSindex As Long
        pTSindex = pSoilFracTable.FindField("TimeSeries")
        Dim pSandindex As Long
        pSandindex = pSoilFracTable.FindField("Sand")
        Dim pSiltindex As Long
        pSiltindex = pSoilFracTable.FindField("Silt")
        Dim pClayindex As Long
        pClayindex = pSoilFracTable.FindField("Clay")
                
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            soilFDict.Add pRow.value(pTSindex), Array(pRow.value(pSandindex), pRow.value(pSiltindex), pRow.value(pClayindex))
            Set pRow = pCursor.NextRow
        Loop
    End If
    
    Dim pLuCursor As esriGeoDatabase.ICursor
    Set pLuCursor = pLUReclassTable.Search(Nothing, False)
    
    Dim pTSFileindex As Long
    pTSFileindex = pLUReclassTable.FindField("TimeSeries")
        
    Set pRow = pLuCursor.NextRow
    Do While Not pRow Is Nothing
        oRs.AddNew
        oRs.Fields(0).value = pRow.value(pTSFileindex)
        If soilFDict.Exists(pRow.value(pTSFileindex)) Then
            oRs.Fields(1).value = soilFDict.Item(pRow.value(pTSFileindex))(0)
            oRs.Fields(2).value = soilFDict.Item(pRow.value(pTSFileindex))(1)
            oRs.Fields(3).value = soilFDict.Item(pRow.value(pTSFileindex))(2)
        Else
            oRs.Fields(1).value = 0
            oRs.Fields(2).value = 0
            oRs.Fields(3).value = 0
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    '* Set datagrid value, header caption and width
    Set DataGridSoilFracs.DataSource = oRs
    DataGridSoilFracs.ColumnHeaders = True
    DataGridSoilFracs.Columns(0).Caption = "Timeseries File"
    DataGridSoilFracs.Columns(0).Locked = True
    DataGridSoilFracs.Columns(0).Width = 2400
    DataGridSoilFracs.Columns(1).Caption = "Sand Fraction"
    DataGridSoilFracs.Columns(1).Width = 1300
    DataGridSoilFracs.Columns(2).Caption = "Silt Fraction"
    DataGridSoilFracs.Columns(2).Width = 1300
    DataGridSoilFracs.Columns(3).Caption = "Clay Fraction"
    DataGridSoilFracs.Columns(3).Width = 1300
End Sub
