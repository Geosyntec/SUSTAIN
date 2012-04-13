VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDataManage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Management"
   ClientHeight    =   6930
   ClientLeft      =   5160
   ClientTop       =   3285
   ClientWidth     =   7725
   Icon            =   "frmDataManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox SUBBASINR 
      Height          =   315
      Left            =   5040
      TabIndex        =   26
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox SUBBASIN 
      Height          =   315
      Left            =   5040
      TabIndex        =   25
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton findSTREAM 
      Height          =   340
      Left            =   6960
      Picture         =   "frmDataManage.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1545
      Width           =   600
   End
   Begin VB.ComboBox STREAM 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   21
      ToolTipText     =   "Select NHD layer from Map"
      Top             =   1560
      Width           =   3300
   End
   Begin VB.CommandButton findDemLayer 
      Height          =   340
      Left            =   6960
      Picture         =   "frmDataManage.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   600
   End
   Begin VB.ComboBox DEM 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   19
      ToolTipText     =   "Select DEM layer from Map"
      Top             =   135
      Width           =   3300
   End
   Begin VB.Frame Frame4 
      Caption         =   "BMP ET Calculation Options"
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   7455
      Begin VB.OptionButton optMonCons 
         Caption         =   "Constant Monthly"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optFromTS 
         Caption         =   "Daily ET from Timeseries"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton optCalculate 
         Caption         =   "Calculate Using Daily Temperature"
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox TSFILE 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton findTSFile 
         Height          =   340
         Left            =   6480
         Picture         =   "frmDataManage.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox LATITUDE 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGridET 
         Height          =   1455
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2566
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
      Begin VB.Label ETLabel 
         Caption         =   "Enter Monthly ET rate (in/day)"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label LabelTSFilePath 
         Caption         =   "Climate Time Series File Path"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label LabelLatitude 
         Caption         =   "Latitude (Decimal degrees)"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.ComboBox Landuse 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Select land use layer from Map"
      Top             =   615
      Width           =   3300
   End
   Begin VB.CommandButton findLuGrid 
      Height          =   340
      Left            =   6960
      Picture         =   "frmDataManage.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   600
   End
   Begin VB.ComboBox lulookup 
      Height          =   315
      ItemData        =   "frmDataManage.frx":0CD2
      Left            =   3600
      List            =   "frmDataManage.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select landuse lookup table from Map"
      Top             =   1080
      Width           =   3300
   End
   Begin VB.CommandButton findlulu 
      Height          =   340
      Left            =   6960
      Picture         =   "frmDataManage.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1065
      Width           =   600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   6480
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Select the Downstream ID (COM_ID2 or SUBBASINR):"
      Height          =   315
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Select the Stream ID (COM_ID or SUBBASIN):"
      Height          =   315
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Define Stream Feature Layer"
      Height          =   315
      Left            =   240
      TabIndex        =   24
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Define DEM Raster Layer (optional)"
      Height          =   315
      Left            =   240
      TabIndex        =   23
      Top             =   135
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Define Landuse Raster Layer"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   615
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Define Landuse Lookup Table"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "frmDataManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdOk_Click()
    Dim strError  As String
    Dim strWarning1 As String
    Dim strWarning2 As String
    strError = ""
    strWarning1 = ""
    strWarning2 = ""
    
    Dim pLanduseLayerName As String
    Dim pLanduseTableName As String
    Dim pTempFolderName As String
    Dim pDEMLayerName As String
    Dim pStreamLayerName As String
    
    Dim missTsError As Boolean
    Dim invalidETCoeff As String
    
    'Get error messages
    If (Landuse.Text = "") Then
        strError = strError & " Landuse raster layer. " & vbNewLine
    End If
    If (lulookup.Text = "") Then
        strError = strError & " Landuse lookup table " & vbNewLine
    End If
    
    'Check ET related fields
    
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridET.DataSource
    invalidETCoeff = ""
    oRs.MoveFirst
    Do Until oRs.EOF
        If Not IsNumeric(oRs.Fields(1).value) Then
            invalidETCoeff = invalidETCoeff & oRs.Fields(0).value & vbNewLine
        End If
        oRs.MoveNext
    Loop
    
    If (optFromTS.value Or optCalculate.value) Then
        If TSFILE.Text = "" Then
            strError = strError & " Climate time series file " & vbNewLine
        Else
            Dim fso As New FileSystemObject
            If Not fso.FileExists(Trim(TSFILE.Text)) Then
                missTsError = True
            End If
        End If
    End If
    If (optCalculate.value) Then
        If LATITUDE.Text = "" Then
            strError = strError & " Latitude value can not be empty" & vbNewLine
        ElseIf IsNumeric(Trim(LATITUDE.Text)) = False Then
            strError = strError & " Latitude value must be numeric" & vbNewLine
        ElseIf CDbl(LATITUDE.Text) > 90 Or CDbl(LATITUDE.Text) < -90 Then
            strError = strError & " Latitude value must be in [-90, 90]" & vbNewLine
        End If
    End If
    
    If (DEM.Text = "") Then
        strWarning1 = strWarning1 & " DEM raster Layer. " & vbNewLine
    End If
    If (STREAM.Text = "") Then
        strWarning2 = strWarning2 & " STREAM feature Layer. " & vbNewLine
    End If
    
    If (STREAM.Text <> "") Then
        gSUBBASINFieldName = SUBBASIN.Text
        gSUBBASINRFieldName = SUBBASINR.Text
    End If
    
    If (strError = "") Then
        pLanduseLayerName = Landuse.Text
        pLanduseTableName = lulookup.Text
    Else
        strError = "Define Following " & vbNewLine & strError
        MsgBox strError, vbExclamation
        Exit Sub
    End If
    
    If missTsError Then
        MsgBox " Climate time series file does not exists " & vbNewLine
        Exit Sub
    End If
    If invalidETCoeff <> "" Then
        MsgBox " ET coefficients for the following months are not numbers " & vbNewLine & invalidETCoeff
        Exit Sub
    End If
    
    If strWarning1 = "" Then
        pDEMLayerName = DEM.Text
    End If
    If (strWarning2 = "") Then
        pStreamLayerName = STREAM.Text
    End If
           
    If (gLayerNameDictionary Is Nothing) Then
        Set gLayerNameDictionary = CreateObject("Scripting.Dictionary")
    End If
    gLayerNameDictionary.Item("DEM") = pDEMLayerName: gDEMLayer = pDEMLayerName: RenderRasterLayer (gDEMLayer)
    gLayerNameDictionary.Item("STREAM") = pStreamLayerName: gStreamLayer = pStreamLayerName
    gLayerNameDictionary.Item("Landuse") = pLanduseLayerName: RenderLanduseSymbology pLanduseLayerName, pLanduseTableName
    gLayerNameDictionary.Item("lulookup") = pLanduseTableName
    gLayerNameDictionary.Item("TEMP") = gMapTempFolder
    gLayerNameDictionary.Item("SUBBASIN") = gSUBBASINFieldName
    gLayerNameDictionary.Item("SUBBASINR") = gSUBBASINRFieldName
    
    Dim etOption As Integer
    If optMonCons.value Then
        etOption = 0
    ElseIf optFromTS.value Then
        etOption = 1
    Else
        etOption = 2
    End If
    gLayerNameDictionary.Item("ETOPTION") = etOption
    If etOption <> 0 Then
        gLayerNameDictionary.Item("TSFILE") = Trim(TSFILE.Text)
        If etOption = 2 Then
            gLayerNameDictionary.Item("LATITUDE") = Trim(LATITUDE.Text)
        End If
    End If
        
    
    oRs.MoveFirst
    Do Until oRs.EOF
        gLayerNameDictionary.Item("MonET" & oRs.Fields(0).value) = oRs.Fields(1).value
        oRs.MoveNext
    Loop
    
    If (CheckInputDataProjection = False) Then
        FrmErrors.lblErrors.Caption = "Some of the datasets may not have projections defined or different projections. Please define same projection for all input datasets."
        FrmErrors.Show vbModal
        Exit Sub
    End If
'    If (CheckInputTableFieldFormats = False) Then
'        FrmErrors.lblErrors.Caption = "Some of the datasets may have missing fields. Please refer below for information on required fields for input datasets."
'        FrmErrors.Show vbModal
'        Exit Sub
'    End If
'
    'Write the gLayerNameDictionary information to src file
    Call WriteLayerTagDictionaryToSRCFile
    gDefLayers = True
    
    'Close the form
    Unload Me
''    'Render the Landuse rasterlayer and join to lulookup table
''    Call JoinAndRenderLanduseRasterLayer
    
End Sub

Public Sub RenderRasterLayer(strRasterLayer As String)
    
 
    '----------------change to stretched layer-------------------------
    Dim pStretchRen As IRasterStretchColorRampRenderer
    Set pStretchRen = New RasterStretchColorRampRenderer
    Dim pRasRen As IRasterRenderer
    Set pRasRen = pStretchRen
    
    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = GetInputRasterLayer(strRasterLayer)
    If pRasterLayer Is Nothing Then Exit Sub
    
    'Set raster for the renderer and update
    Set pRasRen.Raster = pRasterLayer.Raster
    pRasRen.Update
    
    'Define two colors
    Dim pFromColor As IColor
    Dim pToColor As IColor
    Set pFromColor = New RgbColor
    Set pToColor = New RgbColor
    pFromColor.RGB = RGB(30, 0, 150)
    pToColor.RGB = RGB(150, 0, 0)
    
    ' Create color ramp
    Dim pColorRamp As esriDisplay.IAlgorithmicColorRamp
    Set pColorRamp = New esriDisplay.AlgorithmicColorRamp
    pColorRamp.Size = 255
    pColorRamp.FromColor = pFromColor
    pColorRamp.ToColor = pToColor
    pColorRamp.CreateRamp True
    
    '## This renders the raster ##
    pStretchRen.ColorRamp = pColorRamp 'sets a custom colorramp
        
    pRasRen.Update
    Set pRasterLayer.Renderer = pRasRen
    pRasterLayer.Renderer.Update

    gMxDoc.UpdateContents

End Sub

Public Sub RenderLanduseSymbology(pLanduseLayerName As String, pLanduseTableName As String)
    On Error GoTo ShowError
    Dim NumOfValues As Integer
    
    Call SetLanduseColorDictionary
    
    Dim pLuRLayer As IRasterLayer
    Set pLuRLayer = GetInputRasterLayer(pLanduseLayerName)
    
    If pLuRLayer Is Nothing Then Err.Raise vbObjectError + 5002, , "Missing landuse layer. Validate Layer first"
    
    Dim pLuRaster As IRaster
    Set pLuRaster = pLuRLayer.Raster
    
    Dim LuDescDict As Scripting.Dictionary
    Set LuDescDict = GetLuLuDictionary(pLanduseTableName)
    If LuDescDict Is Nothing Then Exit Sub
    
    Dim pTable As iTable
    Dim pBand As esriDataSourcesRaster.IRasterBand ' IRasterBand
    Dim pBandCol As IRasterBandCollection
    Set pBandCol = pLuRaster
    Set pBand = pBandCol.Item(0)
    Dim TableExist As Boolean
    pBand.HasTable TableExist
    'TableExist = pBand.HasTable
    If Not TableExist Then Exit Sub
    Set pTable = pBand.AttributeTable
    
    'Get the number of rows from raster table
    NumOfValues = pTable.RowCount(Nothing)
  
   ' Specified a field and get the field index for the specified field to be rendered.
    Dim FieldIndex As Integer
    Dim fieldname As String
    fieldname = "Value" ' Value is the default field, you can specify other field here
    FieldIndex = pTable.FindField(fieldname)
  
  ' Create random color
    Dim pRamp As IRandomColorRamp
    Set pRamp = New RandomColorRamp
    pRamp.Size = NumOfValues
    pRamp.Seed = 100
    pRamp.CreateRamp (True)
    Dim pFSymbol As ISimpleFillSymbol
    Dim pHashlineSymbol As ILineFillSymbol
  
  ' Create UniqueValue renderer and QI RasterRenderer
    Dim pUVRen As IRasterUniqueValueRenderer
    Set pUVRen = New RasterUniqueValueRenderer
    Dim pRasRen As IRasterRenderer
    Set pRasRen = pUVRen
  
  ' Connect the renderer and the raster
    Set pRasRen.Raster = pLuRaster
    pRasRen.Update
  
  ' Set UniqueValue renderer
    pUVRen.HeadingCount = 1   ' Use one heading
    'pUVRen.Heading(0) = "All Data Values"
    pUVRen.ClassCount(0) = NumOfValues
    pUVRen.Field = fieldname
    Dim i As Long
    Dim pRow As iRow
    Dim LabelValue As Variant
    
    Set pFSymbol = New SimpleFillSymbol
    
    Set pHashlineSymbol = New LineFillSymbol
    pHashlineSymbol.Angle = 45
    pHashlineSymbol.Separation = 5
    pHashlineSymbol.Offset = 1
    
    'pMarkerSymbol.MarkerSymbol.Size = 1
    
    Dim pRGBColor As IRgbColor
    Set pRGBColor = New RgbColor
    
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    i = 0
    Do Until pRow Is Nothing
'    For i = 0 To NumOfValues - 1
'        Set pRow = pTable.GetRow(i)
        LabelValue = pRow.value(FieldIndex)
        If (LuDescDict.Exists(LabelValue)) Then
            pUVRen.AddValue 0, i, LabelValue 'Get a row from the table
            pUVRen.Label(0, i) = CStr(LuDescDict.Item(LabelValue))

            If gColorDict.Exists(CStr(LuDescDict.Item(LabelValue))) Then
                pRGBColor.RGB = FrmColors.colors.Item(gColorDict.Item(CStr(LuDescDict.Item(LabelValue)))(1)).BackColor

                If gColorDict.Item(CStr(LuDescDict.Item(LabelValue)))(0) Then
                    pFSymbol.Color = pRGBColor
                    pUVRen.Symbol(0, i) = pFSymbol
                Else
                    pHashlineSymbol.Color = pRGBColor
                    pUVRen.Symbol(0, i) = pHashlineSymbol
                End If

            Else
                pFSymbol.Color = pRamp.Color(i)
                pUVRen.Symbol(0, i) = pFSymbol
            End If

            'pFSymbol.Color = pRamp.Color(i)
            'pUVRen.Symbol(0, i) = pFSymbol  'Set symbol
''        Else
''            pUVRen.Label(0, i) = CStr(LabelValue)
        End If
'    Next i
        i = i + 1
        Set pRow = pCursor.NextRow
    Loop
  
    'Update render and refresh layer
    pRasRen.Update
    Set pLuRLayer.Renderer = pUVRen
    gMxDoc.ActiveView.Refresh
    gMxDoc.UpdateContents
    
    ' Clean up
    GoTo CleanUp
ShowError:
    MsgBox "Error in RenderLanduseSymbology :" & Err.description
CleanUp:
    Set pLuRLayer = Nothing
    Set pUVRen = Nothing
    Set pRasRen = Nothing
    Set pRamp = Nothing
    Set pFSymbol = Nothing
    Set pLuRaster = Nothing
    Set pBand = Nothing
    Set pBandCol = Nothing
    Set pTable = Nothing
    Set pRow = Nothing
End Sub


'******************************************************************************
'Subroutine: GetLuLuDictionary
'Author:     Sabu Paul
'Purpose:    Creates a dictionary for mapping land use description and lu code
'******************************************************************************

Public Function GetLuLuDictionary(pLanduseTableName As String) As Scripting.Dictionary
On Error GoTo ShowError
    Dim pLuLuDict As Scripting.Dictionary
    Set pLuLuDict = CreateObject("Scripting.Dictionary")
    
    Dim pLULookup As iTable
    Set pLULookup = GetInputDataTable(pLanduseTableName)
    
    If (pLULookup Is Nothing) Then Err.Raise vbObjectError + 5001, , "Validate layers first"
        
    Dim pCursor As ICursor
    Set pCursor = pLULookup.Search(Nothing, True)
    
    Dim pDescFldInd As Long
    pDescFldInd = pLULookup.FindField("LUName")
    
    Dim pIdFldInd As Long
    pIdFldInd = pLULookup.FindField("LUCODE")
        
    If pDescFldInd < 0 Then Err.Raise vbObjectError + 5001, , "No Description (LUName) field in land use lookup table"
    If pIdFldInd < 0 Then Err.Raise vbObjectError + 5001, , "No luCode field in land use lookup table"
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        If pLuLuDict.Exists(pRow.value(pIdFldInd)) Then
            Err.Raise vbObjectError + 5001, , "ID is not unique in land use lookup table"
        Else
            pLuLuDict.add pRow.value(pIdFldInd), pRow.value(pDescFldInd)
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    Set GetLuLuDictionary = pLuLuDict
    GoTo CleanUp
    
ShowError:
    MsgBox "Error in GetLuLuDictionary: " & Err.description
CleanUp:
    Set pLULookup = Nothing
    Set pCursor = Nothing
    Set pLuLuDict = Nothing
End Function
Public Sub SetLanduseColorDictionary()
  On Error GoTo ErrorHandler

    Set gColorDict = CreateObject("Scripting.Dictionary")
    gColorDict.add "Low Density Residential", Array(True, 0)
    gColorDict.add "Medium Density Residential", Array(True, 1)
    gColorDict.add "High Density Residential", Array(True, 2)
    gColorDict.add "Commercial", Array(True, 3)
    gColorDict.add "Industrial", Array(True, 4)
    gColorDict.add "Institutional", Array(True, 5)
    gColorDict.add "Open Urban Land", Array(True, 6)
    gColorDict.add "Cropland", Array(True, 7)
    gColorDict.add "Pasture", Array(True, 8)
    gColorDict.add "Orchards/Vine yard/Horticul", Array(True, 9)
    gColorDict.add "Urban Herbaceous", Array(True, 10)
    gColorDict.add "Deciduous Forest", Array(True, 11)
    gColorDict.add "Evergreen Forest", Array(True, 12)
    gColorDict.add "Mixed Forest", Array(True, 13)
    gColorDict.add "Brush", Array(True, 14)
    gColorDict.add "Water", Array(True, 15)
    gColorDict.add "Wetlands", Array(True, 16)
    gColorDict.add "Bare Ground", Array(True, 17)
    gColorDict.add "Extractive", Array(True, 18)
    gColorDict.add "Highway Corridors", Array(True, 19)
    gColorDict.add "Railroad Corridors", Array(True, 20)
    gColorDict.add "Agricultural Buildings", Array(True, 21)

  Exit Sub
ErrorHandler:
  HandleError True, "SetLanduseColorDictionary " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
End Sub

Private Sub findDemLayer_Click()
    Dim strDemLayer As String
    strDemLayer = SelectLayerOrTableData("Raster")
    If (Trim(strDemLayer) <> "") Then
        DEM.AddItem strDemLayer
        DEM.ListIndex = DEM.ListCount - 1
    End If
End Sub
Private Sub findLuGrid_Click()
    Dim strLuLayer As String
    strLuLayer = SelectLayerOrTableData("Raster")
    If (Trim(strLuLayer) <> "") Then
        Landuse.AddItem strLuLayer
        Landuse.ListIndex = Landuse.ListCount - 1
    End If
End Sub

Private Sub findlulu_Click()
    Dim strLuLuTable As String
    strLuLuTable = SelectLayerOrTableData("GeoTable")
    If (Trim(strLuLuTable) <> "") Then
        lulookup.AddItem strLuLuTable
        lulookup.ListIndex = lulookup.ListCount - 1
    End If
End Sub

Private Sub findSTREAM_Click()
    Dim strSTREAMLayer As String
    strSTREAMLayer = SelectLayerOrTableData("Feature")
    If (Trim(strSTREAMLayer) <> "") Then
        STREAM.AddItem strSTREAMLayer
        STREAM.ListIndex = STREAM.ListCount - 1
    End If
    

    If (Trim(STREAM.Text) <> "") Then
        Call InitializeStreamFields(STREAM.Text)
    End If

End Sub



Public Function GetWindowsDir() As String

    Dim sRet As String, lngRet As Long
    sRet = String$(MAX_PATH, 0)
    lngRet = GetWindowsDirectory(sRet, MAX_PATH)
    GetWindowsDir = Left(sRet, lngRet)
End Function

Private Sub InitializeLayers()

    On Error GoTo ShowError
    Dim pLayer As ILayer
    Dim pFLayer2 As IFeatureLayer2
    
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If pLayer.Valid Then
            If (TypeOf pLayer Is IFeatureLayer) Then
                Set pFLayer2 = pLayer
                If pFLayer2.ShapeType = esriGeometryLine Or pFLayer2.ShapeType = esriGeometryPolyline Then
                    STREAM.AddItem pLayer.name
                End If
            ElseIf (TypeOf pLayer Is IRasterLayer) Then
                DEM.AddItem pLayer.name
                Landuse.AddItem pLayer.name
            End If
        End If
    Next
    Dim pTabCollection As IStandaloneTableCollection
    Set pTabCollection = gMap
    ReDim TableNames(pTabCollection.StandaloneTableCount)
    For i = 0 To (pTabCollection.StandaloneTableCount - 1)
        lulookup.AddItem pTabCollection.StandaloneTable(i).name
    Next
    
    Exit Sub
    
ShowError:
    MsgBox "InitializeLayers: " & Err.description

CleanUp:
    Set pLayer = Nothing

End Sub


Private Sub InitializeDataSourcesForm()
On Error GoTo ShowError
    
    'Read all the layer names and load default layers
    Dim FeatureLayerNames() As String
    ReDim FeatureLayerNames(0)
    Dim TableNames() As String
    Dim RasterLayerNames() As String
    ReDim RasterLayerNames(0)
    Dim i As Integer
    Dim totalFlayers As Integer
    Dim totalRlayers As Integer
    
    Dim lineFLayerNames() As String
    ReDim lineFLayerNames(0)
    Dim totalLineFLayers As Integer
    Dim pFLayer2 As IFeatureLayer2
    
    Dim pLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If pLayer.Valid Then
            If (TypeOf pLayer Is IFeatureLayer) Then
                totalFlayers = UBound(FeatureLayerNames) + 1
                ReDim Preserve FeatureLayerNames(totalFlayers)
                FeatureLayerNames(totalFlayers) = pLayer.name
                Set pFLayer2 = pLayer
                If pFLayer2.ShapeType = esriGeometryLine Or pFLayer2.ShapeType = esriGeometryPolyline Then
                    totalLineFLayers = UBound(lineFLayerNames) + 1
                    ReDim Preserve lineFLayerNames(totalLineFLayers)
                    lineFLayerNames(totalLineFLayers) = pLayer.name
                End If
            ElseIf (TypeOf pLayer Is IRasterLayer) Then
                totalRlayers = UBound(RasterLayerNames) + 1
                ReDim Preserve RasterLayerNames(totalRlayers)
                RasterLayerNames(totalRlayers) = pLayer.name
            End If
        End If
    Next
    Dim pTabCollection As IStandaloneTableCollection
    Set pTabCollection = gMap
    ReDim TableNames(pTabCollection.StandaloneTableCount)
    For i = 0 To (pTabCollection.StandaloneTableCount - 1)
        TableNames(i + 1) = pTabCollection.StandaloneTable(i).name
    Next
       
    Dim pDEMLayerName As String
    Dim pLanduseLayerName As String
    Dim pStreamLayerName As String
    Dim pLanduseLookupTable As String
    
    If Not (gLayerNameDictionary Is Nothing) Then
        pDEMLayerName = gLayerNameDictionary.Item("DEM")
        pLanduseLayerName = gLayerNameDictionary.Item("Landuse")
        pStreamLayerName = gLayerNameDictionary.Item("STREAM")
        pLanduseLookupTable = gLayerNameDictionary.Item("lulookup")
    End If
    
    'Load these values, only if nothing is entered in them
    If (DEM.ListCount = 0) Then
        LoadDefaultLayerNames DEM, "DEM", RasterLayerNames, pDEMLayerName
    End If
    If (Landuse.ListCount = 0) Then
        LoadDefaultLayerNames Landuse, "Landuse", RasterLayerNames, pLanduseLayerName
    End If
    If (STREAM.ListCount = 0) Then
        'LoadDefaultLayerNames STREAM, "STREAM", FeatureLayerNames, pStreamLayerName
        LoadDefaultLayerNames STREAM, "STREAM", lineFLayerNames, pStreamLayerName
    End If
    If (lulookup.ListCount = 0) Then
        LoadDefaultLayerNames lulookup, "Lulookup", TableNames, pLanduseLookupTable
    End If
        
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = Replace(gApplication.Document, ".mxd", "")
    
    Exit Sub
    
ShowError:
    MsgBox "InitializeDataSourcesForm: " & Err.description

CleanUp:
    Set pLayer = Nothing
    Set pControl = Nothing
End Sub

Private Sub InitializeDataFromFile()
On Error GoTo ShowError

    '** Set null values for all controls if not value to initialize
    For Each pControl In Controls
        If (TypeOf pControl Is TextBox) Then
            pControl.Text = ""
        End If
        If (TypeOf pControl Is ComboBox) Then
            pControl.AddItem ""
        End If
    Next pControl
       
    If (gLayerNameDictionary Is Nothing) Then
        Exit Sub
    End If
    
    'Initialize from default values
    'InitializeDataSourcesForm
    InitializeLayers
    
    'Define variables for layer name dictionary access
    Dim pKeys
    pKeys = gLayerNameDictionary.keys
    Dim pLayerKey
    Dim ikey

    'Open datasources.txt file if present and initialize datasource -- Sabu Paul; Aug 24, 2004
    For Each pControl In Controls
        If (TypeOf pControl Is TextBox Or TypeOf pControl Is ComboBox) Then
            For ikey = 0 To gLayerNameDictionary.Count - 1
                pLayerKey = pKeys(ikey)
                If pControl.name = pLayerKey Then
                    If (TypeOf pControl Is TextBox) Then
                        pControl.Text = gLayerNameDictionary.Item(pLayerKey)
                    Else
                        pControl.Text = (gLayerNameDictionary.Item(pLayerKey))
                        'pControl.ListIndex = 0
                    End If
                    Exit For
                End If
            Next
        End If
    Next pControl
        
    
    'Read ET Related information
    
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "Month", adVarChar, 10
    oRs.Fields.Append "Value", adDouble
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    With DataGridET
    Set .DataSource = oRs
        .ColumnHeaders = True
        .Columns(0).Caption = "Month"
        .Columns(0).Locked = True
        .Columns(0).Visible = True
        .Columns(0).Width = 1500
        .Columns(1).Caption = "Value"
        .Columns(1).Locked = False
        .Columns(1).Visible = True
        .Columns(1).Width = 2000
    End With
    
    Dim etOption As Integer
    etOption = 0
    Dim monI As Integer
    Dim months
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    If gLayerNameDictionary.Exists("ETOPTION") Then
        etOption = gLayerNameDictionary.Item("ETOPTION")
        For i = 0 To 11
            oRs.AddNew
            oRs.Fields(0).value = months(i)
            oRs.Fields(1).value = gLayerNameDictionary.Item("MonET" & months(i))
        Next
    Else
        For i = 0 To 11
            oRs.AddNew
            oRs.Fields(0).value = months(i)
            oRs.Fields(1).value = 0.0055
        Next
    End If
    
    If etOption = 0 Then
        optMonCons.value = True
    Else
        'LabelTSFilePath.Enabled = True
        'TSFILE.Enabled = True
        'findTSFile.Enabled = True
        If etOption = 1 Then
            optFromTS.value = True
        Else
            optCalculate.value = True
            'LabelLatitude.Enabled = True
            'LATITUDE.Enabled = True
        End If
    End If

    
    GoTo CleanUp
ShowError:
    MsgBox "InitializeDataFromFile: " & Err.description
CleanUp:
    Set pDataLines = Nothing
    Set pDataSrcFile = Nothing
    Set fso = Nothing
    Set pControl = Nothing
End Sub
'*** Subroutine to load default names to list box controls
Private Sub LoadDefaultLayerNames(pControl As ComboBox, datasetname As String, LayerNames() As String, _
                                    strDefaultName As String)
    Dim i As Integer
    For i = 1 To (UBound(LayerNames))
        pControl.AddItem LayerNames(i)
        If (strDefaultName = "") Then
            If (Replace(UCase(LayerNames(i)), " ", "") = Replace(UCase(datasetname), " ", "")) Then
                 pControl.ListIndex = i - 1
            End If
        Else
            If (Replace(UCase(LayerNames(i)), " ", "") = Replace(UCase(strDefaultName), " ", "")) Then
                pControl.ListIndex = i - 1
            End If
        End If
    Next
End Sub



Private Sub findTSFile_Click()
    CommonDialog1.ShowOpen
    TSFILE.Text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    'Read data layer information from src file
    Call ReadLayerTagDictionaryToSRCFile
    
    'Check if the dataset define is present on map
    Call CheckDatasetExistenceOnMap
    
    'Initialize form parameters from file loaded dictionary names
    InitializeDataFromFile
    
    '** Resize the height of the form
    Dim pLayerStream As IFeatureLayer
    Set pLayerStream = GetInputFeatureLayer(STREAM.Text)
    Dim pTable As IDisplayTable
    Set pTable = pLayerStream
    If Not pLayerStream Is Nothing Then
        If pTable.DisplayTable.FindField(gSUBBASINFieldName) > -1 Then SUBBASIN.AddItem gSUBBASINFieldName: SUBBASIN.Text = gSUBBASINFieldName
        If pTable.DisplayTable.FindField(gSUBBASINRFieldName) > -1 Then SUBBASINR.AddItem gSUBBASINRFieldName: SUBBASINR.Text = gSUBBASINRFieldName
    End If
    
End Sub


Private Sub CheckDatasetExistenceOnMap()

    If (gLayerNameDictionary Is Nothing) Then
        Exit Sub
    End If
    
    'Check if any of the dataset(except temp folder) from map, remove from glayernamedictionary
    Dim pRLayer As IRasterLayer
    Set pRLayer = GetInputRasterLayer("Landuse")
    If (pRLayer Is Nothing And (gLayerNameDictionary.Exists("Landuse"))) Then
        gLayerNameDictionary.Remove ("Landuse")
    End If
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("lulookup")
    If (pTable Is Nothing And (gLayerNameDictionary.Exists("lulookup"))) Then
        gLayerNameDictionary.Remove ("lulookup")
    End If
    
    Dim pFlayer As IFeatureLayer
    Set pFlayer = GetInputFeatureLayer("STREAM")
    If ((pFlayer Is Nothing) And (gLayerNameDictionary.Exists("STREAM"))) Then
        gLayerNameDictionary.Remove ("STREAM")
    End If
    
    Set pRLayer = GetInputRasterLayer("DEM")
    If ((pRLayer Is Nothing) And (gLayerNameDictionary.Exists("DEM"))) Then
        gLayerNameDictionary.Remove ("DEM")
    End If
End Sub

Public Sub CheckLayerFromFilePresentInMap(ByVal pControl As ComboBox)
    'Read all the layer names and load default layers
    Dim pDataName As String
    pDataName = pControl.Text
    pControl.Text = ""
    Dim i As Integer
    For i = 0 To (gMap.LayerCount - 1)
        If (gMap.Layer(i).name = pDataName) Then
            pControl.Text = pDataName
            Exit Sub
        End If
    Next
    Dim pTabCollection As IStandaloneTableCollection
    Set pTabCollection = gMap
    For i = 0 To (pTabCollection.StandaloneTableCount - 1)
        If (pTabCollection.StandaloneTable(i).name = pDataName) Then
            pControl.Text = pDataName
            Exit Sub
        End If
    Next
End Sub

Private Sub optCalculate_Click()
    ETLabel.Caption = "Enter Monthly Variable Coefficient to Calculate ET Values"
    LabelTSFilePath.Enabled = True
    LabelLatitude.Enabled = True
    TSFILE.Enabled = True
    findTSFile.Enabled = True
    LATITUDE.Enabled = True
End Sub

Private Sub optFromTS_Click()
    ETLabel.Caption = "Enter Monthly Pan Coefficient"
    LabelTSFilePath.Enabled = True
    LabelLatitude.Enabled = False
    TSFILE.Enabled = True
    findTSFile.Enabled = True
    LATITUDE.Enabled = False
End Sub

Private Sub optMonCons_Click()
    ETLabel.Caption = "Enter Monthly ET rate (in/day)"
    LabelTSFilePath.Enabled = False
    LabelLatitude.Enabled = False
    TSFILE.Enabled = False
    findTSFile.Enabled = False
    LATITUDE.Enabled = False
End Sub

Private Sub STREAM_Change()
    If (Trim(STREAM.Text) <> "") Then
        Call InitializeStreamFields(STREAM.Text)
    End If
End Sub


Private Sub InitializeStreamFields(pStreamLayerName As String)

    Dim pSTREAMFLayer As IFeatureLayer
    Set pSTREAMFLayer = GetInputFeatureLayer(pStreamLayerName)
    If (pSTREAMFLayer Is Nothing) Then
        MsgBox "Stream feature layer not found."
        Exit Sub
    End If
    Dim pSTREAMFClass As IFeatureClass
    Set pSTREAMFClass = pSTREAMFLayer.FeatureClass
    
    Dim pTable As IDisplayTable
    Set pTable = pSTREAMFLayer
    
    Dim pFields As esriGeoDatabase.IFields
    Set pFields = pTable.DisplayTable.Fields
    Dim pField As esriGeoDatabase.IField
    Dim i As Integer
    Dim pSUBBASINindex As Integer
    Dim pSUBBASINRindex As Integer
    pSUBBASINindex = 0
    pSUBBASINRindex = 0
    SUBBASIN.Clear
    SUBBASINR.Clear
    For i = 0 To (pFields.FieldCount - 1)
      Set pField = pFields.Field(i)
      SUBBASIN.AddItem pField.name
      SUBBASINR.AddItem pField.name
      '**check if SUBBASIN/COM_ID field is present
      If (pField.name = "SUBBASIN" Or pField.name = "COM_ID") Then
        pSUBBASINindex = i
      End If
      '**check if SUBBASIN/COM_ID field is present
      If (pField.name = "SUBBASINR" Or pField.name = "COM_ID2") Then
        pSUBBASINRindex = i
      End If
    Next
    SUBBASIN.ListIndex = pSUBBASINindex
    SUBBASINR.ListIndex = pSUBBASINRindex
    
    '** Cleanup
    Set pField = Nothing
    Set pFields = Nothing
    Set pSTREAMFClass = Nothing
    Set pSTREAMFLayer = Nothing
End Sub

Private Sub STREAM_Click()
    If (Trim(STREAM.Text) <> "") Then
        Call InitializeStreamFields(STREAM.Text)
    End If
End Sub

