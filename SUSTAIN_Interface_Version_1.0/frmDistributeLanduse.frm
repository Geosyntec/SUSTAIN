VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDistributeLanduse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribute Landuse for Aggregate BMPs"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "frmDistributeLanduse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   360
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   360
      Left            =   6240
      TabIndex        =   3
      Top             =   240
      Width           =   960
   End
   Begin VB.ComboBox cmbAggregate 
      Height          =   315
      ItemData        =   "frmDistributeLanduse.frx":08CA
      Left            =   1560
      List            =   "frmDistributeLanduse.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2400
   End
   Begin MSDataGridLib.DataGrid DataGridAgg 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8705
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
   Begin VB.Label Label1 
      Caption         =   "Select Aggregate"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   285
      Width           =   1320
   End
End
Attribute VB_Name = "frmDistributeLanduse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pBmpIdDictionary As Scripting.Dictionary
Private pBmpCatDictionary As Scripting.Dictionary
Private pLUReclassDict As Scripting.Dictionary
Private pModifyFlag As Boolean




Private Sub cmbAggregate_Click()
    
    If pModifyFlag Then
        If MsgBox("There are changes that are not saved.  Do you want to continue?", vbInformation + vbYesNo, "SUSTAIN") = vbNo Then
            cmbAggregate.Text = ""
            Exit Sub
        End If
    End If
    
    ' *****************************************
    If (gSubWaterLandUseDict Is Nothing) Then
        Call FindAndConvertWatershedFeatureLayerToRaster
        Call ComputeLanduseAreaForEachSubBasin
    End If
    
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "Landuse", adVarChar, 50
    oRs.Fields.Append "Interception", adInteger
    oRs.Fields.Append "Treatment", adInteger
    oRs.Fields.Append "Routing", adInteger
    oRs.Fields.Append "Storage", adInteger
    oRs.Fields.Append "Outlet", adInteger
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    ' Get the Watershed ID...
    Dim SubCatchmentID As Long
    SubCatchmentID = cmbAggregate.Text
    
    Dim bmpId As Long
    bmpId = cmbAggregate.Text
        
    Dim pLanduseDictionary As Scripting.Dictionary
    Set pLanduseDictionary = gSubWaterLandUseDict.Item(SubCatchmentID)
    
    Set pBmpCatDictionary = Get_BMP_Categories(pBmpIdDictionary.Item(bmpId))
    With DataGridAgg
        If Not pBmpCatDictionary.Exists("On-Site Interception") Then .Columns(1).Locked = True
        If Not pBmpCatDictionary.Exists("On-Site Treatment") Then .Columns(2).Locked = True
        If Not pBmpCatDictionary.Exists("Routing Attenuation") Then .Columns(3).Locked = True
        If Not pBmpCatDictionary.Exists("Regional Storage/Treatment") Then .Columns(4).Locked = True
    End With
    
    ' *********************************************************
    Set pBmpCatDictionary = CreateObject("Scripting.Dictionary")
    Dim pKeys
    pKeys = pLanduseDictionary.keys
    Dim pkey As String
    Dim ikey As Integer
    For ikey = 0 To pLanduseDictionary.Count - 1
        pkey = pKeys(ikey)
        If pBmpCatDictionary.Exists(pLUReclassDict.Item(pkey)) Then
            pBmpCatDictionary.Item(pLUReclassDict.Item(pkey)) = pLanduseDictionary.Item(pkey) + pBmpCatDictionary.Item(pLUReclassDict.Item(pkey))
        Else
            pBmpCatDictionary.Add pLUReclassDict.Item(pkey), pLanduseDictionary.Item(pkey)
        End If
    Next
    
    ' Now add the reocrds to the Datagrid.....
    pKeys = pBmpCatDictionary.keys
    For ikey = 0 To pBmpCatDictionary.Count - 1
        pkey = pKeys(ikey)
        oRs.AddNew
        oRs.Fields(0).value = pkey
        oRs.Fields(5).value = pBmpCatDictionary.Item(pkey)
    Next

    
    pModifyFlag = False
    
End Sub


Private Function Get_BMP_Categories(ByVal strBMP As String) As Scripting.Dictionary
    
    Set Get_BMP_Categories = CreateObject("Scripting.Dictionary")
        
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
        
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    pQueryFilter.WhereClause = "PropName='Type' And PropValue='" & strBMP & "'"
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow

    Dim pQueryFilter2 As IQueryFilter
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Do While Not pRow Is Nothing
        Set pQueryFilter2 = New QueryFilter
        pQueryFilter2.WhereClause = "PropName='Category' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then Get_BMP_Categories.Add pRow2.value(pNameIndex), pRow2.value(pNameIndex)
        Set pRow = pCursor.NextRow
    Loop

End Function




Private Sub cmdSave_Click()
    
    pModifyFlag = False
    
End Sub

Private Sub DataGridAgg_AfterUpdate()
    pModifyFlag = True
End Sub


Private Sub Form_Load()
    
    Dim curBMPType As String
    curBMPType = gNewBMPType
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    Set pLUReclassDict = CreateObject("Scripting.Dictionary")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "PropName='Type' And PropValue LIKE '%Aggregate%' ORDER BY PropValue"
    Dim pCursor As ICursor
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    
    Dim pRow As iRow
    cmbAggregate.Clear
    
    If pSelRowCount > 0 Then
        Do
            Set pRow = pCursor.NextRow
            If Not (pRow Is Nothing) Then
                If Not pBmpIdDictionary.Exists(pRow.value(pNameIndex)) Then
                    cmbAggregate.AddItem pRow.value(pNameIndex)
                    pBmpIdDictionary.Add pRow.value(pNameIndex), pRow.value(pIDindex)
                End If
            End If
        Loop Until (pRow Is Nothing)
        cmbAggregate.ListIndex = 0
    Else
        MsgBox "no bmp of this type"
    End If
    
    With DataGridAgg
            .ColumnHeaders = True
            .Columns(0).Caption = "Landuse"
            .Columns(0).Locked = True
            .Columns(0).Visible = True
            .Columns(0).Width = 1500
            .Columns(1).Caption = "Interception (%)"
            .Columns(1).Locked = False
            .Columns(1).Visible = True
            .Columns(1).Width = 2000
            .Columns(2).Caption = "Treatment (%)"
            .Columns(2).Locked = False
            .Columns(2).Visible = True
            .Columns(2).Width = 2000
            .Columns(3).Caption = "Routing (%)"
            .Columns(3).Locked = False
            .Columns(3).Visible = True
            .Columns(3).Width = 2000
            .Columns(4).Caption = "Storage (%)"
            .Columns(4).Locked = False
            .Columns(4).Visible = True
            .Columns(4).Width = 2000
            .Columns(5).Caption = "Outlet (%)"
            .Columns(5).Locked = True
            .Columns(5).Visible = True
            .Columns(5).Width = 2000
    End With
    
    ' Now Get the ID from the BMPs Layer....
    Dim pBMPLayer As iTable
    Set pBMPLayer = GetInputDataTable("BMPDetail")
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    Dim pQueryFilter2 As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName='Type'"
    Set pCursor = pBMPLayer.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
         pBmpIdDictionary.Add pRow.value(pBMPLayer.FindField("ID")), pRow.value(pBMPLayer.FindField("PropValue"))
    Loop
        
    ' Get the Landuse Codes....
    Set pBMPLayer = GetInputDataTable("LUReclass")
    Set pCursor = pBMPLayer.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
         pLUReclassDict.Add pRow.value(pBMPLayer.FindField("LUCode")), pRow.value(pBMPLayer.FindField("LUGroup"))
        Set pRow = pCursor.NextRow
    Loop
    
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    
    
End Sub
