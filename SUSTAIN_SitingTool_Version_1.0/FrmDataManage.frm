VERSION 5.00
Begin VB.Form FrmDataManage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Data Path"
   ClientHeight    =   2985
   ClientLeft      =   6825
   ClientTop       =   6615
   ClientWidth     =   6150
   Icon            =   "FrmDataManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtDEMpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ComboBox txtStreampath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ComboBox txtLandusepath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ComboBox txtRoadpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ComboBox txtSoilpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.ComboBox txtStreampath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   19
      Top             =   2040
      Width           =   3000
   End
   Begin VB.ComboBox txtLandusepath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   18
      Top             =   1680
      Width           =   3000
   End
   Begin VB.ComboBox txtRoadpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Top             =   1320
      Width           =   3000
   End
   Begin VB.ComboBox txtSoilpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   16
      Top             =   960
      Width           =   3000
   End
   Begin VB.CommandButton cmdBrowseStream 
      Height          =   350
      Left            =   5520
      Picture         =   "FrmDataManage.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Browse to select a directory"
      Top             =   2040
      Width           =   500
   End
   Begin VB.CommandButton cmdBrowselanduse 
      Height          =   350
      Left            =   5520
      Picture         =   "FrmDataManage.frx":081E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Browse to select a directory"
      Top             =   1680
      Width           =   500
   End
   Begin VB.CommandButton cmdBrowseRoad 
      Height          =   350
      Left            =   5520
      Picture         =   "FrmDataManage.frx":1030
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Browse to select a directory"
      Top             =   1320
      Width           =   500
   End
   Begin VB.CommandButton cmdBrowseSoil 
      Height          =   350
      Left            =   5520
      Picture         =   "FrmDataManage.frx":1842
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Browse to select a directory"
      Top             =   960
      Width           =   500
   End
   Begin VB.CommandButton cmdBrowseDEM 
      Height          =   350
      Left            =   5520
      Picture         =   "FrmDataManage.frx":2054
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Browse to select a directory"
      Top             =   580
      Width           =   500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox txtDEMpath 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   15
      Top             =   600
      Width           =   3000
   End
   Begin VB.Label Label7 
      Caption         =   "Select Stream shapefile"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Select UrbanLanduse shapefile"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Select Road shapefile"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select Soil shapefile"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Select the DEM grid"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblHelp 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Help"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label4 
      Caption         =   "You can use the right browse buttons to browse and select the Datasets."
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   5535
   End
End
Attribute VB_Name = "FrmDataManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmDataManage.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
' Enumerator for the DatasetType....
Private Enum datasetType
    dtRaster = 1
    dtFeature = 2
End Enum
'Private Variables.....
Private m_DEMFlag As Boolean
Private m_SoilFlag As Boolean
Private m_RoadFlag As Boolean
Private m_LanduseFlag As Boolean
Private m_StreamFlag As Boolean
' dictionary for the Layers....
Private m_LayerDict As Scripting.Dictionary



Private Sub cmdBrowseDEM_Click()

    On Error GoTo ErrorHandler
    If Browse_Dataset(dtRaster, txtDEMpath(0)) Then m_DEMFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseDEM_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Function Browse_Dataset(pdatasetType As datasetType, ByRef pControl As ComboBox) As Boolean

    On Error GoTo ErrorHandler
    Dim pDlg As IGxDialog
      Dim pGXSelect As IEnumGxObject
      Dim pGxObject As IGxObject
      Dim pGXDataset As IGxDataset
      Dim pFeatCls As IFeatureClass
      Dim pFeatLyr As IFeatureLayer
      Dim className As String
      Dim pObjectFilter As IGxObjectFilter
      Dim i As Long
      Dim pActiveView As IActiveView
      Set pActiveView = gMap
      Browse_Dataset = False
      ' set up filters on the files that will be browsed
      Set pDlg = New GxDialog
      If (pdatasetType = dtRaster) Then
        Set pObjectFilter = New GxFilterRasterDatasets
      ElseIf (pdatasetType = dtFeature) Then
        Set pObjectFilter = New GxFilterShapefiles
      End If
    
      pDlg.AllowMultiSelect = False
      pDlg.Title = "Select Data"
      Set pDlg.ObjectFilter = pObjectFilter
    
      If (pDlg.DoModalOpen(pActiveView.ScreenDisplay.hwnd, pGXSelect) = False) Then Exit Function
    
        ' got a valid selection from the GX Dialog, now extract the feature classes datasets etc.
        ' loop through the selection enumeration
        pGXSelect.Reset
        Set pGxObject = pGXSelect.Next
        
         If (Not pGxObject Is Nothing) Then
                ' We could be handed objects of various types, work out what types we have been handed and then open
                ' them up and add a feature layer to handle them
                Set pGXDataset = pGxObject
                If (TypeOf pGxObject Is IGxDataset) Then
                  If (pGXDataset.Type = esriDTFeatureClass) Then
                        Set pFeatCls = pGXDataset.Dataset
                        If pFeatCls.FeatureType = esriFTSimple Then
                          Set pFeatLyr = New FeatureLayer
                          Set pFeatLyr.FeatureClass = pFeatCls
                          pFeatLyr.Name = pFeatCls.AliasName
                          pFeatLyr.Visible = False
                          gMap.AddLayer pFeatLyr
                          Browse_Dataset = True
                          If pControl.Style = ComboBoxConstants.vbComboDropdownList Then
                            pControl.AddItem pGxObject.Name
                          End If
                          pControl.Text = pGxObject.Name
                        End If
                    ElseIf (pGXDataset.Type = esriDTRasterDataset) Then
                            Dim pRasterLayer As IRasterLayer
                            Set pRasterLayer = New RasterLayer
                            pRasterLayer.CreateFromDataset pGXDataset.Dataset
                            gMap.AddLayer pRasterLayer
                            gMxDoc.ActiveView.Refresh
                            Browse_Dataset = True
                            If pControl.Style = ComboBoxConstants.vbComboDropdownList Then
                                pControl.AddItem pGxObject.Name
                            End If
                            pControl.Text = pGxObject.Name
                    End If
                End If
            End If

    
  Exit Function
ErrorHandler:
  HandleError True, "Browse_Dataset " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function

Private Sub cmdBrowselanduse_Click()

    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtLandusepath(0)) Then m_LanduseFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowselanduse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Sub cmdBrowseRoad_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtRoadpath(0)) Then m_RoadFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseRoad_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdBrowseSoil_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtSoilpath(0)) Then m_SoilFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseSoil_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdBrowseStream_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtStreampath(0)) Then m_StreamFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseStream_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub



Private Sub cmdOK_Click()
  
  On Error GoTo ErrorHandler
  
   'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = DefineApplicationPath
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = gApplication.Document
    dataSrcFN = Replace(dataSrcFN, ".mxd", ".src")
    dataSrcFN = pAppPath & dataSrcFN
        
    Dim pDataSrcFile
    Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForWriting, True, TristateUseDefault)
   
   
   ' Store into Global variables...
   If txtDEMpath(0).Text = "" Then
        gDEMdata = txtDEMpath(1).Text
        pDataSrcFile.WriteLine "gDEMdata" & vbTab & gDEMdata
   Else
        gDEMdata = txtDEMpath(0).Text
        pDataSrcFile.WriteLine "gDEMdata" & vbTab & gDEMdata
   End If
   If gDEMdata = "" Then gDEMdata = "Not Available"
   If txtLandusepath(0).Text = "" Then
        gLandusedata = txtLandusepath(1).Text
        pDataSrcFile.WriteLine "gLandusedata" & vbTab & gLandusedata
   Else
        gLandusedata = txtLandusepath(0).Text
        pDataSrcFile.WriteLine "gLandusedata" & vbTab & gLandusedata
   End If
   If gLandusedata = "" Then gLandusedata = "Not Available"
   If txtRoadpath(0).Text = "" Then
        gRoaddata = txtRoadpath(1).Text
        pDataSrcFile.WriteLine "gRoaddata" & vbTab & gRoaddata
   Else
        gRoaddata = txtRoadpath(0).Text
        pDataSrcFile.WriteLine "gRoaddata" & vbTab & gRoaddata
   End If
   If gRoaddata = "" Then gRoaddata = "Not Available"
   If txtSoilpath(0).Text = "" Then
        gSoildata = txtSoilpath(1).Text
        pDataSrcFile.WriteLine "gSoildata" & vbTab & gSoildata
   Else
        gSoildata = txtSoilpath(0).Text
        pDataSrcFile.WriteLine "gSoildata" & vbTab & gSoildata
   End If
   If gSoildata = "" Then gSoildata = "Not Available"
   If txtStreampath(0).Text = "" Then
        gWaterdata = txtStreampath(1).Text
        pDataSrcFile.WriteLine "gWaterdata" & vbTab & gWaterdata
   Else
        gWaterdata = txtStreampath(0).Text
        pDataSrcFile.WriteLine "gWaterdata" & vbTab & gWaterdata
   End If
   If gWaterdata = "" Then gWaterdata = "Not Available"
   
   'Validation on Each Layer.......
   
   
       
   
   
    ' Close the form..........
    Unload Me
    pDataSrcFile.Close
    gDataValid = True
    
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub Form_Load()

    On Error GoTo ErrorHandler
   
   ' Create the Layer Dictionary.....
   Set m_LayerDict = CreateObject("Scripting.Dictionary")
   m_LayerDict.RemoveAll
      
    'If the map has subwatershed layer, remove it
    Dim i As Integer
    Dim pLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If TypeOf pLayer Is IFeatureLayer Then
            txtLandusepath(1).AddItem pLayer.Name
            txtRoadpath(1).AddItem pLayer.Name
            txtSoilpath(1).AddItem pLayer.Name
            txtStreampath(1).AddItem pLayer.Name
            txtLandusepath(1).Visible = True
            txtRoadpath(1).Visible = True
            txtSoilpath(1).Visible = True
            txtStreampath(1).Visible = True
        ElseIf TypeOf pLayer Is IRasterLayer Then
            txtDEMpath(1).AddItem pLayer.Name
            txtDEMpath(1).Visible = True
        End If
        m_LayerDict.Add pLayer.Name, pLayer
    Next
    
    ' Now validate the layers and add to the Form......
    If gDEMdata <> "" Then
        If m_LayerDict.Exists(gDEMdata) And Not GetInputFeatureLayer(gDEMdata) Is Nothing Then
            txtDEMpath(1).Text = gDEMdata
        End If
    End If
    If gLandusedata <> "" Then
        If m_LayerDict.Exists(gLandusedata) And Not GetInputFeatureLayer(gLandusedata) Is Nothing Then
            txtLandusepath(1).Text = gLandusedata
        End If
    End If
    If gRoaddata <> "" Then
        If m_LayerDict.Exists(gRoaddata) And Not GetInputFeatureLayer(gRoaddata) Is Nothing Then
            txtRoadpath(1).Text = gRoaddata
        End If
    End If
    If gSoildata <> "" Then
        If m_LayerDict.Exists(gSoildata) And Not GetInputFeatureLayer(gSoildata) Is Nothing Then
            txtSoilpath(1).Text = gSoildata
        End If
    End If
    If gWaterdata <> "" Then
        If m_LayerDict.Exists(gWaterdata) And Not GetInputFeatureLayer(gWaterdata) Is Nothing Then
            txtStreampath(1).Text = gWaterdata
        End If
    End If
    
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub


Private Sub txtDEMpath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtDEMpath(1).Text Or txtLandusepath(1).Text = txtDEMpath(1).Text Or txtRoadpath(1).Text = txtDEMpath(1).Text Or txtStreampath(1).Text = txtDEMpath(1).Text Then
        txtDEMpath(1).ListIndex = -1
    End If
        
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "txtDEMpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        

End Sub

Private Sub txtLandusepath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtLandusepath(1).Text Or txtLandusepath(1).Text = txtDEMpath(1).Text Or txtRoadpath(1).Text = txtLandusepath(1).Text Or txtStreampath(1).Text = txtLandusepath(1).Text Then
        txtLandusepath(1).ListIndex = -1
    End If
        
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "txtLandusepath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtRoadpath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtRoadpath(1).Text Or txtLandusepath(1).Text = txtRoadpath(1).Text Or txtRoadpath(1).Text = txtDEMpath(1).Text Or txtStreampath(1).Text = txtRoadpath(1).Text Then
        txtRoadpath(1).ListIndex = -1
    End If
        
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "txtRoadpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtSoilpath_Click(Index As Integer)

    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtDEMpath(1).Text Or txtLandusepath(1).Text = txtSoilpath(1).Text Or txtRoadpath(1).Text = txtSoilpath(1).Text Or txtStreampath(1).Text = txtSoilpath(1).Text Then
        txtSoilpath(1).ListIndex = -1
    End If
        
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "txtSoilpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Sub txtStreampath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtStreampath(1).Text Or txtDEMpath(1).Text = txtStreampath(1).Text Or txtRoadpath(1).Text = txtStreampath(1).Text Or txtStreampath(1).Text = txtLandusepath(1).Text Then
        txtStreampath(1).ListIndex = -1
    End If
        
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "txtStreampath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub
