VERSION 5.00
Begin VB.Form FrmImportWatershed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Watershed"
   ClientHeight    =   1305
   ClientLeft      =   7125
   ClientTop       =   6270
   ClientWidth     =   5250
   Icon            =   "FrmImportWatershed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5250
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton findWATERSHED 
      Height          =   340
      Left            =   4440
      Picture         =   "FrmImportWatershed.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   600
   End
   Begin VB.ComboBox WATERSHED 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select the watershed to import"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmImportWatershed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'On Select close the dialog box
    If (WATERSHED.Text <> "") Then
        'Add watershed layer in input file
        gLayerNameDictionary.Item("Watershed") = WATERSHED.Text
        'Call the subroutine to write this layer into src file
        Call WriteLayerTagDictionaryToSRCFile
                
        Dim pFeatureLayer As IFeatureLayer
        Set pFeatureLayer = GetInputFeatureLayer("Watershed")
        If (Not pFeatureLayer Is Nothing) Then
            If (CheckInputDataProjection = True) Then
                RenderWatershedLayer pFeatureLayer
                RenumberWatershedFeatures
                gManualDelineationFlag = True   'Set the flag to enabled manual delineation tools
            Else
                DeleteLayerFromMap ("Watershed")
                MsgBox "Projection of imported watershed does not match input dataset projection.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub findWATERSHED_Click()
    
    Dim strWatershedLayer As String
    strWatershedLayer = SelectLayerOrTableData("Feature")
    If (Trim(strWatershedLayer) <> "") Then
        WATERSHED.AddItem strWatershedLayer
        WATERSHED.ListIndex = WATERSHED.ListCount - 1
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Dim pLayer As ILayer
    Dim pFeatureLayer As IFeatureLayer
    Dim i As Integer
    For i = 0 To gMap.LayerCount - 1
        Set pLayer = gMap.Layer(i)
        If (TypeOf pLayer Is IFeatureLayer And pLayer.Valid) Then
            Set pFeatureLayer = pLayer
            If (pFeatureLayer.FeatureClass.ShapeType = esriGeometryPolygon) Then
                WATERSHED.AddItem pLayer.name
            End If
        End If
    Next
    If (WATERSHED.ListCount > 0) Then WATERSHED.ListIndex = 0

    For i = 0 To WATERSHED.ListCount - 1
        If gLayerNameDictionary.Item("Watershed") = Trim(WATERSHED.List(i)) Then
            WATERSHED.ListIndex = i
            Exit For
        End If
    Next
End Sub
