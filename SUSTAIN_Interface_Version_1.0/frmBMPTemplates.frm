VERSION 5.00
Begin VB.Form frmBMPTemplates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define BMP Templates"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBMPTemplates.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "BMP Templates"
      Height          =   7320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6600
      Begin VB.PictureBox imgPorousPave 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   4755
         Picture         =   "frmBMPTemplates.frx":08CA
         ScaleHeight     =   1575
         ScaleWidth      =   1470
         TabIndex        =   12
         Top             =   5400
         Width           =   1500
      End
      Begin VB.OptionButton OptionPorousPave 
         Caption         =   "Porous Pavement"
         Height          =   450
         Left            =   4755
         TabIndex        =   11
         Top             =   4920
         Width           =   1785
      End
      Begin VB.OptionButton OptionBioRB 
         Caption         =   "Bioretention Basin"
         Height          =   450
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1890
      End
      Begin VB.OptionButton OptionDryPond 
         Caption         =   "Dry Pond"
         Height          =   450
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   1530
      End
      Begin VB.OptionButton OptionRainB 
         Caption         =   "Rain Barrels"
         Height          =   450
         Left            =   4755
         TabIndex        =   2
         Top             =   240
         Width           =   1530
      End
      Begin VB.OptionButton OptionCistern 
         Caption         =   "Cistern"
         Height          =   450
         Left            =   270
         TabIndex        =   3
         Top             =   2520
         Width           =   1800
      End
      Begin VB.OptionButton OptionWetPond 
         Caption         =   "Wet Pond / Wet Land"
         Height          =   450
         Left            =   2520
         TabIndex        =   4
         Top             =   2520
         Width           =   1860
      End
      Begin VB.OptionButton OptionInfilTrench 
         Caption         =   "Infiltration Trench"
         Height          =   450
         Left            =   4755
         TabIndex        =   5
         Top             =   2520
         Width           =   1800
      End
      Begin VB.OptionButton OptionVegSwale 
         Caption         =   "Vegetative Swale"
         Height          =   450
         Left            =   270
         TabIndex        =   6
         Top             =   4920
         Width           =   1860
      End
      Begin VB.OptionButton OptionGreenRoof 
         Caption         =   "Green Roof"
         Height          =   450
         Left            =   2520
         TabIndex        =   7
         Top             =   4920
         Width           =   1785
      End
      Begin VB.Image imgBioRet 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   390
         Picture         =   "frmBMPTemplates.frx":2846
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1500
      End
      Begin VB.Image imgRainB 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   4800
         Picture         =   "frmBMPTemplates.frx":1E398
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1500
      End
      Begin VB.Image imgDryPond 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   2640
         Picture         =   "frmBMPTemplates.frx":21F5C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1500
      End
      Begin VB.Image imgCistern 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   270
         Picture         =   "frmBMPTemplates.frx":2E81C
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image imgGreenRoof 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   2520
         Picture         =   "frmBMPTemplates.frx":2FC59
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Image imgWetPond 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   2520
         Picture         =   "frmBMPTemplates.frx":34C80
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Image imgVegSwale 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   270
         Picture         =   "frmBMPTemplates.frx":3F2DA
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   1500
      End
      Begin VB.Image imgInfiltT 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1605
         Left            =   4755
         Picture         =   "frmBMPTemplates.frx":4C54C
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdSetDim 
      Caption         =   "Set Dimensions"
      Height          =   480
      Left            =   3960
      TabIndex        =   9
      Top             =   7560
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   5640
      TabIndex        =   10
      Top             =   7560
      Width           =   1080
   End
End
Attribute VB_Name = "frmBMPTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    OptionBioRB.value = True
End Sub

Private Sub imgBioRet_Click()
    OptionBioRB = True
End Sub

Private Sub imgCistern_Click()
    OptionCistern = True
End Sub

Private Sub imgDryPond_Click()
    OptionDryPond = True
End Sub

Private Sub imgInfiltT_Click()
    OptionInfilTrench = True
End Sub

Private Sub imgGreenRoof_Click()
    OptionGreenRoof = True
End Sub

Private Sub imgPorousPave_Click()
    OptionPorousPave = True
End Sub

Private Sub imgRainB_Click()
    OptionRainB = True
End Sub

Private Sub imgVegSwale_Click()
    OptionVegSwale = True
End Sub

Private Sub imgWetPond_Click()
    OptionWetPond = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSetDim_Click()
    
    'frmDataManager.Show
    Dim pBMPType As String
        If OptionBioRB = True Then
            pBMPType = "BioRetentionBasin"
        ElseIf OptionWetPond = True Then
            pBMPType = "WetPond"
        ElseIf OptionCistern = True Then
            pBMPType = "Cistern"
        ElseIf OptionDryPond = True Then
            pBMPType = "DryPond"
        ElseIf OptionInfilTrench = True Then
            pBMPType = "InfiltrationTrench"
        ElseIf OptionGreenRoof = True Then
            pBMPType = "GreenRoof"
        ElseIf OptionPorousPave = True Then
            pBMPType = "PorousPavement"
        ElseIf OptionRainB = True Then
            pBMPType = "RainBarrel"
        ElseIf OptionVegSwale = True Then
            pBMPType = "VegetativeSwale"
        End If
   
    'Setting the global variable which VegetativeSwaleidentifies
    'that the BMP forms are loaded to define the type
    'and not to add a new BMP
    
    Me.Hide
'    gIsBMPTemplate = True
    Call Set_Dimensions(pBMPType)
        
    Unload Me
CleanUp:
    Set pQueryFilter = Nothing
End Sub

Private Sub Set_Dimensions(pBMPType As String)
    
    gNewBMPType = pBMPType

    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
    
    If (pBMPTypesTable Is Nothing) Then
        Set pBMPTypesTable = CreateBMPTypesDBF("BMPTypes")
        AddTableToMap pBMPTypesTable
    End If
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "Type = '" & pBMPType & "'"
 
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
            
    If pSelRowCount > 0 Then
        Load FrmBMPTypes
        FrmBMPTypes.Form_Initialize
        FrmBMPTypes.Show vbModal
    Else
        'Create a new BMP of the current type
        Call CreateNewBmpType(pBMPType)
    End If

End Sub
