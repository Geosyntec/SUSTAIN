VERSION 5.00
Begin VB.Form frmBMPDesign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP Siting Tool - Design"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "frmBMPDesign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   38
      Top             =   7680
      Width           =   1800
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   4200
      TabIndex        =   37
      Top             =   7680
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   35
      Top             =   7200
      Width           =   1800
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   4200
      TabIndex        =   34
      Top             =   7200
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   4200
      TabIndex        =   33
      Top             =   6720
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   4200
      TabIndex        =   32
      Top             =   6240
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4200
      TabIndex        =   31
      Top             =   5760
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   4200
      TabIndex        =   30
      Top             =   5280
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   29
      Top             =   4800
      Width           =   600
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   28
      Top             =   4320
      Width           =   600
   End
   Begin VB.TextBox lblDC_RD 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   27
      Top             =   6720
      Width           =   1800
   End
   Begin VB.TextBox lblDC_WT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   23
      Top             =   6240
      Width           =   1800
   End
   Begin VB.TextBox lblDC_HG 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   22
      Top             =   5760
      Width           =   1800
   End
   Begin VB.TextBox lblDC_LS 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      Top             =   5280
      Width           =   1800
   End
   Begin VB.TextBox lblDC_DS 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   20
      Top             =   4800
      Width           =   1800
   End
   Begin VB.TextBox lblDC_DA 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   19
      Top             =   4320
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   1500
      Picture         =   "frmBMPDesign.frx":000C
      TabIndex        =   3
      Top             =   8200
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Proceed"
      Height          =   400
      Left            =   2625
      Picture         =   "frmBMPDesign.frx":03B9
      TabIndex        =   2
      Top             =   8200
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "BMP Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cmbBMPType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Building Buffer (m)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   39
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stream Buffer (m)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   36
      Top             =   7200
      Width           =   1545
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Road Buffer (m)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   26
      Top             =   6720
      Width           =   1350
   End
   Begin VB.Label lblRoad 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   25
      Top             =   3345
      Width           =   60
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Roads"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   200
      TabIndex        =   24
      Top             =   3360
      Width           =   525
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   120
      Y1              =   3840
      Y2              =   8080
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4900
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4900
      Y1              =   8080
      Y2              =   8080
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   4920
      X2              =   4920
      Y1              =   3840
      Y2              =   8080
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      X1              =   2160
      X2              =   2160
      Y1              =   4200
      Y2              =   8080
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblDEM 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   18
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label lblSoil 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   17
      Top             =   1905
      Width           =   60
   End
   Begin VB.Label lblLanduse 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   16
      Top             =   2385
      Width           =   60
   End
   Begin VB.Label lblWatertab 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   15
      Top             =   2865
      Width           =   60
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label10 
      Caption         =   "Design Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   3930
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Drainage Area (ac)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   13
      Top             =   4320
      Width           =   1605
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Drainage Slope (%)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   12
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Hydraulic Conductivity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   11
      Top             =   5280
      Width           =   1905
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Watertable Depth (ft)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   10
      Top             =   6240
      Width           =   1845
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Hydrological Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   5760
      Width           =   1605
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Streams"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   200
      TabIndex        =   8
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "UrbanLanduse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Soil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   200
      TabIndex        =   6
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   1455
      Width           =   360
   End
   Begin VB.Label Label4 
      Caption         =   "Data Requirement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   4900
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      X1              =   1440
      X2              =   1440
      Y1              =   1320
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   4920
      X2              =   4920
      Y1              =   960
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4900
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4900
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   3720
   End
End
Attribute VB_Name = "frmBMPDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\frmBMPDesign.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms




Private Sub cmbBMPType_Click()

    On Error GoTo ErrorHandler
    
    Dim oBmp As BMPobj
    Set oBmp = gBMPtypeDict.Item(cmbBMPType.Text)
    If oBmp Is Nothing Then Exit Sub
    
    ' Now Set the Properties for the form controls.....
    With oBmp
        If .DR_EL Then
            lblDEM.Caption = gDEMdata
        Else
            lblDEM.Caption = "Not Required"
        End If
        If .DR_UL Then
            lblLanduse.Caption = gLandusedata
        Else
            lblLanduse.Caption = "Not Required"
        End If
        If .DR_RD Then
            lblRoad.Caption = gRoaddata
        Else
            lblRoad.Caption = "Not Required"
        End If
        If .DR_SL Then
            lblSoil.Caption = gSoildata
        Else
            lblSoil.Caption = "Not Required"
        End If
        If .DR_WT Then
            lblWatertab.Caption = gWaterdata
        Else
            lblWatertab.Caption = "Not Required"
        End If
        lblDC_DA.Text = .DC_DA
        lblDC_DS.Text = .DC_DS
        lblDC_HG.Text = .DC_HG
        lblDC_LS.Text = .DC_LS
        lblDC_RD.Text = .DC_RD
        lblDC_WT.Text = .DC_WT
    End With
        

CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "cmbBMPType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub



Private Sub cmdCancel_Click()
    
    On Error GoTo ErrorHandler

    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdEdit_Click(Index As Integer)

    On Error GoTo ErrorHandler
    Select Case Index
    Case 0
        lblDC_DA.Enabled = True
        lblDC_DA.SetFocus
        lblDC_DA.SelStart = 0
        lblDC_DA.SelLength = Len(lblDC_DA.Text)
    Case 1
        lblDC_DS.Enabled = True
        lblDC_DS.SetFocus
        lblDC_DS.SelStart = 0
        lblDC_DS.SelLength = Len(lblDC_DS.Text)
    Case 2
        lblDC_LS.Enabled = True
        lblDC_LS.SetFocus
        lblDC_LS.SelStart = 0
        lblDC_LS.SelLength = Len(lblDC_LS.Text)
    Case 3
        lblDC_HG.Enabled = True
        lblDC_HG.SetFocus
        lblDC_HG.SelStart = 0
        lblDC_HG.SelLength = Len(lblDC_HG.Text)
    Case 4
        lblDC_WT.Enabled = True
        lblDC_WT.SetFocus
        lblDC_WT.SelStart = 0
        lblDC_WT.SelLength = Len(lblDC_WT.Text)
    Case 5
        lblDC_RD.Enabled = True
        lblDC_RD.SetFocus
        lblDC_RD.SelStart = 0
        lblDC_RD.SelLength = Len(lblDC_RD.Text)
    End Select


CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "cmdEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Set gBMPtypeDict = New Scripting.Dictionary
    gBMPtypeDict.RemoveAll
    
    ' Now Create the BMP Type objects with Props.....
    Dim oBmp As BMPobj
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Dry pond"
    oBmp.BMPId = 1
    ' DC
    oBmp.DC_DA = ">10"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = False
    'Add to Dictionary......
    gBMPtypeDict.Add "Dry pond", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Wet pond"
    oBmp.BMPId = 2
    ' DC
    oBmp.DC_DA = ">20"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    'Add to Dictionary......
    gBMPtypeDict.Add "Wet pond", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Infiltration basin"
    oBmp.BMPId = 3
    ' DC
    oBmp.DC_DA = "<10"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = ">.5"
    oBmp.DC_WT = ">4"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    'Add to Dictionary......
    gBMPtypeDict.Add "Infiltration basin", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Infiltration trench"
    oBmp.BMPId = 4
    ' DC
    oBmp.DC_DA = "<5"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = ">.5"
    oBmp.DC_WT = ">4"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = True
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Infiltration trench", oBmp
    
'    Set oBmp = New BMPobj
'    oBmp.BMPType = "Porous pavement"
'    oBmp.BMPId = 5
'    ' DC
'    oBmp.DC_DA = "<1/3"
'    oBmp.DC_DS = "<1"
'    oBmp.DC_LS = "0"
'    oBmp.DC_HG = "A-B"
'    oBmp.DC_WT = ">2"
'    oBmp.DC_RD = ""
'    'DR
'    oBmp.DR_EL = True
'    oBmp.DR_SL = True
'    oBmp.DR_UL = True
'    oBmp.DR_WT = True
'    oBmp.DR_RD = False
'    'SU
'    oBmp.SU_ND = False
'    oBmp.SU_DW = True
'    oBmp.SU_PL = True
'    oBmp.SU_RD = False
'    gBMPtypeDict.Add "Porous pavement", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Bioretention"
    oBmp.BMPId = 6
    ' DC
    oBmp.DC_DA = "<5"
    oBmp.DC_DS = "<5"
    oBmp.DC_LS = "<5"
    oBmp.DC_HG = "A-B"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = "30"
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = True
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = True
    gBMPtypeDict.Add "Bioretention", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Sand filter (surface)"
    oBmp.BMPId = 7
    ' DC
    oBmp.DC_DA = "<10"
    oBmp.DC_DS = "<10"
    oBmp.DC_LS = "0"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Sand filter (surface)", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Sand filter (non-surface)"
    oBmp.BMPId = 8
    ' DC
    oBmp.DC_DA = "<2"
    oBmp.DC_DS = "<10"
    oBmp.DC_LS = "0"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Sand filter (non-surface)", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Stormwater wetland"
    oBmp.BMPId = 9
    ' DC
    oBmp.DC_DA = ">20"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Stormwater wetland", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Perimeter filter"
    oBmp.BMPId = 10
    ' DC
    oBmp.DC_DA = "<2"
    oBmp.DC_DS = "Flat"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = "A-D"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = True
    oBmp.DR_WT = True
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Perimeter filter", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Grassed swales"
    oBmp.BMPId = 11
    ' DC
    oBmp.DC_DA = "<5"
    oBmp.DC_DS = "<6"
    oBmp.DC_LS = "1-6"
    oBmp.DC_HG = "A-C"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = "30"
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = False
    oBmp.DR_WT = True
    oBmp.DR_RD = True
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = True
    gBMPtypeDict.Add "Grassed swales", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Vegetated filterstrip"
    oBmp.BMPId = 12
    ' DC
    oBmp.DC_DA = ""
    oBmp.DC_DS = "<10"
    oBmp.DC_LS = "2-10"
    oBmp.DC_HG = "A-C"
    oBmp.DC_WT = ">2"
    oBmp.DC_RD = "30"
    'DR
    oBmp.DR_EL = True
    oBmp.DR_SL = True
    oBmp.DR_UL = False
    oBmp.DR_WT = True
    oBmp.DR_RD = True
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = True
    gBMPtypeDict.Add "Vegetated filterstrip", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Rain barrel"
    oBmp.BMPId = 13
    ' DC
    oBmp.DC_DA = "10 m buffer around house/building"
    oBmp.DC_DS = ""
    oBmp.DC_LS = ""
    oBmp.DC_HG = ""
    oBmp.DC_WT = ""
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = False
    oBmp.DR_SL = False
    oBmp.DR_UL = True
    oBmp.DR_WT = False
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Rain barrel", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Cistern"
    oBmp.BMPId = 14
    ' DC
    oBmp.DC_DA = "10 m buffer around house/building"
    oBmp.DC_DS = ""
    oBmp.DC_LS = ""
    oBmp.DC_HG = ""
    oBmp.DC_WT = ""
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = False
    oBmp.DR_SL = False
    oBmp.DR_UL = True
    oBmp.DR_WT = False
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Cistern", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Green roof"
    oBmp.BMPId = 15
    ' DC
    oBmp.DC_DA = "10 m buffer around house/building"
    oBmp.DC_DS = ""
    oBmp.DC_LS = ""
    oBmp.DC_HG = ""
    oBmp.DC_WT = ""
    oBmp.DC_RD = ""
    'DR
    oBmp.DR_EL = False
    oBmp.DR_SL = False
    oBmp.DR_UL = True
    oBmp.DR_WT = False
    oBmp.DR_RD = False
    'SU
    oBmp.SU_ND = False
    oBmp.SU_DW = False
    oBmp.SU_PL = False
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Green roof", oBmp
    
    
    ' Now Add the Items to the Combo Box......
    Dim pKeys
    pKeys = gBMPtypeDict.Keys
    Dim pKey As String
    Dim iKey As Integer
    For iKey = 0 To gBMPtypeDict.Count - 1
        pKey = pKeys(iKey)
        Set oBmp = gBMPtypeDict.Item(pKey)
        cmbBMPType.AddItem oBmp.BMPType
    Next
    
    

CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

