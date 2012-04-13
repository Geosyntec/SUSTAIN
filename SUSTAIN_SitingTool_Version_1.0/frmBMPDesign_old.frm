VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBMPDesign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP Siting Tool - Design"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmBMPDesign_old.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ImgSU_DW 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4110
      Picture         =   "frmBMPDesign_old.frx":000C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      Top             =   2985
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgSU_RD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2925
      Picture         =   "frmBMPDesign_old.frx":0219
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      Top             =   2985
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgSU_PL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1740
      Picture         =   "frmBMPDesign_old.frx":0426
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   23
      Top             =   2985
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgSU_ND 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   555
      Picture         =   "frmBMPDesign_old.frx":0633
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   2985
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgDR_RD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4125
      Picture         =   "frmBMPDesign_old.frx":0840
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgDR_WT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3165
      Picture         =   "frmBMPDesign_old.frx":0A4D
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgDR_UL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2220
      Picture         =   "frmBMPDesign_old.frx":0C5A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgDR_SL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1275
      Picture         =   "frmBMPDesign_old.frx":0E67
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox ImgDR_EL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      Picture         =   "frmBMPDesign_old.frx":1074
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   700
      Left            =   4920
      Picture         =   "frmBMPDesign_old.frx":1281
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3945
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Proceed"
      Height          =   700
      Left            =   6045
      Picture         =   "frmBMPDesign_old.frx":162E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3945
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criteria Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   7200
      TabIndex        =   2
      Top             =   960
      Width           =   4815
      Begin MSComctlLib.ListView lvDataTab 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "  DEM"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Soil"
            Object.Width           =   1341
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Landuse"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Streams"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Roads"
            Object.Width           =   1799
         EndProperty
      End
      Begin MSComctlLib.ListView lvDesignTab 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "  DA"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "DS"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "LS"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "HG"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "WT"
            Object.Width           =   1341
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "RD"
            Object.Width           =   1359
         EndProperty
      End
      Begin MSComctlLib.ListView lvSuitTab 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "     ND"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "PL"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "RD"
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "DW"
            Object.Width           =   2066
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Data Requirement"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Data Requirement"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data Requirement"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Width           =   1455
      End
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
      Width           =   4695
      Begin VB.ComboBox cmbBMPType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   3575
      Y2              =   4630
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   4800
      X2              =   4800
      Y1              =   3575
      Y2              =   4630
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000080FF&
      X1              =   1605
      X2              =   1605
      Y1              =   3920
      Y2              =   4630
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000080FF&
      X1              =   2400
      X2              =   2400
      Y1              =   3920
      Y2              =   4630
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   3910
      Y2              =   3910
   End
   Begin VB.Line Line19 
      BorderColor     =   &H000080FF&
      X1              =   840
      X2              =   840
      Y1              =   3920
      Y2              =   4630
   End
   Begin VB.Line Line20 
      BorderColor     =   &H000080FF&
      X1              =   4005
      X2              =   4005
      Y1              =   3920
      Y2              =   4630
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   4200
      Y2              =   4200
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
      Left            =   1815
      TabIndex        =   45
      Top             =   3625
      Width           =   1440
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "DA"
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
      Left            =   360
      TabIndex        =   44
      Top             =   3940
      Width           =   270
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "DS"
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
      Left            =   1080
      TabIndex        =   43
      Top             =   3940
      Width           =   255
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "LS"
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
      Left            =   1920
      TabIndex        =   42
      Top             =   3940
      Width           =   225
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "WT"
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
      Left            =   3480
      TabIndex        =   41
      Top             =   3940
      Width           =   285
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "RD"
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
      Left            =   4320
      TabIndex        =   40
      Top             =   3940
      Width           =   270
   End
   Begin VB.Line Line23 
      BorderColor     =   &H000080FF&
      X1              =   3210
      X2              =   3210
      Y1              =   3920
      Y2              =   4630
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "HG"
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
      Left            =   2700
      TabIndex        =   39
      Top             =   3940
      Width           =   255
   End
   Begin VB.Label lblDC_DA 
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
      Left            =   360
      TabIndex        =   38
      Top             =   4335
      Width           =   60
   End
   Begin VB.Label lblDC_DS 
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
      Left            =   1080
      TabIndex        =   37
      Top             =   4335
      Width           =   60
   End
   Begin VB.Label lblDC_LS 
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
      Left            =   1920
      TabIndex        =   36
      Top             =   4335
      Width           =   60
   End
   Begin VB.Label lblDC_WT 
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
      Left            =   3480
      TabIndex        =   35
      Top             =   4335
      Width           =   60
   End
   Begin VB.Label lblDC_RD 
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
      Left            =   4245
      TabIndex        =   34
      Top             =   4335
      Width           =   60
   End
   Begin VB.Label lblDC_HG 
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
      Left            =   2700
      TabIndex        =   33
      Top             =   4335
      Width           =   60
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   4800
      Y1              =   3550
      Y2              =   3550
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   4800
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmBMPDesign_old.frx":1A53
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
      Height          =   2535
      Left            =   4920
      TabIndex        =   32
      Top             =   1260
      Width           =   2160
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   5520
      Picture         =   "frmBMPDesign_old.frx":1B79
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label18 
      Caption         =   $"frmBMPDesign_old.frx":2138
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   7200
      TabIndex        =   31
      Top             =   4800
      Width           =   4440
   End
   Begin VB.Line Line31 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   4800
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "DW"
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
      Left            =   4125
      TabIndex        =   30
      Top             =   2655
      Width           =   315
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "RD"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   2655
      Width           =   270
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "PL"
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
      Left            =   1800
      TabIndex        =   28
      Top             =   2655
      Width           =   225
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "ND"
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
      Left            =   600
      TabIndex        =   27
      Top             =   2655
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "Suitability"
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
      Left            =   2040
      TabIndex        =   26
      Top             =   2355
      Width           =   960
   End
   Begin VB.Line Line30 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line29 
      BorderColor     =   &H000080FF&
      X1              =   3720
      X2              =   3720
      Y1              =   2640
      Y2              =   3430
   End
   Begin VB.Line Line27 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line26 
      BorderColor     =   &H000080FF&
      X1              =   2520
      X2              =   2520
      Y1              =   2640
      Y2              =   3430
   End
   Begin VB.Line Line25 
      BorderColor     =   &H000080FF&
      X1              =   1320
      X2              =   1320
      Y1              =   2640
      Y2              =   3430
   End
   Begin VB.Line Line24 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   4800
      X2              =   4800
      Y1              =   2280
      Y2              =   3430
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   4800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   2280
      Y2              =   3430
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "RD"
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
      Left            =   4200
      TabIndex        =   21
      Top             =   1340
      Width           =   270
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "WT"
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
      Left            =   3240
      TabIndex        =   20
      Top             =   1340
      Width           =   285
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "UL"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   1340
      Width           =   225
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "SL"
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
      Left            =   1320
      TabIndex        =   18
      Top             =   1340
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "EL"
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
      Left            =   480
      TabIndex        =   17
      Top             =   1335
      Width           =   210
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
      Left            =   1695
      TabIndex        =   16
      Top             =   1035
      Width           =   1800
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000080FF&
      X1              =   3840
      X2              =   3840
      Y1              =   1320
      Y2              =   2150
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000080FF&
      X1              =   960
      X2              =   960
      Y1              =   1320
      Y2              =   2150
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000080FF&
      X1              =   120
      X2              =   4800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      X1              =   2880
      X2              =   2880
      Y1              =   1320
      Y2              =   2150
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      X1              =   1920
      X2              =   1920
      Y1              =   1320
      Y2              =   2150
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   4800
      X2              =   4800
      Y1              =   960
      Y2              =   2150
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4800
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   4800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      DrawMode        =   9  'Not Mask Pen
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   2150
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
        ImgDR_EL.Visible = .DR_EL
        ImgDR_RD.Visible = .DR_RD
        ImgDR_SL.Visible = .DR_SL
        ImgDR_UL.Visible = .DR_UL
        ImgDR_WT.Visible = .DR_WT
        ImgSU_DW.Visible = .SU_DW
        ImgSU_ND.Visible = .SU_ND
        ImgSU_PL.Visible = .SU_PL
        ImgSU_RD.Visible = .SU_RD
        lblDC_DA.Caption = .DC_DA
        lblDC_DS.Caption = .DC_DS
        lblDC_HG.Caption = .DC_HG
        lblDC_LS.Caption = .DC_LS
        lblDC_RD.Caption = .DC_RD
        lblDC_WT.Caption = .DC_WT
    End With
        

CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "cmbBMPType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Set gBMPtypeDict = New Scripting.Dictionary
    gBMPtypeDict.RemoveAll
    
    ' Now Create the BMP Type objects with Props.....
    Dim oBmp As BMPobj
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Dry extended detention pond"
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
    gBMPtypeDict.Add "Dry extended detention pond", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Wet retention pond"
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
    gBMPtypeDict.Add "Wet retention pond", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Pocket wet pond"
    oBmp.BMPId = 3
    ' DC
    oBmp.DC_DA = "<5"
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
    gBMPtypeDict.Add "Pocket wet pond", oBmp
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Infiltration trench"
    oBmp.BMPId = 4
    ' DC
    oBmp.DC_DA = "<5"
    oBmp.DC_DS = "<15"
    oBmp.DC_LS = "Flat"
    oBmp.DC_HG = "A-B"
    oBmp.DC_WT = ">2"
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
    
    Set oBmp = New BMPobj
    oBmp.BMPType = "Porous pavement"
    oBmp.BMPId = 5
    ' DC
    oBmp.DC_DA = "<1/3"
    oBmp.DC_DS = "<1"
    oBmp.DC_LS = "0"
    oBmp.DC_HG = "A-B"
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
    oBmp.SU_DW = True
    oBmp.SU_PL = True
    oBmp.SU_RD = False
    gBMPtypeDict.Add "Porous pavement", oBmp
    
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
    oBmp.BMPType = "Pocket wetland"
    oBmp.BMPId = 10
    ' DC
    oBmp.DC_DA = "<5"
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
    gBMPtypeDict.Add "Pocket wetland", oBmp
    
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
