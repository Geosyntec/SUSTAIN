VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBMPCriteria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP Siting Tool"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmBMPCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tbsBMP 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Data Management"
      TabPicture(0)   =   "frmBMPCriteria.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Select BMP Types"
      TabPicture(1)   =   "frmBMPCriteria.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCancel(1)"
      Tab(1).Control(1)=   "cmdDel"
      Tab(1).Control(2)=   "cmdAdd"
      Tab(1).Control(3)=   "lstBMP"
      Tab(1).Control(4)=   "cmdBack"
      Tab(1).Control(5)=   "cmdNext"
      Tab(1).Control(6)=   "lstBMPSel"
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(8)=   "Label20"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "BMP Siting Criteria"
      TabPicture(2)   =   "frmBMPCriteria.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPrev"
      Tab(2).Control(1)=   "chkBB"
      Tab(2).Control(2)=   "chkSB"
      Tab(2).Control(3)=   "chkRB"
      Tab(2).Control(4)=   "chkWT"
      Tab(2).Control(5)=   "chkHG"
      Tab(2).Control(6)=   "chkIMP"
      Tab(2).Control(7)=   "chkDS"
      Tab(2).Control(8)=   "chkDA"
      Tab(2).Control(9)=   "lblDC_BB"
      Tab(2).Control(10)=   "lblDC_SB"
      Tab(2).Control(11)=   "lblDC_RB"
      Tab(2).Control(12)=   "lblDC_WT"
      Tab(2).Control(13)=   "lblDC_HG"
      Tab(2).Control(14)=   "lblDC_IMP"
      Tab(2).Control(15)=   "lblDC_DS"
      Tab(2).Control(16)=   "lblDC_DA"
      Tab(2).Control(17)=   "Frame1"
      Tab(2).Control(18)=   "cmdCancel(0)"
      Tab(2).Control(19)=   "cmdProceed"
      Tab(2).Control(20)=   "ImgBMP"
      Tab(2).Control(21)=   "Label9"
      Tab(2).Control(22)=   "Line16"
      Tab(2).Control(23)=   "Label2"
      Tab(2).Control(24)=   "Line15"
      Tab(2).Control(25)=   "Line14"
      Tab(2).Control(26)=   "Label3"
      Tab(2).Control(27)=   "Line10"
      Tab(2).Control(28)=   "Line28"
      Tab(2).Control(29)=   "Line27"
      Tab(2).Control(30)=   "Line26"
      Tab(2).Control(31)=   "Line25"
      Tab(2).Control(32)=   "Line24"
      Tab(2).Control(33)=   "Line22"
      Tab(2).Control(34)=   "Line20"
      Tab(2).Control(35)=   "Line12"
      Tab(2).Control(36)=   "Line8"
      Tab(2).Control(37)=   "Label10"
      Tab(2).Control(38)=   "Label11"
      Tab(2).Control(39)=   "Label12"
      Tab(2).Control(40)=   "Label13"
      Tab(2).Control(41)=   "Label14"
      Tab(2).Control(42)=   "Label16"
      Tab(2).Control(43)=   "Label1"
      Tab(2).ControlCount=   44
      Begin VB.Frame Frame3 
         Caption         =   "Select Vector Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   60
         Top             =   2520
         Width           =   9135
         Begin VB.CommandButton cmdBrowseSoil_lk 
            Height          =   350
            Left            =   8520
            Picture         =   "frmBMPCriteria.frx":0060
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Browse to select a directory"
            Top             =   1200
            Width           =   500
         End
         Begin VB.CommandButton cmdBrowseWT 
            Height          =   350
            Left            =   8520
            Picture         =   "frmBMPCriteria.frx":0872
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Browse to select a directory"
            Top             =   720
            Width           =   500
         End
         Begin VB.CommandButton cmdBrowseRoad 
            Height          =   350
            Left            =   8520
            Picture         =   "frmBMPCriteria.frx":1084
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Browse to select a directory"
            Top             =   240
            Width           =   500
         End
         Begin VB.ComboBox txtSoil_lk 
            Height          =   315
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   1200
            Width           =   1995
         End
         Begin VB.ComboBox txtSoilpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   1200
            Width           =   1995
         End
         Begin VB.ComboBox txtSoilpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   77
            Top             =   1200
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.CommandButton cmdBrowseSoil 
            Height          =   350
            Left            =   4080
            Picture         =   "frmBMPCriteria.frx":1896
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Browse to select a directory"
            Top             =   1200
            Width           =   500
         End
         Begin VB.ComboBox txtWTPath 
            Height          =   315
            Index           =   1
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox txtWTPath 
            Height          =   315
            Index           =   0
            Left            =   6480
            Style           =   1  'Simple Combo
            TabIndex        =   73
            Top             =   720
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.ComboBox txtLandusepath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox txtLandusepath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   70
            Top             =   720
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.CommandButton cmdBrowselanduse 
            Height          =   350
            Left            =   4080
            Picture         =   "frmBMPCriteria.frx":20A8
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Browse to select a directory"
            Top             =   720
            Width           =   500
         End
         Begin VB.ComboBox txtRoadpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   240
            Width           =   1995
         End
         Begin VB.ComboBox txtRoadpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   6480
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   66
            Top             =   240
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.ComboBox txtStreampath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   240
            Width           =   1995
         End
         Begin VB.ComboBox txtStreampath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2040
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   63
            Top             =   255
            Visible         =   0   'False
            Width           =   2000
         End
         Begin VB.CommandButton cmdBrowseStream 
            Height          =   350
            Left            =   4080
            Picture         =   "frmBMPCriteria.frx":28BA
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Browse to select a directory"
            Top             =   255
            Width           =   500
         End
         Begin VB.Label Label23 
            Caption         =   "Soil lookup table"
            Height          =   255
            Left            =   4800
            TabIndex        =   80
            Top             =   1275
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "SSURGO soil shapefile"
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   1260
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "GWT depth shapefile"
            Height          =   255
            Left            =   4800
            TabIndex        =   74
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Urban Landuse shapefile"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Road shapefile"
            Height          =   240
            Left            =   4800
            TabIndex        =   67
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Stream shapefile"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   315
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Raster Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   9135
         Begin VB.CommandButton cmdBrowseMrlc_lk 
            Height          =   350
            Left            =   8520
            Picture         =   "frmBMPCriteria.frx":30CC
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Browse to select a directory"
            Top             =   720
            Width           =   500
         End
         Begin VB.ComboBox txtMrlc_lk 
            Height          =   315
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox txtImp 
            Height          =   315
            Index           =   1
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   240
            Width           =   1995
         End
         Begin VB.CommandButton cmdBrowseIMP 
            Height          =   350
            Left            =   8520
            Picture         =   "frmBMPCriteria.frx":38DE
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Browse to select a directory"
            Top             =   240
            Width           =   500
         End
         Begin VB.ComboBox txtImp 
            Height          =   315
            Index           =   0
            Left            =   6555
            Style           =   1  'Simple Combo
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ComboBox txtMRLC 
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox txtMRLC 
            Height          =   315
            Index           =   0
            Left            =   2115
            Style           =   1  'Simple Combo
            TabIndex        =   51
            Top             =   720
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CommandButton cmdBrowseMRLC 
            Height          =   350
            Left            =   4080
            Picture         =   "frmBMPCriteria.frx":40F0
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Browse to select a directory"
            Top             =   720
            Width           =   500
         End
         Begin VB.ComboBox txtDEMpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   240
            Width           =   1995
         End
         Begin VB.ComboBox txtDEMpath 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   2115
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.CommandButton cmdBrowseDEM 
            Height          =   350
            Left            =   4080
            Picture         =   "frmBMPCriteria.frx":4902
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Browse to select a directory"
            Top             =   240
            Width           =   500
         End
         Begin VB.Label Label22 
            Caption         =   "Landuse lookup table"
            Height          =   255
            Left            =   4800
            TabIndex        =   59
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Impervious grid"
            Height          =   255
            Left            =   4800
            TabIndex        =   56
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Landuse grid"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "DEM grid"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   400
         Index           =   2
         Left            =   3480
         Picture         =   "frmBMPCriteria.frx":5114
         TabIndex        =   41
         Top             =   5000
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   400
         Index           =   1
         Left            =   -72075
         Picture         =   "frmBMPCriteria.frx":54C1
         TabIndex        =   40
         Top             =   5000
         Width           =   1050
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Back"
         Height          =   400
         Left            =   -70920
         Picture         =   "frmBMPCriteria.frx":586E
         TabIndex        =   39
         Top             =   4995
         Width           =   1050
      End
      Begin VB.CommandButton cmdDel 
         Height          =   735
         Left            =   -70800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBMPCriteria.frx":5C1B
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   735
         Left            =   -70800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBMPCriteria.frx":6147
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin MSComctlLib.ListView lstBMP 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   35
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "BMP Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   400
         Left            =   -70920
         Picture         =   "frmBMPCriteria.frx":652F
         TabIndex        =   34
         Top             =   5000
         Width           =   1050
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   400
         Left            =   -69765
         Picture         =   "frmBMPCriteria.frx":68DC
         TabIndex        =   33
         Top             =   5000
         Width           =   1050
      End
      Begin VB.CheckBox chkBB 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   22
         Top             =   4425
         Width           =   255
      End
      Begin VB.CheckBox chkSB 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   21
         Top             =   3945
         Width           =   255
      End
      Begin VB.CheckBox chkRB 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   20
         Top             =   3465
         Width           =   255
      End
      Begin VB.CheckBox chkWT 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   19
         Top             =   2985
         Width           =   255
      End
      Begin VB.CheckBox chkHG 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   18
         Top             =   2505
         Width           =   255
      End
      Begin VB.CheckBox chkIMP 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   17
         Top             =   2025
         Width           =   255
      End
      Begin VB.CheckBox chkDS 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   16
         Top             =   1545
         Width           =   255
      End
      Begin VB.CheckBox chkDA 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66180
         TabIndex        =   15
         Top             =   1065
         Width           =   255
      End
      Begin VB.TextBox lblDC_BB 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   14
         Top             =   4440
         Width           =   1400
      End
      Begin VB.TextBox lblDC_SB 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   13
         Top             =   3960
         Width           =   1400
      End
      Begin VB.TextBox lblDC_RB 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   12
         Top             =   3480
         Width           =   1400
      End
      Begin VB.TextBox lblDC_WT 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   11
         Top             =   3000
         Width           =   1400
      End
      Begin VB.TextBox lblDC_HG 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   10
         Top             =   2520
         Width           =   1400
      End
      Begin VB.TextBox lblDC_IMP 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   9
         Top             =   2040
         Width           =   1400
      End
      Begin VB.TextBox lblDC_DS 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   8
         Top             =   1560
         Width           =   1400
      End
      Begin VB.TextBox lblDC_DA 
         Appearance      =   0  'Flat
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
         Left            =   -67755
         TabIndex        =   7
         Top             =   1080
         Width           =   1400
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select BMP Type"
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
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   4215
         Begin VB.ComboBox cmbBMPType 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   400
         Index           =   0
         Left            =   -72075
         Picture         =   "frmBMPCriteria.frx":6D01
         TabIndex        =   4
         Top             =   4995
         Width           =   1050
      End
      Begin VB.CommandButton cmdProceed 
         Caption         =   "&Proceed"
         Height          =   400
         Left            =   -69765
         Picture         =   "frmBMPCriteria.frx":70AE
         TabIndex        =   3
         Top             =   4995
         Width           =   1050
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Next"
         Height          =   400
         Left            =   4635
         TabIndex        =   2
         Top             =   5000
         Width           =   1050
      End
      Begin MSComctlLib.ListView lstBMPSel 
         Height          =   4095
         Left            =   -69840
         TabIndex        =   36
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Selected BMP Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label21 
         Caption         =   "Selected BMP Types"
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
         Left            =   -69000
         TabIndex        =   43
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "Available BMP Types"
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
         Left            =   -73920
         TabIndex        =   42
         Top             =   420
         Width           =   1935
      End
      Begin VB.Image ImgBMP 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3495
         Left            =   -74760
         Stretch         =   -1  'True
         Top             =   1365
         Width           =   4215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Building Buffer (ft)"
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
         Left            =   -70215
         TabIndex        =   31
         Top             =   4440
         Width           =   1545
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stream Buffer (ft)"
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
         Left            =   -70215
         TabIndex        =   30
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Road Buffer (ft)"
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
         Left            =   -70215
         TabIndex        =   29
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         DrawMode        =   9  'Not Mask Pen
         X1              =   -70320
         X2              =   -70320
         Y1              =   600
         Y2              =   4840
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         DrawMode        =   9  'Not Mask Pen
         X1              =   -70305
         X2              =   -65890
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         DrawMode        =   9  'Not Mask Pen
         X1              =   -70305
         X2              =   -65890
         Y1              =   4845
         Y2              =   4845
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         DrawMode        =   9  'Not Mask Pen
         X1              =   -65880
         X2              =   -65880
         Y1              =   600
         Y2              =   4840
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00000000&
         X1              =   -68040
         X2              =   -68040
         Y1              =   960
         Y2              =   4840
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         X1              =   -70305
         X2              =   -65890
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "Siting Criteria"
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
         Left            =   -68640
         TabIndex        =   28
         Top             =   690
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
         Left            =   -70215
         TabIndex        =   27
         Top             =   1080
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
         Left            =   -70215
         TabIndex        =   26
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Imperviousness (%)"
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
         Left            =   -70215
         TabIndex        =   25
         Top             =   2040
         Width           =   1755
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
         Left            =   -70215
         TabIndex        =   24
         Top             =   3000
         Width           =   1845
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Hydrological Soil Groups"
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
         Left            =   -70215
         TabIndex        =   23
         Top             =   2520
         Width           =   2040
      End
      Begin VB.Label Label17 
         Caption         =   "You can use the right browse buttons to browse and select the Datasets."
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   5535
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Caption         =   "BMP image not available"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin MSComctlLib.ImageList Imglst 
      Left            =   120
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   245
      ImageHeight     =   164
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":74D3
            Key             =   "Dry Pond"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":E085
            Key             =   "Wet Pond"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":13E5F
            Key             =   "Infiltration Trench"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":1AD0E
            Key             =   "Cistern"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":27EE9
            Key             =   "Grassed Swales"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":2AAE4
            Key             =   "Porous Pavement"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":2CA70
            Key             =   "Green Roof"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":32C9E
            Key             =   "Rain Barrel"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":37171
            Key             =   "Sand Filter (Surface)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":39A70
            Key             =   "Stormwater Wetland"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":4986D
            Key             =   "Vegetated Filterstrip"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPCriteria.frx":4C97D
            Key             =   "Bioretention"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBMPCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\frmBMPDesign.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

' Tab1 variables.....
' Enumerator for the DatasetType....
Private Enum datasetType
    dtRaster = 1
    dtFeature = 2
    dtTable = 3
End Enum
'Private Variables.....
Private m_DEMFlag As Boolean
Private m_MRLCFlag As Boolean
Private m_SoilFlag As Boolean
Private m_RoadFlag As Boolean
Private m_LanduseFlag As Boolean
Private m_StreamFlag As Boolean
Private m_WTFlag As Boolean
Private m_IMPFlag As Boolean
Private m_MRLC_lkFlag As Boolean
Private m_Soil_lkFlag As Boolean

' dictionary for the Layers....
Private m_LayerDict As Scripting.Dictionary

Private m_Prev_BMP As String
Public m_BMPType As String
Private m_PassFlag As Boolean
Private m_backFlag As Boolean
Private strResult As String
Private GP As Object


Private Sub Initialize_Controls()

    On Error GoTo ErrorHandler
    
    lblDC_BB.Text = ""
    lblDC_DA.Text = ""
    lblDC_DS.Text = ""
    lblDC_IMP.Text = ""
    lblDC_HG.Text = ""
    lblDC_RB.Text = ""
    lblDC_SB.Text = ""
    lblDC_WT.Text = ""
    
    chkBB.Value = vbChecked
    chkDA.Value = vbChecked
    chkDS.Value = vbChecked
    chkIMP.Value = vbChecked
    chkHG.Value = vbChecked
    chkRB.Value = vbChecked
    chkSB.Value = vbChecked
    chkWT.Value = vbChecked
    
    chkBB.Enabled = False
    chkDA.Enabled = False
    chkDS.Enabled = False
    chkIMP.Enabled = False
    chkHG.Enabled = False
    chkRB.Enabled = False
    chkSB.Enabled = False
    chkWT.Enabled = False
    
    Label1.Visible = True

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Initialize_Controls " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub

Private Sub chkBB_Click()
    
     On Error GoTo ErrorHandler
    If chkBB.Value = vbChecked Then
        lblDC_BB.Enabled = True
    Else
        lblDC_BB.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkBB_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkDA_Click()
    
    On Error GoTo ErrorHandler
    If chkDA.Value = vbChecked Then
        lblDC_DA.Enabled = True
    Else
        lblDC_DA.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkDA_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
End Sub

Private Sub chkDS_Click()
    
     On Error GoTo ErrorHandler
    If chkDS.Value = vbChecked Then
        lblDC_DS.Enabled = True
    Else
        lblDC_DS.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkDS_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkIMP_Click()
    
     On Error GoTo ErrorHandler
    If chkIMP.Value = vbChecked Then
        lblDC_IMP.Enabled = True
    Else
        lblDC_IMP.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkimp_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkHG_Click()
    
     On Error GoTo ErrorHandler
    If chkHG.Value = vbChecked Then
        lblDC_HG.Enabled = True
    Else
        lblDC_HG.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkHG_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkRB_Click()
    
     On Error GoTo ErrorHandler
    If chkRB.Value = vbChecked Then
        lblDC_RB.Enabled = True
    Else
        lblDC_RB.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkRB_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkSB_Click()
    
     On Error GoTo ErrorHandler
    If chkSB.Value = vbChecked Then
        lblDC_SB.Enabled = True
    Else
        lblDC_SB.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkSB_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub chkWT_Click()
    
     On Error GoTo ErrorHandler
    If chkWT.Value = vbChecked Then
        lblDC_WT.Enabled = True
    Else
        lblDC_WT.Enabled = False
    End If
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "chkWT_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub cmbBMPType_Click()

    On Error GoTo ErrorHandler
    
    ' ***********************************
    ' Write back the BmpType to the dictionary.....
    ' ***********************************
    Dim oDefaultBMP As BMPobj
    Dim oBMP As BMPobj
    If m_Prev_BMP <> "" And Not gBMPCriteriaDictionary Is Nothing Then
        If gBMPCriteriaDictionary.Exists(m_Prev_BMP) Then
            Set oBMP = gBMPCriteriaDictionary.Item(m_Prev_BMP)
            ' Now Set the Properties for the form controls.....
            With oBMP
                .DC_DA = lblDC_DA.Text
                .DC_DS = lblDC_DS.Text
                .DC_HG = lblDC_HG.Text
                .DC_IMP = lblDC_IMP.Text
                .DC_RB = lblDC_RB.Text
                .DC_SB = lblDC_SB.Text
                .DC_BB = lblDC_BB.Text
                .DC_WT = lblDC_WT.Text
                'States......
                .DC_DA_State = chkDA.Value
                .DC_DS_State = chkDS.Value
                .DC_HG_State = chkHG.Value
                .DC_IMP_State = chkIMP.Value
                .DC_RB_State = chkRB.Value
                .DC_SB_State = chkSB.Value
                .DC_BB_State = chkBB.Value
                .DC_WT_State = chkWT.Value
            End With
            gBMPCriteriaDictionary.Remove m_Prev_BMP
            gBMPCriteriaDictionary.Add m_Prev_BMP, oBMP
            Set oBMP = Nothing
        End If
    End If
    
    ' Init.....
    Call Initialize_Controls

    If Not gBMPCriteriaDictionary Is Nothing Then
        If gBMPCriteriaDictionary.Exists(cmbBMPType.Text) Then
            Set oBMP = gBMPCriteriaDictionary.Item(cmbBMPType.Text)
        End If
    End If
    If gBMPtypeDict.Exists(cmbBMPType.Text) Then Set oDefaultBMP = gBMPtypeDict.Item(cmbBMPType.Text)
    If oBMP Is Nothing Then
'        If gBMPtypeDict.Exists(cmbBMPType.Text) Then
'            Set oBMP = gBMPtypeDict.Item(cmbBMPType.Text)
'        End If
        If Not oDefaultBMP Is Nothing Then
            Set oBMP = New BMPobj
            With oBMP
                .DC_DA = oDefaultBMP.DC_DA
                .DC_DS = oDefaultBMP.DC_DS
                .DC_HG = oDefaultBMP.DC_HG
                .DC_IMP = oDefaultBMP.DC_IMP
                .DC_RB = oDefaultBMP.DC_RB
                .DC_SB = oDefaultBMP.DC_SB
                .DC_BB = oDefaultBMP.DC_BB
                .DC_WT = oDefaultBMP.DC_WT
                'States......
                .DC_DA_State = oDefaultBMP.DC_DA_State
                .DC_DS_State = oDefaultBMP.DC_DS_State
                .DC_HG_State = oDefaultBMP.DC_HG_State
                .DC_IMP_State = oDefaultBMP.DC_IMP_State
                .DC_RB_State = oDefaultBMP.DC_RB_State
                .DC_SB_State = oDefaultBMP.DC_SB_State
                .DC_BB_State = oDefaultBMP.DC_BB_State
                .DC_WT_State = oDefaultBMP.DC_WT_State
            End With
        End If
    End If

    If oBMP Is Nothing Then Exit Sub
    
    ' Now Set the Properties for the form controls.....
    With oBMP
        lblDC_DA.Text = .DC_DA
        lblDC_DA.ToolTipText = "Default criteria is " & oDefaultBMP.DC_DA
        If gDEMdata <> "" Then
            lblDC_DA.Enabled = True
            chkDA.Enabled = True
            chkDA.Value = .DC_DA_State
        Else
            lblDC_DA.Enabled = False
        End If
        
        
        lblDC_DS.Text = .DC_DS
        lblDC_DS.ToolTipText = "Default criteria is " & oDefaultBMP.DC_DS
        If gDEMdata <> "" Then
            lblDC_DS.Enabled = True
            chkDS.Enabled = True
            chkDS.Value = .DC_DS_State
        Else
            lblDC_DS.Enabled = False
        End If
        
        lblDC_HG.Text = .DC_HG
        lblDC_HG.ToolTipText = "Default criteria is " & oDefaultBMP.DC_HG
        If gSoildata <> "" Then
            lblDC_HG.Enabled = True
            chkHG.Enabled = True
            chkHG.Value = .DC_HG_State
        Else
            lblDC_HG.Enabled = False
        End If
        
        
        lblDC_IMP.Text = .DC_IMP
        lblDC_IMP.ToolTipText = "Default criteria is " & oDefaultBMP.DC_IMP
        If gImperviousdata <> "" Then
            lblDC_IMP.Enabled = True
            chkIMP.Enabled = True
            chkIMP.Value = .DC_IMP_State
        Else
            lblDC_IMP.Enabled = False
        End If
        
        lblDC_RB.Text = .DC_RB
        lblDC_RB.ToolTipText = "Default criteria is " & oDefaultBMP.DC_RB & ". Three options for values are <100 for buffering inside 100 ft, >100 for buffering outside 100ft, and 50-100 for a buffering between 50 to 100 ft."
        If gRoaddata <> "" Then
            lblDC_RB.Enabled = True
            chkRB.Enabled = True
            chkRB.Value = .DC_RB_State
        Else
            lblDC_RB.Enabled = False
        End If
        
        lblDC_SB.Text = .DC_SB
        lblDC_SB.ToolTipText = "Default criteria is " & oDefaultBMP.DC_SB & ". Three options for values are <100 for buffering inside 100 ft, >100 for buffering outside 100ft, and 50-100 for a buffering between 50 to 100 ft."
        If gStreamdata <> "" Then
            lblDC_SB.Enabled = True
            chkSB.Enabled = True
            chkSB.Value = .DC_SB_State
        Else
            lblDC_SB.Enabled = False
        End If
        
        lblDC_BB.Text = .DC_BB
        lblDC_BB.ToolTipText = "Default criteria is " & oDefaultBMP.DC_BB & ". Three options for values are <100 for buffering inside 100 ft, >100 for buffering outside 100ft, and 50-100 for a buffering between 50 to 100 ft."
        If gLandusedata <> "" Then
            lblDC_BB.Enabled = True
            chkBB.Enabled = True
            chkBB.Value = .DC_BB_State
        Else
            lblDC_BB.Enabled = False
        End If
        
        lblDC_WT.Text = .DC_WT
        lblDC_WT.ToolTipText = "Default criteria is " & oDefaultBMP.DC_WT
        If gWTdata <> "" Then
            lblDC_WT.Enabled = True
            chkWT.Enabled = True
            chkWT.Value = .DC_WT_State
        Else
            lblDC_WT.Enabled = False
        End If
        
    End With
  
    ' Display the Image......
    ImgBMP.Picture = Nothing
    Dim iCnt As Integer
    For iCnt = 1 To Imglst.ListImages.Count
        If UCase(Imglst.ListImages.Item(iCnt).Key) = UCase(oBMP.BMPName) Then
            ImgBMP.Picture = Imglst.ListImages(iCnt).Picture
            Exit For
        End If
    Next iCnt
    
    'Store the Previous BMP......
    m_Prev_BMP = cmbBMPType.Text

Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "cmbBMPType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub



Private Sub cmdAdd_Click()
    
    On Error GoTo ErrorHandler
    Dim oColItems As Collection
    Set oColItems = New Collection
    Dim lstItem As ListItem
    'Collect the Items........
    For Each lstItem In lstBMP.ListItems
        If lstItem.Selected = True Then
            oColItems.Add lstItem, lstItem.Text
        End If
    Next
    'Remove the Items........
    For Each lstItem In oColItems
        lstBMPSel.ListItems.Add , lstItem.Text, lstItem.Text
        lstBMP.ListItems.Remove lstItem.Index
    Next
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

   
End Sub

Private Sub cmdBack_Click()
    
    On Error GoTo ErrorHandler
    
    tbsBMP.TabEnabled(0) = True
    tbsBMP.Tab = 0
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdBack_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdBrowseIMP_Click()
    
        
     On Error GoTo ErrorHandler
    If Browse_Dataset(dtRaster, txtImp(1)) Then m_IMPFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseIMP_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdBrowseMRLC_Click()
    
     On Error GoTo ErrorHandler
    If Browse_Dataset(dtRaster, txtMRLC(1)) Then m_MRLCFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseDEM_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdBrowseMrlc_lk_Click()
    
            
     On Error GoTo ErrorHandler
    If Browse_Dataset(dtTable, txtMrlc_lk) Then m_MRLC_lkFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseMrlc_lk_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdBrowseSoil_lk_Click()
    
      On Error GoTo ErrorHandler
    If Browse_Dataset(dtTable, txtSoil_lk) Then m_Soil_lkFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseSoil_lk_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdBrowseWT_Click()
        
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtWTPath(1)) Then m_WTFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseWT_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
End Sub



Private Sub cmdCancel_Click(Index As Integer)
    
     On Error GoTo ErrorHandler
    
    Call Generate_Siting_Cache
    'Write the Configuration data.....
    Call Generate_Criteria_Cache
    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdDel_Click()
    
    On Error GoTo ErrorHandler
    Dim oColItems As Collection
    Set oColItems = New Collection
    Dim lstItem As ListItem
    'Collect the Items........
    For Each lstItem In lstBMPSel.ListItems
        If lstItem.Selected = True Then
            oColItems.Add lstItem, lstItem.Text
        End If
    Next
    'Remove the Items........
    For Each lstItem In oColItems
        lstBMP.ListItems.Add , lstItem.Text, lstItem.Text
        lstBMPSel.ListItems.Remove lstItem.Index
    Next
        
Exit Sub
ErrorHandler:
  HandleError True, "cmdDel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdNext_Click()
    
 On Error GoTo ErrorHandler
    
    If lstBMPSel.ListItems.Count > 0 Then
        
        ' Check if Selection list is Empty.....
        If gBMPSelDict Is Nothing Then Set gBMPSelDict = New Scripting.Dictionary
        gBMPSelDict.RemoveAll
        Dim pKeys
        pKeys = gBMPtypeDict.Keys
        Dim pKey As String
        Dim iKey As Integer
        For iKey = 0 To gBMPtypeDict.Count - 1
            pKey = pKeys(iKey)
            gBMPSelDict.Add pKey, gBMPtypeDict.Item(pKey)
        Next
        ' Now Initialize the BMPCriteria Dict....
        If gBMPCriteriaDictionary Is Nothing Then
            Set gBMPCriteriaDictionary = New Scripting.Dictionary
            pKeys = gBMPSelDict.Keys
            For iKey = 0 To gBMPSelDict.Count - 1
                pKey = pKeys(iKey)
                gBMPCriteriaDictionary.Add pKey, gBMPSelDict.Item(pKey)
            Next
        End If
        ' Now fill the Selected BMP list....
        Dim iCnt As Integer
        Dim lstItem As ListItem
        For Each lstItem In lstBMP.ListItems
            If gBMPSelDict.Exists(lstItem.Text) Then gBMPSelDict.Remove lstItem.Text
        Next
        
        m_backFlag = True
        tbsBMP.TabEnabled(2) = True
        tbsBMP.Tab = 2
        
    Else
        MsgBox "Please select atleast one BMP.", vbInformation, "BMP Siting Tool"
    End If
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdNext_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

Private Sub cmdOK_Click()
  
  On Error GoTo ErrorHandler
  
  ' ***********************************
  ' Validate the Datasets...........
  ' ***********************************
  
  If Not Check_Datasets Then Exit Sub
  If Not ValidateDatasets_ST(False) Then Exit Sub
    
   ' Write to Source text File......
   Call Generate_Siting_Cache
   Call SetDataDirectory_ST
   
    ' Close the form..........
    tbsBMP.TabEnabled(1) = True
    tbsBMP.Tab = 1
    gDataValid = True
    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Function Check_Datasets() As Boolean
    
    On Error GoTo ErrorHandler
    Check_Datasets = False
    
    If txtDEMpath(1).Text = "" And txtImp(1).Text = "" And txtMRLC(1).Text = "" And txtStreampath(1).Text = "" And txtRoadpath(1).Text = "" And txtLandusepath(1).Text = "" And txtWTPath(1).Text = "" And txtSoilpath(1).Text = "" Then
        MsgBox "Please select atleast one data layer", vbCritical, "BMP Siting Tool"
        Exit Function
    End If
    
    ' MRLC
    If txtMRLC(1).Text <> "" And txtMrlc_lk.Text = "" Then
        MsgBox "Please select Landuse lookup table", vbCritical, "BMP Siting Tool"
        Exit Function
    End If
    If txtMRLC(1).Text = "" And txtMrlc_lk.Text <> "" Then
        MsgBox "Please select MRLC data layer", vbCritical, "BMP Siting Tool"
        Exit Function
    End If
    
    'SOIL
    If txtSoilpath(1).Text <> "" And txtSoil_lk.Text = "" Then
        MsgBox "Please select Soil lookup table", vbCritical, "BMP Siting Tool"
        Exit Function
    End If
    If txtSoilpath(1).Text = "" And txtSoil_lk.Text <> "" Then
        MsgBox "Please select Soil data layer", vbCritical, "BMP Siting Tool"
        Exit Function
    End If
    
    Check_Datasets = True

Exit Function
ErrorHandler:
  HandleError True, "Check_Datasets " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
  
End Function

Private Sub Generate_Siting_Cache()
    
    On Error GoTo ErrorHandler
    
    'Create a file for writing the datasources -- Arun Raj; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = DefineApplicationPath_ST
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dataSrcFN As String 'Arun Raj -- October 2004
    dataSrcFN = gApplication.Document
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_Siting.src" 'dataSrcFN = Replace(dataSrcFN, ".mxd", "_Siting.src")
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
   'If gDEMdata = "" Then gDEMdata = "Not Available"
   If txtMRLC(0).Text = "" Then
        gMRLCdata = txtMRLC(1).Text
        pDataSrcFile.WriteLine "gMRLCdata" & vbTab & gMRLCdata
   Else
        gMRLCdata = txtMRLC(0).Text
        pDataSrcFile.WriteLine "gMRLCdata" & vbTab & gMRLCdata
   End If
   'If gMRLCdata = "" Then gMRLCdata = "Not Available"
   If txtLandusepath(0).Text = "" Then
        gLandusedata = txtLandusepath(1).Text
        pDataSrcFile.WriteLine "gLandusedata" & vbTab & gLandusedata
   Else
        gLandusedata = txtLandusepath(0).Text
        pDataSrcFile.WriteLine "gLandusedata" & vbTab & gLandusedata
   End If
   'If gLandusedata = "" Then gLandusedata = "Not Available"
   If txtRoadpath(0).Text = "" Then
        gRoaddata = txtRoadpath(1).Text
        pDataSrcFile.WriteLine "gRoaddata" & vbTab & gRoaddata
   Else
        gRoaddata = txtRoadpath(0).Text
        pDataSrcFile.WriteLine "gRoaddata" & vbTab & gRoaddata
   End If
   'If gRoaddata = "" Then gRoaddata = "Not Available"
   If txtSoilpath(0).Text = "" Then
        gSoildata = txtSoilpath(1).Text
        pDataSrcFile.WriteLine "gSoildata" & vbTab & gSoildata
   Else
        gSoildata = txtSoilpath(0).Text
        pDataSrcFile.WriteLine "gSoildata" & vbTab & gSoildata
   End If
   'If gSoildata = "" Then gSoildata = "Not Available"
   If txtStreampath(0).Text = "" Then
        gStreamdata = txtStreampath(1).Text
        pDataSrcFile.WriteLine "gStreamdata" & vbTab & gStreamdata
   Else
        gStreamdata = txtStreampath(0).Text
        pDataSrcFile.WriteLine "gStreamdata" & vbTab & gStreamdata
   End If
   'If gStreamdata = "" Then gStreamdata = "Not Available"
   If txtImp(0).Text = "" Then
        gImperviousdata = txtImp(1).Text
        pDataSrcFile.WriteLine "gImperviousdata" & vbTab & gImperviousdata
   Else
        gImperviousdata = txtImp(0).Text
        pDataSrcFile.WriteLine "gImperviousdata" & vbTab & gImperviousdata
   End If
   'If gImperviousdata = "" Then gImperviousdata = "Not Available"
   If txtWTPath(0).Text = "" Then
        gWTdata = txtWTPath(1).Text
        pDataSrcFile.WriteLine "gWTdata" & vbTab & gWTdata
   Else
        gWTdata = txtWTPath(0).Text
        pDataSrcFile.WriteLine "gWTdata" & vbTab & gWTdata
   End If
   'If gWTdata = "" Then gWTdata = "Not Available"
   
   If txtMrlc_lk.Text <> "" Then
        gMrlcTable = txtMrlc_lk.Text
        pDataSrcFile.WriteLine "gMrlcTable" & vbTab & gMrlcTable
   End If
   If txtSoil_lk.Text <> "" Then
        gSoilTable = txtSoil_lk.Text
        pDataSrcFile.WriteLine "gSoilTable" & vbTab & gSoilTable
   End If
   
   ' Close the File......
   pDataSrcFile.Close

Exit Sub
ErrorHandler:
  HandleError True, "Generate_Siting_Cache " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

Private Sub Generate_Criteria_Cache()
    
    On Error GoTo ErrorHandler
    
    If gBMPCriteriaDictionary Is Nothing Then Exit Sub
    
    'Create a file for writing the datasources -- Arun Raj; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = DefineApplicationPath_ST
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dataSrcFN As String
    dataSrcFN = gApplication.Document
    'dataSrcFN = Replace(dataSrcFN, ".mxd", "_criteria.src")
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_criteria.src"
    dataSrcFN = pAppPath & dataSrcFN
    
    ' ***********************************
    ' Write back the BmpType to the dictionary.....
    ' ***********************************
    Dim oBMP As BMPobj
    If m_Prev_BMP <> "" And gBMPCriteriaDictionary.Exists(m_Prev_BMP) Then
        Set oBMP = gBMPCriteriaDictionary.Item(m_Prev_BMP)
        If oBMP Is Nothing Then MsgBox "Dictionary object failed.": Exit Sub
        ' Now Set the Properties for the form controls.....
        With oBMP
            .DC_DA = lblDC_DA.Text
            .DC_DS = lblDC_DS.Text
            .DC_HG = lblDC_HG.Text
            .DC_IMP = lblDC_IMP.Text
            .DC_RB = lblDC_RB.Text
            .DC_SB = lblDC_SB.Text
            .DC_BB = lblDC_BB.Text
            .DC_WT = lblDC_WT.Text
            'States......
            .DC_DA_State = chkDA.Value
            .DC_DS_State = chkDS.Value
            .DC_HG_State = chkHG.Value
            .DC_IMP_State = chkIMP.Value
            .DC_RB_State = chkRB.Value
            .DC_SB_State = chkSB.Value
            .DC_BB_State = chkBB.Value
            .DC_WT_State = chkWT.Value
        End With
        gBMPCriteriaDictionary.Remove m_Prev_BMP
        gBMPCriteriaDictionary.Add m_Prev_BMP, oBMP
    End If
    Set oBMP = Nothing ' Required.....
    
    ' ***********************************
    ' Write the BMP Paramenters to Source File....
    ' ***********************************
    Dim pDataSrcFile
    Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForWriting, True, TristateUseDefault)
    Dim pKeys
    Dim pKey As String
    Dim iKey As Integer
    ' First write the Selected BMPs.......
    pKeys = gBMPSelDict.Keys
    For iKey = 0 To gBMPSelDict.Count - 1
        pKey = pKeys(iKey)
        pDataSrcFile.WriteLine pKey
    Next
    pDataSrcFile.WriteLine ""
    
    ' Now write the BMP object properties....
    pKeys = gBMPCriteriaDictionary.Keys
    For iKey = 0 To gBMPCriteriaDictionary.Count - 1
        pKey = pKeys(iKey)
        If pKey <> "" Then
            Set oBMP = gBMPCriteriaDictionary.Item(pKey)
            pDataSrcFile.WriteLine "BMPName" & vbTab & oBMP.BMPName
            pDataSrcFile.WriteLine "DC_DA" & vbTab & oBMP.DC_DA & vbTab & oBMP.DC_DA_State
            pDataSrcFile.WriteLine "DC_BB" & vbTab & oBMP.DC_BB & vbTab & oBMP.DC_BB_State
            pDataSrcFile.WriteLine "DC_DS" & vbTab & oBMP.DC_DS & vbTab & oBMP.DC_DS_State
            pDataSrcFile.WriteLine "DC_HG" & vbTab & oBMP.DC_HG & vbTab & oBMP.DC_HG_State
            pDataSrcFile.WriteLine "DC_IMP" & vbTab & oBMP.DC_IMP & vbTab & oBMP.DC_IMP_State
            pDataSrcFile.WriteLine "DC_RB" & vbTab & oBMP.DC_RB & vbTab & oBMP.DC_RB_State
            pDataSrcFile.WriteLine "DC_SB" & vbTab & oBMP.DC_SB & vbTab & oBMP.DC_SB_State
            pDataSrcFile.WriteLine "DC_WT" & vbTab & oBMP.DC_WT & vbTab & oBMP.DC_WT_State
            pDataSrcFile.WriteLine ""
        End If
        Set oBMP = Nothing
    Next
    pDataSrcFile.Close
    Set oBMP = Nothing ' Required.....

Exit Sub
ErrorHandler:
  HandleError True, "cmdPrev_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdPrev_Click()
        
    On Error GoTo ErrorHandler
    
    'Write the Configuration data.....
    Call Generate_Criteria_Cache
    
    tbsBMP.TabEnabled(1) = True
    tbsBMP.Tab = 1
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdPrev_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
        
End Sub

Private Sub cmdProceed_Click()

    On Error GoTo ErrorHandler
    m_PassFlag = True
    
    'Check if the BMP Type is picked....
    If cmbBMPType.ListCount = 0 Then
        MsgBox "Please select a BMP Type to place and proceed.", vbInformation, "BMP_Siting_Tool"
        Exit Sub
    End If
    
    'Write the Configuration data.....
    Call Generate_Criteria_Cache
    
    ' ########################################
    ' Show the Splash Screen..............................................................
    ' ########################################
    Dim f As BMP_Siting_Tool.frmSplash
    Set f = New frmSplash
    Me.Hide
    f.Show
    f.Refresh
    'Set the window to be on top....
    ModuleSustainGlobal.AlwaysOnTop_ST f, -1
    
    ' ########################################
    ' Check for Spatial Analyst License information and Proceed.........
    ' ########################################
    If Not CheckSpatialAnalystLicense Then m_PassFlag = False: GoTo Cleanup
    
    ' ########################################
    ' Copy All featureClasses to the Process folder....
    ' ########################################
    If Not Copy_Data_toWorkfolder Then m_PassFlag = False: GoTo Cleanup

    ' ########################################
    ' Validate the Datasets and then Proceed..........
    ' ########################################
    'Update the Status....
    f.lblStatus.Caption = "Validating data. Please wait!!!"
    f.Refresh
    ' Initialize the operators...
    Call ModuleRasterUtils.InitializeOperators
    
    m_PassFlag = ValidateDatasets_ST(True)
    If Not m_PassFlag Then GoTo Cleanup
    
    
    ' ########################################
    ' Check for Input Projection information and Proceed....................
    ' ########################################
    'Update the Status....
    f.lblStatus.Caption = "Validating input projections. Please wait!!!"
    f.Refresh
    If Not CheckInputDataProjection_ST Then m_PassFlag = False: GoTo Cleanup
            
    ' ########################################
    ' Now Loop through the BMPs in the list and process each BMP...
    ' ########################################
    Dim pcolResult As Scripting.Dictionary
    Set pcolResult = New Scripting.Dictionary
    If GP Is Nothing Then Set GP = CreateObject("esriGeoprocessing.GpDispatch.1") ' Create a GP object...........
    Dim pKeys
    Dim pKey As String
    Dim iKey As Integer
    For iKey = 0 To cmbBMPType.ListCount - 1
        cmbBMPType.ListIndex = iKey
        
        ' *****************************************
        ' First Check if atleast one Criteria is Checked............
        If Check_Criteria Then
            ' *******************************
            ' Delete any layers from the last result....
            Dim pLayer As ILayer
            Set pLayer = GetInputFeatureLayer(Replace(cmbBMPType.Text, " ", "_"))
            If Not pLayer Is Nothing Then
                MsgBox "Please delete the analysis layer and proceed.", vbInformation, "BMP Siting Tool"
                GoTo Cleanup
            End If
            Set pLayer = GetInputFeatureLayer("Composite")
            If Not pLayer Is Nothing Then
                MsgBox "Please delete the analysis layer and proceed.", vbInformation, "BMP Siting Tool"
                GoTo Cleanup
            End If
            ' Call the Process...
            f.lblBMP.Caption = "Analyzing " & cmbBMPType.Text & " BMP"
            Call Start_Analyze(pcolResult, f)
            If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        Else
            If cmbBMPType.ListCount = 1 Then m_PassFlag = False
        End If
        
    Next
    
    ' **************************************************
    ' Now Create a Composite Layer if selected BMPs are multiple ....
    ' **************************************************
    If pcolResult.Count > 1 Then
        f.lblBMP.Caption = ""
        Create_Composite_Layer pcolResult, f
        ' Add the Composite Layer....
        pcolResult.Add "Composite", "Composite"
    End If
        
Cleanup:

    If Not f Is Nothing Then
        Unload f
        If m_PassFlag Then
            ' Turn off layers....
            Call Turn_Off_Layers(pcolResult)
            MsgBox "Analysis completed.", vbInformation, "BMP Siting Tool"
        End If
        Me.Show
        AlwaysOnTop_ST Me, -1
        cmdCancel(0).SetFocus ' Get the focus out of Proceed...
        ' Set the Globals back to normal....
        gSoildata = txtSoilpath(1).Text
        gMRLCdata = txtMRLC(1).Text
        'Unload Me
    End If
    
    ' Cleanup.....
    Set pcolResult = Nothing

  Exit Sub
ErrorHandler:
  HandleError True, "cmdProceed_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
  GoTo Cleanup
End Sub

Private Function Check_Criteria() As Boolean
    
    On Error GoTo ErrorHandler
    Check_Criteria = False
    
    ' Now Check if the Buffer distance is > 0.......................
    If lblDC_RB.Enabled = True And Not Check_Expression(lblDC_RB.Text) Then
        MsgBox "Road buffer should be in one of the following formats: " & vbCrLf & ">min, <max, min-max, NA", vbInformation, "BMP Siting Tool"
        Exit Function
    End If
    If lblDC_SB.Enabled = True And Not Check_Expression(lblDC_SB.Text) Then
        MsgBox "Stream buffer should be in one of the following formats: " & vbCrLf & ">min, <max, min-max, NA", vbInformation, "BMP Siting Tool"
        Exit Function
    End If
    If lblDC_BB.Enabled = True And Not Check_Expression(lblDC_BB.Text) Then
        MsgBox "Building buffer should be in one of the following formats: " & vbCrLf & ">min, <max, min-max, NA", vbInformation, "BMP Siting Tool"
        Exit Function
    End If
            
    
    '/////////////////////////////////////////
    If lblDC_SB.Enabled = True Then
        If lblDC_BB.Enabled = True Or lblDC_DA.Enabled = True Or lblDC_DS.Enabled = True Or lblDC_HG.Enabled = True Or lblDC_IMP.Enabled = True Or lblDC_RB.Enabled = True Or lblDC_SB.Enabled = True Or lblDC_WT.Enabled = True Then
            Check_Criteria = True
        Else
            MsgBox "Please select one more other than Stream criteria for " & cmbBMPType.Text, vbInformation, "BMP Siting Tool"
            Exit Function
        End If
    End If
    ' If Selected BMP is Porous Pavement....
    Check_Criteria = False
    If (UCase(cmbBMPType.Text) = "POROUS PAVEMENT" Or UCase(cmbBMPType.Text) = "GREEN ROOF") And chkBB.Value = vbChecked And chkBB.Enabled = True Then
        Check_Criteria = True
    ElseIf lblDC_BB.Enabled = True Or lblDC_DA.Enabled = True Or lblDC_DS.Enabled = True Or lblDC_HG.Enabled = True Or lblDC_IMP.Enabled = True Or lblDC_RB.Enabled = True Or lblDC_SB.Enabled = True Or lblDC_WT.Enabled = True Then
        Check_Criteria = True
    End If
    If Not Check_Criteria Then MsgBox "Please select atleast one criteria for " & cmbBMPType.Text, vbInformation, "BMP Siting Tool"
        
Exit Function
ErrorHandler:
  HandleError True, "Check_Criteria " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function

'Feb 19, 2009: edited by Ying to accept >min, <max, min-max formats
Private Function Check_Expression(ByVal strExp As String) As Boolean
    
    On Error GoTo ErrorHandler
    Dim strTmp As String
    Dim iCnt As Integer
    Dim strTmp2 As String
    
    iCnt = 1
    Check_Expression = False
    
    'check NA
    If InStr(1, strExp, "NA", vbTextCompare) > 0 Then Check_Expression = True
    
    'check if it begins with >
    iCnt = InStr(1, Trim(strExp), ">", vbTextCompare)
    If iCnt > 0 Then
        strTmp = Right(Trim(strExp), Len(Trim(strExp)) - iCnt)
        If IsNumeric(strTmp) And InStr(1, strTmp, "-", vbTextCompare) = 0 Then  'is number and is positive
            Check_Expression = True
        End If
    End If
    
    'check if it begins with <
    iCnt = InStr(1, Trim(strExp), "<", vbTextCompare)
    If iCnt > 0 Then
        strTmp = Right(Trim(strExp), Len(Trim(strExp)) - iCnt)
        If IsNumeric(strTmp) And InStr(1, strTmp, "-", vbTextCompare) = 0 Then  'is number and is positive
            Check_Expression = True
        End If
    End If
    
    'check if its format is min-max
    iCnt = InStr(1, Trim(strExp), "-", vbTextCompare)
    If iCnt > 1 Then
        strTmp = Left(Trim(strExp), iCnt - 1)
        strTmp2 = Right(Trim(strExp), Len(Trim(strExp)) - iCnt)
        If IsNumeric(strTmp) And InStr(1, strTmp, "-", vbTextCompare) = 0 And IsNumeric(strTmp2) And InStr(1, strTmp2, "-", vbTextCompare) = 0 Then
            Check_Expression = True
        End If
    End If
   
Exit Function
ErrorHandler:

  HandleError True, "Check_Expression " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function

Private Sub Start_Analyze(ByRef pcolResult As Scripting.Dictionary, ByVal f As frmSplash)
    
    On Error GoTo ErrorHandler
    
    ' #########################################
    '............................... Now Validate the Data....................................
    ' #########################################
    If Not Validate_Input_Data Then Exit Sub
    
    ' #########################################
    '...................... Now Start Pre-Processing the Data..........................
    ' #########################################

    Dim pActiveView As IActiveView
    Set pActiveView = gMap
    pActiveView.Refresh
    gMxDoc.UpdateContents
    strResult = Replace(cmbBMPType.Text, " ", "_")
          
    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ' Pre-Process FEATURE Data...........
    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

    Dim dScalefac As Double
    ' Get the Map units....
    dScalefac = Get_Scale_Factor(gMap.MapUnits)
        
    ' Declarations.........
    Dim pResult As IFeatureLayer
    Dim pQueryFilter As IQueryFilter
    Dim pSelset As ISelectionSet
    Dim pSelectionSet As ISelectionSet
    Dim pFeatureSelection As IFeatureSelection
    Dim pExportFclass As IFeatureClass
    
    Dim Bldg_Flag As Boolean
    Dim Stream_Flag As Boolean
    Dim Skip_Soil_Flag As Boolean
    Dim strWhereClause As String
    Dim iCnt As Integer
    
    ' Get the Source Layers.........
    Dim pRoadLayer As IFeatureLayer
    Set pRoadLayer = GetFeatureLayer(gWorkingfolder, gRoaddata)
    Dim pStreamLayer As IFeatureLayer
    Set pStreamLayer = GetFeatureLayer(gWorkingfolder, gStreamdata)
    Dim pSoilLayer As IFeatureLayer
    Set pSoilLayer = Nothing
    Dim pWTLayer As IFeatureLayer
    Set pWTLayer = GetFeatureLayer(gWorkingfolder, gWTdata)
    Dim pLanduseLayer As IFeatureLayer
    Set pLanduseLayer = GetFeatureLayer(gWorkingfolder, gLandusedata)
    
    Dim pTmpDataset As IDataset
    
    ' ******************************************
    If lblDC_RB.Text <> "NA" And chkRB.Enabled = True And chkRB.Value = vbChecked And Not pRoadLayer Is Nothing Then
                
        'Update the Status....
        f.lblStatus.Caption = "Buffering Road data. Please wait!!!"
        f.Refresh
        
        PrepareBufferFeature lblDC_RB.Text, gWorkingfolder, gRoaddata
        Set pRoadLayer = GetInputFeatureLayer(gRoaddata & "_Buf")
'
        ' **************************************************
        ' Was some problem with Line features with creating features.....
        ' **************************************************
'        If pRoadLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
'            Set pRoadLayer = CreateShapefile(gRoaddata, pRoadLayer)
'            CreateBufferFeatureClass pRoadLayer, lblDC_RB.Text, True
'        Else
'            Call Create_Line_Buffer(gWorkingfolder, gRoaddata, lblDC_RB.Text * dScalefac)
'            Set pRoadLayer = GetInputFeatureLayer(gRoaddata & "_Dis")
'            MsgBox "pRoadlayer name:" & pRoadLayer.Name
'        End If
        
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    ' ******************************************
    If lblDC_SB.Text <> "NA" And chkSB.Enabled = True And chkSB.Value = vbChecked And Not pStreamLayer Is Nothing Then
        
        'Update the Status....
        f.lblStatus.Caption = "Buffering Stream data. Please wait!!!"
        f.Refresh
        
        PrepareBufferFeature lblDC_SB.Text, gWorkingfolder, gStreamdata
        Set pStreamLayer = GetInputFeatureLayer(gStreamdata & "_Buf")

'        ' **************************************************
'        ' Was some problem with Line features with creating features....
'        ' **************************************************
'        If pStreamLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
'            Set pStreamLayer = CreateShapefile(gStreamdata, pStreamLayer)
'            CreateBufferFeatureClass pStreamLayer, lblDC_SB.Text, True
'        Else
'            Call Create_Line_Buffer(gWorkingfolder, gStreamdata, lblDC_SB.Text * dScalefac)
'            Set pStreamLayer = GetInputFeatureLayer(gStreamdata & "_Dis")
'        End If
        
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    ' ******************************************
    If lblDC_WT.Text <> "NA" And chkWT.Enabled = True And chkWT.Value = vbChecked And Not pWTLayer Is Nothing Then
        
        'Update the Status....
        f.lblStatus.Caption = "Filtering Water table data. Please wait!!!"
        f.Refresh
        ' Select the Features from the Water table Data.....
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "GWdep_ft " & lblDC_WT.Text
        'Select all features of the feature class
        Set pFeatureSelection = pWTLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultAdd, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count > 0 Then
            ' Export the Selected features to a new Shapefile.....
            Set pExportFclass = Get_ExportedShapefile(pWTLayer, strResult & "WT_Filter", True)
            gMap.ClearSelection ' Clear the Selection.....
            If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
            ' If the layer is added to map... Delete......
            Set pWTLayer = GetFeatureLayer(gWorkingfolder, strResult & "WT_Filter")
            If Not pWTLayer Is Nothing Then gMap.DeleteLayer pWTLayer
            ' Add the new Layer to map.........
            Set pWTLayer = New FeatureLayer
            Set pWTLayer.FeatureClass = pExportFclass
            pWTLayer.Visible = True
            pWTLayer.Name = pExportFclass.AliasName
            gMap.AddLayer pWTLayer
        Else
            Set pWTLayer = Nothing
        End If
        
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
     ' ******************************************
     ' Filter buildings.............................................................
     
    ' *******************************************
    ' Calculate the Area of the Buildings.....
    ' *******************************************
    If lblDC_DA.Text <> "NA" And chkDA.Enabled = True And chkDA.Value = vbChecked And Not pLanduseLayer Is Nothing And UCase(cmbBMPType.Text) = "GREEN ROOF" Then
        Set pExportFclass = pLanduseLayer.FeatureClass
        If pExportFclass.Fields.FindField("Acreage") = -1 Then
            ' the following creates and adds the Acreage field
            Dim pFieldEdit As IFieldEdit
            Set pFieldEdit = New Field
            pFieldEdit.Name = "Acreage"
            pFieldEdit.Type = esriFieldTypeDouble
            pExportFclass.AddField pFieldEdit
        End If
        Dim pCalc As ICalculator
        Set pCalc = New Calculator
        Dim pCursor As ICursor
        Set pCursor = pExportFclass.Update(Nothing, True)
        With pCalc
            Set .Cursor = pCursor
            .ShowErrorPrompt = False
            ' The PreExpression is the actual equation. The field area must be converted into acres.
            .PreExpression = "Dim dblArea As Double" & vbNewLine & "Dim pArea as IArea" & vbNewLine & "Set pArea = [shape]" & vbNewLine & "dblArea = round((pArea.Area * 0.000247), 2)"
            .Expression = "dblArea"
            .Field = "Acreage"
        End With
        pCalc.Calculate
        
        ' Select the Features from the Landuse Data.....
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "Acreage" & lblDC_DA.Text
        'Select all features of the feature class
        Set pFeatureSelection = pLanduseLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count > 0 Then
            ' Export the Selected features to a new Shapefile.....
            Set pExportFclass = Get_ExportedShapefile(pLanduseLayer, strResult & "BB_Area", False)
            gMap.ClearSelection ' Clear the Selection.....
            ' Add the new Layer to map.........
            Set pLanduseLayer = New FeatureLayer
            Set pLanduseLayer.FeatureClass = pExportFclass
            pLanduseLayer.Visible = True
            pLanduseLayer.Name = pExportFclass.AliasName
            gMap.AddLayer pLanduseLayer
            If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        Else
            Set pLanduseLayer = Nothing
        End If
    End If
    ' ************************************
    ' Filter for Buildings.....
    ' ************************************
    If Not pLanduseLayer Is Nothing And lblDC_BB.Enabled = True Then
        'Update the Status....
        f.lblStatus.Caption = "Filtering building data. Please wait!!!"
        f.Refresh
        ' Select the Features from the Water table Data.....
        Set pQueryFilter = New QueryFilter
        If cmbBMPType.Text = "Porous Pavement" Then
            pQueryFilter.WhereClause = "LU_DESC Like '%Roadways%' or LU_DESC Like '%Parking%'"
        Else
            pQueryFilter.WhereClause = "LU_DESC = 'Buildings'"
        End If
        'Select all features of the feature class
        Set pFeatureSelection = pLanduseLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultAdd, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count > 0 Then
            ' Export the Selected features to a new Shapefile.....
            If lblDC_BB.Text = "NA" And Not pLanduseLayer Is Nothing Then
                Set pExportFclass = Get_ExportedShapefile(pLanduseLayer, strResult & "BB_Filter", True) ' Dissolve...
            Else
                Set pExportFclass = Get_ExportedShapefile(pLanduseLayer, strResult & "BB_Filter", False) ' Dont Dissolve.....
            End If
            gMap.ClearSelection ' Clear the Selection.....
            Set pLanduseLayer = GetInputFeatureLayer(strResult & "BB_Area")
            If Not pLanduseLayer Is Nothing Then gMap.DeleteLayer pLanduseLayer
            ' Add the new Layer to map.........
            Set pLanduseLayer = New FeatureLayer
            Set pLanduseLayer.FeatureClass = pExportFclass
            pLanduseLayer.Visible = True
            pLanduseLayer.Name = pExportFclass.AliasName
            gMap.AddLayer pLanduseLayer
            If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        Else
            Set pLanduseLayer = Nothing
            Set pLanduseLayer = GetInputFeatureLayer(strResult & "BB_Area")
            If Not pLanduseLayer Is Nothing Then gMap.DeleteLayer pLanduseLayer
            MsgBox "No features found with this criteria.", vbInformation, "BMP Siting Tool"
        End If
    End If
    
    ' ********************************************
    ' Buffer the Buldings....
    ' ********************************************
    If lblDC_BB.Text <> "NA" And chkBB.Enabled = True And chkBB.Value = vbChecked And Not pLanduseLayer Is Nothing Then
        'Update the Status....
        f.lblStatus.Caption = "Buffering building data. Please wait!!!"
        f.Refresh
        Dim ptmpLayer As IFeatureLayer
        Set ptmpLayer = pLanduseLayer
                   
        Set pTmpDataset = ptmpLayer
        
        'PrepareBufferFeature lblDC_BB.Text, gWorkingfolder, gLandusedata
        PrepareBufferFeature lblDC_BB.Text, pTmpDataset.Workspace.PathName, pTmpDataset.BrowseName
        'Set pLanduseLayer = GetInputFeatureLayer(gLandusedata & "_Buf")
        Set pLanduseLayer = GetInputFeatureLayer(pLanduseLayer.Name & "_Buf")

'        Set pLanduseLayer = CreateShapefile(gLandusedata, pLanduseLayer)
'        CreateBufferFeatureClass pLanduseLayer, lblDC_BB.Text, True
'        gMap.DeleteLayer ptmpLayer
'        Delete_Dataset_ST gWorkingfolder, ptmpLayer.Name
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    ' ***********************************************************************
    ' Start Analyzing with Soil Layer....................
    ' ***********************************************************************
    
    ' ******************************************
    If lblDC_HG.Text <> "NA" And chkHG.Enabled = True And chkHG.Value = vbChecked Then
    
        'Update the Status....
        f.lblStatus.Caption = "Filtering Soil data. Please wait!!!"
        f.Refresh

        ' ******************************************
        Set pSoilLayer = GetFeatureLayer(gWorkingfolder, gSoildata)
        strWhereClause = "HYDGRP In ("
        Dim pValues As Collection
        Set pValues = Get_Soil_Groups(lblDC_HG.Text, "-")
        For iCnt = 1 To pValues.Count
            strWhereClause = strWhereClause & "'" & pValues(iCnt) & "',"
        Next iCnt
        strWhereClause = Generic_Trim(strWhereClause, ",", "") & ")"

        ' Select the Features from the Soil Data.....
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = strWhereClause
        'Select all features of the feature class
        Set pFeatureSelection = pSoilLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count = 0 Then GoTo Check
        
        ' Export the Selected features to a new Shapefile.....
        Set pExportFclass = Get_ExportedShapefile(pSoilLayer, strResult & "Soil_Filter", True)
        gMap.ClearSelection ' Clear the Selection.....
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        ' If the layer is added to map... Delete......
        Set pSoilLayer = GetFeatureLayer(gWorkingfolder, strResult & "Soil_Filter")
        If Not pSoilLayer Is Nothing Then gMap.DeleteLayer pSoilLayer
        ' Add the new Layer to map.........
        Set pSoilLayer = New FeatureLayer
        Set pSoilLayer.FeatureClass = pExportFclass
        pSoilLayer.Visible = True
        pSoilLayer.Name = pExportFclass.AliasName
        gMap.AddLayer pSoilLayer
        
        ' ****************************************
        ' Exclude building Area..............................................
        If lblDC_BB.Text <> "NA" And chkBB.Enabled = True And chkBB.Value = vbChecked And Not pLanduseLayer Is Nothing Then
            'Update the Status....
            f.lblStatus.Caption = "Analyzing building data. Please wait!!!"
            f.Refresh
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pLanduseLayer, strResult & "_BB")
            'Set pResult = Get_IntersectLayer(pSoilLayer, pRoadLayer, strResult)
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            gMap.DeleteLayer pLanduseLayer
            Delete_Dataset_ST gWorkingfolder, pLanduseLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
            Else
                Set pSoilLayer = Nothing
            End If
        End If
        
        ' *****************************************
        ' Intersect SOIL with Watertable................................
        If pSoilLayer Is Nothing Then GoTo Check
        If lblDC_WT.Text <> "NA" And chkWT.Enabled = True And chkWT.Value = vbChecked And Not pWTLayer Is Nothing Then
            'Update the Status....
            f.lblStatus.Caption = "Analyzing water data. Please wait!!!"
            f.Refresh
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pWTLayer, strResult & "_WT")
            'Set pResult = Get_IntersectLayer(pSoilLayer, pRoadLayer, strResult)
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            gMap.DeleteLayer pWTLayer
            Delete_Dataset_ST gWorkingfolder, pWTLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
            Else
                Set pSoilLayer = Nothing
            End If
        End If
        
        '*****************************************
        ' Intersect SOIL with ROAD.........................................
        If pSoilLayer Is Nothing Then GoTo Check
        If lblDC_RB.Text <> "NA" And chkRB.Enabled = True And chkRB.Value = vbChecked And Not pRoadLayer Is Nothing Then
            'Update the Status....
            f.lblStatus.Caption = "Analyzing soil data. Please wait!!!"
            f.Refresh
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pRoadLayer, strResult & "_RB")
            'Set pResult = Get_IntersectLayer(pSoilLayer, pRoadLayer, strResult)
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            gMap.DeleteLayer pRoadLayer
            Delete_Dataset_ST gWorkingfolder, pRoadLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
            Else
                Set pSoilLayer = Nothing
            End If
        'New buffer already removes the road polygon
''        ElseIf lblDC_RB.Text = "NA" And chkRB.Enabled = True And chkRB.Value = vbChecked And Not pRoadLayer Is Nothing Then
''            If pRoadLayer.FeatureClass.ShapeType = esriGeometryPolygon Then
''                'Update the Status....
''                f.lblStatus.Caption = "Excluding roads. Please wait!!!"
''                f.Refresh
''                ' Intersect the FeatureClasses......
''                Set pResult = Get_EraseLayer(pSoilLayer, pRoadLayer)
''                gMap.DeleteLayer pRoadLayer
''                If Not pResult Is Nothing Then
''                    Set pSoilLayer = pResult
''                Else
''                    Set pSoilLayer = Nothing
''                End If
''            End If
        End If
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
        '*****************************************
        ' Erase SOIL/RESULT with STREAM........................
        If pSoilLayer Is Nothing Then GoTo Check
        If lblDC_SB.Text <> "NA" And chkSB.Enabled = True And chkSB.Value = vbChecked And Not pStreamLayer Is Nothing Then
            'Update the Status....
            f.lblStatus.Caption = "Analyzing stream data. Please wait!!!"
            f.Refresh
            Stream_Flag = True
            ' Earse the FeatureClasses......
            'Set pResult = Get_EraseLayer(pSoilLayer, pStreamLayer)
            'Changed to intersect
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pStreamLayer, strResult & "_SB")
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            gMap.DeleteLayer pStreamLayer
            Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
            If Not pResult Is Nothing Then
                'ChangeStyle pResult
                Set pSoilLayer = pResult
'                gMap.DeleteLayer pStreamLayer
'                Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
            End If
        End If
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....

    Else
    
Skip_Soil:
        
        Skip_Soil_Flag = True
        
        ' *******************************************
        ' First intersect......
        ' *******************************************
        ' Intersect Road & Water layers......
        If lblDC_RB.Text <> "NA" Then
            If chkRB.Enabled = True And chkWT.Enabled = True And chkRB.Value = vbChecked And chkWT.Value = vbChecked Then
                If Not (pRoadLayer Is Nothing And pWTLayer Is Nothing) Then
                    ' Intersect the FeatureClasses......
                    Set pResult = Get_Intersect_FeatureLayer(pRoadLayer, pWTLayer, strResult & "_RB")
                    gMap.DeleteLayer pRoadLayer
                    Delete_Dataset_ST gWorkingfolder, pRoadLayer.Name
                    gMap.DeleteLayer pWTLayer
                    Delete_Dataset_ST gWorkingfolder, pWTLayer.Name
                    If Not pResult Is Nothing Then
                        'ChangeStyle pResult
                        Set pSoilLayer = pResult
                    End If
                    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
                End If
            End If
        End If
        If pSoilLayer Is Nothing Then
            If Not pRoadLayer Is Nothing And chkRB.Enabled = True And chkRB.Value = vbChecked And lblDC_RB.Text <> "NA" Then Set pSoilLayer = pRoadLayer: Set pRoadLayer = Nothing
            If Not pWTLayer Is Nothing And chkWT.Enabled = True And chkWT.Value = vbChecked And lblDC_WT.Text <> "NA" Then Set pSoilLayer = pWTLayer: Set pWTLayer = Nothing
        End If
        
        ' Erase result & Stream layers......
        If chkSB.Enabled = True And chkSB.Value = vbChecked And lblDC_SB.Text <> "NA" And Not pSoilLayer Is Nothing Then
            If Not pSoilLayer Is Nothing Then
                'Update the Status....
                f.lblStatus.Caption = "Excluding stream data. Please wait!!!"
                f.Refresh
                Stream_Flag = True
                'Set pResult = Get_EraseLayer(pSoilLayer, pStreamLayer)
                Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pStreamLayer, strResult & "_SB")
                gMap.DeleteLayer pSoilLayer
                Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
                gMap.DeleteLayer pStreamLayer
                Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
                
                If Not pResult Is Nothing Then
                    'ChangeStyle pResult
                    Set pSoilLayer = pResult
'                    gMap.DeleteLayer pStreamLayer
'                    Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
                End If
                If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
            End If
        End If
        
        ' ****************************************
        ' Exclude building Area..............................................
        ' ****************************************
        If lblDC_BB.Text = "NA" And chkBB.Enabled = True And chkBB.Value = vbChecked And UCase(cmbBMPType.Text) <> "GREEN ROOF" Then
            If Not pSoilLayer Is Nothing Then
                'Update the Status....
                f.lblStatus.Caption = "Excluding building data. Please wait!!!"
                f.Refresh
                Bldg_Flag = True
                ' Earse the FeatureClasses......
                Set pResult = Get_EraseLayer(pSoilLayer, pLanduseLayer)
                If Not pResult Is Nothing Then
                    Set pSoilLayer = pResult
                    gMap.DeleteLayer pLanduseLayer
                    Delete_Dataset_ST gWorkingfolder, pLanduseLayer.Name
                End If
            End If
        End If
        
        If pSoilLayer Is Nothing Then
            If chkBB.Enabled = True And chkBB.Value = vbChecked Then
                If UCase(cmbBMPType.Text) = "GREEN ROOF" Or UCase(cmbBMPType.Text) = "RAIN BARREL" Or UCase(cmbBMPType.Text) = "CISTERN" Then
                    Set pSoilLayer = pLanduseLayer
                    Set pLanduseLayer = Nothing
                End If
            End If
        End If
    
    End If
    

    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ' Pre-Process RASTER Data...........
    ' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
     ' ******************************************
    If lblDC_IMP.Text <> "NA" And chkIMP.Enabled = True And chkIMP.Value = vbChecked Then
    
        'Update the Status....
        f.lblStatus.Caption = "Analyzing impervious data. Please wait!!!"
        f.Refresh
        
        Set pResult = Nothing
        ' CHeck if the Impervious GRID is already cerated.....
        ' If Created then Use it......
        Dim pInImpLayer As IFeatureLayer
        Set pInImpLayer = GetFeatureLayer(gWorkingfolder, gImperviousdata & "_Ras")
        If pInImpLayer Is Nothing Then
            Dim pInShp_IMP As IFeatureClass
            Set pInImpLayer = New FeatureLayer
            Set pInShp_IMP = ConvertRastertoFeature(gRasterfolder, gImperviousdata, False, False)
            Set pInImpLayer.FeatureClass = pInShp_IMP
        End If
        
        ' Select the Features from the Soil Data.....
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "GRIDCODE " & lblDC_IMP.Text
        'Select all features of the feature class
        Set pFeatureSelection = pInImpLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultAdd, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count = 0 Then
            MsgBox "No features satisfy impervious criteria.", vbInformation, "BMP Siting Tool"
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            m_PassFlag = False
            GoTo Cleanup
        End If
        
        'Update the Status....
        f.lblStatus.Caption = "Filtering........... Please wait!!!"
        f.Refresh
        ' Export the Selected features to a new Shapefile.....
        Set pExportFclass = Get_ExportedShapefile(pInImpLayer, strResult & "IMP_Filter", True)
        gMap.ClearSelection ' Clear the Selection.....
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        ' Add the new Layer to map.........
        Set pInImpLayer = New FeatureLayer
        Set pInImpLayer.FeatureClass = pExportFclass
        pInImpLayer.Visible = True
        pInImpLayer.Name = pExportFclass.AliasName
        gMap.AddLayer pInImpLayer
        
        If Not pSoilLayer Is Nothing Then
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pInImpLayer, strResult & "_IMP")
            gMap.DeleteLayer pInImpLayer
            Delete_Dataset_ST gWorkingfolder, pInImpLayer.Name
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
                Set pInImpLayer = Nothing
            Else
                Set pSoilLayer = Nothing
            End If
        Else
            Set pSoilLayer = pInImpLayer
            Set pInImpLayer = Nothing
        End If
       
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    
    ' #######################
    ' Calculate the Slope from DEM...........
    Dim pInRasterLay As IRasterLayer
    Set pInRasterLay = GetInputFeatureLayer(gDEMdata)
    If Not pInRasterLay Is Nothing Then
        Dim pDEMRaster As IRaster
        Set pDEMRaster = pInRasterLay.Raster
        'Get the raster props
        Dim pDEMRasterProps As IRasterProps
        Set pDEMRasterProps = pDEMRaster
        'Get the raster cell size
        gCellSize = (pDEMRasterProps.MeanCellSize.x + pDEMRasterProps.MeanCellSize.y) / 2
        ' Convert the Cellsize to Feet....
        gCellSize = dScalefac * gCellSize
    End If

    ' First Calculate the Slope & Down stream area......
    If lblDC_DS.Text <> "NA" And chkDS.Enabled = True And chkDS.Value = vbChecked Then
        
        'Update the Status....
        f.lblStatus.Caption = "Analyzing slope data. Please wait!!!"
        f.Refresh
        
        Set pResult = Nothing
        gDACriteria = lblDC_DS.Text
        ' CHeck if the SLOPE GRID is already cerated.....
        ' If Created then Use it......
        Dim pInSlopeLayer As IFeatureLayer
        Dim pRaster As IRaster
        Set pRaster = OpenRasterDatasetFromDisk("SLOPE")
        If pRaster Is Nothing Then
            Dim pInShp_Slope As IFeatureClass
            Set pInSlopeLayer = New FeatureLayer
            'Update the Status....
            f.lblStatus.Caption = "Creating slope data. Please wait!!!"
            f.Refresh
            ' Calculate the Slope....
            Calculate_Slope pInRasterLay
            Set pInShp_Slope = ConvertRastertoFeature(gRasterfolder, "SLOPE", False, False)
            Set pInSlopeLayer.FeatureClass = pInShp_Slope
        Else
            Set pInShp_Slope = ConvertRastertoFeature(gRasterfolder, "SLOPE", True, False)
            Set pInSlopeLayer = GetFeatureLayer(gWorkingfolder, "SLOPE_Ras")
        End If
        
        ' Filter the Slope Layer.....
        If pInSlopeLayer Is Nothing Then GoTo Cleanup
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "GRIDCODE=1" '& gDACriteria '  GRIDCODE default field created when Converted Raster to Feature....
        Set pFeatureSelection = pInSlopeLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultNew, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count = 0 Then
            MsgBox "No features satisfy slope criteria.", vbInformation, "BMP Siting Tool"
            GoTo Cleanup
        End If
        
        ' Export the Selected features to a new Shapefile.....
        Set pExportFclass = Get_ExportedShapefile(pInSlopeLayer, "Slope_Filter", True)
        gMap.ClearSelection ' Clear the Selection.....
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        ' If the layer is added to map... Delete......
        Set pInSlopeLayer = GetFeatureLayer(gWorkingfolder, "Slope_Filter")
        If Not pInSlopeLayer Is Nothing Then gMap.DeleteLayer pInSlopeLayer
        ' Add the new Layer to map.........
        Set pInSlopeLayer = New FeatureLayer
        Set pInSlopeLayer.FeatureClass = pExportFclass
        pInSlopeLayer.Visible = True
        pInSlopeLayer.Name = pExportFclass.AliasName
        gMap.AddLayer pInSlopeLayer
        
        If Not pSoilLayer Is Nothing Then
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pInSlopeLayer, strResult & "_Slope")
            gMap.DeleteLayer pInSlopeLayer
            Delete_Dataset_ST gWorkingfolder, pInSlopeLayer.Name
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
                Set pInSlopeLayer = Nothing
            Else
                Set pSoilLayer = Nothing
            End If
        Else
            Set pSoilLayer = pInSlopeLayer
            Set pInSlopeLayer = Nothing
        End If
        
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    If lblDC_DA.Text <> "NA" And chkDA.Enabled = True And chkDA.Value = vbChecked And UCase(cmbBMPType.Text) <> "GREEN ROOF" Then
    
        'Update the Status....
        f.lblStatus.Caption = "Creating drainage area. Please wait!!!"
        f.Refresh
        
        Set pResult = Nothing
        gDACriteria = lblDC_DA.Text
        ' CHeck if the FLOW GRID is already cerated.....
        ' If Created then Use it......
        Dim pInFlowLayer As IFeatureLayer
        Set pRaster = OpenRasterDatasetFromDisk("FLOW")
        If pRaster Is Nothing Then
            Dim pInShp_Flow As IFeatureClass
            Set pInFlowLayer = New FeatureLayer
            ' Calculate the Flow....
            Create_FlowDirectionandAccumulation pInRasterLay.Raster ' Pass the DEM grid......
            Set pInShp_Flow = ConvertRastertoFeature(gRasterfolder, "FLOW", True, True)
            Set pInFlowLayer.FeatureClass = pInShp_Flow
        Else
            Set pInShp_Flow = ConvertRastertoFeature(gRasterfolder, "FLOW", True, True)
            Set pInFlowLayer = GetFeatureLayer(gWorkingfolder, "FLOW_Ras")
        End If

        'Update the Status....
        f.lblStatus.Caption = "Analyzing drainage area. Please wait!!!"
        f.Refresh
        
        ' Filter the Flow Layer.....
        If pInFlowLayer Is Nothing Then GoTo Cleanup
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "GRIDCODE=1" '& Parse_Expression(gDACriteria)
        Set pFeatureSelection = pInFlowLayer 'QI
        pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultAdd, False
        Set pSelectionSet = pFeatureSelection.SelectionSet
        If pSelectionSet.Count = 0 Then
            MsgBox "No features satisfy drainage area criteria.", vbInformation, "BMP Siting Tool"
            GoTo Cleanup
        End If
        
        ' Export the Selected features to a new Shapefile.....
        Set pExportFclass = Get_ExportedShapefile(pInFlowLayer, "Flow_Filter", True)
        gMap.ClearSelection ' Clear the Selection.....
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
        ' If the layer is added to map... Delete......
        Set pInFlowLayer = GetFeatureLayer(gWorkingfolder, "Flow_Filter")
        If Not pInFlowLayer Is Nothing Then gMap.DeleteLayer pInFlowLayer
        ' Add the new Layer to map.........
        Set pInFlowLayer = New FeatureLayer
        Set pInFlowLayer.FeatureClass = pExportFclass
        pInFlowLayer.Visible = True
        pInFlowLayer.Name = pExportFclass.AliasName
        gMap.AddLayer pInFlowLayer
        
        If Not pSoilLayer Is Nothing Then
            ' Intersect the FeatureClasses......
            Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pInFlowLayer, strResult & "_Flow")
            gMap.DeleteLayer pInFlowLayer
            Delete_Dataset_ST gWorkingfolder, pInFlowLayer.Name
            gMap.DeleteLayer pSoilLayer
            Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
            If Not pResult Is Nothing Then
                Set pSoilLayer = pResult
                Set pInFlowLayer = Nothing
            Else
                Set pSoilLayer = Nothing
            End If
        Else
            Set pSoilLayer = pInFlowLayer
            Set pInFlowLayer = Nothing
        End If
        
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    
    '########################
    ' Filter the MRLC Code............
    If lblDC_DA.Text <> "NA" And chkDA.Enabled = True And chkDA.Value = vbChecked And gMRLCdata <> "" And UCase(cmbBMPType.Text) <> "GREEN ROOF" Then
        
        'Update the Status....
        f.lblStatus.Caption = "Analyzing mrlc data. Please wait!!!"
        f.Refresh
        
        Set pResult = Nothing
        Dim pInmrlcLayer As IFeatureLayer
        Set pInmrlcLayer = GetFeatureLayer(gWorkingfolder, gMRLCdata)
        If Not pInmrlcLayer Is Nothing Then
            ' Build the Query String from joined Table......
            Dim strQry As String
            Dim pTable As ITable
            Set pQueryFilter = New QueryFilter
            Set pTable = pInmrlcLayer
            pQueryFilter.WhereClause = "SUITABLE = 1"
            Set pCursor = pTable.Search(pQueryFilter, False)
            Dim pRow As IRow
            Set pRow = pCursor.NextRow
            Do While Not pRow Is Nothing
                strQry = strQry & "," & pRow.Value(pRow.Fields.FindField("LUCODE"))
                Set pRow = pCursor.NextRow
            Loop
            strQry = Mid(strQry, 2)
            
            ' Filter the Slope Layer.....
            pQueryFilter.WhereClause = "GRIDCODE In (" & strQry & ")"
            Set pFeatureSelection = pInmrlcLayer 'QI
            pFeatureSelection.SelectFeatures pQueryFilter, esriSelectionResultAdd, False
            Set pSelectionSet = pFeatureSelection.SelectionSet
            If pSelectionSet.Count = 0 Then
                MsgBox "No features satisfy mrlc criteria.", vbInformation, "BMP Siting Tool"
                GoTo Cleanup
            End If
            
            ' Export the Selected features to a new Shapefile.....
            Set pExportFclass = Get_ExportedShapefile(pInmrlcLayer, strResult & "MRLC_Filter", True)
            gMap.ClearSelection ' Clear the Selection.....
            Set pInmrlcLayer = Nothing ' Clear pointer......
            ' Add the new Layer to map.........
            Set pInmrlcLayer = New FeatureLayer
            Set pInmrlcLayer.FeatureClass = pExportFclass
            pInmrlcLayer.Visible = True
            pInmrlcLayer.Name = pExportFclass.AliasName
            gMap.AddLayer pInmrlcLayer
            If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
            
            If Not pSoilLayer Is Nothing Then
                ' Intersect the FeatureClasses......
                Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pInmrlcLayer, strResult & "_MRLC")
                gMap.DeleteLayer pInmrlcLayer
                Delete_Dataset_ST gWorkingfolder, pInmrlcLayer.Name
                gMap.DeleteLayer pSoilLayer
                Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
                If Not pResult Is Nothing Then
                    Set pSoilLayer = pResult
                    Set pInmrlcLayer = Nothing
                Else
                    Set pSoilLayer = Nothing
                End If
            Else
                Set pSoilLayer = pInmrlcLayer
                Set pInmrlcLayer = Nothing
            End If
        End If
    End If
    If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....

    
    ' /////////////////////////////////////////////////////////////////////////////
    ' Finally exclude the Streams.......
    ' /////////////////////////////////////////////////////////////////////////////
    
    ' Erase result & Stream layers......
    If Stream_Flag Then GoTo Check
    If chkSB.Enabled = True And chkSB.Value = vbChecked And lblDC_SB.Text <> "NA" And Not pSoilLayer Is Nothing Then
        'Update the Status....
        'f.lblStatus.Caption = "Excluding stream data. Please wait!!!"
        f.lblStatus.Caption = "Intersecting with stream buffer. Please wait!!!"
        f.Refresh
        'Set pResult = Get_EraseLayer(pSoilLayer, pStreamLayer)
        Set pResult = Get_Intersect_FeatureLayer(pSoilLayer, pStreamLayer, strResult & "_SB")
        gMap.DeleteLayer pSoilLayer
        Delete_Dataset_ST gWorkingfolder, pSoilLayer.Name
        gMap.DeleteLayer pStreamLayer
        Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
        If Not pResult Is Nothing Then
            'ChangeStyle pResult
            Set pSoilLayer = pResult
'            gMap.DeleteLayer pStreamLayer
'            Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
        End If
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    End If
    
    'Erase result & Building layers......
    If Bldg_Flag Then GoTo Check
    If chkBB.Enabled = True And chkBB.Value = vbChecked And lblDC_BB.Text = "NA" And Not pSoilLayer Is Nothing And Not pLanduseLayer Is Nothing Then
        'Update the Status....
        f.lblStatus.Caption = "Excluding building data. Please wait!!!"
        f.Refresh
        Set pResult = Get_EraseLayer(pSoilLayer, pLanduseLayer)
        If Not pResult Is Nothing Then
            'ChangeStyle pResult
            Set pSoilLayer = pResult
            gMap.DeleteLayer pLanduseLayer
            Delete_Dataset_ST gWorkingfolder, pLanduseLayer.Name
        End If
        If Not m_PassFlag Then GoTo Cleanup ' Check if the Process is Successful....
    End If
    
Check:

    ' *******************************
    ' Check if any result is available..............
    ' *******************************
    If pSoilLayer Is Nothing Then
        MsgBox "No results found for " & cmbBMPType.Text, vbInformation, "BMP Siting Tool"
        GoTo Cleanup
    End If
    
    ' *******************************
    ' Clean the Final layer..............................
    ' Drop all fields and Add a result Field......
    ' *******************************
    pSoilLayer.Name = strResult
    If cmbBMPType.ListCount > 1 Then Set pSoilLayer = Clean_Layer(pSoilLayer)
    ChangeStyle pSoilLayer
    
    ' Add the result to the Collection.....
    pcolResult.Add strResult, strResult
    
    ' Display the result......
    If Not pFeatureSelection Is Nothing Then pFeatureSelection.Clear
    pActiveView.Refresh
    gMxDoc.UpdateContents
        
Cleanup:
    
     ' Clean any layers....
     Set pRoadLayer = GetInputFeatureLayer(gRoaddata & "_Dis")
     Set pStreamLayer = GetInputFeatureLayer(gStreamdata & "_Dis")
     If Not pWTLayer Is Nothing Then gMap.DeleteLayer pWTLayer
    If Not pRoadLayer Is Nothing Then gMap.DeleteLayer pRoadLayer: Delete_Dataset_ST gWorkingfolder, pRoadLayer.Name
    If Not pStreamLayer Is Nothing Then gMap.DeleteLayer pStreamLayer: Delete_Dataset_ST gWorkingfolder, pStreamLayer.Name
    If Not pLanduseLayer Is Nothing Then gMap.DeleteLayer pLanduseLayer: Delete_Dataset_ST gWorkingfolder, pLanduseLayer.Name
    If Not pInFlowLayer Is Nothing Then gMap.DeleteLayer pInFlowLayer
    If Not pInImpLayer Is Nothing Then gMap.DeleteLayer pInImpLayer
    If Not pInmrlcLayer Is Nothing Then gMap.DeleteLayer pInmrlcLayer
    Dim pLayer As IFeatureLayer
    Set pLayer = GetInputFeatureLayer(gRoaddata & "_BufMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gRoaddata & "_BufMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gRoaddata & "_DisMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gRoaddata & "_DisMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gRoaddata & "_Buf")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    
    Set pLayer = GetInputFeatureLayer(gStreamdata & "_BufMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gStreamdata & "_BufMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gStreamdata & "_DisMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gStreamdata & "_DisMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gStreamdata & "_Buf")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    
    Set pLayer = GetInputFeatureLayer(gLandusedata & "_BufMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gLandusedata & "_BufMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gLandusedata & "_DisMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gLandusedata & "_DisMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(gLandusedata & "_Buf")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    
    Set pLayer = GetInputFeatureLayer(strResult & "BB_Filter_DisMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(strResult & "BB_Filter_DisMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(strResult & "BB_Filter_BufMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(strResult & "BB_Filter_BufMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(strResult & "BB_Filter")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    
    Set pLayer = Nothing
    Delete_Dataset_ST gWorkingfolder, gRoaddata & "_BufMax"
    Delete_Dataset_ST gWorkingfolder, gRoaddata & "_BufMin"
    Delete_Dataset_ST gWorkingfolder, gRoaddata & "_DisMax"
    Delete_Dataset_ST gWorkingfolder, gRoaddata & "_DisMin"

    Delete_Dataset_ST gWorkingfolder, gStreamdata & "_BufMax"
    Delete_Dataset_ST gWorkingfolder, gStreamdata & "_BufMin"
    Delete_Dataset_ST gWorkingfolder, gStreamdata & "_DisMax"
    Delete_Dataset_ST gWorkingfolder, gStreamdata & "_DisMin"
    
    Delete_Dataset_ST gWorkingfolder, gLandusedata & "_BufMax"
    Delete_Dataset_ST gWorkingfolder, gLandusedata & "_BufMin"
    Delete_Dataset_ST gWorkingfolder, gLandusedata & "_DisMax"
    Delete_Dataset_ST gWorkingfolder, gLandusedata & "_DisMin"
    
    Delete_Dataset_ST gWorkingfolder, strResult & "BB_Filter_DisMin"
    Delete_Dataset_ST gWorkingfolder, strResult & "BB_Filter_DisMax"
    Delete_Dataset_ST gWorkingfolder, strResult & "BB_Filter_BufMin"
    Delete_Dataset_ST gWorkingfolder, strResult & "BB_Filter_BufMax"
    Delete_Dataset_ST gWorkingfolder, strResult & "BB_Filter"
    
    ' Cleanup.....
    Set pActiveView = Nothing
    Set pCalc = Nothing
    Set pCursor = Nothing
    Set pDEMRaster = Nothing
    Set pExportFclass = Nothing
    Set pFeatureSelection = Nothing
    Set pFieldEdit = Nothing
    Set pInFlowLayer = Nothing
    Set pInImpLayer = Nothing
    Set pInmrlcLayer = Nothing
    Set pInRasterLay = Nothing
    Set pInShp_Flow = Nothing
    Set pInShp_Slope = Nothing
    Set pInSlopeLayer = Nothing
    Set pLanduseLayer = Nothing
    Set pQueryFilter = Nothing
    Set pRaster = Nothing
    Set pResult = Nothing
    Set pSelectionSet = Nothing
    Set pSelset = Nothing
    Set pTable = Nothing
    Set ptmpLayer = Nothing
    Set pValues = Nothing
    Set pSoilLayer = Nothing
    Set pRoadLayer = Nothing
    Set pStreamLayer = Nothing
    
    
Exit Sub
ErrorHandler:
  m_PassFlag = False
  HandleError True, "Start_Analyze " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
  GoTo Cleanup

End Sub

Private Function Clean_Layer(ByVal pLayer As IFeatureLayer) As IFeatureLayer
    
    On Error GoTo ErrorHandler
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = pLayer.FeatureClass
    
    ' **************************************************************
    ' Seems to be problem deletin/adding fields with result.........
    ' Copy the result Featureclass and Clean the featureClass.......
    ' **************************************************************
    Dim pDataset As IDataset
    Set pDataset = pFeatClass
    If pDataset.CanCopy Then
        Dim pWkspace As IWorkspace
        Set pWkspace = GetWorkspace(gWorkingfolder)
        pDataset.Copy strResult, pWkspace
        Set pWkspace = Nothing
    End If
    Set pFeatClass = OpenShapeFile(gWorkingfolder, strResult)
    
    ' Now add the Field.....
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
    Set pField = New Field
    Set pFieldEdit = pField
    pFieldEdit.Name = "BMP"
    pFieldEdit.AliasName = "BMP"
    pFieldEdit.Type = esriFieldType.esriFieldTypeString
    pFieldEdit.Length = 50
    pFeatClass.AddField pField
    
    ' Now Calulate the new field value.............
    Dim pCalc As ICalculator
    Set pCalc = New Calculator
    Dim pCursor As ICursor
    Set pCursor = pFeatClass.Update(Nothing, True)
    With pCalc
      Set .Cursor = pCursor
        .Expression = """" & strResult & """"
        .Field = "BMP"
    End With
    pCalc.Calculate ' Execute the Query.......
  
    ' Delete all fields from the featureClass......
    Dim pFields As IFields
    Dim l As Long
    Set pFields = pFeatClass.Fields
    For l = pFields.FieldCount - 1 To 0 Step -1
      If pFields.Field(l).Type <> esriFieldTypeOID And pFields.Field(l).Type <> esriFieldTypeGeometry And pFields.Field(l).Name <> "BMP" Then
        pFeatClass.DeleteField pFields.Field(l)
      End If
    Next l
    
    ' Remove the old Layer.....
    gMap.DeleteLayer pLayer
    Set pLayer = New FeatureLayer
    Set pLayer.FeatureClass = pFeatClass
    pLayer.Name = strResult
    pLayer.Visible = True
    'Add the New Layer....
    gMap.AddLayer pLayer
    ChangeStyle pLayer
    
    Set Clean_Layer = pLayer
    
Exit Function
ErrorHandler:

  HandleError True, "Clean_Layer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

  
End Function

Private Sub Create_Composite_Layer(ByVal pcolResult As Scripting.Dictionary, ByVal f As frmSplash)
        
    On Error GoTo ErrorHandler
    
    '*******************************************
    ' Union the Features to make one feature....
    ' ******************************************
    
    'Update the Status....
    f.lblStatus.Caption = "Creating composite layer. Please wait!!!"
    f.Refresh
    
    Dim pResult As IFeatureLayer
    Dim strLayers As String
    Dim pKeys
    pKeys = pcolResult.Keys
    Dim pKey As String
    Dim iKey As Integer
    Dim strDissolve_Flds As String
    ' InitialiZe the Stings...........
    strLayers = "''"
    strDissolve_Flds = "BMP"

    ' Get the ARC installation path....
    Dim strInstallPath As String
    strInstallPath = GetArcGISPath
        
    ' Loop and build the Dissolve fields......
    For iKey = 0 To pcolResult.Count - 1
        pKey = pKeys(iKey)
        'strLayers = strLayers & ";'" & gWorkingfolder & "\" & pcolResult.Item(pKey) & ".shp' ''"
        If iKey > 0 Then strDissolve_Flds = strDissolve_Flds & ";" & "BMP_" & iKey
    Next
    
    ' Clean up the Map......
    Dim pDataset As IDataset
    Set pDataset = GetFeatureLayer(gWorkingfolder, "Union")
    If Not pDataset Is Nothing Then
        If pDataset.CanDelete Then pDataset.Delete
    End If
        
    'Set the toolbox
    GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
    ' Take first 2 items and Merge and continue with the resultant..............
    pKey = pKeys(0)
    strLayers = strLayers & "; '" & gWorkingfolder & "\" & pcolResult.Item(pKey) & ".shp' ''"
    pKey = pKeys(1)
    strLayers = strLayers & "; '" & gWorkingfolder & "\" & pcolResult.Item(pKey) & ".shp' ''"
    strLayers = Mid(strLayers, 5) ' trim the leading Delim.....
    '# Union 3 other feature classes, but specify some ranks for each since parcels has better spatial accuracy
    GP.Union strLayers, gWorkingfolder & "\Union_1.shp", "ALL", "#", "GAPS"
    
    ' Loop and Merge all results......
    strLayers = "''"
    pKeys = pcolResult.Keys
    For iKey = 2 To pcolResult.Count - 1
        pKey = pKeys(iKey)
        strLayers = "'" & gWorkingfolder & "\Union_" & iKey - 1 & ".shp' ''; '" & gWorkingfolder & "\" & pcolResult.Item(pKey) & ".shp' ''"
        '# Union 3 other feature classes, but specify some ranks for each since parcels has better spatial accuracy
        GP.Union strLayers, gWorkingfolder & "\Union_" & iKey & ".shp", "ALL", "#", "GAPS"
        
        ' Rename the Fields.....
        Dim pTable As ITable
        Set pTable = GetFeatureLayer(gWorkingfolder, "Union_" & iKey)
        Dim pFields As IFields
        Dim pFieldEdit As IFieldEdit
        Set pFields = pTable.Fields
        
        ' Loop through the Fields and Find BMP and rename...
        Dim iCnt As Integer, iFldCnt As Integer
        iFldCnt = 1
        For iCnt = 0 To pFields.FieldCount - 1
            If InStr(1, pFields.Field(iCnt).Name, "BMP_", vbTextCompare) > 0 Then
                Set pFieldEdit = pFields.Field(iCnt)
                pFieldEdit.Name = "BMP_" & iFldCnt
                Set pFieldEdit = Nothing
                iFldCnt = iFldCnt + 1 ' Increment the Field flag....
            End If
        Next iCnt
        
        ' Remove the Previous Union layer.....
        Set pResult = GetInputFeatureLayer("Union_" & iKey - 1)
        If Not pResult Is Nothing Then gMap.DeleteLayer pResult: Delete_Dataset_ST gWorkingfolder, "Union_" & iKey - 1
    Next
    
    ' Delete any composite layer............
    Set pDataset = GetFeatureLayer(gWorkingfolder, "Composite")
    If Not pDataset Is Nothing Then
        If pDataset.CanDelete Then pDataset.Delete
    End If
    
    ' # Dissolve the polygons to simplify.............
    'Set the toolbox
    GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
    '# Union 3 other feature classes, but specify some ranks for each since parcels has better spatial accuracy
    GP.Dissolve "Union_" & iKey - 1, gWorkingfolder & "\Composite.shp", strDissolve_Flds, "#", "MULTI_PART"
    'Executing: Dissolve Union "D:\Arun Raj\Work-TT\Projects\Sustain\TestData\SitingTool\Dissolve.shp" BMP;BMP_1 # MULTI_PART

    ' Delete the Combined Layer......
    Set pResult = GetInputFeatureLayer("Union_" & iKey - 1)
    If Not pResult Is Nothing Then gMap.DeleteLayer pResult: Delete_Dataset_ST gWorkingfolder, "Union_" & iKey - 1
    
    ' # If an error occurred when running Union, print out the error message.
    If GP.GetMessages(2) <> "" Then MsgBox GP.GetMessages(2)
    
    ' Change Style of composite Layer.....
    Set pResult = GetInputFeatureLayer("Composite")
    RenderUniqueValueFillSymbol_ST pResult, strDissolve_Flds, "Result"
    
Cleanup:

    Set pcolResult = Nothing
    Set pDataset = Nothing
    Set pTable = Nothing
    Set pResult = Nothing
        
Exit Sub
ErrorHandler:

  HandleError True, "Create_Composite_Layer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub


Private Function Copy_Data_toWorkfolder() As Boolean
    
    On Error GoTo ErrorHandler
    Dim pFeatClass As IFeatureClass
    Dim pFeatSel As IFeatureSelection
    Dim pDataset As IDataset
    Dim pFeatlayer As IFeatureLayer
    Dim fso As New FileSystemObject
    Copy_Data_toWorkfolder = True
    
    'Delete the Working folder if exists. & recreate...
    If fso.FolderExists(gWorkingfolder) Then fso.DeleteFolder gWorkingfolder, True
    fso.CreateFolder gWorkingfolder
    
    ' *************************************************
    ' Create a Cache Folder to store the Raster output.
    ' *************************************************
    If Not fso.FolderExists(gRasterfolder) Then
        fso.CreateFolder gRasterfolder
    Else
        Call Copy_from_Cache
    End If
    
    Dim pWkspace As IWorkspace
    Set pWkspace = GetWorkspace(gWorkingfolder)
    
    'Copy Soil Data
    Set pFeatlayer = GetInputFeatureLayer(gSoildata)
    If Not pFeatlayer Is Nothing Then
        If IsGeodatabaseFClass(pFeatlayer) Then
            Set pFeatSel = pFeatlayer
            pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
            Set pFeatClass = Get_ExportedShapefile(pFeatlayer, gSoildata, False)
        Else
            Set pFeatClass = pFeatlayer.FeatureClass
            Set pDataset = pFeatClass
            If pDataset.CanCopy Then
                pDataset.Copy gSoildata, pWkspace
            End If
        End If
    End If
    'Copy Road Data
    Set pFeatlayer = GetInputFeatureLayer(gRoaddata)
    If Not pFeatlayer Is Nothing Then
        If IsGeodatabaseFClass(pFeatlayer) Then
            Set pFeatSel = pFeatlayer
            pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
            Set pFeatClass = Get_ExportedShapefile(pFeatlayer, gRoaddata, False)
        Else
            Set pFeatClass = pFeatlayer.FeatureClass
            Set pDataset = pFeatClass
            If pDataset.CanCopy Then
                pDataset.Copy gRoaddata, pWkspace
            End If
        End If
    End If
    'Copy Landuse Data
    Set pFeatlayer = GetInputFeatureLayer(gLandusedata)
    If Not pFeatlayer Is Nothing Then
        If IsGeodatabaseFClass(pFeatlayer) Then
            Set pFeatSel = pFeatlayer
            pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
            Set pFeatClass = Get_ExportedShapefile(pFeatlayer, gLandusedata, False)
        Else
            Set pFeatClass = pFeatlayer.FeatureClass
            Set pDataset = pFeatClass
            If pDataset.CanCopy Then
                pDataset.Copy gLandusedata, pWkspace
            End If
        End If
    End If
    'Copy Stream Data
    Set pFeatlayer = GetInputFeatureLayer(gStreamdata)
    If Not pFeatlayer Is Nothing Then
        If IsGeodatabaseFClass(pFeatlayer) Then
            Set pFeatSel = pFeatlayer
            pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
            Set pFeatClass = Get_ExportedShapefile(pFeatlayer, gStreamdata, False)
        Else
            Set pFeatClass = pFeatlayer.FeatureClass
            Set pDataset = pFeatClass
            If pDataset.CanCopy Then
                pDataset.Copy gStreamdata, pWkspace
            End If
        End If
    End If
    'Copy WaterTable Data
    Set pFeatlayer = GetInputFeatureLayer(gWTdata)
    If Not pFeatlayer Is Nothing Then
        If IsGeodatabaseFClass(pFeatlayer) Then
            Set pFeatSel = pFeatlayer
            pFeatSel.SelectFeatures Nothing, esriSelectionResultNew, False
            Set pFeatClass = Get_ExportedShapefile(pFeatlayer, gWTdata, False)
        Else
            Set pFeatClass = pFeatlayer.FeatureClass
            Set pDataset = pFeatClass
            If pDataset.CanCopy Then
                pDataset.Copy gWTdata, pWkspace
            End If
        End If
    End If
    'Copy Slope Grid.......
    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = GetInputFeatureLayer(gDEMdata)
    If Not pRasterLayer Is Nothing Then
        Copy_Raster_Data gDEMdata, fso.GetParentFolderName(pRasterLayer.FilePath), gRasterfolder
    End If
    'Copy MRLC Grid.......
    Set pRasterLayer = GetInputFeatureLayer(gMRLCdata)
    If Not pRasterLayer Is Nothing Then
        Copy_Raster_Data gMRLCdata, fso.GetParentFolderName(pRasterLayer.FilePath), gRasterfolder
    End If
    'Copy IMPERVIOUS Grid.......
    Set pRasterLayer = GetInputFeatureLayer(gImperviousdata)
    If Not pRasterLayer Is Nothing Then
        Copy_Raster_Data gImperviousdata, fso.GetParentFolderName(pRasterLayer.FilePath), gRasterfolder
    End If
    
Cleanup:
    Set pWkspace = Nothing

  Exit Function
ErrorHandler:
  HandleError True, "Copy_Data_toWorkfolder " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
  Copy_Data_toWorkfolder = False

End Function

Private Function IsGeodatabaseFClass(pDLayer As IDataLayer2) As Boolean
    Dim pDSName As IDatasetName
    Set pDSName = pDLayer.DataSourceName
    
    If pDSName.Category = "Personal Geodatabase Feature Class" Then
        IsGeodatabaseFClass = True
    End If
                        
End Function


Private Sub Copy_from_Cache()
    
    On Error GoTo ErrorHandler
    
    Dim pFeatClass As IFeatureClass
    Dim pDataset As IDataset
    Dim pWkspace As IWorkspace
    Set pWkspace = GetWorkspace(gWorkingfolder)
    'Copy Slope Data
    Set pFeatClass = OpenShapeFile(gRasterfolder, "Slope_Ras")
    If Not pFeatClass Is Nothing Then
        Set pDataset = pFeatClass
        If pDataset.CanCopy Then
            pDataset.Copy "Slope_Ras", pWkspace
        End If
    End If
    'Copy Flow Data
    Set pFeatClass = OpenShapeFile(gRasterfolder, "Flow_Ras")
    If Not pFeatClass Is Nothing Then
        Set pDataset = pFeatClass
        If pDataset.CanCopy Then
            pDataset.Copy "Flow_Ras", pWkspace
        End If
    End If
    
Cleanup:
    Set pWkspace = Nothing

Exit Sub
ErrorHandler:
  HandleError True, "Copy_from_Cache " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Sub Copy_Raster_Data(strRaster As String, strCopyFrom As String, strCopyTo As String)
    
    On Error GoTo ErrorHandler
    'create workspace
    Dim pRWorkSpace As IRasterWorkspace
    Dim pRWSF As IWorkspaceFactory
    Set pRWSF = New RasterWorkspaceFactory

    Set pRWorkSpace = pRWSF.OpenFromFile(strCopyFrom, 0)
    Dim pRDataset As IRasterDataset
    Set pRDataset = pRWorkSpace.OpenRasterDataset(strRaster)
    If pRDataset.CanCopy = True Then
        Dim pCopyRWSF As IWorkspaceFactory
        Set pCopyRWSF = New RasterWorkspaceFactory
        Dim pCopyRWorks As IRasterWorkspace
        Set pCopyRWorks = pCopyRWSF.OpenFromFile(strCopyTo, 0)
        Dim fso As New FileSystemObject
        If Not fso.FolderExists(strCopyTo & "\" & strRaster) Then pRDataset.Copy strRaster, pCopyRWorks
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "Copy_Raster_Data " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub


Private Function Get_Scale_Factor(ByVal UnitName As esriUnits) As Double

   On Error GoTo ErrorHandler
   
   Select Case UnitName
        
        Case esriUnits.esriFeet
            Get_Scale_Factor = 1
        Case esriUnits.esriMeters
            Get_Scale_Factor = 0.3048
        Case esriUnits.esriInches
            Get_Scale_Factor = 12
        Case esriUnits.esriCentimeters
            Get_Scale_Factor = 30.48
        Case esriUnits.esriKilometers
            Get_Scale_Factor = 0.0003048009
        Case esriUnits.esriMiles
            Get_Scale_Factor = 0.0001893939
        Case esriUnits.esriMillimeters
            Get_Scale_Factor = 304.8
        Case esriUnits.esriUnknownUnits
            Get_Scale_Factor = 1
        
    End Select
            
  Exit Function
  
ErrorHandler:
  HandleError True, "Get_Scale_Factor " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function

Private Sub Turn_Off_Layers(ByVal pcolResult As Scripting.Dictionary)

    On Error GoTo ErrorHandler
    Dim iCnt As Integer
    
    If pcolResult Is Nothing Then Exit Sub
    For iCnt = 0 To gMap.LayerCount - 1
        If Not pcolResult.Exists(gMap.Layer(iCnt).Name) Then
            gMap.Layer(iCnt).Visible = False
        End If
    Next iCnt
    Dim pActView As IActiveView
    Set pActView = gMap
    pActView.Refresh
    gMxDoc.UpdateContents

  Exit Sub
ErrorHandler:
  HandleError True, "Turn_Off_Layers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Function Validate_Input_Data() As Boolean
        
    On Error GoTo ErrorHandler
    Validate_Input_Data = True
    
    ' *****************************
    ' Validate Data
    ' *****************************
    
    
    
    ' *****************************
    ' Validate Inputs
    ' *****************************
    


    ' *****************************
    ' Validate Criteria
    ' *****************************
    
    

  Exit Function
ErrorHandler:
  HandleError True, "Validate_Input_Data " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function


Private Sub Form_Load()

    On Error GoTo ErrorHandler
   
   ' Refresh Controls....
   Call Initialize_List
    
    ' Now validate the layers and add to the Form......
    If gDEMdata <> "" Then
        If m_LayerDict.Exists(gDEMdata) And Not GetInputFeatureLayer(gDEMdata) Is Nothing Then
            txtDEMpath(1).Text = gDEMdata
        End If
    End If
    If gMRLCdata <> "" Then
        If m_LayerDict.Exists(gMRLCdata) And Not GetInputFeatureLayer(gMRLCdata) Is Nothing Then
            txtMRLC(1).Text = gMRLCdata
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
    If gStreamdata <> "" Then
        If m_LayerDict.Exists(gStreamdata) And Not GetInputFeatureLayer(gStreamdata) Is Nothing Then
            txtStreampath(1).Text = gStreamdata
        End If
    End If
    If gImperviousdata <> "" Then
        If m_LayerDict.Exists(gImperviousdata) And Not GetInputFeatureLayer(gImperviousdata) Is Nothing Then
            txtImp(1).Text = gImperviousdata
        End If
    End If
    If gWTdata <> "" Then
        If m_LayerDict.Exists(gWTdata) And Not GetInputFeatureLayer(gWTdata) Is Nothing Then
            txtWTPath(1).Text = gWTdata
        End If
    End If
    If gSoilTable <> "" Then
        If m_LayerDict.Exists(gSoilTable) And Not GetInputTable(gSoilTable) Is Nothing Then
            txtSoil_lk.Text = gSoilTable
        End If
    End If
    If gMrlcTable <> "" Then
        If m_LayerDict.Exists(gMrlcTable) And Not GetInputTable(gMrlcTable) Is Nothing Then
            txtMrlc_lk.Text = gMrlcTable
        End If
    End If
    
    ' Disable the Design Tab..........
    m_backFlag = False
    tbsBMP.TabEnabled(0) = True
    tbsBMP.TabEnabled(1) = False
    tbsBMP.TabEnabled(2) = False
    tbsBMP.Tab = 0
    m_Prev_BMP = ""
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

Private Sub Initialize_List()
    
    On Error GoTo ErrorHandler
            
    ' Create the Layer Dictionary.....
    Set m_LayerDict = CreateObject("Scripting.Dictionary")
    m_LayerDict.RemoveAll
    ' add a blank item to all combos....
    txtLandusepath(1).AddItem ""
    txtRoadpath(1).AddItem ""
    txtSoilpath(1).AddItem ""
    txtStreampath(1).AddItem ""
    txtDEMpath(1).AddItem ""
    txtMRLC(1).AddItem ""
    txtImp(1).AddItem ""
    txtWTPath(1).AddItem ""
    txtSoil_lk.AddItem ""
    txtMrlc_lk.AddItem ""

    
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
            txtWTPath(1).AddItem pLayer.Name
            txtLandusepath(1).Visible = True
            txtRoadpath(1).Visible = True
            txtSoilpath(1).Visible = True
            txtStreampath(1).Visible = True
            txtWTPath(1).Visible = True
        ElseIf TypeOf pLayer Is IRasterLayer Then
            txtDEMpath(1).AddItem pLayer.Name
            txtDEMpath(1).Visible = True
            txtMRLC(1).AddItem pLayer.Name
            txtMRLC(1).Visible = True
            txtImp(1).AddItem pLayer.Name
            txtImp(1).Visible = True
        End If
        m_LayerDict.Add pLayer.Name, pLayer
    Next
    ' Now add the Standalone tables......
    Dim pStandCol As IStandaloneTableCollection
    Dim pStTab As IStandaloneTable
    Set pStandCol = gMap
    For i = 0 To pStandCol.StandaloneTableCount - 1
      If Not m_LayerDict.Exists(pStandCol.StandaloneTable(i).Name) Then
        txtSoil_lk.AddItem pStandCol.StandaloneTable(i).Name
        txtMrlc_lk.AddItem pStandCol.StandaloneTable(i).Name
        m_LayerDict.Add pStandCol.StandaloneTable(i).Name, pStandCol.StandaloneTable(i).Table
      End If
    Next

  Exit Sub
ErrorHandler:
  HandleError True, "Initialize_List " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub Initialize_BMPs()

    On Error GoTo ErrorHandler
    ' Now Create the BMP Type objects with Props.....
    Dim oBMP As BMPobj

    Set oBMP = New BMPobj
    oBMP.BMPName = "Dry pond"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 1
    ' DC
    oBMP.DC_DA = ">10"
    oBMP.DC_DS = "<15"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">4"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    'Add to Dictionary......
    gBMPtypeDict.Add "Dry pond", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Wet pond"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 2
    ' DC
    oBMP.DC_DA = ">25"
    oBMP.DC_DS = "<15"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">4"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    'Add to Dictionary......
    gBMPtypeDict.Add "Wet pond", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Infiltration basin"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 3
    ' DC
    oBMP.DC_DA = "<10"
    oBMP.DC_DS = "<15"
    oBMP.DC_HG = "A-B"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">4"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    'Add to Dictionary......
    gBMPtypeDict.Add "Infiltration basin", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Infiltration trench"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 4
    ' DC
    oBMP.DC_DA = "<5"
    oBMP.DC_DS = "<15"
    oBMP.DC_HG = "A-B"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">4"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    'Add to Dictionary......
    gBMPtypeDict.Add "Infiltration trench", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Bioretention"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 5
    ' DC
    oBMP.DC_DA = "<2"
    oBMP.DC_DS = "<5"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "<100"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    ' Add to the Dict
    gBMPtypeDict.Add "Bioretention", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Sand filter (surface)"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 6
    ' DC
    oBMP.DC_DA = "<10"
    oBMP.DC_DS = "<10"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Sand filter (surface)", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Sand filter (non-surface)"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 7
    ' DC
    oBMP.DC_DA = "<2"
    oBMP.DC_DS = "<10"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Sand filter (non-surface)", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Constructed wetland"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 8
    ' DC
    oBMP.DC_DA = ">25"
    oBMP.DC_DS = "<15"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">4"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = ">100"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 1
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Constructed wetland", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Porous Pavement"
    oBMP.BMPType = "Area"
    oBMP.BMPId = 9
    ' DC
    oBMP.DC_DA = "<3"
    oBMP.DC_DS = "<1"
    oBMP.DC_HG = "A-B"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Porous Pavement", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Grassed swales"
    oBMP.BMPType = "Line"
    oBMP.BMPId = 10
    ' DC
    oBMP.DC_DA = "<5"
    oBMP.DC_DS = "<4"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "<100"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 1
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Grassed swales", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Vegetated filterstrip"
    oBMP.BMPType = "Line"
    oBMP.BMPId = 11
    ' DC
    oBMP.DC_DA = "NA"
    oBMP.DC_DS = "<10"
    oBMP.DC_HG = "A-D"
    oBMP.DC_IMP = ">0"
    oBMP.DC_WT = ">2"
    oBMP.DC_RB = "<100"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 0
    oBMP.DC_DS_State = 1
    oBMP.DC_HG_State = 1
    oBMP.DC_WT_State = 1
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 1
    gBMPtypeDict.Add "Vegetated filterstrip", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Rain barrel"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 12
    ' DC
    oBMP.DC_DA = "NA"
    oBMP.DC_DS = "NA"
    oBMP.DC_HG = "NA"
    oBMP.DC_IMP = "NA"
    oBMP.DC_WT = "NA"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "<30"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 0
    oBMP.DC_DS_State = 0
    oBMP.DC_HG_State = 0
    oBMP.DC_WT_State = 0
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 0
    gBMPtypeDict.Add "Rain barrel", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Cistern"
    oBMP.BMPType = "Point"
    oBMP.BMPId = 13
    ' DC
    oBMP.DC_DA = "NA"
    oBMP.DC_DS = "NA"
    oBMP.DC_HG = "NA"
    oBMP.DC_IMP = "NA"
    oBMP.DC_WT = "NA"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "<30"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 0
    oBMP.DC_DS_State = 0
    oBMP.DC_HG_State = 0
    oBMP.DC_WT_State = 0
    oBMP.DC_RB_State = 1
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 0
    gBMPtypeDict.Add "Cistern", oBMP
    
    Set oBMP = New BMPobj
    oBMP.BMPName = "Green roof"
    oBMP.BMPType = "Area"
    oBMP.BMPId = 14
    ' DC
    oBMP.DC_DA = "NA"
    oBMP.DC_DS = "NA"
    oBMP.DC_HG = "NA"
    oBMP.DC_IMP = "NA"
    oBMP.DC_WT = "NA"
    oBMP.DC_RB = "NA"
    oBMP.DC_BB = "NA"
    oBMP.DC_SB = "NA"
    oBMP.DC_DA_State = 0
    oBMP.DC_DS_State = 0
    oBMP.DC_HG_State = 0
    oBMP.DC_WT_State = 0
    oBMP.DC_RB_State = 0
    oBMP.DC_BB_State = 1
    oBMP.DC_SB_State = 0
    oBMP.DC_IMP_State = 0
    gBMPtypeDict.Add "Green roof", oBMP

    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "Initialize_BMPs" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Sub

Private Function Get_Soil_Groups(ByVal strSoilRange As String, ByVal strDelim As String) As Collection

    On Error GoTo ErrorHandler
    Dim Group_Range As Collection
    Set Group_Range = New Collection
    Group_Range.Add "A", "A"
    Group_Range.Add "B", "B"
    Group_Range.Add "C", "C"
    Group_Range.Add "D", "D"
    Group_Range.Add "E", "E"
    Group_Range.Add "F", "F"
    Group_Range.Add "G", "G"
    Group_Range.Add "H", "H"
    Group_Range.Add "I", "I"
    Group_Range.Add "J", "J"
    Group_Range.Add "K", "K"
    Group_Range.Add "L", "L"
    Group_Range.Add "M", "M"
    Group_Range.Add "N", "N"
    Group_Range.Add "O", "O"
    Group_Range.Add "P", "P"
    Group_Range.Add "Q", "Q"
    Group_Range.Add "R", "R"
    Group_Range.Add "S", "S"
    Group_Range.Add "T", "T"
    Group_Range.Add "U", "U"
    Group_Range.Add "V", "V"
    Group_Range.Add "W", "W"
    Group_Range.Add "X", "X"
    Group_Range.Add "Y", "Y"
    Group_Range.Add "Z", "Z"
    
    Dim AddFlag As Boolean
    Dim iCnt As Integer
    Dim strSoilRange1 As String
    Dim strSoilRange2 As String
    Dim pVals As Variant
    Set Get_Soil_Groups = New Collection
    
    pVals = Split(strSoilRange, "-")
    strSoilRange1 = pVals(0)
    strSoilRange2 = pVals(1)
    For iCnt = 1 To Group_Range.Count
        If Group_Range.Item(iCnt) = strSoilRange1 Then AddFlag = True
        If AddFlag Then Get_Soil_Groups.Add Group_Range.Item(iCnt)
        If Group_Range.Item(iCnt) = strSoilRange2 Or strSoilRange2 = "" Then AddFlag = False
    Next
    
    ' *************************************
    ' Now Handle the "/" in the attribute Table.........
    ' *************************************
    Dim tmpCol As Collection
    Set tmpCol = New Collection
    ' Transfer the contents to the Tmp collection......
    For iCnt = 1 To Get_Soil_Groups.Count
        tmpCol.Add Get_Soil_Groups.Item(iCnt)
    Next

    Do While tmpCol.Count > 1
        For iCnt = 2 To tmpCol.Count
            Get_Soil_Groups.Add tmpCol.Item(1) & "/" & tmpCol.Item(iCnt)
        Next iCnt
        tmpCol.Remove (1)
    Loop
        
    
Cleanup:

  Exit Function
ErrorHandler:
  HandleError True, "Get_Soil_Groups" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

End Function

Private Function Get_Intersect_FeatureLayer(ByVal pInputFeatLayer As IFeatureLayer, ByVal pOverlayLayer As IFeatureLayer, strResultName As String) As IFeatureLayer

    On Error GoTo ErrorHandler

    
'      ' Use the Itable interface from the Layer (not from the FeatureClass)
'      Dim pInputTable As ITable
'      Set pInputTable = pInputFeatLayer
'
'      ' Get the input feature class.
'      ' The Input feature class properties, such as shape type,
'      ' will be needed for the output
'      Dim pInputFeatClass As IFeatureClass
'      Set pInputFeatClass = pInputFeatLayer.FeatureClass
'
'      ' Get the overlay layer
'      ' Use the Itable interface from the Layer (not from the FeatureClass)
'      Dim pOverlayTable As ITable
'      Set pOverlayTable = pOverlayLayer
'
'      ' Error checking
'      If pInputTable Is Nothing Then
'          MsgBox "Table QI failed"
'          Exit Function
'      End If
'
'      If pOverlayTable Is Nothing Then
'          MsgBox "Table QI failed"
'          Exit Function
'      End If
'
'      ' Check if the Feature Class Exists....
'      Dim pFeatClass As IFeatureClass
'      Set pFeatClass = OpenShapeFile(gWorkingfolder, strResultName)
'      'Delete the FeatureClass.....
'      If Not pFeatClass Is Nothing Then
'        Dim pDataset As IDataset
'        Set pDataset = pFeatClass
'        If pDataset.CanDelete Then pDataset.Delete
'      End If
'
'      ' Define the output feature class name and shape type (taken from the
'      ' properties of the input feature class)
'      Dim pFeatClassName As IFeatureClassName
'      Set pFeatClassName = New FeatureClassName
'      With pFeatClassName
'          .FeatureType = esriFTSimple
'          .ShapeFieldName = "Shape"
'          .ShapeType = pInputFeatClass.ShapeType
'      End With
'
'      ' Set output location and feature class name
'      Dim pNewWSName As IWorkspaceName
'      Set pNewWSName = New WorkspaceName
'      pNewWSName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapeFileWorkspaceFactory.1"
'      pNewWSName.PathName = gWorkingfolder
'
'      Dim pDatasetName As IDatasetName
'      Set pDatasetName = pFeatClassName
'      pDatasetName.Name = strResultName
'      Set pDatasetName.WorkspaceName = pNewWSName
'
'      ' Set the tolerance.  Passing 0.0 causes the default tolerance to be used.
'      ' The default tolerance is 1/10,000 of the extent of the data frame's spatial domain
'      Dim tol As Double
'      tol = 0#
'
'      ' Perform the intersect
'      Dim pBGP As IBasicGeoprocessor
'      Set pBGP = New BasicGeoprocessor
'      Dim pOutputFeatClass As IFeatureClass
'      Set pOutputFeatClass = pBGP.Intersect(pInputTable, False, pOverlayTable, False, tol, pFeatClassName)
'
'      ' Add the output layer to the map
'      Dim pOutputFeatLayer As IFeatureLayer
'      Set pOutputFeatLayer = New FeatureLayer
'      Set pOutputFeatLayer.FeatureClass = pOutputFeatClass
'      pOutputFeatLayer.Name = pOutputFeatClass.AliasName
'      pOutputFeatLayer.Visible = True
'      gMap.AddLayer pOutputFeatLayer ' Add the Layer to the Map.........

    Dim pOutputFeatClass As IFeatureClass
    Dim pOutputFeatLayer As IFeatureLayer
    Dim pFlag As Boolean
    Dim strInstallPath As String
    strInstallPath = GetArcGISPath
'
    ' Get a valid name.....
    strResultName = Get_Available_name(strResultName)
    
    GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
    GP.Intersect pInputFeatLayer.Name & ";" & pOverlayLayer.Name, gWorkingfolder & "\" & strResultName & ".shp", "ALL"
    
    Set pOutputFeatLayer = GetInputFeatureLayer(strResultName)
    Set pOutputFeatClass = pOutputFeatLayer.FeatureClass
    If pOutputFeatClass.FeatureCount(Nothing) > 1 Then
        ' # Dissolve the polygons to simplify.............
        'Set the toolbox
        GP.overwriteouput = 1
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve gWorkingfolder & "\" & strResultName & ".shp", gWorkingfolder & "\" & strResultName & "_Dis.shp", "#", "#", "MULTI_PART"
        ' Delete the Intersect layer.....
        If Not pOutputFeatLayer Is Nothing Then gMap.DeleteLayer pOutputFeatLayer: Delete_Dataset_ST gWorkingfolder, strResultName
        ' Return the Dissolved layer....
        Set pOutputFeatLayer = GetInputFeatureLayer(strResultName & "_Dis")
        Set pOutputFeatClass = pOutputFeatLayer.FeatureClass
    End If
    
    If pOutputFeatClass.FeatureCount(Nothing) = 0 Then gMap.DeleteLayer pOutputFeatLayer: Set pOutputFeatLayer = Nothing

Continue:
      
      Set Get_Intersect_FeatureLayer = pOutputFeatLayer ' Return the output....
     
Cleanup:

    'Set pBGP = Nothing
    Set pOutputFeatClass = Nothing

  Exit Function
ErrorHandler:

  m_PassFlag = False
  HandleError True, "Get_Intersect_FeatureLayer" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    

  End Function
  


Private Function CreateShapefile(strShapefileName As String, pInShapeFile As IFeatureLayer) As IFeatureLayer

      On Error GoTo ErrorHandler
      Const strShapeFieldName As String = "Shape"
      
      Dim fso As New FileSystemObject
      
      ' Open the folder to contain the shapefile as a workspace
      Dim pWs As IFeatureWorkspace
      Dim pWorkspaceFactory As IWorkspaceFactory
      Set pWorkspaceFactory = New ShapefileWorkspaceFactory
      Set pWs = pWorkspaceFactory.OpenFromFile(gWorkingfolder, 0)
      
      ' Open the Input featureClass.....
      Dim pInFClass As IFeatureClass
      Set pInFClass = pInShapeFile.FeatureClass

        Dim pInDs As IDataset
        Set pInDs = pInFClass
        Dim pOutDs As IDataset
        Dim pOutFC As IFeatureClass
        
        If pInDs.CanCopy Then
            Set pOutFC = OpenShapeFile(gWorkingfolder, strShapefileName & "_Buf")
            If Not pOutFC Is Nothing Then
                Set pOutDs = pOutFC
                If pOutDs.CanDelete Then pOutDs.Delete
            End If
            Set pOutDs = pInDs.Copy(strShapefileName & "_Buf", pWs)
            Set pOutFC = pWs.OpenFeatureClass(strShapefileName & "_Buf")
            Dim pFeatLyr As IFeatureLayer
            Set pFeatLyr = New FeatureLayer
            Set pFeatLyr.FeatureClass = pOutFC
            pFeatLyr.Name = pOutFC.AliasName
            gMap.AddLayer pFeatLyr
        End If
        
        Set CreateShapefile = pFeatLyr
                                
Cleanup:

  Exit Function
ErrorHandler:
  m_PassFlag = False
  HandleError True, "CreateShapefile" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function

'Feb 20 2009. Ying: parse min max for buffer area
Private Sub PrepareBufferFeature(strExp As String, strWorkspace As String, InFlayerName As String)
    
    Dim opIndex As Integer
    Dim bufferMin As Integer
    Dim bufferMax As Integer
    Dim dScalefac As Double
    ' Get the Map units....
    dScalefac = Get_Scale_Factor(gMap.MapUnits)

    'parse strExp
    'check >min
    opIndex = InStr(1, Trim(strExp), ">", vbTextCompare)
    If InStr(1, strExp, ">", vbTextCompare) > 0 Then
        bufferMin = CInt(Right(Trim(strExp), Len(Trim(strExp)) - opIndex)) * dScalefac
        bufferMax = -1
    End If
      
    'check <max
    opIndex = InStr(1, Trim(strExp), "<", vbTextCompare)
    If InStr(1, strExp, "<", vbTextCompare) > 0 Then
        bufferMax = CInt(Right(Trim(strExp), Len(Trim(strExp)) - opIndex)) * dScalefac
        bufferMin = -1
    End If
      
    'check min-max
    opIndex = InStr(1, Trim(strExp), "-", vbTextCompare)
    If InStr(1, strExp, "-", vbTextCompare) > 0 Then
        bufferMin = CInt(Left(Trim(strExp), opIndex - 1)) * dScalefac
        bufferMax = CInt(Right(Trim(strExp), Len(Trim(strExp)) - opIndex)) * dScalefac
    End If
    
    CreateSingleBufferFeature bufferMin, bufferMax, strWorkspace, InFlayerName
End Sub

'Feb 20 2009. Ying: create a single buffer feature
'accepts both polygon and poyline as input
'has bufferMin and bufferMax, min/max value is -1 if min/max does not exist,
'if bufferMax is -1, activeView's full extent is used as max_buffer
Private Sub CreateSingleBufferFeature(bufferMin As Integer, bufferMax As Integer, strWorkspacePath As String, InFlayerName As String)
    
    Dim strWorkspace As String
    'if workspace path has no \ at end, add it
    If Right(strWorkspacePath, 1) <> "\" Then
        strWorkspace = strWorkspacePath & "\"
    End If
    Dim pInlayer As IFeatureLayer
    'possible to be polyline layer or polygon layer
    Set pInlayer = GetInputFeatureLayer(InFlayerName)
        
    Dim pEnv As IEnvelope
    Dim pExtentFeature As IFeature
    Dim pFeatCursor As IFeatureCursor
    
    Dim pPolygon As IPolygon
    Dim pPointLL As IPoint
    Dim pPointTL As IPoint
    Dim pPointTR As IPoint
    Dim pPointLR As IPoint
    Dim pPointCollection As IPointCollection

    Dim GP As Object
    If GP Is Nothing Then Set GP = CreateObject("esriGeoprocessing.GpDispatch.1") ' Create a GP object...........
    
    ' Get the ARC installation path....
    Dim strInstallPath As String
    strInstallPath = GetArcGISPath

'    Dim pGPComHelper As IGPComHelper
'    Dim pGPSettings As IGeoProcessorSettings
'
'    Set pGPComHelper = CreateObject("esriGeoprocessing.GPDispatch.1")
'    Set pGPSettings = pGPComHelper.EnvironmentManager
'
'    pGPSettings.AddOutputsToMap = True     'set False if results are expected to be deleted from map after each geoprocessing


    'delete intermediate results from previous if any
    Dim pLayer As IFeatureLayer
    Set pLayer = GetInputFeatureLayer(InFlayerName & "_BufMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(InFlayerName & "_BufMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(InFlayerName & "_DisMax")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(InFlayerName & "_DisMin")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = GetInputFeatureLayer(InFlayerName & "_Buf")
    If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    Set pLayer = Nothing
    
    Delete_Dataset_ST strWorkspace, InFlayerName & "_BufMax"
    Delete_Dataset_ST strWorkspace, InFlayerName & "_BufMin"
    Delete_Dataset_ST strWorkspace, InFlayerName & "_DisMax"
    Delete_Dataset_ST strWorkspace, InFlayerName & "_DisMin"
    Delete_Dataset_ST strWorkspace, InFlayerName & "_Buf"
    
    'Set the toolbox
    If bufferMin <= 0 Then      'no min value or min=0, format: <max
        If pInlayer.FeatureClass.ShapeType = esriGeometryPolygon Then     'polygon
            
            'buffer max
            GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
            GP.Buffer strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_BufMax.shp", bufferMax

            'dissolve max
            GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
            GP.Dissolve strWorkspace & InFlayerName & "_BufMax.shp", strWorkspace & InFlayerName & "_DisMax.shp", "#", "#", "MULTI_PART"

            'dissolve self as min
            GP.Dissolve strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_DisMin.shp", "#", "#", "MULTI_PART"

            'union & remove
            GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
            GP.Union strWorkspace & InFlayerName & ".shp" & ";" & strWorkspace & InFlayerName & "_DisMax.shp", strWorkspace & InFlayerName & "_Buf.shp", "ONLY_FID", "#", "GAPS"

            RemoveOverlap InFlayerName & "_Buf"

''            Dim minLayer As IFeatureLayer
''            Set minLayer = GetFeatureLayer(gWorkingfolder, InFlayerName & "_DisMin.shp")
''
''            Dim maxLayer As IFeatureLayer
''            Set maxLayer = GetFeatureLayer(gWorkingfolder, InFlayerName & "_DisMax.shp")
''            Set maxLayer = Get_EraseLayer(minLayer, maxLayer)

        Else    'polyline
            'buffer max
'            GP.RefreshCatalog (strWorkspace)
            GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
            GP.Buffer strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_BufMax.shp", bufferMax
            
            'dissolve max
            GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
            GP.Dissolve strWorkspace & InFlayerName & "_BufMax.shp", strWorkspace & InFlayerName & "_Buf.shp", "#", "#", "MULTI_PART"

        End If
        
    ElseIf bufferMax <= 0 Then  'no max value, then max is assumed to be full extent, format: >min
        'same for polygon and polyline
        'buffer min
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
        GP.Buffer strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_BufMin.shp", bufferMin
        
        'dissolve min
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve strWorkspace & InFlayerName & "_BufMin.shp", strWorkspace & InFlayerName & "_DisMin.shp", "#", "#", "MULTI_PART"
    
        'create max as full extent
        Set pEnv = gMxDoc.ActiveView.FullExtent
                
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        'Process:  Create a feature class to update
        GP.CreateFeatureClass strWorkspace, InFlayerName & "_BufMax.shp", "POLYGON", strWorkspace & InFlayerName & ".shp"

        'create a polygon which covers the full extent
        Set pPointLL = New Point
        pPointLL.x = pEnv.XMin
        pPointLL.y = pEnv.YMin
    
        Set pPointTL = New Point
        pPointTL.x = pEnv.XMin
        pPointTL.y = pEnv.YMax
    
        Set pPointTR = New Point
        pPointTR.x = pEnv.XMax
        pPointTR.y = pEnv.YMax
    
        Set pPointLR = New Point
        pPointLR.x = pEnv.XMax
        pPointLR.y = pEnv.YMin
    
        Set pPolygon = New Polygon
        Set pPointCollection = pPolygon
        pPointCollection.AddPoint pPointLL
        pPointCollection.AddPoint pPointTL
        pPointCollection.AddPoint pPointTR
        pPointCollection.AddPoint pPointLR
        pPolygon.Close
    
        'create the feature in buf_max, assign the polygon
        Set pInlayer = GetInputFeatureLayer(InFlayerName & "_BufMax")
        Set pFeatCursor = pInlayer.FeatureClass.Update(Nothing, False)
        Set pExtentFeature = pInlayer.FeatureClass.CreateFeature
        Set pExtentFeature.Shape = pPolygon
        pExtentFeature.Store
        
        'union and remove
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
        GP.Union strWorkspace & InFlayerName & "_BufMin.shp" & ";" & strWorkspace & InFlayerName & "_BufMax.shp", strWorkspace & InFlayerName & "_Buf.shp", "ONLY_FID", "#", "GAPS"
        RemoveOverlap InFlayerName & "_Buf"
    
    Else        'format min-max
        
        'buffer max
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
       
        GP.Buffer strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_BufMax.shp", bufferMax
       
        'buffer min
        'GP.Toolbox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
        GP.Buffer strWorkspace & InFlayerName & ".shp", strWorkspace & InFlayerName & "_BufMin.shp", bufferMin
        
        'dissolve max
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve strWorkspace & InFlayerName & "_BufMax.shp", strWorkspace & InFlayerName & "_DisMax.shp", "#", "#", "MULTI_PART"
        
        'dissolve min
        'GP.Toolbox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve strWorkspace & InFlayerName & "_BufMin.shp", strWorkspace & InFlayerName & "_DisMin.shp", "#", "#", "MULTI_PART"
        
        'union and remove
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
        GP.Union strWorkspace & InFlayerName & "_DisMin.shp" & ";" & strWorkspace & InFlayerName & "_DisMax.shp", strWorkspace & InFlayerName & "_Buf.shp", "ONLY_FID", "#", "GAPS"
       
        RemoveOverlap InFlayerName & "_Buf"
        
    End If

End Sub

'Feb 20 2009, Ying: Remove overlap part from unioned layer
Private Sub RemoveOverlap(fileName As String)
On Error GoTo ShowError
    Dim swslayer As IFeatureLayer
    Dim pFeatCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pTopoOp As ITopologicalOperator
    Dim pPolygon As IPolygon
    Dim pPolyline As IPolyline
    Dim pFieldIndex As Integer
    
    Dim toRemoveFile As String
    
    Set swslayer = GetInputFeatureLayer(fileName)

    toRemoveFile = "FID_" & swslayer.Name
    If Len(toRemoveFile) > 10 Then
        toRemoveFile = Left(toRemoveFile, 10)
    End If

'    pFieldIndex = swslayer.FeatureClass.Fields.FindField(toRemoveFile)
'
'    Set pFeatCursor = swslayer.FeatureClass.Search(Nothing, False)
'    Set pFeature = pFeatCursor.NextFeature
'
'    Do Until pFeature Is Nothing       'sws feature
'        If pFeature.Value(pFieldIndex) <> -1 Then
'            pFeature.Delete
'        End If
'
'        Set pFeature = pFeatCursor.NextFeature
'    Loop
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pTable As ITable
    If Not swslayer Is Nothing Then
        Set pTable = swslayer.FeatureClass
        pQueryFilter.WhereClause = " " & toRemoveFile & " <> -1 "
        pTable.DeleteSearchedRows pQueryFilter
    End If
    Exit Sub

ShowError:
    MsgBox "Error in RemoveOverlap: " & Err.Description
End Sub

Private Sub CreateBufferFeatureClass(pFeatLyr As IFeatureLayer, iBufferDistance As Integer, pDiff As Boolean)

      On Error GoTo ErrorHandler
      Dim pActiveView As IActiveView
      Dim pEnumFeature As IEnumFeature
      Dim pFeature As IFeature
      Dim pPolygon As IPolygon
      Dim pTopoOp As ITopologicalOperator
      Dim pFeatCursor As IFeatureCursor
      
      'Buffer all the selected features by the BufferDistance
      'and create a new polygon element from each result
      Set pFeatCursor = pFeatLyr.Search(Nothing, False)
      Set pFeature = pFeatCursor.NextFeature
      
      Do While Not pFeature Is Nothing
        Set pTopoOp = pFeature.Shape
        Set pTopoOp = pTopoOp.Buffer(iBufferDistance)
        If pDiff Then
            Set pFeature.Shape = pTopoOp.Difference(pFeature.ShapeCopy)
        Else
            Set pFeature.Shape = pTopoOp
        End If
        'Store the new feature......
        pFeature.Store
        Set pFeature = pFeatCursor.NextFeature
      Loop

    Exit Sub
    
Cleanup:

  Exit Sub
ErrorHandler:
  m_PassFlag = False
  HandleError True, "CreateBufferFeatureClass" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub

Private Sub Create_Line_Buffer(strWorkspace As String, InFlayerName As String, iBufferDistance As Integer)
    
        On Error GoTo ErrorHandler
  
        ' Get the ARC installation path....
        Dim strInstallPath As String
        strInstallPath = GetArcGISPath
        
        ' Delete the Dataset of Exists...
        Dim pDataset As IDataset
        Dim pFeatClass As IFeatureClass
        Set pFeatClass = OpenShapeFile(strWorkspace, InFlayerName & "_Buf")
        If Not pFeatClass Is Nothing Then
            Set pDataset = pFeatClass
            If pDataset.CanDelete Then pDataset.Delete
        End If
        
        'Set the toolbox
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Analysis Tools.tbx"
        GP.Buffer strWorkspace & "\" & InFlayerName & ".shp", strWorkspace & "\" & InFlayerName & "_Buf.shp", iBufferDistance
        
        ' ************************************
        ' Dissolve the Features to make one feature....
        ' ************************************

        'Set the toolbox
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve strWorkspace & "\" & InFlayerName & "_Buf.shp", strWorkspace & "\" & InFlayerName & "_Dis.shp", "#", "#", "MULTI_PART"
        
        ' Delete the Buf layer before the dissolve......
        Dim pResult As IFeatureLayer
        Set pResult = GetInputFeatureLayer(InFlayerName & "_Buf")
        If Not pResult Is Nothing Then gMap.DeleteLayer pResult: Delete_Dataset_ST strWorkspace, InFlayerName & "_Buf"
        
        'Clean up........
        Set pResult = Nothing
        Set pDataset = Nothing
        Set pFeatClass = Nothing
        
'        ' ************************************
'        ' Dissolve the Features to make one feature....
'        ' ************************************
'         'Use the Itable interface from the Layer (not from the FeatureClass)
'         Dim pInputFeatClass As IFeatureClass
'         Set pInputFeatClass = OpenShapeFile(strWorkspace, InFlayerName & "_Buf")
'
'          Dim pInputTable As ITable
'          Set pInputTable = pInputFeatClass
'
'
'          ' Define the output feature class name and shape type (taken from the
'          ' properties of the input feature class)
'          Dim pFeatClassName As IFeatureClassName
'          Set pFeatClassName = New FeatureClassName
'          With pFeatClassName
'              .FeatureType = esriFTSimple
'              .ShapeFieldName = "Shape"
'              .ShapeType = pInputFeatClass.ShapeType
'          End With
'
'          ' Set output location and feature class name
'          Dim pNewWSName As IWorkspaceName
'          Set pNewWSName = New WorkspaceName
'          pNewWSName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapeFileWorkspaceFactory.1"
'          pNewWSName.PathName = strWorkspace
'
'          Dim pDatasetName As IDatasetName
'          Set pDatasetName = pFeatClassName
'          pDatasetName.Name = InFlayerName & "_Dis"
'          Set pDatasetName.WorkspaceName = pNewWSName
'
'          ' Check if the Feature Class Exists....
'          Set pFeatClass = OpenShapeFile(strWorkspace, InFlayerName & "_Dis")
'          'Delete the FeatureClass.....
'          If Not pFeatClass Is Nothing Then
'            Set pDataset = pFeatClass
'            If pDataset.CanDelete Then pDataset.Delete
'          End If
'
'          ' Set the tolerance.  Passing 0.0 causes the default tolerance to be used.
'          ' The default tolerance is 1/10,000 of the extent of the data frame's spatial domain
'          Dim tol As Double
'          tol = 0#
'
''          ' Perform the intersect
'          Dim pBGP As IBasicGeoprocessor
'          Set pBGP = New BasicGeoprocessor
'          Dim pOutputTable As ITable
'          Dim pOutputFeatClass As IFeatureClass
'         Set pOutputTable = pBGP.Dissolve(pInputTable, False, "ID", "Dissolve.Shape, Minimum.ID", pDatasetName)
'          Set pOutputFeatClass = pOutputTable
     
        
Exit Sub
ErrorHandler:
  m_PassFlag = False
  HandleError True, "Create_Line_Buffer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Function Dissolve(pFeatCursor As IFeatureCursor) As IFeature
   
    Dim pTopoOp As ITopologicalOperator
    Dim pFeature As IFeature
    Dim pinitfeature As IFeature
    Dim pNewGeom As IGeometry
    
    Set pinitfeature = pFeatCursor.NextFeature
    Set pTopoOp = pinitfeature.ShapeCopy
    Set pFeature = pFeatCursor.NextFeature
    While Not pFeature Is Nothing
        Set pNewGeom = pTopoOp.Union(pFeature.Shape)
        Set pTopoOp = pNewGeom
        pFeature.Delete
        Set pFeature = pFeatCursor.NextFeature
    Wend
    Set pinitfeature.Shape = pTopoOp
    pinitfeature.Store

    Set Dissolve = pTopoOp
End Function

' ****************************************************************
' To Duplicate the Erase Command.....
' ****************************************************************

Public Function Get_EraseLayer(ByVal pPoly1 As IFeatureLayer, ByVal pPoly2 As IFeatureLayer) As IFeatureLayer
    
    On Error GoTo ErrorHandler

      Dim pFeature As IFeature
      Dim pSelFeature As IFeature
      Dim pPolygon As IPolygon
      Dim pTopoOp As ITopologicalOperator
      Dim pFeatCursor As IFeatureCursor
      Dim pSelFeatCursor As IFeatureCursor

      'Buffer all the selected features by the BufferDistance
      'and create a new polygon element from each result
      Set pFeatCursor = pPoly1.Search(Nothing, False)
      Set pFeature = pFeatCursor.NextFeature
      
      Do While Not pFeature Is Nothing
        Set pTopoOp = pFeature.Shape
        
        Set pSelFeatCursor = SelectByLocationIN(gMap, pPoly2, pFeature.Shape, pPoly1.FeatureClass.ShapeFieldName)
        Set pSelFeature = pSelFeatCursor.NextFeature
        If Not pSelFeature Is Nothing Then
            Do While Not pSelFeature Is Nothing
                Set pTopoOp = pTopoOp.Difference(pSelFeature.ShapeCopy)
                Set pSelFeature = pSelFeatCursor.NextFeature
            Loop
            Set pFeature.Shape = pTopoOp
            'Store the new feature......
            pFeature.Store
        End If
        
        Set pFeature = pFeatCursor.NextFeature
      Loop
      
      Set Get_EraseLayer = pPoly1
    
Cleanup:

  Exit Function
ErrorHandler:
  m_PassFlag = False
  HandleError True, "Get_EraseLayer" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Function

' ****************************************************************
' To Duplicate the Intersect Command.....
' ****************************************************************

Public Function Get_IntersectLayer(ByVal pPoly1 As IFeatureLayer, ByVal pPoly2 As IFeatureLayer, ByVal strResultName As String) As IFeatureLayer
    
    On Error GoTo ErrorHandler

      Dim pFeature As IFeature
      Dim pSelFeature As IFeature
      Dim pPolygon As IPolygon
      Dim pTopoOp As ITopologicalOperator
      Dim pFeatCursor As IFeatureCursor
      Dim pSelFeatCursor As IFeatureCursor

      'Buffer all the selected features by the BufferDistance
      'and create a new polygon element from each result
      Set pFeatCursor = pPoly1.Search(Nothing, False)
      Set pFeature = pFeatCursor.NextFeature
      
      Do While Not pFeature Is Nothing
        Set pTopoOp = pFeature.Shape
        
        Set pSelFeatCursor = SelectByLocationIN(gMap, pPoly2, pFeature.Shape, pPoly1.FeatureClass.ShapeFieldName)
        Set pSelFeature = pSelFeatCursor.NextFeature
        If Not pSelFeature Is Nothing Then
            Do While Not pSelFeature Is Nothing
                Set pTopoOp = pTopoOp.Intersect(pSelFeature.ShapeCopy, esriGeometry2Dimension)
                Set pSelFeature = pSelFeatCursor.NextFeature
            Loop
            Set pFeature.Shape = pTopoOp
            'Store the new feature......
            pFeature.Store
        Else
            'Store the new feature......
            pFeature.Delete
        End If
        
        Set pFeature = pFeatCursor.NextFeature
      Loop
      
      pPoly1.Name = strResultName
      Set Get_IntersectLayer = pPoly1
    
Cleanup:

  Exit Function
ErrorHandler:
  m_PassFlag = False
  HandleError True, "Get_IntersectLayer" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Function

Private Function SelectByLocationIN(ByVal pMap As IMap, ByVal pFLayer_poly As IFeatureLayer, ByVal pGeometry As IGeometry, ByVal pShapeFieldName As String) As IFeatureCursor
    
    On Error GoTo ErrorHandler:
    
    Dim pAView As IActiveView
    Dim pFSelection_poly As IFeatureSelection
    Dim pSelectionSet As ISelectionSet
    Dim pFCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pSpatialFilter As ISpatialFilter
    

    Set pAView = pMap
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pGeometry
    pSpatialFilter.GeometryField = pShapeFieldName
    pSpatialFilter.SpatialRel = esriSpatialRelOverlaps Or esriSpatialRelContains
    Set pFSelection_poly = pFLayer_poly
    pFSelection_poly.SelectFeatures pSpatialFilter, esriSelectionResultNew, False
    
    Set pSelectionSet = pFSelection_poly.SelectionSet
    pSelectionSet.Search Nothing, False, pFCursor
    Set SelectByLocationIN = pFCursor

    Exit Function

ErrorHandler:
  MsgBox Err.Number & Err.Description & "In SelectbyLocationCompIN"
    
    End Function




Private Function Get_ExportedShapefile(pFLayer As IFeatureLayer, strExportName As String, pDissolve As Boolean) As IFeatureClass

    On Error GoTo ErrorHandler

    Dim pFc As IFeatureClass
    Set pFc = pFLayer.FeatureClass
    
    'Get the FcName from the featureclass
    Dim pINFeatureClassName As IFeatureClassName
    Dim pDataset As IDataset
    Set pDataset = pFc
    
    Set pINFeatureClassName = pDataset.FullName
    Dim pInDSName As IDatasetName
    Set pInDSName = pINFeatureClassName
    
    'Get the selection set
    Dim pFSel As IFeatureSelection
    Set pFSel = pFLayer
    
    Dim pSelset As ISelectionSet
    Set pSelset = pFSel.SelectionSet
    
    'Create a new feature class name
    ' Define the output feature class name
    '
    Dim pFeatureClassName As IFeatureClassName
    Set pFeatureClassName = New FeatureClassName
    
    Dim pOutDatasetName As IDatasetName
    Set pOutDatasetName = pFeatureClassName
    
    Dim pWorkspaceName As IWorkspaceName
    Set pWorkspaceName = New WorkspaceName
    
    pWorkspaceName.PathName = gWorkingfolder
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"
    Set pOutDatasetName.WorkspaceName = pWorkspaceName
    
    
    ' Check if the Feature Class Exists....
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = OpenShapeFile(gWorkingfolder, strExportName)
    'Delete the FeatureClass.....
    If Not pFeatClass Is Nothing Then
        strExportName = Get_Available_name(strExportName)
    End If
    
    pFeatureClassName.FeatureType = esriFTSimple
    pFeatureClassName.ShapeType = esriGeometryAny
    pFeatureClassName.ShapeFieldName = "Shape"
    pOutDatasetName.Name = strExportName
    
    'Export
    Dim pExportOp As IExportOperation
    Set pExportOp = New ExportOperation
    Dim pOutFclass As IFeatureClass
    pExportOp.ExportFeatureClass pInDSName, Nothing, pSelset, Nothing, pOutDatasetName, 0
    
    If pDissolve Then
        ' ************************************
        ' Dissolve the Features to make one feature....
        ' ************************************

        ' Get the ARC installation path....
        Dim strInstallPath As String
        strInstallPath = GetArcGISPath
    
        'Set the toolbox
        GP.ToolBox = strInstallPath & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        GP.Dissolve gWorkingfolder & "\" & strExportName & ".shp", gWorkingfolder & "\" & strExportName & "_Dis.shp", "#", "#", "MULTI_PART"
    
        ' Delete the undissolved Shapefile......
        Set pDataset = OpenShapeFile(gWorkingfolder, strExportName)
        If pDataset.CanDelete Then pDataset.Delete
    
        ' Rename the dissolved dataset........
        Set pDataset = OpenShapeFile(gWorkingfolder, strExportName & "_Dis")
        If pDataset.CanRename Then pDataset.Rename strExportName

        ' Delete the renamed Layer form the Map.......
        Dim pLayer As IFeatureLayer
        Set pLayer = GetInputFeatureLayer(strExportName & "_Dis")
        If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
    End If
    
    'Return the exported shape file.....
    Set Get_ExportedShapefile = OpenShapeFile(gWorkingfolder, strExportName)
    
Cleanup:
    
    Set pDataset = Nothing
    Set pExportOp = Nothing
    Set pFc = Nothing
    Set pFeatClass = Nothing
    Set pFeatureClassName = Nothing
    Set pFLayer = Nothing
    Set pFSel = Nothing
    Set pInDSName = Nothing
    Set pINFeatureClassName = Nothing
    Set pLayer = Nothing
    Set pOutDatasetName = Nothing
    Set pOutFclass = Nothing
    Set pSelset = Nothing
    Set pWorkspaceName = Nothing
    
     
    Exit Function
ErrorHandler:
  m_PassFlag = False
  HandleError True, "Get_ExportedShapefile" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
       
End Function

Private Function Get_Available_name(ByVal strName As String) As String
    
    On Error GoTo ErrorHandler
    
    Dim strTmp As String
    Dim pFeatClass As IFeatureClass
    Dim iCnt As Integer
    iCnt = 0
    strTmp = strName
    Set pFeatClass = OpenShapeFile(gWorkingfolder, strTmp)
    Do While Not pFeatClass Is Nothing
        iCnt = iCnt + 1
        strTmp = strName & "_" & iCnt
        Set pFeatClass = OpenShapeFile(gWorkingfolder, strTmp)
    Loop
    
    Get_Available_name = strTmp
    
    Exit Function
ErrorHandler:
  HandleError True, "Get_Available_name" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function

' this macro will handle feature symbology according to the following rules:
'     if we have point features, then use a marker symbol for rendering
'     else, if we have line features, then use a line symbol
'     else, if we have polygon features, then use a fill symbol
'     else, (we don't have any of these feature types) so do not assign renderer to layer

Private Sub ChangeStyle(pFeatLyr As IFeatureLayer)

    On Error GoTo ErrorHandler
          
      Dim pColor As IRgbColor
      Set pColor = New RgbColor
      ' use red. it's a good color
      pColor.RGB = vbBlack
      
      Dim pOutlineSymbol As ISimpleLineSymbol
      Set pOutlineSymbol = New SimpleLineSymbol
      pOutlineSymbol.Color = pColor   ' Put the color R G B u Want
      pOutlineSymbol.Width = 0
      
      ' use Green.....
      pColor.RGB = vbGreen
      
      Dim pSym As ISymbol
      ' based on feature type, make proper symbol, then assign to pSym
      Select Case pFeatLyr.FeatureClass.ShapeType
        Case esriGeometryPoint     ' set up a marker symbol
            Dim pMarkerSym As ISimpleMarkerSymbol
            Set pMarkerSym = New SimpleMarkerSymbol
            With pMarkerSym
                .Size = 12
                .Color = pColor
                .Style = esriSMSX
            End With
            Set pSym = pMarkerSym
      
        Case esriGeometryPolyline    ' set up a line symbol
            Dim pLineSymbol As ISimpleLineSymbol
            Set pLineSymbol = New SimpleLineSymbol
            With pLineSymbol
                .Width = 1
                .Color = pColor
                .Style = esriSLSDashDotDot
            End With
            Set pSym = pLineSymbol
     
        Case esriGeometryPolygon    ' setup a fill symbol
            Dim pFillSymbol As ISimpleFillSymbol
            Set pFillSymbol = New SimpleFillSymbol
            With pFillSymbol
                .Color = pColor
                .Style = esriSFSSolid
                .Outline = pOutlineSymbol
            End With
            Set pSym = pFillSymbol
      
        Case Else
            Exit Sub
      End Select
        
      Dim pRend As IFeatureRenderer
      Set pRend = New SimpleRenderer
      
      ' set symbol.
      Dim pSimpleRend As ISimpleRenderer
      Set pSimpleRend = pRend
      Set pSimpleRend.Symbol = pSym
      
      Dim pGeoFL As IGeoFeatureLayer
      Set pGeoFL = pFeatLyr
      
      ' finally, set the new renderer to the layer and refresh the map
      Set pGeoFL.Renderer = pRend
      'Upate the TOC.....
      Dim pMapFrame As IMapFrame
      Set pMapFrame = New MapFrame
      Set pMapFrame.Map = gMap
      gMxDoc.CurrentContentsView.Refresh 0
      ' Zoom to Layer Extents....
      Dim pActiveView As IActiveView
      Set pActiveView = gMap
      pActiveView.Extent = pFeatLyr.AreaOfInterest
      pActiveView.Refresh

  Exit Sub
ErrorHandler:
  HandleError True, "ChangeStyle" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Sub






Private Sub tbsBMP_Click(PreviousTab As Integer)
    
    On Error GoTo ErrorHandler
    Dim oBMP As BMPobj
    Dim pKeys
    Dim pKey As String
    Dim iKey As Integer
        
    m_Prev_BMP = "" ' clear any selected BMP.......
    
    If tbsBMP.Tab = 0 Then
        m_backFlag = False
        tbsBMP.TabEnabled(1) = False ' Disable 2 tab..............
        tbsBMP.TabEnabled(2) = False ' Disable 3 tab..............
        Call SetDataDirectory_ST
        
    ElseIf tbsBMP.Tab = 1 Then
        
        If Not m_backFlag Then
            ' Initialize the BMP Dictionary.....
            Set gBMPtypeDict = New Scripting.Dictionary
            gBMPtypeDict.RemoveAll
            ' Initialize all BMPs.....
            Call Initialize_BMPs
            ' Now Add the Items to the List Box......
            pKeys = gBMPtypeDict.Keys
            'Clear the Listview....
            lstBMP.ListItems.Clear
            lstBMPSel.ListItems.Clear
            For iKey = 0 To gBMPtypeDict.Count - 1
                pKey = pKeys(iKey)
                Set oBMP = gBMPtypeDict.Item(pKey)
                If oBMP.BMPType = m_BMPType Or m_BMPType = "*" Then
                    lstBMP.ListItems.Add , oBMP.BMPName, oBMP.BMPName
                End If
            Next
            ' Now Fill the selected BMP list in the Listbox....
            If Not gBMPSelDict Is Nothing Then
                Dim lstItem As ListItem
                pKeys = gBMPSelDict.Keys
                For iKey = 0 To gBMPSelDict.Count - 1
                    pKey = pKeys(iKey)
                    lstBMPSel.ListItems.Add , pKey, pKey
                    Set lstItem = lstBMP.FindItem(pKey)
                    lstBMP.ListItems.Remove lstItem.Index
                Next
            End If
            
        End If
        tbsBMP.TabEnabled(2) = False ' Disable 3 tab..............
        ' Make the Column size to fit the control....
        lstBMP.ColumnHeaders.Item(1).Width = lstBMP.Width
        lstBMPSel.ColumnHeaders.Item(1).Width = lstBMPSel.Width
        lstBMP.SelectedItem = Nothing
        lstBMPSel.SelectedItem = Nothing
        
    ElseIf tbsBMP.Tab = 2 Then
            
        'Clear the Combo box....
        cmbBMPType.Clear
        cmbBMPType.ListIndex = -1
        Call Initialize_Controls
        ' Now Add the Items to the Combo Box......
        pKeys = gBMPSelDict.Keys
        For iKey = 0 To gBMPSelDict.Count - 1
            pKey = pKeys(iKey)
            Set oBMP = gBMPSelDict.Item(pKey)
            If oBMP.BMPType = m_BMPType Or m_BMPType = "*" Then cmbBMPType.AddItem oBMP.BMPName
        Next
         
    End If
    
    Exit Sub
ErrorHandler:
  HandleError True, "tbsBMP_Click" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
    
End Sub




Private Sub cmdBrowseDEM_Click()

    On Error GoTo ErrorHandler
    If Browse_Dataset(dtRaster, txtDEMpath(1)) Then m_DEMFlag = True

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
      ElseIf (pdatasetType = dtTable) Then
        Set pObjectFilter = New GxFilterdBASEFiles
      End If
    
      pDlg.AllowMultiSelect = False
      pDlg.Title = "Select Data"
      Set pDlg.ObjectFilter = pObjectFilter
      
      Dim fso As New FileSystemObject
     
     Me.Hide
      If (pDlg.DoModalOpen(pActiveView.ScreenDisplay.hwnd, pGXSelect) = False) Then Me.Show vbModal: Exit Function
    
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
                            pControl.AddItem fso.GetBaseName(pGxObject.Name)
                          End If
                          pControl.Text = fso.GetBaseName(pGxObject.Name)
                        End If
                    ElseIf (pGXDataset.Type = esriDTRasterDataset) Then
                        Dim pRasterLayer As IRasterLayer
                        Set pRasterLayer = New RasterLayer
                        pRasterLayer.CreateFromDataset pGXDataset.Dataset
                        gMap.AddLayer pRasterLayer
                        gMxDoc.ActiveView.Refresh
                        Browse_Dataset = True
                        If pControl.Style = ComboBoxConstants.vbComboDropdownList Then
                            pControl.AddItem fso.GetBaseName(pGxObject.Name)
                        End If
                        pControl.Text = fso.GetBaseName(pGxObject.Name)
                    ElseIf (pGXDataset.Type = esriDTTable) Then
                        Dim pTable As ITable
                        Dim pStTab As IStandaloneTable
                        Dim pStTabColl As IStandaloneTableCollection
                        Set pTable = pGXDataset.Dataset
                        Set pStTab = New StandaloneTable
                        Set pStTab.Table = pTable
                        Set pStTabColl = gMap
                        pStTabColl.AddStandaloneTable pStTab
                        If pControl.Style = ComboBoxConstants.vbComboDropdownList Then
                            pControl.AddItem fso.GetBaseName(pGxObject.Name)
                        End If
                        pControl.Text = fso.GetBaseName(pGxObject.Name)
                    End If
                End If
            End If
            
            Me.Show vbModal
    
  Exit Function
ErrorHandler:
  HandleError True, "Browse_Dataset " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function

Private Sub cmdBrowselanduse_Click()

    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtLandusepath(1)) Then m_LanduseFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowselanduse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Sub cmdBrowseRoad_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtRoadpath(1)) Then m_RoadFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseRoad_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub

Private Sub cmdBrowseSoil_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtSoilpath(1)) Then m_SoilFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseSoil_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Sub

Private Sub cmdBrowseStream_Click()
    
    On Error GoTo ErrorHandler
    If Browse_Dataset(dtFeature, txtStreampath(1)) Then m_StreamFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowseStream_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

    
End Sub




Private Sub txtDEMpath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtDEMpath(1).Text Or txtLandusepath(1).Text = txtDEMpath(1).Text Or txtRoadpath(1).Text = txtDEMpath(1).Text Or txtStreampath(1).Text = txtDEMpath(1).Text Then
        txtDEMpath(1).ListIndex = -1
    End If
    
    gDEMdata = txtDEMpath(1).Text
    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtDEMpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        

End Sub

Private Sub txtImp_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtImp(1).Text = txtMRLC(1).Text Or txtImp(1).Text = txtDEMpath(1).Text Then
        txtImp(1).ListIndex = -1
    End If
    
    gImperviousdata = txtImp(1).Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtImp_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtLandusepath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtLandusepath(1).Text Or txtLandusepath(1).Text = txtDEMpath(1).Text Or txtRoadpath(1).Text = txtLandusepath(1).Text Or txtStreampath(1).Text = txtLandusepath(1).Text Then
        txtLandusepath(1).ListIndex = -1
    End If
        
    gLandusedata = txtLandusepath(1).Text
    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtLandusepath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtMRLC_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtMRLC(1).Text = txtImp(1).Text Or txtMRLC(1).Text = txtDEMpath(1).Text Then
        txtMRLC(1).ListIndex = -1
    End If
    
    gMRLCdata = txtMRLC(1).Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtMRLC_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtMrlc_lk_Click()
    
    On Error GoTo ErrorHandler
    If txtMrlc_lk.Text = txtSoil_lk.Text Then
        txtMrlc_lk.ListIndex = -1
    End If
    
    gMrlcTable = txtMrlc_lk.Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtMrlc_lk_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtRoadpath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtRoadpath(1).Text Or txtLandusepath(1).Text = txtRoadpath(1).Text Or txtRoadpath(1).Text = txtDEMpath(1).Text Or txtStreampath(1).Text = txtRoadpath(1).Text Then
        txtRoadpath(1).ListIndex = -1
    End If
    
    gRoaddata = txtRoadpath(1).Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtRoadpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtSoil_lk_Click()
    
     On Error GoTo ErrorHandler
    If txtSoil_lk.Text = txtMrlc_lk.Text Then
        txtSoil_lk.ListIndex = -1
    End If
    
    gSoilTable = txtSoil_lk.Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtMrlc_lk_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        

End Sub

Private Sub txtSoilpath_Click(Index As Integer)

    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtDEMpath(1).Text Or txtLandusepath(1).Text = txtSoilpath(1).Text Or txtRoadpath(1).Text = txtSoilpath(1).Text Or txtStreampath(1).Text = txtSoilpath(1).Text Then
        txtSoilpath(1).ListIndex = -1
    End If
        
    gSoildata = txtSoilpath(1).Text
    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtSoilpath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

Private Sub txtStreampath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtSoilpath(1).Text = txtStreampath(1).Text Or txtDEMpath(1).Text = txtStreampath(1).Text Or txtRoadpath(1).Text = txtStreampath(1).Text Or txtStreampath(1).Text = txtLandusepath(1).Text Then
        txtStreampath(1).ListIndex = -1
    End If
    
    gStreamdata = txtStreampath(1).Text
        
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtStreampath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub

Private Sub txtWTPath_Click(Index As Integer)
    
    On Error GoTo ErrorHandler
    If txtWTPath(1).Text = txtLandusepath(1).Text Or txtWTPath(1).Text = txtDEMpath(1).Text Or txtRoadpath(1).Text = txtWTPath(1).Text Or txtStreampath(1).Text = txtWTPath(1).Text Then
        txtWTPath(1).ListIndex = -1
    End If
        
    gWTdata = txtWTPath(1).Text
    
Cleanup:

  Exit Sub
ErrorHandler:
  HandleError True, "txtWTPath_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
        
    
End Sub
