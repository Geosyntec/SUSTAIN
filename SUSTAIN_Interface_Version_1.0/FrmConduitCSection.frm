VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmConduitCSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conduit Cross Section"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "FrmConduitCSection.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6150
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   6000
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   706
      TabCaption(0)   =   "Conduit Cross-section"
      TabPicture(0)   =   "FrmConduitCSection.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameGEOM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Conduit Dimension Group"
      TabPicture(1)   =   "FrmConduitCSection.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtInitialFlow"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtManningCoeff"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtLength"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Decay Factors"
      TabPicture(2)   =   "FrmConduitCSection.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DecayFactorGRID"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sediment"
      TabPicture(3)   =   "FrmConduitCSection.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame5"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame2"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame7 
         Caption         =   "Clay Transport Parameters"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   82
         Top             =   4320
         Width           =   6975
         Begin VB.TextBox txtClayDia 
            Height          =   300
            Left            =   1320
            TabIndex        =   88
            Text            =   "0"
            ToolTipText     =   "Effective diameter"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtClayVel 
            Height          =   300
            Left            =   1320
            TabIndex        =   87
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtClayDens 
            Height          =   300
            Left            =   3480
            TabIndex        =   86
            Text            =   "2.65"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtClayScour 
            Height          =   300
            Left            =   3480
            TabIndex        =   85
            Text            =   "1E10"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtClayDep 
            Height          =   300
            Left            =   5880
            TabIndex        =   84
            Text            =   "1E10"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtClayErod 
            Height          =   300
            Left            =   5880
            TabIndex        =   83
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Diameter (in)"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Velocity (in/sec)"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Density (lb/ft³)"
            Height          =   255
            Left            =   2280
            TabIndex        =   92
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "Scour Stress (lb/ft²)"
            Height          =   375
            Left            =   2280
            TabIndex        =   91
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label27 
            Caption         =   "Deposition Stress (lb/ft²)"
            Height          =   375
            Left            =   4560
            TabIndex        =   90
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Erodibility (lb/ft²/day)"
            Height          =   375
            Left            =   4560
            TabIndex        =   89
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Silt Transport Parameters"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   62
         Top             =   2880
         Width           =   6975
         Begin VB.TextBox txtSiltErod 
            Height          =   300
            Left            =   5880
            TabIndex        =   68
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtSiltDep 
            Height          =   300
            Left            =   5880
            TabIndex        =   67
            Text            =   "1E10"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSiltScour 
            Height          =   300
            Left            =   3480
            TabIndex        =   66
            Text            =   "1E10"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtSiltDens 
            Height          =   300
            Left            =   3480
            TabIndex        =   65
            Text            =   "2.65"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSiltVel 
            Height          =   300
            Left            =   1320
            TabIndex        =   64
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtSiltDia 
            Height          =   300
            Left            =   1320
            TabIndex        =   63
            Text            =   "0"
            ToolTipText     =   "Effective diameter"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Erodibility (lb/ft²/day)"
            Height          =   375
            Left            =   4560
            TabIndex        =   74
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Deposition Stress (lb/ft²)"
            Height          =   375
            Left            =   4560
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Scour Stress (lb/ft²)"
            Height          =   375
            Left            =   2280
            TabIndex        =   72
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Density (lb/ft³)"
            Height          =   255
            Left            =   2280
            TabIndex        =   71
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Velocity (in/sec)"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Diameter (in)"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Sand Transport Parameters"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   51
         Top             =   1680
         Width           =   6975
         Begin VB.TextBox txtSandVel 
            Height          =   300
            Left            =   1320
            TabIndex        =   55
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSandDens 
            Height          =   300
            Left            =   3480
            TabIndex        =   54
            Text            =   "2.65"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSandCoeff 
            Height          =   300
            Left            =   3480
            TabIndex        =   53
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSandExp 
            Height          =   300
            Left            =   5880
            TabIndex        =   52
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSandDia 
            Height          =   300
            Left            =   1320
            TabIndex        =   56
            Text            =   "0"
            ToolTipText     =   "Effective diameter"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Diameter (in)"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Velocity (in/sec)"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Density (lb/ft³)"
            Height          =   255
            Left            =   2280
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Coefficient"
            Height          =   255
            Left            =   2280
            TabIndex        =   58
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Exponent"
            Height          =   255
            Left            =   4560
            TabIndex        =   57
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "General Parameters"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   6975
         Begin VB.TextBox txtClay 
            Height          =   300
            Left            =   5880
            TabIndex        =   44
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSilt 
            Height          =   300
            Left            =   5880
            TabIndex        =   43
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSand 
            Height          =   300
            Left            =   3480
            TabIndex        =   42
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtPorosity 
            Height          =   300
            Left            =   3480
            TabIndex        =   41
            Text            =   "0.5"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtDepth 
            Height          =   300
            Left            =   1320
            TabIndex        =   40
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtWidth 
            Height          =   300
            Left            =   1320
            TabIndex        =   39
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Clay Fraction"
            Height          =   255
            Left            =   4560
            TabIndex        =   50
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Silt Fraction"
            Height          =   255
            Left            =   4560
            TabIndex        =   49
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Sand Fraction"
            Height          =   255
            Left            =   2280
            TabIndex        =   48
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Bed Porosity"
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label label2 
            Caption         =   "Bed Depth (ft)"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label label2 
            Caption         =   "Bed Width (ft)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Invert Level Elevation"
         Height          =   1815
         Left            =   -71160
         TabIndex        =   33
         Top             =   2100
         Width           =   2295
         Begin VB.TextBox txtSlopeExit 
            Height          =   300
            Left            =   1200
            TabIndex        =   35
            Text            =   "0"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtSlopeEntrance 
            Height          =   300
            Left            =   1200
            TabIndex        =   34
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "At Exit:"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "At Entrance:"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Head Loss Coefficient"
         Height          =   1815
         Left            =   -74640
         TabIndex        =   26
         Top             =   2100
         Width           =   3135
         Begin VB.TextBox txtHeadLossLength 
            Height          =   300
            Left            =   1920
            TabIndex        =   29
            Text            =   "0"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtHeadLossExit 
            Height          =   300
            Left            =   1320
            TabIndex        =   28
            Text            =   "0"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtHeadLossEntrance 
            Height          =   300
            Left            =   1320
            TabIndex        =   27
            Text            =   "0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Along Conduit Length:"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "At Exit:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "At Entrance:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtInitialFlow 
         Height          =   300
         Left            =   -72120
         TabIndex        =   25
         Text            =   "0"
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox txtManningCoeff 
         Height          =   300
         Left            =   -72120
         TabIndex        =   24
         Text            =   "0.14"
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtLength 
         Height          =   300
         Left            =   -72120
         TabIndex        =   23
         Text            =   "1"
         Top             =   660
         Width           =   855
      End
      Begin VB.Frame frameGEOM 
         Height          =   2895
         Left            =   2880
         TabIndex        =   8
         Top             =   660
         Width           =   3375
         Begin VB.CommandButton cmdEdit 
            Height          =   300
            Left            =   2880
            Picture         =   "FrmConduitCSection.frx":093A
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   2445
            Width           =   300
         End
         Begin VB.ComboBox cmbTransect 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   2445
            Width           =   1455
         End
         Begin VB.TextBox txtDimension 
            BackColor       =   &H80000000&
            Height          =   330
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Feet"
            Top             =   480
            Width           =   900
         End
         Begin VB.TextBox txtParam1 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Text            =   "0"
            Top             =   1200
            Width           =   900
         End
         Begin VB.CommandButton cmdIncrease 
            Height          =   200
            Left            =   750
            Picture         =   "FrmConduitCSection.frx":0A24
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   435
            Width           =   300
         End
         Begin VB.CommandButton cmdDecrease 
            Height          =   200
            Left            =   750
            Picture         =   "FrmConduitCSection.frx":0B0E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   630
            Width           =   300
         End
         Begin VB.TextBox txtBarrel 
            Height          =   360
            Left            =   120
            TabIndex        =   12
            Text            =   "1"
            Top             =   465
            Width           =   615
         End
         Begin VB.TextBox txtParam2 
            Height          =   375
            Left            =   2040
            TabIndex        =   11
            Text            =   "0"
            Top             =   1200
            Width           =   900
         End
         Begin VB.TextBox txtParam3 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   1920
            Width           =   900
         End
         Begin VB.TextBox txtParam4 
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Text            =   "0"
            Top             =   1920
            Width           =   900
         End
         Begin VB.Label lblTransect 
            Caption         =   "Transect Name"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label lblBarrel 
            Caption         =   "Barrels"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblDimension 
            Caption         =   "Dimensions"
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblParam1 
            Caption         =   "Max. Depth (ft.)"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblParam2 
            Caption         =   "Label"
            Height          =   255
            Left            =   2040
            TabIndex        =   19
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblParam3 
            Caption         =   "Label"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblParam4 
            Caption         =   "Label"
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Top             =   1680
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   2535
         Begin VB.ComboBox cmbCrossSection 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   2055
         End
         Begin VB.PictureBox imgCrossSection 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   480
            ScaleHeight     =   1425
            ScaleWidth      =   1425
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label LabelDummy 
            Caption         =   "Instant Connection"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Shape"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DecayFactorGRID 
         Height          =   1695
         Left            =   -74520
         TabIndex        =   75
         Top             =   840
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.Label Label7 
         Caption         =   "Initial flow in the conduit (cfs):"
         Height          =   300
         Left            =   -74640
         TabIndex        =   78
         Top             =   1620
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Manning's Roughness Coefficient:"
         Height          =   300
         Left            =   -74640
         TabIndex        =   77
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Conduit Length (ft):"
         Height          =   300
         Left            =   -74640
         TabIndex        =   76
         Top             =   660
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmConduitCSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UpdateCrossSectionPicture()

    Dim crossSectionType
    crossSectionType = cmbCrossSection.Text
    EnableOtherParams
    Select Case crossSectionType
        Case "CIRCULAR":
            imgCrossSection.Picture = LoadResPicture("XCIRCULAR", 0)
            EnableParam1 ("Max. Depth (ft.)")
        Case "FILLED_CIRCULAR":
            imgCrossSection.Picture = LoadResPicture("XFILLED_CIRCULAR", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Filled Depth (ft.)")
        Case "RECT_CLOSED":
            imgCrossSection.Picture = LoadResPicture("XRECT_CLOSED", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Bottom Width (ft.)")
        Case "RECT_OPEN":
            imgCrossSection.Picture = LoadResPicture("XRECT_OPEN", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Bottom Width (ft.)")
        Case "TRAPEZOIDAL":
            imgCrossSection.Picture = LoadResPicture("XTRAPEZOIDAL", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Bottom Width (ft.)")
            EnableParam3 ("Left Slope")
            EnableParam4 ("Right Slope")
        Case "TRIANGULAR":
            imgCrossSection.Picture = LoadResPicture("XTRIANGULAR", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Top Width (ft.)")
        Case "PARABOLIC":
            imgCrossSection.Picture = LoadResPicture("XPARABOLIC", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Top Width (ft.)")
        Case "RECT_TRIANGULAR":
            imgCrossSection.Picture = LoadResPicture("XRECT_TRIANGULAR", 0)
            EnableParam1 ("Max. Depth (ft.)")
            EnableParam2 ("Top Width (ft.)")
            EnableParam3 ("Triangle Height (ft.)")
        Case "IRREGULAR":
            imgCrossSection.Picture = LoadResPicture("XIRREGULAR", 0)
            lblTransect.Visible = True
            cmbTransect.Visible = True
            cmdEdit.Visible = True
            lblBarrel.Visible = False
            txtBarrel.Visible = False
            cmdIncrease.Visible = False
            cmdDecrease.Visible = False
            txtDimension.Visible = False
            lblDimension.Visible = False
        Case "DUMMY":
            imgCrossSection.Picture = Nothing
            imgCrossSection.Visible = False
            LabelDummy.Visible = True
            DisableOtherParams
        Case Else
           imgCrossSection.Picture = LoadResPicture("XCIRCULAR", 0)
           EnableParam1 ("Max. Depth (ft.)")
    End Select
    
End Sub


Private Sub DisableOtherParams()
    lblBarrel.Visible = False
    txtBarrel.Visible = False
    txtBarrel.Text = 0
    cmdDecrease.Visible = False
    cmdIncrease.Visible = False
    txtDimension.Visible = False
    lblDimension.Visible = False
End Sub


Private Sub EnableOtherParams()
    lblBarrel.Visible = True
    txtBarrel.Visible = True
    txtBarrel.Text = 1
    cmdDecrease.Visible = True
    cmdIncrease.Visible = True
    txtDimension.Visible = True
    lblDimension.Visible = True
    imgCrossSection.Visible = True
End Sub

Private Sub EnableParam1(txtValue As String)
    lblParam1.Visible = True
    lblParam1.Caption = txtValue
    txtParam1.Visible = True
    txtParam1.Text = 1
End Sub

Private Sub EnableParam2(txtValue As String)
    lblParam2.Visible = True
    lblParam2.Caption = txtValue
    txtParam2.Visible = True
End Sub

Private Sub EnableParam3(txtValue As String)
    lblParam3.Visible = True
    lblParam3.Caption = txtValue
    txtParam3.Visible = True
End Sub

Private Sub EnableParam4(txtValue As String)
    lblParam4.Visible = True
    lblParam4.Caption = txtValue
    txtParam4.Visible = True
End Sub


Private Sub cmbCrossSection_Click()
'Call subroutine to disable all additional parameters
 DisableAdditionalParameters
'Call subroutine to change the cross section picture
 UpdateCrossSectionPicture
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDecrease_Click()
    Dim totalBarrels
    totalBarrels = txtBarrel.Text
    If (totalBarrels > 1) Then
        txtBarrel.Text = CInt(totalBarrels) - 1
    End If
End Sub

Private Sub cmdEdit_Click()
    
    frmTransect.Show vbModal
    
    ' Load any new Transects......
    Dim strTransect As String
    strTransect = cmbTransect.Text
    Call LoadTransectNamesforform
    If strTransect <> "" And GetListBoxIndex(cmbTransect, strTransect) > -1 Then cmbTransect.Text = strTransect
    
End Sub

Private Sub cmdIncrease_Click()
    Dim totalBarrels
    totalBarrels = txtBarrel.Text
    txtBarrel.Text = CInt(totalBarrels) + 1
End Sub

Private Sub cmdOk_Click()
    'Call subroutine to update the conduit cross-section
    UpdateConduitCrossSectionInformation
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
   ' Load the TransactNames....
   LoadTransectNamesforform
   
  'Load the different cross-section types
   cmbCrossSection.AddItem "DUMMY"
   cmbCrossSection.AddItem "CIRCULAR"
   cmbCrossSection.AddItem "FILLED_CIRCULAR"
   cmbCrossSection.AddItem "RECT_CLOSED"
   cmbCrossSection.AddItem "RECT_OPEN"
   cmbCrossSection.AddItem "TRAPEZOIDAL"
   cmbCrossSection.AddItem "TRIANGULAR"
   cmbCrossSection.AddItem "PARABOLIC"
   cmbCrossSection.AddItem "RECT_TRIANGULAR"
   cmbCrossSection.AddItem "IRREGULAR"
   cmbCrossSection.ListIndex = 0
  'Load the image with first cross section - DUMMY
   imgCrossSection.Picture = Nothing
   imgCrossSection.Visible = False
   LabelDummy.Visible = True
   
   '** Initialize the conduit decay factors
   InitPollutantData
    
   'Call subroutine to read existing values from table
   ReadExistingConduitCrossSectionInformation
        
   'Initialize Data grid for conduit decay factors
   InitializeDataGrid
      
End Sub

Private Sub LoadTransectNamesforform()
    
    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadTransectNames
    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    cmbTransect.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        cmbTransect.AddItem pPollutantCollection.Item(iCount)
    Next
    
    Set pPollutantCollection = Nothing

End Sub


Private Sub txtBarrel_LostFocus()
    Dim totalBarrels
    totalBarrels = txtBarrel.Text
      
    Dim intBarrels
    If (IsNumeric(totalBarrels)) Then
        intBarrels = CInt(totalBarrels)
        If (CDbl(intBarrels) <> CDbl(totalBarrels)) Then
            txtBarrel.Text = intBarrels
            MsgBox "Number of Barrels should be a non zero integer"
            txtBarrel.SetFocus
        End If
    Else
        txtBarrel.Text = 1
        MsgBox "Number of Barrels should be a non zero integer"
        txtBarrel.SetFocus
    End If
End Sub

'Subroutine to disable additional parameters
Private Sub DisableAdditionalParameters()
    lblBarrel.Visible = True
    txtBarrel.Visible = True
    cmdIncrease.Visible = True
    cmdDecrease.Visible = True
    txtDimension.Visible = True
    lblDimension.Visible = True
    lblParam1.Visible = False
    lblParam2.Visible = False
    lblParam3.Visible = False
    lblParam4.Visible = False
    txtParam1.Visible = False
    txtParam2.Visible = False
    txtParam3.Visible = False
    txtParam4.Visible = False
    txtParam1.Text = 0
    txtParam2.Text = 0
    txtParam3.Text = 0
    txtParam4.Text = 0
    LabelDummy.Visible = False
    cmbTransect.Visible = False
    cmdEdit.Visible = False
    lblTransect.Visible = False
End Sub


'*** Subroutine to Add Conduit Cross-section and dimension group values
Private Sub UpdateConduitCrossSectionInformation()
On Error GoTo ShowError

    '*** Get Conduit Detail table
    Dim pConduitDetailTable As iTable
    Set pConduitDetailTable = GetInputDataTable("BMPDetail")
    
    If (pConduitDetailTable Is Nothing) Then
        Set pConduitDetailTable = CreatePropertiesTableDBF("BMPDetail")
        AddTableToMap pConduitDetailTable
    End If
    
    '*** Get field names from conduit detail table
    Dim pIDindex As Long
    pIDindex = pConduitDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pConduitDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pConduitDetailTable.FindField("PropValue")
  
    '*** Get all form parameters, validating input parameters
    Dim conduitShape As String
    conduitShape = FrmConduitCSection.cmbCrossSection.Text
    '*** BARREL
    Dim conduitBarrels As Integer
    Dim strconduitBarrels As String
    strconduitBarrels = FrmConduitCSection.txtBarrel.Text
    If (Not IsNumeric(strconduitBarrels)) Then
        MsgBox "Number of Barrels must be a valid number."
        Exit Sub
    End If
    If (Trim(strconduitBarrels) = "") Then
        conduitBarrels = 0
    Else
        conduitBarrels = CInt(strconduitBarrels)
    End If
    If (conduitBarrels < 0) Then
        MsgBox "Number of Barrels must be a positive number."
        Exit Sub
    End If
    
    '*** GEOMETRY1 - MAXIMUM DEPTH
    Dim conduitGeom1 As Double  'Max. depth
    Dim strconduitGeom1 As String
    strconduitGeom1 = FrmConduitCSection.txtParam1.Text
    If (Not IsNumeric(strconduitGeom1)) Then
        MsgBox FrmConduitCSection.lblParam1.Caption & " must be a valid number."
        Exit Sub
    End If
    If (Trim(strconduitGeom1) = "") Then
        conduitGeom1 = 0
    Else
        conduitGeom1 = CDbl(strconduitGeom1)
    End If
    If (conduitGeom1 < 0) Then
        MsgBox FrmConduitCSection.lblParam1.Caption & " must be a positive number."
        Exit Sub
    End If
    
    '*** GEOMETRY2
    Dim conduitGeom2 As Double
    Dim strconduitGeom2 As String
    strconduitGeom2 = FrmConduitCSection.txtParam2.Text
    If (Not IsNumeric(strconduitGeom2)) Then
        MsgBox FrmConduitCSection.lblParam2.Caption & " must be a valid number."
        Exit Sub
    End If
    If (Trim(strconduitGeom2) = "") Then
        conduitGeom2 = 0
    Else
        conduitGeom2 = CDbl(strconduitGeom2)
    End If
    If (conduitGeom2 < 0) Then
        MsgBox FrmConduitCSection.lblParam2.Caption & " must be a positive number."
        Exit Sub
    End If
    
    '*** GEOMETRY3
    Dim conduitGeom3 As Double
    Dim strconduitGeom3 As String
    strconduitGeom3 = FrmConduitCSection.txtParam3.Text
    If (Not IsNumeric(strconduitGeom3)) Then
        MsgBox FrmConduitCSection.lblParam3.Caption & " must be a valid number."
        Exit Sub
    End If
    If (Trim(strconduitGeom3) = "") Then
        conduitGeom3 = 0
    Else
        conduitGeom3 = CDbl(strconduitGeom3)
    End If
    If (conduitGeom3 < 0) Then
        MsgBox FrmConduitCSection.lblParam3.Caption & " must be a positive number."
        Exit Sub
    End If

    '*** GEOMETRY4
    Dim conduitGeom4 As Double
    Dim strconduitGeom4 As String
    strconduitGeom4 = FrmConduitCSection.txtParam4.Text
    If (Not IsNumeric(strconduitGeom4)) Then
        MsgBox FrmConduitCSection.lblParam4.Caption & " must be a valid number."
        Exit Sub
    End If
    If (Trim(strconduitGeom4) = "") Then
        conduitGeom4 = 0
    Else
        conduitGeom4 = CDbl(strconduitGeom4)
    End If
    If (conduitGeom4 < 0) Then
        MsgBox FrmConduitCSection.lblParam4.Caption & " must be a positive number."
        Exit Sub
    End If
    
    '*** Focus to next tab
    FrmConduitCSection.SSTab1.Tab = 1
    
    '*** CONDUIT LENGTH
    Dim conduitLength As Double  'Conduit length
    Dim strconduitLength As String
    strconduitLength = FrmConduitCSection.txtLength.Text
    If (strconduitLength <> "" And Not IsNumeric(strconduitLength)) Then
        MsgBox "Conduit length must be a valid number."
        Exit Sub
    End If
    conduitLength = CDbl(strconduitLength)
    If (conduitLength <= 0) Then
        MsgBox "Conduit length must be a positive number (greater than 0)."
        Exit Sub
    End If

    '*** MANNING'S COEFFICIENT
    Dim conduitManning As Double  'Mannings Co-efficient
    Dim strconduitManning As String
    strconduitManning = FrmConduitCSection.txtManningCoeff.Text
    If (strconduitManning <> "" And Not IsNumeric(strconduitManning)) Then
        MsgBox "Manning's Coefficient must be a valid number."
        Exit Sub
    End If
    conduitManning = CDbl(strconduitManning)
    If (conduitManning <= 0) Then
        MsgBox "Manning's Coefficient must be a positive number (greater than zero)."
        Exit Sub
    End If
    
    '*** INITIAL FLOW
    Dim conduitInitFlow As Double  'Initial Flow
    Dim strconduitInitFlow As String
    strconduitInitFlow = FrmConduitCSection.txtInitialFlow.Text
    If (strconduitInitFlow = "" Or (Not IsNumeric(strconduitInitFlow))) Then
        MsgBox "Initial Flow must be a valid number."
        Exit Sub
    End If
    conduitInitFlow = CDbl(strconduitInitFlow)
    If (conduitInitFlow < 0) Then
        MsgBox "Initial Flow must be a positive number."
        Exit Sub
    End If
    
    '*** HEAD LOSS AT ENTRANCE
    Dim conduitHeadLossEnt As Double  'Head loss coefficient at conduit entrance
    Dim strconduitHeadLossEnt As String
    strconduitHeadLossEnt = FrmConduitCSection.txtHeadLossEntrance.Text
    If (strconduitHeadLossEnt = "" Or Not IsNumeric(strconduitHeadLossEnt)) Then
        MsgBox "Head loss coefficient at conduit entrance must be a valid number."
        Exit Sub
    End If
    conduitHeadLossEnt = CDbl(strconduitHeadLossEnt)
    If (conduitHeadLossEnt < 0) Then
        MsgBox "Head loss coefficient at conduit entrance must be a positive number."
        Exit Sub
    End If
    
    '*** HEAD LOSS AT EXIT
    Dim conduitHeadLossExit As Double  'Head loss coefficient at conduit exit
    Dim strconduitHeadLossExit As String
    strconduitHeadLossExit = FrmConduitCSection.txtHeadLossExit.Text
    If (strconduitHeadLossExit = "" Or Not IsNumeric(strconduitHeadLossExit)) Then
        MsgBox "Head loss coefficient at conduit exit must be a valid number."
        Exit Sub
    End If
    conduitHeadLossExit = CDbl(strconduitHeadLossExit)
    If (conduitHeadLossExit < 0) Then
        MsgBox "Head loss coefficient at conduit exit must be a positive number."
        Exit Sub
    End If


    '*** AVERAGE HEAD LOSS ALONG THE CONDUIT
    Dim conduitHeadLossLen As Double  'Average Head loss coefficient along the conduit
    Dim strconduitHeadLossLen As String
    strconduitHeadLossLen = FrmConduitCSection.txtHeadLossLength.Text
    If (strconduitHeadLossLen = "" Or Not IsNumeric(strconduitHeadLossLen)) Then
        MsgBox "Head loss coefficient along conduit length must be a valid number."
        Exit Sub
    End If
    conduitHeadLossLen = CDbl(strconduitHeadLossLen)
    If (conduitHeadLossLen < 0) Then
        MsgBox "Head loss coefficient along conduit length must be a positive number."
        Exit Sub
    End If
    
    '*** INVERT ELEVATION AT CONDUIT ENTRANCE
    Dim conduitSlopeEnt As Double  'Invert Elevation at Conduit Entrance
    Dim strconduitSlopeEnt As String
    strconduitSlopeEnt = FrmConduitCSection.txtSlopeEntrance.Text
    If (strconduitSlopeEnt = "" Or Not IsNumeric(strconduitSlopeEnt)) Then
        MsgBox "Invert level elevation at conduit entrance must be a valid number."
        Exit Sub
    End If
    conduitSlopeEnt = CDbl(strconduitSlopeEnt)
    If (conduitSlopeEnt < 0) Then
        MsgBox "Invert level elevation at conduit entrance must be a positive number."
        Exit Sub
    End If
    
    '*** INVERT ELEVATION AT CONDUIT EXIT
    Dim conduitSlopeExit As Double  'Invert Elevation at Conduit Exit
    Dim strconduitSlopeExit As String
    strconduitSlopeExit = FrmConduitCSection.txtSlopeExit.Text
    If (strconduitSlopeExit = "" Or Not IsNumeric(strconduitSlopeExit)) Then
        MsgBox "Invert level elevation at conduit exit must be a valid number."
        Exit Sub
    End If
    conduitSlopeExit = CDbl(strconduitSlopeExit)
    If (conduitSlopeExit < 0) Then
        MsgBox "Invert level elevation at conduit exit must be a positive number."
        Exit Sub
    End If
    
    '*** Invert elevation at entrance should be greater or equal to that of exit
    If (conduitSlopeEnt < conduitSlopeExit) Then
        MsgBox "Invert level elevation at conduit entrance must be greater than or equal to the invert level elevation at conduit exit", vbExclamation
        Exit Sub
    End If
        
    '*** Add all parameters to a dictionary to add them to a table
    Dim pConduitDetailDict As Scripting.Dictionary
    Set pConduitDetailDict = CreateObject("Scripting.Dictionary")
    pConduitDetailDict.add "BMPClass", "C"
    pConduitDetailDict.add "TYPE", conduitShape
    pConduitDetailDict.add "BARRELS", conduitBarrels
    pConduitDetailDict.add "GEOM1", conduitGeom1
    pConduitDetailDict.add "GEOM2", conduitGeom2
    pConduitDetailDict.add "GEOM3", conduitGeom3
    pConduitDetailDict.add "GEOM4", conduitGeom4
    pConduitDetailDict.add "LENGTH", conduitLength
    pConduitDetailDict.add "MANN_N", conduitManning
    pConduitDetailDict.add "INIFLOW", conduitInitFlow
    pConduitDetailDict.add "ENTLOSS", conduitHeadLossEnt
    pConduitDetailDict.add "EXTLOSS", conduitHeadLossExit
    pConduitDetailDict.add "AVGLOSS", conduitHeadLossLen
    pConduitDetailDict.add "ENTINVERTLEV", conduitSlopeEnt
    pConduitDetailDict.add "EXTINVERTLEV", conduitSlopeExit
    
    ' Add the Sediment Section values.................
    pConduitDetailDict.add "Transect", cmbTransect.Text
    pConduitDetailDict.add "Bed width", txtWidth.Text
    pConduitDetailDict.add "Bed depth", txtDepth.Text
    pConduitDetailDict.add "Porosity", txtPorosity.Text
    pConduitDetailDict.add "Sand fraction", txtSand.Text
    pConduitDetailDict.add "Silt fraction", txtSilt.Text
    pConduitDetailDict.add "Clay fraction", txtClay.Text
    pConduitDetailDict.add "Sand effective diameter", txtSandDia.Text
    pConduitDetailDict.add "Sand velocity", txtSandVel.Text
    pConduitDetailDict.add "Sand density", txtSandDens.Text
    pConduitDetailDict.add "Sand coefficient", txtSandCoeff.Text
    pConduitDetailDict.add "Sand exponent", txtSandExp.Text
    pConduitDetailDict.add "Silt effective diameter", txtSiltDia.Text
    pConduitDetailDict.add "Silt velocity", txtSiltVel.Text
    pConduitDetailDict.add "Silt density", txtSiltDens.Text
    pConduitDetailDict.add "Silt Deposition stress", txtSiltDep.Text
    pConduitDetailDict.add "Silt Scour stress", txtSiltScour.Text
    pConduitDetailDict.add "Silt Erodibility", txtSiltErod.Text
    
    pConduitDetailDict.add "Clay effective diameter", txtClayDia.Text
    pConduitDetailDict.add "Clay velocity", txtClayVel.Text
    pConduitDetailDict.add "Clay density", txtClayDens.Text
    pConduitDetailDict.add "Clay Deposition stress", txtClayDep.Text
    pConduitDetailDict.add "Clay Scour stress", txtClayScour.Text
    pConduitDetailDict.add "Clay Erodibility", txtClayErod.Text
       
    '* Read from TEMPTable and add values in ConduitDetail dictionary
'    Dim pTempTable As iTable
'    Set pTempTable = GetInputDataTable("TempTable")
'    If Not (pTempTable Is Nothing) Then
'        Dim pTempCursor As ICursor
'        Dim pTempRow As iRow
'        Dim iValue As Long
'        iValue = pTempTable.FindField("DECAY")
'        Set pTempCursor = pTempTable.Search(Nothing, True)
'        Set pTempRow = pTempCursor.NextRow
'        Dim iDecayCount As Integer
'        iDecayCount = 1
'        Do While Not (pTempRow Is Nothing)
'            pConduitDetailDict.Add "Decay" & iDecayCount, pTempRow.value(iValue)
'            iDecayCount = iDecayCount + 1
'            'move to next row
'            Set pTempRow = pTempCursor.NextRow
'        Loop
'    End If
    
    Dim oRs As ADODB.Recordset
    Set oRs = DecayFactorGRID.DataSource
    oRs.MoveFirst
    
    Dim iDecayCount As Integer
    iDecayCount = 1
    Do Until oRs.EOF
        pConduitDetailDict.add "Decay" & iDecayCount, oRs.Fields(1).value
        iDecayCount = iDecayCount + 1
        oRs.MoveNext
    Loop
    '*** Check if BMPDetail table has the record, if so delete all records for that id
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & gConduitIDValue
    pConduitDetailTable.DeleteSearchedRows pQueryFilter
    
    '*** Iterate over the entire dictionary, save record in table
    Dim pRow As iRow
    Dim pConduitKeys
    pConduitKeys = pConduitDetailDict.keys
    Dim i As Integer
    Dim pPropertyName As String
    Dim pPropertyValue As String
    For i = 0 To (pConduitDetailDict.Count - 1)
        pPropertyName = pConduitKeys(i)
        pPropertyValue = pConduitDetailDict.Item(pPropertyName)
        Set pRow = pConduitDetailTable.CreateRow
        pRow.value(pIDindex) = gConduitIDValue
        pRow.value(pPropNameIndex) = pPropertyName
        pRow.value(pPropValueIndex) = pPropertyValue
        pRow.Store
    Next

    Set pConduitKeys = Nothing
    Set pConduitDetailDict = Nothing
     
    'Unload the form
    Unload Me
        
    GoTo CleanUp
    

ShowError:
    MsgBox "UpdateConduitCrossSectionInformation :" & Err.description
    
CleanUp:
End Sub



Private Sub ReadExistingConduitCrossSectionInformation()

On Error GoTo ShowError

    '*** Get Conduit Detail table
    Dim pConduitDetailTable As iTable
    Set pConduitDetailTable = GetInputDataTable("BMPDetail")
    
    If (pConduitDetailTable Is Nothing) Then
        Exit Sub
    End If
    
    '* Dictionary to store decay factors for conduits
    Dim pDecayFactor As Scripting.Dictionary
    Set pDecayFactor = CreateObject("Scripting.Dictionary")
    
    '*** Check if BMPDetail table has the record, if so READ all records, else exit sub
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & gConduitIDValue
    If (pConduitDetailTable.RowCount(pQueryFilter) = 0) Then
        Exit Sub
    End If
    
    '*** Add all parameters to a dictionary to add them to form controls
    Dim pConduitDetailDict As Scripting.Dictionary
    Set pConduitDetailDict = CreateObject("Scripting.Dictionary")
    
    '*** Iterate over the entire dictionary, save record in table
    Dim pRow As iRow
    Dim pCursor As ICursor
    Set pCursor = pConduitDetailTable.Search(pQueryFilter, True)
    '*** Get field names from conduit detail table
    Dim iPropNameIndex As Long
    iPropNameIndex = pConduitDetailTable.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pConduitDetailTable.FindField("PropValue")
    Set pRow = pCursor.NextRow
    Dim pPropName As String
    Dim pPropValue As String
    
    Do While Not (pRow Is Nothing)
        pPropName = pRow.value(iPropNameIndex)
        pPropValue = pRow.value(iPropValueIndex)
        
        Select Case pPropName
            Case "TYPE":
                FrmConduitCSection.cmbCrossSection.Text = pPropValue
                UpdateCrossSectionPicture   'Will update the cross-section image
            Case "BARRELS":
                FrmConduitCSection.txtBarrel.Text = CInt(pPropValue)
            Case "GEOM1":
                FrmConduitCSection.txtParam1.Text = CDbl(pPropValue)
            Case "GEOM2":
                FrmConduitCSection.txtParam2.Text = CDbl(pPropValue)
            Case "GEOM3":
                FrmConduitCSection.txtParam3.Text = CDbl(pPropValue)
            Case "GEOM4":
                FrmConduitCSection.txtParam4.Text = CDbl(pPropValue)
            Case "LENGTH":
'                frmConduitCSection.txtLength.Text = CDbl(pPropValue)
            Case "MANN_N":
                FrmConduitCSection.txtManningCoeff.Text = CDbl(pPropValue)
            Case "INIFLOW":
                FrmConduitCSection.txtInitialFlow.Text = CDbl(pPropValue)
            Case "ENTLOSS":
                FrmConduitCSection.txtHeadLossEntrance.Text = CDbl(pPropValue)
            Case "EXTLOSS":
                FrmConduitCSection.txtHeadLossExit.Text = CDbl(pPropValue)
            Case "AVGLOSS":
                FrmConduitCSection.txtHeadLossLength.Text = CDbl(pPropValue)
            Case "ENTINVERTLEV":
                FrmConduitCSection.txtSlopeEntrance.Text = CDbl(pPropValue)
            Case "EXTINVERTLEV":
                FrmConduitCSection.txtSlopeExit.Text = CDbl(pPropValue)
        End Select
        
        '** Add decay factors
        If (Left(pPropName, 5) = "Decay") Then
            pDecayFactor.add pPropName, pPropValue
        End If
        
        pConduitDetailDict.add pPropName, pPropValue
        'Move to next row
        Set pRow = pCursor.NextRow
    Loop
    
    ' Load the Sediment Section values.................
    If pConduitDetailDict.Item("Transect") <> "" Then
        If GetListBoxIndex(cmbTransect, pConduitDetailDict.Item("Transect")) > -1 Then
            cmbTransect.Text = pConduitDetailDict.Item("Transect")
        End If
    End If
    txtWidth.Text = pConduitDetailDict.Item("Bed width")
    txtDepth.Text = pConduitDetailDict.Item("Bed depth")
    txtPorosity.Text = pConduitDetailDict.Item("Porosity")
    txtSand.Text = pConduitDetailDict.Item("Sand fraction")
    txtSilt.Text = pConduitDetailDict.Item("Silt fraction")
    txtClay.Text = pConduitDetailDict.Item("Clay fraction")
    txtSandDia.Text = pConduitDetailDict.Item("Sand effective diameter")
    txtSandVel.Text = pConduitDetailDict.Item("Sand velocity")
    txtSandDens.Text = pConduitDetailDict.Item("Sand density")
    txtSandCoeff.Text = pConduitDetailDict.Item("Sand coefficient")
    txtSandExp.Text = pConduitDetailDict.Item("Sand exponent")
    txtSiltDia.Text = pConduitDetailDict.Item("Silt effective diameter")
    txtSiltVel.Text = pConduitDetailDict.Item("Silt velocity")
    txtSiltDens.Text = pConduitDetailDict.Item("Silt density")
    txtSiltDep.Text = pConduitDetailDict.Item("Silt Deposition stress")
    txtSiltScour.Text = pConduitDetailDict.Item("Silt Scour stress")
    txtSiltErod.Text = pConduitDetailDict.Item("Silt Erodibility")
      
    txtClayDia.Text = pConduitDetailDict.Item("Clay effective diameter")
    txtClayVel.Text = pConduitDetailDict.Item("Clay velocity")
    txtClayDens.Text = pConduitDetailDict.Item("Clay density")
    txtClayDep.Text = pConduitDetailDict.Item("Clay Deposition stress")
    txtClayScour.Text = pConduitDetailDict.Item("Clay Scour stress")
    txtClayErod.Text = pConduitDetailDict.Item("Clay Erodibility")
    
    'If pdecayfactor is empty, initialize pollutant list, else load it
    If (pDecayFactor.Count > 0) Then
        Call LoadPollutantData(pDecayFactor)
    End If
    
   GoTo CleanUp
    
ShowError:
    MsgBox "ReadExistingConduitCrossSectionInformation :" & Err.description
    
CleanUp:
    Set pConduitDetailTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pConduitDetailDict = Nothing
    Set pDecayFactor = Nothing

End Sub

'* Initialize data grid with values
Private Sub InitializeDataGrid()
On Error GoTo ShowError

    '* get total pollutant count
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gParamInfos) + 1
    
    '* Find temporary table to store decay factor values
''    Dim pTableDF As iTable
''    Set pTableDF = GetInputDataTable("TempTable")
''    If (pTableDF Is Nothing) Then
''        Set pTableDF = CreateTEMPDBFTable("TempTable")
''        AddTableToMap pTableDF
''    Else
''        'delete all rows
''        pTableDF.DeleteSearchedRows Nothing
''    End If
''
''    '* Iterate over table
''    Dim iPollutant As Long
''    iPollutant = pTableDF.FindField("POLLUTANT")
''
''    Dim iValue As Long
''    iValue = pTableDF.FindField("DECAY")
''
''    Dim pRow As iRow
''    Dim iR As Integer
''
''    'Add all pollutants and their decay factors
''    For iR = 1 To pTotalPollutants
''        Set pRow = pTableDF.CreateRow
''        pRow.value(iPollutant) = gParamInfos(iR - 1).name
''        pRow.value(iValue) = CDbl(gParamInfos(iR - 1).Decay)
''        pRow.Store
''    Next
''    Set pRow = Nothing
''    Set pTableDF = Nothing
    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "POLLUTANT", adVarChar, 50
    oRs.Fields.Append "DECAY", adDouble
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    'Add all pollutants and their decay factors
    Dim iR As Integer
    For iR = 1 To pTotalPollutants
        oRs.AddNew
        oRs.Fields(0).value = gParamInfos(iR - 1).name
        oRs.Fields(1).value = CDbl(gParamInfos(iR - 1).Decay)
    Next
   
'    Dim oConn As New ADODB.Connection
'    oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
'           "SourceType=DBF;" & _
'           "SourceDB=" & gMapTempFolder & ";" & _
'           "Exclusive=No"
'    'Note: Specify the filename in the SQL statement. For example:
'    Dim oRs As New ADODB.Recordset
'    oRs.CursorLocation = adUseClient
'    oRs.Open "Select * From TempTable.dbf", oConn, adOpenDynamic, adLockOptimistic, adCmdText
'
    '* Set datagrid value
    Set DecayFactorGRID.DataSource = oRs
    DecayFactorGRID.ColumnHeaders = True
    DecayFactorGRID.Columns(0).Caption = "Pollutant"
    DecayFactorGRID.Columns(0).Locked = True
    DecayFactorGRID.Columns(0).Width = 2000
    
    DecayFactorGRID.Columns(1).Caption = "Decay Factor (1/hr)"
    DecayFactorGRID.Columns(1).Width = 1800
    
    GoTo CleanUp
ShowError:
    MsgBox "InitializeDataGrid: " & Err.description
CleanUp:


End Sub
