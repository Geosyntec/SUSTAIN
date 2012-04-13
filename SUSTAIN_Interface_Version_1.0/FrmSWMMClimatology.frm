VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSWMMClimatology 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meteorological Data"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "FrmSWMMClimatology.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin TabDlg.SSTab SSTabClimatology 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   882
      TabCaption(0)   =   "Temperature"
      TabPicture(0)   =   "FrmSWMMClimatology.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Evaporation"
      TabPicture(1)   =   "FrmSWMMClimatology.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Wind Speed"
      TabPicture(2)   =   "FrmSWMMClimatology.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Snow Melt"
      TabPicture(3)   =   "FrmSWMMClimatology.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Areal  Depletion"
      TabPicture(4)   =   "FrmSWMMClimatology.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   80
         Top             =   600
         Width           =   5655
         Begin MSComctlLib.ListView listPerviousArealDepletion 
            Height          =   2655
            Left            =   2880
            TabIndex        =   87
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CommandButton cmdArealPervNaturalArea 
            Caption         =   "Natural Area"
            Height          =   375
            Left            =   4200
            TabIndex        =   86
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdArealImperNaturalArea 
            Caption         =   "Natural Area"
            Height          =   375
            Left            =   1440
            TabIndex        =   85
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdArealPervNoDepletion 
            Caption         =   "No Depletion"
            Height          =   375
            Left            =   3000
            TabIndex        =   84
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CommandButton cmdArealImperNoDepletion 
            Caption         =   "No Depletion"
            Height          =   375
            Left            =   240
            TabIndex        =   83
            Top             =   3240
            Width           =   1215
         End
         Begin MSComctlLib.ListView listImperviousArealDepletion 
            Height          =   2655
            Left            =   120
            TabIndex        =   81
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label34 
            Caption         =   "Fraction of Area Covered by Snow"
            Height          =   255
            Left            =   1560
            TabIndex        =   82
            Top             =   150
            Width           =   2655
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   67
         Top             =   600
         Width           =   6375
         Begin VB.TextBox txtSNOWLongitudeCorr 
            Height          =   315
            Left            =   5280
            TabIndex        =   79
            Text            =   "0.0"
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox txtSNOWLatitude 
            Height          =   315
            Left            =   5280
            TabIndex        =   77
            Text            =   "50.0"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtSNOWElevation 
            Height          =   315
            Left            =   5280
            TabIndex        =   75
            Text            =   "0.0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox txtSNOWNegativeMeltRatio 
            Height          =   315
            Left            =   5280
            TabIndex        =   73
            Text            =   "0.6"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtSNOWATIWeight 
            Height          =   315
            Left            =   5280
            TabIndex        =   71
            Text            =   "0.5"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSNOWDivTemp 
            Height          =   315
            Left            =   5280
            TabIndex        =   69
            Text            =   "34"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label33 
            Caption         =   "Longitude Correction (+/- minutes)"
            Height          =   315
            Left            =   240
            TabIndex        =   78
            Top             =   2640
            Width           =   4335
         End
         Begin VB.Label Label32 
            Caption         =   "Latitude (degrees)"
            Height          =   315
            Left            =   240
            TabIndex        =   76
            Top             =   2160
            Width           =   4335
         End
         Begin VB.Label Label31 
            Caption         =   "Elevation above MSL (feet)"
            Height          =   315
            Left            =   240
            TabIndex        =   74
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Label Label30 
            Caption         =   "Negative Melt Ratio (fraction)"
            Height          =   315
            Left            =   240
            TabIndex        =   72
            Top             =   1200
            Width           =   4335
         End
         Begin VB.Label Label29 
            Caption         =   "ATI Weight (fraction)"
            Height          =   315
            Left            =   240
            TabIndex        =   70
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label Label28 
            Caption         =   "Dividing Temperature Between Snow and Rain ( degrees F)"
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   39
         Top             =   600
         Width           =   6375
         Begin VB.OptionButton optionWINDClimateFile 
            Caption         =   "From Climate File  (see Temperature Page)"
            Height          =   315
            Left            =   360
            TabIndex        =   53
            Top             =   360
            Value           =   -1  'True
            Width           =   3735
         End
         Begin VB.OptionButton optionWINDMonthlyAvg 
            Caption         =   "Monthly Averages"
            Height          =   315
            Left            =   360
            TabIndex        =   52
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtWINDJan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   360
            TabIndex        =   51
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDOct 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   50
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtWINDSep 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   49
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtWINDAug 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            TabIndex        =   48
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtWINDJul 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   360
            TabIndex        =   47
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtWINDJun 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   46
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDMay 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            TabIndex        =   45
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDApr 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   44
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDMar 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   43
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDFeb 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            TabIndex        =   42
            Top             =   1860
            Width           =   615
         End
         Begin VB.TextBox txtWINDDec 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   41
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox txtWINDNov 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            TabIndex        =   40
            Top             =   2460
            Width           =   615
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jan"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   360
            TabIndex        =   66
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Feb"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   960
            TabIndex        =   65
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mar"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   64
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Apr"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2160
            TabIndex        =   63
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "May"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2760
            TabIndex        =   62
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jun"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3360
            TabIndex        =   61
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jul"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   360
            TabIndex        =   60
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aug"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   960
            TabIndex        =   59
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sep"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   58
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Oct"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2160
            TabIndex        =   57
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nov"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2760
            TabIndex        =   56
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dec"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3360
            TabIndex        =   55
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "Monthly Windspeed (mph)"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   1200
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   6375
         Begin VB.TextBox txtEVAPNov 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            TabIndex        =   37
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPDec 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   36
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPFeb 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            TabIndex        =   35
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtEVAPMar 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   34
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtEVAPApr 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   33
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtEVAPMay 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            TabIndex        =   32
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtEVAPJun 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            TabIndex        =   31
            Top             =   2220
            Width           =   615
         End
         Begin VB.TextBox txtEVAPJul 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   360
            TabIndex        =   30
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPAug 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            TabIndex        =   29
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPSep 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            TabIndex        =   28
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPOct 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            TabIndex        =   27
            Top             =   2820
            Width           =   615
         End
         Begin VB.TextBox txtEVAPJan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   360
            TabIndex        =   26
            Top             =   2220
            Width           =   615
         End
         Begin VB.OptionButton optionEVAPMonthlyAvg 
            Caption         =   "Monthly Averages"
            Height          =   315
            Left            =   360
            TabIndex        =   13
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton optionEVAPClimateFile 
            Caption         =   "From Climate File  (see Temperature Page)"
            Height          =   315
            Left            =   360
            TabIndex        =   12
            Top             =   780
            Width           =   3735
         End
         Begin VB.TextBox txtEVAPConstant 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Text            =   "0.0"
            Top             =   360
            Width           =   915
         End
         Begin VB.OptionButton optionEVAPConstant 
            Caption         =   "Constant Value"
            Height          =   315
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.Label Label14 
            Caption         =   "Monthly Evaporation (in/day)"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dec"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3360
            TabIndex        =   25
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nov"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2760
            TabIndex        =   24
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Oct"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2160
            TabIndex        =   23
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sep"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aug"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   960
            TabIndex        =   21
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jul"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   360
            TabIndex        =   20
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jun"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3360
            TabIndex        =   19
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "May"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2760
            TabIndex        =   18
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Apr"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mar"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1560
            TabIndex        =   16
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Feb"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   960
            TabIndex        =   15
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jan"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   360
            TabIndex        =   14
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "(in/day)"
            Height          =   315
            Left            =   3480
            TabIndex        =   11
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Source of Temperature Data"
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   6375
         Begin VB.CommandButton cmdTEMPBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   5760
            TabIndex        =   7
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtTEMPClimateFile 
            Height          =   315
            Left            =   600
            TabIndex        =   6
            Top             =   1560
            Width           =   5055
         End
         Begin VB.OptionButton optionTEMPClimateFile 
            Caption         =   "Climate File"
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optionTEMPNoData 
            Caption         =   "No Data"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Value           =   -1  'True
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmSWMMClimatology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArealImperNaturalArea_Click()
    Dim pNaturalAreaValues
    pNaturalAreaValues = Array(0.1, 0.5, 0.65, 0.675, 0.75, 0.775, 0.85, 0.875, 0.925, 0.95)
    
    Dim iCount As Integer
    Dim iValue As Double
    Dim pItem As ListItem
    For iCount = 0 To 9
        iValue = pCount / 10
        Set pItem = listImperviousArealDepletion.ListItems.Item(iCount + 1)
        pItem.SubItems(1) = pNaturalAreaValues(iCount)
    Next
End Sub

Private Sub cmdArealImperNoDepletion_Click()
    Dim iCount As Integer
    Dim iValue As Double
    Dim pItem As ListItem
    For iCount = 0 To 9
        iValue = pCount / 10
        Set pItem = listImperviousArealDepletion.ListItems.Item(iCount + 1)
        pItem.SubItems(1) = "1.0"
    Next
End Sub

Private Sub cmdArealPervNaturalArea_Click()
    Dim pNaturalAreaValues
    pNaturalAreaValues = Array(0.1, 0.5, 0.65, 0.675, 0.75, 0.775, 0.85, 0.875, 0.925, 0.95)
    
    Dim iCount As Integer
    Dim iValue As Double
    Dim pItem As ListItem
    For iCount = 0 To 9
        iValue = pCount / 10
        Set pItem = listPerviousArealDepletion.ListItems.Item(iCount + 1)
        pItem.SubItems(1) = pNaturalAreaValues(iCount)
    Next
End Sub

Private Sub cmdArealPervNoDepletion_Click()
    Dim iCount As Integer
    Dim iValue As Double
    Dim pItem As ListItem
    For iCount = 0 To 9
        iValue = pCount / 10
        Set pItem = listPerviousArealDepletion.ListItems.Item(iCount + 1)
        pItem.SubItems(1) = "1.0"
    Next
End Sub

Private Sub cmdCancel_Click()
    '** close the dialog box
    Unload Me
End Sub

Private Sub cmdSave_Click()
    '** input data validation
    
    '** TEMPERATURE TAB
        Dim pTemperature As String
        pTemperature = ""
        If (optionTEMPNoData.value = True) Then
            pTemperature = "NO DATA"
        ElseIf (optionTEMPClimateFile.value = True) Then
            If (txtTEMPClimateFile.Text = "") Then
                SSTabClimatology.Tab = 0
                txtTEMPClimateFile.SetFocus
                MsgBox "Select climate data file"
                Exit Sub
            End If
            pTemperature = "FILE: " & """" & Trim(txtTEMPClimateFile.Text) & """"
        End If
    
    '** EVAPORATION TAB
        Dim pEvaporation As String
        pEvaporation = ""
        If (optionEVAPConstant.value = True) Then
            If (txtEVAPConstant.Text = "") Then
                SSTabClimatology.Tab = 1
                txtEVAPConstant.SetFocus
                MsgBox "Enter evaporation constant value"
                Exit Sub
            End If
            pEvaporation = "CONSTANT: " & txtEVAPConstant.Text
        ElseIf (optionEVAPMonthlyAvg.value = True) Then
            pEvaporation = "MONTHLY: " & txtEVAPJan.Text & "," & txtEVAPFeb.Text & "," & _
                           txtEVAPMar.Text & "," & txtEVAPApr.Text & "," & _
                           txtEVAPMay.Text & "," & txtEVAPJun.Text & "," & _
                           txtEVAPJul.Text & "," & txtEVAPAug.Text & "," & _
                           txtEVAPSep.Text & "," & txtEVAPOct.Text & "," & _
                           txtEVAPNov.Text & "," & txtEVAPDec.Text
        ElseIf (optionEVAPClimateFile.value = True) Then
            pEvaporation = "FILE: " & txtEVAPJan.Text & "," & txtEVAPFeb.Text & "," & _
                           txtEVAPMar.Text & "," & txtEVAPApr.Text & "," & _
                           txtEVAPMay.Text & "," & txtEVAPJun.Text & "," & _
                           txtEVAPJul.Text & "," & txtEVAPAug.Text & "," & _
                           txtEVAPSep.Text & "," & txtEVAPOct.Text & "," & _
                           txtEVAPNov.Text & "," & txtEVAPDec.Text
        End If
    
    '** WIND SPEED
        Dim pWindSpeed As String
        pWindSpeed = ""
        If (optionWINDClimateFile.value = True) Then
            If (optionTEMPClimateFile.value = False) Then
                 MsgBox "Specify climate data file on Temperature Tab"
                 Exit Sub
            End If
            pWindSpeed = "FILE: " & txtTEMPClimateFile.Text
        ElseIf (optionWINDMonthlyAvg.value = True) Then
            pWindSpeed = "MONTHLY: " & txtWINDJan.Text & "," & txtWINDFeb.Text & "," & _
                       txtWINDMar.Text & "," & txtWINDApr.Text & "," & _
                       txtWINDMay.Text & "," & txtWINDJun.Text & "," & _
                       txtWINDJul.Text & "," & txtWINDAug.Text & "," & _
                       txtWINDSep.Text & "," & txtWINDOct.Text & "," & _
                       txtWINDNov.Text & "," & txtWINDDec.Text
        End If
    
  
    '** AREAL DEPLETION - Impervious
    Dim pImpArealDepletion As String
    pImpArealDepletion = ""
    Dim iCount As Integer
    For iCount = 1 To 9
        pImpArealDepletion = pImpArealDepletion & listImperviousArealDepletion.ListItems.Item(iCount).SubItems(1) & ","
    Next
    pImpArealDepletion = pImpArealDepletion & listImperviousArealDepletion.ListItems.Item(10).SubItems(1)
    
    '** AREAL DEPLETION - Pervious
    Dim pPervArealDepletion As String
    pPervArealDepletion = ""
    For iCount = 1 To 9
        pPervArealDepletion = pPervArealDepletion & listPerviousArealDepletion.ListItems.Item(iCount).SubItems(1) & ","
    Next
    pPervArealDepletion = pPervArealDepletion & listPerviousArealDepletion.ListItems.Item(10).SubItems(1)
       
    '** get input values and write them to a dictionary
    Dim pPropertyDict As Scripting.Dictionary
    Set pPropertyDict = CreateObject("Scripting.Dictionary")
    pPropertyDict.add "Temperature", pTemperature
    pPropertyDict.add "Evaporation", pEvaporation
    pPropertyDict.add "Windspeed", pWindSpeed
    pPropertyDict.add "SnowDividingTemp", txtSNOWDivTemp.Text
    pPropertyDict.add "ATI Weight", txtSNOWATIWeight.Text
    pPropertyDict.add "Negative Melt Ratio", txtSNOWNegativeMeltRatio.Text
    pPropertyDict.add "Elevation", txtSNOWElevation.Text
    pPropertyDict.add "Latitude", txtSNOWLatitude.Text
    pPropertyDict.add "Longitude Correction", txtSNOWLongitudeCorr.Text
    pPropertyDict.add "ADC: IMPERVIOUS", pImpArealDepletion
    pPropertyDict.add "ADC: PERVIOUS", pPervArealDepletion
        
    '** write these values to a table
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDClimatology", 0, pPropertyDict
    
    '** cleanup
    Set pPropertyDict = Nothing
    
    '** close the form
    Unload Me
    
End Sub

Private Sub cmdTEMPBrowse_Click()
    CommonDialog.Filter = "Rain Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
    CommonDialog.ShowOpen
    txtTEMPClimateFile.Text = CommonDialog.FileName
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '** load the areal depletion list box
    listImperviousArealDepletion.ColumnHeaders.add , , "Depth Ratio", listImperviousArealDepletion.Width * 0.49
    listImperviousArealDepletion.ColumnHeaders.add , , "Impervious", listImperviousArealDepletion.Width * 0.48
    listPerviousArealDepletion.ColumnHeaders.add , , "Depth Ratio", listPerviousArealDepletion.Width * 0.49
    listPerviousArealDepletion.ColumnHeaders.add , , "Pervious", listPerviousArealDepletion.Width * 0.48
    
    Dim iCount As Integer
    Dim iValue As String
    Dim pItem As ListItem
    For iCount = 0 To 9
        iValue = CStr(CDbl(iCount / 10))
        Set pItem = listImperviousArealDepletion.ListItems.add(iCount + 1, , iValue)
        pItem.SubItems(1) = "1.0"
        Set pItem = listPerviousArealDepletion.ListItems.add(iCount + 1, , iValue)
        pItem.SubItems(1) = "1.0"
    Next
    
   '** call the subroutine to make the monthly average evaporation values disabled
   ChangeEvaporationMonthlyAvgEnable False, ""
   
   '** call the subroutine to make the monthly average wind speed values disabled
   ChangeWindSpeedMonthlyAvgEnable False
   
   '** call the subroutine to make the temperature control disabled
   ChangeTemperatureOptionControl False
   
   Dim pClimatologyDict As Scripting.Dictionary
   Set pClimatologyDict = LoadSWMMClimatologyDataToDictionary
   
   '** Load data back on the interface
   If Not (pClimatologyDict Is Nothing) Then
        '** TEMPERATURE DATA
        If (StringContains(pClimatologyDict.Item("Temperature"), "FILE")) Then
            optionTEMPClimateFile.value = True
            txtTEMPClimateFile.Text = Trim(Replace(Replace(pClimatologyDict.Item("Temperature"), "FILE:", ""), """", ""))
        Else
            optionTEMPNoData.value = True
        End If
        
        '** EVAPORATION DATA
        Dim pMonthlyEvap
        If (StringContains(pClimatologyDict.Item("Evaporation"), "CONSTANT")) Then
            optionEVAPConstant.value = True
            txtEVAPConstant.Text = Replace(pClimatologyDict.Item("Evaporation"), "CONSTANT:", "")
        ElseIf (StringContains(pClimatologyDict.Item("Evaporation"), "FILE")) Then
            optionEVAPClimateFile.value = True
            pMonthlyEvap = Split(Replace(pClimatologyDict.Item("Evaporation"), "FILE:", ""), ",")
            txtEVAPJan.Text = pMonthlyEvap(0)
            txtEVAPFeb.Text = pMonthlyEvap(1)
            txtEVAPMar.Text = pMonthlyEvap(2)
            txtEVAPApr.Text = pMonthlyEvap(3)
            txtEVAPMay.Text = pMonthlyEvap(4)
            txtEVAPJun.Text = pMonthlyEvap(5)
            txtEVAPJul.Text = pMonthlyEvap(6)
            txtEVAPAug.Text = pMonthlyEvap(7)
            txtEVAPSep.Text = pMonthlyEvap(8)
            txtEVAPOct.Text = pMonthlyEvap(9)
            txtEVAPNov.Text = pMonthlyEvap(10)
            txtEVAPDec.Text = pMonthlyEvap(11)
        Else
            optionEVAPMonthlyAvg.value = True
            pMonthlyEvap = Split(Replace(pClimatologyDict.Item("Evaporation"), "MONTHLY:", ""), ",")
            txtEVAPJan.Text = pMonthlyEvap(0)
            txtEVAPFeb.Text = pMonthlyEvap(1)
            txtEVAPMar.Text = pMonthlyEvap(2)
            txtEVAPApr.Text = pMonthlyEvap(3)
            txtEVAPMay.Text = pMonthlyEvap(4)
            txtEVAPJun.Text = pMonthlyEvap(5)
            txtEVAPJul.Text = pMonthlyEvap(6)
            txtEVAPAug.Text = pMonthlyEvap(7)
            txtEVAPSep.Text = pMonthlyEvap(8)
            txtEVAPOct.Text = pMonthlyEvap(9)
            txtEVAPNov.Text = pMonthlyEvap(10)
            txtEVAPDec.Text = pMonthlyEvap(11)
        End If
            
        
        '** WIND SPEED DATA
         If (StringContains(pClimatologyDict.Item("Windspeed"), "FILE")) Then
            optionWINDClimateFile.value = True
         Else
            optionWINDMonthlyAvg.value = True
            Dim pWindSpeed
            pWindSpeed = Replace(pClimatologyDict.Item("Windspeed"), "MONTHLY: ", "")
            Dim pMonthlyWindSpeed
            pMonthlyWindSpeed = Split(pWindSpeed, ",")
            txtWINDJan.Text = pMonthlyWindSpeed(0)
            txtWINDFeb.Text = pMonthlyWindSpeed(1)
            txtWINDMar.Text = pMonthlyWindSpeed(2)
            txtWINDApr.Text = pMonthlyWindSpeed(3)
            txtWINDMay.Text = pMonthlyWindSpeed(4)
            txtWINDJun.Text = pMonthlyWindSpeed(5)
            txtWINDJul.Text = pMonthlyWindSpeed(6)
            txtWINDAug.Text = pMonthlyWindSpeed(7)
            txtWINDSep.Text = pMonthlyWindSpeed(8)
            txtWINDOct.Text = pMonthlyWindSpeed(9)
            txtWINDNov.Text = pMonthlyWindSpeed(10)
            txtWINDDec.Text = pMonthlyWindSpeed(11)
         End If
   
   
         '** SNOW MELT DATA
         txtSNOWATIWeight.Text = pClimatologyDict.Item("ATI Weight")
         txtSNOWDivTemp.Text = pClimatologyDict.Item("SnowDividingTemp")
         txtSNOWElevation.Text = pClimatologyDict.Item("Elevation")
         txtSNOWLatitude.Text = pClimatologyDict.Item("Latitude")
         txtSNOWLongitudeCorr.Text = pClimatologyDict.Item("Longitude Correction")
         txtSNOWNegativeMeltRatio.Text = pClimatologyDict.Item("Negative Melt Ratio")
   
        '** AREAL DEPLETION DATA
        Dim pADCImpervious, pADCPervious
        pADCImpervious = Split(pClimatologyDict.Item("ADC: IMPERVIOUS"), ",")
        pADCPervious = Split(pClimatologyDict.Item("ADC: PERVIOUS"), ",")
   
        listImperviousArealDepletion.ListItems.Clear
        listPerviousArealDepletion.ListItems.Clear
        For iCount = LBound(pADCImpervious) To UBound(pADCImpervious)
             iValue = CStr(CDbl(iCount / 10))
             Set pItem = listImperviousArealDepletion.ListItems.add(iCount + 1, , iValue)
             pItem.SubItems(1) = pADCImpervious(iCount)
             Set pItem = listPerviousArealDepletion.ListItems.add(iCount + 1, , iValue)
             pItem.SubItems(1) = pADCPervious(iCount)
        Next
    
   End If
   
      
End Sub


Private Sub listImperviousArealDepletion_DblClick()
    Dim pItem As ListItem
    Set pItem = listImperviousArealDepletion.SelectedItem

    'Get existing value
    Dim pDefault
    pDefault = pItem.SubItems(1)
    'Get new input value
    Dim bValue
    bValue = InputBox("Enter value for Areal Depletion for Impervious Area", "Areal Depletion", pDefault)

    If (Trim(bValue) = "") Then
        bValue = pDefault
    End If
    pItem.SubItems(1) = bValue
End Sub

Private Sub listperviousArealDepletion_DblClick()
    Dim pItem As ListItem
    Set pItem = listPerviousArealDepletion.SelectedItem

    'Get existing value
    Dim pDefault
    pDefault = pItem.SubItems(1)
    'Get new input value
    Dim bValue
    bValue = InputBox("Enter value for Areal Depletion for Pervious Area", "Areal Depletion", pDefault)

    If (Trim(bValue) = "") Then
        bValue = pDefault
    End If
    pItem.SubItems(1) = bValue
End Sub

Private Sub optionEVAPClimateFile_Click()
   '** call the subroutine to make the monthly average evaporation values disabled
   ChangeEvaporationMonthlyAvgEnable True, "1.0"
   
   Dim pClimatologyDict As Scripting.Dictionary
   Set pClimatologyDict = LoadSWMMClimatologyDataToDictionary
   If pClimatologyDict Is Nothing Then Exit Sub
   
    '** EVAPORATION DATA
    Dim pMonthlyEvap
    If (StringContains(pClimatologyDict.Item("Evaporation"), "FILE")) Then
        pMonthlyEvap = Split(Replace(pClimatologyDict.Item("Evaporation"), "FILE:", ""), ",")
        txtEVAPJan.Text = pMonthlyEvap(0)
        txtEVAPFeb.Text = pMonthlyEvap(1)
        txtEVAPMar.Text = pMonthlyEvap(2)
        txtEVAPApr.Text = pMonthlyEvap(3)
        txtEVAPMay.Text = pMonthlyEvap(4)
        txtEVAPJun.Text = pMonthlyEvap(5)
        txtEVAPJul.Text = pMonthlyEvap(6)
        txtEVAPAug.Text = pMonthlyEvap(7)
        txtEVAPSep.Text = pMonthlyEvap(8)
        txtEVAPOct.Text = pMonthlyEvap(9)
        txtEVAPNov.Text = pMonthlyEvap(10)
        txtEVAPDec.Text = pMonthlyEvap(11)
    End If
    
    ' Change the label Text....
    Label14.Caption = "Pan Coefficients"
    
End Sub

Private Sub optionEVAPConstant_Click()
   '** call the subroutine to make the monthly average evaporation values disabled
   ChangeEvaporationMonthlyAvgEnable False, ""
End Sub

Private Sub optionEVAPMonthlyAvg_Click()
   '** call the subroutine to make them active
   ChangeEvaporationMonthlyAvgEnable True, "0.0"
   
   Dim pClimatologyDict As Scripting.Dictionary
   Set pClimatologyDict = LoadSWMMClimatologyDataToDictionary
   If pClimatologyDict Is Nothing Then Exit Sub
   
    '** EVAPORATION DATA
    Dim pMonthlyEvap
    If (StringContains(pClimatologyDict.Item("Evaporation"), "MONTHLY")) Then
        pMonthlyEvap = Split(Replace(pClimatologyDict.Item("Evaporation"), "MONTHLY:", ""), ",")
        txtEVAPJan.Text = pMonthlyEvap(0)
        txtEVAPFeb.Text = pMonthlyEvap(1)
        txtEVAPMar.Text = pMonthlyEvap(2)
        txtEVAPApr.Text = pMonthlyEvap(3)
        txtEVAPMay.Text = pMonthlyEvap(4)
        txtEVAPJun.Text = pMonthlyEvap(5)
        txtEVAPJul.Text = pMonthlyEvap(6)
        txtEVAPAug.Text = pMonthlyEvap(7)
        txtEVAPSep.Text = pMonthlyEvap(8)
        txtEVAPOct.Text = pMonthlyEvap(9)
        txtEVAPNov.Text = pMonthlyEvap(10)
        txtEVAPDec.Text = pMonthlyEvap(11)
    End If
    
    ' Change the label Text....
    Label14.Caption = "Monthly Evaporation (in/day)"
    
End Sub

Private Sub optionTEMPClimateFile_Click()
   '** call the subroutine to make the temperature control enable
   ChangeTemperatureOptionControl True
End Sub

Private Sub optionTEMPNoData_Click()
   '** call the subroutine to make the temperature control disabled
   ChangeTemperatureOptionControl False
End Sub

Private Sub optionWINDClimateFile_Click()
   '** call the subroutine to make the monthly average wind speed values disabled
   ChangeWindSpeedMonthlyAvgEnable False
End Sub

Private Sub optionWINDMonthlyAvg_Click()
   '** call the subroutine to make the monthly average wind speed values enabled
   ChangeWindSpeedMonthlyAvgEnable True
   
   Dim pClimatologyDict As Scripting.Dictionary
   Set pClimatologyDict = LoadSWMMClimatologyDataToDictionary
   If pClimatologyDict Is Nothing Then Exit Sub
   
   '** Load data back on the interface
        
    '** WIND SPEED DATA
    If (StringContains(pClimatologyDict.Item("Windspeed"), "MONTHLY")) Then
        Dim pWindSpeed
        pWindSpeed = Replace(pClimatologyDict.Item("Windspeed"), "MONTHLY: ", "")
        Dim pMonthlyWindSpeed
        pMonthlyWindSpeed = Split(pWindSpeed, ",")
        txtWINDJan.Text = pMonthlyWindSpeed(0)
        txtWINDFeb.Text = pMonthlyWindSpeed(1)
        txtWINDMar.Text = pMonthlyWindSpeed(2)
        txtWINDApr.Text = pMonthlyWindSpeed(3)
        txtWINDMay.Text = pMonthlyWindSpeed(4)
        txtWINDJun.Text = pMonthlyWindSpeed(5)
        txtWINDJul.Text = pMonthlyWindSpeed(6)
        txtWINDAug.Text = pMonthlyWindSpeed(7)
        txtWINDSep.Text = pMonthlyWindSpeed(8)
        txtWINDOct.Text = pMonthlyWindSpeed(9)
        txtWINDNov.Text = pMonthlyWindSpeed(10)
        txtWINDDec.Text = pMonthlyWindSpeed(11)
    End If

End Sub

Private Sub ChangeEvaporationMonthlyAvgEnable(bEnable As Boolean, bValue As String)
    '** change the enable value to the passed parameter
    txtEVAPJan.Enabled = bEnable
    txtEVAPFeb.Enabled = bEnable
    txtEVAPMar.Enabled = bEnable
    txtEVAPApr.Enabled = bEnable
    txtEVAPMay.Enabled = bEnable
    txtEVAPJun.Enabled = bEnable
    txtEVAPJul.Enabled = bEnable
    txtEVAPAug.Enabled = bEnable
    txtEVAPSep.Enabled = bEnable
    txtEVAPOct.Enabled = bEnable
    txtEVAPNov.Enabled = bEnable
    txtEVAPDec.Enabled = bEnable
    If (bEnable = True) Then
        txtEVAPJan.Text = bValue
        txtEVAPFeb.Text = bValue
        txtEVAPMar.Text = bValue
        txtEVAPApr.Text = bValue
        txtEVAPMay.Text = bValue
        txtEVAPJun.Text = bValue
        txtEVAPJul.Text = bValue
        txtEVAPAug.Text = bValue
        txtEVAPSep.Text = bValue
        txtEVAPOct.Text = bValue
        txtEVAPNov.Text = bValue
        txtEVAPDec.Text = bValue
    End If
    
End Sub



Private Sub ChangeWindSpeedMonthlyAvgEnable(bEnable As Boolean)
    '** change the enable value to the passed parameter
    txtWINDJan.Enabled = bEnable
    txtWINDFeb.Enabled = bEnable
    txtWINDMar.Enabled = bEnable
    txtWINDApr.Enabled = bEnable
    txtWINDMay.Enabled = bEnable
    txtWINDJun.Enabled = bEnable
    txtWINDJul.Enabled = bEnable
    txtWINDAug.Enabled = bEnable
    txtWINDSep.Enabled = bEnable
    txtWINDOct.Enabled = bEnable
    txtWINDNov.Enabled = bEnable
    txtWINDDec.Enabled = bEnable
    If (bEnable = True) Then
        txtWINDJan.Text = "0.0"
        txtWINDFeb.Text = "0.0"
        txtWINDMar.Text = "0.0"
        txtWINDApr.Text = "0.0"
        txtWINDMay.Text = "0.0"
        txtWINDJun.Text = "0.0"
        txtWINDJul.Text = "0.0"
        txtWINDAug.Text = "0.0"
        txtWINDSep.Text = "0.0"
        txtWINDOct.Text = "0.0"
        txtWINDNov.Text = "0.0"
        txtWINDDec.Text = "0.0"
    End If
End Sub


Private Sub ChangeTemperatureOptionControl(bEnable As Boolean)
    '** change the enable value to the passed parameter
    txtTEMPClimateFile.Enabled = bEnable
    cmdTEMPBrowse.Enabled = bEnable
End Sub



