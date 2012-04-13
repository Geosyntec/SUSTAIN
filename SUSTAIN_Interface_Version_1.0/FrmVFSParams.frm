VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVFSParams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Buffer Strip Parameters"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "FrmVFSParams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   7440
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Overland Flow"
      TabPicture(0)   =   "FrmVFSParams.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtVFSID"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Infiltration Properties"
      TabPicture(1)   =   "FrmVFSParams.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vegetation Properties"
      TabPicture(2)   =   "FrmVFSParams.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Sediment Characteristics"
      TabPicture(3)   =   "FrmVFSParams.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "WQ Parameters"
      TabPicture(4)   =   "FrmVFSParams.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label17"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "DataGridSedDec"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "DataGridDissDec"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Cost Factors"
      TabPicture(5)   =   "FrmVFSParams.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label78"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label79"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lstComponents"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Frame7"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cmdRemove"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdAdd"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "txtSourceDetails"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "cmdEdit"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).ControlCount=   8
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   -68160
         TabIndex        =   93
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtSourceDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   -69480
         MultiLine       =   -1  'True
         TabIndex        =   87
         Top             =   915
         Width           =   3615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   -69360
         TabIndex        =   86
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   -66960
         TabIndex        =   85
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         Caption         =   "Select components and sources from the list"
         Height          =   4215
         Left            =   -74760
         TabIndex        =   68
         Top             =   480
         Width           =   5175
         Begin VB.TextBox txtCostUnits 
            Height          =   285
            Left            =   2040
            TabIndex        =   92
            Text            =   "1"
            Top             =   3000
            Width           =   2895
         End
         Begin VB.ComboBox cbxUnit 
            Height          =   315
            Left            =   2040
            TabIndex        =   77
            Top             =   2595
            Width           =   2895
         End
         Begin VB.CheckBox chkNRCS 
            Caption         =   "Include NRCS Sources"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ComboBox cbxComponent 
            Height          =   315
            Left            =   2040
            TabIndex        =   75
            Top             =   600
            Width           =   2895
         End
         Begin VB.ComboBox cbxSource 
            Height          =   315
            Left            =   2040
            TabIndex        =   74
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox txtCost 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   73
            Top             =   3360
            Width           =   2895
         End
         Begin VB.ComboBox cbxYear 
            Height          =   315
            Left            =   2040
            TabIndex        =   72
            Top             =   2190
            Width           =   2895
         End
         Begin VB.ComboBox cbxLocation 
            Height          =   315
            Left            =   2040
            TabIndex        =   71
            Top             =   1395
            Width           =   2895
         End
         Begin VB.CheckBox chkCCI 
            Caption         =   "Adjust cost based on ENR Construction Cost Index"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   3720
            Value           =   1  'Checked
            Width           =   4575
         End
         Begin VB.TextBox txtUserComponent 
            Height          =   285
            Left            =   2040
            TabIndex        =   69
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label15 
            Caption         =   "Number of Units (Per unit)"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label77 
            Caption         =   "Functional Components"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label76 
            Caption         =   "Source"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label75 
            Caption         =   "Unit"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   2595
            Width           =   2415
         End
         Begin VB.Label Label74 
            Caption         =   "Unit Cost"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label73 
            Caption         =   "Source Year"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   2190
            Width           =   1815
         End
         Begin VB.Label Label72 
            Caption         =   "Source Locale"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   1395
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "User Defined Component"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buffer Dimensions"
         Height          =   1455
         Left            =   3720
         TabIndex        =   59
         Top             =   2580
         Width           =   4455
         Begin VB.TextBox BufferWidth 
            Height          =   330
            Left            =   3240
            TabIndex        =   61
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox BufferLength 
            Height          =   330
            Left            =   3240
            TabIndex        =   60
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Width of the Strip [FWIDTH] (ft)"
            Height          =   210
            Left            =   240
            TabIndex        =   63
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label10 
            Caption         =   "Buffer Length [VL]  (ft) "
            Height          =   210
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Information"
         Height          =   855
         Left            =   3720
         TabIndex        =   56
         Top             =   1500
         Width           =   4455
         Begin VB.TextBox txtName 
            Height          =   330
            Left            =   1200
            TabIndex        =   57
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtVFSID 
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Text            =   "HIDDEN"
         Top             =   3900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         Picture         =   "FrmVFSParams.frx":0972
         ScaleHeight     =   2535
         ScaleWidth      =   3255
         TabIndex        =   54
         Top             =   1380
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         Caption         =   "Green Ampt Infiltration Parameters"
         Height          =   3255
         Left            =   -74760
         TabIndex        =   41
         Top             =   840
         Width           =   7815
         Begin VB.TextBox VKS 
            Height          =   285
            Left            =   5280
            TabIndex        =   47
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox Sav 
            Height          =   285
            Left            =   5280
            TabIndex        =   46
            Top             =   912
            Width           =   1000
         End
         Begin VB.TextBox OI 
            Height          =   285
            Left            =   5280
            TabIndex        =   45
            Top             =   1344
            Width           =   1000
         End
         Begin VB.TextBox OS 
            Height          =   285
            Left            =   5280
            TabIndex        =   44
            Top             =   1776
            Width           =   1000
         End
         Begin VB.TextBox SM 
            Height          =   285
            Left            =   5280
            TabIndex        =   43
            Top             =   2208
            Width           =   1000
         End
         Begin VB.TextBox SCHK 
            Height          =   285
            Left            =   5280
            TabIndex        =   42
            Top             =   2640
            Width           =   1000
         End
         Begin VB.Label Label2 
            Caption         =   "Vertical Saturated Hydraulic Conductivity, VKS (in/hr)"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label3 
            Caption         =   "Average Suction at the Wetting Front, Sav (ft)"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   912
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   "Initial Water Content, OI (fraction)"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   1344
            Width           =   3735
         End
         Begin VB.Label Label5 
            Caption         =   "Saturated Water Content, OS (fraction)"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   1776
            Width           =   3855
         End
         Begin VB.Label Label6 
            Caption         =   "Maximum Surface Storage, SM (ft)"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   2208
            Width           =   3135
         End
         Begin VB.Label Label7 
            Caption         =   "Fraction of the filter where ponding is checked, (0 <= SCHK <=1)"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   2640
            Width           =   5055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Vegetation Properties"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   27
         Top             =   780
         Width           =   8295
         Begin VB.TextBox SS 
            Height          =   285
            Left            =   3480
            TabIndex        =   36
            Top             =   360
            Width           =   1000
         End
         Begin VB.TextBox H 
            Height          =   285
            Left            =   3480
            TabIndex        =   35
            Top             =   840
            Width           =   1000
         End
         Begin VB.TextBox VN 
            Height          =   285
            Left            =   3480
            TabIndex        =   34
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox Vn2 
            Height          =   285
            Left            =   3480
            TabIndex        =   33
            Top             =   1800
            Width           =   1000
         End
         Begin VB.Frame Frame8 
            Caption         =   "Grass Segments"
            Height          =   2175
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   7695
            Begin VB.TextBox NPROP 
               Height          =   285
               Left            =   3120
               TabIndex        =   31
               Top             =   360
               Width           =   975
            End
            Begin VB.CommandButton cmdUpdateDG 
               Caption         =   "Update Segment Grid"
               Height          =   375
               Left            =   4320
               TabIndex        =   29
               Top             =   360
               Width           =   1815
            End
            Begin MSDataGridLib.DataGrid DataGridSegments 
               Height          =   1215
               Left            =   120
               TabIndex        =   30
               Top             =   840
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   2143
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
            Begin VB.Label Label27 
               Caption         =   "Number of Segments, NPROP"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   2535
            End
         End
         Begin VB.Label Label8 
            Caption         =   "Spacing for Grass Stems, SS (in)"
            Height          =   255
            Left            =   360
            TabIndex        =   40
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label11 
            Caption         =   "Height of Grass, H (in)"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label12 
            Caption         =   "Grass Manning's n -VN"
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   1320
            Width           =   3735
         End
         Begin VB.Label Label13 
            Caption         =   "Bare Surface Manning's n - Vn2"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   1800
            Width           =   3855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Incoming Sediment Properties"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   3
         Top             =   900
         Width           =   9135
         Begin VB.TextBox PORSand 
            Height          =   285
            Left            =   5040
            TabIndex        =   18
            Top             =   795
            Width           =   1000
         End
         Begin VB.TextBox NPARTSand 
            Height          =   285
            Left            =   5040
            TabIndex        =   17
            Top             =   1350
            Width           =   1000
         End
         Begin VB.TextBox COARSESand 
            Height          =   285
            Left            =   5040
            TabIndex        =   16
            Top             =   1920
            Width           =   1000
         End
         Begin VB.TextBox DPSand 
            Height          =   285
            Left            =   5040
            TabIndex        =   15
            Top             =   2445
            Width           =   1000
         End
         Begin VB.TextBox SGSand 
            Height          =   285
            Left            =   5040
            TabIndex        =   14
            Top             =   3000
            Width           =   1000
         End
         Begin VB.TextBox SGSilt 
            Height          =   285
            Left            =   6240
            TabIndex        =   13
            Top             =   2985
            Width           =   1000
         End
         Begin VB.TextBox DPSilt 
            Height          =   285
            Left            =   6240
            TabIndex        =   12
            Top             =   2430
            Width           =   1000
         End
         Begin VB.TextBox COARSESilt 
            Height          =   285
            Left            =   6240
            TabIndex        =   11
            Top             =   1905
            Width           =   1000
         End
         Begin VB.TextBox NPARTSilt 
            Height          =   285
            Left            =   6240
            TabIndex        =   10
            Top             =   1335
            Width           =   1000
         End
         Begin VB.TextBox PORSilt 
            Height          =   285
            Left            =   6240
            TabIndex        =   9
            Top             =   780
            Width           =   1000
         End
         Begin VB.TextBox SGClay 
            Height          =   285
            Left            =   7680
            TabIndex        =   8
            Top             =   2985
            Width           =   1000
         End
         Begin VB.TextBox DPClay 
            Height          =   285
            Left            =   7680
            TabIndex        =   7
            Top             =   2430
            Width           =   1000
         End
         Begin VB.TextBox COARSEClay 
            Height          =   285
            Left            =   7680
            TabIndex        =   6
            Top             =   1905
            Width           =   1000
         End
         Begin VB.TextBox NPARTClay 
            Height          =   285
            Left            =   7680
            TabIndex        =   5
            Top             =   1335
            Width           =   1000
         End
         Begin VB.TextBox PORClay 
            Height          =   285
            Left            =   7680
            TabIndex        =   4
            Top             =   780
            Width           =   1000
         End
         Begin VB.Label Label21 
            Caption         =   "Porosity of Deposited Sediment, POR (fraction)"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   3855
         End
         Begin VB.Label Label22 
            Caption         =   "Sediment Particle Class According to the USDA, NPART"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1350
            Width           =   4695
         End
         Begin VB.Label Label23 
            Caption         =   "Portion of Particle with Diameter > 0.0037 cm, COARSE"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   1890
            Width           =   4335
         End
         Begin VB.Label Label24 
            Caption         =   "Sediment Particle Size d50, DP (in)"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   2445
            Width           =   4335
         End
         Begin VB.Label Label25 
            Caption         =   "Sediment Particle Density, SG (lb/ft^3)"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   3000
            Width           =   3375
         End
         Begin VB.Label Label20 
            Caption         =   "Sand"
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
            Left            =   5040
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "Silt"
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
            Left            =   6240
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "Clay"
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
            Left            =   7680
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid DataGridDissDec 
         Height          =   1605
         Left            =   -74640
         TabIndex        =   64
         Top             =   3300
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2831
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
      Begin MSDataGridLib.DataGrid DataGridSedDec 
         Height          =   1605
         Left            =   -74640
         TabIndex        =   65
         Top             =   1260
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2831
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
      Begin MSComctlLib.ListView lstComponents 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   89
         Top             =   5040
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Component"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Component_ID"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Locale"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Source"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Year"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Volume Types"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Num Units"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Unit Cost"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Adj Cost"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label79 
         Caption         =   "Selected components"
         Height          =   255
         Left            =   -74760
         TabIndex        =   90
         Top             =   4800
         Width           =   3495
      End
      Begin VB.Label Label78 
         Caption         =   "Source Details"
         Height          =   255
         Left            =   -69480
         TabIndex        =   88
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Define first order decay (/day) and temperature correction factor for sediment fraction"
         Height          =   375
         Left            =   -74640
         TabIndex        =   67
         Top             =   900
         Width           =   6735
      End
      Begin VB.Label Label26 
         Caption         =   "Define first order decay (/day) and temperature correction factor for dissolved fraction"
         Height          =   255
         Left            =   -74640
         TabIndex        =   66
         Top             =   3060
         Width           =   6855
      End
   End
End
Attribute VB_Name = "FrmVFSParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Public bContinue As Boolean
Private pAdoConn As ADODB.Connection
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmVFSParams.frm"
Private costAdjDict As Scripting.Dictionary
Private maxCCIYear As Integer

'Private Sub cbxComponent_Click()
'  On Error GoTo ErrorHandler
'
'
'    Dim pRs As ADODB.Recordset
'    Set pRs = New ADODB.Recordset
'
'    Dim strSqlLocale As String
'    cbxLocation.Clear
'
'    If cbxComponent.List(cbxComponent.ListIndex) = "User Defined" Then
'        txtUserComponent.Enabled = True
'    Else
'        txtUserComponent.Text = ""
'        txtUserComponent.Enabled = False
'        If cbxComponent.List(cbxComponent.ListIndex) <> "Land Cost" Then
'            strSqlLocale = "SELECT DISTINCT Component_Costs.Locale " & _
'                " FROM (BMP_Components INNER JOIN Component_Costs ON BMP_Components.Component_ID = Component_Costs.Component_ID) INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID " & _
'                " WHERE BMP_Components.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
'                " AND Component_Costs.Cost > 0"
'
'            If chkNRCS.value = 0 Then
'                strSqlLocale = strSqlLocale & " AND  Reference_Sources.Author NOT LIKE '%NRCS%'"
'            End If
'
'            strSqlLocale = strSqlLocale & " ORDER BY Component_Costs.Locale"
'
'            pRs.Open strSqlLocale, pAdoConn, adOpenDynamic, adLockOptimistic
'
'
'            If pRs.EOF Then
'                cbxSource.Clear
'                cbxYear.Clear
'                'cbxUnit.Clear
'                txtCost.Text = ""
'                'txtUnit.Text = ""
'                Exit Sub
'            End If
'
'            pRs.MoveFirst
'            Do Until pRs.EOF
'                cbxLocation.AddItem pRs("Locale")
'                pRs.MoveNext
'            Loop
'            pRs.Close
'            cbxLocation.AddItem "Average of available data"
'        End If
'    End If
'
'    cbxLocation.AddItem "User Defined"
'    cbxLocation.ListIndex = 0
'
'  Exit Sub
'ErrorHandler:
'  HandleError True, "cbxComponent_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
'End Sub
'
'Private Sub cbxLocation_Click()
'  On Error GoTo ErrorHandler
'
'
'    Dim pRs As ADODB.Recordset
'    Set pRs = New ADODB.Recordset
'
'    Dim strSqlSource As String
'    cbxSource.Clear
'
'    If cbxLocation.List(cbxLocation.ListIndex) = "Average of available data" Then
'        cbxSource.AddItem "Average"
'    ElseIf cbxLocation.List(cbxLocation.ListIndex) <> "User Defined" Then
'        strSqlSource = "SELECT DISTINCT Reference_Sources.Author" & _
'            " FROM Component_Costs INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID " & _
'            " WHERE Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
'            " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
'            " AND Component_Costs.Use_Flag <> 0 "
'
'        pRs.Open strSqlSource, pAdoConn, adOpenDynamic, adLockOptimistic
'
'        pRs.MoveFirst
'        Do Until pRs.EOF
'            'cbxSource.AddItem pRs("Author") & " - " & pRs("Title")
'            cbxSource.AddItem pRs("Author")
'            'cbxSource.ItemData(cbxSource.NewIndex) = pRs("Reference_Source_ID")
'            pRs.MoveNext
'        Loop
'        pRs.Close
'    End If
'    'Add another option for user defined
'    cbxSource.AddItem "User Defined"
'    cbxSource.ListIndex = 0
'
'  Exit Sub
'ErrorHandler:
'  HandleError True, "cbxLocation_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
'End Sub
'
'
'Private Sub cbxSource_Click()
'  On Error GoTo ErrorHandler
'    cbxYear.Clear
'
'    Dim pRs As ADODB.Recordset
'    Set pRs = New ADODB.Recordset
'
'    Dim strSqlYear As String
'    Dim strSource As String
'    strSource = cbxSource.List(cbxSource.ListIndex)
'
'    If strSource = "User Defined" Then
'        cbxYear.AddItem Year(Now)
'    Else
'        If strSource = "Average" Then
'            strSqlYear = "SELECT DISTINCT Year " & _
'                " FROM Average_Costs " & _
'                " WHERE Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex)
'        Else 'strSource <> "User Defined" Then
'            strSqlYear = "SELECT DISTINCT Reference_Sources.Year " & _
'                " FROM Component_Costs INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID" & _
'                " WHERE Reference_Sources.Author ='" & cbxSource.List(cbxSource.ListIndex) & "'" & _
'                " AND Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
'                " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
'                " AND Component_Costs.Use_Flag <> 0"
'        End If
'        pRs.Open strSqlYear, pAdoConn, adOpenDynamic, adLockOptimistic
'
'        pRs.MoveFirst
'        Do Until pRs.EOF
'            cbxYear.AddItem pRs("Year")
'            'cbxYear.ItemData(cbxYear.NewIndex) = pRs("Reference_Source_ID")
'            pRs.MoveNext
'        Loop
'        pRs.Close
'    End If
'
'    cbxYear.ListIndex = 0
'
'
'  Exit Sub
'ErrorHandler:
'  HandleError True, "cbxSource_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
'End Sub
'
'
'

'
'Private Sub cbxYear_Click()
'  On Error GoTo ErrorHandler
'    Dim strSource As String
'    strSource = cbxSource.List(cbxSource.ListIndex)
'
'    Dim pRs As ADODB.Recordset
'    Set pRs = New ADODB.Recordset
'
'    Dim strSql As String
'    If strSource <> "User Defined" And strSource <> "Average" Then
'
'        strSql = "SELECT DISTINCT Reference_Sources.Title, " & _
'            " Reference_Sources.Year, Reference_Sources.Author, " & _
'            " Reference_Sources.Publication_Street, " & _
'            " Reference_Sources.Publication_City, " & _
'            " Reference_Sources.Publication_State, " & _
'            " Reference_Sources.Reference_Number, " & _
'            " Reference_Sources.Prepared_By, " & _
'            " UnitTypes.UnitType_Desc, Component_Costs.Unit, Component_Costs.Cost " & _
'            " FROM (Component_Costs INNER JOIN Reference_Sources ON " & _
'            " Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID) " & _
'            " INNER JOIN UnitTypes ON Component_Costs.Unit = UnitTypes.UnitType_Code " & _
'            " WHERE Reference_Sources.Author ='" & cbxSource.List(cbxSource.ListIndex) & "'" & _
'            " AND Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
'            " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
'            " AND Reference_Sources.Year =" & cbxYear.List(cbxYear.ListIndex)
'
'        pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
'
'        pRs.MoveFirst
'        If Not IsNull(pRs("Cost")) Then
'            txtCost.Text = pRs("Cost")
'        Else
'            txtCost.Text = ""
'        End If
'
'        If Not IsNull(pRs("UnitType_Desc")) Then
'            Set_Unit (pRs("UnitType_Desc"))
'        Else
'            Set_Unit ("Unknown")
'        End If
'
'        txtSourceDetails = pRs("Author") & " . " & _
'                pRs("Year") & ". " & pRs("Title") & ". " & _
'                pRs("Publication_Street") & ". " & _
'                pRs("Publication_City") & ". " & _
'                pRs("Publication_State") & ". " & _
'                pRs("Reference_Number")
'        pRs.Close
'        txtCost.Enabled = False
'    ElseIf strSource = "Average" Then
'        'Find the average...
'
'    Else
'        Set_Unit ("Unknown")
'        txtSourceDetails = ""
'        txtCost.Enabled = True
'        txtCost.Text = ""
'    End If
'  Exit Sub
'ErrorHandler:
'  HandleError True, "cbxYear_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
'End Sub

Private Sub cbxComponent_Click()
  On Error GoTo ErrorHandler

    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
        
    Dim pRsAvg As ADODB.Recordset
    
    Dim strSqlLocale As String
    cbxLocation.Clear
       
    If cbxComponent.List(cbxComponent.ListIndex) = "User Defined" Then
        txtUserComponent.Enabled = True
    Else
        txtUserComponent.Text = ""
        txtUserComponent.Enabled = False
        If cbxComponent.List(cbxComponent.ListIndex) <> "Land Cost" Then
            strSqlLocale = "SELECT DISTINCT Component_Costs.Locale " & _
                " FROM (BMP_Components INNER JOIN Component_Costs ON BMP_Components.Component_ID = Component_Costs.Component_ID) INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID " & _
                " WHERE BMP_Components.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
                " AND Component_Costs.Cost > 0 AND Component_Costs.Use_Flag <> 0"
        
        
        '    strSqlLocale = SELECT Component_Costs.Locale, Reference_Sources.Title, Reference_Sources.Author
        '        FROM (BMP_Components INNER JOIN Component_Costs ON BMP_Components.Component_ID = Component_Costs.Component_ID) INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID;
        
            If chkNRCS.value = 0 Then
                strSqlLocale = strSqlLocale & " AND  Reference_Sources.Author NOT LIKE '%NRCS%'"
            End If
            
            strSqlLocale = strSqlLocale & " ORDER BY Component_Costs.Locale"
            
            pRs.Open strSqlLocale, pAdoConn, adOpenDynamic, adLockOptimistic
                        
            If pRs.EOF Then
                cbxSource.Clear
                cbxYear.Clear
                'cbxUnit.Clear
                txtCost.Text = ""
                'txtUnit.Text = ""
                Exit Sub
            End If
            
            pRs.MoveFirst
            Do Until pRs.EOF
                cbxLocation.AddItem pRs("Locale")
                pRs.MoveNext
            Loop
            pRs.Close
            
            Dim strSqlAvg As String
            
            strSqlAvg = " SELECT Count(*) " & _
                    " FROM Cost_Unit_Check " & _
                    " WHERE [Component_ID] = " & cbxComponent.ItemData(cbxComponent.ListIndex)
            
            Set pRsAvg = New ADODB.Recordset
            pRsAvg.Open strSqlAvg, pAdoConn, adOpenDynamic, adLockOptimistic
            pRsAvg.MoveFirst
            If pRsAvg(0).value > 0 Then cbxLocation.AddItem "Average of available data"
            pRsAvg.Close
        End If
    End If
    
    cbxLocation.AddItem "User Defined"
    cbxLocation.ListIndex = 0
    

  GoTo CleanUp
ErrorHandler:
  HandleError True, "cbxComponent_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
CleanUp:
    Set pRs = Nothing
    Set pRsAvg = Nothing
End Sub

Private Sub cbxLocation_Click()
  On Error GoTo ErrorHandler

   
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
    
    Dim strSqlSource As String
    cbxSource.Clear
    
    If cbxLocation.List(cbxLocation.ListIndex) = "Average of available data" Then
        cbxSource.AddItem "Average"
    ElseIf cbxLocation.List(cbxLocation.ListIndex) <> "User Defined" Then
        strSqlSource = "SELECT DISTINCT Reference_Sources.Author" & _
            " FROM Component_Costs INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID " & _
            " WHERE Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
            " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
            " AND Component_Costs.Use_Flag <> 0 "
    
        pRs.Open strSqlSource, pAdoConn, adOpenDynamic, adLockOptimistic
                   
        pRs.MoveFirst
        Do Until pRs.EOF
            'cbxSource.AddItem pRs("Author") & " - " & pRs("Title")
            cbxSource.AddItem pRs("Author")
            'cbxSource.ItemData(cbxSource.NewIndex) = pRs("Reference_Source_ID")
            pRs.MoveNext
        Loop
        pRs.Close
    End If
    'Add another option for user defined
    cbxSource.AddItem "User Defined"
    cbxSource.ListIndex = 0
    
  Exit Sub
ErrorHandler:
  HandleError True, "cbxLocation_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


Private Sub cbxSource_Click()
  On Error GoTo ErrorHandler
    cbxYear.Clear
    
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
    
    Dim strSqlYear As String
    Dim strSource As String
    strSource = cbxSource.List(cbxSource.ListIndex)
    
    If strSource = "User Defined" Then
        cbxYear.AddItem Year(Now)
    Else
        If strSource = "Average" Then
            strSqlYear = "SELECT DISTINCT Year " & _
                " FROM Average_Costs " & _
                " WHERE Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex)
        Else 'strSource <> "User Defined" Then
            strSqlYear = "SELECT DISTINCT Reference_Sources.Year " & _
                " FROM Component_Costs INNER JOIN Reference_Sources ON Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID" & _
                " WHERE Reference_Sources.Author ='" & cbxSource.List(cbxSource.ListIndex) & "'" & _
                " AND Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
                " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
                " AND Component_Costs.Use_Flag <> 0"
        End If
        pRs.Open strSqlYear, pAdoConn, adOpenDynamic, adLockOptimistic
                   
        pRs.MoveFirst
        Do Until pRs.EOF
            cbxYear.AddItem pRs("Year")
            'cbxYear.ItemData(cbxYear.NewIndex) = pRs("Reference_Source_ID")
            pRs.MoveNext
        Loop
        pRs.Close
    End If
    
    cbxYear.ListIndex = 0
    
  Exit Sub
ErrorHandler:
  HandleError True, "cbxSource_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


Private Sub cbxYear_Click()
  On Error GoTo ErrorHandler
    Dim strSource As String
    strSource = cbxSource.List(cbxSource.ListIndex)
   
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
        
    Dim strSql As String
    If strSource <> "User Defined" And strSource <> "Average" Then
    
        strSql = "SELECT DISTINCT Reference_Sources.Title, " & _
            " Reference_Sources.Year, Reference_Sources.Author, " & _
            " Reference_Sources.Publication_Street, " & _
            " Reference_Sources.Publication_City, " & _
            " Reference_Sources.Publication_State, " & _
            " Reference_Sources.Reference_Number, " & _
            " Reference_Sources.Prepared_By, " & _
            " UnitTypes.UnitType_Desc, Component_Costs.Unit, Component_Costs.Cost " & _
            " FROM (Component_Costs INNER JOIN Reference_Sources ON " & _
            " Component_Costs.Source_ID = Reference_Sources.Reference_Source_ID) " & _
            " INNER JOIN UnitTypes ON Component_Costs.Unit = UnitTypes.UnitType_Code " & _
            " WHERE Reference_Sources.Author ='" & cbxSource.List(cbxSource.ListIndex) & "'" & _
            " AND Component_Costs.Locale = '" & Trim(cbxLocation.List(cbxLocation.ListIndex)) & "'" & _
            " AND Component_Costs.Component_ID =" & cbxComponent.ItemData(cbxComponent.ListIndex) & _
            " AND Component_Costs.Use_Flag <> 0 " & _
            " AND Reference_Sources.Year = " & cbxYear.Text 'cbxYear.List(cbxYear.ListIndex)
        
        pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
           
        pRs.MoveFirst
        If Not pRs.EOF Then
            If Not IsNull(pRs("Cost")) Then
                txtCost.Text = pRs("Cost")
            Else
                txtCost.Text = ""
            End If
            
            If Not IsNull(pRs("UnitType_Desc")) Then
                Set_Unit (pRs("UnitType_Desc"))
            Else
                Set_Unit ("Unknown")
            End If
        End If
        txtSourceDetails = pRs("Author") & " . " & _
                pRs("Year") & ". " & pRs("Title") & ". " & _
                pRs("Publication_Street") & ". " & _
                pRs("Publication_City") & ". " & _
                pRs("Publication_State") & ". " & _
                pRs("Reference_Number")
        pRs.Close
        'txtCost.Enabled = False
    ElseIf strSource = "Average" Then
        'Find the average...
        strSql = "SELECT DISTINCT UnitTypes.UnitType_Desc,  Average_Costs.AvgOfCost" & _
                " FROM Average_Costs INNER JOIN UnitTypes ON Average_Costs.Unit = UnitTypes.UnitType_Code " & _
                " WHERE Component_ID = " & cbxComponent.ItemData(cbxComponent.ListIndex) & _
                " AND Average_Costs.Year = " & cbxYear.Text
        pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
           
        pRs.MoveFirst
        If Not pRs.EOF Then
            If Not IsNull(pRs("AvgOfCost")) Then
                txtCost.Text = pRs("AvgOfCost")
            Else
                txtCost.Text = ""
            End If
            
            If Not IsNull(pRs("UnitType_Desc")) Then
                Set_Unit (pRs("UnitType_Desc"))
            Else
                Set_Unit ("Unknown")
            End If
        End If
        txtSourceDetails = "Average cost - Combination of many sources"
        pRs.Close
    Else
        Set_Unit ("Unknown")
        txtSourceDetails = ""
        txtCost.Enabled = True
        txtCost.Text = ""
    End If
  Exit Sub
ErrorHandler:
  HandleError True, "cbxYear_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub cbxUnit_Click()
    txtCostUnits.Enabled = False
    If cbxUnit.List(cbxUnit.ListIndex) = "Per Unit" Then
        txtCostUnits.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo ErrorHandler
    
    Dim costAdj As Double
    costAdj = 1
    
    Dim curYear As Integer
    
    If cbxUnit.List(cbxUnit.ListIndex) = "Unknown" Then
        MsgBox "Unit cannot be 'Unknown'. Please select an appropriate unit from the list", vbExclamation
        Exit Sub
    ElseIf cbxUnit.List(cbxUnit.ListIndex) = "Per Unit" Then
        If Trim(txtCostUnits.Text) = "" Or Not IsNumeric(txtCostUnits.Text) Then
            MsgBox "Please enter a valid number for Number of Unit (Cost Tab)", vbExclamation
            Exit Sub
        End If
    End If
    Dim strComponent As String
    strComponent = cbxComponent.List(cbxComponent.ListIndex)
    If strComponent = "User Defined" Then
        If Trim(txtUserComponent.Text) = "" Then
            MsgBox "Please enter a name for the 'User Defined' component", vbExclamation
            Exit Sub
        Else
            strComponent = Trim(txtUserComponent.Text)
        End If
    End If
    If Trim(txtCost.Text) = "" Or Not IsNumeric(txtCost.Text) Then
        MsgBox "Please enter a valid number for Unit Cost", vbExclamation
        Exit Sub
    End If
    curYear = CInt(cbxYear.List(cbxYear.ListIndex))
    
    If UCase(cbxUnit.List(cbxUnit.ListIndex)) <> "PERCENTAGE" Then
        If costAdjDict.Exists(curYear) Then costAdj = costAdjDict.Item(curYear)
    End If
    
    Dim pCostVolType As Integer
    Dim pCostNumUnits As Integer
    
    pCostVolType = COST_VOLUME_TYPE_TOTAL
'    If optVolMedia.value Then
'        pCostVolType = COST_VOLUME_TYPE_MEDIA
'    ElseIf optVolUnderDrain.value Then
'        pCostVolType = COST_VOLUME_TYPE_UNDERDRAIN
'    End If
    
    pCostNumUnits = 1
    If cbxUnit.List(cbxUnit.ListIndex) = "Per Unit" Then
        pCostNumUnits = CInt(txtCostUnits.Text)
    End If
    
    If cbxComponent.ListCount = 0 Then Exit Sub
    
    Dim lstItem As ListItem
    Set lstItem = lstComponents.ListItems.add(, , strComponent) 'cbxComponent.List(cbxComponent.ListIndex))
    lstItem.ListSubItems.add , , cbxComponent.ItemData(cbxComponent.ListIndex)
    lstItem.ListSubItems.add , , cbxLocation.List(cbxLocation.ListIndex)
    lstItem.ListSubItems.add , , cbxSource.List(cbxSource.ListIndex)
    lstItem.ListSubItems.add , , cbxYear.List(cbxYear.ListIndex)
    lstItem.ListSubItems.add , , cbxUnit.List(cbxUnit.ListIndex)
    lstItem.ListSubItems.add , , pCostVolType
    lstItem.ListSubItems.add , , pCostNumUnits
    lstItem.ListSubItems.add , , txtCost.Text
    lstItem.ListSubItems.add , , CStr(FormatNumber(CDbl(txtCost.Text) * costAdj, 2))

  Exit Sub
ErrorHandler:
  HandleError True, "cmdAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

'edit cost factor Ying Cao Dec 8, 2008
Private Sub cmdEdit_Click()
  On Error GoTo ErrorHandler
    Dim pIndex As Integer
    Dim pCompIndex As Integer
    Dim pComponent As String
    Dim pComponent_Id As String
    Dim curYear As Integer
    Dim costAdj As Double
    costAdj = 1
    Dim strComponent As String
    Dim pCostVolType As Integer
    Dim pCostNumUnits As Integer
    
    Dim hasSelected As Boolean
    
    
    If cmdEdit.Caption = "Save" Then    'save the component values to listview
        For pIndex = lstComponents.ListItems.Count To 1 Step -1
            If lstComponents.ListItems.Item(pIndex).Selected Then
                'validate user input
                If cbxUnit.List(cbxUnit.ListIndex) = "Unknown" Then
                    MsgBox "Unit cannot be 'Unknown'. Please select an appropriate unit from the list", vbExclamation
                    Exit Sub
                ElseIf cbxUnit.List(cbxUnit.ListIndex) = "Per Unit" Then
                    If Trim(txtCostUnits.Text) = "" Or Not IsNumeric(txtCostUnits.Text) Then
                        MsgBox "Please enter a valid number for Number of Unit (Cost Tab)", vbExclamation
                        Exit Sub
                    End If
                End If
                
                strComponent = cbxComponent.List(cbxComponent.ListIndex)
                If strComponent = "User Defined" Then
                    If Trim(txtUserComponent.Text) = "" Then
                        MsgBox "Please enter a name for the 'User Defined' component", vbExclamation
                        Exit Sub
                    Else
                        strComponent = Trim(txtUserComponent.Text)
                    End If
                End If
                If Trim(txtCost.Text) = "" Or Not IsNumeric(txtCost.Text) Then
                    MsgBox "Please enter a valid number for Unit Cost", vbExclamation
                    Exit Sub
                End If
    
                If txtUserComponent.Text <> "" Then
                    lstComponents.ListItems.Item(pIndex).ListSubItems(1).Text = txtUserComponent.Text
                Else
                    lstComponents.ListItems.Item(pIndex).ListSubItems(1).Text = cbxComponent.Text
                End If
                
                If cbxComponent.Text <> "User Defined" Then
                    lstComponents.ListItems.Item(pIndex) = cbxComponent.Text
                Else
                    lstComponents.ListItems.Item(pIndex) = txtUserComponent.Text
                End If
                lstComponents.ListItems.Item(pIndex).ListSubItems(2).Text = cbxLocation.Text
                lstComponents.ListItems.Item(pIndex).ListSubItems(3).Text = cbxSource.Text
                lstComponents.ListItems.Item(pIndex).ListSubItems(4).Text = cbxYear.Text
                lstComponents.ListItems.Item(pIndex).ListSubItems(5).Text = cbxUnit.Text
                lstComponents.ListItems.Item(pIndex).ListSubItems(6).Text = txtCostUnits.Text
                lstComponents.ListItems.Item(pIndex).ListSubItems(6).Text = CStr(pCostVolType)
                lstComponents.ListItems.Item(pIndex).ListSubItems(8).Text = txtCost.Text

                curYear = CInt(cbxYear.List(cbxYear.ListIndex))
                If UCase(cbxUnit.List(cbxUnit.ListIndex)) <> "PERCENTAGE" Then
                    If costAdjDict.Exists(curYear) Then
                        costAdj = costAdjDict.Item(curYear)
                    End If
                End If
                lstComponents.ListItems.Item(pIndex).ListSubItems(9).Text = CStr(FormatNumber(CDbl(txtCost.Text) * costAdj, 2))
                cmdEdit.Caption = "Edit"
                Exit For
            End If
        Next
        
    ElseIf cmdEdit.Caption = "Edit" Then    'populate the component values
        hasSelected = False
        For pIndex = lstComponents.ListItems.Count To 1 Step -1
            If lstComponents.ListItems.Item(pIndex).Selected Then
                pComponent = lstComponents.ListItems.Item(pIndex)
                'populate values to fields
                
'                For pCompIndex = cbxComponent.ListCount To 1 Step -1
'                    If cbxComponent.List(pCompIndex - 1) = lstComponents.ListItems.Item(pIndex) Then
'                        cbxComponent.Text = lstComponents.ListItems.Item(pIndex)
'                        txtUserComponent.Text = ""
'                        txtUserComponent.Enabled = False
'                        Exit For
'                    End If
'                Next
                
                pCompIndex = ModuleUtility.SetComboItemIndex(cbxComponent, lstComponents.ListItems.Item(pIndex))
                
                If pCompIndex = 0 Then  'no match found
                    cbxComponent.ListIndex = cbxComponent.ListCount - 1     'last item: User Defined
                    txtUserComponent.Text = lstComponents.ListItems.Item(pIndex)
                    txtUserComponent.Enabled = True
                Else
                    txtUserComponent.Text = ""
                    txtUserComponent.Enabled = False
                End If
                
                pCompIndex = ModuleUtility.SetComboItemIndex(cbxLocation, lstComponents.ListItems.Item(pIndex).ListSubItems(2).Text)
                pCompIndex = ModuleUtility.SetComboItemIndex(cbxSource, lstComponents.ListItems.Item(pIndex).ListSubItems(3).Text)
                pCompIndex = ModuleUtility.SetComboItemIndex(cbxYear, lstComponents.ListItems.Item(pIndex).ListSubItems(4).Text)
                pCompIndex = ModuleUtility.SetComboItemIndex(cbxUnit, lstComponents.ListItems.Item(pIndex).ListSubItems(5).Text)
                txtCostUnits.Text = lstComponents.ListItems.Item(pIndex).ListSubItems(6).Text
                txtCost.Text = lstComponents.ListItems.Item(pIndex).ListSubItems(8).Text

                hasSelected = True
            End If
        Next
        If hasSelected = False Then
            MsgBox "Please highlight a component before editing."
        Else
            cmdEdit.Caption = "Save"
        End If
        
    End If
    
  Exit Sub
ErrorHandler:
  HandleError True, "cmdEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub cmdOk_Click()
    If ValidateInputs Then
        bContinue = True
        
        Set gBufferStripDetailDict = CreateObject("Scripting.Dictionary")
        
        gBufferStripDetailDict.Item("ID") = CInt(txtVFSID.Text)
        gBufferStripDetailDict.Item("Name") = txtName.Text
        
        Dim oRs As ADODB.Recordset
        Set oRs = DataGridSedDec.DataSource
        
        oRs.MoveFirst
        If Not oRs.EOF Then
            Do Until oRs.EOF
                gBufferStripDetailDict.Item("SedFrac" & oRs.Fields(0).value) = oRs.Fields(2).value
                gBufferStripDetailDict.Item("SedDec" & oRs.Fields(0).value) = oRs.Fields(3).value
                gBufferStripDetailDict.Item("SedCorr" & oRs.Fields(0).value) = oRs.Fields(4).value
                gBufferStripDetailDict.Item("WatDec" & oRs.Fields(0).value) = oRs.Fields(5).value
                gBufferStripDetailDict.Item("WatCorr" & oRs.Fields(0).value) = oRs.Fields(6).value
                oRs.MoveNext
            Loop
        End If
        
        Dim pControl
        For Each pControl In Controls
            If ((TypeOf pControl Is TextBox) And (pControl.Enabled)) Then
               If pControl.name <> "txtName" And pControl.name <> "txtVFSID" Then
                    gBufferStripDetailDict.Item(pControl.name) = Trim(pControl.Text)
               End If
            End If
        Next
'        If optKPG0.value = True Then
'            gBufferStripDetailDict.Item("KPG") = 0
'        Else
'            gBufferStripDetailDict.Item("KPG") = 1
'        End If
        
        Dim oRsSeg As ADODB.Recordset
        Set oRsSeg = DataGridSegments.DataSource
        
        If Not oRsSeg.EOF Then
            oRsSeg.MoveFirst
            Dim rowCnt As Integer
            rowCnt = 1
            Do Until oRsSeg.EOF
                gBufferStripDetailDict.Item("SX" & rowCnt) = oRsSeg.Fields(0).value
                gBufferStripDetailDict.Item("RNA" & rowCnt) = oRsSeg.Fields(1).value
                gBufferStripDetailDict.Item("SOA" & rowCnt) = oRsSeg.Fields(2).value
                rowCnt = rowCnt + 1
                oRsSeg.MoveNext
            Loop
        End If
        
        Dim costComps As String
        Dim costCompIds As String
        Dim costLocs As String
        Dim costSrcs As String
        Dim costYears As String
        Dim costUnits As String
        Dim costUnitCosts As String
        Dim costAdjUnitCosts As String
        Dim costVolTypes As String
        Dim costNumUnits As String
        
        costComps = ""
        costCompIds = ""
        costLocs = ""
        costSrcs = ""
        costYears = ""
        costUnits = ""
        costUnitCosts = ""
        costAdjUnitCosts = ""
        costVolTypes = ""
        costNumUnits = ""
        
        
        Dim pIndex As Integer
        
        If lstComponents.ListItems.Count > 0 Then
             costComps = lstComponents.ListItems.Item(1)
             costCompIds = lstComponents.ListItems.Item(1).SubItems(1)
             costLocs = lstComponents.ListItems.Item(1).SubItems(2)
             costSrcs = lstComponents.ListItems.Item(1).SubItems(3)
             costYears = lstComponents.ListItems.Item(1).SubItems(4)
             costUnits = lstComponents.ListItems.Item(1).SubItems(5)
             costVolTypes = lstComponents.ListItems.Item(1).SubItems(6)
             costNumUnits = lstComponents.ListItems.Item(1).SubItems(7)
             costUnitCosts = lstComponents.ListItems.Item(1).SubItems(8)
             costAdjUnitCosts = lstComponents.ListItems.Item(1).SubItems(9)
        
            For pIndex = 2 To lstComponents.ListItems.Count
                 costComps = costComps & ";" & lstComponents.ListItems.Item(pIndex)
                 costCompIds = costCompIds & ";" & lstComponents.ListItems.Item(pIndex).SubItems(1)
                 costLocs = costLocs & ";" & lstComponents.ListItems.Item(pIndex).SubItems(2)
                 costSrcs = costSrcs & ";" & lstComponents.ListItems.Item(pIndex).SubItems(3)
                 costYears = costYears & ";" & lstComponents.ListItems.Item(pIndex).SubItems(4)
                 costUnits = costUnits & ";" & lstComponents.ListItems.Item(pIndex).SubItems(5)
                 costVolTypes = costVolTypes & ";" & lstComponents.ListItems.Item(pIndex).SubItems(6)
                 costNumUnits = costNumUnits & ";" & lstComponents.ListItems.Item(pIndex).SubItems(7)
                 costUnitCosts = costUnitCosts & ";" & lstComponents.ListItems.Item(pIndex).SubItems(8)
                 costAdjUnitCosts = costAdjUnitCosts & ";" & lstComponents.ListItems.Item(pIndex).SubItems(9)
            Next
            
            gBufferStripDetailDict.add "CostComponents", costComps
            gBufferStripDetailDict.add "CostComponentIds", costCompIds
            gBufferStripDetailDict.add "CostLocations", costLocs
            gBufferStripDetailDict.add "CostSources", costSrcs
            gBufferStripDetailDict.add "CostYears", costYears
            gBufferStripDetailDict.add "CostUnits", costUnits
            gBufferStripDetailDict.add "CostVolTypes", costVolTypes
            gBufferStripDetailDict.add "CostNumUnits", costNumUnits
            gBufferStripDetailDict.add "CostUnitCosts", costUnitCosts
            gBufferStripDetailDict.add "CostAdjUnitCosts", costAdjUnitCosts
        End If
        
    Else
        bContinue = False
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    '** close the form
    bContinue = False
    Unload Me
End Sub

Private Function ValidateInputs() As Boolean
On Error GoTo ErrorHandler
    
    Dim pMessageStr As String
    pMessageStr = ""
    
    If (Trim(txtName.Text) = "") Then
        pMessageStr = "Buffer strip name (can not be empty)."
    End If
    
    Dim pControl
    For Each pControl In Controls
        If ((TypeOf pControl Is TextBox) And (pControl.Enabled)) Then
           If Not (pControl.name = "txtName" Or pControl.name = "txtUserComponent" Or pControl.name = "txtCost" Or pControl.name = "txtCostUnits") Then
                If Trim(pControl.Text) = "" Or Not (IsNumeric(Trim(pControl.Text))) Then
                    pMessageStr = pMessageStr & pControl.name & vbNewLine
                End If
           End If
        End If
    Next
    
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSedDec.DataSource
    
    If Not oRs.EOF Then
        oRs.MoveFirst
        Do Until oRs.EOF
            If Not IsNumeric(oRs.Fields(2).value) Then _
                pMessageStr = pMessageStr & "Sediment Fraction " & oRs.Fields(1).value & vbNewLine
            If Not IsNumeric(oRs.Fields(3).value) Then _
                pMessageStr = pMessageStr & "Sediment dDecay factor for " & oRs.Fields(1).value & vbNewLine
            If Not IsNumeric(oRs.Fields(4).value) Then _
                pMessageStr = pMessageStr & "Sediment temperature Correction factor for " & oRs.Fields(1).value & vbNewLine
            If Not IsNumeric(oRs.Fields(5).value) Then _
                pMessageStr = pMessageStr & "Dissolved decay factor for " & oRs.Fields(1).value & vbNewLine
            If Not IsNumeric(oRs.Fields(6).value) Then _
                pMessageStr = pMessageStr & "Dissolved temperature Correction factor for " & oRs.Fields(1).value & vbNewLine
            oRs.MoveNext
        Loop
    End If
    
    Dim oRsSeg As ADODB.Recordset
    Set oRsSeg = DataGridSegments.DataSource
    
    If CInt(NPROP.Text) <> oRsSeg.RecordCount Then
        pMessageStr = pMessageStr & "The NPROP and number of rows in grass segments grid do not match "
    End If
    
    
    Dim rowCnt As Integer
    rowCnt = 1
    
    If Not oRsSeg.EOF Then
        oRsSeg.MoveFirst
        Do Until oRsSeg.EOF
            If Not IsNumeric(oRsSeg.Fields(0).value) Then _
                pMessageStr = pMessageStr & "X-Distance (segment grid) for row number  " & rowCnt & vbNewLine
            If Not IsNumeric(oRsSeg.Fields(1).value) Then _
                pMessageStr = pMessageStr & "Mannings N (segment grid) for row number  " & rowCnt & vbNewLine
            If Not IsNumeric(oRsSeg.Fields(2).value) Then _
                pMessageStr = pMessageStr & "Segment Slope (segment grid) for row number  " & rowCnt & vbNewLine
            oRsSeg.MoveNext
            rowCnt = rowCnt + 1
        Loop
    End If
    
    If pMessageStr = "" Then
        ValidateInputs = True
    Else
        ValidateInputs = False
        MsgBox "Enter valid inputs for following parameters" & vbNewLine & pMessageStr
    End If
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error in Validating the inputs " & Err.description
CleanUp:
    
End Function


Private Sub cmdRemove_Click()
  On Error GoTo ErrorHandler

    Dim pIndex As Integer
    Dim pComponent As String
    Dim pComponent_Id As Integer
    
    For pIndex = lstComponents.ListItems.Count To 1 Step -1
        If lstComponents.ListItems.Item(pIndex).Selected Then
            pComponent = lstComponents.ListItems.Item(pIndex)
            'pComponent_Id = lstComponents.ListItems.Item(pIndex).ListSubItems(1)
            lstComponents.ListItems.Remove (pIndex)
            Exit For
        End If
    Next
  Exit Sub
ErrorHandler:
  HandleError True, "cmdRemove_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

'Private Sub Form_Load()
'On Error GoTo ShowError
'
'
'    Exit Sub
'ShowError:
'    MsgBox "Error loading VFSParameter Form: " & err.description
'End Sub

Private Sub cmdUpdateDG_Click()
    Dim iNProp As Integer
    If IsNumeric(Trim(NPROP.Text)) Then
        iNProp = CInt(Trim(NPROP.Text))
    Else
        MsgBox "Enter a valid (integer) number for NPROP"
        Exit Sub
    End If
    Dim oRs As ADODB.Recordset
    Set oRs = DataGridSegments.DataSource
    Dim i As Integer, recCnt As Integer
    recCnt = oRs.RecordCount
    If recCnt > iNProp Then
        oRs.MoveLast
        For i = recCnt To iNProp + 1 Step -1
          oRs.Delete
          oRs.MoveLast
        Next
    ElseIf recCnt < iNProp Then
        oRs.MoveLast
        For i = recCnt To iNProp - 1
          oRs.AddNew
          oRs.MoveLast
        Next
    End If
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If SSTab1.TabVisible(gBMPDefTab) Then SSTab1.Tab = gBMPDefTab
    Set costAdjDict = New Scripting.Dictionary
    Call SetCostAdjDict(costAdjDict)
    
End Sub


Public Sub InitCostFromDB(Optional componentsToExclude As String)
On Error GoTo ErrorHandler
    Dim BMPType As String
    BMPType = "Buffer Strip"
    Dim ConnStr As String
    'ConnStr = "D:\SUSTAIN\CostDB\BMPCosts.mdb"
    If Trim(gCostDBpath) = "" Then Exit Sub
    ConnStr = Trim(gCostDBpath)
    
    Set pAdoConn = New ADODB.Connection
    pAdoConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnStr & ";"
    
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
           
    Dim strSqlComp As String
    strSqlComp = " SELECT Components.Components_ID, Components.Components_TXT " & _
                " FROM (BMPTypes INNER JOIN BMP_Components ON BMPTypes.BMPType_ID = BMP_Components.BMPType_ID) INNER JOIN Components ON BMP_Components.Component_ID = Components.Components_ID " & _
                " WHERE BMPTypes.BMPType_Code = '" & BMPType & "'"

''    strSqlComp = " SELECT DISTINCT Component_ID, Components " & _
''                " FROM Cost_Unit_Check " & _
''                " WHERE [BMP Type] = '" & BMPType & "'"
    
    If componentsToExclude <> "" Then
        strSqlComp = strSqlComp & " AND Components_TXT NOT IN " & componentsToExclude
    End If
    
    pRs.Open strSqlComp, pAdoConn, adOpenDynamic, adLockOptimistic
       
    cbxComponent.Clear
    
    If Not pRs.EOF Then
        pRs.MoveFirst
        Do Until pRs.EOF
            cbxComponent.AddItem Trim(pRs("Components_TXT"))
            cbxComponent.ItemData(cbxComponent.NewIndex) = Trim(pRs("Components_ID"))
            pRs.MoveNext
        Loop
    End If
    cbxComponent.AddItem "Land Cost"
    cbxComponent.AddItem "User Defined"
    
    cbxComponent.ListIndex = 0
    pRs.Close
    
    If cbxUnit.ListCount = 0 Then Populate_Cost_Units

    Exit Sub
    
ErrorHandler:
    MsgBox "Error in Initializing Cost Data:" & Err.description
End Sub

Private Sub Set_Unit(curUnit As String)
On Error GoTo ErrorHandler
    Dim i As Integer
    For i = 0 To cbxUnit.ListCount
        If UCase(cbxUnit.List(i)) = UCase(curUnit) Then
            cbxUnit.ListIndex = i
            Exit For
        End If
    Next
    Exit Sub
ErrorHandler:
  HandleError True, "Set_Unit " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub Populate_Cost_Units()
On Error GoTo ErrorHandler
   
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
        
    Dim strSqlUnit As String
    strSqlUnit = "SELECT UnitType_Desc " & _
        " FROM UnitTypes"

    pRs.Open strSqlUnit, pAdoConn, adOpenDynamic, adLockOptimistic
       
    cbxUnit.Clear
    pRs.MoveFirst
    If Not pRs.EOF Then
        Do Until pRs.EOF
            cbxUnit.AddItem pRs("UnitType_Desc")
            pRs.MoveNext
        Loop
    End If
    cbxUnit.ListIndex = 0
    pRs.Close

  Exit Sub
ErrorHandler:
  HandleError True, "Populate_Cost_Units " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

'call this function to update component list
Public Sub Update_Component_List(pBufferStripDetailDict As Scripting.Dictionary)
On Error GoTo ErrorHandler
    Dim costComps
    Dim costCompIds
    Dim costLocs
    Dim costSrcs
    Dim costYears
    Dim costUnits
    Dim costUnitCosts
    Dim costAdjUnitCosts
    Dim costVolTypes
    Dim costNumUnits

    Dim lstItem As ListItem
    Dim pIndex As Integer
    
    'first clear the list
    lstComponents.ListItems.Clear
    
    If pBufferStripDetailDict Is Nothing Then Exit Sub

    'populate items in lstComponents
    If pBufferStripDetailDict.Exists("CostComponents") Then

        costComps = Split(pBufferStripDetailDict.Item("CostComponents"), ";", , vbTextCompare)
        costCompIds = Split(pBufferStripDetailDict.Item("CostComponentIds"), ";", , vbTextCompare)
        costLocs = Split(pBufferStripDetailDict.Item("CostLocations"), ";", , vbTextCompare)
        costSrcs = Split(pBufferStripDetailDict.Item("CostSources"), ";", , vbTextCompare)
        costYears = Split(pBufferStripDetailDict.Item("CostYears"), ";", , vbTextCompare)
        costUnits = Split(pBufferStripDetailDict.Item("CostUnits"), ";", , vbTextCompare)
        costVolTypes = Split(pBufferStripDetailDict.Item("CostVolTypes"), ";", , vbTextCompare)
        costNumUnits = Split(pBufferStripDetailDict.Item("CostNumUnits"), ";", , vbTextCompare)
        costUnitCosts = Split(pBufferStripDetailDict.Item("CostUnitCosts"), ";", , vbTextCompare)
        costAdjUnitCosts = Split(pBufferStripDetailDict.Item("CostAdjUnitCosts"), ";")

        For pIndex = 0 To UBound(costComps)
            Set lstItem = lstComponents.ListItems.add(, , costComps(pIndex))
            lstItem.ListSubItems.add , , costCompIds(pIndex)
            lstItem.ListSubItems.add , , costLocs(pIndex)
            lstItem.ListSubItems.add , , costSrcs(pIndex)
            lstItem.ListSubItems.add , , costYears(pIndex)
            lstItem.ListSubItems.add , , costUnits(pIndex)
            lstItem.ListSubItems.add , , costVolTypes(pIndex)
            lstItem.ListSubItems.add , , costNumUnits(pIndex)
            lstItem.ListSubItems.add , , costUnitCosts(pIndex)
            lstItem.ListSubItems.add , , costAdjUnitCosts(pIndex)
            'Remove_Component CStr(costComps(pIndex))
        Next
    End If

    Exit Sub
ErrorHandler:
  HandleError True, "Update_Component_List " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


