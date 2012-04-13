VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmSWMMOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Simulation Options"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "FrmSWMMOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   45
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   44
      Top             =   4920
      Width           =   975
   End
   Begin TabDlg.SSTab TabOptions 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSWMMOptions.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dates"
      TabPicture(1)   =   "FrmSWMMOptions.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAntecedentDryDays"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "sweepingEndDate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "sweepingStartDate"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "analysisEndTime"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "analysisEndDate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "reportingStartTime"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "reportingStartDate"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "analysisStartTime"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "analysisStartDate"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label14"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label13"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label12"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label10"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label9"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label8"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label7"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Time Steps"
      TabPicture(2)   =   "FrmSWMMOptions.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDryWeatherRunoff"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dtReporting"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "UpDown4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtReporting"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "dtRouting"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "UpDown3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtRouting"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dtWetWeatherRunoff"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "UpDown2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtWetWeatherRunoff"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "UpDown1"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "dtDryWeatherRunoff"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label21"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label20"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label19"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label18"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label17"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label16"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label15"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Files"
      TabPicture(3)   =   "FrmSWMMOptions.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdBrowseSFile"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtSWMMInputFile"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtPost"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdBrowsePost"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "CommonDialog"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label26"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label23"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.CommandButton cmdBrowseSFile 
         Caption         =   "..."
         Height          =   375
         Left            =   -69480
         TabIndex        =   56
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtSWMMInputFile 
         Height          =   375
         Left            =   -73560
         TabIndex        =   55
         Top             =   720
         Width           =   3855
      End
      Begin VB.Frame Frame 
         Caption         =   "Pre-Developed Landuse Scenario"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   49
         Top             =   1800
         Width           =   6015
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   375
            Left            =   5400
            TabIndex        =   59
            Top             =   885
            Width           =   495
         End
         Begin VB.TextBox txtOutflowFile 
            Height          =   375
            Left            =   1320
            TabIndex        =   58
            Top             =   885
            Width           =   3855
         End
         Begin VB.CommandButton cmdPreDevBrowse 
            Caption         =   "..."
            Height          =   375
            Left            =   5400
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtPreDevInputFile 
            Height          =   375
            Left            =   1320
            TabIndex        =   51
            Top             =   360
            Width           =   3855
         End
         Begin VB.ComboBox cmbPredevLanduse 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label22 
            Caption         =   "Output File:"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "Input File:"
            Height          =   495
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label24 
            Caption         =   "Predeveloped Landuse:"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   1335
         End
      End
      Begin VB.TextBox txtPost 
         Height          =   375
         Left            =   -73560
         TabIndex        =   47
         Top             =   1245
         Width           =   3855
      End
      Begin VB.CommandButton cmdBrowsePost 
         Caption         =   "..."
         Height          =   375
         Left            =   -69480
         TabIndex        =   46
         Top             =   1245
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   -69600
         Top             =   3960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDryWeatherRunoff 
         Height          =   375
         Left            =   -72600
         TabIndex        =   42
         Text            =   "0"
         Top             =   1260
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtReporting 
         Height          =   375
         Left            =   -71040
         TabIndex        =   39
         Top             =   3060
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         OLEDropMode     =   1
         CustomFormat    =   "HH:mm:ss"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494.0416666667
      End
      Begin MSComCtl2.UpDown UpDown4 
         Height          =   375
         Left            =   -72000
         TabIndex        =   38
         Top             =   3060
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtReporting"
         BuddyDispid     =   196625
         OrigLeft        =   2520
         OrigTop         =   3840
         OrigRight       =   2760
         OrigBottom      =   4215
         Max             =   366
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtReporting 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72600
         TabIndex        =   37
         Text            =   "0"
         Top             =   3060
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtRouting 
         Height          =   375
         Left            =   -71040
         TabIndex        =   35
         Top             =   2460
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "HH:mm:ss"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494.0006944444
      End
      Begin MSComCtl2.UpDown UpDown3 
         Height          =   375
         Left            =   -72000
         TabIndex        =   34
         Top             =   2460
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtRouting"
         BuddyDispid     =   196626
         OrigLeft        =   2520
         OrigTop         =   3000
         OrigRight       =   2760
         OrigBottom      =   3375
         Max             =   366
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtRouting 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72600
         TabIndex        =   33
         Text            =   "0"
         Top             =   2460
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtWetWeatherRunoff 
         Height          =   375
         Left            =   -71040
         TabIndex        =   31
         Top             =   1860
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "HH:mm:ss"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494.0104166667
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   375
         Left            =   -72000
         TabIndex        =   30
         Top             =   1860
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtWetWeatherRunoff"
         BuddyDispid     =   196627
         OrigLeft        =   2520
         OrigTop         =   2160
         OrigRight       =   2760
         OrigBottom      =   2535
         Max             =   366
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtWetWeatherRunoff 
         Height          =   375
         Left            =   -72600
         TabIndex        =   29
         Text            =   "0"
         Top             =   1860
         Width           =   615
      End
      Begin VB.TextBox txtAntecedentDryDays 
         Height          =   375
         Left            =   -72480
         TabIndex        =   24
         Text            =   "5"
         Top             =   3960
         Width           =   735
      End
      Begin MSComCtl2.DTPicker sweepingEndDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   22
         Top             =   3330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "MM/dd"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   35796
      End
      Begin MSComCtl2.DTPicker sweepingStartDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   20
         Top             =   2715
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "MM/dd"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   35796
      End
      Begin MSComCtl2.DTPicker analysisEndTime 
         Height          =   375
         Left            =   -70440
         TabIndex        =   18
         Top             =   2085
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "HH:mm"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494.5
      End
      Begin MSComCtl2.DTPicker analysisEndDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   17
         Top             =   2085
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   35796
      End
      Begin MSComCtl2.DTPicker reportingStartTime 
         Height          =   375
         Left            =   -70440
         TabIndex        =   15
         Top             =   1470
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker reportingStartDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   14
         Top             =   1470
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   35796
      End
      Begin MSComCtl2.DTPicker analysisStartTime 
         Height          =   375
         Left            =   -70440
         TabIndex        =   12
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "HH:mm"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker analysisStartDate 
         Height          =   375
         Left            =   -72480
         TabIndex        =   11
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   35796
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   3975
         Begin VB.Label Label4 
            Caption         =   "Kinematic Wave"
            Height          =   375
            Left            =   1920
            TabIndex        =   7
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Routing Method"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1020
            Width           =   1995
         End
         Begin VB.Label Label6 
            Caption         =   "Green Ampt"
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Infiltration Model"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label2 
            Caption         =   "CFS"
            Height          =   375
            Left            =   1920
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Flow Units"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   1995
         End
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   -72000
         TabIndex        =   41
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         BuddyControl    =   "txtDryWeatherRunoff"
         BuddyDispid     =   196624
         OrigLeft        =   3000
         OrigTop         =   1260
         OrigRight       =   3240
         OrigBottom      =   1635
         Max             =   366
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtDryWeatherRunoff 
         Height          =   375
         Left            =   -71040
         TabIndex        =   43
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   36494.0416666667
         MaxDate         =   40543
         MinDate         =   2
      End
      Begin VB.Label Label26 
         Caption         =   "Input File:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   57
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Output File:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Runoff: Dry Weather"
         Height          =   255
         Left            =   -74640
         TabIndex        =   40
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Reporting"
         Height          =   255
         Left            =   -74640
         TabIndex        =   36
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Routing"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Runoff: Wet Weather"
         Height          =   255
         Left            =   -74640
         TabIndex        =   28
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Runoff: Dry Weather"
         Height          =   255
         Left            =   -74640
         TabIndex        =   27
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Time (HH:MM:SS)"
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
         Left            =   -71160
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Days"
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
         Left            =   -72600
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Antecedent Dry Days"
         Height          =   375
         Left            =   -74640
         TabIndex        =   23
         Top             =   3960
         Width           =   1800
      End
      Begin VB.Label Label13 
         Caption         =   "End Sweeping On"
         Height          =   375
         Left            =   -74640
         TabIndex        =   21
         Top             =   3330
         Width           =   1800
      End
      Begin VB.Label Label12 
         Caption         =   "Start Sweeping On"
         Height          =   375
         Left            =   -74640
         TabIndex        =   19
         Top             =   2715
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "End Analysis On"
         Height          =   375
         Left            =   -74640
         TabIndex        =   16
         Top             =   2085
         Width           =   1800
      End
      Begin VB.Label Label10 
         Caption         =   "Start Reporting On"
         Height          =   375
         Left            =   -74640
         TabIndex        =   13
         Top             =   1470
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Start Analysis On"
         Height          =   375
         Left            =   -74640
         TabIndex        =   10
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label8 
         Caption         =   "Time (HH:MM)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -70440
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Date (MM/DD/YYYY)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72600
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmSWMMOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdBrowse_Click()
    CommonDialog.Filter = "Text (*.txt)|*.txt"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtOutflowFile.Text = CommonDialog.FileName
End Sub

Private Sub cmdBrowsePost_Click()
    CommonDialog.Filter = "Text (*.txt)|*.txt"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtPost.Text = CommonDialog.FileName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If (Trim(txtAntecedentDryDays.Text) = "") Then
        MsgBox "Enter Antecedent Dry Days to continue.", vbExclamation
        FrmSWMMOptions.TabOptions.Tab = 1
        Exit Sub
    End If
    
    If (Trim(txtOutflowFile.Text) = "") Then
        MsgBox "Please Select the outflow file to continue.", vbExclamation
        FrmSWMMOptions.TabOptions.Tab = 3
        Exit Sub
    End If
             
    'All values are entered, save it a dictionary, and call a routine
    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = CreateObject("Scripting.Dictionary")
    
    
    '** Get the dry step, wet step, routing step, reporting step
    Dim pDryStep
    pDryStep = CInt(txtDryWeatherRunoff.Text) * 24 + Hour(dtDryWeatherRunoff.value)
    pDryStep = pDryStep & ":" & Minute(dtDryWeatherRunoff.value) & ":" & Second(dtDryWeatherRunoff.value)
        
    Dim pWetStep
    pWetStep = CInt(txtWetWeatherRunoff.Text) * 24 + Hour(dtWetWeatherRunoff.value)
    pWetStep = pWetStep & ":" & Minute(dtWetWeatherRunoff.value) & ":" & Second(dtWetWeatherRunoff.value)
      
    Dim pRoutingStep
    pRoutingStep = Hour(dtRouting.value) & ":" & Minute(dtRouting.value) & ":" & Second(dtRouting.value)
    
    Dim pReportingStep
    pReportingStep = CInt(txtReporting.Text) * 24 + Hour(dtReporting.value)
    pReportingStep = pReportingStep & ":" & Minute(dtReporting.value) & ":" & Second(dtReporting.value)
    
    
    pOptionProperty.add "FLOW_UNITS", "CFS"             'CONSTANT
    pOptionProperty.add "INFILTRATION", "GREEN_AMPT"    'CONSTANT
    pOptionProperty.add "FLOW_ROUTING", "KINWAVE"       'CONSTANT
    pOptionProperty.add "START_DATE", Format(analysisStartDate.value, "mm/dd/yyyy")
    pOptionProperty.add "START_TIME", Format(analysisStartTime.value, "hh:mm:ss") 'Hour(analysisStartTime.value) & ":" & Minute(analysisStartTime.value) & ":" & Second(analysisStartTime.value)
    pOptionProperty.add "REPORT_START_DATE", Format(reportingStartDate.value, "mm/dd/yyyy")
    pOptionProperty.add "REPORT_START_TIME", Format(reportingStartTime.value, "hh:mm:ss") 'Hour(reportingStartTime.value) & ":" & Minute(reportingStartTime.value) & ":" & Second(reportingStartTime.value)
    pOptionProperty.add "END_DATE", Format(analysisEndDate.value, "mm/dd/yyyy")
    pOptionProperty.add "END_TIME", Format(analysisEndTime.value, "hh:mm:ss") ' Hour(analysisEndTime.value) & ":" & Minute(analysisEndTime.value) & ":" & Second(analysisEndTime.value)
    pOptionProperty.add "SWEEP_START", Format(sweepingStartDate.value, "mm/dd")   ' Month(sweepingStartDate.value) & "/" & Day(sweepingStartDate.value)
    pOptionProperty.add "SWEEP_END", Format(sweepingEndDate.value, "mm/dd") 'Month(sweepingEndDate.value) & "/" & Day(sweepingEndDate.value)
    pOptionProperty.add "DRY_DAYS", txtAntecedentDryDays.Text
    pOptionProperty.add "REPORT_STEP", Format(pReportingStep, "hh:mm:ss")
    pOptionProperty.add "WET_STEP", Format(pWetStep, "hh:mm:ss")
    pOptionProperty.add "DRY_STEP", Format(pDryStep, "hh:mm:ss")
    pOptionProperty.add "ROUTING_STEP", Format(pRoutingStep, "hh:mm:ss")
    pOptionProperty.add "ALLOW_PONDING", "NO"
    pOptionProperty.add "INERTIAL_DAMPING", "NONE"
    pOptionProperty.add "VARIABLE_STEP", "0.00"
    pOptionProperty.add "LENGTHENING_STEP", "0"
    pOptionProperty.add "MIN_SURFAREA", "0"
    pOptionProperty.add "NORMAL_FLOW_LIMITED", "NO"
    pOptionProperty.add "SKIP_STEADY_STATE", "NO"
    pOptionProperty.add "IGNORE_RAINFALL", "NO"
    pOptionProperty.add "SAVE POST OUTFLOWS", """" & txtPost.Text & """"
    pOptionProperty.add "SAVE OUTFLOWS", """" & txtOutflowFile.Text & """"
        
    '** Call the subroutine to save the values in a table
    ModuleSWMMFunctions.SaveSWMMOptionsToTable pOptionProperty
    
    ' Run simulation......
    Call Run_Simulation
    
End Sub


Private Sub dtRouting_Change()
    Dim pTimeStep
    pTimeStep = dtRouting.value
    If (Hour(pTimeStep) > 1) Then
        pTimeStep = "01:" & Minute(pTimeStep) & ":" & Second(pTimeStep)
        dtRouting.value = pTimeStep
    End If
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    'Make first tab active
    FrmSWMMOptions.TabOptions.Tab = 0
    
    '** Read the SWMMOptions table and update the form
    Dim pOptionDictionary As Scripting.Dictionary
    Set pOptionDictionary = LoadSWMMOptionsToDictionary
    
    If Not (pOptionDictionary Is Nothing) Then
        analysisStartDate.value = pOptionDictionary.Item("START_DATE")
        analysisStartTime.value = pOptionDictionary.Item("START_TIME")
        reportingStartDate.value = pOptionDictionary.Item("REPORT_START_DATE")
        reportingStartTime.value = pOptionDictionary.Item("REPORT_START_TIME")
        analysisEndDate.value = pOptionDictionary.Item("END_DATE")
        analysisEndTime.value = pOptionDictionary.Item("END_TIME")
        txtAntecedentDryDays.Text = pOptionDictionary.Item("DRY_DAYS")
        txtOutflowFile.Text = Replace(pOptionDictionary.Item("SAVE OUTFLOWS"), """", "")
        txtPost.Text = Replace(pOptionDictionary.Item("SAVE POST OUTFLOWS"), """", "")
        txtSWMMInputFile.Text = pOptionDictionary.Item("INPUT")
        txtPreDevInputFile.Text = pOptionDictionary.Item("OUTPUT")
        
        '** Define value for start sweeping date
        Dim pSweepStartDate
        pSweepStartDate = Replace(pOptionDictionary.Item("SWEEP_START"), ":", "/") & "/2000"
        sweepingStartDate.value = pSweepStartDate
        
        '** Define value for stop sweeping date
        Dim pSweepStopDate
        pSweepStopDate = Replace(pOptionDictionary.Item("SWEEP_END"), ":", "/") & "/2000"
        sweepingEndDate.value = pSweepStopDate
               
        '** GET ALL VALUES OF STEP
        Dim pDryStepVals, pDryStep, pDryStepD, pDryStepH, pDryStepM, pDryStepS
        Dim pWetStepVals, pWetStep, pWetStepD, pWetStepH, pWetStepM, pWetStepS
        Dim pRoutingStep
        Dim pRptStepVals, pRptStep, pRptStepD, pRptStepH, pRptStepM, pRptStepS
        
        pDryStep = pOptionDictionary.Item("DRY_STEP")
        pWetStep = pOptionDictionary.Item("WET_STEP")
        pRoutingStep = pOptionDictionary.Item("ROUTING_STEP")
        pRptStep = pOptionDictionary.Item("REPORT_STEP")
 
        '** Define the dry step value
        pDryStepVals = Split(pDryStep, ":")
        pDryStepH = pDryStepVals(0)
        pDryStepM = pDryStepVals(1)
        pDryStepS = pDryStepVals(2)
        If (pDryStepH >= 24) Then
            pDryStepD = pDryStepH \ 24
            txtDryWeatherRunoff.Text = pDryStepD
            pDryStepH = pDryStepH - (pDryStepD * 24)
        End If
        pDryStep = "1/1/2000 " & pDryStepH & ":" & pDryStepM & ":" & pDryStepS
        dtDryWeatherRunoff.value = pDryStep
        
        '** Define the wet step value
        pWetStepVals = Split(pWetStep, ":")
        pWetStepH = pWetStepVals(0)
        pWetStepM = pWetStepVals(1)
        pWetStepS = pWetStepVals(2)
        If (pWetStepH >= 24) Then
            pWetStepD = pWetStepH \ 24
            txtWetWeatherRunoff.Text = pWetStepD
            pWetStepH = pWetStepH - (pWetStepD * 24)
        End If
        pWetStep = "1/1/2000 " & pWetStepH & ":" & pWetStepM & ":" & pWetStepS
        dtWetWeatherRunoff.value = pWetStep
    
        '** Define the routing step
        dtRouting.value = pRoutingStep
        
        '** Define the reporting step
        pRptStepVals = Split(pRptStep, ":")
        pRptStepH = pRptStepVals(0)
        pRptStepM = pRptStepVals(1)
        pRptStepS = pRptStepVals(2)
        If (pRptStepH >= 24) Then
            pRptStepD = pRptStepH \ 24
            txtReporting.Text = pRptStepD
            pRptStepH = pRptStepH - (pRptStepD * 24)
        End If
        pRptStep = "1/1/2000 " & pRptStepH & ":" & pRptStepM & ":" & pRptStepS
        dtReporting.value = pRptStep
    
    End If
    
    Call Load_Timeseries
    
    Call Simulation_Load

End Sub

Private Sub Load_Timeseries()
    
    Dim strFile As String
    Dim fso As New FileSystemObject
    Dim pFile As TextStream
    Dim strLine As String
    Dim pVals
    Dim pDate As String
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDRainGages")
    If Not pTable Is Nothing Then
        
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "PropName='Source Name'"
        Dim pCursor As ICursor
        Set pCursor = pTable.Search(pQueryFilter, True)
        Dim pRow As iRow
        Set pRow = pCursor.NextRow
        If Not pRow Is Nothing Then
            strFile = Replace(pRow.value(pTable.FindField("PropValue")), """", "")
            Set pFile = fso.OpenTextFile(strFile, ForReading, True, TristateUseDefault)
            
            ' Read the First line....
            strLine = pFile.ReadLine
            'pVals = Split(strLine, " ")
            pVals = CustomSplit(strLine)
            pDate = pVals(2) & "/" & pVals(3) & "/" & pVals(1)
            If IsDate(pDate) Then analysisStartDate.value = pDate: reportingStartDate.value = pDate
            pDate = pVals(4) & ":" & pVals(5)
            If IsDate(pDate) Then analysisStartTime.value = pDate: reportingStartTime.value = pDate
            
            ' Go to Last line..........
            Do While pFile.AtEndOfStream <> True
                strLine = pFile.ReadLine
            Loop
            'pVals = Split(strLine, " ")
            pVals = CustomSplit(strLine)
            pDate = pVals(2) & "/" & pVals(3) & "/" & pVals(1)
            If IsDate(pDate) Then analysisEndDate.value = pDate
            pDate = pVals(4) & ":" & pVals(5)
            If IsDate(pDate) Then analysisEndTime.value = pDate
            
            'Close the stream....
            pFile.Close
        End If
        
    End If


End Sub

Private Sub Simulation_Load()
    
    On Error GoTo ShowError
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LUReclass")
    If (pTable Is Nothing) Then
        MsgBox "LUReclass table not found."
        Exit Sub
    End If
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
'    Dim iTimeSeries As Long
'    iTimeSeries = pCursor.FindField("TimeSeries")
    Dim iLuGroup As Long
    iLuGroup = pCursor.FindField("LUGroup")
    Dim iLuGroupID As Long
    iLuGroupID = pCursor.FindField("LUGroupID")
    Dim iLuImp As Long
    iLuImp = pCursor.FindField("Impervious")
    
    Dim pLanduseDict As Scripting.Dictionary
    Set pLanduseDict = CreateObject("Scripting.Dictionary")
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
'    Dim pTimeSeriesFile As String
    Dim pLuGroupID As Integer
    Dim pLuGroup As String
    Dim pImp As Integer

    Do While Not (pRow Is Nothing)
'        pTimeSeriesFile = pRow.value(iTimeSeries)
        pLuGroupID = pRow.value(iLuGroupID)
        pImp = CInt(pRow.value(iLuImp))
        If pImp = 1 Then
            pLuGroup = pRow.value(iLuGroup) & "_imp"
        Else
            pLuGroup = pRow.value(iLuGroup) & "_perv"
        End If
        If (Not pLanduseDict.Exists(pLuGroup)) Then
            pLanduseDict.add pLuGroup, pLuGroupID
            'Add to the predeveloped landuse combo control
            cmbPredevLanduse.AddItem pLuGroup
        End If
        'End If
        Set pRow = pCursor.NextRow
    Loop
    cmbPredevLanduse.ListIndex = 0
        
     GoTo CleanUp
ShowError:
    MsgBox "Error loading form: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
    Set pLanduseDict = Nothing

End Sub

Private Sub cmdBrowseSFile_Click()
    CommonDialog.Filter = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtSWMMInputFile.Text = CommonDialog.FileName
End Sub

Private Sub cmdPreDevBrowse_Click()
    CommonDialog.Filter = "LAND Predeveloped Input File (*.inp)|*.inp"
    CommonDialog.FileName = ""
    CommonDialog.CancelError = False
    CommonDialog.ShowSave
    txtPreDevInputFile.Text = CommonDialog.FileName
End Sub

Private Sub Run_Simulation()
    
    '** get all input values and validate them
    Dim pInputFileName As String
    Dim pPreDevFileName As String
    Dim pInfilConductivity As Double
    Dim pSuctionHead As Double
    Dim pInitialDef As Double
    Dim pPreDevLanduse As String
    
    '** Validate swmm input file
    pInputFileName = txtSWMMInputFile.Text
    If (Trim(pInputFileName) = "") Then
        MsgBox "Please specify SWMM input file to continue."
        Exit Sub
    End If
    
    '** Validate predeveloped landuse file
    pPreDevFileName = txtPreDevInputFile.Text
    If (Trim(pPreDevFileName) = "") Then
        MsgBox "Please specify SWMM predeveloped landuse file to continue."
        Exit Sub
    End If
        
    '** Get predeveloped landuse name
    pPreDevLanduse = cmbPredevLanduse.Text
    
    ' store to the Globals.....
    gPostDevfile = txtSWMMInputFile.Text
    gPreDevfile = txtPreDevInputFile.Text
    
    'Unload the form
    Unload Me
    
    '** write the SWMM output file
    ModuleSWMMFunctions.WriteSWMMProjectDetails pInputFileName
    
    '** write the SWMM output file
    ModuleSWMMFunctions.WriteSWMMPredevelopedLanduseFile pPreDevFileName, pPreDevLanduse, pInfilConductivity, pSuctionHead, pInitialDef
    
    
End Sub

Private Function LoadSWMMOptionsToDictionary() As Scripting.Dictionary
   
    Dim pTable As iTable
    Set pTable = GetInputDataTable("LANDOptions")
    If (pTable Is Nothing) Then
        Set LoadSWMMOptionsToDictionary = Nothing
        Exit Function
    End If
    
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim iPropName As Long
    iPropName = pTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim pOptionDictionary As Scripting.Dictionary
    Set pOptionDictionary = CreateObject("Scripting.Dictionary")
    Do While Not (pRow Is Nothing)
        pOptionDictionary.Item(pRow.value(iPropName)) = pRow.value(iPropValue)
        Set pRow = pCursor.NextRow
    Loop
    
    '** return the option dictionary back
    Set LoadSWMMOptionsToDictionary = pOptionDictionary
    
    GoTo CleanUp

    
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
    
End Function
