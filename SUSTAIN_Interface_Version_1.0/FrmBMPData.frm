VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form FrmBMPData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define BMP Parameters"
   ClientHeight    =   9690
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   11625
   Icon            =   "FrmBMPData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImgLstDgm 
      Left            =   10560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   265
      ImageHeight     =   199
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBMPData.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBMPData.frx":273E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBMPData.frx":4EB66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBMPData.frx":8D668
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   10560
      TabIndex        =   56
      Top             =   1080
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   480
      Left            =   10560
      TabIndex        =   55
      Top             =   360
      Width           =   960
   End
   Begin TabDlg.SSTab SSTabBMP 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   16748
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Dimensions"
      TabPicture(0)   =   "FrmBMPData.frx":C84C6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Dimensions"
      TabPicture(1)   =   "FrmBMPData.frx":C84E2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgBmpb"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Substrate Properties"
      TabPicture(2)   =   "FrmBMPData.frx":C84FE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GreenAmptON"
      Tab(2).Control(1)=   "FrameGreenAmpt"
      Tab(2).Control(2)=   "Frame9"
      Tab(2).Control(3)=   "Frame10"
      Tab(2).Control(4)=   "UnderDrainON"
      Tab(2).Control(5)=   "imgSoil"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Growth Index"
      TabPicture(3)   =   "FrmBMPData.frx":C851A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Cistern Release"
      TabPicture(4)   =   "FrmBMPData.frx":C8536
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame12"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Water Quality Parameters"
      TabPicture(5)   =   "FrmBMPData.frx":C8552
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(1)=   "Frame14"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Cost Factors"
      TabPicture(6)   =   "FrmBMPData.frx":C856E
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdEdit"
      Tab(6).Control(1)=   "cmdRemove"
      Tab(6).Control(2)=   "cmdAdd"
      Tab(6).Control(3)=   "txtSourceDetails"
      Tab(6).Control(4)=   "Frame7"
      Tab(6).Control(5)=   "lstComponents"
      Tab(6).Control(6)=   "Label78"
      Tab(6).Control(7)=   "Label79"
      Tab(6).ControlCount=   8
      TabCaption(7)   =   "Sediment"
      TabPicture(7)   =   "FrmBMPData.frx":C858A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame3"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Sediment Transport Related Parameters"
         Height          =   8175
         Left            =   -74760
         TabIndex        =   207
         Top             =   720
         Width           =   9495
         Begin MSDataGridLib.DataGrid DataGridSed 
            Height          =   6015
            Left            =   240
            TabIndex        =   208
            Top             =   480
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   10610
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
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   -67920
         TabIndex        =   204
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CheckBox GreenAmptON 
         Caption         =   "Use Green Ampt Infiltration"
         Height          =   255
         Left            =   -74760
         TabIndex        =   203
         Top             =   7200
         Width           =   2775
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   -66480
         TabIndex        =   182
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   -69360
         TabIndex        =   181
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtSourceDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   -69360
         MultiLine       =   -1  'True
         TabIndex        =   180
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Frame Frame7 
         Caption         =   "Select components and sources from the list"
         Height          =   4695
         Left            =   -74640
         TabIndex        =   166
         Top             =   1200
         Width           =   5175
         Begin VB.TextBox txtCostExp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   3840
            TabIndex        =   205
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Frame FrameVolumeType 
            Caption         =   "Volume Type"
            Height          =   615
            Left            =   120
            TabIndex        =   199
            Top             =   3300
            Width           =   4815
            Begin VB.OptionButton optVolUnderDrain 
               Caption         =   "Underdrain Volume"
               Height          =   255
               Left            =   3000
               TabIndex        =   202
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optVolMedia 
               Caption         =   "Soil Media"
               Height          =   255
               Left            =   1680
               TabIndex        =   201
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optVolTotal 
               Caption         =   "Total Volume"
               Height          =   255
               Left            =   120
               TabIndex        =   200
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txtCostUnits 
            Height          =   285
            Left            =   2040
            TabIndex        =   198
            Text            =   "1"
            Top             =   3000
            Width           =   2895
         End
         Begin VB.TextBox txtUserComponent 
            Height          =   285
            Left            =   2040
            TabIndex        =   187
            Top             =   1005
            Width           =   2895
         End
         Begin VB.CheckBox chkCCI 
            Caption         =   "Adjust cost based on ENR Construction Cost Index"
            Height          =   255
            Left            =   120
            TabIndex        =   186
            Top             =   4320
            Value           =   1  'Checked
            Width           =   4575
         End
         Begin VB.ComboBox cbxLocation 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   1380
            Width           =   2895
         End
         Begin VB.ComboBox cbxYear 
            Height          =   315
            Left            =   2040
            TabIndex        =   172
            Top             =   2190
            Width           =   2895
         End
         Begin VB.TextBox txtCost 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   960
            TabIndex        =   171
            Top             =   3960
            Width           =   1095
         End
         Begin VB.ComboBox cbxSource 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   1800
            Width           =   2895
         End
         Begin VB.ComboBox cbxComponent 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   600
            Width           =   2895
         End
         Begin VB.CheckBox chkNRCS 
            Caption         =   "Include NRCS Sources"
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   240
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ComboBox cbxUnit 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   167
            Top             =   2595
            Width           =   2895
         End
         Begin VB.Label Label70 
            Caption         =   "Cost Exponent"
            Height          =   255
            Left            =   2280
            TabIndex        =   206
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Number of Units (Per Unit)"
            Height          =   255
            Left            =   120
            TabIndex        =   197
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Label21 
            Caption         =   "User Defined Component"
            Height          =   255
            Left            =   120
            TabIndex        =   188
            Top             =   1000
            Width           =   1815
         End
         Begin VB.Label Label72 
            Caption         =   "Source Locale"
            Height          =   255
            Left            =   120
            TabIndex        =   179
            Top             =   1400
            Width           =   2055
         End
         Begin VB.Label Label73 
            Caption         =   "Source Year"
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   2200
            Width           =   1815
         End
         Begin VB.Label Label74 
            Caption         =   "Unit Cost"
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   3960
            Width           =   2175
         End
         Begin VB.Label Label75 
            Caption         =   "Unit"
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   2600
            Width           =   1815
         End
         Begin VB.Label Label76 
            Caption         =   "Source"
            Height          =   255
            Left            =   120
            TabIndex        =   175
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label77 
            Caption         =   "Functional Components"
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame FrameGreenAmpt 
         Height          =   1095
         Left            =   -74760
         TabIndex        =   159
         Top             =   7440
         Width           =   7695
         Begin VB.TextBox txtDeficit 
            Height          =   360
            Left            =   5040
            TabIndex        =   162
            Text            =   "0.3"
            Top             =   600
            Width           =   1440
         End
         Begin VB.TextBox txtConduct 
            Height          =   360
            Left            =   2520
            TabIndex        =   161
            Text            =   "0.5"
            Top             =   600
            Width           =   1440
         End
         Begin VB.TextBox txtSuction 
            Height          =   360
            Left            =   240
            TabIndex        =   160
            Text            =   "3.0"
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label Label71 
            Caption         =   "Initial Deficit (fraction)"
            Height          =   360
            Left            =   5040
            TabIndex        =   165
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label Label6 
            Caption         =   "Conductivity (in/hr)"
            Height          =   360
            Left            =   2520
            TabIndex        =   164
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label Label5 
            Caption         =   "Suction Head (in)"
            Height          =   360
            Left            =   240
            TabIndex        =   163
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Define Decay Factors (1/day)"
         Height          =   4095
         Left            =   -74880
         TabIndex        =   152
         Top             =   1380
         Width           =   9495
         Begin MSDataGridLib.DataGrid DataGridDECAY 
            Height          =   2295
            Left            =   360
            TabIndex        =   153
            Top             =   480
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   4048
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
      End
      Begin VB.Frame Frame15 
         Caption         =   "Define Underdrain Removal Rates (fraction [0-1])"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   151
         Top             =   5580
         Width           =   9495
         Begin MSDataGridLib.DataGrid DataGridREMOVAL 
            Height          =   2295
            Left            =   480
            TabIndex        =   154
            Top             =   600
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   4048
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
      End
      Begin VB.Frame Frame12 
         Caption         =   "Define Percapita Hourly Release Rate (ft3/sec)"
         Height          =   6840
         Left            =   -74760
         TabIndex        =   100
         Top             =   1620
         Width           =   9240
         Begin VB.TextBox txtHr1 
            Height          =   360
            Left            =   1230
            TabIndex        =   126
            Text            =   "0"
            Top             =   360
            Width           =   840
         End
         Begin VB.TextBox txtHr2 
            Height          =   360
            Left            =   1230
            TabIndex        =   125
            Text            =   "0"
            Top             =   905
            Width           =   840
         End
         Begin VB.TextBox txtHr3 
            Height          =   360
            Left            =   1230
            TabIndex        =   124
            Text            =   "0"
            Top             =   1451
            Width           =   840
         End
         Begin VB.TextBox txtHr4 
            Height          =   360
            Left            =   1230
            TabIndex        =   123
            Text            =   "0"
            Top             =   2040
            Width           =   840
         End
         Begin VB.TextBox txtHr8 
            Height          =   360
            Left            =   1230
            TabIndex        =   122
            Text            =   "0"
            Top             =   4080
            Width           =   840
         End
         Begin VB.TextBox txtHr5 
            Height          =   360
            Left            =   1200
            TabIndex        =   121
            Text            =   "0"
            Top             =   2520
            Width           =   840
         End
         Begin VB.TextBox txtHr6 
            Height          =   360
            Left            =   1230
            TabIndex        =   120
            Text            =   "0"
            Top             =   3087
            Width           =   840
         End
         Begin VB.TextBox txtHr7 
            Height          =   360
            Left            =   1230
            TabIndex        =   119
            Text            =   "0"
            Top             =   3600
            Width           =   840
         End
         Begin VB.TextBox txtHr12 
            Height          =   360
            Left            =   1230
            TabIndex        =   118
            Text            =   "0"
            Top             =   6360
            Width           =   840
         End
         Begin VB.TextBox txtHr9 
            Height          =   360
            Left            =   1230
            TabIndex        =   117
            Text            =   "0"
            Top             =   4680
            Width           =   840
         End
         Begin VB.TextBox txtHr10 
            Height          =   360
            Left            =   1230
            TabIndex        =   116
            Text            =   "0"
            Top             =   5268
            Width           =   840
         End
         Begin VB.TextBox txtHr11 
            Height          =   360
            Left            =   1230
            TabIndex        =   115
            Text            =   "0"
            Top             =   5760
            Width           =   840
         End
         Begin VB.TextBox txtHr24 
            Height          =   360
            Left            =   3269
            TabIndex        =   114
            Text            =   "0"
            Top             =   6360
            Width           =   840
         End
         Begin VB.TextBox txtHr13 
            Height          =   360
            Left            =   3269
            TabIndex        =   113
            Text            =   "0"
            Top             =   360
            Width           =   840
         End
         Begin VB.TextBox txtHr14 
            Height          =   360
            Left            =   3269
            TabIndex        =   112
            Text            =   "0"
            Top             =   905
            Width           =   840
         End
         Begin VB.TextBox txtHr15 
            Height          =   360
            Left            =   3269
            TabIndex        =   111
            Text            =   "0"
            Top             =   1451
            Width           =   840
         End
         Begin VB.TextBox txtHr16 
            Height          =   360
            Left            =   3269
            TabIndex        =   110
            Text            =   "0"
            Top             =   2040
            Width           =   840
         End
         Begin VB.TextBox txtHr20 
            Height          =   360
            Left            =   3269
            TabIndex        =   109
            Text            =   "0"
            Top             =   4080
            Width           =   840
         End
         Begin VB.TextBox txtHr17 
            Height          =   360
            Left            =   3269
            TabIndex        =   108
            Text            =   "0"
            Top             =   2520
            Width           =   840
         End
         Begin VB.TextBox txtHr18 
            Height          =   360
            Left            =   3269
            TabIndex        =   107
            Text            =   "0"
            Top             =   3087
            Width           =   840
         End
         Begin VB.TextBox txtHr19 
            Height          =   360
            Left            =   3269
            TabIndex        =   106
            Text            =   "0"
            Top             =   3600
            Width           =   840
         End
         Begin VB.TextBox txtHr21 
            Height          =   360
            Left            =   3269
            TabIndex        =   105
            Text            =   "0"
            Top             =   4680
            Width           =   840
         End
         Begin VB.TextBox txtHr22 
            Height          =   360
            Left            =   3269
            TabIndex        =   104
            Text            =   "0"
            Top             =   5268
            Width           =   840
         End
         Begin VB.TextBox txtHr23 
            Height          =   360
            Left            =   3240
            TabIndex        =   103
            Text            =   "0"
            Top             =   5760
            Width           =   840
         End
         Begin VB.CommandButton cmdReleaseCurve 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   4320
            TabIndex        =   101
            Top             =   360
            Width           =   855
         End
         Begin MSChart20Lib.MSChart ReleaseChart 
            Height          =   3975
            Left            =   4320
            OleObjectBlob   =   "FrmBMPData.frx":C85A6
            TabIndex        =   102
            Top             =   960
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.Label Label46 
            Caption         =   "Hour 1"
            Height          =   360
            Left            =   270
            TabIndex        =   150
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label47 
            Caption         =   "Hour 2"
            Height          =   360
            Left            =   270
            TabIndex        =   149
            Top             =   905
            Width           =   720
         End
         Begin VB.Label Label48 
            Caption         =   "Hour 3"
            Height          =   360
            Left            =   270
            TabIndex        =   148
            Top             =   1451
            Width           =   720
         End
         Begin VB.Label Label49 
            Caption         =   "Hour 4"
            Height          =   360
            Left            =   270
            TabIndex        =   147
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label50 
            Caption         =   "Hour 5"
            Height          =   360
            Left            =   270
            TabIndex        =   146
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label Label51 
            Caption         =   "Hour 6"
            Height          =   360
            Left            =   270
            TabIndex        =   145
            Top             =   3087
            Width           =   720
         End
         Begin VB.Label Label52 
            Caption         =   "Hour 7"
            Height          =   360
            Left            =   270
            TabIndex        =   144
            Top             =   3600
            Width           =   720
         End
         Begin VB.Label Label53 
            Caption         =   "Hour 8"
            Height          =   360
            Left            =   270
            TabIndex        =   143
            Top             =   4080
            Width           =   720
         End
         Begin VB.Label Label54 
            Caption         =   "Hour 9"
            Height          =   360
            Left            =   270
            TabIndex        =   142
            Top             =   4680
            Width           =   720
         End
         Begin VB.Label Label55 
            Caption         =   "Hour 10"
            Height          =   360
            Left            =   270
            TabIndex        =   141
            Top             =   5268
            Width           =   720
         End
         Begin VB.Label Label56 
            Caption         =   "Hour 11"
            Height          =   360
            Left            =   270
            TabIndex        =   140
            Top             =   5760
            Width           =   720
         End
         Begin VB.Label Label57 
            Caption         =   "Hour 12"
            Height          =   360
            Left            =   270
            TabIndex        =   139
            Top             =   6360
            Width           =   720
         End
         Begin VB.Label Label58 
            Caption         =   "Hour 13"
            Height          =   360
            Left            =   2310
            TabIndex        =   138
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label59 
            Caption         =   "Hour 14"
            Height          =   360
            Left            =   2310
            TabIndex        =   137
            Top             =   905
            Width           =   720
         End
         Begin VB.Label Label60 
            Caption         =   "Hour 15"
            Height          =   360
            Left            =   2310
            TabIndex        =   136
            Top             =   1451
            Width           =   720
         End
         Begin VB.Label Label61 
            Caption         =   "Hour 16"
            Height          =   360
            Left            =   2310
            TabIndex        =   135
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label62 
            Caption         =   "Hour 17"
            Height          =   360
            Left            =   2310
            TabIndex        =   134
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label Label63 
            Caption         =   "Hour 18"
            Height          =   360
            Left            =   2310
            TabIndex        =   133
            Top             =   3087
            Width           =   720
         End
         Begin VB.Label Label64 
            Caption         =   "Hour 19"
            Height          =   360
            Left            =   2310
            TabIndex        =   132
            Top             =   3600
            Width           =   720
         End
         Begin VB.Label Label65 
            Caption         =   "Hour 20"
            Height          =   360
            Left            =   2310
            TabIndex        =   131
            Top             =   4080
            Width           =   720
         End
         Begin VB.Label Label66 
            Caption         =   "Hour 21"
            Height          =   360
            Left            =   2310
            TabIndex        =   130
            Top             =   4680
            Width           =   720
         End
         Begin VB.Label Label67 
            Caption         =   "Hour 22"
            Height          =   360
            Left            =   2310
            TabIndex        =   129
            Top             =   5268
            Width           =   720
         End
         Begin VB.Label Label68 
            Caption         =   "Hour 23"
            Height          =   360
            Left            =   2310
            TabIndex        =   128
            Top             =   5760
            Width           =   720
         End
         Begin VB.Label Label69 
            Caption         =   "Hour 24"
            Height          =   360
            Left            =   2310
            TabIndex        =   127
            Top             =   6360
            Width           =   720
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Define Growth Index"
         Height          =   6480
         Left            =   -74760
         TabIndex        =   75
         Top             =   1620
         Width           =   4080
         Begin VB.TextBox Month1 
            Height          =   360
            Left            =   2350
            TabIndex        =   87
            Text            =   "0"
            Top             =   360
            Width           =   1080
         End
         Begin VB.TextBox Month2 
            Height          =   360
            Left            =   2350
            TabIndex        =   86
            Text            =   "0"
            Top             =   839
            Width           =   1080
         End
         Begin VB.TextBox Month3 
            Height          =   360
            Left            =   2350
            TabIndex        =   85
            Text            =   "0"
            Top             =   1320
            Width           =   1080
         End
         Begin VB.TextBox Month4 
            Height          =   360
            Left            =   2350
            TabIndex        =   84
            Text            =   "0"
            Top             =   1800
            Width           =   1080
         End
         Begin VB.TextBox Month5 
            Height          =   360
            Left            =   2350
            TabIndex        =   83
            Text            =   "0"
            Top             =   2279
            Width           =   1080
         End
         Begin VB.TextBox Month6 
            Height          =   360
            Left            =   2350
            TabIndex        =   82
            Text            =   "0"
            Top             =   2760
            Width           =   1080
         End
         Begin VB.TextBox Month7 
            Height          =   360
            Left            =   2350
            TabIndex        =   81
            Text            =   "0"
            Top             =   3240
            Width           =   1080
         End
         Begin VB.TextBox Month8 
            Height          =   360
            Left            =   2350
            TabIndex        =   80
            Text            =   "0"
            Top             =   3719
            Width           =   1080
         End
         Begin VB.TextBox Month9 
            Height          =   360
            Left            =   2350
            TabIndex        =   79
            Text            =   "0"
            Top             =   4200
            Width           =   1080
         End
         Begin VB.TextBox Month10 
            Height          =   360
            Left            =   2350
            TabIndex        =   78
            Text            =   "0"
            Top             =   4680
            Width           =   1080
         End
         Begin VB.TextBox Month11 
            Height          =   360
            Left            =   2350
            TabIndex        =   77
            Text            =   "0"
            Top             =   5159
            Width           =   1080
         End
         Begin VB.TextBox Month12 
            Height          =   360
            Left            =   2350
            TabIndex        =   76
            Text            =   "0"
            Top             =   5640
            Width           =   1080
         End
         Begin VB.Label Label34 
            Caption         =   "January"
            Height          =   360
            Left            =   770
            TabIndex        =   99
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label35 
            Caption         =   "February"
            Height          =   360
            Left            =   770
            TabIndex        =   98
            Top             =   839
            Width           =   1080
         End
         Begin VB.Label Label36 
            Caption         =   "March"
            Height          =   360
            Left            =   770
            TabIndex        =   97
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Label Label37 
            Caption         =   "April"
            Height          =   360
            Left            =   770
            TabIndex        =   96
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label Label38 
            Caption         =   "May"
            Height          =   360
            Left            =   770
            TabIndex        =   95
            Top             =   2279
            Width           =   1080
         End
         Begin VB.Label Label39 
            Caption         =   "June"
            Height          =   360
            Left            =   770
            TabIndex        =   94
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Label Label40 
            Caption         =   "July"
            Height          =   360
            Left            =   770
            TabIndex        =   93
            Top             =   3240
            Width           =   1080
         End
         Begin VB.Label Label41 
            Caption         =   "August"
            Height          =   360
            Left            =   770
            TabIndex        =   92
            Top             =   3719
            Width           =   1080
         End
         Begin VB.Label Label42 
            Caption         =   "September"
            Height          =   360
            Left            =   770
            TabIndex        =   91
            Top             =   4200
            Width           =   1080
         End
         Begin VB.Label Label43 
            Caption         =   "October"
            Height          =   360
            Left            =   770
            TabIndex        =   90
            Top             =   4680
            Width           =   1080
         End
         Begin VB.Label Label44 
            Caption         =   "November"
            Height          =   360
            Left            =   770
            TabIndex        =   89
            Top             =   5159
            Width           =   1080
         End
         Begin VB.Label Label45 
            Caption         =   "December"
            Height          =   360
            Left            =   770
            TabIndex        =   88
            Top             =   5640
            Width           =   1080
         End
      End
      Begin VB.Frame Frame9 
         Height          =   3960
         Left            =   -69480
         TabIndex        =   66
         Top             =   1500
         Width           =   4440
         Begin VB.TextBox txtWilting 
            Height          =   360
            Left            =   2310
            TabIndex        =   156
            Text            =   "0.15"
            Top             =   2160
            Width           =   1440
         End
         Begin VB.TextBox txtCapacity 
            Height          =   360
            Left            =   2310
            TabIndex        =   155
            Text            =   "0.3"
            Top             =   1560
            Width           =   1440
         End
         Begin VB.TextBox SoilDepth 
            Height          =   360
            Left            =   2310
            TabIndex        =   70
            Top             =   360
            Width           =   1440
         End
         Begin VB.TextBox SoilPorosity 
            Height          =   360
            Left            =   2310
            TabIndex        =   69
            Top             =   960
            Width           =   1440
         End
         Begin VB.TextBox VegetativeParam 
            Height          =   360
            Left            =   2310
            TabIndex        =   68
            Top             =   2760
            Width           =   1440
         End
         Begin VB.TextBox SoilLayerInfiltration 
            Height          =   360
            Left            =   2310
            TabIndex        =   67
            Top             =   3360
            Width           =   1440
         End
         Begin VB.Label Label4 
            Caption         =   "Soil Wilting Point"
            Height          =   360
            Left            =   270
            TabIndex        =   158
            Top             =   2160
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "Soil Field Capacity"
            Height          =   360
            Left            =   270
            TabIndex        =   157
            Top             =   1560
            Width           =   1680
         End
         Begin VB.Image SoilDOptimized2 
            Height          =   360
            Left            =   3870
            Picture         =   "FrmBMPData.frx":CBEF7
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label27 
            Caption         =   "Depth of Soil, Ds (ft):"
            Height          =   360
            Left            =   270
            TabIndex        =   74
            Top             =   360
            Width           =   1680
         End
         Begin VB.Label Label28 
            Caption         =   "Soil Porosity (0-1):"
            Height          =   360
            Left            =   270
            TabIndex        =   73
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label29 
            Caption         =   "Vegetative Parameter A:"
            Height          =   360
            Left            =   270
            TabIndex        =   72
            Top             =   2760
            Width           =   2040
         End
         Begin VB.Label Label30 
            Caption         =   "Soil Layer Infiltration (in/hr):"
            Height          =   360
            Left            =   270
            TabIndex        =   71
            Top             =   3360
            Width           =   2160
         End
         Begin VB.Image SoilDOptimized 
            Height          =   360
            Left            =   3870
            Picture         =   "FrmBMPData.frx":CC201
            Stretch         =   -1  'True
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame Frame10 
         Height          =   1200
         Left            =   -74760
         TabIndex        =   59
         Top             =   5865
         Width           =   7680
         Begin VB.TextBox StorageDepth 
            Height          =   360
            Left            =   270
            TabIndex        =   62
            Top             =   600
            Width           =   1440
         End
         Begin VB.TextBox VoidFraction 
            Height          =   360
            Left            =   2550
            TabIndex        =   61
            Top             =   600
            Width           =   1440
         End
         Begin VB.TextBox BackgroundInfiltration 
            Height          =   360
            Left            =   5070
            TabIndex        =   60
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label Label31 
            Caption         =   "Storage Depth (Du, ft)"
            Height          =   360
            Left            =   270
            TabIndex        =   65
            Top             =   240
            Width           =   1800
         End
         Begin VB.Label Label32 
            Caption         =   "Media Void Fraction (0-1):"
            Height          =   360
            Left            =   2550
            TabIndex        =   64
            Top             =   240
            Width           =   1920
         End
         Begin VB.Label Label33 
            Caption         =   "Background Infiltration (in/hr):"
            Height          =   360
            Left            =   5070
            TabIndex        =   63
            Top             =   240
            Width           =   2280
         End
      End
      Begin VB.CheckBox UnderDrainON 
         Caption         =   "Consider Underdrain Structure:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   58
         Top             =   5565
         Width           =   3300
      End
      Begin VB.Frame Frame6 
         Caption         =   "Basic Dimensions"
         Height          =   1080
         Left            =   240
         TabIndex        =   36
         Top             =   1500
         Width           =   9840
         Begin VB.TextBox BMPUnitsA 
            Height          =   300
            Left            =   1710
            TabIndex        =   190
            Text            =   "1"
            Top             =   630
            Width           =   2400
         End
         Begin VB.TextBox BMPDrainAreaA 
            Height          =   300
            Left            =   6240
            TabIndex        =   189
            Text            =   "0"
            Top             =   620
            Width           =   2775
         End
         Begin VB.TextBox BMPWidthA 
            Height          =   300
            Left            =   6240
            TabIndex        =   4
            Top             =   255
            Width           =   2775
         End
         Begin VB.TextBox BMPLengthA 
            Height          =   300
            Left            =   1710
            TabIndex        =   3
            Top             =   260
            Width           =   2400
         End
         Begin VB.Label Label23 
            Caption         =   "Number of Units"
            Height          =   300
            Left            =   360
            TabIndex        =   192
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Drainage Area (ac)"
            Height          =   360
            Left            =   4560
            TabIndex        =   191
            Top             =   620
            Width           =   1395
         End
         Begin VB.Image NumUnitsOptimized 
            Height          =   360
            Left            =   4200
            Picture         =   "FrmBMPData.frx":CC50B
            Stretch         =   -1  'True
            ToolTipText     =   "Decision parameters for basin length"
            Top             =   620
            Width           =   360
         End
         Begin VB.Image BWidthOptimized 
            Height          =   360
            Left            =   9120
            Picture         =   "FrmBMPData.frx":CC815
            Stretch         =   -1  'True
            ToolTipText     =   "Decision parameters for basin width"
            Top             =   240
            Width           =   360
         End
         Begin VB.Image BLengthOptimized 
            Height          =   360
            Left            =   4200
            Picture         =   "FrmBMPData.frx":CCB1F
            Stretch         =   -1  'True
            ToolTipText     =   "Decision parameters for basin length"
            Top             =   240
            Width           =   360
         End
         Begin VB.Label labelWidth 
            Caption         =   "              Width (ft)"
            Height          =   300
            Left            =   4680
            TabIndex        =   38
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label labelLength 
            Caption         =   " Length (ft)"
            Height          =   300
            Left            =   360
            TabIndex        =   37
            Top             =   315
            Width           =   1200
         End
         Begin VB.Image BLengthOptimized2 
            Height          =   360
            Left            =   4200
            Picture         =   "FrmBMPData.frx":CCE29
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image BWidthOptimized2 
            Height          =   360
            Left            =   9120
            Picture         =   "FrmBMPData.frx":CD133
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image NumUnitsOptimized2 
            Height          =   360
            Left            =   4200
            Picture         =   "FrmBMPData.frx":CD43D
            Stretch         =   -1  'True
            Top             =   620
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Weir Configuration"
         Height          =   2640
         Left            =   240
         TabIndex        =   27
         Top             =   6700
         Width           =   9840
         Begin VB.TextBox BMPWeirHeight 
            Height          =   360
            Left            =   7470
            TabIndex        =   9
            Top             =   360
            Width           =   1560
         End
         Begin VB.TextBox BMPTriangularWeirAngle 
            Height          =   300
            Left            =   8430
            TabIndex        =   11
            Top             =   1560
            Width           =   1080
         End
         Begin VB.TextBox BMPRectWeirWidth 
            Height          =   300
            Left            =   8430
            TabIndex        =   10
            Top             =   960
            Width           =   1080
         End
         Begin VB.Frame Frame5 
            Caption         =   "Weir Type"
            Height          =   2400
            Left            =   150
            TabIndex        =   28
            Top             =   200
            Width           =   5640
            Begin VB.OptionButton WeirType2 
               Height          =   300
               Left            =   4200
               TabIndex        =   30
               Top             =   2040
               Width           =   375
            End
            Begin VB.OptionButton WeirType1 
               Height          =   300
               Left            =   1680
               TabIndex        =   29
               Top             =   2040
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.Image imgBmpa3 
               Height          =   1800
               Left            =   150
               Picture         =   "FrmBMPData.frx":CD747
               Stretch         =   -1  'True
               Top             =   240
               Width           =   5280
            End
         End
         Begin VB.Image WHeightOptimized 
            Height          =   360
            Left            =   9120
            Picture         =   "FrmBMPData.frx":E736D
            Stretch         =   -1  'True
            ToolTipText     =   "Decision parameters for weir height"
            Top             =   315
            Width           =   360
         End
         Begin VB.Label Label7 
            Caption         =   "Weir Height (Hw, ft)"
            Height          =   360
            Left            =   5910
            TabIndex        =   35
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label Label11 
            Caption         =   "Triangular Weir"
            Height          =   240
            Left            =   5910
            TabIndex        =   34
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label Label10 
            Caption         =   "Vertex Angle (theta, deg)"
            Height          =   360
            Left            =   5910
            TabIndex        =   33
            Top             =   1680
            Width           =   2040
         End
         Begin VB.Label Label9 
            Caption         =   "Weir Crest Width (B, ft)"
            Height          =   360
            Left            =   5910
            TabIndex        =   32
            Top             =   1080
            Width           =   1920
         End
         Begin VB.Label Label8 
            Caption         =   "Rectangular Weir"
            Height          =   240
            Left            =   5910
            TabIndex        =   31
            Top             =   840
            Width           =   1800
         End
         Begin VB.Image WHeightOptimized2 
            Height          =   360
            Left            =   9120
            Picture         =   "FrmBMPData.frx":E7677
            Stretch         =   -1  'True
            Top             =   315
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Surface Storage Configuration"
         Height          =   4100
         Left            =   240
         TabIndex        =   17
         Top             =   2600
         Width           =   9840
         Begin VB.Frame FrameReleaseOption 
            Caption         =   "Release Option"
            Height          =   1800
            Left            =   4590
            TabIndex        =   19
            Top             =   2160
            Width           =   4680
            Begin VB.OptionButton OptionRelNone 
               Caption         =   "None"
               Height          =   360
               Left            =   150
               TabIndex        =   22
               Top             =   1200
               Value           =   -1  'True
               Width           =   1470
            End
            Begin VB.TextBox NumDays 
               Height          =   360
               Left            =   3630
               TabIndex        =   8
               Top             =   720
               Width           =   840
            End
            Begin VB.OptionButton OptionRelRainB 
               Caption         =   "Rain Barrel"
               Height          =   300
               Left            =   150
               TabIndex        =   21
               Top             =   780
               Width           =   1350
            End
            Begin VB.TextBox NumPeople 
               Height          =   360
               Left            =   3630
               TabIndex        =   7
               Top             =   240
               Width           =   840
            End
            Begin VB.OptionButton OptionRelCistern 
               Caption         =   "Cistern"
               Height          =   300
               Left            =   150
               TabIndex        =   20
               Top             =   300
               Width           =   1470
            End
            Begin VB.Label Label15 
               Caption         =   "Number of Dry Days"
               Height          =   240
               Left            =   1800
               TabIndex        =   24
               Top             =   780
               Width           =   1680
            End
            Begin VB.Label Label14 
               Caption         =   "Number of People"
               Height          =   240
               Left            =   1800
               TabIndex        =   23
               Top             =   300
               Width           =   1320
            End
         End
         Begin VB.TextBox BMPOrificeHeight 
            Height          =   300
            Left            =   2430
            TabIndex        =   6
            Top             =   3700
            Width           =   1080
         End
         Begin VB.TextBox BMPOrificeDiameter 
            Height          =   300
            Left            =   2430
            TabIndex        =   5
            Top             =   3360
            Width           =   1080
         End
         Begin VB.Frame FrameExitType 
            Caption         =   "Exit Type"
            Height          =   2040
            Left            =   4560
            TabIndex        =   18
            Top             =   120
            Width           =   4680
            Begin VB.OptionButton ExitType2 
               Height          =   270
               Left            =   1590
               TabIndex        =   13
               Top             =   1680
               Width           =   375
            End
            Begin VB.OptionButton ExitType4 
               Height          =   270
               Left            =   3929
               TabIndex        =   15
               Top             =   1680
               Width           =   375
            End
            Begin VB.OptionButton ExitType3 
               Height          =   270
               Left            =   2729
               TabIndex        =   14
               Top             =   1680
               Width           =   375
            End
            Begin VB.OptionButton ExitType1 
               Height          =   270
               Left            =   270
               TabIndex        =   12
               Top             =   1680
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.Image imgBmpa2 
               Height          =   1275
               Left            =   150
               Picture         =   "FrmBMPData.frx":E7981
               Top             =   240
               Width           =   4440
            End
         End
         Begin VB.Image imgBmpa1 
            Height          =   2970
            Left            =   120
            Picture         =   "FrmBMPData.frx":FA09B
            Top             =   240
            Width           =   3960
         End
         Begin VB.Label LabelOrificeHeight 
            Caption         =   "Orifice Height (Ho, ft)"
            Height          =   300
            Left            =   750
            TabIndex        =   26
            Top             =   3700
            Width           =   1680
         End
         Begin VB.Label LabelOrificeDiameter 
            Caption         =   "Orifice Diameter (in)"
            Height          =   300
            Left            =   750
            TabIndex        =   25
            Top             =   3360
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Information"
         Height          =   600
         Left            =   240
         TabIndex        =   1
         Top             =   920
         Width           =   9840
         Begin VB.TextBox BMPNameA 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2670
            TabIndex        =   2
            Top             =   200
            Width           =   6465
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   360
            Left            =   270
            TabIndex        =   16
            Top             =   200
            Width           =   2161
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "BMP Information"
         Height          =   2520
         Left            =   -74880
         TabIndex        =   39
         Top             =   5820
         Width           =   9120
         Begin VB.TextBox BMPUnitsB 
            Height          =   360
            Left            =   1470
            TabIndex        =   194
            Text            =   "1"
            Top             =   1980
            Width           =   2400
         End
         Begin VB.TextBox BMPDrainAreaB 
            Height          =   375
            Left            =   6000
            TabIndex        =   193
            Text            =   "0"
            Top             =   1980
            Width           =   2775
         End
         Begin VB.TextBox BMPWidthB 
            Height          =   360
            Left            =   870
            TabIndex        =   57
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox BMPManningsN 
            Height          =   360
            Left            =   6390
            TabIndex        =   46
            Top             =   360
            Width           =   1080
         End
         Begin VB.TextBox BMPSlope3 
            Height          =   360
            Left            =   6150
            TabIndex        =   45
            Top             =   1440
            Width           =   1080
         End
         Begin VB.TextBox BMPSlope2 
            Height          =   360
            Left            =   3390
            TabIndex        =   44
            Top             =   1440
            Width           =   1080
         End
         Begin VB.TextBox BMPSlope1 
            Height          =   360
            Left            =   870
            TabIndex        =   43
            Top             =   1440
            Width           =   1080
         End
         Begin VB.TextBox BMPNameB 
            Height          =   360
            Left            =   1230
            TabIndex        =   42
            Top             =   360
            Width           =   3585
         End
         Begin VB.TextBox BMPMaxDepth 
            Height          =   360
            Left            =   6120
            TabIndex        =   41
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox BMPLengthB 
            Height          =   360
            Left            =   3390
            TabIndex        =   40
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label25 
            Caption         =   "Number of Units"
            Height          =   240
            Left            =   120
            TabIndex        =   196
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Design Drainage Area (ac)"
            Height          =   360
            Left            =   4320
            TabIndex        =   195
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Image NumUnitsOptimizedB 
            Height          =   360
            Left            =   3960
            Picture         =   "FrmBMPData.frx":12056D
            Stretch         =   -1  'True
            ToolTipText     =   "Decision parameters for basin length"
            Top             =   1965
            Width           =   360
         End
         Begin VB.Image BDepthBOptimized 
            Height          =   360
            Left            =   7440
            Picture         =   "FrmBMPData.frx":120877
            Stretch         =   -1  'True
            Top             =   960
            Width           =   360
         End
         Begin VB.Image BLengthBOptimized 
            Height          =   360
            Left            =   4560
            Picture         =   "FrmBMPData.frx":120B81
            Stretch         =   -1  'True
            Top             =   960
            Width           =   360
         End
         Begin VB.Image BWidthBOptimized 
            Height          =   360
            Left            =   2040
            Picture         =   "FrmBMPData.frx":120E8B
            Stretch         =   -1  'True
            Top             =   960
            Width           =   360
         End
         Begin VB.Label Label20 
            Caption         =   "Manning's 'N'"
            Height          =   360
            Left            =   5070
            TabIndex        =   54
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label19 
            Caption         =   "Slope 3"
            Height          =   360
            Left            =   4920
            TabIndex        =   53
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label18 
            Caption         =   "Slope 2"
            Height          =   360
            Left            =   2520
            TabIndex        =   52
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label17 
            Caption         =   "Slope 1"
            Height          =   360
            Left            =   150
            TabIndex        =   51
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label16 
            Caption         =   "BMP Name"
            Height          =   360
            Left            =   150
            TabIndex        =   50
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label13 
            Caption         =   "Max. Depth (ft)"
            Height          =   240
            Left            =   4950
            TabIndex        =   49
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Label Label12 
            Caption         =   "Length (ft)"
            Height          =   240
            Left            =   2520
            TabIndex        =   48
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label Label2 
            Caption         =   "Width (ft)"
            Height          =   240
            Left            =   120
            TabIndex        =   47
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Image BWidthBOptimized2 
            Height          =   360
            Left            =   2040
            Picture         =   "FrmBMPData.frx":121195
            Stretch         =   -1  'True
            Top             =   960
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image BLengthBOptimized2 
            Height          =   360
            Left            =   4560
            Picture         =   "FrmBMPData.frx":12149F
            Stretch         =   -1  'True
            Top             =   960
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image BDepthBOptimized2 
            Height          =   360
            Left            =   7440
            Picture         =   "FrmBMPData.frx":1217A9
            Stretch         =   -1  'True
            Top             =   960
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Image NumUnitsOptimizedB2 
            Height          =   360
            Left            =   3960
            Picture         =   "FrmBMPData.frx":121AB3
            Stretch         =   -1  'True
            Top             =   1965
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin MSComctlLib.ListView lstComponents 
         Height          =   2175
         Left            =   -74640
         TabIndex        =   183
         Top             =   6240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         NumItems        =   11
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
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Cost Exponent"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label78 
         Caption         =   "Source Details"
         Height          =   255
         Left            =   -69360
         TabIndex        =   185
         Top             =   1250
         Width           =   2175
      End
      Begin VB.Label Label79 
         Caption         =   "Selected components"
         Height          =   255
         Left            =   -74640
         TabIndex        =   184
         Top             =   6000
         Width           =   3495
      End
      Begin VB.Image imgSoil 
         Height          =   3825
         Left            =   -74760
         Picture         =   "FrmBMPData.frx":121DBD
         Top             =   1605
         Width           =   5160
      End
      Begin VB.Image imgBmpb 
         Height          =   3375
         Left            =   -74520
         Picture         =   "FrmBMPData.frx":1621F7
         Top             =   1980
         Width           =   7680
      End
   End
End
Attribute VB_Name = "FrmBMPData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private pAdoConn As ADODB.Connection
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmBMPData.frm"
Private costAdjDict As Scripting.Dictionary
Private maxCCIYear As Integer



Private Sub UpdateCost()
On Error GoTo ErrorHandler:
' Changed the design of cost module - Sabu Paul, September 2007
''    Dim currentCost As Double
''    Dim dblDepth As Double
''    Dim dblAb As Double
''    Dim dblAa As Double
''    Dim dblDb As Double
''    Dim dblDa As Double
''    Dim dblConstCost As Double
''    Dim dblLdCost As Double
''
''    If (Aa.Text <> "") Then
''        dblAa = CDbl(Trim(Aa.Text))
''    End If
''    If (Ab.Text <> "") Then
''        dblAb = CDbl(Trim(Ab.Text))
''    End If
''    If (Da.Text <> "") Then
''        dblDa = CDbl(Trim(Da.Text))
''    End If
''    If (Db.Text <> "") Then
''        dblDb = CDbl(Trim(Db.Text))
''    End If
''    If (ConstCost.Text <> "") Then
''        dblConstCost = CDbl(Trim(ConstCost.Text))
''    End If
''    If (LdCost.Text <> "") Then
''        dblLdCost = CDbl(Trim(LdCost.Text))
''    End If

''    Dim bmpArea As Double
''    bmpArea = 0
''    If SSTabBMP.TabVisible(0) = True Then
''        If (BMPWidthA.Text <> "" And BMPLengthA.Text <> "") Then
''            If (OptionRelRainB.value = True Or OptionRelCistern.value = True) Then
''                '** area = 3.14 * r2 * number of units
''                bmpArea = (3.14285 / 4) * (CDbl(Trim(BMPLengthA.Text))) * (CDbl(Trim(BMPLengthA.Text))) * CDbl(Trim(BMPWidthA.Text))
''            Else
''                bmpArea = CDbl(Trim(BMPWidthA.Text)) * CDbl(Trim(BMPLengthA.Text))
''            End If
''        End If
''    Else
''        If (BMPWidthB.Text <> "" And BMPLengthB.Text <> "") Then
''            bmpArea = CDbl(Trim(BMPWidthB.Text)) * CDbl(Trim(BMPLengthB.Text))
''        End If
''    End If
 ' Changed the design of cost module - Sabu Paul, September 2007
''    dblDepth = 0
''    If SSTabBMP.TabVisible(0) = True Then   'TYPE A
''        If (BMPWeirHeight.Text <> "") Then
''             dblDepth = CDbl(Trim(BMPWeirHeight.Text))
''        End If
''    ElseIf SSTabBMP.TabVisible(1) = True Then   'TYPE B
''        If (BMPMaxDepth.Text <> "") Then
''            dblDepth = CDbl(Trim(BMPMaxDepth.Text))
''        End If
''    End If
''    If (SSTabBMP.TabVisible(2) = True) Then 'soil properties
''        If (SoilDepth.Text <> "") Then
''            dblDepth = dblDepth + CDbl(Trim(SoilDepth.Text))
''        End If
''        If (StorageDepth.Text <> "") Then
''            dblDepth = dblDepth + CDbl(Trim(StorageDepth.Text))
''        End If
''    End If
''
''    currentCost = ((dblDa * (dblDepth ^ dblDb)) * (dblAa * (bmpArea ^ dblAb))) + (bmpArea * dblLdCost) + dblConstCost
''    totalCost.Text = currentCost
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in updating BMP cost " & Err.description
End Sub

''Private Sub Aa_Change()
''  On Error GoTo ErrorHandler
''
''    If Not IsNumeric(Aa.Text) Then
''        MsgBox "Aa must be a valid number."
''        Aa.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "Aa_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

''Private Sub Ab_Change()
''  On Error GoTo ErrorHandler
''
''    If Not (IsNumeric(Ab.Text)) Then
''        MsgBox "Ab must be a valid number."
''        Ab.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "Ab_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

''Private Sub BMPLengthA_Change()
''  On Error GoTo ErrorHandler
''
''    If (BMPLengthA.Text <> "") Then
''        If IsNumeric(BMPLengthA.Text) Then
''            Call UpdateCost
''        End If
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "BMPLengthA_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub
''
''Private Sub BMPLengthB_Change()
''  On Error GoTo ErrorHandler
''
''    If (BMPLengthB.Text <> "") Then
''        If IsNumeric(BMPLengthB.Text) Then
''            Call UpdateCost
''        End If
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "BMPLengthB_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

Private Sub BMPManningsN_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler

    If Trim(BMPManningsN.Text) <> "" Then
        If IsNumeric(Trim(BMPManningsN.Text)) Then
            If Not (CDbl(Trim(BMPManningsN.Text)) >= 0# And CDbl(Trim(BMPManningsN.Text)) <= 1#) Then
                MsgBox "Manning's N should be between 0 and 1", vbExclamation
            End If
        Else
             MsgBox "Manning's N should be a number", vbExclamation
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "BMPManningsN_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


''Private Sub BMPMaxDepth_Change()
''  On Error GoTo ErrorHandler
''
''    If (BMPMaxDepth.Text <> "") Then
''        If IsNumeric(BMPMaxDepth.Text) Then
''            Call UpdateCost
''        End If
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "BMPMaxDepth_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

Private Sub BMPTriangularWeirAngle_Change()
  On Error GoTo ErrorHandler

    If IsNumeric(Trim(BMPTriangularWeirAngle.Text)) Then
        If ((CDbl(Trim(BMPTriangularWeirAngle.Text)) < 0) Or (CDbl(Trim(BMPTriangularWeirAngle.Text)) > 180)) Then
            MsgBox "Triangular weir angle should be between 0 and 180", vbExclamation
        End If
    Else
        MsgBox "Triangular weir angle should be a number", vbExclamation
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "BMPTriangularWeirAngle_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


''Private Sub BMPWidthA_Change()
''  On Error GoTo ErrorHandler
''
''    If (BMPWidthA.Text <> "") Then
''        If IsNumeric(BMPWidthA.Text) Then
''            Call UpdateCost
''        End If
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "BMPWidthA_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

''Private Sub BMPWidthB_Change()
''  On Error GoTo ErrorHandler
''
''    If (BMPWidthB.Text <> "") Then
''        If IsNumeric(BMPWidthB.Text) Then
''            Call UpdateCost
''        End If
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "BMPWidthB_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub


Private Sub cbxUnit_Click()
    optVolTotal.value = True
    txtCostUnits.Enabled = False
    FrameVolumeType.Enabled = False
    If cbxUnit.List(cbxUnit.ListIndex) = "Cubic Feet" Then
        FrameVolumeType.Enabled = True
    ElseIf cbxUnit.List(cbxUnit.ListIndex) = "Per Unit" Then
        txtCostUnits.Enabled = True
    End If
    
End Sub

Private Sub chkNRCS_Click()
    Call cbxComponent_Click
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    Set gBMPDetailDict = Nothing
    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub cmdGetCost_Click()
  On Error GoTo ErrorHandler

    Call UpdateCost

  Exit Sub
ErrorHandler:
  HandleError False, "cmdGetCost_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub cmdOk_Click()
 On Error GoTo ShowError
    Dim pBMPName As String
    Dim pEstCost As Double
        
    Dim pBasinWidth As Double, pBasinLength As Double
    Dim pOHeigth As Double, pODia As Double
    
    Dim pExitType As Integer 'String
    
    Dim pRelOption As String
    Dim pNumPeople As Double
    Dim pNumDDays As Double
    
    Dim pWeirType As Integer
    
    Dim pWHeight As Double
    Dim pWWidth As Double
    Dim pWAngle As Double
    
    Dim pNumUnits As Integer
    Dim pDrainArea As Double
    
    Dim orificecoef As Double
    If Not ValidateInputs Then
        Exit Sub
    End If
    
    'if no cost information is defined, popup a warning box
    If lstComponents.ListItems.Count = 0 Then
        If MsgBox("There is no cost information defined.  Do you want to continue?", vbInformation + vbYesNo, "SUSTAIN") = vbNo Then
            Exit Sub
        End If
        'MsgBox "There is no cost information defined.", vbExclamation + vbOKOnly, "Warning"
    End If

    gBMPDetailDict.add "BMPType", gNewBMPType
    If SSTabBMP.TabVisible(0) = True Then
        pBMPName = Trim(BMPNameA.Text)
        
        pBasinWidth = 0
        If BMPWidthA.Enabled Then pBasinWidth = CDbl(Trim(BMPWidthA.Text))
        pBasinLength = CDbl(Trim(BMPLengthA.Text))
        pOHeigth = CDbl(Trim(BMPOrificeHeight.Text))
        pODia = CDbl(Trim(BMPOrificeDiameter.Text))
                   
        pNumUnits = CInt(Trim(BMPUnitsA.Text))
        pDrainArea = CDbl(Trim(BMPDrainAreaA.Text))
        
        If ExitType1.value = True Then
            pExitType = 1 '"ExitType1"
            orificecoef = 1#
        ElseIf ExitType2.value = True Then
            pExitType = 2 '"ExitType2"
            orificecoef = 0.61
        ElseIf ExitType3.value = True Then
            pExitType = 3 '"ExitType3"
            orificecoef = 0.61
        ElseIf ExitType4.value = True Then
            pExitType = 4 '"ExitType4"
            orificecoef = 0.5
        Else
            pExitType = 1
            orificecoef = 1
        End If
        
        
        If OptionRelCistern.value = True Then
            pRelOption = "Cistern"
        ElseIf OptionRelRainB.value = True Then
            pRelOption = "RainBarrel"
        Else
            pRelOption = "None"
        End If
        
        
        If WeirType1.value = True Then
            pWeirType = 1
        Else
            pWeirType = 2
        End If
    
        
        If pRelOption = "Cistern" Then
            pNumPeople = CInt(Trim(NumPeople.Text))
        ElseIf pRelOption = "RainBarrel" Then
            pNumDDays = CInt(Trim(NumDays.Text))
        End If
        
        pWHeight = CDbl(Trim(BMPWeirHeight.Text))
    
        If pWeirType = 1 Then
            pWWidth = CDbl(Trim(BMPRectWeirWidth.Text))
        Else
            pWAngle = CDbl(Trim(BMPTriangularWeirAngle.Text))
        End If
        
        
        'Insert into the global dictionary
        gBMPDetailDict.add "BMPName", pBMPName
        'gBMPDetailDict.add "BMPType", gNewBMPType 'Same for Class A & B
        gBMPDetailDict.add "BMPClass", "A"
        
        gBMPDetailDict.add "BMPWidth", pBasinWidth
        gBMPDetailDict.add "BMPLength", pBasinLength
        gBMPDetailDict.add BMPOrificeHeight.name, pOHeigth
        gBMPDetailDict.add BMPOrificeDiameter.name, pODia
        gBMPDetailDict.add "ExitType", pExitType
        gBMPDetailDict.add "OrificeCoef", orificecoef
        gBMPDetailDict.add "ReleaseOption", pRelOption
        gBMPDetailDict.add "WeirType", pWeirType

        gBMPDetailDict.add BMPWeirHeight.name, pWHeight
        If pWeirType = 1 Then
           gBMPDetailDict.add BMPRectWeirWidth.name, pWWidth
        Else
           gBMPDetailDict.add BMPTriangularWeirAngle.name, pWAngle
        End If
        If pRelOption = "Cistern" Then
           gBMPDetailDict.add NumPeople.name, pNumPeople
        ElseIf pRelOption = "RainBarrel" Then
           gBMPDetailDict.add NumDays.name, pNumDDays
        End If
    ElseIf SSTabBMP.TabVisible(1) = True Then
        gBMPDetailDict.add "BMPName", Trim(BMPNameB.Text)
        'gBMPDetailDict.add "BMPType", gNewBMPType 'Same for Class A & B
        gBMPDetailDict.add "BMPClass", "B"
        gBMPDetailDict.add "BMPManningsN", CDbl(Trim(BMPManningsN.Text))
        gBMPDetailDict.add "BMPWidth", CDbl(Trim(BMPWidthB.Text))
        gBMPDetailDict.add "BMPLength", CDbl(Trim(BMPLengthB.Text))
        gBMPDetailDict.add "BMPMaxDepth", CDbl(Trim(BMPMaxDepth.Text))
        gBMPDetailDict.add "BMPSlope1", CDbl(Trim(BMPSlope1.Text))
        gBMPDetailDict.add "BMPSlope2", CDbl(Trim(BMPSlope2.Text))
        gBMPDetailDict.add "BMPSlope3", CDbl(Trim(BMPSlope3.Text))
        
        pNumUnits = CInt(Trim(BMPUnitsB.Text))
        pDrainArea = CDbl(Trim(BMPDrainAreaB.Text))
    End If
    
    gBMPDetailDict.add "NumUnits", pNumUnits
    gBMPDetailDict.add "DrainArea", pDrainArea
        
    '** Add cost related parameters
    ' Changed the design of cost module - Sabu Paul, September 2007
''    gBMPDetailDict.Add "Aa", Aa.Text
''    gBMPDetailDict.Add "Ab", Ab.Text
''    gBMPDetailDict.Add "Da", Da.Text
''    gBMPDetailDict.Add "Db", Db.Text
''    gBMPDetailDict.Add "LdCost", LdCost.Text
''    gBMPDetailDict.Add "ConstCost", ConstCost.Text
    
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
    Dim costExponents As String
    
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
    costExponents = ""
    
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
         costExponents = lstComponents.ListItems.Item(1).SubItems(10)
         
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
             costExponents = costExponents & ";" & lstComponents.ListItems.Item(pIndex).SubItems(10)
        Next
        
        gBMPDetailDict.add "CostComponents", costComps
        gBMPDetailDict.add "CostComponentIds", costCompIds
        gBMPDetailDict.add "CostLocations", costLocs
        gBMPDetailDict.add "CostSources", costSrcs
        gBMPDetailDict.add "CostYears", costYears
        gBMPDetailDict.add "CostUnits", costUnits
        gBMPDetailDict.add "CostVolTypes", costVolTypes
        gBMPDetailDict.add "CostNumUnits", costNumUnits
        gBMPDetailDict.add "CostUnitCosts", costUnitCosts
        gBMPDetailDict.add "CostAdjUnitCosts", costAdjUnitCosts
        gBMPDetailDict.add "costExponents", costExponents
    End If
    '** Add growth parameters
    Dim iGrowthIndex(1 To 12) As Double
    Dim strGrowthIndex As String
    Dim pMonthIndex As Integer

    Dim pControl As Control
    For Each pControl In FrmBMPData.Controls
        If ((TypeOf pControl Is TextBox)) Then
            If (Left(pControl.name, 5) = "Month") Then
                    pMonthIndex = CInt(Replace(pControl.name, "Month", ""))
                    iGrowthIndex(pMonthIndex) = CDbl(pControl.Text)
            End If
        End If
    Next pControl
    Set pControl = Nothing
    strGrowthIndex = CStr(iGrowthIndex(1))
    Dim i As Integer
    For i = 2 To 12
        strGrowthIndex = strGrowthIndex & ";" & CStr(iGrowthIndex(i))
    Next
    gBMPDetailDict.add "GrowthIndex", strGrowthIndex
    
    'if current BMP is a cistern
    If SSTabBMP.TabVisible(4) = True Then 'Sabu Paul -- March 31, 2005 Tab index modified
        Dim waterRelease(24) As Double
        Dim incr As Long
        Dim waterRelStr As String
        
        For incr = 1 To 24
            For Each pControl In Controls
            If TypeOf pControl Is TextBox Then
                If pControl.Enabled Then
                    If pControl.name = "txtHr" & incr Then
                         waterRelease(incr) = CDbl(Trim(pControl.Text))
                    End If
                End If
            End If
            Next pControl
        Next incr
        waterRelStr = waterRelease(1)
        For incr = 2 To 24
            waterRelStr = waterRelStr & ";" & waterRelease(incr)
        Next incr
        gBMPDetailDict.add "CisternFlow", waterRelStr
    End If
    
    Dim boolDecayRemoval As Boolean
    boolDecayRemoval = False
    If SSTabBMP.TabVisible(5) = True Then  'Sabu Paul -- March 31, 2005 Tab index modified
        boolDecayRemoval = True
    End If
   
    If SSTabBMP.TabVisible(2) Then
        gBMPDetailDict.add "SoilDepth", CDbl(Trim(SoilDepth.Text))
        gBMPDetailDict.add "SoilPorosity", CDbl(Trim(SoilPorosity.Text))
        gBMPDetailDict.add "SoilFieldCapacity", CDbl(Trim(txtCapacity.Text))
        gBMPDetailDict.add "SoilWiltingPoint", CDbl(Trim(txtWilting.Text))
        gBMPDetailDict.add "VegetativeParam", CDbl(Trim(VegetativeParam.Text))
        gBMPDetailDict.add "SoilLayerInfiltration", CDbl(Trim(SoilLayerInfiltration.Text))
        gBMPDetailDict.add "StorageDepth", CDbl(Trim(StorageDepth.Text))
        gBMPDetailDict.add "VoidFraction", CDbl(Trim(VoidFraction.Text))
        gBMPDetailDict.add "BackgroundInfiltration", CDbl(Trim(BackgroundInfiltration.Text))
        gBMPDetailDict.add "UnderDrainON", CBool(UnderDrainON.value)
        gBMPDetailDict.add "GreenAmptON", CBool(GreenAmptON.value)
        gBMPDetailDict.add "SuctionHead", CDbl(Trim(txtSuction.Text))
        gBMPDetailDict.add "Conductivity", CDbl(Trim(txtConduct.Text))
        gBMPDetailDict.add "InitialDeficit", CDbl(Trim(txtDeficit.Text))
    End If
    
    
    
    '** If 5th tab is visible, enter decay and percent removal factors in the dictionary
    If boolDecayRemoval = True Then
        Dim oRs As ADODB.Recordset
        Set oRs = DataGridDECAY.DataSource
        oRs.MoveFirst
        
        Dim iDecayCount As Integer
        iDecayCount = 1
        Do Until oRs.EOF
            gBMPDetailDict.add "Decay" & iDecayCount, oRs.Fields(1).value
            gBMPDetailDict.add "K" & iDecayCount, oRs.Fields(2).value
            gBMPDetailDict.add "C" & iDecayCount, oRs.Fields(3).value
            gBMPDetailDict.add "PctRem" & iDecayCount, oRs.Fields(4).value
            iDecayCount = iDecayCount + 1
            oRs.MoveNext
        Loop
        '* Read from TEMPTable and add values in gBMPDetailDict dictionary
''        Dim pTempTable As iTable
''        Set pTempTable = GetInputDataTable("TempTable")
''        If Not (pTempTable Is Nothing) Then
''            Dim pTempCursor As ICursor
''            Dim pTempRow As iRow
''            Dim iDecay As Long
''            iDecay = pTempTable.FindField("DECAY")
''            Dim iK As Long
''            iK = pTempTable.FindField("K")
''            Dim iC As Long
''            iC = pTempTable.FindField("C")
''            Dim iRemoval As Long
''            iRemoval = pTempTable.FindField("REMOVAL")
''            Set pTempCursor = pTempTable.Search(Nothing, True)
''            Set pTempRow = pTempCursor.NextRow
''            Dim iDecayCount As Integer
''            iDecayCount = 1
''            Do While Not (pTempRow Is Nothing)
''                gBMPDetailDict.Add "Decay" & iDecayCount, pTempRow.value(iDecay)
''                gBMPDetailDict.Add "K" & iDecayCount, pTempRow.value(iK)
''                gBMPDetailDict.Add "C" & iDecayCount, pTempRow.value(iC)
''                gBMPDetailDict.Add "PctRem" & iDecayCount, pTempRow.value(iRemoval)
''                iDecayCount = iDecayCount + 1
''                'move to next row
''                Set pTempRow = pTempCursor.NextRow
''            Loop
''        End If
    End If
    
    Dim sedParams
'    sedParams = Array("Bed width", "Bed depth", "Porosity", "Sand fraction", "Silt fraction", "Clay fraction", "Sand effective diameter" _
'        , "Sand velocity", "Sand density", "Sand coeff", "Sand exponent", "Silt effective diameter", "Silt velocity", "Silt density" _
'        , "Deposition stress", "Scour stress", "Erodibility")
    sedParams = Array("Bed width", "Bed depth", "Porosity", "Sand fraction", "Silt fraction", "Clay fraction", "Sand effective diameter" _
        , "Sand velocity", "Sand density", "Sand coefficient", "Sand exponent", "Silt effective diameter", "Silt velocity", "Silt density" _
        , "Silt Deposition stress", "Silt Scour stress", "Silt Erodibility" _
        , "Clay effective diameter", "Clay velocity", "Clay density", "Clay Deposition stress", "Clay Scour stress", "Clay Erodibility")

        
    Dim oRsSed As ADODB.Recordset
    Set oRsSed = DataGridSed.DataSource
    oRsSed.MoveFirst
    Dim sedIncr As Integer
    sedIncr = 0
    Do Until oRsSed.EOF
        gBMPDetailDict.Item(sedParams(sedIncr)) = oRsSed.Fields(1).value
        sedIncr = sedIncr + 1
        oRsSed.MoveNext
    Loop
    oRsSed.Close
    
    'gBMPDetailDict.add "UnderDrainON", CBool(UnderDrainON.value)
    If CBool(GreenAmptON.value) Then
        gBMPDetailDict.add "Infiltration Method", 1
    Else
        gBMPDetailDict.add "Infiltration Method", 0
    End If
    
    'Ying: add option para Feb 10, 2009
    If (Not gBMPOptionsDict Is Nothing) Then
        'gBMPDetailDict.add "Infiltration Method", gBMPOptionsDict.Item("Infiltration Method")
        gBMPDetailDict.add "Pollutant Removal Method", gBMPOptionsDict.Item("Pollutant Removal Method")
        gBMPDetailDict.add "Pollutant Routing Method", gBMPOptionsDict.Item("Pollutant Routing Method")
    End If
    
    
    'Close the form
    Unload Me
    
    GoTo CleanUp
    
ShowError:
    Unload Me
    MsgBox "Read BMP Details: " & Err.description
CleanUp:
End Sub

Public Sub Cistern_Initialize()
  On Error GoTo ErrorHandler

    txtHr1.Text = 0.020498
    txtHr2.Text = "0.017824"
    txtHr3.Text = "0.019606"
    txtHr4.Text = "0.023171"
    txtHr5.Text = "0.035648"
    txtHr6.Text = "0.053472"
    txtHr7.Text = "0.08912"
    txtHr8.Text = "0.106944"
    txtHr9.Text = "0.115857"
    txtHr10.Text = "0.121204"
    txtHr11.Text = "0.120313"
    txtHr12.Text = "0.117639"
    txtHr13.Text = "0.114074"
    txtHr14.Text = "0.108727"
    txtHr15.Text = "0.106944"
    txtHr16.Text = "0.108727"
    txtHr17.Text = "0.115857"
    txtHr18.Text = "0.133681"
    txtHr19.Text = "0.151505"
    txtHr20.Text = "0.160417"
    txtHr21.Text = "0.151505"
    txtHr22.Text = "0.08912"
    txtHr23.Text = "0.053472"
    txtHr24.Text = "0.035648"

  Exit Sub
ErrorHandler:
  HandleError True, "Cistern_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub
Private Sub BDepthBOptimized_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BDepthB"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BDepthBOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BDepthBOptimized2_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BDepthB"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinBDepth") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinBDepth")
    End If
    If gBMPDetailDict.Exists("MaxBasinBDepth") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinBDepth")
    End If
    If gBMPDetailDict.Exists("BasinBDepthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinBDepthIncr")
    End If

    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BDepthBOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BLengthBOptimized_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BLengthB"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BLengthBOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BLengthBOptimized2_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BLengthB"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinBLength") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinBLength")
    End If
    If gBMPDetailDict.Exists("MaxBasinBLength") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinBLength")
    End If
    If gBMPDetailDict.Exists("BasinBLengthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinBLengthIncr")
    End If

    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BLengthBOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BWidthBOptimized_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BWidthB"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BWidthBOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BWidthBOptimized2_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "BWidthB"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinBWidth") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinBWidth")
    End If
    If gBMPDetailDict.Exists("MaxBasinBWidth") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinBWidth")
    End If
    If gBMPDetailDict.Exists("BasinBWidthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinBWidthIncr")
    End If

    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BWidthBOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BLengthOptimized_Click()
  On Error GoTo ErrorHandler

    'BLengthOptimized = True
    gCurOptParam = "BLength"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BLengthOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BLengthOptimized2_Click()
  On Error GoTo ErrorHandler

    'BLengthOptimized = True
    gCurOptParam = "BLength"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinLength") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinLength")
    End If
    If gBMPDetailDict.Exists("MaxBasinLength") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinLength")
    End If
    If gBMPDetailDict.Exists("BasinLengthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinLengthIncr")
    End If

    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BLengthOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub BWidthOptimized_Click()
  On Error GoTo ErrorHandler

    'BWidthOptimized = True
    gCurOptParam = "BWidth"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BWidthOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


Private Sub BWidthOptimized2_Click()
  On Error GoTo ErrorHandler

    'BLengthOptimized = True
    gCurOptParam = "BWidth"
    'Set frmOptimizer values
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinWidth") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinWidth")
    End If
    If gBMPDetailDict.Exists("MaxBasinWidth") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinWidth")
    End If
    If gBMPDetailDict.Exists("BasinWidthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinWidthIncr")
    End If
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "BWidthOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub cmdReleaseCurve_Click()
  On Error GoTo ErrorHandler

   Call RefreshChartingCurve

  Exit Sub
ErrorHandler:
  HandleError True, "cmdReleaseCurve_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


Private Sub RefreshChartingCurve()
  On Error GoTo ErrorHandler


    Dim pHr As Double
    Dim pStringArray(1 To 24) As String
    Dim i As Integer
    Dim pValue As Double
    Dim pMinValue As Double
    Dim pMaxValue As Double
    
    pStringArray(1) = txtHr1.Text
    pStringArray(2) = txtHr2.Text
    pStringArray(3) = txtHr3.Text
    pStringArray(4) = txtHr4.Text
    pStringArray(5) = txtHr5.Text
    pStringArray(6) = txtHr6.Text
    pStringArray(7) = txtHr7.Text
    pStringArray(8) = txtHr8.Text
    pStringArray(9) = txtHr9.Text
    pStringArray(10) = txtHr10.Text
    pStringArray(11) = txtHr11.Text
    pStringArray(12) = txtHr12.Text
    pStringArray(13) = txtHr13.Text
    pStringArray(14) = txtHr14.Text
    pStringArray(15) = txtHr15.Text
    pStringArray(16) = txtHr16.Text
    pStringArray(17) = txtHr17.Text
    pStringArray(18) = txtHr18.Text
    pStringArray(19) = txtHr19.Text
    pStringArray(20) = txtHr20.Text
    pStringArray(21) = txtHr21.Text
    pStringArray(22) = txtHr22.Text
    pStringArray(23) = txtHr23.Text
    pStringArray(24) = txtHr24.Text
    
    'Initialize min and max values
    pMaxValue = -1
    pMinValue = 9999
    
    For i = 1 To 24
        If (Not IsNumeric(pStringArray(i))) Then
            MsgBox "Value for Hour " & i & " is not valid."
            Exit Sub
        End If
        If (pValue > pMaxValue) Then
            pMaxValue = pValue
        End If
        If (pValue < pMinValue) Then
            pMinValue = pValue
        End If
        pValue = CDbl(pStringArray(i))
        ReleaseChart.DataGrid.SetData i, 1, pValue, 0 ' nullflag
        ReleaseChart.Row = i
        ReleaseChart.RowLabel = i
    Next
    
    With ReleaseChart.Plot.Axis(VtChAxisIdY)
      .CategoryScale.Auto = False
      .ValueScale.Maximum = pMaxValue * 1.1
      .ValueScale.Minimum = pMinValue
      .ValueScale.MajorDivision = 5 ' (pMaxValue - pMinValue) / 5
      .ValueScale.MinorDivision = 1 '(pMaxValue - pMinValue) / 5
    End With
    
    Dim olabel
    For Each olabel In ReleaseChart.Plot.Axis(VtChAxisIdX).Labels
        olabel.TextLayout.Orientation = VtOrientationVertical
    Next

    'Make the chart visible
    ReleaseChart.Visible = True
    

 

  Exit Sub
ErrorHandler:
  HandleError False, "RefreshChartingCurve " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

' Changed the design of cost module - Sabu Paul, September 2007
''Private Sub ConstCost_Change()
''  On Error GoTo ErrorHandler
''
''    If Not (IsNumeric(ConstCost.Text)) Then
''        MsgBox "Fixed Cost must be a valid number."
''        ConstCost.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "ConstCost_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

''Private Sub Da_Change()
''  On Error GoTo ErrorHandler
''
''    If Not (IsNumeric(Da.Text)) Then
''        MsgBox "Da must be a valid number."
''        Da.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "Da_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub
''
''Private Sub Db_Change()
''  On Error GoTo ErrorHandler
''
''    If Not (IsNumeric(Db.Text)) Then
''        MsgBox "Db must be a valid number."
''        Db.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "Db_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorHandler

    Call RefreshChartingCurve
    If (UnderDrainON.value = 0) Then
        StorageDepth.Enabled = False
        VoidFraction.Enabled = False
        BackgroundInfiltration.Enabled = False
    End If
    
    If SSTabBMP.TabVisible(gBMPDefTab) Then SSTabBMP.Tab = gBMPDefTab
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Activate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    'Initialize decay rate and removal rate for wq parameters
    InitializeDecayAndRemovalRates
        
    'Call to refresh the release chart
    Call RefreshChartingCurve

    'Set the cost components
    'Call Update_Component_List
    
    'Call SetCostAdjDict
    Set costAdjDict = New Scripting.Dictionary
    
    Call SetCostAdjDict(costAdjDict)
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pAdoConn.Close
    Set pAdoConn = Nothing
End Sub

Private Sub GreenAmptON_Click()
  On Error GoTo ErrorHandler
    If (GreenAmptON.value = 0) Then
        FrameGreenAmpt.Enabled = False
    Else
        FrameGreenAmpt.Enabled = True
    End If
  Exit Sub
ErrorHandler:
  HandleError True, "GreenAmptON_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
    
End Sub



Private Sub lstComponents_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'change mode from "save" to "edit"
    cmdEdit.Caption = "Edit"
End Sub

Private Sub NumUnitsOptimized_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "NumUnitsA"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "NumUnitsOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub NumUnitsOptimized2_Click()
  On Error GoTo ErrorHandler
    gCurOptParam = "NumUnitsA"
    frmOptimizer.OptimizerOnCheck.value = 1
    
    If gBMPDetailDict.Exists("MinNumUnits") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinNumUnits")
    End If
    If gBMPDetailDict.Exists("MaxNumUnits") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxNumUnits")
    End If
    If gBMPDetailDict.Exists("NumUnitsIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("NumUnitsIncr")
    End If

    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "NumUnitsOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub NumUnitsOptimizedB_Click()
  On Error GoTo ErrorHandler

    gCurOptParam = "NumUnitsB"
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "NumUnitsOptimizedB_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description

End Sub

Private Sub NumUnitsOptimizedB2_Click()
  On Error GoTo ErrorHandler
    gCurOptParam = "NumUnitsB"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinNumUnitsB") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinNumUnitsB")
    End If
    If gBMPDetailDict.Exists("MaxNumUnitsB") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxNumUnitsB")
    End If
    If gBMPDetailDict.Exists("NumUnitsIncrB") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("NumUnitsIncrB")
    End If
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "NumUnitsOptimizedB2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description

End Sub

' Changed the design of cost module - Sabu Paul, September 2007
''Private Sub LdCost_Change()
''  On Error GoTo ErrorHandler
''
''    If Not (IsNumeric(LdCost.Text)) Then
''        MsgBox "Land Cost must be a valid number."
''        LdCost.SetFocus
''    Else
''        Call UpdateCost
''    End If
''
''  Exit Sub
''ErrorHandler:
''  HandleError True, "LdCost_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
''End Sub

Private Sub SoilDepth_Change()
  On Error GoTo ErrorHandler

    If (SoilDepth.Text <> "") Then
        If IsNumeric(SoilDepth.Text) Then
            Call UpdateCost
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "SoilDepth_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub SoilPorosity_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler

    If Trim(SoilPorosity.Text) <> "" Then
        If IsNumeric(Trim(SoilPorosity.Text)) Then
            If Not (CDbl(Trim(SoilPorosity.Text)) >= 0# And CDbl(Trim(SoilPorosity.Text)) <= 1#) Then
                MsgBox "Soil Porosity should be between 0 and 1", vbExclamation
            End If
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "SoilPorosity_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub





Private Sub StorageDepth_Change()
  On Error GoTo ErrorHandler

    If (StorageDepth.Text <> "") Then
        If IsNumeric(StorageDepth.Text) Then
            Call UpdateCost
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "StorageDepth_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub




Private Sub UnderDrainON_Click()
  On Error GoTo ErrorHandler

    If (UnderDrainON.value = 0) Then
        StorageDepth.Enabled = False
        VoidFraction.Enabled = False
        BackgroundInfiltration.Enabled = False
    Else
        StorageDepth.Enabled = True
        VoidFraction.Enabled = True
        BackgroundInfiltration.Enabled = True
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "UnderDrainON_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub VoidFraction_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler

    If Trim(VoidFraction.Text) <> "" Then
        If IsNumeric(Trim(VoidFraction.Text)) Then
            If Not (CDbl(Trim(VoidFraction.Text)) >= 0# And CDbl(Trim(VoidFraction.Text)) <= 1#) Then
                MsgBox "Void Fraction should be between 0 and 1", vbExclamation
            End If
        Else
            MsgBox "Void Fraction should be a number", vbExclamation
        End If
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "VoidFraction_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub



Private Sub WeirType1_Click()
  On Error GoTo ErrorHandler

    BMPRectWeirWidth.Enabled = True
    BMPRectWeirWidth.BackColor = vbWhite
    BMPTriangularWeirAngle.Enabled = False
    BMPTriangularWeirAngle.BackColor = &H80000016

  Exit Sub
ErrorHandler:
  HandleError True, "WeirType1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub WeirType2_Click()
  On Error GoTo ErrorHandler

    BMPTriangularWeirAngle.Enabled = True
    BMPTriangularWeirAngle.BackColor = vbWhite
    BMPRectWeirWidth.Enabled = False
    BMPRectWeirWidth.BackColor = &H80000016

  Exit Sub
ErrorHandler:
  HandleError True, "WeirType2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Private Sub WHeightOptimized_Click()
  On Error GoTo ErrorHandler

    'WHeightOptimized = True
    gCurOptParam = "WHeight"
    frmOptimizer.Show vbModal


  Exit Sub
ErrorHandler:
  HandleError True, "WHeightOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub
Private Sub WHeightOptimized2_Click()
  On Error GoTo ErrorHandler

    'WHeightOptimized = True
    gCurOptParam = "WHeight"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinWeirHeight") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinWeirHeight")
    End If
    If gBMPDetailDict.Exists("MaxWeirHeight") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxWeirHeight")
    End If
    If gBMPDetailDict.Exists("WeirHeightIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("WeirHeightIncr")
    End If
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "WHeightOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub
Private Sub SoilDOptimized_Click()
  On Error GoTo ErrorHandler

    'SoilDOptimized = True
    gCurOptParam = "SoilD"
    frmOptimizer.Show vbModal


  Exit Sub
ErrorHandler:
  HandleError True, "SoilDOptimized_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub
Private Sub SoilDOptimized2_Click()
  On Error GoTo ErrorHandler

    'SoilDOptimized = True
    gCurOptParam = "SoilD"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinSoilDepth") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinSoilDepth")
    End If
    If gBMPDetailDict.Exists("MaxSoilDepth") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxSoilDepth")
    End If
    If gBMPDetailDict.Exists("SoilDepthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("SoilDepthIncr")
    End If
    frmOptimizer.Show vbModal

  Exit Sub
ErrorHandler:
  HandleError True, "SoilDOptimized2_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub


Private Function ValidateInputs() As Boolean

On Error GoTo ErrorHandler

    Dim pControl
    Dim pMessageStr As String
    pMessageStr = ""
    If SSTabBMP.TabVisible(0) = True Then
        If Not (OptionRelCistern.value Or OptionRelRainB.value) Then
            If Trim(BMPWidthA.Text) = "" Or Not (IsNumeric(Trim(BMPWidthA.Text))) Then
                pMessageStr = pMessageStr & "BMPWidth" & vbNewLine
            End If
        End If
        If Trim(BMPLengthA.Text) = "" Or Not (IsNumeric(Trim(BMPLengthA.Text))) Then
            pMessageStr = pMessageStr & "BMPLength" & vbNewLine
        End If
        If Trim(BMPOrificeHeight.Text) = "" Or Not (IsNumeric(Trim(BMPOrificeHeight.Text))) Then
            pMessageStr = pMessageStr & "BMPOrificeHeight" & vbNewLine
        End If
        If Trim(BMPOrificeDiameter.Text) = "" Or Not (IsNumeric(Trim(BMPOrificeDiameter.Text))) Then
            pMessageStr = pMessageStr & "BMPOrificeDiameter" & vbNewLine
        End If
        If OptionRelCistern.value = True Then
            If Trim(NumPeople.Text) = "" Or Not (IsNumeric(Trim(NumPeople.Text))) Then
                pMessageStr = pMessageStr & "NumPeople" & vbNewLine
            End If
        ElseIf OptionRelRainB.value = True Then
            If Trim(NumDays.Text) = "" Or Not (IsNumeric(Trim(NumDays.Text))) Then
                pMessageStr = pMessageStr & "NumDays" & vbNewLine
            End If
        End If
        If WeirType1.value = True Then
            If Trim(BMPRectWeirWidth.Text) = "" Or Not (IsNumeric(Trim(BMPRectWeirWidth.Text))) Then
                pMessageStr = pMessageStr & "BMPRectWeirWidth" & vbNewLine
            End If
        Else
            If Trim(BMPTriangularWeirAngle.Text) = "" Or Not (IsNumeric(Trim(BMPTriangularWeirAngle.Text))) Then
                pMessageStr = pMessageStr & "BMPTriangularWeirAngle" & vbNewLine
            End If
            'Mira Chokshi - added on 03/30/2005
            If (IsNumeric(Trim(BMPTriangularWeirAngle.Text))) Then
                If (CDbl(BMPTriangularWeirAngle.Text <= 0) Or CDbl(BMPTriangularWeirAngle.Text >= 180)) Then
                    pMessageStr = pMessageStr & "BMPTriangularWeirAngle: Greater than 0 And Less than 180)"
                End If
            End If
        End If
        If Trim(BMPWeirHeight.Text) = "" Or Not (IsNumeric(Trim(BMPWeirHeight.Text))) Then
            pMessageStr = pMessageStr & "BMPWeirHeight" & vbNewLine
        End If
        
        If Trim(BMPUnitsA.Text) = "" Or Not (IsNumeric(Trim(BMPUnitsA.Text))) Then
            pMessageStr = pMessageStr & "Num of Units" & vbNewLine
        End If
        If Trim(BMPDrainAreaA.Text) = "" Or Not (IsNumeric(Trim(BMPDrainAreaA.Text))) Then
            pMessageStr = pMessageStr & "Drainage Area" & vbNewLine
        End If
    ElseIf SSTabBMP.TabVisible(1) = True Then
        If Trim(BMPWidthB.Text) = "" Or Not (IsNumeric(Trim(BMPWidthB.Text))) Then
            pMessageStr = pMessageStr & "BMPWidth" & vbNewLine
        End If
        If Trim(BMPLengthB.Text) = "" Or Not (IsNumeric(Trim(BMPLengthB.Text))) Then
            pMessageStr = pMessageStr & "BMPLength" & vbNewLine
        End If
        If Trim(BMPMaxDepth.Text) = "" Or Not (IsNumeric(Trim(BMPMaxDepth.Text))) Then
            pMessageStr = pMessageStr & "BMPMaxDepth" & vbNewLine
        End If
        If Trim(BMPSlope1.Text) = "" Or Not (IsNumeric(Trim(BMPSlope1.Text))) Then
            pMessageStr = pMessageStr & "BMPSlope1" & vbNewLine
        End If
        If Trim(BMPSlope2.Text) = "" Or Not (IsNumeric(Trim(BMPSlope2.Text))) Then
            pMessageStr = pMessageStr & "BMPSlope2" & vbNewLine
        End If
        If Trim(BMPSlope3.Text) = "" Or Not (IsNumeric(Trim(BMPSlope3.Text))) Then
            pMessageStr = pMessageStr & "BMPSlope3" & vbNewLine
        End If

        If Trim(BMPManningsN.Text) = "" Or Not (IsNumeric(Trim(BMPManningsN.Text))) Then
            pMessageStr = pMessageStr & "BMPManningsN" & vbNewLine
        ElseIf Not (CDbl(Trim(BMPManningsN.Text)) >= 0# And CDbl(Trim(BMPManningsN.Text)) <= 1#) Then
            pMessageStr = pMessageStr & "BMPManningsN" & vbNewLine
        End If
        If Trim(BMPUnitsB.Text) = "" Or Not (IsNumeric(Trim(BMPUnitsB.Text))) Then
            pMessageStr = pMessageStr & "Num of Units" & vbNewLine
        End If
        If Trim(BMPDrainAreaB.Text) = "" Or Not (IsNumeric(Trim(BMPDrainAreaB.Text))) Then
            pMessageStr = pMessageStr & "Drainage Area" & vbNewLine
        End If
    End If
    Dim pMonthIndex As Integer
    For Each pControl In Controls
        If TypeOf pControl Is TextBox Then
            If pControl.Enabled Then
                If InStr(1, pControl.name, "Month", vbTextCompare) > 0 Then
                    If Trim(pControl.Text) = "" Or Not (IsNumeric(Trim(pControl.Text))) Then
                        pMonthIndex = CInt(Replace(pControl.name, "Month", ""))
                        pMessageStr = pMessageStr & "Growth Index for month " & pMonthIndex & vbNewLine
                    End If
                End If
            End If
        End If
    Next pControl

    Dim incr As Long
    If SSTabBMP.TabVisible(4) = True Then 'Sabu Paul -- March 31, 2005 Tab index modified
        For incr = 1 To 24
            For Each pControl In Controls
            If TypeOf pControl Is TextBox Then
                If pControl.Enabled Then
                    If pControl.name = "txtHr" & incr Then
                         If Trim(pControl.Text) = "" Or Not (IsNumeric(Trim(pControl.Text))) Then
                             pMessageStr = pMessageStr & "Cistern Flow for month " & pMonthIndex & vbNewLine
                         End If
                    End If
                End If
            End If
            Next pControl
        Next incr
    End If
    If SSTabBMP.TabVisible(5) = True Then 'Sabu Paul -- March 31, 2005 Tab index modified
        For incr = 1 To 5
            For Each pControl In Controls
                If TypeOf pControl Is TextBox Then
                    If pControl.Enabled Then
                        If pControl.name = "Decay" & incr Then
                            If Trim(pControl.Text) = "" Or Not (IsNumeric(Trim(pControl.Text))) Then
                                pMessageStr = pMessageStr & "Decay factor for parameter " & incr & vbNewLine
                            End If
                        End If
                        If pControl.name = "PctRem" & incr Then
                            If Trim(pControl.Text) = "" Or Not (IsNumeric(Trim(pControl.Text))) Then
                                pMessageStr = pMessageStr & "Percent Removal for parameter " & incr & vbNewLine
                            End If
                        End If
                    End If
                End If
            Next pControl
        Next incr
    End If
    
    ' Changed the design of cost module - Sabu Paul, September 2007
''    If Trim(Aa.Text) = "" Or Not (IsNumeric(Trim(Aa.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param Aa" & vbNewLine
''    End If
''    If Trim(Ab.Text) = "" Or Not (IsNumeric(Trim(Ab.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param Ab" & vbNewLine
''    End If
''    If Trim(Da.Text) = "" Or Not (IsNumeric(Trim(Da.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param Da" & vbNewLine
''    End If
''    If Trim(Db.Text) = "" Or Not (IsNumeric(Trim(Db.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param Db" & vbNewLine
''    End If
''    If Trim(LdCost.Text) = "" Or Not (IsNumeric(Trim(LdCost.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param LdCost" & vbNewLine
''    End If
''    If Trim(ConstCost.Text) = "" Or Not (IsNumeric(Trim(ConstCost.Text))) Then
''        pMessageStr = pMessageStr & "Cost Param ConstCost" & vbNewLine
''    End If
    

    If SSTabBMP.TabVisible(2) = True Then
        If Trim(SoilPorosity.Text) = "" Or Not (IsNumeric(Trim(SoilPorosity.Text))) Then
            pMessageStr = pMessageStr & "SoilPorosity" & vbNewLine
        ElseIf Not (CDbl(Trim(SoilPorosity.Text)) >= 0# And CDbl(Trim(SoilPorosity.Text)) <= 1#) Then
                pMessageStr = pMessageStr & "SoilPorosity" & vbNewLine
        End If
        If Trim(SoilDepth.Text) = "" Or Not (IsNumeric(Trim(SoilDepth.Text))) Then
            pMessageStr = pMessageStr & "SoilDepth" & vbNewLine
        End If
    
        If Trim(VegetativeParam.Text) = "" Or Not (IsNumeric(Trim(VegetativeParam.Text))) Then
            pMessageStr = pMessageStr & "VegetativeParam" & vbNewLine
        End If
        If Trim(SoilLayerInfiltration.Text) = "" Or Not (IsNumeric(Trim(SoilLayerInfiltration.Text))) Then
            pMessageStr = pMessageStr & "SoilLayerInfiltration" & vbNewLine
        End If
        If Trim(StorageDepth.Text) = "" Or Not (IsNumeric(Trim(StorageDepth.Text))) Then
            pMessageStr = pMessageStr & "StorageDepth" & vbNewLine
        End If
        If Trim(VoidFraction.Text) = "" Or Not (IsNumeric(Trim(VoidFraction.Text))) Then
            pMessageStr = pMessageStr & "VoidFraction" & vbNewLine
        ElseIf Not (CDbl(Trim(VoidFraction.Text)) >= 0# And CDbl(Trim(VoidFraction.Text)) <= 1#) Then
            pMessageStr = pMessageStr & "VoidFraction" & vbNewLine
        End If
        If Trim(BackgroundInfiltration.Text) = "" Or Not (IsNumeric(Trim(BackgroundInfiltration.Text))) Then
            pMessageStr = pMessageStr & "BackgroundInfiltration" & vbNewLine
        End If
    End If
    
    'Check sediment porosity and fractions
'    sedParams = Array("Bed width", "Bed depth", "Porosity", "Sand fraction", "Silt fraction", "Clay fraction", "Sand effective diameter" _
'        , "Sand velocity", "Sand density", "Sand coeff", "Sand exponent", "Silt effective diameter", "Silt velocity", "Silt density" _
'        , "Deposition stress", "Scour stress", "Erodibility")
        
    Dim oRsSed As ADODB.Recordset
    Set oRsSed = DataGridSed.DataSource
    oRsSed.Move 2, adBookmarkFirst
    If oRsSed.Fields(1).value < 0 Or oRsSed.Fields(1).value > 1 Then
        pMessageStr = pMessageStr & "Sediment Porosity (0-1)" & vbNewLine
    End If
    
    Dim bedDepth As Double
    oRsSed.Move 1, adBookmarkFirst
    bedDepth = oRsSed.Fields(1).value
    
    Dim soilFraction As Double
    oRsSed.Move 3, adBookmarkFirst
    soilFraction = oRsSed.Fields(1).value
    oRsSed.MoveNext
    soilFraction = soilFraction + oRsSed.Fields(1).value
    oRsSed.MoveNext
    soilFraction = soilFraction + oRsSed.Fields(1).value
    
    If bedDepth <> 0 Then
        If soilFraction <> 1 Then pMessageStr = pMessageStr & "Sediment soil fraction (Sand, Silt, Clay) (sum <> 1.0)" & vbNewLine
    Else
        If Not (soilFraction = 1 Or soilFraction = 0) Then
            pMessageStr = pMessageStr & "Sediment soil fraction (Sand, Silt, Clay) (sum <> 0.0 or 1.0)" & vbNewLine
        End If
    End If
    Set oRsSed = Nothing
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

Public Sub InitializeSedimentParameters(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    Dim sedParams
    sedParams = Array("Bed width", "Bed depth", "Porosity", "Sand fraction", "Silt fraction", "Clay fraction", "Sand effective diameter" _
        , "Sand velocity", "Sand density", "Sand coefficient", "Sand exponent", "Silt effective diameter", "Silt velocity", "Silt density" _
        , "Silt Deposition stress", "Silt Scour stress", "Silt Erodibility" _
        , "Clay effective diameter", "Clay velocity", "Clay density", "Clay Deposition stress", "Clay Scour stress", "Clay Erodibility")
    
    Dim sedUnits
    sedUnits = Array("(ft)", "(ft)", "", "", "", "", "(in)", "(in/sec)", "(lb/ft³)", "", "", "(in)", "(in/sec)", "(lb/ft³)", "(lb/ft²)", "(lb/ft²)", "(lb/ft²)", "(in)", "(in/sec)", "(lb/ft³)", "(lb/ft²)", "(lb/ft²)", "(lb/ft²)")
    
    Dim sedValues
    sedValues = Array(0, 0, 0.5, 0, 0, 0, 0, 0, 2.65, 0, 0, 0, 0, 2.65, 10000000000#, 10000000000#, 0, 0, 0, 2.65, 10000000000#, 10000000000#, 0)
    
        
    Dim oRsSed As ADODB.Recordset
    Set oRsSed = New ADODB.Recordset
    oRsSed.Fields.Append "Parameter", adVarChar, 40
    oRsSed.Fields.Append "Value", adDouble
    oRsSed.CursorType = adOpenDynamic
    oRsSed.Open
    
    Dim sedIncr As Integer
    For sedIncr = 0 To UBound(sedParams)
        oRsSed.AddNew
        oRsSed.Fields(0).value = sedParams(sedIncr) & " " & sedUnits(sedIncr)
        
        If Not pBmpDetailDict Is Nothing Then
            If pBmpDetailDict.Exists(sedParams(sedIncr)) Then
                'oRsSed.Fields(1).value = CDbl(pBmpDetailDict.Item(sedParams(sedIncr)))
                sedValues(sedIncr) = pBmpDetailDict.Item(sedParams(sedIncr))
            End If
        End If
        oRsSed.Fields(1).value = CDbl(sedValues(sedIncr))
    Next
    
    Set DataGridSed.DataSource = oRsSed
    DataGridSed.ColumnHeaders = True
    DataGridSed.Columns(0).Caption = "Parameter"
    DataGridSed.Columns(0).Locked = True
    DataGridSed.Columns(0).Width = 3000
    DataGridSed.Columns(1).Caption = "Value"
    DataGridSed.Columns(1).Width = 1800
        
    Exit Sub
ShowError:
    MsgBox " Error in InitializeSedimentParameters : " & Err.description
End Sub

'* Initialize data grid with values
Private Sub InitializeDecayAndRemovalRates()
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
''    Dim iDecayFact As Long
''    iDecayFact = pTableDF.FindField("DECAY")
''
''    Dim iK As Long
''    iK = pTableDF.FindField("K")
''
''    Dim iC As Long
''    iC = pTableDF.FindField("C")
''
''    Dim iPctRemoval As Long
''    iPctRemoval = pTableDF.FindField("REMOVAL")
''
''    Dim pRow As iRow
''    Dim iR As Integer
''
''    'Add all pollutants and their decay factors
''    For iR = 1 To pTotalPollutants
''        Set pRow = pTableDF.CreateRow
''        pRow.value(iPollutant) = gParamInfos(iR - 1).name
''        pRow.value(iDecayFact) = CDbl(gParamInfos(iR - 1).Decay)
''        pRow.value(iPctRemoval) = CDbl(gParamInfos(iR - 1).PctRem)
''        pRow.value(iK) = CDbl(gParamInfos(iR - 1).K)
''        pRow.value(iC) = CDbl(gParamInfos(iR - 1).C)
''        pRow.Store
''    Next
''    Set pRow = Nothing
''    Set pTableDF = Nothing
''
''
''    Dim oConn As New ADODB.Connection
''    oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
''           "SourceType=DBF;" & _
''           "SourceDB=" & gMapTempFolder & ";" & _
''           "Exclusive=No"
''    'Note: Specify the filename in the SQL statement. For example:
''    Dim oRs As New ADODB.Recordset
''    oRs.CursorLocation = adUseClient
''    oRs.Open "Select * From TempTable.dbf", oConn, adOpenDynamic, adLockOptimistic, adCmdText

    Dim oRs As ADODB.Recordset
    Set oRs = New ADODB.Recordset
    oRs.Fields.Append "POLLUTANT", adVarChar, 50
    oRs.Fields.Append "DECAY", adDouble
    oRs.Fields.Append "K", adDouble
    oRs.Fields.Append "C", adDouble
    oRs.Fields.Append "REMOVAL", adDouble
    oRs.CursorType = adOpenDynamic
    oRs.Open
    
    'Add all pollutants and their decay factors
    Dim iR As Integer
    For iR = 1 To pTotalPollutants
        oRs.AddNew
        oRs.Fields(0).value = gParamInfos(iR - 1).name
        oRs.Fields(1).value = CDbl(gParamInfos(iR - 1).Decay)
        oRs.Fields(2).value = CDbl(gParamInfos(iR - 1).K)
        oRs.Fields(3).value = CDbl(gParamInfos(iR - 1).C)
        oRs.Fields(4).value = CDbl(gParamInfos(iR - 1).PctRem)
    Next

    '* Set datagrid value, header caption and width
    Set DataGridDECAY.DataSource = oRs
    DataGridDECAY.ColumnHeaders = True
    DataGridDECAY.Columns(0).Caption = "Pollutant"
    DataGridDECAY.Columns(0).Locked = True
    DataGridDECAY.Columns(0).Width = 2000
    DataGridDECAY.Columns(1).Caption = "Decay Factor (1/day)"
    DataGridDECAY.Columns(1).Width = 1800
    DataGridDECAY.Columns(2).Caption = "K (ft/hr)"
    DataGridDECAY.Columns(2).Width = 900
    DataGridDECAY.Columns(3).Caption = "C* (lb/ft³)"
    DataGridDECAY.Columns(3).Width = 900
    DataGridDECAY.Columns(4).Visible = False
    
    
    '* Set datagrid value, header caption and width
    Set DataGridREMOVAL.DataSource = oRs
    DataGridREMOVAL.ColumnHeaders = True
    DataGridREMOVAL.Columns(0).Caption = "Pollutant"
    DataGridREMOVAL.Columns(0).Locked = True
    DataGridREMOVAL.Columns(0).Width = 2000
    DataGridREMOVAL.Columns(1).Visible = False
    DataGridREMOVAL.Columns(2).Visible = False
    DataGridREMOVAL.Columns(3).Visible = False
    DataGridREMOVAL.Columns(4).Caption = "Removal Rate (0-1)"
    DataGridREMOVAL.Columns(4).Width = 1800
    
    GoTo CleanUp
ShowError:
    MsgBox "InitializeDecayAndRemovalRates: " & Err.description
CleanUp:

End Sub



Public Sub InitCostFromDB(BMPType As String, Optional componentsToExclude As String)
On Error GoTo ErrorHandler
    'txtSourceDetails.Text = "CALTRANS    BMP Retrofit Pilot Program  1999    California Department of Transportation; CALTRANS Environmental Program 1120 N. St., MS 27  Sacramento  CA  95814       Robert Bein, William Frost & Associates "
    Dim ConnStr As String
    'ConnStr = "D:\SUSTAIN\CostDB\BMPCosts.mdb"
    If Trim(gCostDBpath) = "" Then Exit Sub
    ConnStr = Trim(gCostDBpath)
    
    Set pAdoConn = New ADODB.Connection
    pAdoConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnStr & ";"
    
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
           
    Dim strSqlComp As String
    strSqlComp = " SELECT DISTINCT Components.Components_ID, Components.Components_TXT " & _
                " FROM (BMPTypes INNER JOIN BMP_Components ON BMPTypes.BMPType_ID = BMP_Components.BMPType_ID) INNER JOIN Components ON BMP_Components.Component_ID = Components.Components_ID " & _
                " WHERE BMPTypes.BMPType_Code = '" & BMPType & "'"
    
'    strSqlComp = " SELECT DISTINCT Component_ID, Components " & _
'                " FROM Cost_Unit_Check " & _
'                " WHERE [BMP Type] = '" & BMPType & "'"
                                
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
    
    If Trim(txtCostExp.Text) <> "" Then
        If Not IsNumeric(txtCostExp.Text) Then
            MsgBox "Please enter a valid number for Cost Exponent", vbExclamation
            Exit Sub
        End If
    End If
    
''    If chkCCI.value <> 0 Then
''        Dim thisday As Date
''        thisday = Date$
''        Dim curYear As Integer, curIndex As Double
''        curIndex = 0#
''        Dim maxYear As Integer, maxIndex As Double
''        'curYear = Year(thisday)
''
''        'Cost is adjusted to last year of the index
''        Dim pRs As ADODB.Recordset
''        Set pRs = New ADODB.Recordset
''
''        Dim strSql As String
''        curYear = CInt(cbxYear.List(cbxYear.ListIndex))
''        strSql = "SELECT Year, Dec_CCI" & _
''                " From ConsolidatedDecCCI " & _
''                " WHERE Year = " & curYear & _
''                " AND ucase(Location)='NATIONAL'"
''        pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
''
''        pRs.MoveFirst
''        If Not pRs.EOF Then
''            curIndex = pRs("Dec_CCI")
''        End If
''
''        pRs.Close
''        strSql = "SELECT Year, Dec_CCI" & _
''                " From ConsolidatedDecCCI " & _
''                " WHERE Year =( Select Max(Year) from ConsolidatedDecCCI) " & _
''                " AND ucase(Location)='NATIONAL'"
''        pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
''
''        pRs.MoveFirst
''        If Not pRs.EOF Then
''            maxYear = pRs("Year")
''            maxIndex = pRs("Dec_CCI")
''        End If
''        pRs.Close
''        Set pRs = Nothing
''
''        'costAdj = 1.019 ^ (curYear - CDbl(cbxYear.List(cbxYear.ListIndex)))
''        If curIndex > 0 Then costAdj = maxIndex / curIndex
''    End If
    'curYear = CInt(cbxYear.List(cbxYear.ListIndex))
    curYear = CInt(cbxYear.Text)
    
    If UCase(cbxUnit.List(cbxUnit.ListIndex)) <> "PERCENTAGE" Then
        If costAdjDict.Exists(curYear) Then costAdj = costAdjDict.Item(curYear)
    End If
    
    Dim pCostVolType As Integer
    Dim pCostNumUnits As Integer
    
    pCostVolType = COST_VOLUME_TYPE_TOTAL
    If optVolMedia.value Then
        pCostVolType = COST_VOLUME_TYPE_MEDIA
    ElseIf optVolUnderDrain.value Then
        pCostVolType = COST_VOLUME_TYPE_UNDERDRAIN
    End If
    
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
    lstItem.ListSubItems.add , , cbxYear.Text 'cbxYear.List(cbxYear.ListIndex)
    lstItem.ListSubItems.add , , cbxUnit.List(cbxUnit.ListIndex)
    lstItem.ListSubItems.add , , pCostVolType
    lstItem.ListSubItems.add , , pCostNumUnits
    lstItem.ListSubItems.add , , txtCost.Text
    lstItem.ListSubItems.add , , CStr(FormatNumber(CDbl(txtCost.Text) * costAdj, 2))
    
    Dim costExp As String
    If Trim(txtCostExp.Text) <> "" Then
        costExp = Trim(txtCostExp.Text)
    Else
        costExp = "1"
    End If
    lstItem.ListSubItems.add , , costExp
    
    'cbxComponent.RemoveItem cbxComponent.ListIndex
'    If cbxComponent.ListCount > 0 Then
'        cbxComponent.ListIndex = 0
'    Else
'        cbxLocation.Clear
'        cbxSource.Clear
'        cbxYear.Clear
'        cbxUnit.Clear
'        txtCost.Text = ""
'        txtSourceDetails.Text = ""
'    End If

  Exit Sub
ErrorHandler:
  HandleError True, "cmdAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

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
'    cbxComponent.AddItem pComponent
'    cbxComponent.ItemData(cbxComponent.NewIndex) = pComponent_Id
'    cbxComponent.ListIndex = cbxComponent.NewIndex

  Exit Sub
ErrorHandler:
  HandleError True, "cmdRemove_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
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
    Dim costExp As String
    
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
                
                If Trim(txtCostExp.Text) <> "" Then
                    If Not IsNumeric(txtCostExp.Text) Then
                        MsgBox "Please enter a valid number for Cost Exponent", vbExclamation
                        Exit Sub
                    End If
                End If
    
                If txtUserComponent.Text <> "" Then
                    lstComponents.ListItems.Item(pIndex).ListSubItems(1).Text = txtUserComponent.Text
                Else
                    lstComponents.ListItems.Item(pIndex).ListSubItems(1).Text = cbxComponent.Text
                End If
                
                pCostVolType = COST_VOLUME_TYPE_TOTAL
                If optVolMedia.value Then
                    pCostVolType = COST_VOLUME_TYPE_MEDIA
                ElseIf optVolUnderDrain.value Then
                    pCostVolType = COST_VOLUME_TYPE_UNDERDRAIN
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
                
                If Trim(txtCostExp.Text) <> "" Then
                    costExp = Trim(txtCostExp.Text)
                Else
                    costExp = "1"
                End If
                lstComponents.ListItems.Item(pIndex).ListSubItems(10).Text = costExp
    
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

                If lstComponents.ListItems.Item(pIndex).ListSubItems(10).Text <> "1" Then
                    txtCostExp.Text = lstComponents.ListItems.Item(pIndex).ListSubItems(10).Text
                Else
                    txtCostExp.Text = ""
                End If
                
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
Private Sub Populate_Cost_Units()
On Error GoTo ErrorHandler
   
    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
        
    Dim strSqlUnit As String
    strSqlUnit = "SELECT DISTINCT UnitType_Desc " & _
        " FROM UnitTypes"

    pRs.Open strSqlUnit, pAdoConn, adOpenDynamic, adLockOptimistic
       
    cbxUnit.Clear
    pRs.MoveFirst
    Do Until pRs.EOF
        cbxUnit.AddItem pRs("UnitType_Desc")
        pRs.MoveNext
    Loop
    cbxUnit.ListIndex = 0
    pRs.Close

  Exit Sub
ErrorHandler:
  HandleError True, "Populate_Cost_Units " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
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

Private Sub Remove_Component(curComponent As String)
On Error GoTo ErrorHandler
    Dim i As Integer
    For i = 0 To cbxComponent.ListCount - 1
        If UCase(cbxComponent.List(i)) = UCase(curComponent) Then
            cbxComponent.RemoveItem i
            If cbxComponent.ListCount > 0 Then cbxComponent.ListIndex = 0
            Exit For
        End If
    Next
    Exit Sub
ErrorHandler:
  HandleError True, "Remove_Component " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub

Public Sub Update_Component_List(pBmpDetailDict As Scripting.Dictionary)
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
    Dim costExponents
    
    Dim lstItem As ListItem
    Dim pIndex As Integer
    
    'First clear the current content
    lstComponents.ListItems.Clear
    
    If pBmpDetailDict Is Nothing Then Exit Sub
    
    If pBmpDetailDict.Exists("CostComponents") Then
        costComps = Split(pBmpDetailDict.Item("CostComponents"), ";", , vbTextCompare)
        costCompIds = Split(pBmpDetailDict.Item("CostComponentIds"), ";", , vbTextCompare)
        costLocs = Split(pBmpDetailDict.Item("CostLocations"), ";", , vbTextCompare)
        costSrcs = Split(pBmpDetailDict.Item("CostSources"), ";", , vbTextCompare)
        costYears = Split(pBmpDetailDict.Item("CostYears"), ";", , vbTextCompare)
        costUnits = Split(pBmpDetailDict.Item("CostUnits"), ";", , vbTextCompare)
        costVolTypes = Split(pBmpDetailDict.Item("CostVolTypes"), ";", , vbTextCompare)
        costNumUnits = Split(pBmpDetailDict.Item("CostNumUnits"), ";", , vbTextCompare)
        costUnitCosts = Split(pBmpDetailDict.Item("CostUnitCosts"), ";", , vbTextCompare)
        costAdjUnitCosts = Split(pBmpDetailDict.Item("CostAdjUnitCosts"), ";")
        costExponents = Split(pBmpDetailDict.Item("costExponents"), ";")
        
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
            lstItem.ListSubItems.add , , costExponents(pIndex)
            'Remove_Component CStr(costComps(pIndex))
        Next
    End If
    Exit Sub
ErrorHandler:
  HandleError True, "Update_Component_List " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description
End Sub
