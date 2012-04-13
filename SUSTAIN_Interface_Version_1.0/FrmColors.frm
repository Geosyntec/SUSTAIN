VERSION 5.00
Begin VB.Form FrmColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Palletes"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   ClipControls    =   0   'False
   ForeColor       =   &H80000006&
   Icon            =   "FrmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Bar Chart Colors"
      Height          =   1935
      Left            =   120
      TabIndex        =   45
      Top             =   6480
      Width           =   8295
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   5
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   1
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton SeriesColors 
         BackColor       =   &H00FF8080&
         Height          =   375
         Index           =   0
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Series 6"
         Height          =   255
         Left            =   4080
         TabIndex        =   57
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label27 
         Caption         =   "Series 5"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label26 
         Caption         =   "Series 4"
         Height          =   255
         Left            =   4080
         TabIndex        =   53
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label25 
         Caption         =   "Series 3"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label24 
         Caption         =   "Series 2"
         Height          =   255
         Left            =   4080
         TabIndex        =   48
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label23 
         Caption         =   "Series 1"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lu Colors"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.CommandButton colors 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   6
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1950
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   15
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H0000C0C0&
         Height          =   375
         Index           =   13
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3375
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   8
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H0000FF00&
         Height          =   375
         Index           =   7
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00004000&
         Height          =   375
         Index           =   12
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3405
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00400040&
         Height          =   375
         Index           =   14
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3930
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   11
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2895
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Index           =   10
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   20
         Left            =   2760
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5325
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00004080&
         Height          =   375
         Index           =   18
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4845
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H008080FF&
         Height          =   375
         Index           =   9
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2415
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00000040&
         Height          =   375
         Index           =   21
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5325
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H000000C0&
         Height          =   375
         Index           =   19
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4845
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00004000&
         Height          =   375
         Index           =   17
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4350
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00C0C000&
         Height          =   375
         Index           =   16
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4365
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Index           =   4
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1350
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   855
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00808080&
         Height          =   375
         Index           =   1
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   0
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00004040&
         Height          =   375
         Index           =   5
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1335
         Width           =   975
      End
      Begin VB.CommandButton colors 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Open Urban Land"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1995
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Cropland"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   1995
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Pasture"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2490
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Orchards/Vine yard/Horticul"
         Height          =   255
         Left            =   4200
         TabIndex        =   41
         Top             =   2490
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Urban Herbaceous"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2985
         Width           =   2415
      End
      Begin VB.Label Label12 
         Caption         =   "Evergreen Forest"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Brush"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3975
         Width           =   2415
      End
      Begin VB.Label Label14 
         Caption         =   "Deciduous Forest"
         Height          =   255
         Left            =   4200
         TabIndex        =   37
         Top             =   2985
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Mixed Forest"
         Height          =   255
         Left            =   4200
         TabIndex        =   36
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Water"
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   3975
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Wetlands"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4455
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Bare Ground"
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   4455
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Extractive"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4950
         Width           =   2415
      End
      Begin VB.Label Label20 
         Caption         =   "Highway Corridors"
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   4950
         Width           =   1935
      End
      Begin VB.Label Label21 
         Caption         =   "Railroad Corridors"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5445
         Width           =   2415
      End
      Begin VB.Label Label22 
         Caption         =   "Agricultural Buildings"
         Height          =   255
         Left            =   4200
         TabIndex        =   29
         Top             =   5445
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Low Density Residential"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   420
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Medium Density Residential"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "High Density Residential"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   915
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Commercial"
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   915
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Industrial"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1410
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Institutional"
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   1410
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
