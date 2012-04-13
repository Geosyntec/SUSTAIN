VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBufferStripTemplate 
   Caption         =   "Define Buffer Strip Template"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "frmBufferStripTemplate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Overland Flow"
      TabPicture(0)   =   "frmBufferStripTemplate.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Infiltration Soil Properties"
      TabPicture(1)   =   "frmBufferStripTemplate.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Buffer Vegetation Properties"
      TabPicture(2)   =   "frmBufferStripTemplate.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Buffer Dimensions"
         Height          =   1215
         Left            =   -74520
         TabIndex        =   3
         Top             =   720
         Width           =   5775
         Begin VB.TextBox txtBufferLength 
            Height          =   330
            Left            =   2640
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtBufferWidth 
            Height          =   330
            Left            =   2640
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtBufferSlope 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4680
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBufferRoughness 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4680
            TabIndex        =   4
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Buffer Length (m) [VL]"
            Height          =   330
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Width of the Strip (m) [FWIDTH]"
            Height          =   330
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label8 
            Caption         =   "Slope"
            Enabled         =   0   'False
            Height          =   330
            Left            =   3720
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Roughness"
            Enabled         =   0   'False
            Height          =   330
            Left            =   3720
            TabIndex        =   8
            Top             =   720
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "frmBufferStripTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

    '** Input Validation
    If (txtBufferLength.Text = "") Then
        MsgBox "Please specify Buffer Length"
        Exit Sub
    End If
    If (txtBufferWidth.Text = "") Then
        MsgBox "Please specify Buffer Width"
        Exit Sub
    End If
    
''    If (txtBufferRoughness.Text = "") Then
''        MsgBox "Please specify Buffer Roughness"
''        Exit Sub
''    End If
''    If (txtBufferSlope.Text = "") Then
''        MsgBox "Please specify Buffer Slope"
''        Exit Sub
''    End If
    
    '** create the dictionary
    Set gBufferStripDetailDict = CreateObject("Scripting.Dictionary")
    gBufferStripDetailDict.Add "BufferLength", txtBufferLength.Text
    gBufferStripDetailDict.Add "BufferWidth", txtBufferWidth.Text
    
        '** call the generic function to create and add rows for values
    ModuleVFSFunctions.SaveVFSPropertiesTable "BufferStripDefault", 0, gBufferStripDetailDict
        
    '** set it to nothing
    Set gBufferStripDetailDict = Nothing
    
    '** close the form
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    '** close the form
    Unload Me
End Sub


