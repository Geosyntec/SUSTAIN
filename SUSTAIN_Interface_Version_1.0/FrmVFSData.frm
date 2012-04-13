VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form FrmVFSData 
   Caption         =   "Define Buffer Strip Template"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   Icon            =   "FrmVFSData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Overland Flow"
      TabPicture(0)   =   "FrmVFSData.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtVFSID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Infiltration Properties"
      TabPicture(1)   =   "FrmVFSData.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Vegetation Properties"
      TabPicture(2)   =   "FrmVFSData.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.PictureBox Picture 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         Picture         =   "FrmVFSData.frx":091E
         ScaleHeight     =   2535
         ScaleWidth      =   3255
         TabIndex        =   12
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtVFSID 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Text            =   "HIDDEN"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Information"
         Height          =   855
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   4455
         Begin VB.TextBox txtName 
            Height          =   330
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buffer Dimensions"
         Height          =   1455
         Left            =   3720
         TabIndex        =   3
         Top             =   1680
         Width           =   4455
         Begin VB.TextBox txtBufferLength 
            Height          =   330
            Left            =   3240
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtBufferWidth 
            Height          =   330
            Left            =   3240
            TabIndex        =   4
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Buffer Length [VL]  (m) "
            Height          =   210
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Width of the Strip [FWIDTH] (m)"
            Height          =   210
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   2295
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "FrmVFSData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bContinue As Boolean
Private Sub cmdOk_Click()

    '** Input Validation
    If (Trim(txtName.Text) = "") Then
        MsgBox "Please specify buffer strip name to continue."
        Exit Sub
    End If
    If Not (IsNumeric(txtBufferLength.Text)) Then
        MsgBox "Buffer length value must be a valid number."
        Exit Sub
    End If
    If Not (IsNumeric(txtBufferWidth.Text)) Then
        MsgBox "Buffer width value must be a valid number."
        Exit Sub
    End If
      
    '** close the form
    bContinue = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    '** close the form
    bContinue = False
    Unload Me
End Sub


