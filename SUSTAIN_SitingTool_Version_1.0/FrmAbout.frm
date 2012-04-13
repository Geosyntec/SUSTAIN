VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About BMP Siting Tool"
   ClientHeight    =   1875
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7740
   ClipControls    =   0   'False
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1294.159
   ScaleMode       =   0  'User
   ScaleWidth      =   7268.258
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      Picture         =   "FrmAbout.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   3000
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmAbout.frx":5A88
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BMP Siting Tool"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   4365
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version 1.0: Dated January 09, 2007"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

