VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1035
   ClientLeft      =   7665
   ClientTop       =   9735
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4305
      Begin VB.Label lblBMP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2130
         TabIndex        =   2
         Top             =   120
         Width           =   75
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   120
         Picture         =   "FrmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   675
         Width           =   4080
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Processing........ Please Wait !!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

