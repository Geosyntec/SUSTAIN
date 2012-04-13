VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1200
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   5385
      Begin VB.Label lblProcess 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   120
         Picture         =   "FrmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   5160
      End
      Begin VB.Label lblProductName 
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
         Top             =   135
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

