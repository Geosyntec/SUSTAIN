VERSION 5.00
Begin VB.Form FrmErrors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projections"
   ClientHeight    =   3315
   ClientLeft      =   5760
   ClientTop       =   4170
   ClientWidth     =   5415
   Icon            =   "FrmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox ListProjections 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label lblErrors 
      Caption         =   "ERRORS"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "FrmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
