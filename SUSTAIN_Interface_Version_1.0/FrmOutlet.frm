VERSION 5.00
Begin VB.Form FrmOutlet 
   Caption         =   "Select Outlet Type:"
   ClientHeight    =   1185
   ClientLeft      =   7740
   ClientTop       =   6900
   ClientWidth     =   2955
   ClipControls    =   0   'False
   Icon            =   "FrmOutlet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2955
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.OptionButton UnderDrain 
      Caption         =   "Under Drain"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton Orifice 
      Caption         =   "Orifice"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Weir 
      Caption         =   "Weir"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmOutlet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    If (Weir.value = True) Then 'WEIR
        gBMPOutletType = 2
    ElseIf (Orifice.value = True) Then  'ORIFICE
        gBMPOutletType = 3
    ElseIf (UnderDrain.value = True) Then 'UNDERDRAIN
        gBMPOutletType = 4
    Else
        gBMPOutletType = -1
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
