VERSION 5.00
Begin VB.Form FrmVFSSimOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define VFSMOD Solution Parameters"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "FrmVFSSimOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Kinematic Wave Numerical Solution Parameters"
      Height          =   3495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.TextBox N 
         Height          =   285
         Left            =   4800
         TabIndex        =   10
         Text            =   "99"
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox THETAW 
         Height          =   285
         Left            =   4800
         TabIndex        =   9
         Text            =   "0.5"
         ToolTipText     =   "0.5 is recommended"
         Top             =   840
         Width           =   1000
      End
      Begin VB.TextBox NPOL 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Text            =   "3"
         Top             =   1335
         Width           =   1000
      End
      Begin VB.TextBox CR 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Text            =   "0.6"
         ToolTipText     =   "Between 0.5 - 0.8"
         Top             =   2520
         Width           =   1000
      End
      Begin VB.TextBox MAXITER 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Text            =   "150"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.Frame Frame7 
         Caption         =   "Solution Method (KPG)"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   6135
         Begin VB.OptionButton optKPG0 
            Caption         =   "Regular Finite Element"
            Height          =   195
            Left            =   3120
            TabIndex        =   5
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton optKPG1 
            Caption         =   "Petrov-Galerkin solution"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Label Label14 
         Caption         =   "Number of Nodes in Solution Domain (N)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label15 
         Caption         =   "Time Weight factor (THETAW)"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label16 
         Caption         =   "Number of Element Nodal Points (NPOL)"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1335
         Width           =   3255
      End
      Begin VB.Label Label18 
         Caption         =   "Courant Number (CR)"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Maximum Iterations (MAXITER)"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3698
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2018
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "FrmVFSSimOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ShowError
    Dim N_num As Integer
    Dim THETAW_num As Double
    Dim NPOL_num As Integer
    Dim KPG_num As Integer
    Dim CR_num As Double
    Dim MAXITER_num As Integer
    
    Dim IELOUT As Integer
    IELOUT = 0
    
    If (Trim(N.Text) = "" Or Not IsNumeric(N.Text)) Then
        MsgBox "Please specify integer value for N."
        Exit Sub
    End If
    If (Trim(THETAW.Text) = "" Or Not IsNumeric(THETAW.Text)) Then
        MsgBox "Please specify real value for THETAW."
        Exit Sub
    End If
    If (Trim(NPOL.Text) = "" Or Not IsNumeric(NPOL.Text)) Then
        MsgBox "Please specify integer value for NPOL."
        Exit Sub
    End If
    If (Trim(CR.Text) = "" Or Not IsNumeric(CR.Text)) Then
        MsgBox "Please specify real value (05. to 0.8) for CR."
        Exit Sub
    End If

    If (Trim(MAXITER.Text Or Not IsNumeric(MAXITER.Text)) = "") Then
        MsgBox "Please specify integer value for MAXITER."
        Exit Sub
    End If
    
    N_num = CInt(Trim(N.Text))
    THETAW_num = CDbl(Trim(THETAW.Text))
    NPOL_num = CInt(Trim(NPOL.Text))
    CR_num = CDbl(Trim(CR.Text))
    MAXITER_num = CInt(Trim(MAXITER.Text))
    
    If optKPG0.value = True Then
        KPG_num = 0
    Else
        KPG_num = 1
    End If
    
    DefineVFSSimulationOptions N_num, THETAW_num, NPOL_num, KPG_num, CR_num, MAXITER_num, IELOUT
    
    Unload Me
    Exit Sub
    
ShowError:
    MsgBox "Error in setting VFS Simulation options (Form): " & Err.description
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
