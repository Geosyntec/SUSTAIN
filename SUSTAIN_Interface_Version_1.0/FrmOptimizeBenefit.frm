VERSION 5.00
Begin VB.Form FrmOptimizeBenefit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maximize Control Benefit"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "FrmOptimizeBenefit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Number of Near Optimal Solutions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox NumBest 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Number of near optimal solutions for output"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cost Limit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5655
      Begin VB.TextBox CostLimit 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "$"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Input cost limit"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Search Stopping Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
      Begin VB.TextBox MaxRunTime 
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox StopDelta 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Maximum search time allowed"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "hour"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Stop the search and output near optimal solutions when the control benefit has NOT been improved by"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmOptimizeBenefit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim pCostLimit As Double
    Dim pStopDelta As Double
    Dim pMaxRunTime As Double
    Dim pNumBest As Integer
    
    '*** Cost Limit
    If Not (IsNumeric(CostLimit)) Then
        MsgBox "Cost limit should be a numeric value."
        CostLimit.SetFocus
        Exit Sub
    End If
    If (CDbl(CostLimit) < 0) Then
        MsgBox "Cost limit should be a positive numeric value."
        CostLimit.SetFocus
        Exit Sub
    End If
    pCostLimit = CDbl(CostLimit)
    
    '*** Percentage Benefit Control i.e. Stop Delta
    If Not (IsNumeric(StopDelta)) Then
        MsgBox "Control benefit should be a numeric value."
        StopDelta.SetFocus
        Exit Sub
    End If
    If (CDbl(StopDelta) < 0 Or CDbl(StopDelta) > 100) Then
        MsgBox "Control benefit should be a percentage value between 0-100%."
        StopDelta.SetFocus
        Exit Sub
    End If
    pStopDelta = CDbl(StopDelta)
    
    '*** Max. run time
    If Not (IsNumeric(MaxRunTime)) Then
        MsgBox "Search time allowed should be a numeric value."
        MaxRunTime.SetFocus
        Exit Sub
    End If
    If (CDbl(MaxRunTime) < 0) Then
        MsgBox "Search time allowed should be a positive numeric value."
        MaxRunTime.SetFocus
        Exit Sub
    End If
    pMaxRunTime = CDbl(MaxRunTime)
    
    '*** Number of Best Solutions
    If Not (IsNumeric(NumBest)) Then
        MsgBox "Number of near optimal solutions should be a numeric value."
        NumBest.SetFocus
        Exit Sub
    End If
    If (CInt(NumBest) < 0) Then
        MsgBox "Number of near optimal solutions should be a positive numeric value."
        NumBest.SetFocus
        Exit Sub
    End If
    pNumBest = CInt(NumBest)
        
    'Close form
    Unload Me
    
    'Call the subroutine to update this option
    DefineOptimizationMethod 2, pCostLimit, pStopDelta, pMaxRunTime, pNumBest
            
   

End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
