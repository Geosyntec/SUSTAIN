VERSION 5.00
Begin VB.Form FrmOptimizeCost 
   Caption         =   "Minimize Cost"
   ClientHeight    =   3645
   ClientLeft      =   4875
   ClientTop       =   4185
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "FrmOptimizeCost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5655
      Begin VB.TextBox NumBest 
         Height          =   375
         Left            =   4080
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Number of near optimal solutions for output"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   3120
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox StopDelta 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox MaxRunTime 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "$"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   400
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Stop the search and output near optimal solutions when the total cost has NOT been reduced by"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "hour"
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   1300
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Maximum search time allowed"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   3135
      End
   End
End
Attribute VB_Name = "FrmOptimizeCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    Dim pStopDelta As Double
    Dim pMaxRunTime As Double
    Dim pNumBest As Integer
    
    '*** Cost Limit i.e. Stop Delta
    If Not (IsNumeric(StopDelta)) Then
        MsgBox "Cost should be a numeric value."
        StopDelta.SetFocus
        Exit Sub
    End If
    If (CDbl(StopDelta) < 0) Then
        MsgBox "Cost should be a positive numeric value."
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
    DefineOptimizationMethod 1, 0, pStopDelta, pMaxRunTime, pNumBest
                        
        
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
