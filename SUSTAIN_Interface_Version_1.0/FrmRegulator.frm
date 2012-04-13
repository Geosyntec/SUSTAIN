VERSION 5.00
Begin VB.Form FrmRegulator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regulator"
   ClientHeight    =   8115
   ClientLeft      =   2940
   ClientTop       =   2175
   ClientWidth     =   9780
   Icon            =   "FrmRegulator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   9780
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8520
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   8520
      TabIndex        =   31
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Weir Configuration"
      Height          =   2880
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   9480
      Begin VB.Frame frameTriangularWeir 
         Caption         =   "Triangular Weir"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6000
         TabIndex        =   29
         Top             =   1920
         Width           =   3255
         Begin VB.TextBox BMPTriangularWeirAngle 
            Enabled         =   0   'False
            Height          =   360
            Left            =   1920
            TabIndex        =   14
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblTriangularWeir 
            Caption         =   "Vertex Angle (theta, deg)"
            Enabled         =   0   'False
            Height          =   240
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1800
         End
      End
      Begin VB.Frame frameRectangularWeir 
         Caption         =   "Rectangular Weir"
         Height          =   855
         Left            =   6000
         TabIndex        =   27
         Top             =   840
         Width           =   3255
         Begin VB.TextBox BMPRectWeirWidth 
            Height          =   360
            Left            =   1920
            TabIndex        =   13
            Text            =   "3"
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblRectangularWeir 
            Caption         =   "Weir Crest Width (B, ft)"
            Height          =   360
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1920
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Weir Type"
         Height          =   2385
         Left            =   150
         TabIndex        =   24
         Top             =   240
         Width           =   5640
         Begin VB.OptionButton WeirType 
            Height          =   240
            Index           =   1
            Left            =   1680
            TabIndex        =   10
            Top             =   2040
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton WeirType 
            Height          =   240
            Index           =   2
            Left            =   4200
            TabIndex        =   11
            Top             =   2040
            Width           =   255
         End
         Begin VB.Image imgBmpa3 
            Height          =   1800
            Left            =   120
            Picture         =   "FrmRegulator.frx":08CA
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5160
         End
      End
      Begin VB.TextBox BMPWeirHeight 
         Height          =   360
         Left            =   7920
         TabIndex        =   12
         Text            =   "3.5"
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label7 
         Caption         =   "Weir Height (Hw, ft)"
         Height          =   360
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orifice Configuration"
      Height          =   3000
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   8625
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         Picture         =   "FrmRegulator.frx":1A4F0
         ScaleHeight     =   2415
         ScaleWidth      =   3375
         TabIndex        =   26
         Top             =   360
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exit Type"
         Height          =   2040
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Width           =   4680
         Begin VB.OptionButton OrificeExitType 
            Height          =   270
            Index           =   1
            Left            =   270
            TabIndex        =   4
            Top             =   1680
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton OrificeExitType 
            Height          =   270
            Index           =   3
            Left            =   2729
            TabIndex        =   6
            Top             =   1680
            Width           =   375
         End
         Begin VB.OptionButton OrificeExitType 
            Height          =   270
            Index           =   4
            Left            =   3929
            TabIndex        =   7
            Top             =   1680
            Width           =   375
         End
         Begin VB.OptionButton OrificeExitType 
            Height          =   270
            Index           =   2
            Left            =   1590
            TabIndex        =   5
            Top             =   1680
            Width           =   375
         End
         Begin VB.Image imgBmpa2 
            Height          =   1275
            Left            =   120
            Picture         =   "FrmRegulator.frx":355B2
            Top             =   240
            Width           =   4440
         End
      End
      Begin VB.TextBox BMPOrificeDiameter 
         Height          =   360
         Left            =   5040
         TabIndex        =   8
         Text            =   "15"
         Top             =   2520
         Width           =   720
      End
      Begin VB.TextBox BMPOrificeHeight 
         Height          =   360
         Left            =   7560
         TabIndex        =   9
         Text            =   "0"
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "Orifice Diameter (in)"
         Height          =   360
         Left            =   3600
         TabIndex        =   22
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "Orifice Height (Ho, ft)"
         Height          =   360
         Left            =   6000
         TabIndex        =   21
         Top             =   2520
         Width           =   1560
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Basin Dimensions"
      Height          =   720
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   6360
      Begin VB.TextBox BMPLength 
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Text            =   "5"
         Top             =   260
         Width           =   1095
      End
      Begin VB.TextBox BMPWidth 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Text            =   "5"
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   " Length (ft)"
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   320
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Width (ft)"
         Height          =   240
         Left            =   3840
         TabIndex        =   17
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Information"
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6360
      Begin VB.TextBox BMPName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Text            =   "Regulator"
         Top             =   240
         Width           =   5025
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   360
         Left            =   270
         TabIndex        =   15
         Top             =   240
         Width           =   2161
      End
   End
End
Attribute VB_Name = "FrmRegulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BLengthOptimized_Click()
    gCurOptParam = "BLength"
    frmOptimizer.Show vbModal
End Sub

Private Sub BLengthOptimized2_Click()
    'BLengthOptimized = True
    gCurOptParam = "BLength"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinLength") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinLength")
    End If
    If gBMPDetailDict.Exists("MaxBasinLength") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinLength")
    End If
    If gBMPDetailDict.Exists("BasinLengthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinLengthIncr")
    End If

    frmOptimizer.Show vbModal
End Sub

Private Sub BWidthOptimized_Click()
    'BWidthOptimized = True
    gCurOptParam = "BWidth"
    frmOptimizer.Show vbModal

End Sub

Private Sub BWidthOptimized2_Click()
    'BLengthOptimized = True
    gCurOptParam = "BWidth"
    'Set frmOptimizer values
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinBasinWidth") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinBasinWidth")
    End If
    If gBMPDetailDict.Exists("MaxBasinWidth") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxBasinWidth")
    End If
    If gBMPDetailDict.Exists("BasinWidthIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("BasinWidthIncr")
    End If
    frmOptimizer.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Set gBMPDetailDict = Nothing
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '*** ERROR CHECKING ROUTINE **************************************
    Dim pBMPName As String  '*** NAME
    pBMPName = Trim(FrmRegulator.BMPName.Text)
    If (pBMPName = "") Then
        MsgBox "Enter Regulator Name."
        Exit Sub
    End If
    
    Dim pBMPLength As Double    '*** LENGTH
    Dim pStrBMPLength As String
    pStrBMPLength = Trim(FrmRegulator.BMPLength.Text)
    If Not IsNumeric(pStrBMPLength) Then
        MsgBox "BMP Length should be a valid dimension."
        Exit Sub
    End If
    If (CDbl(pStrBMPLength) <= 0) Then
        MsgBox "BMP Length should be a positive dimension."
        Exit Sub
    End If
    pBMPLength = CDbl(pStrBMPLength)
    
    Dim pBMPWidth As Double     '*** WIDTH
    Dim pStrBMPWidth As String
    pStrBMPWidth = Trim(FrmRegulator.BMPWidth.Text)
    If Not IsNumeric(pStrBMPWidth) Then
        MsgBox "BMP Length should be a valid dimension."
        Exit Sub
    End If
    If (CDbl(pStrBMPWidth) <= 0) Then
        MsgBox "BMP Width should be a positive dimension."
        Exit Sub
    End If
    pBMPWidth = CDbl(pStrBMPWidth)
    
    Dim pExitType As Integer    '*** ORIFICE EXIT TYPE
    Dim iC As Integer
    For iC = 1 To FrmRegulator.OrificeExitType.Count
        If (FrmRegulator.OrificeExitType.Item(iC).value = True) Then
            pExitType = iC
        End If
    Next
    
    Dim pOrificeCoef As Double  '*** ORIFICE COEFFICIENT
    Select Case pExitType
        Case 1:
            pOrificeCoef = 1#
        Case 2:
            pOrificeCoef = 0.61
        Case 3:
            pOrificeCoef = 0.61
        Case 4:
            pOrificeCoef = 0.5
    End Select
    
    
    Dim pOrificeDia As Double   '*** ORIFICE DIAMETER
    Dim pStrOrificeDia As String
    pStrOrificeDia = Trim(FrmRegulator.BMPOrificeDiameter.Text)
    If Not IsNumeric(pStrOrificeDia) Then
        MsgBox "Orifice Diameter should be a valid dimension."
        Exit Sub
    End If
    If (CDbl(pStrOrificeDia) <= 0) Then
        MsgBox "Orifice Diameter should be a positive dimension."
        Exit Sub
    End If
    pOrificeDia = CDbl(pStrOrificeDia)
    
    Dim pOrificeHeight As Double  '*** ORIFICE HEIGHT
    Dim pStrOrificeHeight As String
    pStrOrificeHeight = Trim(FrmRegulator.BMPOrificeHeight.Text)
    If Not IsNumeric(pStrOrificeHeight) Then
        MsgBox "Orifice Height should be a valid dimension."
        Exit Sub
    End If
    If (CDbl(pStrOrificeHeight) < 0) Then
        MsgBox "Orifice Height should be a non-negative dimension."
        Exit Sub
    End If
    pOrificeHeight = CDbl(pStrOrificeHeight)
    
    Dim pWeirType As Double     '*** WEIR TYPE
    For iC = 1 To FrmRegulator.WeirType.Count
        If (FrmRegulator.WeirType.Item(iC).value = True) Then
            pWeirType = iC
        End If
    Next
    
    Dim pWeirHeight As Double    '*** WEIR HEIGHT
    Dim pStrWeirHeight As String
    pStrWeirHeight = Trim(FrmRegulator.BMPWeirHeight.Text)
    If Not IsNumeric(pStrWeirHeight) Then
        MsgBox "Weir Height should be a valid dimension."
        Exit Sub
    End If
    If (CDbl(pStrWeirHeight) <= 0) Then
        MsgBox "Weir Height should be a non-negative dimension."
        Exit Sub
    End If
    pWeirHeight = CDbl(pStrWeirHeight)
    
    Dim pWeirCrestWidth As Double
    Dim pStrWeirCrestWidth As String
    Dim pWeirVertexAngle As Double
    Dim pStrWeirVertexAngle As String
    
    If (pWeirType = 1) Then     'RECTANGULAR CROSS-SECTION
        pStrWeirCrestWidth = Trim(FrmRegulator.BMPRectWeirWidth.Text)
        If Not IsNumeric(pStrWeirCrestWidth) Then
            MsgBox "Rectangular Crest Width should be a valid dimension."
            Exit Sub
        End If
        If (CDbl(pStrWeirCrestWidth) <= 0) Then
            MsgBox "Rectangular Crest Width should be a positive dimension."
            Exit Sub
        End If
        pWeirCrestWidth = CDbl(pStrWeirCrestWidth)
    ElseIf (pWeirType = 2) Then     'TRIANGULAR CROSS-SECTION
        pStrWeirVertexAngle = Trim(FrmRegulator.BMPTriangularWeirAngle.Text)
        If Not IsNumeric(pStrWeirVertexAngle) Then
            MsgBox "Triangular Weir Angle should be a valid dimension."
            Exit Sub
        End If
        If (CDbl(pStrWeirVertexAngle) <= 0) Then
            MsgBox "Triangular Weir Angle should be a positive dimension."
            Exit Sub
        End If
        pWeirVertexAngle = CDbl(pStrWeirVertexAngle)
    End If
    
    '*** Since all input variables are validated, insert them in gBMPDetailDict
    Set gBMPDetailDict = Nothing
    Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
    gBMPDetailDict.add "BMPName", pBMPName
    gBMPDetailDict.add "BMPType", "Regulator"
    gBMPDetailDict.add "BMPClass", "A"
    gBMPDetailDict.add "BMPWidth", pBMPWidth
    gBMPDetailDict.add "BMPLength", pBMPLength
    gBMPDetailDict.add "BMPOrificeHeight", pOrificeHeight
    gBMPDetailDict.add "BMPOrificeDiameter", pOrificeDia
    gBMPDetailDict.add "OrificeExitType", pExitType  'add exit type
    gBMPDetailDict.add "OrificeCoef", pOrificeCoef
    gBMPDetailDict.add "ReleaseOption", "None"
    gBMPDetailDict.add "WeirType", pWeirType
    gBMPDetailDict.add "BMPWeirHeight", pWeirHeight
    
    If (pWeirType = 1) Then     'RECTANGULAR
        gBMPDetailDict.add "BMPRectWeirWidth", pWeirCrestWidth
    ElseIf (pWeirType = 2) Then     'TRIANGULAR
        gBMPDetailDict.add "BMPTriangularWeirAngle", pWeirVertexAngle
    End If
    
    gBMPDetailDict.add "isAssessmentPoint", False
    
    'CLOSE THE FORM
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub

Private Sub WeirType_Click(Index As Integer)
    If (Index = 1) Then     'RECTANGULAR WEIR C/S
        FrmRegulator.frameRectangularWeir.Enabled = True
        FrmRegulator.lblRectangularWeir.Enabled = True
        FrmRegulator.BMPRectWeirWidth.Enabled = True
        FrmRegulator.frameTriangularWeir.Enabled = False
        FrmRegulator.lblTriangularWeir.Enabled = False
        FrmRegulator.BMPTriangularWeirAngle.Enabled = False
    ElseIf (Index = 2) Then     'TRIANGULAR WEIR C/S
        FrmRegulator.frameRectangularWeir.Enabled = False
        FrmRegulator.lblRectangularWeir.Enabled = False
        FrmRegulator.BMPRectWeirWidth.Enabled = False
        FrmRegulator.frameTriangularWeir.Enabled = True
        FrmRegulator.lblTriangularWeir.Enabled = True
        FrmRegulator.BMPTriangularWeirAngle.Enabled = True
    End If
End Sub

Private Sub WHeightOptimized_Click()
    'WHeightOptimized = True
    gCurOptParam = "WHeight"
    frmOptimizer.Show vbModal
End Sub
Private Sub WHeightOptimized2_Click()
    'WHeightOptimized = True
    gCurOptParam = "WHeight"
    frmOptimizer.OptimizerOnCheck.value = 1
    If gBMPDetailDict.Exists("MinWeirHeight") Then
        frmOptimizer.txtMinValue.Text = gBMPDetailDict.Item("MinWeirHeight")
    End If
    If gBMPDetailDict.Exists("MaxWeirHeight") Then
        frmOptimizer.txtMaxValue.Text = gBMPDetailDict.Item("MaxWeirHeight")
    End If
    If gBMPDetailDict.Exists("WeirHeightIncr") Then
        frmOptimizer.txtIncrValue.Text = gBMPDetailDict.Item("WeirHeightIncr")
    End If
    frmOptimizer.Show vbModal
End Sub
