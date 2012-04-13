VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExternalTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define External Time Series"
   ClientHeight    =   3630
   ClientLeft      =   4560
   ClientTop       =   5055
   ClientWidth     =   6555
   Icon            =   "frmExternalTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6555
   Begin VB.TextBox txtClayFrac 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Text            =   "0.0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtSiltFrac 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Text            =   "0.0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtSandFrac 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Text            =   "0.0"
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtBMP 
      Height          =   405
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtMultiplier 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "1"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdFindTimeSeries 
      Caption         =   "..."
      Height          =   360
      Left            =   5640
      TabIndex        =   1
      Top             =   1440
      Width           =   600
   End
   Begin VB.TextBox txtTimeSeriesFile 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   3240
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Clay Fraction  (0 -1)"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Silt Fraction  (0 -1)"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Sand Fraction (0 -1)"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Multiplier"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Select Time-Series file:"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmExternalTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindTimeSeries_Click()
    frmExternalTS.CommonDialog.ShowOpen
    txtTimeSeriesFile.Text = frmExternalTS.CommonDialog.FileName
End Sub

Private Sub cmdOk_Click()
    'Check if multiplier and time series file are specified
    Dim strMultiplier As String
    strMultiplier = txtMultiplier.Text
    If Not (IsNumeric(strMultiplier)) Then
        MsgBox "Multiplier should be a valid number"
        Exit Sub
    End If
    
    Dim dblMultiplier As Double
    dblMultiplier = CDbl(strMultiplier)
    If (dblMultiplier < 0) Then
        MsgBox "Multiplier should be a positive number."
        Exit Sub
    End If
    
    Dim strTimeSeries As String
    strTimeSeries = Trim(txtTimeSeriesFile.Text)
    If (strTimeSeries = "") Then
        MsgBox "Select external time series file."
        Exit Sub
    End If
    
    Dim strSandFrac As String
    strSandFrac = Trim(txtSandFrac.Text)
    If Not (IsNumeric(strSandFrac)) Then
        MsgBox "Sand Fraction should be a valid number"
        Exit Sub
    End If
    
    Dim strSiltFrac As String
    strSiltFrac = Trim(txtSiltFrac.Text)
    If Not (IsNumeric(strSiltFrac)) Then
        MsgBox "Silt Fraction should be a valid number"
        Exit Sub
    End If
    
    Dim strClayFrac As String
    strClayFrac = Trim(txtClayFrac.Text)
    If Not (IsNumeric(strClayFrac)) Then
        MsgBox "Clay Fraction should be a valid number"
        Exit Sub
    End If
    
    Dim strDescription As String
    strDescription = Trim(txtDescription)
        
    Dim bmpId As Integer
    bmpId = CInt(txtBMP.Text)
    
    Unload Me
    
    'Call subroutine to create table/add row for this bmp
'    AddExternalTimeSeriesForBMP BMPId, strDescription, dblMultiplier, strTimeSeries
    AddExternalTimeSeriesForBMP bmpId, strDescription, dblMultiplier, strTimeSeries, CDbl(strSandFrac), CDbl(strSiltFrac), CDbl(strClayFrac)
    
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
