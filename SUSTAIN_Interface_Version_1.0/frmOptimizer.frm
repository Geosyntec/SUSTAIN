VERSION 5.00
Begin VB.Form frmOptimizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimization Parameters"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptimizer.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox OptimizerOnCheck 
      Caption         =   "Consider As Decision Variable"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "If checked the dimension will be selected as a decision variable during optimization"
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtMinValue 
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Minimum value of the decision variable"
      Top             =   480
      Width           =   1320
   End
   Begin VB.TextBox txtMaxValue 
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Maximum value of the decision variable"
      Top             =   1020
      Width           =   1320
   End
   Begin VB.TextBox txtIncrValue 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Increment value of the decision variable"
      Top             =   1560
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   850
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   850
   End
   Begin VB.Label Label1 
      Caption         =   "Minimum"
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum"
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Increment"
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   960
   End
End
Attribute VB_Name = "frmOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim pMin As Double
    Dim pMax As Double
    Dim pIncr As Double
    
    If OptimizerOnCheck.value = vbUnchecked Then
        Select Case UCase(gCurOptParam)
            Case "BLENGTH":
                FrmBMPData.BLengthOptimized2.Visible = False
                FrmBMPData.BLengthOptimized.Visible = True
                
                If gBMPDetailDict.Exists("BLengthOptimized") Then
                    gBMPDetailDict.Remove ("BLengthOptimized")
                End If
                If gBMPDetailDict.Exists("MinBasinLength") Then
                    gBMPDetailDict.Remove ("MinBasinLength")
                End If
                If gBMPDetailDict.Exists("MaxBasinLength") Then
                    gBMPDetailDict.Remove ("MaxBasinLength")
                End If
                If gBMPDetailDict.Exists("BasinLengthIncr") Then
                    gBMPDetailDict.Remove ("BasinLengthIncr")
                End If

            Case "BWIDTH":
                FrmBMPData.BWidthOptimized2.Visible = False
                FrmBMPData.BWidthOptimized.Visible = True
                If gBMPDetailDict.Exists("BWidthOptimized") Then
                    gBMPDetailDict.Remove ("BWidthOptimized")
                End If
                If gBMPDetailDict.Exists("MinBasinWidth") Then
                    gBMPDetailDict.Remove ("MinBasinWidth")
                End If
                If gBMPDetailDict.Exists("MaxBasinWidth") Then
                    gBMPDetailDict.Remove ("MaxBasinWidth")
                End If
                If gBMPDetailDict.Exists("BasinWidthIncr") Then
                    gBMPDetailDict.Remove ("BasinWidthIncr")
                End If
            Case "WHEIGHT":
                FrmBMPData.WHeightOptimized2.Visible = False
                FrmBMPData.WHeightOptimized.Visible = True
                
                If gBMPDetailDict.Exists("WHeightOptimized") Then
                    gBMPDetailDict.Remove ("WHeightOptimized")
                End If
                If gBMPDetailDict.Exists("MinWeirHeight") Then
                    gBMPDetailDict.Remove ("MinWeirHeight")
                End If
                If gBMPDetailDict.Exists("MaxWeirHeight") Then
                    gBMPDetailDict.Remove ("MaxWeirHeight")
                End If
                If gBMPDetailDict.Exists("WeirHeightIncr") Then
                    gBMPDetailDict.Remove ("WeirHeightIncr")
                End If
            Case "SOILD":
                FrmBMPData.SoilDOptimized2.Visible = False
                FrmBMPData.SoilDOptimized.Visible = True
                
                If gBMPDetailDict.Exists("SoilDOptimized") Then
                    gBMPDetailDict.Remove ("SoilDOptimized")
                End If
                If gBMPDetailDict.Exists("MinSoilDepth") Then
                    gBMPDetailDict.Remove ("MinSoilDepth")
                End If
                If gBMPDetailDict.Exists("MaxSoilDepth") Then
                    gBMPDetailDict.Remove ("MaxSoilDepth")
                End If
                If gBMPDetailDict.Exists("SoilDepthIncr") Then
                    gBMPDetailDict.Remove ("SoilDepthIncr")
                End If
            Case "BDEPTHB":
                FrmBMPData.BDepthBOptimized2.Visible = False
                FrmBMPData.BDepthBOptimized.Visible = True
                
                If gBMPDetailDict.Exists("BDepthBOptimized") Then
                    gBMPDetailDict.Remove ("BDepthBOptimized")
                End If
                If gBMPDetailDict.Exists("MinBasinBDepth") Then
                    gBMPDetailDict.Remove ("MinBasinBDepth")
                End If
                If gBMPDetailDict.Exists("MaxBasinBDepth") Then
                    gBMPDetailDict.Remove ("MaxBasinBDepth")
                End If
                If gBMPDetailDict.Exists("BasinBDepthIncr") Then
                    gBMPDetailDict.Remove ("BasinBDepthIncr")
                End If
            Case "BLENGTHB":
                FrmBMPData.BLengthBOptimized2.Visible = False
                FrmBMPData.BLengthBOptimized.Visible = True
                
                If gBMPDetailDict.Exists("BLengthBOptimized") Then
                    gBMPDetailDict.Remove ("BLengthBOptimized")
                End If
                If gBMPDetailDict.Exists("MinBasinBLength") Then
                    gBMPDetailDict.Remove ("MinBasinBLength")
                End If
                If gBMPDetailDict.Exists("MaxBasinBLength") Then
                    gBMPDetailDict.Remove ("MaxBasinBLength")
                End If
                If gBMPDetailDict.Exists("BasinBLengthIncr") Then
                    gBMPDetailDict.Remove ("BasinBLengthIncr")
                End If
            Case "BWIDTHB":
                FrmBMPData.BWidthBOptimized2.Visible = False
                FrmBMPData.BWidthBOptimized.Visible = True
                
                If gBMPDetailDict.Exists("BWidthBOptimized") Then
                    gBMPDetailDict.Remove ("BWidthBOptimized")
                End If
                If gBMPDetailDict.Exists("MinBasinBWidth") Then
                    gBMPDetailDict.Remove ("MinBasinBWidth")
                End If
                If gBMPDetailDict.Exists("MaxBasinBWidth") Then
                    gBMPDetailDict.Remove ("MaxBasinBWidth")
                End If
                If gBMPDetailDict.Exists("BasinBWidthIncr") Then
                    gBMPDetailDict.Remove ("BasinBWidthIncr")
                End If
            Case "NUMUNITSA":
                FrmBMPData.NumUnitsOptimized2.Visible = False
                FrmBMPData.NumUnitsOptimized.Visible = True
                
                If gBMPDetailDict.Exists("NumUnitsOptimized") Then
                    gBMPDetailDict.Remove ("NumUnitsOptimized")
                End If
                If gBMPDetailDict.Exists("MinNumUnits") Then
                    gBMPDetailDict.Remove ("MinNumUnits")
                End If
                If gBMPDetailDict.Exists("MaxNumUnits") Then
                    gBMPDetailDict.Remove ("MaxNumUnits")
                End If
                If gBMPDetailDict.Exists("NumUnitsIncr") Then
                    gBMPDetailDict.Remove ("NumUnitsIncr")
                End If
            Case "NUMUNITSB":
                FrmBMPData.NumUnitsOptimizedB2.Visible = False
                FrmBMPData.NumUnitsOptimizedB.Visible = True
                
                If gBMPDetailDict.Exists("NumUnitsOptimizedB") Then
                    gBMPDetailDict.Remove ("NumUnitsOptimizedB")
                End If
                If gBMPDetailDict.Exists("MinNumUnitsB") Then
                    gBMPDetailDict.Remove ("MinNumUnitsB")
                End If
                If gBMPDetailDict.Exists("MaxNumUnitsB") Then
                    gBMPDetailDict.Remove ("MaxNumUnitsB")
                End If
                If gBMPDetailDict.Exists("NumUnitsIncrB") Then
                    gBMPDetailDict.Remove ("NumUnitsIncrB")
                End If
        End Select
        Unload Me
        Exit Sub
    End If
    
    If Trim(txtMinValue.Text) = "" Or Trim(txtMaxValue.Text) = "" Or Trim(txtIncrValue.Text) = "" _
            Or Not (IsNumeric(Trim(txtMinValue.Text))) Or Not (IsNumeric(Trim(txtMaxValue.Text))) _
            Or Not (IsNumeric(Trim(txtIncrValue.Text))) Then
        MsgBox "Enter valid numbers for the parameters", vbExclamation
        Exit Sub
    End If

    pMin = CDbl(Trim(txtMinValue.Text))
    pMax = CDbl(Trim(txtMaxValue.Text))
    pIncr = CDbl(Trim(txtIncrValue.Text))
    
    If pMin >= pMax Then
        MsgBox "Minimun should be less than maximum value", vbExclamation
        Exit Sub
    End If
    If pIncr > (pMax - pMin) Then
        MsgBox "Increment should be smaller than the range", vbExclamation
        Exit Sub
    End If
    
    Select Case UCase(gCurOptParam)
    Case "BLENGTH":
        FrmBMPData.BLengthOptimized.Visible = False
        FrmBMPData.BLengthOptimized2.Visible = True
        gBMPDetailDict.Item("BLengthOptimized") = "True"
        gBMPDetailDict.Item("MinBasinLength") = txtMinValue.Text
        gBMPDetailDict.Item("MaxBasinLength") = txtMaxValue.Text
        gBMPDetailDict.Item("BasinLengthIncr") = txtIncrValue.Text
     
    Case "BWIDTH":
        FrmBMPData.BWidthOptimized.Visible = False
        FrmBMPData.BWidthOptimized2.Visible = True
        gBMPDetailDict.Item("BWidthOptimized") = "True"
        gBMPDetailDict.Item("MinBasinWidth") = txtMinValue.Text
        gBMPDetailDict.Item("MaxBasinWidth") = txtMaxValue.Text
        gBMPDetailDict.Item("BasinWidthIncr") = txtIncrValue.Text
        
    Case "WHEIGHT":
        FrmBMPData.WHeightOptimized.Visible = False
        FrmBMPData.WHeightOptimized2.Visible = True
        gBMPDetailDict.Item("WHeightOptimized") = "True"
        gBMPDetailDict.Item("MinWeirHeight") = txtMinValue.Text
        gBMPDetailDict.Item("MaxWeirHeight") = txtMaxValue.Text
        gBMPDetailDict.Item("WeirHeightIncr") = txtIncrValue.Text
        
    Case "SOILD":
        FrmBMPData.SoilDOptimized.Visible = False
        FrmBMPData.SoilDOptimized2.Visible = True
        gBMPDetailDict.Item("SoilDOptimized") = "True"
        gBMPDetailDict.Item("MinSoilDepth") = txtMinValue.Text
        gBMPDetailDict.Item("MaxSoilDepth") = txtMaxValue.Text
        gBMPDetailDict.Item("SoilDepthIncr") = txtIncrValue.Text
        
    Case "BDEPTHB":
        FrmBMPData.BDepthBOptimized.Visible = False
        FrmBMPData.BDepthBOptimized2.Visible = True
        gBMPDetailDict.Item("BDepthBOptimized") = "True"
        gBMPDetailDict.Item("MinBasinBDepth") = txtMinValue.Text
        gBMPDetailDict.Item("MaxBasinBDepth") = txtMaxValue.Text
        gBMPDetailDict.Item("BasinBDepthIncr") = txtIncrValue.Text
    Case "BLENGTHB":
        FrmBMPData.BLengthBOptimized.Visible = False
        FrmBMPData.BLengthBOptimized2.Visible = True
        gBMPDetailDict.Item("BLengthBOptimized") = "True"
        gBMPDetailDict.Item("MinBasinBLength") = txtMinValue.Text
        gBMPDetailDict.Item("MaxBasinBLength") = txtMaxValue.Text
        gBMPDetailDict.Item("BasinBLengthIncr") = txtIncrValue.Text
    Case "BWIDTHB":
        FrmBMPData.BWidthBOptimized.Visible = False
        FrmBMPData.BWidthBOptimized2.Visible = True
        gBMPDetailDict.Item("BWidthBOptimized") = "True"
        gBMPDetailDict.Item("MinBasinBWidth") = txtMinValue.Text
        gBMPDetailDict.Item("MaxBasinBWidth") = txtMaxValue.Text
        gBMPDetailDict.Item("BasinBWidthIncr") = txtIncrValue.Text
    Case "NUMUNITSA":
        FrmBMPData.NumUnitsOptimized.Visible = False
        FrmBMPData.NumUnitsOptimized2.Visible = True
        gBMPDetailDict.Item("NumUnitsOptimized") = "True"
        gBMPDetailDict.Item("MinNumUnits") = txtMinValue.Text
        gBMPDetailDict.Item("MaxNumUnits") = txtMaxValue.Text
        gBMPDetailDict.Item("NumUnitsIncr") = txtIncrValue.Text
    Case "NUMUNITSB":
        FrmBMPData.NumUnitsOptimizedB.Visible = False
        FrmBMPData.NumUnitsOptimizedB2.Visible = True
        gBMPDetailDict.Item("NumUnitsOptimizedB") = "True"
        gBMPDetailDict.Item("MinNumUnitsB") = txtMinValue.Text
        gBMPDetailDict.Item("MaxNumUnitsB") = txtMaxValue.Text
        gBMPDetailDict.Item("NumUnitsIncrB") = txtIncrValue.Text
    End Select
    Unload Me
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    OptimizerOnCheck.value = 1
End Sub


