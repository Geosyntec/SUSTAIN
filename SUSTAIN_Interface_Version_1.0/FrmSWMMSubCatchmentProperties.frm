VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSWMMSubCatchmentProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Subcatchment Properties"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "FrmSWMMSubCatchmentProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSubCatchment 
      Height          =   375
      Left            =   7200
      TabIndex        =   47
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   43
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7200
      TabIndex        =   42
      Top             =   240
      Width           =   855
   End
   Begin TabDlg.SSTab CatchmentTab 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FrmSWMMSubCatchmentProperties.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Additional"
      TabPicture(1)   =   "FrmSWMMSubCatchmentProperties.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Landuses"
      TabPicture(2)   =   "FrmSWMMSubCatchmentProperties.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Initial Buildup"
      TabPicture(3)   =   "FrmSWMMSubCatchmentProperties.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Groundwater"
      TabPicture(4)   =   "FrmSWMMSubCatchmentProperties.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "Groundwater Flow Editor"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   6495
         Begin VB.TextBox txtThershold 
            Height          =   315
            Left            =   4920
            TabIndex        =   71
            Text            =   "0"
            Top             =   2340
            Width           =   1335
         End
         Begin VB.TextBox txtSurDep 
            Height          =   315
            Left            =   4920
            TabIndex        =   69
            Text            =   "0"
            Top             =   1860
            Width           =   1335
         End
         Begin VB.TextBox txtSurGW 
            Height          =   315
            Left            =   4920
            TabIndex        =   67
            Text            =   "0"
            Top             =   1380
            Width           =   1335
         End
         Begin VB.TextBox txtSurExp 
            Height          =   315
            Left            =   4920
            TabIndex        =   65
            Text            =   "0"
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox txtSurCoeff 
            Height          =   315
            Left            =   4920
            TabIndex        =   63
            Text            =   "0"
            Top             =   375
            Width           =   1335
         End
         Begin VB.TextBox txtGWExp 
            Height          =   315
            Left            =   1920
            TabIndex        =   61
            Text            =   "0"
            Top             =   2295
            Width           =   1335
         End
         Begin VB.ComboBox cmbAquifer 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSurElev 
            Height          =   315
            Left            =   1920
            TabIndex        =   55
            Text            =   "0"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtReceiveNode 
            Height          =   315
            Left            =   1920
            TabIndex        =   54
            Text            =   "9"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtGWCoeff 
            Height          =   315
            Left            =   1920
            TabIndex        =   53
            Text            =   "0"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "Thershold Groundwater Elev. (ft)"
            Height          =   615
            Left            =   3480
            TabIndex        =   72
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Fixed Surface Water Depth (ft)"
            Height          =   375
            Left            =   3480
            TabIndex        =   70
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label29 
            Caption         =   "Surface-GW Interaction Coeff."
            Height          =   375
            Left            =   3480
            TabIndex        =   68
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Surface Water Flow Expon."
            Height          =   375
            Left            =   3480
            TabIndex        =   66
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Surface Water Flow Coeff."
            Height          =   375
            Left            =   3480
            TabIndex        =   64
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Groundwater Flow Expon."
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Aquifer Name"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   400
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Surface Elevation (ft)"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   1365
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Receiving Node"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   870
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Groundwater Flow Coeff."
            Height          =   375
            Left            =   240
            TabIndex        =   56
            Top             =   1780
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Initial Buildup Values"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   4215
         Begin MSComctlLib.ListView listInitialBuildUp 
            Height          =   2775
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Landuse Distribution"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   4215
         Begin MSComctlLib.ListView listLanduses 
            Height          =   2775
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   35
         Top             =   2520
         Width           =   6255
         Begin VB.ComboBox cmbSnowPacks 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkGrndWater 
            Height          =   255
            Left            =   1680
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtCurbLength 
            Height          =   315
            Left            =   1680
            TabIndex        =   39
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Curb Length (ft)"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Snow Packs"
            Height          =   255
            Left            =   3360
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label24 
            Caption         =   "Ground Water"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   6255
         Begin VB.TextBox txtPercentRouted 
            Height          =   315
            Left            =   4920
            TabIndex        =   25
            Text            =   "100"
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbSubareaRouting 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "% Routed"
            Height          =   255
            Left            =   3600
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Subarea Routing"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1680
            TabIndex        =   49
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtArea 
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            Top             =   1320
            Width           =   1200
         End
         Begin VB.TextBox txtPercentZeroImpervious 
            Height          =   315
            Left            =   4680
            TabIndex        =   20
            Text            =   "25"
            Top             =   2370
            Width           =   1200
         End
         Begin VB.TextBox txtDPervious 
            Height          =   315
            Left            =   4680
            TabIndex        =   18
            Text            =   "0.05"
            Top             =   1890
            Width           =   1200
         End
         Begin VB.TextBox txtDImpervious 
            Height          =   315
            Left            =   1680
            TabIndex        =   16
            Text            =   "0.05"
            Top             =   2835
            Width           =   1200
         End
         Begin VB.TextBox txtNPervious 
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            Text            =   "0.001"
            Top             =   1380
            Width           =   1200
         End
         Begin VB.TextBox txtNImpervious 
            Height          =   315
            Left            =   1680
            TabIndex        =   12
            Text            =   "0.10"
            Top             =   2325
            Width           =   1200
         End
         Begin VB.TextBox txtPercentImpervious 
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Top             =   870
            Width           =   1200
         End
         Begin VB.TextBox txtPercentSlope 
            Height          =   315
            Left            =   1680
            TabIndex        =   8
            Text            =   "0.005"
            Top             =   1830
            Width           =   1200
         End
         Begin VB.TextBox txtWidth 
            Height          =   315
            Left            =   4680
            TabIndex        =   6
            Top             =   375
            Width           =   1200
         End
         Begin VB.ComboBox cmbRainGauge 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   825
            Width           =   1200
         End
         Begin VB.Label Label6 
            Caption         =   "Rain Gauge"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "% Zero Impervious"
            Height          =   255
            Left            =   3240
            TabIndex        =   19
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Dstore-Pervious"
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Dstore-Impervious"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2865
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "N-Pervious"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   1410
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "N-Impervious"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   2355
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "% Impervious"
            Height          =   255
            Left            =   3240
            TabIndex        =   9
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "% Slope"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Width (ft)"
            Height          =   255
            Left            =   3240
            TabIndex        =   5
            Top             =   405
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Area (acre)"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1350
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Infiltration"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   6255
         Begin VB.TextBox txtInitialDeficit 
            Height          =   315
            Left            =   4920
            TabIndex        =   34
            Text            =   "4"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtConductivity 
            Height          =   315
            Left            =   1680
            TabIndex        =   32
            Text            =   "0.5"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtSuctionHead 
            Height          =   315
            Left            =   4920
            TabIndex        =   30
            Text            =   "3.0"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Initial Deficit"
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Conductivity (in/hr)"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "GREEN_AMPT"
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Suction Head (in)"
            Height          =   255
            Left            =   3600
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Infiltration Method"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmSWMMSubCatchmentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Outletnode As Long

Private Sub chkGrndWater_Click()
    If chkGrndWater.value = vbChecked Then
        CatchmentTab.TabEnabled(4) = True
    Else
        CatchmentTab.TabEnabled(4) = False
    End If
End Sub

Private Sub cmdCancel_Click()
    '** close the dialog box
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = New Scripting.Dictionary
    
    '** All input validation - check if the values are entered correctly
    CatchmentTab.Tab = 0
    If Not IsNumeric(txtArea.Text) Then
        MsgBox "Area must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "Width must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtPercentSlope.Text) Then
        MsgBox "Slope must be a valid number."
        Exit Sub
    End If
    
    If (CDbl(txtPercentSlope.Text) < 0) Or (CDbl(txtPercentSlope.Text) > 100) Then
        MsgBox "Slope must be a valid percentage (0-100)%."
        Exit Sub
    End If
    
    If Not IsNumeric(txtPercentImpervious.Text) Then
        MsgBox "Percent impervious must be a valid number."
        Exit Sub
    End If
    
    If (CDbl(txtPercentImpervious.Text) < 0) Or (CDbl(txtPercentImpervious.Text) > 100) Then
        MsgBox "Percent impervious must be a valid percentage (0-100)%."
        Exit Sub
    End If
    
    If Not IsNumeric(txtNImpervious.Text) Then
        MsgBox "NImpervious must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtNPervious.Text) Then
        MsgBox "NPervious must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtDImpervious.Text) Then
        MsgBox "DImpervious must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtDPervious.Text) Then
        MsgBox "DPervious must be a valid number."
        Exit Sub
    End If
        
    If Not IsNumeric(txtPercentZeroImpervious.Text) Then
        MsgBox "Percent zero impervious must be a valid number."
        Exit Sub
    End If
    
    If (CDbl(txtPercentZeroImpervious.Text) < 0) Or (CDbl(txtPercentZeroImpervious.Text) > 100) Then
        MsgBox "Percent zero impervious must be a valid percentage (0-100)%."
        Exit Sub
    End If
       
    CatchmentTab.Tab = 1
    If Not IsNumeric(txtPercentRouted.Text) Then
        MsgBox "Percent routed must be a valid number."
        Exit Sub
    End If
    
    If (CDbl(txtPercentRouted.Text) < 0) Or (CDbl(txtPercentRouted.Text) > 100) Then
        MsgBox "Percent routed must be a valid percentage (0-100)%."
        Exit Sub
    End If
    
    If Not IsNumeric(txtSuctionHead.Text) Then
        MsgBox "Suction head must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtConductivity.Text) Then
        MsgBox "Conductivity must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtInitialDeficit.Text) Then
        MsgBox "Initial deficit must be a valid number."
        Exit Sub
    End If
    
    If Not IsNumeric(txtCurbLength.Text) Then
        MsgBox "Curb length must be a valid number."
        Exit Sub
    End If

    '** general information
    pOptionProperty.add "Name", txtName.Text
    pOptionProperty.add "Outlet", m_Outletnode
    pOptionProperty.add "Rain Gauge", cmbRainGauge.Text
    pOptionProperty.add "Area", txtArea.Text
    pOptionProperty.add "Width", txtWidth.Text
    pOptionProperty.add "%Slope", txtPercentSlope.Text
    pOptionProperty.add "%Impervious", txtPercentImpervious.Text
    pOptionProperty.add "NImpervious", txtNImpervious.Text
    pOptionProperty.add "NPervious", txtNPervious.Text
    pOptionProperty.add "DImpervious", txtDImpervious.Text
    pOptionProperty.add "DPervious", txtDPervious.Text
    pOptionProperty.add "%ZeroImpervious", txtPercentZeroImpervious.Text
    
    '** additional information
    pOptionProperty.add "SubareaRouting", cmbSubareaRouting.Text
    pOptionProperty.add "%Routing", txtPercentRouted.Text
    pOptionProperty.add "InfiltrationMethod", "GREEN_AMPT"
    pOptionProperty.add "Suction Head", txtSuctionHead.Text
    pOptionProperty.add "Conductivity", txtConductivity.Text
    pOptionProperty.add "Initial Deficit", txtInitialDeficit.Text
    pOptionProperty.add "Ground Water", chkGrndWater.value
    pOptionProperty.add "Snow Packs", cmbSnowPacks.Text
    pOptionProperty.add "Curb Length", txtCurbLength.Text
        
    ' Store the Ground Water Table Props.....
    pOptionProperty.add "GWAquiferName", cmbAquifer.Text
    pOptionProperty.add "GWReceivingNode", txtReceiveNode.Text
    pOptionProperty.add "GWSurfaceElevation", txtSurElev.Text
    pOptionProperty.add "GWCoeff", txtGWCoeff.Text
    pOptionProperty.add "GWExp", txtGWExp.Text
    pOptionProperty.add "GWSurfaceCoeff", txtSurCoeff.Text
    pOptionProperty.add "GWSurfaceExp", txtSurExp.Text
    pOptionProperty.add "GWSurfaceGWCoeff", txtSurGW.Text
    pOptionProperty.add "GWSurfaceDepth", txtSurDep.Text
    pOptionProperty.add "GWThresholdElev", txtThershold.Text

    '** landuse information
    Dim iCount As Integer
    Dim pItem As ListItem
    Dim pTotalPercent As Double
    pTotalPercent = 0
    For iCount = 1 To listLanduses.ListItems.Count
        Set pItem = listLanduses.ListItems.Item(iCount)
        If Not (IsNumeric(pItem.SubItems(1))) Then
            CatchmentTab.Tab = 2
            MsgBox "Landuse percentage should be a valid percentage."
            Exit Sub
        End If
        pOptionProperty.add "Landuse: " & pItem.Text, pItem.SubItems(1)
        pTotalPercent = pTotalPercent + CDbl(pItem.SubItems(1))
    Next
    
    '** Validate total percentage
    If (pTotalPercent <> 100#) And (100# - pTotalPercent > 0.1) Then
        CatchmentTab.Tab = 2
        MsgBox "Total percentage for all landuses should total to 100%."
        Exit Sub
    End If
    
    '** initial buildup information
    Set pItem = Nothing
    For iCount = 1 To listInitialBuildUp.ListItems.Count
        Set pItem = listInitialBuildUp.ListItems.Item(iCount)
        If Not (IsNumeric(pItem.SubItems(1))) Then
            CatchmentTab.Tab = 3
            MsgBox "Pollutant - Initial buildup should be a valid number."
            Exit Sub
        End If
        pOptionProperty.add "Pollutant: " & pItem.Text, pItem.SubItems(1)
    Next
    
    '** get the subcatchment ID
    Dim pSubCatchmentID As String
    pSubCatchmentID = txtSubCatchment.Text
    
    '** call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDSubCatchments", pSubCatchmentID, pOptionProperty
    
    
    'Unload the form
    Unload Me
    
End Sub

Private Sub Form_Load()
        
    On Error GoTo ErrHandler
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '** Define variables
    Dim pCount As Integer
    Dim iCount As Integer
    Dim itmX As ListItem
    Dim pValue As String
    CatchmentTab.Tab = 0
    CatchmentTab.TabEnabled(4) = False
    
    '** LOAD all values for rain gages options
    Dim pRainGaugeCollection As Collection
    Set pRainGaugeCollection = ModuleSWMMFunctions.LoadRainGaugeNames
    pCount = pRainGaugeCollection.Count
    For iCount = 1 To pCount
        cmbRainGauge.AddItem pRainGaugeCollection.Item(iCount)
    Next
    cmbRainGauge.ListIndex = 0
    
    '** LOAD all values for snow Pack options
    Dim pSnowPackCollection As Collection
    Set pSnowPackCollection = ModuleSWMMFunctions.LoadSnowPackNames
    
    cmbSnowPacks.AddItem ""
    If Not pSnowPackCollection Is Nothing Then
        pCount = pSnowPackCollection.Count
        For iCount = 1 To pCount
            cmbSnowPacks.AddItem pSnowPackCollection.Item(iCount)
        Next
    End If
    'cmbSnowPacks.ListIndex = 0
    
    '** LOAD all values for snow Pack options
    Dim pAquiferCollection As Collection
    Set pAquiferCollection = ModuleSWMMFunctions.LoadAquiferNames
    cmbAquifer.AddItem ""
    If Not pAquiferCollection Is Nothing Then
        pCount = pAquiferCollection.Count
        For iCount = 1 To pCount
            cmbAquifer.AddItem pAquiferCollection.Item(iCount)
        Next
    End If
    'cmbAquifer.ListIndex = 0

    '** LOAD all values for subarea routing options
    cmbSubareaRouting.AddItem "OUTLET"
    cmbSubareaRouting.AddItem "IMPERVIOUS"
    cmbSubareaRouting.AddItem "PERVIOUS"
    cmbSubareaRouting.ListIndex = 0
    
    '** get values from dictionary, if value is present
    Dim pPropertyDictionary As Scripting.Dictionary
    Set pPropertyDictionary = LoadSubCatchmentProperties(gSubCatchmentID)
   
    '** Refresh landuse names
    Dim pLanduseCollection As Collection
    Set pLanduseCollection = ModuleSWMMFunctions.LoadLanduseNames
    
    '** define landuse list header
    listLanduses.ColumnHeaders.add , , "Landuse Name", listLanduses.Width * 0.5
    listLanduses.ColumnHeaders.add , , "% of Area", listLanduses.Width * 0.48
    
    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadPollutantNames
     
    '** add pollutant name to the column count
    listInitialBuildUp.ColumnHeaders.add , , "Pollutant Name", listInitialBuildUp.Width * 0.5
    listInitialBuildUp.ColumnHeaders.add , , "Initial Buildup (lbs/ac)", listInitialBuildUp.Width * 0.48
           
    '** load from dictionary, if dictionary is not null
    If Not (pPropertyDictionary Is Nothing) Then
         '** general information
        txtName.Text = pPropertyDictionary.Item("Name")
        cmbRainGauge.Text = pPropertyDictionary.Item("Rain Gauge")
        txtArea.Text = pPropertyDictionary.Item("Area")
        txtWidth.Text = pPropertyDictionary.Item("Width")
        txtPercentSlope.Text = pPropertyDictionary.Item("%Slope")
        txtPercentImpervious.Text = pPropertyDictionary.Item("%Impervious")
        txtNImpervious.Text = pPropertyDictionary.Item("NImpervious")
        txtNPervious.Text = pPropertyDictionary.Item("NPervious")
        txtDImpervious.Text = pPropertyDictionary.Item("DImpervious")
        txtDPervious.Text = pPropertyDictionary.Item("DPervious")
        txtPercentZeroImpervious.Text = pPropertyDictionary.Item("%ZeroImpervious")
           
        '** additional information
        cmbSubareaRouting.Text = pPropertyDictionary.Item("SubareaRouting")
        txtPercentRouted.Text = pPropertyDictionary.Item("%Routing")
        txtSuctionHead.Text = pPropertyDictionary.Item("Suction Head")
        txtConductivity.Text = pPropertyDictionary.Item("Conductivity")
        txtInitialDeficit.Text = pPropertyDictionary.Item("Initial Deficit")
        txtCurbLength.Text = pPropertyDictionary.Item("Curb Length")
        If pPropertyDictionary.Item("Ground Water") = "1" Then
            chkGrndWater.value = vbChecked
        Else
            chkGrndWater.value = vbUnchecked
        End If
                
        '** landuses
        pCount = pLanduseCollection.Count
        For iCount = 1 To pCount
            '** populate the list
            Set itmX = listLanduses.ListItems.add(, , pLanduseCollection.Item(iCount))
            pValue = "0.0"
            If ((Not pPropertyDictionary Is Nothing)) Then
                pValue = pPropertyDictionary.Item("Landuse: " & pLanduseCollection.Item(iCount))
            End If
            itmX.SubItems(1) = pValue
        Next
                
        '** pollutant - initial buildup
        pCount = pPollutantCollection.Count
        For iCount = 1 To pCount
            '** populate the list
            Set itmX = listInitialBuildUp.ListItems.add(, , pPollutantCollection.Item(iCount))
            pValue = "0.0"
            If ((Not pPropertyDictionary Is Nothing)) Then
                pValue = pPropertyDictionary.Item("Pollutant: " & pPollutantCollection.Item(iCount))
            End If
            itmX.SubItems(1) = pValue
        Next
            
    Else    '** If the properties for this subwatershed are not available
            '** load the subwatershed/landuse distribution
        If (gSubWaterLandUseDict Is Nothing) Then
            Call FindAndConvertWatershedFeatureLayerToRaster
            Call ComputeLanduseAreaForEachSubBasin
        End If
        Dim pLanduseDictionary As Scripting.Dictionary
        Set pLanduseDictionary = gSubWaterLandUseDict.Item(gSubCatchmentID)
        
        If Not (pLanduseDictionary Is Nothing) Then
                        
            Dim pLanduseAreaDictionary As Scripting.Dictionary
            Set pLanduseAreaDictionary = CreateObject("Scripting.Dictionary")
            Dim pLanduseImPervDictionary As Scripting.Dictionary
            Set pLanduseImPervDictionary = CreateObject("Scripting.Dictionary")
        
            Dim pSWMMLuReclass As iTable
            Set pSWMMLuReclass = GetInputDataTable("LUReclass")
            Dim pQueryFilter As IQueryFilter
            Set pQueryFilter = New QueryFilter
            Dim pCursor As ICursor
            Dim pRow As iRow
            pCount = pLanduseCollection.Count
            Dim pLUCode As Integer
            Dim pImperviousValue As Double
            Dim pLUGroupArea As Double
            Dim pAreaLanduseKey As String
''            For iCount = 1 To pCount
''                pQueryFilter.WhereClause = "LUGroup = '" & pLanduseCollection.Item(iCount) & "'"
''                Set pCursor = pSWMMLuReclass.Search(pQueryFilter, True)
''                Set pRow = pCursor.NextRow
''                pLUGroupArea = 0
''                Do While Not (pRow Is Nothing)
''                    pLUCode = pRow.value(pCursor.FindField("LUCode"))
''                    pImperviousValue = pRow.value(pCursor.FindField("Percentage"))
''
''                    pAreaLanduseKey = "Landuse: " & pLanduseCollection.Item(iCount)
''                    If (pLanduseAreaDictionary.Exists(pAreaLanduseKey)) Then
''                        pLUGroupArea = pLanduseAreaDictionary.Item(pAreaLanduseKey)
''                    End If
''                    pLanduseAreaDictionary.Item(pAreaLanduseKey) = pLUGroupArea + pLanduseDictionary.Item(pLUCode)
''                    pLanduseImPervDictionary.Item(pAreaLanduseKey) = pImperviousValue
''                    Set pRow = pCursor.NextRow
''                Loop
''            Next
            
            Dim iGroupName As Long
            iGroupName = pSWMMLuReclass.FindField("LUGroup")
            Dim iLuCode As Long
            iLuCode = pSWMMLuReclass.FindField("LUCode")
            Dim iPercentage As Long
            iPercentage = pSWMMLuReclass.FindField("Percentage")
            Dim iImpervious As Long
            iImpervious = pSWMMLuReclass.FindField("Impervious")
            Dim groupName As String
            Dim fraction As Double
            
            Set pCursor = pSWMMLuReclass.Search(Nothing, False)
            Set pRow = pCursor.NextRow
            pLUGroupArea = 0
            Do While Not (pRow Is Nothing)
                pLUCode = pRow.value(iLuCode)
                fraction = pRow.value(iPercentage)
                If CInt(pRow.value(iImpervious)) = 1 Then
                    groupName = pRow.value(iGroupName) & "_imp"
                    pImperviousValue = 1#   'pRow.value(iGroupCode)
                Else
                    groupName = pRow.value(iGroupName) & "_perv"
                    pImperviousValue = 0#   ' 1 - pRow.value(iGroupCode)
                End If
                pAreaLanduseKey = "Landuse: " & groupName 'pLanduseCollection.Item(iCount)
                If (pLanduseAreaDictionary.Exists(pAreaLanduseKey)) Then
                    pLUGroupArea = pLanduseAreaDictionary.Item(pAreaLanduseKey)
                Else
                    pLUGroupArea = 0
                End If
                pLanduseAreaDictionary.Item(pAreaLanduseKey) = pLUGroupArea + pLanduseDictionary.Item(pLUCode) * fraction
                pLanduseImPervDictionary.Item(pAreaLanduseKey) = pImperviousValue
                Set pRow = pCursor.NextRow
            Loop
            
        End If
        
        If gMetersPerUnit = 0# Then GetMetersPerLinearUnit
        Dim pSQMeterFactor As Double
        pSQMeterFactor = gMetersPerUnit * gMetersPerUnit    'Per sq. unit area converted to sq. meter
        Dim pSQAcreFactor As Double
        pSQAcreFactor = pSQMeterFactor * 0.0002471044       'sq meter to acre conversion
    
        '** TOTAL AREA COMPUTATION
        Dim pTotalArea As Double
        Dim pWeightedImpervious As Double
        Dim pTotalImpervious As Double
        Dim pKeys
        
        For Each pkey In pLanduseAreaDictionary.keys
            pTotalArea = pTotalArea + pLanduseAreaDictionary.Item(pkey)
            pWeightedImpervious = pWeightedImpervious + (pLanduseImPervDictionary.Item(pkey) * pLanduseAreaDictionary.Item(pkey))
        Next
        
        pTotalImpervious = pWeightedImpervious / pTotalArea
        '** Load the total area in the area box
        'txtArea.Text = Format(pTotalArea * 0.00002295675, "0.00")
        txtArea.Text = Format(pTotalArea * pSQAcreFactor, "0.00")
        txtWidth.Text = Format(Sqr(pTotalArea), "0.00")
        txtPercentImpervious.Text = Format(pTotalImpervious * 100, "0.00")
                
        '** add landuse description to the column count
        pCount = pLanduseCollection.Count
        Dim pLUGroupPctArea As Double
        Dim pFormattedValue As Double
        For iCount = 1 To pCount
            '** populate the list
            Set itmX = listLanduses.ListItems.add(, , pLanduseCollection.Item(iCount))
            pValue = "0.0"
            If ((Not pPropertyDictionary Is Nothing)) Then
                pValue = pPropertyDictionary.Item("Landuse: " & pLanduseCollection.Item(iCount))
            Else
                pValue = pLanduseAreaDictionary.Item("Landuse: " & pLanduseCollection.Item(iCount))
            End If
            pLUGroupPctArea = Format(pValue / pTotalArea * 100, "0.00")
            itmX.SubItems(1) = pLUGroupPctArea
        Next
                
        '** add pollutant name to the column count
        pCount = pPollutantCollection.Count
        Set itmX = Nothing
        For iCount = 1 To pCount
            '** populate the list
            Set itmX = listInitialBuildUp.ListItems.add(, , pPollutantCollection.Item(iCount))
            pValue = "0.0"
            If ((Not pPropertyDictionary Is Nothing)) Then
                pValue = pPropertyDictionary.Item("Pollutant: " & pPollutantCollection.Item(iCount))
            End If
            itmX.SubItems(1) = pValue
        Next
    End If
    
    ' Update the Receiving Node Attr.....
    Dim pFeatLayer As IFeatureLayer
    Set pFeatLayer = GetInputFeatureLayer("Watershed")
    If pFeatLayer Is Nothing Then Exit Sub
    Dim pFCursor As IFeatureCursor
    Dim pFRow As IFeature
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & gSubCatchmentID
    Set pFCursor = pFeatLayer.Search(pQueryFilter, True)
    Set pFRow = pFCursor.NextFeature
    txtReceiveNode.Text = pFRow.value(pFCursor.FindField("BMPID"))
    m_Outletnode = pFRow.value(pFCursor.FindField("BMPID"))

    Exit Sub
    
ErrHandler:
    MsgBox "Form Load : " & Err.description
    
End Sub

Private Sub listLanduses_DblClick()
    
    Dim pItem As ListItem
    Set pItem = listLanduses.SelectedItem

    'Get existing value
    Dim pDefault
    pDefault = pItem.SubItems(1)
    'Get new input value
    Dim bValue
    bValue = InputBox("Enter value for % of Area", "% of Area ", pDefault)

    If (Trim(bValue) = "") Then
        bValue = pDefault
    End If
    pItem.SubItems(1) = bValue
End Sub


Private Sub listInitialBuildUp_DblClick()
    
    Dim pItem As ListItem
    Set pItem = listInitialBuildUp.SelectedItem

    'Get existing value
    Dim pDefault
    pDefault = pItem.SubItems(1)
    'Get new input value
    Dim bValue
    bValue = InputBox("Enter value for Initial Buildup in lbs/ac", "Initial Buildup", pDefault)

    If (Trim(bValue) = "") Then
        bValue = pDefault
    End If
    pItem.SubItems(1) = bValue
End Sub


Private Function LoadSubCatchmentProperties(pID As Integer) As Scripting.Dictionary

    '* get the table from map
    Dim pSWMMSubCatchmentTable As iTable
    Set pSWMMSubCatchmentTable = GetInputDataTable("LANDSubCatchments")
    If (pSWMMSubCatchmentTable Is Nothing) Then
        Exit Function
    End If
    
    '* define indexes for field names
    Dim iPropName As Long
    iPropName = pSWMMSubCatchmentTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMSubCatchmentTable.FindField("PropValue")
    
    '* define query filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pID
    
    '* get the cursor to iterate over the table
    Dim pCursor As ICursor
    Set pCursor = pSWMMSubCatchmentTable.Search(pQueryFilter, True)
    
    '* define a row variable to loop over the table
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Dim pPropertyDict As Scripting.Dictionary
    If Not pRow Is Nothing Then
        Set pPropertyDict = CreateObject("Scripting.Dictionary")
    End If
    Do While Not pRow Is Nothing
        pPropertyDict.add Trim(pRow.value(iPropName)), Trim(pRow.value(iPropValue))
        Set pRow = pCursor.NextRow
    Loop
        
    '** cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMSubCatchmentTable = Nothing
    Set pQueryFilter = Nothing
  
    '** return the property dictionary
    Set LoadSubCatchmentProperties = pPropertyDict
    
End Function
