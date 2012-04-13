VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBMPDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP Definition"
   ClientHeight    =   5715
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10665
   Icon            =   "frmBMPDef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10665
   Begin VB.TextBox EditType 
      Height          =   285
      Left            =   840
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Surface Properties"
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10455
      Begin VB.TextBox BMPNameA 
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
         Left            =   4440
         TabIndex        =   21
         Top             =   615
         Width           =   5745
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Edit Cost Function"
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   20
         Top             =   4440
         Width           =   1850
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Water Quality Parameters"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   19
         Top             =   4440
         Width           =   2000
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Subsurface Properties"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   18
         Top             =   3960
         Width           =   1850
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Close"
         Height          =   855
         Left            =   8400
         TabIndex        =   15
         Top             =   3960
         Width           =   1800
      End
      Begin VB.ComboBox cmbBMPCategory 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cmbBMPType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1350
         Width           =   3975
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "BMP Dimensions"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   12
         Top             =   3960
         Width           =   2000
      End
      Begin VB.Frame Frame2 
         Caption         =   "Infiltration Method"
         Height          =   735
         Left            =   4440
         TabIndex        =   9
         Top             =   1080
         Width           =   5775
         Begin VB.OptionButton optHal 
            Caption         =   "Holtan"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optGreen 
            Caption         =   "Green Ampt"
            Height          =   255
            Left            =   3120
            TabIndex        =   10
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pollutant Removal Method"
         Height          =   735
         Left            =   4440
         TabIndex        =   6
         Top             =   1920
         Width           =   5775
         Begin VB.OptionButton optKadlac 
            Caption         =   "K-C* method                           (Kadlec and Knight Method)"
            Height          =   375
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton optDecay 
            Caption         =   "1st Order Decay"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pollutant Routing Method"
         Height          =   1095
         Left            =   4440
         TabIndex        =   2
         Top             =   2760
         Width           =   5775
         Begin VB.TextBox txtCSTR 
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
            Left            =   4680
            TabIndex        =   24
            Top             =   600
            Width           =   945
         End
         Begin VB.OptionButton optMixed 
            Caption         =   "Completely Mixed"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optPlug 
            Caption         =   "Plug Flow"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2640
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optSeries 
            Caption         =   "CSTRs in series"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label4 
            Caption         =   "No. of CSTRs"
            Height          =   360
            Left            =   3400
            TabIndex        =   23
            Top             =   705
            Width           =   1320
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Select BMP Category"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Select BMP Type"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Image ImgBMP 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   3975
      End
   End
   Begin TabDlg.SSTab TabBMPType 
      Height          =   340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   609
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Point"
      TabPicture(0)   =   "frmBMPDef.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Line"
      TabPicture(1)   =   "frmBMPDef.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Area"
      TabPicture(2)   =   "frmBMPDef.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Aggregate"
      TabPicture(3)   =   "frmBMPDef.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
   Begin MSComctlLib.ImageList Imglst 
      Left            =   120
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   763
      ImageHeight     =   581
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPDef.frx":093A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPDef.frx":3F3F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPDef.frx":9A762
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBMPDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    
    If gBMPEditMode = True Then
        'still need to populate optionDict coz setting was removed from detailDict
        Set gBMPOptionsDict = CreateObject("Scripting.Dictionary")
        gBMPOptionsDict.add "Infiltration Method", Abs(CInt(Not optHal.value))
        gBMPOptionsDict.add "Pollutant Removal Method", Abs(CInt(Not optDecay.value))
        If optPlug.value = True Then
            gBMPOptionsDict.add "Pollutant Routing Method", "0"
        ElseIf optMixed.value = True Then
            gBMPOptionsDict.add "Pollutant Routing Method", "1"
        Else
            gBMPOptionsDict.add "Pollutant Routing Method", txtCSTR.Text
        End If
        
    End If
    
    Unload Me
'    gBMPEditMode = False
End Sub

Private Sub cmbBMPCategory_Click()
    Call Update_BMP_Types
    ImgBMP.Picture = Imglst.ListImages(cmbBMPCategory.ListIndex + 1).Picture
End Sub


Private Sub cmbBMPType_Click()
    
    Frame2.Visible = True
    Frame3.Visible = True
    Frame4.Visible = True
    
    Dim pBMPType As String
     Select Case cmbBmpType.Text
        Case "Infiltration Trench"
            pBMPType = "InfiltrationTrench"
        Case "Vegetative Swale"
            pBMPType = "VegetativeSwale"
        Case "Wet Pond"
            pBMPType = "WetPond"
        Case "Dry Pond"
            pBMPType = "DryPond"
        Case "Bioretention"
            pBMPType = "BioRetentionBasin"
        Case "Rain Barrel"
            pBMPType = "RainBarrel"
        Case "Cistern"
            pBMPType = "Cistern"
        Case "Porous Pavement"
            pBMPType = "PorousPavement"
        Case "Green Roof"
            pBMPType = "GreenRoof"
        Case "Buffer Strip"
            pBMPType = "Buffer Strip"
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
    End Select
    gNewBMPType = pBMPType
    
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
    
    If (pBMPTypesTable Is Nothing) Then
        Set pBMPTypesTable = CreateBMPTypesDBF("BMPTypes")
        AddTableToMap pBMPTypesTable
    End If
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "Type = '" & pBMPType & "'"
 
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    
    BMPNameA.Text = gNewBMPType & pSelRowCount + 1
        
End Sub

Private Sub cmdDimensions_Click(Index As Integer)
       
    Dim pBMPType As String
     Select Case cmbBmpType.Text
        Case "Infiltration Trench"
            pBMPType = "InfiltrationTrench"
        Case "Vegetative Swale"
            pBMPType = "VegetativeSwale"
        Case "Wet Pond"
            pBMPType = "WetPond"
        Case "Dry Pond"
            pBMPType = "DryPond"
        Case "Bioretention"
            pBMPType = "BioRetentionBasin"
        Case "Rain Barrel"
            pBMPType = "RainBarrel"
        Case "Cistern"
            pBMPType = "Cistern"
        Case "Porous Pavement"
            pBMPType = "PorousPavement"
        Case "Green Roof"
            pBMPType = "GreenRoof"
        Case "Buffer Strip"
            pBMPType = "Buffer Strip"
        Case "Buffer Strip"
            pBMPType = "Conduit"
    End Select
    
    Dim cstrError As Boolean: cstrError = False
    If optSeries.value Then
        If Not IsNumeric(txtCSTR.Text) Then
            cstrError = True
        ElseIf CInt(txtCSTR.Text) <= 1 Then
            cstrError = True
        End If
    End If
    If cstrError Then
        MsgBox "Please enter a valid number (>1) of CSTRs in series", vbExclamation
        Exit Sub
    End If
    '  set the Tabindex ........
    If pBMPType = "" Then Exit Sub
    If Index = 1 Then
        gBMPDefTab = 2
    ElseIf Index = 2 Then
        gBMPDefTab = 5
    ElseIf Index = 3 Then
        gBMPDefTab = 6
    Else
        gBMPDefTab = 0
        If pBMPType = "VegetativeSwale" Then gBMPDefTab = 1
    End If
    gNewBMPName = BMPNameA.Text
    Set gBMPOptionsDict = CreateObject("Scripting.Dictionary")
    gBMPOptionsDict.add "Infiltration Method", Abs(CInt(Not optHal.value))
    gBMPOptionsDict.add "Pollutant Removal Method", Abs(CInt(Not optDecay.value))
    If optPlug.value = True Then
        gBMPOptionsDict.add "Pollutant Routing Method", "0"
    ElseIf optMixed.value = True Then
        gBMPOptionsDict.add "Pollutant Routing Method", "1"
    Else
        gBMPOptionsDict.add "Pollutant Routing Method", txtCSTR.Text
    End If
    
    ' ********************************************
    ' Check for duplicate Record.....
    ' ********************************************

    If gBMPEditMode Then GoTo Proceed
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "Name = '" & gNewBMPName & "'"
    Dim pCount As Long
    pCount = pBMPTypesTable.RowCount(pQueryFilter)
    If pCount > 0 Then
        MsgBox "A BMP already exists with this name. Please rename and proceed.", vbInformation, "SUSTAIN"
        Exit Sub
    End If
      
Proceed:
    ' Load the Details form...
'    Dim tagType As String
'    tagType = EditType.Text
'    EditType.Text = ""
    Unload Me
    
    Call Edit_BMPParameters   '(tagType)
    
End Sub

Public Sub Form_Initialize()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    cmbBMPCategory.Clear
    cmbBMPCategory.AddItem "Generalized Practices"
    cmbBMPCategory.AddItem "Conventional Practices"
    cmbBMPCategory.AddItem "Low-Impact Development Practices"
                
    ' ***************************************
    'If redefining then load the data from the Table....
    ' ***************************************
    If gBMPEditMode Then
        Dim pBmpDetailDict As Scripting.Dictionary
        'If EditType.Text <> "BMPOnMap" Then
        If gBMPTypeTag <> "BMPOnMap" Then
            Set pBmpDetailDict = GetBMPPropDict(gNewBMPId)
        Else
            Set pBmpDetailDict = GetBMPDetailDict(gNewBMPId)
        End If
        If (Not pBmpDetailDict Is Nothing) Then
            If BMPNameA.Text = "" Then
                If pBmpDetailDict.Exists("BMPName") Then BMPNameA.Text = pBmpDetailDict.Item("BMPName")
            End If
'            optHal.value = (Not CBool(CInt(pBmpDetailDict.Item("Infiltration Method"))))
'            optDecay.value = (Not CBool(CInt(pBmpDetailDict.Item("Pollutant Removal Method"))))
            
            If CInt(pBmpDetailDict.Item("Infiltration Method")) = 0 Then
                optHal.value = True
            Else
                optGreen.value = True
            End If
            If CInt(pBmpDetailDict.Item("Pollutant Removal Method")) = 0 Then
                optDecay.value = True
            Else
                optKadlac.value = True
            End If
            
            If CInt(pBmpDetailDict.Item("Pollutant Routing Method")) = 0 Then
                optPlug.value = True
            ElseIf CInt(pBmpDetailDict.Item("Pollutant Routing Method")) = 1 Then
                optMixed.value = True
            ElseIf CInt(pBmpDetailDict.Item("Pollutant Routing Method")) > 1 Then
                optSeries.value = True
                txtCSTR.Text = CInt(pBmpDetailDict.Item("Pollutant Routing Method"))
            End If
        End If
    End If
    
    TabBMPType.TabEnabled(3) = False
End Sub


Private Sub TabBMPType_Click(PreviousTab As Integer)
    Call Update_BMP_Types
End Sub


Public Sub Update_BMP_Types()
    BMPNameA.Text = ""
    cmbBmpType.Clear
    Dim pKeys
    pKeys = gBMPTypeDict.keys
    Dim pkey As String
    Dim ikey As Integer
    For ikey = 0 To gBMPTypeDict.Count - 1
        pkey = pKeys(ikey)
        If gBMPTypeDict.Item(pkey) = TabBMPType.TabCaption(TabBMPType.Tab) And _
            gBMPCatDict.Item(pkey) = cmbBMPCategory.Text Then cmbBmpType.AddItem pkey
    Next
    
End Sub

Private Sub Edit_BMPParameters()
    If gNewBMPType = "Buffer Strip" Then
        If gBMPDefTab = 2 Then gBMPDefTab = 1
        If gBMPDefTab = 5 Then gBMPDefTab = 4
        If gBMPDefTab = 6 Then gBMPDefTab = 0
        Call Define_Bufferstrip
    ElseIf gBMPEditMode Then
        gBMPEditMode = False
        ' Modify....
        'If strTag = "BMPOnMap" Then
        If gBMPTypeTag = "BMPOnMap" Then
            Call ModifyBmpDetails(gNewBMPId, gNewBMPName)
        Else
            Call ModifyBmpTypeDetails(gNewBMPId, gNewBMPName)
        End If
    Else
        'Create a new BMP of the current type
        Call CreateNewBmpType(gNewBMPType, gNewBMPName)
    End If
    
End Sub


