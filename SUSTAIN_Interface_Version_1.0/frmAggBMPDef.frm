VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAggBMPDef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BMP Definition"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10665
   Icon            =   "frmAggBMPDef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10665
   Begin VB.Frame Frame1 
      Caption         =   "Surface Properties"
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10455
      Begin VB.TextBox BMPType 
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
         Left            =   7440
         TabIndex        =   18
         Top             =   615
         Width           =   2745
      End
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
         TabIndex        =   14
         Top             =   615
         Width           =   2745
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Edit Cost Function"
         Height          =   375
         Index           =   3
         Left            =   6480
         TabIndex        =   13
         Top             =   4440
         Width           =   1850
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Water Quality Parameters"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   12
         Top             =   4440
         Width           =   2000
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "Subsurface Properties"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   11
         Top             =   3960
         Width           =   1850
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Close"
         Height          =   855
         Left            =   8400
         TabIndex        =   8
         Top             =   3960
         Width           =   1800
      End
      Begin VB.ComboBox cmbBMPCategory 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cmbBMPType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1350
         Width           =   3975
      End
      Begin VB.CommandButton cmdDimensions 
         Caption         =   "BMP Dimensions"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   5
         Top             =   3960
         Width           =   2000
      End
      Begin VB.Frame Frame2 
         Caption         =   "Infiltration Method"
         Height          =   735
         Left            =   4440
         TabIndex        =   4
         Top             =   1080
         Width           =   5775
         Begin VB.OptionButton optGreen 
            Caption         =   "Green Amprt"
            Height          =   255
            Left            =   3120
            TabIndex        =   21
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optHal 
            Caption         =   "Holtan"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pollutant Removal Method"
         Height          =   735
         Left            =   4440
         TabIndex        =   3
         Top             =   1920
         Width           =   5775
         Begin VB.OptionButton optDecay 
            Caption         =   "1st Order Decay"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optKadlac 
            Caption         =   "K-C* method                           (Kadlec and Knight Method)"
            Height          =   375
            Left            =   3120
            TabIndex        =   22
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pollutant Routing Method"
         Height          =   1095
         Left            =   4440
         TabIndex        =   2
         Top             =   2760
         Width           =   5775
         Begin VB.OptionButton optSeries 
            Caption         =   "CSTRs in series"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton optPlug 
            Caption         =   "Plug Flow"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   320
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton optMixed 
            Caption         =   "Completely Mixed"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   320
            Value           =   -1  'True
            Width           =   1935
         End
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
            TabIndex        =   17
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label4 
            Caption         =   "No. of CSTRs"
            Height          =   360
            Left            =   3400
            TabIndex        =   16
            Top             =   705
            Width           =   1320
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Type"
         Height          =   255
         Left            =   7440
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Select BMP Category"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Select BMP Type"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Image ImgBMP 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   240
         Picture         =   "frmAggBMPDef.frx":08CA
         Stretch         =   -1  'True
         Top             =   1800
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
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Point"
      TabPicture(0)   =   "frmAggBMPDef.frx":F8AD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Line"
      TabPicture(1)   =   "frmAggBMPDef.frx":F8C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Area"
      TabPicture(2)   =   "frmAggBMPDef.frx":F8E5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Aggregate"
      TabPicture(3)   =   "frmAggBMPDef.frx":F901
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
   End
End
Attribute VB_Name = "frmAggBMPDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_Close As Boolean
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\projects\sustain\SUSTAIN_9_3\frmAggBMPDef.frm"


Private Sub BMPType_Change()
  On Error GoTo ErrorHandler

    If Not gBMPPlacedDict Is Nothing Then gBMPPlacedDict.RemoveAll

  Exit Sub
ErrorHandler:
  HandleError True, "BMPType_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

'Public Sub Initialize_Form(strTag As String)
'    Me.Tag = strTag
'End Sub

Private Sub CancelButton_Click()
  On Error GoTo ErrorHandler

    Set gBMPDetailDict = Nothing
    
'    ' Validate the BMP Placed & Warn the user if missing.......
'    If Not gBMPEditMode Then
'        If gBMPPlacedDict.Count < 4 Then
'            If MsgBox("You should place one BMP from each category." & vbNewLine & "Closing the form now will delete the BMPs for this Aggregate BMP. Do you want to proceed?", vbCritical + vbYesNo, "SUSTAIN") = vbYes Then
'
'                ' Delete the BMP.....
'                Dim pBMPTypesTable As iTable
'                Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
'
'                Dim pIDindex As Long
'                pIDindex = pBMPTypesTable.FindField("ID")
'                Dim pNameIndex As Long
'                pNameIndex = pBMPTypesTable.FindField("PropValue")
'
'                'Check the existence of BioRetBasin in the table
'                Dim pQueryFilter As IQueryFilter
'                Set pQueryFilter = New QueryFilter
'                Dim pCursor As ICursor
'                Dim pRow As iRow
'
'                Dim pkey, pKeys, ikey As Integer, strID As String
'                Dim pBMPtoDel As Scripting.Dictionary
'                Set pBMPtoDel = CreateObject("Scripting.Dictionary")
'                pKeys = gBMPPlacedDict.keys
'                For ikey = 0 To gBMPPlacedDict.Count - 1
'                    pkey = pKeys(ikey)
'                    pQueryFilter.WhereClause = "PropName='Category' And PropValue='" & pkey & "'"
'                    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
'                    Set pRow = pCursor.NextRow
'                    If Not pRow Is Nothing Then pBMPtoDel.Add pRow.value(pIDindex), pRow.value(pIDindex)
'                    Set pCursor = Nothing
'                Next
'
'                pKeys = pBMPtoDel.keys
'                For ikey = 0 To pBMPtoDel.Count - 1
'                    pkey = pKeys(ikey)
'                    pQueryFilter.WhereClause = "ID=" & pkey & " And PropName='Type' And PropValue='" & BMPType.Text & "'"
'                    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
'                    Set pRow = pCursor.NextRow
'                    If pRow Is Nothing Then pBMPtoDel.Remove pkey
'                    Set pCursor = Nothing
'                Next
'                pKeys = pBMPtoDel.keys
'                For ikey = 0 To pBMPtoDel.Count - 1
'                    pkey = pKeys(ikey)
'                    strID = strID & "," & pkey
'                Next
'                strID = Mid(strID, 2)
'
'                ' Clean...
'                Set pRow = Nothing
'                Set pCursor = Nothing
'
'                pQueryFilter.WhereClause = "ID In (" & strID & ")"
'                '** delete records
'                pBMPTypesTable.DeleteSearchedRows pQueryFilter
'                Set pBMPTypesTable = GetInputDataTable("BMPTypes")
'                '** delete records
'                pBMPTypesTable.DeleteSearchedRows pQueryFilter
'
'            Else
'                Exit Sub
'            End If
'        End If
'    End If
    
    m_Close = True
    Unload Me
    gBMPEditMode = False
    

  Exit Sub
ErrorHandler:
  HandleError True, "CancelButton_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Private Sub cmbBMPCategory_Click()
  On Error GoTo ErrorHandler

    Call Update_BMP_Types

  Exit Sub
ErrorHandler:
  HandleError True, "cmbBMPCategory_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub


Private Sub cmbBMPType_Click()
  On Error GoTo ErrorHandler

    
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
        Case "Conduit"
            pBMPType = "Conduit"
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
    pQueryFilter.WhereClause = "Name LIKE '" & pBMPType & "%'"
 
    Dim pSelRowCount As Long
    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    
    BMPNameA.Text = pBMPType & pSelRowCount + 1
        

  Exit Sub
ErrorHandler:
  HandleError True, "cmbBMPType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Private Sub cmdDimensions_Click(Index As Integer)
On Error GoTo ShowError
    'MsgBox "gBMPTypeTag = " & gBMPTypeTag & " Edit mode" & CStr(gBMPEditMode)
    
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
        Case "Conduit"
            pBMPType = "Conduit"
    End Select
    
    '  set the Tabindex ........
    If pBMPType = "" Then Exit Sub
    If Index = 0 Then   'added this assginment on Jan 07, 2009
        gBMPDefTab = 1
    ElseIf Index = 1 Then
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
    gBMPOptionsDict.add "Category", cmbBMPCategory.Text
    gBMPOptionsDict.add "Type", BMPType.Text
    gBMPOptionsDict.add "BMPType", pBMPType 'cmbBmpType.Text
    
    gNewBMPType = pBMPType
    ' ********************************************
    ' Check for duplicate Record.....
    ' ********************************************
    
    If gBMPEditMode Then GoTo Proceed
    'If Me.Tag <> "BMPOnMap" Then
    If gBMPTypeTag <> "BMPOnMap" Then
        Dim pBMPTypesTable As iTable
        Set pBMPTypesTable = GetInputDataTable("BMPTypes")
        
        'Check the existence of a BMP in the table
        'Validate.....
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "Name = '" & gNewBMPName & "'"
        Dim pCount As Long
        pCount = pBMPTypesTable.RowCount(pQueryFilter)
        If pCount > 0 Then
            MsgBox "A BMP already exists with this name. Please rename and proceed.", vbInformation, "SUSTAIN"
            Exit Sub
        End If
        
        'Validate.....
        If gBMPPlacedDict.Count = 0 Then Call Check_PlacedBMPs
    End If
'    If gBMPPlacedDict.Exists(cmbBMPCategory.Text) Then
'        MsgBox "This category BMP exists. Please change the category.", vbInformation, "SUSTAIN"
'        Exit Sub
'    End If
    'If gBMPPlacedDict.Exists(cmbBMPType.Text) Then
    If gBMPPlacedDict.Exists(pBMPType) Then
        MsgBox "This BMP Type exists. Please change the type.", vbInformation, "SUSTAIN"
        Exit Sub
    End If
Proceed:

    ' Load the Details form...
    If gBMPEditMode Then Me.Hide 'Unload Me
    Call Edit_BMPParameters(pBMPType)
    
    Exit Sub
ShowError:
    MsgBox "Error in defining dimensions :" & Err.description
End Sub

Private Sub Check_PlacedBMPs()
  On Error GoTo ErrorHandler

    
    Dim pBMPTypesTable As iTable
    'If Me.Tag = "BMPOnMap" Then Exit Sub
    If gBMPTypeTag = "BMPOnMap" Then Exit Sub
    
    Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    If pBMPTypesTable Is Nothing Then Exit Sub
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("PropValue")
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    pQueryFilter.WhereClause = "PropName='Type' And PropValue='" & BMPType.Text & "'"
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow

    Dim pQueryFilter2 As IQueryFilter
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
    Do While Not pRow Is Nothing
        Set pQueryFilter2 = New QueryFilter
        'pQueryFilter2.WhereClause = "PropName='Category' And ID = " & pRow.value(pIDindex)
        pQueryFilter2.WhereClause = "PropName='BMPType' And ID = " & pRow.value(pIDindex)
        Set pCursor2 = pBMPTypesTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then gBMPPlacedDict.add pRow2.value(pNameIndex), pRow2.value(pNameIndex)
        Set pRow = pCursor.NextRow
    Loop



  Exit Sub
ErrorHandler:
  HandleError False, "Check_PlacedBMPs " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Public Sub Form_Initialize()
  On Error GoTo ErrorHandler


    cmbBMPCategory.Clear
    cmbBMPCategory.AddItem "On-Site Interception"
    cmbBMPCategory.AddItem "On-Site Treatment"
    cmbBMPCategory.AddItem "Routing Attenuation"
    cmbBMPCategory.AddItem "Regional Storage/Treatment"
      
    frmAggBMPDef.TabBMPType.TabEnabled(0) = False
    frmAggBMPDef.TabBMPType.TabEnabled(1) = False
    frmAggBMPDef.TabBMPType.TabEnabled(2) = False
    BMPType.Text = ""
    
' ***************************************
    'If redefining then load the data from the Table....
    ' ***************************************
    If gBMPEditMode Then
        
        Dim pBmpDetailDict As Scripting.Dictionary
        'Set pBmpDetailDict = GetBMPPropDict(gNewBMPId)
               
        'If Me.Tag <> "BMPOnMap" Then
        If gBMPTypeTag <> "BMPOnMap" Then
            Set pBmpDetailDict = GetBMPPropDict(gNewBMPId)
        Else
            Set pBmpDetailDict = GetBMPDetailDict(gNewBMPId, "AgBMPDetail")
        End If
        
        If (Not pBmpDetailDict Is Nothing) Then
            If BMPNameA.Text = "" Then
                If pBmpDetailDict.Exists("BMPName") Then BMPNameA.Text = pBmpDetailDict.Item("BMPName")
            End If
            
            If BMPType.Text = "" Then
                If pBmpDetailDict.Exists("Type") Then BMPType.Text = pBmpDetailDict.Item("Type")
            End If
            
            optHal.value = (Not CBool(CInt(pBmpDetailDict.Item("Infiltration Method"))))
            optGreen.value = (Not optHal.value)
            optDecay.value = (Not CBool(CInt(pBmpDetailDict.Item("Pollutant Removal Method"))))
            optKadlac.value = (Not optDecay.value)
            If CInt(pBmpDetailDict.Item("Pollutant Routing Method")) = 0 Then
                optPlug.value = True
            ElseIf CInt(pBmpDetailDict.Item("Pollutant Routing Method")) = 1 Then
                optMixed.value = True
            ElseIf CInt(pBmpDetailDict.Item("Pollutant Routing Method")) > 1 Then
                optSeries.value = True
                txtCSTR.Text = CInt(pBmpDetailDict.Item("Pollutant Routing Method"))
            End If
        End If
    Else
        optHal.value = True
        optDecay.value = True
        optMixed.value = True
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Initialize " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
   

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler

    If Not m_Close Then Cancel = 1

  Exit Sub
ErrorHandler:
  HandleError True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub


Private Sub TabBMPType_Click(PreviousTab As Integer)
  On Error GoTo ErrorHandler

    Call Update_BMP_Types

  Exit Sub
ErrorHandler:
  HandleError True, "TabBMPType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub


Public Sub Update_BMP_Types()
  On Error GoTo ErrorHandler

    
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
    

  Exit Sub
ErrorHandler:
  HandleError True, "Update_BMP_Types " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

Private Sub Edit_BMPParameters(curBMPType As String)
  On Error GoTo ErrorHandler

    
    If gBMPEditMode Then
        ' Modify....
        
        'If Me.Tag = "BMPOnMap" Then
        If gBMPTypeTag = "BMPOnMap" Then
            Dim pBmpDetailDict As Scripting.Dictionary
            Set pBmpDetailDict = GetBMPDetailDict(gNewBMPId, "AgBMPDetail")
            ModuleBMPData.LoadPollutantData pBmpDetailDict
            'Call ModifyBmpDetails(gNewBMPId, gNewBMPName)
            Set gBMPDetailDict = New Scripting.Dictionary
            ModuleBMPData.CallInitRoutines curBMPType, pBmpDetailDict
        Else
            Call ModifyBmpTypeDetails(gNewBMPId, gNewBMPName)
        End If
    Else 'If a new BMP is added to the agg BMP
''        If Me.Tag = "BMPOnMap" Then
''            ModuleBMPData.CallInitRoutines curBMPType, Nothing
''        Else
''            'Create a new BMP of the current type
''            Call CreateNewBmpType("Aggregate", gNewBMPName, Me.Tag)
''        End If
        Call CreateNewBmpType("Aggregate", gNewBMPName)  ',Me.Tag)
    End If
'    'If Me.Tag = "BMPOnMap" Then
'    If gBMPTypeTag = "BMPOnMap" Then
    'MsgBox "Category (cmbBMPCategory.Text) = " & cmbBMPCategory.Text & "Type(BMPType.Text)" & BMPType.Text
        'Me.Hide
        If Not gBMPDetailDict Is Nothing Then
            gBMPDetailDict.Item("Category") = cmbBMPCategory.Text
            gBMPDetailDict.Item("Type") = BMPType.Text
        End If
        
        If gBMPTypeTag = "BMPOnMap" Then Me.Hide
'    End If
        

  Exit Sub
ErrorHandler:
  HandleError False, "Edit_BMPParameters " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub
        

''Private Sub Define_Bufferstrip()
''
''    FrmVFSParams.txtName.Text = gNewBMPName
''    FrmVFSParams.txtName.Enabled = True
''    FrmVFSParams.BufferLength.Text = ""
''    FrmVFSParams.txtName.Enabled = True
''    FrmVFSParams.BufferWidth.Text = ""
''    FrmVFSParams.BufferWidth.Enabled = True
''
''    '** Open the VFS Defaults table to get default name
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("VFSDefaults")
''    If (pTable Is Nothing) Then
''        '** open the form that defines the buffer strip params
'''        FrmVFSData.txtVFSID.Text = 1
'''        FrmVFSData.txtName.Text = "VFS1"
'''        FrmVFSData.Show vbModal
''
''        Dim pVFSDictionary As Scripting.Dictionary
''        Set pVFSDictionary = GetDefaultsForVFS(1, "VFS1")
''
''        InitializeVFSPropertyForm pVFSDictionary
''        FrmVFSParams.Show vbModal
''    Else
''        FrmVFSTypes.Show vbModal
''    End If
''    Set pTable = Nothing
''
''
''    If (FrmVFSParams.bContinue = True) Then
''        Dim pIDValue As Integer
''        pIDValue = FrmVFSParams.txtVFSID.Text
''
'''        '** create the dictionary
'''        Set gBufferStripDetailDict = CreateObject("Scripting.Dictionary")
'''        gBufferStripDetailDict.Add "Name", FrmVFSData.txtName.Text
'''        gBufferStripDetailDict.Add "BufferLength", FrmVFSData.txtBufferLength.Text
'''        gBufferStripDetailDict.Add "BufferWidth", FrmVFSData.txtBufferWidth.Text
''
''        '** call the generic function to create and add rows for values
''        ModuleVFSFunctions.SaveVFSPropertiesTable "VFSDefaults", CStr(pIDValue), gBufferStripDetailDict
''
''        '** set it to nothing
''        Set gBufferStripDetailDict = Nothing
''        Unload FrmVFSParams
''    End If
''
''End Sub
