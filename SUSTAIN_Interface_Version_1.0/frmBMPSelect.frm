VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBMPSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select BMP Category"
   ClientHeight    =   9690
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13455
   Icon            =   "frmBMPSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   484.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   672.75
   Begin VB.CommandButton cmdAggregate 
      Caption         =   "Aggregate"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   9135
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      Begin VB.Image ImgScale 
         Height          =   8775
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3585
      Left            =   650
      Max             =   3
      Min             =   1
      TabIndex        =   1
      Top             =   360
      Value           =   1
      Width           =   300
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   9240
      Width           =   1215
   End
   Begin MSComctlLib.ImageList Imglst 
      Left            =   480
      Top             =   6960
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
            Picture         =   "frmBMPSelect.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPSelect.frx":3F385
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBMPSelect.frx":690B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   0
      Picture         =   "frmBMPSelect.frx":B0C69
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmBMPSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_iPic As Integer
Private m_BmpName As String

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdAggregate_Click()
    'commented out on Oct 31, 2008 - Sabu Paul -
    'Same code is used elsewhere also
    'So this is made into another subroutine.
'''    ' Load the BMPS with their Types.......
'''    Set gBMPTypeDict = New Scripting.Dictionary
'''    gBMPTypeDict.RemoveAll
'''    gBMPTypeDict.Add "Bioretention", "Aggregate"
'''    gBMPTypeDict.Add "Dry Pond", "Aggregate"
'''    gBMPTypeDict.Add "Wet Pond", "Aggregate"
'''    gBMPTypeDict.Add "Rain Barrel", "Aggregate"
'''    gBMPTypeDict.Add "Cistern", "Aggregate"
'''    gBMPTypeDict.Add "Porous Pavement", "Aggregate"
'''    gBMPTypeDict.Add "Green Roof", "Aggregate"
'''    gBMPTypeDict.Add "Infiltration Trench", "Aggregate"
'''    gBMPTypeDict.Add "Vegetative Swale", "Aggregate"
'''    gBMPTypeDict.Add "Conduit", "Aggregate"
'''    ' Load the BMPS with their Category.......
'''    Set gBMPCatDict = New Scripting.Dictionary
'''    gBMPCatDict.RemoveAll
'''    gBMPCatDict.Add "Bioretention", "On-Site Treatment"
'''    gBMPCatDict.Add "Dry Pond", "Regional Storage/Treatment"
'''    gBMPCatDict.Add "Wet Pond", "Regional Storage/Treatment"
'''    gBMPCatDict.Add "Rain Barrel", "On-Site Interception"
'''    gBMPCatDict.Add "Cistern", "On-Site Interception"
'''    gBMPCatDict.Add "Porous Pavement", "On-Site Treatment"
'''    gBMPCatDict.Add "Green Roof", "On-Site Interception"
'''    gBMPCatDict.Add "Infiltration Trench", "On-Site Treatment"
'''    gBMPCatDict.Add "Vegetative Swale", "Routing Attenuation"
'''    gBMPCatDict.Add "Conduit", "Routing Attenuation"
'''
'''    ' Initialize the Dict.....
'''    Set gBMPPlacedDict = CreateObject("Scripting.Dictionary")
    
    Call InitializeAggBMPTypes
    
    Call Set_Dimensions("Aggregate")
           
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    Call sldScale_Click
    Call VScroll_Change
    
    Call InitBmpTypeCatDict
    
    If gInternalSimulation Then
        cmdAggregate.Enabled = False
    Else
        cmdAggregate.Enabled = True
    End If
End Sub


Private Sub ImgScale_Click()
    
    If ImgScale.Tag <> "" Then

'        Dim frmBMPDef As frmBMPDef
'        Set frmBMPDef = New frmBMPDef
'        frmBMPDef.ImgBMP.Picture = imglst.ListImages(m_iPic).Picture
'        frmBMPDef.Show vbModal

        Dim pBMPType As String
         Select Case ImgScale.Tag
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
                pBMPType = "BufferStrip"
                MsgBox "Buffer strip is currently under testing and will be release in a future patch.", vbInformation  'will be enabled in version 1.0"
                pBMPType = ""
'                Unload Me
'                Call Define_Bufferstrip
'                Exit Sub
            Case ""
                MsgBox "This BMP Type will be implemented in future phase of SUSTAIN", vbInformation
        End Select
        
        m_BmpName = ImgScale.Tag
        If pBMPType = "" Then Exit Sub
        Call Set_Dimensions(pBMPType)

    End If
    
End Sub



Private Sub Set_Dimensions(pBMPType As String)
    
    gNewBMPType = pBMPType
        
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pBMPTypesTable As iTable
    If pBMPType = "BufferStrip" Then
        Set pBMPTypesTable = GetInputDataTable("VFSDefaults")
    Else
        Set pBMPTypesTable = GetInputDataTable("BMPTypes")
        If (pBMPTypesTable Is Nothing) Then
            Set pBMPTypesTable = CreateBMPTypesDBF("BMPTypes")
            AddTableToMap pBMPTypesTable
        End If
        pQueryFilter.WhereClause = "Type = '" & pBMPType & "'"
    End If
    
    If Not pBMPTypesTable Is Nothing Then
        Dim pSelRowCount As Long
        pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
    End If
    
    Dim iTab As Integer
    iTab = VScroll.value - 1
    
    Unload Me
            
    If pSelRowCount > 0 Then
        If pBMPType = "BufferStrip" Then
            FrmVFSTypes.Show vbModal
        ElseIf pBMPType = "Aggregate" Then
            frmAggTypes.Show vbModal
        Else
            Load FrmBMPTypes
            FrmBMPTypes.m_BmpName = m_BmpName
            FrmBMPTypes.m_iTab = iTab
            FrmBMPTypes.Form_Initialize
            FrmBMPTypes.Show vbModal
        End If
    Else
       'Set the name of the BMP to bmpType
       Dim pNameCount As Long
       pNameCount = pSelRowCount + 1
       
       If pBMPType = "Aggregate" Then
            Load frmAggBMPDef
            frmAggBMPDef.Tag = ""
            frmAggBMPDef.Form_Initialize
            frmAggBMPDef.TabBMPType.Tab = 3
            frmAggBMPDef.BMPType.Text = pBMPType & pNameCount
            frmAggBMPDef.Show vbModal
        Else
           Load frmBMPDef
           frmBMPDef.EditType.Text = ""
           frmBMPDef.Form_Initialize
           frmBMPDef.TabBMPType.Tab = Get_Tab_Index(gBMPTypeDict.Item(m_BmpName))
           frmBMPDef.cmbBMPCategory.ListIndex = iTab
           frmBMPDef.Update_BMP_Types
           frmBMPDef.cmbBmpType.Text = m_BmpName
           frmBMPDef.BMPNameA.Text = pBMPType & pNameCount
           frmBMPDef.Show vbModal
       End If
    End If

End Sub



Private Sub ImgScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim icoHand As IPictureDisp
    Set icoHand = LoadResPicture("Hand", vbResCursor)
    
    If m_iPic = 1 Then ' Watershed Scale....
        If (X > 800 And Y < 8000) And (X < 2400 And Y > 7600) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = ""
        ElseIf (X > 4600 And Y < 3400) And (X < 5500 And Y > 3000) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Buffer Strip"
        ElseIf (X > 6900 And Y < 7900) And (X < 8500 And Y > 7600) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = ""
        ElseIf (X > 6200 And Y < 400) And (X < 10200 And Y > 300) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = ""
        Else
            ImgScale.MouseIcon = Nothing: Screen.MousePointer = 0
            ImgScale.Tag = ""
        End If
    End If
    
    If m_iPic = 2 Then ' Community Scale.....
        If (X > 660 And Y < 7100) And (X < 2100 And Y > 6700) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Infiltration Trench"
        ElseIf (X > 4600 And Y < 6500) And (X < 5900 And Y > 6000) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Vegetative Swale"
        ElseIf (X > 9600 And Y < 6600) And (X < 10300 And Y > 6200) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Dry Pond"
        ElseIf (X > 4500 And Y < 1100) And (X < 5200 And Y > 700) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Wet Pond"
        Else
            ImgScale.MouseIcon = Nothing: Screen.MousePointer = 0
            ImgScale.Tag = ""
        End If
    End If
    
     If m_iPic = 3 Then  ' Lot Scale.....
        If (X > 2200 And Y < 4800) And (X < 3200 And Y > 4400) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Bioretention"
        ElseIf (X > 1450 And Y < 2900) And (X < 2400 And Y > 2600) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Cistern"
        ElseIf (X > 1450 And Y < 2300) And (X < 2250 And Y > 2000) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Rain Barrel"
        ElseIf (X > 9000 And Y < 3850) And (X < 10300 And Y > 3400) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Porous Pavement"
        ElseIf (X > 7500 And Y < 1600) And (X < 8400 And Y > 1100) Then
            ImgScale.MouseIcon = icoHand: ImgScale.MousePointer = 99
            ImgScale.Tag = "Green Roof"
        Else
            ImgScale.MouseIcon = Nothing: Screen.MousePointer = 0
            ImgScale.Tag = ""
        End If
    End If
    
End Sub

Private Sub sldScale_Click()
    
    ImgScale.Picture = Imglst.ListImages(VScroll.value).Picture
    'If VScroll.value = 3 Then ImgScale.Picture = Imglst.ListImages(4).Picture
End Sub

Private Sub VScroll_Change()
    
'    ImgScale.Picture = Imglst.ListImages(VScroll.value).Picture
'    m_iPic = VScroll.value
    m_iPic = VScroll.value
    'If m_iPic = 3 Then m_iPic = VScroll.value + 1
    ImgScale.Picture = Imglst.ListImages(m_iPic).Picture
    
    
End Sub
