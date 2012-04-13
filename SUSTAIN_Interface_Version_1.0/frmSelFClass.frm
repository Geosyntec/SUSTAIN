VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelFClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Delete data from Datamodel"
   ClientHeight    =   1785
   ClientLeft      =   4890
   ClientTop       =   5145
   ClientWidth     =   4095
   Icon            =   "frmSelFClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4095
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Exit"
      Height          =   555
      Left            =   2400
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   3480
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelFClass.frx":57E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   550
      Left            =   720
      Picture         =   "frmSelFClass.frx":58F4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton optDel 
         Caption         =   "Delete Data from DataModel"
         DisabledPicture =   "frmSelFClass.frx":59F6
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add Data to DataModel"
         DisabledPicture =   "frmSelFClass.frx":5E38
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   150
         Picture         =   "frmSelFClass.frx":65FA
         Top             =   360
         Width           =   600
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   150
         Picture         =   "frmSelFClass.frx":6DBC
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame frmFclist 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3855
      Begin MSComctlLib.ListView lstFCLass 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmSelFClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmSelFClass.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms
Private m_Flag As Boolean


Private Sub cmdDelete_Click()

    On Error GoTo ErrorHandler
    ' First Check if any item is Checked.....
    Dim lstItem As ListItem
    Dim pChecked As Boolean
    Dim iCnt As Integer
    Dim edataType As esriDatasetType
    Dim pTable As iTable, pFC As IFeatureClass
    
    
    pChecked = False
    For iCnt = 1 To lstFCLass.ListItems.Count
        If lstFCLass.ListItems(iCnt).Checked = True Then
           pChecked = True
           Exit For
        End If
    Next iCnt
    
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New FileGDBWorkspaceFactory
    
        
    If pChecked Then
        If (MsgBox("The selected data will be deleted from the database. Are you sure you want to delete the selected data?", vbQuestion + vbYesNo) = vbYes) Then
            Dim pLayer As ILayer
            Dim pWorkspace As IWorkspace
            'Set pWorkspace = ModuleUtility.OpenAccessWorkspace(gGDBpath)
            Set pWorkspace = pWorkspaceFactory.OpenFromFile(gGDBpath, 0)
            With lstFCLass
                For iCnt = 1 To .ListItems.Count
                    Set lstItem = .ListItems(iCnt): Set pTable = Nothing: Set pFC = Nothing
                    If lstItem.Tag = "Table" Then
                        edataType = esriDTTable
                        Set pTable = GetTable(gGDBpath, .ListItems(iCnt).Text)
                    Else
                        edataType = esriDTFeatureClass
                        Set pFC = GetFeatureClass(gGDBpath, .ListItems(iCnt).Text)
                    End If
                    If .ListItems(iCnt).Checked = True And (Not pTable Is Nothing Or Not pFC Is Nothing) Then
                       If DeleteGDBData(pWorkspace, .ListItems(iCnt).Text, edataType) Then
                            gFeatClassDictionary.Remove .ListItems(iCnt).Text ' Remove from the Dictionary....
                            Set pLayer = GetLayerFromMap(.ListItems(iCnt).Text)    ' remove from the Map.........
                            If Not pLayer Is Nothing Then gMap.DeleteLayer pLayer
                            .ListItems.Remove (iCnt) ' Remove from the Listview...
                            iCnt = iCnt - 1
                       End If
                    End If
                    If .ListItems.Count = iCnt Then Exit For
                Next iCnt
            End With
            MsgBox "Successfully deleted selected data from the datamodel.", vbInformation, "SUSTAIN"
        End If
        
    End If
    
    
    Exit Sub
ErrorHandler:
  HandleError True, "cmdDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub cmdExit_Click()
    
    On Error GoTo ErrorHandler
    Unload Me
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdExit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub cmdQuit_Click()
    
    On Error GoTo ErrorHandler
    Unload Me
    
Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If gFeatClassDictionary Is Nothing Then
        If Not CreateList_FromGDB(gGDBpath) Then Exit Sub
    End If
    Image2.Width = 250
    optAdd.value = False
    optDel.value = False
    lstFCLass.GridLines = True
    lstFCLass.View = lvwReport
    lstFCLass.ColumnHeaders.add , , "Feature Class", lstFCLass.Width / 2
    lstFCLass.ColumnHeaders.add , , "Type", lstFCLass.Width / 2
    
    Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub



Private Sub Form_Unload(Cancel As Integer)
    m_Flag = False
End Sub


Private Sub lstFCLass_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo ErrorHandler
    With Item
        If .Checked = True Then
            Item.SmallIcon = 1
            .Selected = True
            lstFCLass.Refresh
        Else
            Item.SmallIcon = 0
            .Selected = False
            lstFCLass.Refresh
        End If
    End With

Exit Sub
ErrorHandler:
  HandleError True, "lstFCLass_ItemCheck " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub optAdd_Click()

    On Error GoTo ErrorHandler
    'If m_Flag Then
        If optAdd.value = True Then
            cmdExit.Visible = True
            frmFclist.Visible = False
            Me.Height = 2265
            Me.Hide
            ModuleUtility.Load_ShapeFile_to_PGDB (gGDBpath)
            optAdd.value = False
            Me.Show vbModal
        End If
    'End If
    m_Flag = True
    
Exit Sub
ErrorHandler:
  HandleError True, "optAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
   

End Sub

Private Sub optDel_Click()
    
    On Error GoTo ErrorHandler
    If optDel.value = True Then
        cmdExit.Visible = False
        frmFclist.Visible = True
        lstFCLass.ListItems.Clear
        Dim lstItem As ListItem
        Dim pKeys
        
        If gFeatClassDictionary Is Nothing Then Exit Sub
        
        pKeys = gFeatClassDictionary.keys
        Dim pkey As String
        Dim ikey As Integer
        Dim pFeatureclass As IFeatureClass
        lstFCLass.SmallIcons = imglst
        For ikey = 0 To gFeatClassDictionary.Count - 1
            pkey = pKeys(ikey)
            'Set pFeatureclass = GetFeatureClass(gGDBpath, gFeatClassDictionary.Item(pkey))
            'If Not pFeatureclass Is Nothing Then
                Set lstItem = lstFCLass.ListItems.add(ikey + 1, , pkey)
                lstItem.ListSubItems.add , , gFeatClassDictionary.Item(pkey)
                lstItem.Tag = gFeatClassDictionary.Item(pkey)
            'End If
        Next
    
        Me.Height = 6030
    End If
    
Exit Sub
ErrorHandler:
  HandleError True, "optDel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
    
End Sub
