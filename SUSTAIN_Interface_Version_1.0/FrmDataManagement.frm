VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmDataManagement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Data Path"
   ClientHeight    =   2670
   ClientLeft      =   6825
   ClientTop       =   6615
   ClientWidth     =   5970
   Icon            =   "FrmDataManagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCost 
      Height          =   350
      Left            =   5400
      Picture         =   "FrmDataManagement.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Browse to select a directory"
      Top             =   840
      Width           =   500
   End
   Begin VB.TextBox txtCostDBPath 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   873
      Width           =   3300
   End
   Begin VB.TextBox txtTempDirectory 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   1755
      Width           =   3300
   End
   Begin VB.CommandButton cmdTmpDir 
      Height          =   350
      Left            =   5400
      Picture         =   "FrmDataManagement.frx":592C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Browse to select a directory"
      Top             =   1725
      Width           =   500
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   5040
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.TextBox txtGDBPath 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   3300
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   350
      Left            =   5400
      Picture         =   "FrmDataManagement.frx":5A76
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Browse to select a directory"
      Top             =   1275
      Width           =   500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Cost Database"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   888
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   $"FrmDataManagement.frx":5BC0
      Height          =   600
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "&Temporary Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "&Geodatabase"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1312
      Width           =   2175
   End
   Begin VB.Label lblHelp 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmDataManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\FrmDataManage.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

Private Sub cmdBrowse_Click()

    On Error GoTo ErrorHandler
    
''    Dim pattern As String
''    pattern = "Personal GeoDatabase (*.mdb)|*.mdb"
''    'pattern = "File Based GeoDatabase (*.gdb)|*.gdb"
''    dlgSave.Filter = pattern
''    dlgSave.CancelError = True
''    dlgSave.ShowSave
''
''    If (Err <> cdlCancel) Then
''
''        Dim fso As Scripting.FileSystemObject
''        Set fso = CreateObject("Scripting.FileSystemObject")
''        Dim FilePath As String, FileName As String
''        FilePath = fso.GetParentFolderName(dlgSave.FileName)
''        FileName = fso.GetFileName(dlgSave.FileName)
''        txtGDBPath.Text = dlgSave.FileName
''
''    End If
''
''    Exit Sub

    Dim pGxDialog As IGxDialog
    Dim pFileGDBFilter As IGxObjectFilter
    Set pFileGDBFilter = New GxFilterFileGeodatabases
    Set pGxDialog = New GxDialog
    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog
    pFilterCol.AddFilter pFileGDBFilter, True
    Dim pEnumGx As IEnumGxObject
    pGxDialog.Title = "Browse for File GeoDatabase"
    If Not pGxDialog.DoModalOpen(0, pEnumGx) Then
        Exit Sub 'Exit if user press Cancel
    End If
    txtGDBPath.Text = pEnumGx.Next.FullName
    Exit Sub
ErrorHandler:
  If Err = cdlCancel Then Exit Sub
  HandleError True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

    Unload Me

  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
End Sub



Private Sub cmdCost_Click()
    On Error GoTo ErrorHandler
    
    Dim pattern As String
    pattern = "Access Database (*.mdb)|*.mdb"
    dlgSave.Filter = pattern
    dlgSave.CancelError = True
    dlgSave.ShowOpen
        
    If (Err <> cdlCancel) Then
        
        Dim fso As Scripting.FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim FilePath As String, FileName As String
        FilePath = fso.GetParentFolderName(dlgSave.FileName)
        FileName = fso.GetFileName(dlgSave.FileName)
        txtCostDBPath.Text = dlgSave.FileName
    End If
        
    Exit Sub
ErrorHandler:
  If Err = cdlCancel Then Exit Sub
  HandleError True, "cmdCost " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Private Sub cmdOk_Click()
  On Error GoTo ErrorHandler
   
   'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
   Call DefineApplicationPath
    Dim pAppPath As String
    pAppPath = gApplicationPath
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim dataSrcFN As String
    dataSrcFN = gApplication.Document.Title
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_data.src"
    dataSrcFN = pAppPath & dataSrcFN
        
    Dim pDataSrcFile
    Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForWriting, True, TristateUseDefault)
      
    ' Check if the Cost Database is Selected....
    If txtCostDBPath.Text = "" Then
        MsgBox "Select the Cost Database", vbCritical
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
    ElseIf Not (fso.FileExists(txtCostDBPath.Text)) Then
        MsgBox "Cost database not found. Please check.", vbCritical
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
    Else
        gCostDBpath = txtCostDBPath.Text
        pDataSrcFile.WriteLine "gCostDBpath" & vbTab & gCostDBpath
    End If
    
    ' Check if the Geodatabase is Selected....
    If txtGDBPath.Text = "" Then
        MsgBox "Select the Geodatabase to load the data.", vbCritical
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
''    ElseIf Not (fso.FileExists(txtGDBPath.Text)) Then
''        MsgBox "Geodatabase not found. Please check.", vbCritical
''        'Close the file....
''        pDataSrcFile.Close
''        fso.DeleteFile dataSrcFN
''        Exit Sub
    ElseIf Not (fso.FolderExists(txtGDBPath.Text)) Then
        MsgBox "Geodatabase not found. Please check.", vbCritical
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
    Else
        'Store the GeoDatabase Path... Load the data....
        If gGDBpath <> "" Then gDataLoad = True
        gGDBpath = txtGDBPath.Text
        gGDBFlag = True
        pDataSrcFile.WriteLine "gGDBpath" & vbTab & gGDBpath
    End If
    ' Check for the Temp Folder....
    If (fso.FolderExists(fso.GetParentFolderName(txtTempDirectory.Text))) Then
        If Not (fso.FolderExists(txtTempDirectory.Text)) Then
            fso.CreateFolder txtTempDirectory.Text
        End If
        gMapTempFolder = Trim(Me.txtTempDirectory)
        pDataSrcFile.WriteLine "gMapTempFolder" & vbTab & gMapTempFolder
    ElseIf gMapTempFolder = "" Then
        MsgBox "Set the temperory directory"
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
    Else
        MsgBox txtTempDirectory.Text & " folder does not exist"
        'Close the file....
        pDataSrcFile.Close
        fso.DeleteFile dataSrcFN
        Exit Sub
    End If
    

    'Close the file....
    pDataSrcFile.Close
    ' Close the form..........
    Me.Hide
    Unload Me
    
    '' ####################################
    If gDataLoad Then
        Set gFeatClassDictionary = CreateObject("Scripting.Dictionary") ' to store the FeatureClass Names.....
        gFeatClassDictionary.RemoveAll
'        Dim iCnt As Integer
'        Dim pLayer As ILayer
'        Dim pFeatLayer As IFeatureLayer
'        For iCnt = 0 To (gMap.LayerCount - 1)
'            Set pLayer = gMap.Layer(iCnt)
'            If pLayer.Valid Then
'                If TypeOf pLayer Is IFeatureLayer Then
'                    Set pFeatLayer = pLayer
'                    gFeatClassDictionary.Add pFeatLayer.name, "FeatureClass"
'                End If
'            End If
'        Next iCnt
'        Dim pTabCollection As IStandaloneTableCollection
'        Dim pStTable As IStandaloneTable
'        Set pTabCollection = gMap
'        For iCnt = 0 To (pTabCollection.StandaloneTableCount - 1)
'           Set pStTable = pTabCollection.StandaloneTable(iCnt)
'           gFeatClassDictionary.Add pStTable.name, "Table"
'        Next
    End If
    
CleanUp:

  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
End Sub


Private Sub cmdTmpDir_Click()
    On Error GoTo ShowError
        Dim strTmpDir As String
       'now fill the strPath with the choice by user
        'strTmpDir = BrowseForFolder(0, "Select the project temporary directory")
        strTmpDir = BrowseForSpecificFolder("Select the project temporary directory", gApplicationPath)
        If (Trim(strTmpDir) <> "") Then
            txtTempDirectory.Text = strTmpDir
        End If
        Exit Sub
ShowError:
        MsgBox "Find Temp Directory Click :" & Err.description
End Sub




Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    
    Dim dataSrcFN As String
    dataSrcFN = gApplication.Document
    ' ** Activate the Def Pollutants form......
    'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = gApplicationPath
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & ".src"
    dataSrcFN = pAppPath & dataSrcFN
    
    Dim fso As New FileSystemObject
    If fso.FileExists(dataSrcFN) Then
        'SUSTAIN.frmDataManage.Show
        Load SUSTAIN.frmDataManage
        If frmDataManage.Landuse.Text <> "" And frmDataManage.lulookup.Text <> "" And frmDataManage.DEM.Text <> "" And frmDataManage.STREAM.Text <> "" Then
            SUSTAIN.frmDataManage.cmdOk_Click
        Else
            Unload SUSTAIN.frmDataManage
        End If
    End If
    
    If gCostDBpath <> "" Then
        txtCostDBPath.Text = gCostDBpath
    Else
        Dim pModelFolder As String
        pModelFolder = ""
        pModelFolder = ModuleUtility.GetApplicationPath & "\etc\"
        If pModelFolder <> "" Then txtCostDBPath.Text = pModelFolder & "BMPCosts.mdb"
    End If
    
    If (gGDBpath <> "") Then txtGDBPath.Text = gGDBpath
    
    If (gMapTempFolder <> "") Then txtTempDirectory.Text = gMapTempFolder
    
'    Dim pTable As iTable
'    Set pTable = GetInputDataTable("SimulationOption")
'
'    If (pTable Is Nothing) Then Exit Sub
'
'    Dim pCursor As ICursor
'    Set pCursor = pTable.Search(Nothing, True)
'    Dim pRow As iRow
'    Set pRow = pCursor.NextRow
'
'    Do While Not pRow Is Nothing
'        If pRow.value(pTable.FindField("PropValue")) = "False" Then
'            gInternalSimulation = False
'        Else
'            gInternalSimulation = True
'        End If
'        Set pRow = pCursor.NextRow
'    Loop
'    gExternalSimulation = (Not gInternalSimulation)
'
'    '** cleanup
'    Set pRow = Nothing
'    Set pCursor = Nothing
'    Set pTable = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND


End Sub

