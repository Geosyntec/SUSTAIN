VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About SUSTAIN"
   ClientHeight    =   4725
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   9720
   ClipControls    =   0   'False
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAbout.frx":08CA
   ScaleHeight     =   3261.278
   ScaleMode       =   0  'User
   ScaleWidth      =   9127.581
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "Help"
      Default         =   -1  'True
      Height          =   345
      Left            =   8640
      TabIndex        =   7
      Top             =   720
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   8640
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   6720
      Picture         =   "FrmAbout.frx":1D2A2
      Top             =   480
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "FrmAbout.frx":1E496
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAbout.frx":1ED9D
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dennis F. Lai, Ph.D., P.E."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Designed and developed by:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Built for ArcGIS 9.3 (Build 1850)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   135
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date: October 16, 2009"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   135
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHelp_Click()
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim helpFileName As String
    helpFileName = "SUSTAIN_V1_Step_by_Step_Guide.pdf"
    'Check if the file is saved properly
    Dim pDocFolder As String
    pDocFolder = ""
  
    pDocFolder = ModuleUtility.GetApplicationPath & "\Documents\"
    
    If Not fso.FileExists(pDocFolder & helpFileName) Then
        Dim pattern As String
        pattern = "PDF File (*.pdf)|*.pdf"
        With FrmFileEditor.CommonDialog
            .DialogTitle = "Select SUSTAIN Step by Step Guide (PDF file)"
            .Filter = pattern
            .CancelError = False
            .ShowOpen
            If (Err <> cdlCancel) Then
                sFile = .FileName
                pDocFolder = fso.GetParentFolderName(sFile)
                helpFileName = fso.GetFileName(sFile)
                pDocFolder = pDocFolder & "\"
            End If
        End With
    End If
    
        
    OpenDoc pDocFolder, helpFileName
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub


Private Sub Image1_Click()
    Dim fso As New FileSystemObject
    If fso.FileExists("C:\Program Files\Internet Explorer\IEXPLORE.EXE") Then
        Call Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.tetratech.com", vbNormalFocus)
    ElseIf fso.FileExists("C:\Program Files\Mozilla Firefox\Firefox.exe") Then
        Call Shell("C:\Program Files\Mozilla Firefox\Firefox.exe http://www.tetratech.com", vbNormalFocus)
    End If
    Set fso = Nothing
End Sub

Private Sub Image2_Click()
    Dim fso As New FileSystemObject
    If fso.FileExists("C:\Program Files\Internet Explorer\IEXPLORE.EXE") Then
        Call Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE http://www.epa.gov", vbNormalFocus)
    ElseIf fso.FileExists("C:\Program Files\Mozilla Firefox\Firefox.exe") Then
        Call Shell("C:\Program Files\Mozilla Firefox\Firefox.exe http://www.epa.gov", vbNormalFocus)
    End If
End Sub
