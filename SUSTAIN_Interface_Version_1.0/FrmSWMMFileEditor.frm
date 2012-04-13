VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSWMMFileEditor 
   Caption         =   "View/Edit Simulation File"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "FrmSWMMFileEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox InputFileStream 
      Height          =   6255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   9120
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label InputFileLabel 
      Caption         =   "Simulation Input File Name"
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
      TabIndex        =   4
      Top             =   120
      Width           =   7845
   End
End
Attribute VB_Name = "FrmSWMMFileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
On Error GoTo ErrorHandler:
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim pattern As String
    pattern = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.Filter = pattern
    CommonDialog.CancelError = True
    
    Dim pInputFileName As String
    CommonDialog.ShowOpen
    Dim pFile As TextStream
    
    If (Err <> cdlCancel) Then
        pInputFileName = CommonDialog.FileName
        InputFileLabel = "Current File: " & pInputFileName
        Set pFile = fso.OpenTextFile(pInputFileName, ForReading, True, TristateUseDefault)
        Dim fileContent As String
        InputFileStream.Text = pFile.ReadAll
        pFile.Close
    End If
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error while reading the input file: " & Err.description
CleanUp:
    Set fso = Nothing
    Set pFile = Nothing
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorHandler:
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim pattern As String
    pattern = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.Filter = pattern
    CommonDialog.CancelError = True
        
    Dim pInputFileName As String
    CommonDialog.FileName = Replace(InputFileLabel, "Current File: ", "")
    Dim pFile As TextStream
    
    CommonDialog.ShowSave
    
    If (Err <> cdlCancel) Then
        pInputFileName = CommonDialog.FileName
        Set pFile = fso.CreateTextFile(pInputFileName, True, False)
        pFile.Write (InputFileStream.Text)
        pFile.Close
        MsgBox "Saved the input file as " & pInputFileName
    End If
    GoTo CleanUp
ErrorHandler:
    MsgBox "Error while saving the input file: " & Err.description
CleanUp:
    Set fso = Nothing
    Set pFile = Nothing
End Sub

Private Sub Form_Resize()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If FrmSWMMFileEditor.Width > 3000 Then
        InputFileStream.Width = FrmSWMMFileEditor.Width - 555
    End If
    If FrmSWMMFileEditor.Height > 3000 Then
        InputFileStream.Height = FrmSWMMFileEditor.Height - 1600
    End If
    
    cmdCancel.Left = FrmSWMMFileEditor.Width - 1275
    cmdCancel.Top = FrmSWMMFileEditor.Height - 1000
    
    cmdSave.Left = FrmSWMMFileEditor.Width - 2595
    cmdSave.Top = FrmSWMMFileEditor.Height - 1000
    
    cmdLoad.Left = FrmSWMMFileEditor.Width - 3915
    cmdLoad.Top = FrmSWMMFileEditor.Height - 1000

End Sub

