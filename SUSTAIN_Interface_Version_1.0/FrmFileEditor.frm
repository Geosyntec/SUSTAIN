VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFileEditor 
   Caption         =   "Edit Input File"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   Icon            =   "FrmFileEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox InputFileStream 
      Height          =   6375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   480
      Width           =   9120
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label InputFileLabel 
      Caption         =   "Input File Name"
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
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "FrmFileEditor"
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
    pattern = "Input File (*.inp)|*.inp"
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
    pattern = "Input File (*.inp)|*.inp"
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

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub

Private Sub Form_Resize()

    If FrmFileEditor.Width > 3000 Then
        InputFileStream.Width = FrmFileEditor.Width - 555
    End If
    If FrmFileEditor.Height > 3000 Then
        InputFileStream.Height = FrmFileEditor.Height - 1600
    End If
    
    cmdCancel.Left = FrmFileEditor.Width - 1275
    cmdCancel.Top = FrmFileEditor.Height - 1000
    
    cmdSave.Left = FrmFileEditor.Width - 2595
    cmdSave.Top = FrmFileEditor.Height - 1000
    
    cmdLoad.Left = FrmFileEditor.Width - 3915
    cmdLoad.Top = FrmFileEditor.Height - 1000

End Sub

