VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmWeatherData 
   Caption         =   "Rainfall Data File"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFileBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtWeatherData 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Select Rainfall Data File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmWeatherData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFileBrowse_Click()
    CommonDialog.Filter = "Rain Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
    CommonDialog.ShowOpen
    txtWeatherData.Text = CommonDialog.FileName
End Sub

Private Sub cmdOk_Click()
    If Trim(txtWeatherData.Text) = "" Then
        MsgBox "Enter a valid path for the rainfall data file", vbExclamation
        Exit Sub
    End If
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    If Not fso.FileExists(txtWeatherData.Text) Then
        MsgBox "File " & txtWeatherData.Text & " cannot be found. Enter a valid path for the rainfall data file.", vbExclamation
        Exit Sub
    End If
    
    gWeatherInputFile = txtWeatherData.Text
    Set fso = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
