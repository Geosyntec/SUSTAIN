VERSION 5.00
Begin VB.Form frmBufferStripFlowProps 
   Caption         =   "Define Overland Flow Properties"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmBufferStripFlowProps.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kinematic Wave Numerical Solution Parameters"
      Height          =   855
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buffer Dimensions"
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox txtBufferRoughness 
         Height          =   330
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtBufferSlope 
         Height          =   330
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtBufferWidth 
         Height          =   330
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBufferLength 
         Height          =   330
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Roughness"
         Height          =   330
         Left            =   3720
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Slope"
         Height          =   330
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Width of the Strip (m) [FWIDTH]"
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Buffer Length (m) [VL]"
         Height          =   330
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox txtSimulationTitle 
      Height          =   330
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtFlowPropertiesFile 
      Height          =   330
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Simulation Title"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Overland Flow Properties File (*.ikw)"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmBufferStripFlowProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '** input validation
    Dim pFlowPropsFile As String
    If (txtFlowPropertiesFile.Text = "") Then
        MsgBox "Please specify Overland Flow Input File."
        Exit Sub
    End If
    pFlowPropsFile = txtFlowPropertiesFile.Text
    
    If (txtSimulationTitle.Text = "") Then
        MsgBox "Please specify Simulation Title."
        Exit Sub
    End If
    If (txtBufferLength.Text = "") Then
        MsgBox "Please specify Buffer Length"
        Exit Sub
    End If
    If (txtBufferWidth.Text = "") Then
        MsgBox "Please specify Buffer Width"
        Exit Sub
    End If
    
    If (txtBufferRoughness.Text = "") Then
        MsgBox "Please specify Buffer Roughness"
        Exit Sub
    End If
    If (txtBufferSlope.Text = "") Then
        MsgBox "Please specify Buffer Slope"
        Exit Sub
    End If
    
    '** get the working directory
    Dim pworkingdir As String
    pworkingdir = gBufferStripDetailDict.Item("Working Directory")
        
    '** save all properties to the input file
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile pworkingdir & "/" & pFlowPropsFile, True
    If (fso.FileExists(pworkingdir & "/" & pFlowPropsFile) = False) Then
        MsgBox "Error creating " & pFlowPropsFile & " file."
        Exit Sub
    End If

    '** write these properties now
''    Dim pFile As TextStream
''    Set pFile = fso.OpenTextFile(pInputFileName, ForWriting, True, TristateUseDefault)
''
    '** close the form
    Unload Me
    
    '** set this value in the template form
    frmBufferStripTemplate.txtFlowInputs = pFlowPropsFile
    
    GoTo CleanUp
    
CleanUp:
    Set pFile = Nothing
    Set fso = Nothing
End Sub
