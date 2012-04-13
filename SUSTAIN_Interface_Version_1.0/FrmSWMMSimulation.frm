VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSWMMSimulation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Simulation Parameters"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "FrmSWMMSimulation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPreDev 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   3885
   End
   Begin VB.CommandButton cmdBrowsePreInputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseInputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtInputFile 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label Label3 
      Caption         =   "Predeveloped Input File:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Input File:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Note: Please save your project before running simulation."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
End
Attribute VB_Name = "FrmSWMMSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_pInputFile As String
Private m_pReportFile As String
Private m_pOutputFile As String
Private m_pSimulationTime As Long

Private Declare Function StartLandSimulation Lib "SUSTAINOPT.dll" (ByVal PreFile As String, ByVal PostFile As String) As Boolean

Private Sub cmdBrowseInputFile_Click()
    
    CommonDialog.Filter = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.CancelError = False
    CommonDialog.ShowOpen
    FrmSWMMSimulation.txtInputFile.Text = CommonDialog.FileName

End Sub


Private Sub cmdBrowsePreInputFile_Click()
    
    CommonDialog.Filter = "LAND Simulation Input File (*.inp)|*.inp"
    CommonDialog.CancelError = False
    CommonDialog.ShowOpen
    FrmSWMMSimulation.txtPreDev.Text = CommonDialog.FileName
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()

    Dim pInputFile As String
    Dim pPreFile As String
    pInputFile = txtInputFile.Text
    If (Trim(pInputFile) = "") Then
        MsgBox "Please specify SWMM Simulation Input file to continue."
        Exit Sub
    End If
    pPreFile = txtPreDev.Text
    If (Trim(pPreFile) = "") Then
        MsgBox "Please specify SWMM Simulation pre-developed file to continue."
        Exit Sub
    End If
    
    '** get all simulation parameters
    GetSimulationParameters pInputFile
    
    '** close the form
    Unload Me
    
    '** call the dll to run
    Dim bRes As Boolean
    bRes = StartLandSimulation(pPreFile, pInputFile)
    If (bRes = False) Then
        MsgBox "Simulation failed !", vbExclamation
        Exit Sub
    End If
    

End Sub


Private Sub GetSimulationParameters(pInputFileName As String)
    
    Dim pStrStartDate As String
    Dim pStrEndDate As String
    Dim pStrOutFileLine As String
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not (fso.FileExists(pInputFileName)) Then
        MsgBox "Input file not found."
        Exit Sub
    End If
    
    Dim pInpFileStream
    Set pInpFileStream = fso.OpenTextFile(pInputFileName, ForReading)
    
    Dim pInpFileData As String
    pInpFileData = pInpFileStream.ReadAll
    
    Dim pDataLines
    pDataLines = Split(pInpFileData, vbNewLine, , vbTextCompare)
    Dim pDataLine As String
    Dim iCount As Integer
    For iCount = 0 To UBound(pDataLines) - 1
        pDataLine = Trim(pDataLines(iCount))
        '** get the start date
        If (Left(pDataLine, 10) = "START_DATE") Then
            pStrStartDate = Trim(Replace(pDataLine, "START_DATE", ""))
        End If
        '** get the end date
        If (Left(pDataLine, 8) = "END_DATE") Then
            pStrEndDate = Trim(Replace(pDataLine, "END_DATE", ""))
        End If
        '** get the save outflows file
        If (Left(pDataLine, 13) = "SAVE OUTFLOWS") Then
            pStrOutFileLine = Trim(Replace(pDataLine, "SAVE OUTFLOWS", ""))
            '** replace the extra quotation marks
        End If
                    
    Next iCount
    
    '** clean up memory
    pInpFileStream.Close
    Set fso = Nothing
    
    '** Get the simulation difference in days
    m_pSimulationTime = DateDiff("d", CDate(pStrStartDate), CDate(pStrEndDate)) + 1
    
    '** define the input, output and report file names
    m_pInputFile = pInputFileName
    m_pOutputFile = Replace(pStrOutFileLine, ".txt", ".out")
    m_pReportFile = Replace(pStrOutFileLine, ".txt", ".rpt")

End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    txtInputFile.Text = gPostDevfile
    txtPreDev.Text = gPreDevfile
End Sub
