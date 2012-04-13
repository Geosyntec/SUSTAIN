VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRunDLL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMP Optimization"
   ClientHeight    =   1920
   ClientLeft      =   5160
   ClientTop       =   3585
   ClientWidth     =   7455
   Icon            =   "FrmRunDLL.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   7455
   Begin VB.PictureBox Warning 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      Picture         =   "FrmRunDLL.frx":08CA
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6840
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   900
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   900
   End
   Begin VB.TextBox txtInputPath 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Note: Please save your project before running simulation."
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblFilePath 
      Caption         =   "Input File Path : "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmRunDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StartSimulation Lib "SUSTAINOPT.dll" (ByVal FilePath As String, ByVal bestPopID As String) As Boolean
'Private Declare Function StartSimulation Lib "D:\SUSTAIN\SUSTAIN_OPT\Release\BMPOPT.dll" (ByVal FilePath As String) As Boolean

Private Sub btnRun_Click()
    Dim strFilePath As String
    strFilePath = txtInputPath.Text
    
    'Close the form
    Unload Me
    
    Dim bRes As Boolean
    'The second argument in StartSimulation is blank
    bRes = StartSimulation(strFilePath, "") ' , maxIter, solNum, maxRun, b1, b2, pSize, localSearch, CostLimit, evalMode)

    If (bRes = False) Then
        MsgBox "Simulation failed !", vbExclamation
        Exit Sub
    End If
    
    
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo ErrorHandler

    Dim pattern As String
    pattern = "Input File (*.inp)|*.inp"
    CommonDialog.Filter = pattern
    CommonDialog.CancelError = True
    
    Dim pInputFileName As String
    CommonDialog.ShowOpen
    
    If (Err <> cdlCancel) Then
        pInputFileName = CommonDialog.FileName
        FrmRunDLL.txtInputPath = pInputFileName
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error while reading the input file: " & Err.description

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If (Trim(pInputFileName) = "") Then
        Dim pTable As iTable
        Set pTable = GetInputDataTable("OptimizationDetail")
        If Not (pTable Is Nothing) Then
            Dim pQueryFilter As IQueryFilter
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "PropName = 'InputFile'"
            Dim pCursor As ICursor
            Set pCursor = pTable.Search(pQueryFilter, True)
            Dim pRow As iRow
            Set pRow = pCursor.NextRow
            Dim iPropValue As Long
            iPropValue = pTable.FindField("PropValue")
            If Not (pRow Is Nothing) Then
                pInputFileName = pRow.value(iPropValue)
            End If
            Set pRow = Nothing
            Set pCursor = Nothing
            Set pQueryFilter = Nothing
            Set pTable = Nothing
        End If
    End If
    txtInputPath.Text = pInputFileName
    
End Sub


