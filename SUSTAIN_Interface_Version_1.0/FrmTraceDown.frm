VERSION 5.00
Begin VB.Form FrmTraceDown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Buffer Strip Location"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "FrmTraceDown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRedefineParams 
      Caption         =   "Redefine Parameters"
      Height          =   555
      Left            =   720
      TabIndex        =   17
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame frameAdditionalInfo 
      Caption         =   "Template Info"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   5055
      Begin VB.OptionButton optionRight 
         Caption         =   "Right Bank"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optionLeft 
         Caption         =   "Left Bank"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbVFSTypes 
         Height          =   315
         ItemData        =   "FrmTraceDown.frx":08CA
         Left            =   240
         List            =   "FrmTraceDown.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame frameTrace 
      Caption         =   "Ending Point"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   5055
      Begin VB.ComboBox cbxTraceBMP 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmTraceDown.frx":08CE
         Left            =   3480
         List            =   "FrmTraceDown.frx":08D8
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optTraceDown 
         Caption         =   "&Trace down stream"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optTraceJunction 
         Caption         =   "Trace to next in-stream &BMP/junction"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox cbxToEnd 
         Caption         =   "to the end of &downstream"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox tbxDistance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTrace 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblUnit 
         Caption         =   "meters"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTemplateParams 
      Caption         =   "Use Template Parameters"
      Default         =   -1  'True
      Height          =   555
      Left            =   2400
      TabIndex        =   12
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   4080
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame frameStart 
      Caption         =   "Starting Point"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cbxSnapBMP 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmTraceDown.frx":08EF
         Left            =   3480
         List            =   "FrmTraceDown.frx":08F9
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optNearestNode 
         Caption         =   "Snap to nearest &end node of the stream"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton optNearestPoint 
         Caption         =   "Snap to nearest &point along the stream"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton optNearestJunction 
         Caption         =   "Snap to nearest &in-stream BMP/Junction"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3255
      End
   End
End
Attribute VB_Name = "FrmTraceDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created: 07/14/2005 - Haihong Yang create this module.
'
'******************************************************************************
Option Explicit

Public bContinue As Boolean
Public nSnapOption As Integer
Public nTraceOption As Integer
Public fTraceDistance As Double
Public bTraceToEnd As Boolean
Public strSnapBMPType As String
Public strTraceBMPType As String
Public pVFSName As String
Public bBankSide As String
Public bVFSDefaultID As Integer
Public bRedefineParams As Boolean

Private Sub DefineVFSOptions()

  If optNearestPoint.value = True Then
    nSnapOption = SNAP_NEAREST_POINT
  ElseIf optNearestNode.value = True Then
    nSnapOption = SNAP_NEAREST_NODE
  Else
    nSnapOption = SNAP_NEAREST_JUNCTION
  End If
  
  If optTraceDown.value = True Then
    nTraceOption = TRACE_DOWN
  ElseIf optTraceJunction.value = True Then
    nTraceOption = TRACE_JUNCTION
  End If
  
  fTraceDistance = -1
  If IsNumeric(tbxDistance) Then
    fTraceDistance = CDbl(tbxDistance) ' unit is meters
  End If
  
  If nTraceOption = TRACE_DOWN And fTraceDistance <= 0 Then
    MsgBox "Please specify a tracing distance as a positive number.", vbExclamation + vbOKOnly
    bContinue = False ' Arun Raj
    Exit Sub
  End If
  
  bTraceToEnd = (cbxToEnd.value = vbChecked)
  
  strSnapBMPType = ""
  If nSnapOption = SNAP_NEAREST_JUNCTION Then
    strSnapBMPType = cbxSnapBMP.Text
  End If
  
  strTraceBMPType = ""
  If nTraceOption = TRACE_JUNCTION Then
    strTraceBMPType = cbxTraceBMP.Text
  End If
  
  pVFSName = cmbVFSTypes.Text
  If (optionLeft.value = True) Then
    bBankSide = "Left"
  ElseIf (optionRight.value = True) Then
    bBankSide = "Right"
  End If
  
  bVFSDefaultID = cmbVFSTypes.ListIndex
  
  bContinue = True

End Sub

Private Sub cmdCancel_Click()
  bContinue = False
  Unload Me
End Sub

Private Sub cmdRedefineParams_Click()
    '** call the function to retrieve params
    DefineVFSOptions
    bRedefineParams = True
    '** Close the form
    If (bContinue = True) Then
        Unload Me
    End If
End Sub

Private Sub cmdTemplateParams_Click()
    
    '** call the function to retrieve params
    DefineVFSOptions
    bRedefineParams = False
    '** Close the form
    If (bContinue = True) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '** clear the combo box
    'cmbVFSTypes.Clear
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDefaults")
    If Not (pTable Is Nothing) Then
        Dim pQueryFilter As IQueryFilter
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "PropName = 'Name'"
        Dim pCursor As ICursor
        Set pCursor = pTable.Search(pQueryFilter, True)
        Dim pRow As iRow
        Set pRow = pCursor.NextRow
        Dim iIDFld As Long
        iIDFld = pTable.FindField("ID")
        Dim iPropValueFld As Long
        iPropValueFld = pTable.FindField("PropValue")
        Dim pNameValue As String
        Do While Not pRow Is Nothing
            pNameValue = pRow.value(iPropValueFld)
            '** Add to the cmbbox
            cmbVFSTypes.AddItem pNameValue
            cmbVFSTypes.ItemData(cmbVFSTypes.NewIndex) = pRow.value(iIDFld)
            Set pRow = pCursor.NextRow
        Loop
    End If

  cmbVFSTypes.ListIndex = 0
  cbxSnapBMP.ListIndex = 0
  cbxTraceBMP.ListIndex = 0
  
  optionLeft.value = True
End Sub

Private Sub optNearestJunction_Click()
  cbxSnapBMP.Enabled = True
End Sub

Private Sub optNearestNode_Click()
  cbxSnapBMP.Enabled = False
End Sub

Private Sub optNearestPoint_Click()
  cbxSnapBMP.Enabled = False
End Sub

Private Sub optTraceDown_Click()
  cbxTraceBMP.Enabled = False
End Sub

Private Sub optTraceJunction_Click()
  cbxTraceBMP.Enabled = True
End Sub
