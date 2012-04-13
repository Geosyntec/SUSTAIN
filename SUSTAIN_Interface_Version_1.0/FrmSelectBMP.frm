VERSION 5.00
Begin VB.Form FrmSelectBMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select BMP Type"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSelectBMP.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTemplate 
      Caption         =   "Use Template Parameters"
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPercentDA 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Text            =   "100"
      Top             =   540
      Width           =   615
   End
   Begin VB.ComboBox listExistBMPs 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox cbSplitter 
      Caption         =   "Splitter"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox cbAssessPoint 
      Caption         =   "Assessment Point"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdRedefine 
      Caption         =   "Redefine Parameters"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblPercentDA 
      Caption         =   "Percentage of Drainage Area (%):"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Select an Existing BMP"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "FrmSelectBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pBmpIdDictionary As Scripting.Dictionary
Private pBmpTypeDictionary As Scripting.Dictionary


Private Sub cmdRedefine_Click()
    AddBMPParameters
    gDisplayBMPTemplate = True
End Sub

Public Sub AddBMPParameters()
On Error GoTo ShowError
    Dim pBMPName As String
    pBMPName = listExistBMPs.Text
    Dim pBMPID As Integer
    pBMPID = pBmpIdDictionary.Item(pBMPName)
    
    gNewBMPName = pBMPName
    If (FrmSelectBMP.cbSplitter.value = 1) Then
        bSplitter = True
    Else
        bSplitter = False
    End If
    
    Dim pPercentDAStr As String
    If (txtPercentDA.Visible = True) Then
        pPercentDAStr = txtPercentDA.Text
    End If
      
    'Create a dictionary to store the property names and property values
    If gBMPTypeToolbox = "Aggregate" Then
        Set gBMPDetailDict = GetAggBMPPropDict(pBMPName)
    Else
        Set gBMPDetailDict = GetBMPPropDict(pBMPID)
    End If
    
    gNewBMPType = pBmpTypeDictionary.Item(pBMPName)
    gBMPDetailDict.add "isAssessmentPoint", "False"
    
    'If the bmp is green roof or porous pavement, write the percentage of drainage area
    If (gNewBMPType = "GreenRoof" Or gNewBMPType = "PorousPavement") Then
        If (Not IsNumeric(pPercentDAStr)) Then
            MsgBox "Percentage should be a valid number", vbExclamation
            Exit Sub
        End If
        gBMPDetailDict.add "PercentDA", CDbl(pPercentDAStr)
    End If

    
    'Close the form
     Unload Me
     Exit Sub
ShowError:
    MsgBox "Error in AddBMPParameters: " & Err.description
End Sub

Private Sub cmdTemplate_Click()
    AddBMPParameters
    gDisplayBMPTemplate = False
End Sub

Public Sub Form_Initialize()
    
    Dim pBMPTypesTable As iTable
    If gBMPTypeToolbox = "Aggregate" Then
        Set pBMPTypesTable = GetInputDataTable("BMPDefaults")
    Else
        Set pBMPTypesTable = GetInputDataTable("BMPTypes")
    End If
    
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    If gBMPTypeToolbox = "Aggregate" Then
        pNameIndex = pBMPTypesTable.FindField("PropValue")
    Else
        pNameIndex = pBMPTypesTable.FindField("Name")
    End If
    Dim pTypeIndex As Long
    pTypeIndex = pBMPTypesTable.FindField("Type")
 
    Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    Set pBmpTypeDictionary = CreateObject("Scripting.Dictionary")

    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    If gBMPTypeToolbox = "Aggregate" Then
        pQueryFilter.WhereClause = "PropName='Type' And PropValue LIKE '%Aggregate%' ORDER BY PropValue"
    Else
        pQueryFilter.WhereClause = "TYPE = '" & gBMPTypeToolbox & "'"
    End If
    
    Set pCursor = pBMPTypesTable.Search(pQueryFilter, False)
    
'    Dim pSelRowCount As Long
'    pSelRowCount = pBMPTypesTable.RowCount(pQueryFilter)
'    Dim pBmpIdCount As Integer
'    pBmpIdCount = 0
    listExistBMPs.Clear
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        If gBMPTypeToolbox = "Aggregate" Then
            If Not pBmpIdDictionary.Exists(pRow.value(pNameIndex)) Then
                pBmpIdDictionary.add pRow.value(pNameIndex), pRow.value(pIDindex)
                pBmpTypeDictionary.add pRow.value(pNameIndex), "Aggregate"
                listExistBMPs.AddItem pRow.value(pNameIndex) ', pBmpIdCount
            End If
        Else
            pBmpIdDictionary.add pRow.value(pNameIndex), pRow.value(pIDindex)
            pBmpTypeDictionary.add pRow.value(pNameIndex), pRow.value(pTypeIndex)
            listExistBMPs.AddItem pRow.value(pNameIndex) ', pBmpIdCount
        End If
        Set pRow = pCursor.NextRow
    Loop
    listExistBMPs.ListIndex = 0
'
'    'Hide the List of existing bmps if only one bmp type is defined
'    If (listExistBMPs.ListCount = 1) Then
'        listExistBMPs.Visible = False
'    Else
'        listExistBMPs.Visible = True
'    End If
    
    'Hide the percentage label and text if bmp type is
    'not green roof or porous pavement
    If (gBMPTypeToolbox = "GreenRoof" Or gBMPTypeToolbox = "PorousPavement") Then
        lblPercentDA.Visible = True
        txtPercentDA.Visible = True
    Else
        lblPercentDA.Visible = False
        txtPercentDA.Visible = False
    End If
    
CleanUp:
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub

Private Sub listExistBMPs_Click()
    If listExistBMPs.Text = "AssessmentPoint" Then
        cbAssessPoint.Enabled = False
        cbSplitter.Enabled = False
    Else
        cbAssessPoint.Enabled = True
        cbSplitter.Enabled = True
    End If
End Sub
