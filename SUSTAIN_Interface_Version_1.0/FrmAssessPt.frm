VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmAssessPt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Assessment Point Evaluation Factors"
   ClientHeight    =   4755
   ClientLeft      =   2760
   ClientTop       =   3285
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAssessPt.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9255
   Begin VB.Frame Frame2 
      Caption         =   "Annual Average Values"
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8895
      Begin MSDBGrid.DBGrid DBGridAssess 
         Height          =   2895
         Left            =   120
         OleObjectBlob   =   "FrmAssessPt.frx":08CA
         TabIndex        =   3
         Top             =   240
         Width           =   8550
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   480
      Left            =   3360
      TabIndex        =   0
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   4680
      TabIndex        =   1
      Top             =   3960
      Width           =   960
   End
End
Attribute VB_Name = "frmAssessPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Set gBMPDetailDict = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ShowError:
 
'    Dim pErrorMessage As String
'    pErrorMessage = ""
'    Dim pControl
'
'    pErrorMessage = "Enter a valid input for: "
'    If isFlowVEval.value = 1 And (Trim(aa_fv.Text) = "" Or Not (IsNumeric(Trim(aa_fv.Text)))) Then
'        pErrorMessage = pErrorMessage & "flow volume" & vbTab
'    End If
'    If isFlowPercEval.value = 1 And (Trim(aa_fv_Exist.Text) = "" Or Not (IsNumeric(Trim(aa_fv_Exist.Text)))) Then
'        pErrorMessage = pErrorMessage & "flow volume percentage" & vbTab
'    End If
'    If isTssLoadEval.value = 1 And (Trim(aa_tssLd.Text) = "" Or Not (IsNumeric(Trim(aa_tssLd.Text)))) Then
'        pErrorMessage = pErrorMessage & "TSS load" & vbTab
'    End If
'    If isTssPercEval.value = 1 And (Trim(aa_tssLd_Exist.Text) = "" Or Not (IsNumeric(Trim(aa_tssLd_Exist.Text)))) Then
'        pErrorMessage = pErrorMessage & "TSS percentage"
'    End If
'
'    If (pErrorMessage <> "Enter a valid input for: ") Then
'        MsgBox pErrorMessage
'        Exit Sub
'    End If
'
'    Set pControl = Nothing
'    For Each pControl In Controls
'        If ((TypeOf pControl Is TextBox) And (pControl.Enabled)) Then
'              gBMPDetailDict.Add pControl.Name, pControl.Text
'        End If
'        If ((TypeOf pControl Is CheckBox) And (pControl.Enabled)) Then
'            If pControl.value = 1 Then
'                gBMPDetailDict.Add pControl.Name, "True"
'            Else
'                gBMPDetailDict.Add pControl.Name, "False"
'            End If
'        End If
'    Next pControl
'    Set pControl = Nothing
    
    With gAssessInfos(0)
        gBMPDetailDict.Add "isFlowVolEval", CBool(.isTargetEval)
        gBMPDetailDict.Add "isFlowRednEval", CBool(.isRedEval)
        If .isTargetEval Then
            gBMPDetailDict.Add "FlowVol", CDbl(.Target)
        End If
        If .isRedEval Then
            gBMPDetailDict.Add "FlowRedn", CDbl(.Reduction)
        End If
    End With
    Dim incr As Integer
    For incr = 1 To gMaxPollutants
        With gAssessInfos(incr)
            gBMPDetailDict.Add "isParam" & incr & "LoadEval", CBool(.isTargetEval)
            gBMPDetailDict.Add "isParam" & incr & "RednEval", CBool(.isRedEval)
            If .isTargetEval Then
                gBMPDetailDict.Add "Param" & incr & "Load", CDbl(.Target)
            End If
            If .isRedEval Then
                gBMPDetailDict.Add "Param" & incr & "Redn", CDbl(.Reduction)
            End If
        End With
    Next incr
        
    Unload Me
    Exit Sub
ShowError:
    MsgBox "Assessment Point Info form 1", Err.description
End Sub



Private Sub DBGridAssess_ButtonClick(ByVal ColIndex As Integer)
    Dim curVal As Boolean
    curVal = DBGridAssess.Columns(ColIndex).value
    DBGridAssess.Columns(ColIndex).value = Not (curVal)
End Sub



'Private Sub isFlowPercEval_Click()
'    If isFlowPercEval.value = 1 Then
'        aa_fv_Exist.Enabled = True
'        aa_fv_Exist.BackColor = vbWhite
'    Else
'        aa_fv_Exist.Enabled = False
'        aa_fv_Exist.BackColor = &H80000016
'    End If
'
'End Sub
'
'
'Private Sub isFlowVEval_Click()
'    If isFlowVEval.value = 1 Then
'        aa_fv.Enabled = True
'        aa_fv.BackColor = vbWhite
'    Else
'        aa_fv.Enabled = False
'        aa_fv.BackColor = &H80000016
'    End If
'End Sub
'
'Private Sub isTssLoadEval_Click()
'    If isTssLoadEval.value = 1 Then
'        aa_tssLd.Enabled = True
'        aa_tssLd.BackColor = vbWhite
'    Else
'        aa_tssLd.Enabled = False
'        aa_tssLd.BackColor = &H80000016
'    End If
'End Sub
'
'Private Sub isTssPercEval_Click()
'    If isTssPercEval.value = 1 Then
'        aa_tssLd_Exist.Enabled = True
'        aa_tssLd_Exist.BackColor = vbWhite
'    Else
'        aa_tssLd_Exist.Enabled = False
'        aa_tssLd_Exist.BackColor = &H80000016
'    End If
'End Sub

Private Sub Form_Load()
    Dim Col1
    Set Col1 = DBGridAssess.Columns(0)
    Col1.Locked = True
End Sub


Private Sub DBGridAssess_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim fi As Integer
''For fi = 0 To 5000
''    fi = fi
''Next fi
''Pause for 5 seconds
'Pause (0.05)

Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer

    ' See which direction to read.
    If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    
    ' See if StartLocation is Null.
    If IsNull(StartLocation) Then
        ' Read from the end or beginning of
        ' the data.
        If ReadPriorRows Then
            ' Read backwards from the end.
            row_num = RowBuf.RowCount - 1
        Else
            ' Read from the beginning.
            row_num = 0
        End If
    Else
        ' See where to start reading.
        row_num = CLng(StartLocation) + dr
    End If
    
    ' Copy data from the array into RowBuf.
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        ' Do not run beyond the end of the data.
        If row_num < 0 Or row_num > (gMaxPollutants) Then Exit For
        
        ' Copy the data into the row buffer.
        With gAssessInfos(row_num)
            RowBuf.value(r, 0) = .Factor
            RowBuf.value(r, 1) = .Unit
            RowBuf.value(r, 2) = .isTargetEval
            RowBuf.value(r, 3) = .Target
            RowBuf.value(r, 4) = .isRedEval
            RowBuf.value(r, 5) = Math.Round(.Reduction, 0)
        End With

        ' Use row_num as the bookmark.
        RowBuf.Bookmark(r) = row_num
        
        row_num = row_num + dr
        rows_returned = rows_returned + 1
    Next r

    ' Set the number of rows returned.
    RowBuf.RowCount = rows_returned
    
End Sub



' Save data updated by the control.
Private Sub DBGridAssess_UnboundWriteData(ByVal RowBuf As MSDBGrid.RowBuffer, WriteLocation As Variant)
    ' Update only the values that have changed.
    With gAssessInfos(CInt(WriteLocation))
        If Not IsNull(RowBuf.value(0, 2)) Then
            .isTargetEval = RowBuf.value(0, 2)
        End If
        If Not IsNull(RowBuf.value(0, 3)) Then
            .Target = RowBuf.value(0, 3)
        End If
        If Not IsNull(RowBuf.value(0, 4)) Then
            .isRedEval = RowBuf.value(0, 4)
        End If
        If Not IsNull(RowBuf.value(0, 5)) Then
            
            If StringContains(CStr(RowBuf.value(0, 5)), "%") Then
                Dim rednStr As String
                rednStr = Replace(CStr(RowBuf.value(0, 5)), "%", "")
                .Reduction = CDbl(Math.Round(rednStr / 100, 0))
            ElseIf IsNumeric(CStr(RowBuf.value(0, 5))) Then
                .Reduction = CDbl(Math.Round(RowBuf.value(0, 5), 0))
            End If
        End If
    End With
End Sub


