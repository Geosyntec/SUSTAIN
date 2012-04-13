VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmAggBmpLuDist 
   Caption         =   "Aggregate BMP Landuse Distribution"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   Icon            =   "FrmAggBmpLuDist.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4995
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbxSws 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGridLuDist 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Select Subwatershed"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Land Use Distribution (%)"
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "FrmAggBmpLuDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private pModifyFlag As Boolean
Private curCbxSwsLI As Integer
Private bCheckSws As Boolean
Private bmpIDList

Private Sub cbxSws_Click()
    If pModifyFlag Then
        If bCheckSws Then
            bCheckSws = False
            If MsgBox("There are changes that are not saved.  Do you want to continue?", vbInformation + vbYesNo, "SUSTAIN") = vbNo Then
                cbxSws.ListIndex = curCbxSwsLI
                Exit Sub
            Else
                Call UpdateDataGrid
            End If
        End If
    Else
        Call UpdateDataGrid
    End If
    bCheckSws = True
End Sub

Private Sub cmdClose_Click()
    pModifyFlag = False
    bCheckSws = False
    Set gLuGroupIdDict = Nothing
    Set gLuIdGroupDict = Nothing
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ShowError

    Dim rs As ADODB.Recordset
    Set rs = DataGridLuDist.DataSource
    
    Dim dsIds() As String
    ReDim dsIds(rs.Fields.Count - 3)
    
    Dim i As Integer, j As Integer
    Dim isDsFound As Boolean
    
'    rs.MoveLast
'    For i = 2 To rs.Fields.Count - 2
'        dsIds(i - 2) = rs.Fields(i).value
'    Next
'
'    For i = 0 To UBound(bmpIDList)
'        If bmpIDList(i) = dsIds(i) Then
'            MsgBox "Downstream and upstream BMP IDs are the same for " & bmpIDList(i)
'            Exit Sub
'        End If
'    Next
'
'
'    For i = 0 To UBound(dsIds)
'        isDsFound = False
'        If dsIds(i) <> "0" Then
'            For j = 0 To UBound(bmpIDList)
'                If bmpIDList(j) = dsIds(i) Then isDsFound = True
'            Next
'            If Not isDsFound Then
'                MsgBox "Downstream IDs not from the existing BMP list for " & bmpIDList(i)
'                Exit Sub
'            End If
'        End If
'    Next
    
'    MsgBox "Checked ds"
    
    If gLuGroupIdDict Is Nothing Then Call SetLuGroupIDDict
    gLuGroupIdDict.Item("BMPID") = -98
    gLuGroupIdDict.Item("Downstream ID") = -99
    
    Dim lu As String
    
    Dim pBMPID As Integer
    If cbxSws.ListCount = 0 Then Err.Raise "No subwatershed has aggregate BMPs"
    pBMPID = cbxSws.ItemData(cbxSws.ListIndex)
     
'    MsgBox "Got BMPID"
    
    Dim pAgBmpLuDTable As esriGeoDatabase.iTable
    Set pAgBmpLuDTable = GetInputDataTable("AgLuDistribution")
    
    If pAgBmpLuDTable Is Nothing Then
        Set pAgBmpLuDTable = CreateAgBmpLuDistTableDBF("AgLuDistribution")
        If pAgBmpLuDTable Is Nothing Then
            MsgBox "AgLuDistribution table is missing and can not be created. Try again"
            Exit Sub
        End If
        AddTableToMap pAgBmpLuDTable
    End If
    
'    MsgBox "Got AgLuDistribution"
    
    Dim pRow As esriGeoDatabase.iRow
    Dim pCursor As esriGeoDatabase.ICursor
       
    Dim iFldBmpID As Integer
    Dim iFldLuGroup As Integer
    Dim iFldLuGroupID As Integer
    Dim iFldArea As Integer
    Dim iFldAreaDis As Integer
    
    iFldBmpID = pAgBmpLuDTable.FindField("BMPID")
    iFldLuGroup = pAgBmpLuDTable.FindField("LuGroup")
    iFldLuGroupID = pAgBmpLuDTable.FindField("LuGrpID")
    iFldArea = pAgBmpLuDTable.FindField("TotalArea")
    iFldAreaDis = pAgBmpLuDTable.FindField("AreaDist")
    
'    MsgBox "Got Fields"
    
    If iFldBmpID < 0 Or iFldLuGroup < 0 Or iFldLuGroupID < 0 _
        Or iFldArea < 0 Or iFldAreaDis < 0 Then
        MsgBox "Missing required field(s) [BMPID, LuGroup, LuGroupId, TotalArea, iFldAreaDis] in AgLuDistribution table"
        Exit Sub
    End If
    
    Dim strIncomLus As String
    strIncomLus = ""
    
    Dim sumPerc As Double
    
    rs.MoveFirst
'    MsgBox "1"
    rs.MoveNext
    rs.MoveNext
    rs.MoveNext
'    MsgBox "2"
    Do Until rs.EOF
        lu = CStr(rs("Landuse"))
        If lu <> "Downstream ID" Then
            sumPerc = 0
            For i = 2 To rs.Fields.Count - 1
                sumPerc = sumPerc + CDbl(rs.Fields(i).value)
            Next
            If sumPerc <> 100 Then strIncomLus = strIncomLus & lu & vbNewLine
        End If
        rs.MoveNext
    Loop
    
'    MsgBox "3"
    If strIncomLus <> "" Then
        MsgBox "Total contributions to following landuses do not add up to 100%" & vbNewLine & strIncomLus
        Exit Sub
    End If
    
    Dim boolContinue
    
'    If Not IsNumeric(Trim(txtUpstream.Text)) Then
'        boolContinue = MsgBox("No valid BMPID for 'BMPID Receiving Discharge from Upstream BMP'. Do you want to continue ?", vbYesNo, "Upstream BMP")
'        If boolContinue = vbNo Then Exit Sub
'    End If
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "BMPID = " & pBMPID
    pAgBmpLuDTable.DeleteSearchedRows pQueryFilter
    
'    MsgBox "4"
    rs.MoveFirst
    Set pRow = pAgBmpLuDTable.CreateRow
    lu = CStr(rs("Landuse"))
    pRow.value(iFldBmpID) = pBMPID
    pRow.value(iFldLuGroup) = lu
    pRow.value(iFldLuGroupID) = CInt(gLuGroupIdDict.Item(lu)) 'rs("LanduseID")
    pRow.value(iFldArea) = CDbl(rs("Area"))
    
    Dim strAreaDist As String
'    strAreaDist = rs.Fields(2).value
'    For i = 3 To rs.Fields.Count - 1
'        strAreaDist = strAreaDist & "," & rs.Fields(i).value
'    Next
    
'    MsgBox "5"
    strAreaDist = bmpIDList(0)
    For i = 1 To UBound(bmpIDList)
        strAreaDist = strAreaDist & "," & bmpIDList(i)
    Next
    strAreaDist = strAreaDist & ",0"
    pRow.value(iFldAreaDis) = strAreaDist
    pRow.Store
    
    rs.MoveNext
    rs.MoveNext
    rs.MoveNext
    
'    MsgBox "6"
    
    Do Until rs.EOF
        Set pRow = pAgBmpLuDTable.CreateRow
        lu = CStr(rs("Landuse"))
        pRow.value(iFldBmpID) = pBMPID
        pRow.value(iFldLuGroup) = lu
        pRow.value(iFldLuGroupID) = CInt(gLuGroupIdDict.Item(lu)) 'rs("LanduseID")
        pRow.value(iFldArea) = CDbl(rs("Area"))
        strAreaDist = rs.Fields(2).value
        For i = 3 To rs.Fields.Count - 1
            strAreaDist = strAreaDist & "," & rs.Fields(i).value
        Next
        pRow.value(iFldAreaDis) = strAreaDist
        pRow.Store
        rs.MoveNext
    Loop
    
'    MsgBox "7"
'    Set pRow = pAgBmpLuDTable.CreateRow
'    pRow.value(iFldBmpID) = pBMPID
'    pRow.value(iFldLuGroup) = "UpstreamBMP"
'    pRow.value(iFldLuGroupID) = -97
'    pRow.value(iFldArea) = 0
'    pRow.value(iFldAreaDis) = txtUpstream.Text
'    pRow.Store
    
    pModifyFlag = False
    'Unload Me
    GoTo CleanUp
ShowError:
    MsgBox "Error saving the aggregate BMP land use distribution: " & Err.description
CleanUp:
    Set rs = Nothing
    Set pAgBmpLuDTable = Nothing
    Set pRow = Nothing
End Sub

'Public Sub InitializeDataGrid(luDistDict As Scripting.Dictionary)
Public Sub UpdateDataGrid() '(luDistDict As Scripting.Dictionary)
On Error GoTo ShowError

    Dim lineNumber As String
    lineNumber = "0"
    
    If gLuGroupIdDict Is Nothing Then Call SetLuGroupIDDict
    gLuGroupIdDict.Item("BMPID") = -98
    gLuGroupIdDict.Item("Downstream ID") = -99
        
    Dim luKey, luID As Integer
    Dim pBMPID As Integer
    Dim pSwsID As Integer
    pBMPID = cbxSws.ItemData(cbxSws.ListIndex)
    pSwsID = cbxSws.List(cbxSws.ListIndex)
    
    Dim aggBmpTypeDict As Scripting.Dictionary
    Set aggBmpTypeDict = GetAggBMPTypes(pBMPID)
    
    If aggBmpTypeDict Is Nothing Then MsgBox "No information about the types of BMPs for BMPID :" & pBMPID: Exit Sub

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim pAggBmpName As String
    Dim pAggBmpType As String
    Dim pAggBmpId As Integer
    Dim pAggBmpCat As String
    Dim i As Integer
    
    lineNumber = "1"
    lineNumber = "2"
    
    Dim strAreaDist As String
    Dim luArea As Double
    Dim listAreaDist

    Dim luDistDict As Scripting.Dictionary
    Set luDistDict = Get_Agg_BMP_Lu_Distrib(pBMPID)
    
    Dim pSubWsLuGroupDict As Scripting.Dictionary
    Set pSubWsLuGroupDict = GetSubWsLuGroupDict(pSwsID) 'GetSubWsLuGroupDict(pSubWsID)
    If pSubWsLuGroupDict Is Nothing Then MsgBox "No land use distribution info for watershed :" & pSwsID: Exit Sub
    
    'lineNumber = "3"
    
    Dim errMessage As String
    
    ReDim listAreaDist(0)
    
    luID = -98
    If Not luDistDict Is Nothing Then
        If luDistDict.Exists(luID) Then
            strAreaDist = luDistDict.Item(luID)(1)
            'MsgBox "strAreaDist : " & strAreaDist
            If strAreaDist <> "" Then
                listAreaDist = Split(strAreaDist, ",")
            End If
        End If
        
        'Set the upstream ID
'        If luDistDict.Exists(-97) Then
'            txtUpstream.Text = luDistDict.Item(-97)(1)
'            luDistDict.Remove (-97)
'        Else
'            txtUpstream.Text = ""
'        End If
    End If
    
    Dim bmpInSameOrder As Boolean
    bmpInSameOrder = True
            
    If UBound(listAreaDist) = aggBmpTypeDict.Count Then
        For i = 0 To UBound(listAreaDist) - 1
            If Not aggBmpTypeDict.Exists(CInt(listAreaDist(i))) Then ' aggBmpTypeDict.keys(i)
                bmpInSameOrder = False
            End If
        Next
    Else
        bmpInSameOrder = False
    End If
        
    ' Setup the fields
    rs.Fields.Append "Landuse", adVarChar, 50
    rs.Fields.Append "Area", adDouble
    If bmpInSameOrder Then
        ReDim bmpIDList(UBound(listAreaDist) - 1)
        For i = 0 To UBound(listAreaDist) - 1
            pAggBmpId = listAreaDist(i)
            bmpIDList(i) = pAggBmpId
            pAggBmpName = aggBmpTypeDict.Item(pAggBmpId)(0)
            rs.Fields.Append pAggBmpName, adVarChar, 50
        Next
    Else
        ReDim bmpIDList(aggBmpTypeDict.Count - 1)
        For i = 0 To aggBmpTypeDict.Count - 1
            pAggBmpId = aggBmpTypeDict.keys(i)
            bmpIDList(i) = pAggBmpId
            pAggBmpName = aggBmpTypeDict.Item(pAggBmpId)(0)
            rs.Fields.Append pAggBmpName, adVarChar, 50
        Next
    End If
    rs.Fields.Append "Outlet", adVarChar, 50
    rs.CursorType = adOpenDynamic
    rs.Open
    
    rs.AddNew
    rs.Fields(0).value = "BMPID"
    If bmpInSameOrder Then
        For i = 0 To UBound(listAreaDist)
            rs.Fields(i + 2).value = listAreaDist(i)
        Next
    Else
        For i = 0 To aggBmpTypeDict.Count - 1
            'lineNumber = "6 aggBmpTypeDict i " & i
            pAggBmpId = aggBmpTypeDict.keys(i)
            rs.Fields(i + 2).value = CStr(pAggBmpId)
        Next
        rs.Fields(aggBmpTypeDict.Count + 2).value = "0"
    End If
        
    rs.AddNew
    rs.Fields(0).value = "Category"
    If bmpInSameOrder Then
        For i = 0 To UBound(listAreaDist) - 1
            pAggBmpId = listAreaDist(i)
            pAggBmpCat = Trim(aggBmpTypeDict.Item(pAggBmpId)(1))
            'lineNumber = "7_1 aggBmpTypeDict i " & i & "Category " & pAggBmpCat & "rs.Fields.Count " & rs.Fields.Count
            rs.Fields(i + 2).value = CStr(pAggBmpCat)
        Next
    Else
        For i = 0 To aggBmpTypeDict.Count - 1
            pAggBmpId = aggBmpTypeDict.keys(i)
            pAggBmpCat = Trim(aggBmpTypeDict.Item(pAggBmpId)(1))
            'lineNumber = "7 aggBmpTypeDict i " & i & "Category " & pAggBmpCat & "rs.Fields.Count " & rs.Fields.Count
            rs.Fields(i + 2).value = CStr(pAggBmpCat)
        Next
    End If
    'lineNumber = "7 aggBmpTypeDict i " & i & "Category Outlet"
    rs.Fields(aggBmpTypeDict.Count + 2).value = "Outlet"
    
    rs.AddNew
    rs.Fields(0).value = "BMPType"
    If bmpInSameOrder Then
        For i = 0 To UBound(listAreaDist) - 1
            pAggBmpId = listAreaDist(i)
            pAggBmpType = aggBmpTypeDict.Item(pAggBmpId)(2)
            'lineNumber = "8_1 aggBmpTypeDict i " & i & " pAggBmpType " & pAggBmpType
            rs.Fields(i + 2).value = CStr(pAggBmpType)
        Next
    Else
        For i = 0 To aggBmpTypeDict.Count - 1
            pAggBmpId = aggBmpTypeDict.keys(i)
            pAggBmpType = aggBmpTypeDict.Item(pAggBmpId)(2)
            'lineNumber = "8 aggBmpTypeDict i " & i & " pAggBmpType " & pAggBmpType
            rs.Fields(i + 2).value = CStr(pAggBmpType)
        Next
    End If
    'lineNumber = "8 aggBmpTypeDict i " & i & "Type Outlet"
    rs.Fields(aggBmpTypeDict.Count + 2).value = "Outlet"
    
    For Each luKey In pSubWsLuGroupDict
        'lineNumber = "4 lukey " & luKey
        rs.AddNew
        rs.Fields(0).value = CStr(luKey)
        rs.Fields(1).value = CDbl(FormatNumber(pSubWsLuGroupDict.Item(luKey), 2, vbFalse))
        strAreaDist = ""
        luID = CInt(gLuGroupIdDict.Item(CStr(luKey)))
        If Not luDistDict Is Nothing Then
            If luDistDict.Exists(luID) Then
                luArea = luDistDict.Item(luID)(0)
                strAreaDist = luDistDict.Item(luID)(1)
                If CDbl(FormatNumber(pSubWsLuGroupDict.Item(luKey), 2, vbFalse)) <> CDbl(FormatNumber(luArea, 2, vbFalse)) Then
                    If CDbl(luArea) > 0 Then errMessage = errMessage & vbNewLine & luKey & "From Table = " & FormatNumber(luArea, 2, vbFalse) & " Actual = " & FormatNumber(pSubWsLuGroupDict.Item(luKey), 2, vbFalse)
                End If
            End If
        End If
        
        If bmpInSameOrder Then
            If strAreaDist <> "" Then
                listAreaDist = Split(strAreaDist, ",")
                For i = 0 To UBound(listAreaDist)
                    rs.Fields(i + 2).value = CStr(FormatNumber(listAreaDist(i), 2, vbFalse))
                Next
            Else
                For i = 0 To rs.Fields.Count - 3
                    rs.Fields(i + 2).value = CStr(0)
                Next
            End If
        Else
            For i = 0 To rs.Fields.Count - 3
                rs.Fields(i + 2).value = CStr(0)
            Next
        End If
    Next
        
    'Add another line to enter downstream id
    rs.AddNew
    rs.Fields(0).value = "Downstream ID"
    If bmpInSameOrder Then
        If Not luDistDict Is Nothing Then
            If luDistDict.Exists(-99) Then
                strAreaDist = luDistDict.Item(-99)(1)
                If strAreaDist <> "" Then
                    listAreaDist = Split(strAreaDist, ",")
                    For i = 0 To UBound(listAreaDist)
                        rs.Fields(i + 2).value = CStr(listAreaDist(i))
                    Next
                Else
                    For i = 0 To rs.Fields.Count - 3
                        rs.Fields(i + 2).value = CStr(0)
                    Next
                End If
            End If
        End If
        
    Else
        For i = 0 To rs.Fields.Count - 3
            rs.Fields(i + 2).value = CStr(0)
        Next
    End If

    Set DataGridLuDist.DataSource = rs
    DataGridLuDist.Columns(0).Width = 1800
    DataGridLuDist.Columns(1).Width = 1000

    DataGridLuDist.Columns(0).Caption = "Landuse Group/Info Type"
    DataGridLuDist.Columns(1).Caption = "Area (ac.)"
    For i = 0 To rs.Fields.Count - 3
        DataGridLuDist.Columns(i + 2).Caption = rs.Fields(i + 2).name & " (%)"
    Next
    DataGridLuDist.Columns(aggBmpTypeDict.Count + 2).Caption = "Outlet (%)"
    DataGridLuDist.Columns(0).Locked = True
    DataGridLuDist.Columns(1).Locked = True
    
    If errMessage <> "" Then
        MsgBox "Area(s) currently stored for the following land uses do not match the actual watershed based land use distribution" & vbNewLine & errMessage, vbInformation
    End If
    
    curCbxSwsLI = cbxSws.ListIndex
    pModifyFlag = True
    
    Exit Sub
ShowError:
    MsgBox "Error in UpdateDataGrid in FrmAggBmpLuDist : " & Err.description '& "Line number :" & lineNumber
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub

