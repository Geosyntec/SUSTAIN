VERSION 5.00
Begin VB.Form FrmAggBmpSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Treatment Type"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "FrmAggBmpSelection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   480
      Left            =   1320
      TabIndex        =   5
      Top             =   510
      Width           =   700
   End
   Begin VB.TextBox bmpId 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "bmpId"
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbBmpType 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdRedefine 
      Caption         =   "Redefine"
      Height          =   480
      Left            =   2030
      TabIndex        =   1
      Top             =   510
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   480
      Left            =   3000
      TabIndex        =   0
      Top             =   510
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Select BMP"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "FrmAggBmpSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public pBmpIdDictionary As Scripting.Dictionary
Public pBmpCatDictionary As Scripting.Dictionary
Dim pAggBmpName As String

Public Sub Initialize_Form(curBmpID As Integer)
On Error GoTo ErrorHandler
    
    'Set pBmpIdDictionary = CreateObject("Scripting.Dictionary")
    Set pBmpCatDictionary = CreateObject("Scripting.Dictionary")

    Dim pAggBMPDetailsTable As iTable
    Set pAggBMPDetailsTable = GetInputDataTable("AgBMPDetail")
       
    If pAggBMPDetailsTable Is Nothing Then MsgBox "Missing AgBMPDetail Table": Exit Sub
    
    Dim idField As Long
    idField = pAggBMPDetailsTable.FindField("ID")
    If idField < 0 Then MsgBox "Missing ID Field in AgBMPDetail Table": Exit Sub
    
    Dim propValueField As Long
    propValueField = pAggBMPDetailsTable.FindField("PropValue")
    If propValueField < 0 Then MsgBox "Missing PropValue Field in AgBMPDetail Table": Exit Sub
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName = 'BMPID' AND PropValue = '" & curBmpID & "'"
    
    'Clear all the items from the cmbBmpTYpe
    cmbBmpType.Clear
    
    Dim pRow As iRow
    Dim pCursor As ICursor
    
    Dim tableIDs() As Integer
    Dim tableCnt As Integer
    
    Dim pQueryFilter2 As IQueryFilter
    Set pQueryFilter2 = New QueryFilter
    Dim pRow2 As iRow
    Dim pCursor2 As ICursor
    
    Dim pIDindex As Long
    pIDindex = pAggBMPDetailsTable.FindField("ID")
    Dim pPropValueIndex As Long
    pPropValueIndex = pAggBMPDetailsTable.FindField("PropValue")
    
    If pAggBMPDetailsTable.RowCount(pQueryFilter) > 0 Then
        ReDim tableIDs(pAggBMPDetailsTable.RowCount(pQueryFilter))
        tableCnt = 0
        Set pCursor = pAggBMPDetailsTable.Search(pQueryFilter, False)
        If Not pCursor Is Nothing Then
            Set pRow = pCursor.NextRow
            Do Until pRow Is Nothing
                tableIDs(tableCnt) = pRow.value(idField)
                tableCnt = tableCnt + 1
                Set pRow = pCursor.NextRow
            Loop
        End If
        If tableCnt > 0 Then
            For tableCnt = 0 To UBound(tableIDs)
                pQueryFilter.WhereClause = "ID = " & tableIDs(tableCnt) & " AND PropName = 'BMPName'"
                If pAggBMPDetailsTable.RowCount(pQueryFilter) > 0 Then
                    Set pCursor = pAggBMPDetailsTable.Search(pQueryFilter, False)
                    If Not pCursor Is Nothing Then
                        Set pRow = pCursor.NextRow
                        'Do Until pRow Is Nothing
                            cmbBmpType.AddItem Trim(pRow.value(propValueField))
                            cmbBmpType.ItemData(cmbBmpType.NewIndex) = tableIDs(tableCnt)
'                            Set pRow = pCursor.NextRow
'                        Loop
                    End If
                End If
                    
                'pQueryFilter2.WhereClause = "PropName='Category' And ID = " & tableIDs(tableCnt)
                pQueryFilter2.WhereClause = "PropName='BMPType' And ID = " & tableIDs(tableCnt)
                Set pCursor2 = pAggBMPDetailsTable.Search(pQueryFilter2, False)
                Set pRow2 = pCursor2.NextRow
                If Not pRow2 Is Nothing Then pBmpCatDictionary.add pRow2.value(pIDindex), pRow2.value(pPropValueIndex)
            Next
            cmbBmpType.ListIndex = 0
            pQueryFilter2.WhereClause = "ID = " & tableIDs(0) & " AND PropName = 'Type'"
            Set pCursor2 = pAggBMPDetailsTable.Search(pQueryFilter2, False)
            Set pRow2 = pCursor2.NextRow
            If Not pRow2 Is Nothing Then pAggBmpName = pRow2.value(pPropValueIndex)
                
        End If
    End If
    bmpId.Text = curBmpID
    'If cmbBmpType.ListCount = 4 Then cmdAdd.Enabled = False
    Exit Sub
ErrorHandler:
    MsgBox "Error initializing AggBMPSelection Form:" & Err.description

End Sub

Private Sub cmdAdd_Click()
    Me.Hide
''    'Open the bmp types
''    gNewBMPName = pAggBmpName '"AggregateBMP" & bmpId.Text
''    gNewBMPType = "Aggregate"
    gNewBMPId = CInt(bmpId.Text)
    Call InitializeAggBMPTypes
    
    gBMPTypeTag = "BMPOnMap"

    gBMPEditMode = False
    Load frmAggBMPDef
    'frmAggBMPDef.Tag = "BMPOnMap"
    frmAggBMPDef.Form_Initialize
    frmAggBMPDef.TabBMPType.Tab = 3
    frmAggBMPDef.BMPType.Text = pAggBmpName 'gNewBMPName
    frmAggBMPDef.BMPType.Enabled = False
    
    Set gBMPPlacedDict = New Scripting.Dictionary
    Dim bmpCategory As String
    Dim bBmpId As Integer
    Dim iCnt As Integer
    For iCnt = 0 To cmbBmpType.ListCount - 1
        bBmpId = cmbBmpType.ItemData(iCnt)
        bmpCategory = pBmpCatDictionary.Item(bBmpId)
        gBMPPlacedDict.add bmpCategory, bmpCategory
    Next
    
    frmAggBMPDef.cmbBMPCategory.Enabled = True
    frmAggBMPDef.cmbBmpType.Enabled = True
    frmAggBMPDef.BMPNameA.Enabled = True
    
    
    frmAggBMPDef.Show vbModal
    If gBMPDetailDict Is Nothing Then Exit Sub
    
    Dim bmpRecId As Integer
    bmpRecId = cmbBmpType.ListCount + 1

    Dim curBmpID As Integer
    curBmpID = CInt(bmpId.Text)
    
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("AgBMPDetail")
    
    Dim pIDindex As Long
    pIDindex = pBMPDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pBMPKeys
    Dim i As Integer
    Dim pRow As iRow
    
    If Not (gBMPDetailDict Is Nothing) Then
        gBMPDetailDict.Item("BMPID") = curBmpID
        gBMPDetailDict.Item("BMPName") = pAggBmpName 'gNewBMPName
        pBMPKeys = gBMPDetailDict.keys
        For i = 0 To (gBMPDetailDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            pPropertyValue = gBMPDetailDict.Item(pPropertyName)
            'Create if the row is not already in the table
            Set pRow = pBMPDetailTable.CreateRow
            pRow.value(pIDindex) = bmpRecId
            pRow.value(pPropNameIndex) = pPropertyName
            pRow.value(pPropValueIndex) = pPropertyValue
            pRow.Store
        Next
    End If
    
    frmAggBMPDef.Tag = ""
'    If Not gBMPOptionsDict Is Nothing Then
'        pBMPKeys = gBMPOptionsDict.keys
'        For i = 0 To (gBMPOptionsDict.Count - 1)
'            pPropertyName = pBMPKeys(i)
'            If Not gBMPDetailDict.Exists(pPropertyName) Then
'                pPropertyValue = gBMPOptionsDict.Item(pPropertyName)
'                'Create if the row is not already in the table
'                Set pRow = pBMPDetailTable.CreateRow
'                pRow.value(pIDindex) = bmpRecId
'                pRow.value(pPropNameIndex) = pPropertyName
'                pRow.value(pPropValueIndex) = pPropertyValue
'                pRow.Store
'            End If
'        Next
'    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRedefine_Click()
On Error GoTo ErrorHandler
    
    Me.Hide
    'Open the bmp types
'    gNewBMPName = pAggBmpName '"AggregateBMP" & bmpId.Text
'    gNewBMPType = "Aggregate"
    
    Call InitializeAggBMPTypes
    gNewBMPId = cmbBmpType.ItemData(cmbBmpType.ListIndex)
    
    gBMPTypeTag = "BMPOnMap"
    gBMPEditMode = True
    
    Load frmAggBMPDef
    'frmAggBMPDef.Tag = "BMPOnMap"
    frmAggBMPDef.TabBMPType.Tab = 3
    frmAggBMPDef.Form_Initialize
    frmAggBMPDef.BMPType.Text = pAggBmpName 'gNewBMPName
    frmAggBMPDef.BMPType.Enabled = False
        
    Set gBMPPlacedDict = New Scripting.Dictionary
    Dim bmpCategory As String
    Dim bBmpId As Integer
    Dim iCnt As Integer
    For iCnt = 0 To cmbBmpType.ListCount - 1
        bBmpId = cmbBmpType.ItemData(iCnt)
        bmpCategory = pBmpCatDictionary.Item(bBmpId)
        gBMPPlacedDict.add bmpCategory, bmpCategory
    Next
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = GetBMPDetailDict(gNewBMPId, "AgBMPDetail")

    Dim curBMPType As String
    Dim curCategory As String

    curBMPType = pBmpDetailDict.Item("BMPType")
    curCategory = pBmpDetailDict.Item("Category")

    'set bmp category and then disable it (not editable)
    Dim ikey As Integer
    For ikey = 0 To frmAggBMPDef.cmbBMPCategory.ListCount
        If frmAggBMPDef.cmbBMPCategory.List(ikey) = curCategory Then
            frmAggBMPDef.cmbBMPCategory.ListIndex = ikey
            Exit For
        End If
    Next
    frmAggBMPDef.cmbBMPCategory.Enabled = False

    'set bmp type and then disable it (not editable)
    For ikey = 0 To frmAggBMPDef.cmbBmpType.ListCount
        If frmAggBMPDef.cmbBmpType.List(ikey) = Get_BMP_Name(curBMPType) Then
            frmAggBMPDef.cmbBmpType.ListIndex = ikey
            Exit For
        End If
    Next
    frmAggBMPDef.cmbBmpType.Enabled = False
    
    frmAggBMPDef.BMPNameA.Text = pBmpDetailDict.Item("BMPName")
    frmAggBMPDef.BMPNameA.Enabled = False
    
    frmAggBMPDef.Show vbModal
    If gBMPDetailDict Is Nothing Then Exit Sub
    
    Dim curBmpID As Integer
    curBmpID = CInt(bmpId.Text)
    
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("AgBMPDetail")
    
    Dim pQueryFilter As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pIDindex As Long
    pIDindex = pBMPDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    

'
'    'Set the decay rates and underdrain percentage removals
'    ModuleBMPData.LoadPollutantData pBmpDetailDict
'
'    Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
'    ModuleBMPData.CallInitRoutines gNewBMPType, pBmpDetailDict
    
    
    'Remove records for bmp which are not in the gbmpdetaildictionary
    If Not (gBMPDetailDict Is Nothing) Then
        If Not gBMPDetailDict.Exists("BMPID") Then
            gBMPDetailDict.Item("BMPID") = curBmpID
        End If
    
        Set pQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = "ID = " & gNewBMPId
        Set pCursor = pBMPDetailTable.Update(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        Do While Not pRow Is Nothing
            If (Not gBMPDetailDict.Exists(pRow.value(pPropNameIndex))) Then
                pCursor.DeleteRow
            End If
            Set pRow = pCursor.NextRow
        Loop
        Set pRow = Nothing
        Set pCursor = Nothing
    End If
    
    Dim pBMPKeys
    Dim i As Integer
    
    If Not (gBMPDetailDict Is Nothing) Then
'        gBMPDetailDict.Item("Type") = curType
'        gBMPDetailDict.Item("Category") = curCategory
        
        pBMPKeys = gBMPDetailDict.keys
        For i = 0 To (gBMPDetailDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            pPropertyValue = gBMPDetailDict.Item(pPropertyName)
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = "ID = " & gNewBMPId & " AND PropName = '" & pPropertyName & "'"
            Set pCursor = pBMPDetailTable.Update(pQueryFilter, False)
            Set pRow = pCursor.NextRow
            If Not pRow Is Nothing Then
                pRow.value(pPropValueIndex) = pPropertyValue
                pRow.Store
            Else
                'Create if the row is not already in the table
                Set pRow = pBMPDetailTable.CreateRow
                pRow.value(pIDindex) = gNewBMPId
                pRow.value(pPropNameIndex) = pPropertyName
                pRow.value(pPropValueIndex) = pPropertyValue
                pRow.Store
            End If
        Next
        
'        'Insert the BMP General options
'        If Not gBMPOptionsDict Is Nothing Then
'            pBMPKeys = gBMPOptionsDict.keys
'            For i = 0 To (gBMPOptionsDict.Count - 1)
'                pPropertyName = pBMPKeys(i)
'                'Modified on Feb 23, 2009 - Sabu Paul - No need to insert the entry if it is already in the BMPDetailDict.
'                'If pPropertyName <> "BMPType" Then
'                If Not gBMPDetailDict.Exists(pPropertyName) Then
'                    pPropertyValue = gBMPOptionsDict.Item(pPropertyName)
'                    Set pQueryFilter = New QueryFilter
'                    pQueryFilter.WhereClause = "ID = " & gNewBMPId & " AND PropName = '" & pPropertyName & "'"
'                    Set pCursor = pBMPDetailTable.Update(pQueryFilter, False)
'                    Set pRow = pCursor.NextRow
'                    If Not pRow Is Nothing Then
'                        pRow.value(pPropValueIndex) = pPropertyValue
'                        pRow.Store
'                    Else
'                        'Create if the row is not already in the table
'                        Set pRow = pBMPDetailTable.CreateRow
'                        pRow.value(pIDindex) = gNewBMPId
'                        pRow.value(pPropNameIndex) = pPropertyName
'                        pRow.value(pPropValueIndex) = pPropertyValue
'                        pRow.Store
'                    End If
'                End If
'            Next
'        End If
    End If
    
    frmAggBMPDef.Tag = ""
    GoTo CleanUp
ErrorHandler:
    MsgBox "Error in redefining bmp data: " & Err.description
CleanUp:
    Set pBmpDetailDict = Nothing
    Set pBMPDetailTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
