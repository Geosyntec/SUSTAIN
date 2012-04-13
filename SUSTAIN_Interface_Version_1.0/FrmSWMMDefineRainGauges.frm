VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmSWMMDefineRainGauges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Rain Gage Properties"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "FrmSWMMDefineRainGauges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRainGuage 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   120
      TabIndex        =   24
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   3480
      Width           =   760
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtRainGaugeID 
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Text            =   "HIDDEN"
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   2160
      TabIndex        =   9
      Top             =   2400
      Width           =   6495
      Begin VB.ComboBox cmbRainUnits 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtStationNo 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtFile 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label11 
         Caption         =   "Rain Units"
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Station No"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Rainfall File"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   6495
      Begin VB.ComboBox cmbRainInterval 
         Height          =   315
         ItemData        =   "FrmSWMMDefineRainGauges.frx":08CA
         Left            =   1560
         List            =   "FrmSWMMDefineRainGauges.frx":08CC
         TabIndex        =   20
         Text            =   "cmbRainInterval"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtSnowCatchFactor 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Text            =   "1.0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbRainFormat 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Snow Catch Factor"
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Rain Interval"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Rain Format"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "FrmSWMMDefineRainGauges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    
    Call LoadRainGuages
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    cmdOK.Enabled = True
    txtRainGaugeID.Text = ""
    
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog.Filter = "Rain Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
    CommonDialog.ShowOpen
    txtFile.Text = CommonDialog.FileName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdEdit_Click()
    
    If lstRainGuage.ListCount = 0 Then Exit Sub
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    cmdOK.Enabled = True
    txtRainGaugeID.Text = lstRainGuage.ListIndex + 1
    
    '** pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstRainGuage.ListIndex + 1
    
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = LoadRainGuageDetails(pPollutantID)
    
    '** return if nothing found
    If (pValueDictionary.Count = 0) Then
        Exit Sub
    End If
    
    txtName.Text = pValueDictionary.Item("Name")
    cmbRainFormat.Text = pValueDictionary.Item("Rain Type")
    cmbRainInterval.Text = pValueDictionary.Item("Recd. Freq")
    txtSnowCatchFactor.Text = pValueDictionary.Item("Snow Catch")
    txtFile.Text = Replace(pValueDictionary.Item("Source Name"), """", "")
    txtStationNo.Text = pValueDictionary.Item("Station ID")
    cmbRainUnits.Text = pValueDictionary.Item("Rain Units")

    
End Sub

Private Sub cmdOk_Click()

    '** Input data-validation
    
    ' check if Duplicate exists.....
    If GetListBoxIndex(lstRainGuage, txtName.Text) > -1 And txtRainGaugeID.Text = "" Then Exit Sub
    
    '* input Name
    If (txtName.Text = "") Then
        MsgBox "Please specify rain guage name to continue.", vbExclamation
        If txtName.Enabled Then txtName.SetFocus
        Exit Sub
    End If
    
    '* input data file
    If (txtFile.Text = "") Then
        MsgBox "Please specify rain data file to continue.", vbExclamation
        txtFile.SetFocus
        Exit Sub
    End If
    '* snow catch factor
    If (Not IsNumeric(txtSnowCatchFactor.Text)) Then
        MsgBox "Please specify snow catch factor as a valid number.", vbExclamation
        txtSnowCatchFactor.SetFocus
        Exit Sub
    End If
    'Rain interval
    Dim bRainIntErr As Boolean
    bRainIntErr = False
    If Replace(cmbRainInterval.Text, ":", "") = cmbRainInterval.Text Then
        bRainIntErr = True
    Else
        Dim intWords
        intWords = Split(cmbRainInterval.Text, ":")
        If UBound(intWords) = 1 Then
            If Not (IsNumeric(intWords(0)) And IsNumeric(intWords(1))) Then
                bRainIntErr = True
            End If
        Else
            bRainIntErr = True
        End If
    End If
    If bRainIntErr Then
        MsgBox "Please specify rain interval as HH:MM.", vbExclamation
        cmbRainInterval.SetFocus
        Exit Sub
    End If
    
    '* station no.
    If (Trim(txtStationNo.Text) = "") Then
        MsgBox "Please specify station number to continue.", vbExclamation
        txtStationNo.SetFocus
        Exit Sub
    End If
    
    '** All values are entered, save it a dictionary, and call a routine
    Dim pOptionProperty As Scripting.Dictionary
    Set pOptionProperty = CreateObject("Scripting.Dictionary")
    
    '** Add all values to the dictionary
    pOptionProperty.add "Name", txtName.Text
    pOptionProperty.add "Rain Type", cmbRainFormat.Text
    pOptionProperty.add "Recd. Freq", cmbRainInterval.Text
    pOptionProperty.add "Snow Catch", txtSnowCatchFactor.Text
    pOptionProperty.add "Data Source", "FILE"
    pOptionProperty.add "Source Name", """" & txtFile.Text & """"
    pOptionProperty.add "Station ID", txtStationNo.Text
    pOptionProperty.add "Rain Units", cmbRainUnits.Text
    
    '** get the rain gauge id
    Dim pRGID As String
    If (Trim(txtRainGaugeID.Text) = "") Then
        pRGID = lstRainGuage.ListCount + 1
    Else
        pRGID = txtRainGaugeID.Text
    End If
    
    '** Call the module to create table and add rows for these values
    ModuleSWMMFunctions.SaveSWMMPropertiesTable "LANDRainGages", pRGID, pOptionProperty
       
'    'Get rain gauge feature layer, if not found, create feature class for it
'    Dim pRainGaugeFeatureLayer As IFeatureLayer
'    Set pRainGaugeFeatureLayer = GetInputFeatureLayer("Rain Gauges")
'
'    'Get feature class, if not found, create it, create a feature layer from it
'    Dim pRainGaugeFeatureClass As IFeatureClass
'    Set pRainGaugeFeatureClass = pRainGaugeFeatureLayer.FeatureClass
'
'    'Add a new bmp feature
'    Dim pFeature As IFeature
'    Set pFeature = pRainGaugeFeatureClass.CreateFeature
'    Set pFeature.Shape = gRainGaugePoint
'    pFeature.value(pRainGaugeFeatureClass.FindField("ID")) = pRGID
'    pFeature.value(pRainGaugeFeatureClass.FindField("TYPE")) = "Rain Gauge"
'    pFeature.value(pRainGaugeFeatureClass.FindField("TYPE2")) = "Rain Gauge"
'    pFeature.value(pRainGaugeFeatureClass.FindField("LABEL")) = "RG" & pRGID
'    pFeature.Store
'
'    '* render the rain gauge feature layer
'    RenderRainGaugeLayer pRainGaugeFeatureLayer
    
    
    cmdOK.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    Call LoadRainGuages
    
    '** Clean up
    Set gRainGaugePoint = Nothing
    Set pOptionProperty = Nothing
    Set pFeature = Nothing
    'Set pRainGaugeFeatureClass = Nothing
    'Set pRainGaugeFeatureLayer = Nothing
End Sub

Private Sub cmdRemove_Click()
    
    On Error GoTo ShowError
    If lstRainGuage.ListCount = 0 Then Exit Sub
    
    '** Confirm the deletion
    Dim boolDelete
    boolDelete = MsgBox("Are you sure you want to delete this Rain guage information ?", vbYesNo)
    If (boolDelete = vbNo) Then
        Exit Sub
    End If
    
    '** get pollutant id
    Dim pPollutantID As Integer
    pPollutantID = lstRainGuage.ListIndex + 1
    
    '** get the table to delete records
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDRainGages")
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pPollutantID
    
    '** delete records
    pSWMMPollutantTable.DeleteSearchedRows pQueryFilter
    
    '*** Increment the id's by 1 number for all records after deleted id
    Dim pFromID As Integer
    pFromID = pPollutantID + 1
    Dim bContinue As Boolean
    bContinue = True
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iID As Long
    iID = pSWMMPollutantTable.FindField("ID")
    
    Do While bContinue
        pQueryFilter.WhereClause = "ID = " & pFromID
        Set pCursor = Nothing
        Set pRow = Nothing
        Set pCursor = pSWMMPollutantTable.Search(pQueryFilter, False)
        Set pRow = pCursor.NextRow
        If (pRow Is Nothing) Then
            bContinue = False
        End If
        Do While Not pRow Is Nothing
            pRow.value(iID) = pFromID - 1
            pRow.Store
            Set pRow = pCursor.NextRow
        Loop
        pFromID = pFromID + 1
    Loop
    
    '** clean up
    Set pQueryFilter = Nothing
    Set pSWMMPollutantTable = Nothing
    
    '** load pollutant names
    Call LoadRainGuages
    
    Exit Sub
ShowError:
    MsgBox "cmdRemove :" & Err.description
    
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    '*** Load the RainGauge Format values
    cmbRainFormat.AddItem "INTENSITY"
    cmbRainFormat.AddItem "VOLUME"
    cmbRainFormat.AddItem "CUMULATIVE"
    cmbRainFormat.ListIndex = 0
    
    '*** Load the RainInterval values
    cmbRainInterval.AddItem "0:01"
    cmbRainInterval.AddItem "0:05"
    cmbRainInterval.AddItem "0:10"
    cmbRainInterval.AddItem "0:15"
    cmbRainInterval.AddItem "0:20"
    cmbRainInterval.AddItem "0:30"
    cmbRainInterval.AddItem "1:00"
    cmbRainInterval.AddItem "6:00"
    cmbRainInterval.AddItem "12:00"
    cmbRainInterval.AddItem "24:00"
    cmbRainInterval.Text = "1:00"
    
    '*** Load the RainUnit values
    cmbRainUnits.AddItem "IN"
    cmbRainUnits.AddItem "MM"
    cmbRainUnits.ListIndex = 0
    
    
    Call LoadRainGuages
    
End Sub

Public Sub LoadRainGuages()
On Error GoTo ShowError

    '** Refresh pollutant names
    Dim pPollutantCollection As Collection
    Set pPollutantCollection = ModuleSWMMFunctions.LoadRainGaugeNames
    If (pPollutantCollection Is Nothing) Then
        Exit Sub
    End If
    
    lstRainGuage.Clear
    Dim pCount As Integer
    pCount = pPollutantCollection.Count
    Dim iCount As Integer
    For iCount = 1 To pCount
        lstRainGuage.AddItem pPollutantCollection.Item(iCount)
    Next
    
    lstRainGuage.Selected(0) = True
    txtName.Text = ""
    txtFile.Text = ""
    Set pPollutantCollection = Nothing
    
    Exit Sub
    
ShowError:
    MsgBox "LoadAquifers :" & Err.description
End Sub

Private Function LoadRainGuageDetails(pID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    '* Load the list box with pollutant names
    Dim pSWMMPollutantTable As iTable
    Set pSWMMPollutantTable = GetInputDataTable("LANDRainGages")
    If (pSWMMPollutantTable Is Nothing) Then
        Exit Function
    End If
    Dim iPropName As Long
    iPropName = pSWMMPollutantTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pSWMMPollutantTable.FindField("PropValue")
    
    'Define query filter
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pID
    
    'Get the cursor to iterate over the table
    Dim pCursor As ICursor
    Set pCursor = pSWMMPollutantTable.Search(pQueryFilter, True)
    
    'Define a dictionary to store the values
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = CreateObject("Scripting.Dictionary")
    
    'Define a row variable to loop over the table
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    Do While Not pRow Is Nothing
        pValueDictionary.add Trim(pRow.value(iPropName)), Trim(pRow.value(iPropValue))
        Set pRow = pCursor.NextRow
    Loop
    
    '** Return the dictionary
    Set LoadRainGuageDetails = pValueDictionary
    
    '** Cleanup
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pSWMMPollutantTable = Nothing
    Set pQueryFilter = Nothing
    Exit Function
ShowError:
    MsgBox "LoadRainGuageDetails: " & Err.description
    
End Function


