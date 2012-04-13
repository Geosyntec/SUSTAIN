VERSION 5.00
Begin VB.Form FrmVFSFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select NHD Fields"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "FrmVFSFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cmbSUBBASINR 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cmbSUBBASIN 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Select the Downstream ID (COM_ID2 or SUBBASINR):"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Stream ID (COM_ID or SUBBASIN):"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FrmVFSFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
''
''Private Sub cmdCancel_Click()
''    Unload Me
''End Sub
''
''Private Sub cmdOK_Click()
''
''    '** Get sub-basin, and receiving sub-basin names.
''    gSUBBASINFieldName = cmbSUBBASIN.Text
''    gSUBBASINRFieldName = cmbSUBBASINR.Text
''
''    '** Close the form.
''    Unload Me
''End Sub
''
''Private Sub Form_Load()
''    Dim pSTREAMFLayer As IFeatureLayer
''    Set pSTREAMFLayer = GetInputFeatureLayer("STREAM")
''    If (pSTREAMFLayer Is Nothing) Then
''        MsgBox "Streams feature layer not found."
''        Exit Sub
''    End If
''    Dim pSTREAMFClass As IFeatureClass
''    Set pSTREAMFClass = pSTREAMFLayer.FeatureClass
''
''    Dim pFields As IFields
''    Set pFields = pSTREAMFClass.Fields
''    Dim pField As IField
''    Dim i As Integer
''    Dim pSUBBASINindex As Integer
''    Dim pSUBBASINRindex As Integer
''    pSUBBASINindex = 0
''    pSUBBASINRindex = 0
''    For i = 0 To (pFields.FieldCount - 1)
''      Set pField = pFields.Field(i)
''      cmbSUBBASIN.AddItem pField.Name
''      cmbSUBBASINR.AddItem pField.Name
''      '**check if SUBBASIN/COM_ID field is present
''      If (pField.Name = "SUBBASIN" Or pField.Name = "COM_ID") Then
''        pSUBBASINindex = i
''      End If
''      '**check if SUBBASIN/COM_ID field is present
''      If (pField.Name = "SUBBASINR" Or pField.Name = "COM_ID2") Then
''        pSUBBASINRindex = i
''      End If
''    Next
''    cmbSUBBASIN.ListIndex = pSUBBASINindex
''    cmbSUBBASINR.ListIndex = pSUBBASINRindex
''
''    '** Cleanup
''    Set pField = Nothing
''    Set pFields = Nothing
''    Set pSTREAMFClass = Nothing
''    Set pSTREAMFLayer = Nothing
''
''End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
End Sub
