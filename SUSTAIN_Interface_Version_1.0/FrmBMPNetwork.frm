VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBMPNetwork 
   Caption         =   "Update BMP Network"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   Icon            =   "FrmBMPNetwork.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGridBMP 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Label Label1 
      Caption         =   "Update BMP routing for splitter(s)"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "FrmBMPNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
    ModifyRouteLayer "Schematic Route", "Schematic BMPs"
End Sub

Private Sub Form_Load()
    'Call the subroutine to define the bmp network routing
    LoadBMPNetworkRouting
End Sub

'Subroutine to display decay factors
Private Sub LoadBMPNetworkRouting()

    Dim oConn As New ADODB.Connection
    oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
           "SourceType=DBF;" & _
           "SourceDB=" & gMapTempFolder & ";" & _
           "Exclusive=No"
    'Note: Specify the filename in the SQL statement. For example:
    Dim oRs As New ADODB.Recordset
    oRs.CursorLocation = adUseClient
    oRs.Open "Select * From BMPNetwork.dbf", oConn, adOpenDynamic, adLockOptimistic, adCmdText
    Set DataGridBMP.DataSource = oRs
    DataGridBMP.Columns(0).Locked = True
    DataGridBMP.Columns(1).Locked = True
    DataGridBMP.Columns(3).Locked = True
    
End Sub

