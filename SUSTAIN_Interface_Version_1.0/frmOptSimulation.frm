VERSION 5.00
Begin VB.Form frmOptSimulation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Land Simulation Option"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "frmOptSimulation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   2280
      TabIndex        =   3
      Top             =   750
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Land Simulation"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.OptionButton optInternal 
         Caption         =   "Internal Simulation"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optExternal 
         Caption         =   "External Simulation"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmOptSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOk_Click()
    Call ReadLayerTagDictionaryToSRCFile
    
    If optExternal.value = True Then
        gExternalSimulation = True
        gInternalSimulation = False
        gLayerNameDictionary.Item("SimulationOption") = "External"
    ElseIf optInternal.value = True Then
        'Check for Aggregate BMP
        Dim pBMPFLayer As IFeatureLayer
        Set pBMPFLayer = GetInputFeatureLayer("BMPs")
        
        'If Not (pTable Is Nothing Or pBMPFLayer Is Nothing) Then
        If Not pBMPFLayer Is Nothing Then
            Dim pBMPFClass As IFeatureClass
            Set pBMPFClass = pBMPFLayer.FeatureClass
            
            If Not pBMPFClass Is Nothing Then
                Dim pQueryFilter As IQueryFilter
                Set pQueryFilter = New QueryFilter
                pQueryFilter.WhereClause = "TYPE = 'Aggregate'"
                
                If pBMPFClass.FeatureCount(pQueryFilter) > 0 Then
                    MsgBox "Aggregate BMPs are not allowed for internal simlation. Please remove them and try again", vbExclamation
                    optExternal.value = True
                    GoTo CleanUp
                End If
           End If
        End If

        gInternalSimulation = True
        gExternalSimulation = False
        gLayerNameDictionary.Item("SimulationOption") = "Internal"
    End If
    
    '** All values are entered, save it a dictionary, and call a routine
'    Dim pOptionProperty As Scripting.Dictionary
'    Set pOptionProperty = CreateObject("Scripting.Dictionary")
'
'    '** Add all values to the dictionary
'    pOptionProperty.add "Internal", optInternal.value
'
'    '** Call the module to create table and add rows for these values
'    ModuleSWMMFunctions.SaveSWMMPropertiesTable "SimulationOption", "1", pOptionProperty
    
    ' Now Load the data into the Geodatabase......
'    Dim pTable As iTable
'    Set pTable = GetTable(gGDBpath, "SimulationOption")
'    If pTable Is Nothing Then
'        Import_Shape_To_GDB gMapTempFolder, "SimulationOption.dbf", gGDBpath, esriDTTable
'    End If
    

    Call WriteLayerTagDictionaryToSRCFile
    Unload Me
CleanUp:
    Set pBMPFLayer = Nothing
    Set pBMPFClass = Nothing
    Set pQueryFilter = Nothing
End Sub


Private Sub Form_Load()
    Set Me.Icon = LoadResPicture("SUSTAIN", vbResIcon)
    If gInternalSimulation Then optInternal.value = 1
    If gExternalSimulation Then optExternal.value = 1
        
End Sub

