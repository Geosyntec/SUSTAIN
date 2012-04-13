Attribute VB_Name = "ModuleOptimization"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleOptimization
'   Purpose:     This module creates necessary tables and input for optimization parameters
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created: 08/23/2004 - mira chokshi
'
'******************************************************************************

Option Explicit
Option Base 0


'** Subroutine to delete existing options before overwriting
Public Sub DeletePreviouslyDefinedOptions(pBMPID As Integer)

    'Get landuse reclassification table: LUReclass, Create new if not found
      Dim pOptimizationTable As iTable
      Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
      If Not (pOptimizationTable Is Nothing) Then
          'Query in the table for existing records
          Dim pQueryFilter As IQueryFilter
          Set pQueryFilter = New QueryFilter
          pQueryFilter.WhereClause = "ID = " & pBMPID
          pOptimizationTable.DeleteSearchedRows pQueryFilter
      End If
      Set pQueryFilter = Nothing
      Set pOptimizationTable = Nothing
End Sub


Public Sub DefineOptimizationMethod(optimizeOption As Integer, limitCost As Double, StopDelta As Double, MaxRunTime As Double, NumBest As Integer)
On Error GoTo ShowError

    Dim pOptimDetailTable As iTable
    Set pOptimDetailTable = GetInputDataTable("OptimizationDetail")
    
    If (pOptimDetailTable Is Nothing) Then
        Set pOptimDetailTable = CreatePropertiesTableDBF("OptimizationDetail")
        AddTableToMap pOptimDetailTable
    End If
    Dim iIDindex As Long
    iIDindex = pOptimDetailTable.FindField("ID")
    Dim iPropNameIndex As Long
    iPropNameIndex = pOptimDetailTable.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pOptimDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'Option'"
    
    'Define variables to iterate the tables
    Dim pCursor As ICursor
    Dim pRow As iRow
   
    Set pCursor = pOptimDetailTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pExistingOption As Integer
    If Not (pRow Is Nothing) Then
        pExistingOption = CInt(pRow.value(iPropValueIndex))
    End If
    Set pRow = Nothing
    Set pCursor = Nothing
    

    If (pExistingOption = optimizeOption) Then
            pQueryFilter.WhereClause = "ID = 0"     'It will be overwritten in steps below
            pOptimDetailTable.DeleteSearchedRows pQueryFilter   'Delete all option rows only
    Else
       'Delete rows if already present, delete all rows from Optimization Detail, BMPDetail & BMPs feature layer
        pOptimDetailTable.DeleteSearchedRows Nothing
        Call DeleteAllAssessmentPointsDetails
    End If
    
    'Create new records for option and cost limit and input information
    Set pRow = pOptimDetailTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "Option"
    pRow.value(iPropValueIndex) = optimizeOption
    pRow.Store

    Set pRow = pOptimDetailTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "CostLimit"
    pRow.value(iPropValueIndex) = limitCost
    pRow.Store
        
    Set pRow = pOptimDetailTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "StopDelta"
    pRow.value(iPropValueIndex) = StopDelta
    pRow.Store
    
    Set pRow = pOptimDetailTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "MaxRunTime"
    pRow.value(iPropValueIndex) = MaxRunTime
    pRow.Store
    
    Set pRow = pOptimDetailTable.CreateRow
    pRow.value(iIDindex) = 0
    pRow.value(iPropNameIndex) = "NumBest"
    pRow.value(iPropValueIndex) = NumBest
    pRow.Store
    
    GoTo CleanUp
    
ShowError:
    MsgBox "DefineOptimizationMethod: " & Err.description
CleanUp:
    Set pOptimDetailTable = Nothing
    Set pRow = Nothing
End Sub


Public Sub DeleteAllAssessmentPointsDetails()
On Error GoTo ShowError
     
     '*** Update the BMP Detail table
     Dim pBMPDetail As iTable
     Set pBMPDetail = GetInputDataTable("BMPDetail")
     Dim iPropValue As Long
     iPropValue = pBMPDetail.FindField("PropValue")
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter
     pQueryFilter.WhereClause = "PropName = 'isAssessmentPoint'"
     Dim pCursor As ICursor
     Dim pRow As iRow
     Set pCursor = pBMPDetail.Search(pQueryFilter, True)
     Set pRow = pCursor.NextRow
     Do While Not (pRow Is Nothing)
        pRow.value(iPropValue) = "False"
        pRow.Store
        Set pRow = pCursor.NextRow
     Loop


    '*** Update the BMPs feature layer
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim iType As Long
    iType = pFeatureclass.FindField("TYPE")
    Dim iType2 As Long
    iType2 = pFeatureclass.FindField("TYPE2")
    Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
    Set pFeature = pFeatureCursor.NextFeature
    Do While Not (pFeature Is Nothing)
        pFeature.value(iType2) = pFeature.value(iType)
        pFeature.Store
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    
    '*** Update the BMP rendering
    RenderSchematicBMPLayer pFeatureLayer
    
    GoTo CleanUp
ShowError:
    MsgBox "DeleteAllAssessmentPointsDetails: " & Err.description
CleanUp:
     Set pQueryFilter = Nothing
     Set pBMPDetail = Nothing
     Set pCursor = Nothing
     Set pRow = Nothing
     Set pFeatureLayer = Nothing
     Set pFeatureclass = Nothing
     Set pFeatureCursor = Nothing
     Set pFeature = Nothing
End Sub


Public Function CheckDecisionVariablesPresent() As Boolean
On Error GoTo ShowError
    CheckDecisionVariablesPresent = False

    Dim pBMPDetail As iTable
    Set pBMPDetail = GetInputDataTable("BMPDetail")
    If (pBMPDetail Is Nothing) Then
        MsgBox "BMPDetail table not found."
        Exit Function
    End If
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName LIKE '%Optimized'"
    
    Dim pCursor As ICursor
    Set pCursor = pBMPDetail.Search(pQueryFilter, True)
    
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim iPropValue As Long
    iPropValue = pBMPDetail.FindField("PropValue")
    
    Do While Not pRow Is Nothing
        If (pRow.value(iPropValue) = "True") Then
            CheckDecisionVariablesPresent = True
            GoTo CleanUp
        End If
        'continue iteration
        Set pRow = pCursor.NextRow
    Loop
    

    GoTo CleanUp
ShowError:
    MsgBox "CheckDecisionVariablesPresent :" & Err.description
CleanUp:
    Set pQueryFilter = Nothing
    Set pBMPDetail = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing

End Function

Public Sub UpdateOptimizationParamsForBMP(pBMPID As Integer, pOptimizationColl As Collection)
    
    'Update the parameters in BMPs feature layer and BMPDetail table
    Dim pBMPFLayer As IFeatureLayer
    Set pBMPFLayer = GetInputFeatureLayer("BMPs")
    Dim pBMPFClass As IFeatureClass
    Set pBMPFClass = pBMPFLayer.FeatureClass
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pBMPID
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pBMPFClass.Search(pQueryFilter, True)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Dim pBMPType2Val As String
    If Not (pFeature Is Nothing) Then
        'Set the Type2 parameter to designate assessment point
        pBMPType2Val = Trim(pFeature.value(pFeatureCursor.FindField("TYPE2")))
        If (Right(pBMPType2Val, 1) <> "X") Then
            pFeature.value(pFeatureCursor.FindField("TYPE2")) = pFeature.value(pFeatureCursor.FindField("TYPE2")) & "X"
            pFeature.Store
        End If
    End If
    
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDetail")
    pQueryFilter.WhereClause = "ID = " & pBMPID & " AND PropName = 'isAssessmentPoint'"
    Dim pCursor As ICursor
    Set pCursor = pBMPDetailTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    If Not (pRow Is Nothing) Then
        pRow.value(pCursor.FindField("PropValue")) = "True"
        pRow.Store
    End If
    
    
    Dim pArray As Variant
    Dim pACount As Integer
    Dim pCollCount As Integer
    Dim pStrParam As String
    
    Dim pOptimDetailTable As iTable
    Set pOptimDetailTable = GetInputDataTable("OptimizationDetail")
    If (pOptimDetailTable Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Sub
    End If
    Dim iIDindex As Long
    iIDindex = pOptimDetailTable.FindField("ID")
    Dim iPropNameIndex As Long
    iPropNameIndex = pOptimDetailTable.FindField("PropName")
    Dim iPropValueIndex As Long
    iPropValueIndex = pOptimDetailTable.FindField("PropValue")
    
    Dim pPropertyName As String
    Dim pPropertyValue As String

    'Iterate the array collection
    For pCollCount = 1 To pOptimizationColl.Count
        pArray = pOptimizationColl.Item(pCollCount)
'        pStrParam = ""
'        For pACount = 0 To 4
'            pStrParam = pStrParam & CStr(pArray(pACount)) & ","
'        Next
'        pStrParam = pStrParam & CStr(pArray(5))
        pStrParam = CStr(pArray(0))
        For pACount = 1 To UBound(pArray)
            pStrParam = pStrParam & "," & CStr(pArray(pACount))
        Next
        
        Set pRow = pOptimDetailTable.CreateRow
        pRow.value(iIDindex) = pBMPID
        pRow.value(iPropNameIndex) = "Parameters"
        pRow.value(iPropValueIndex) = pStrParam
        pRow.Store
    Next

    'Render the BMP feature layer and refresh
    RenderSchematicBMPLayer pBMPFLayer
    gMxDoc.ActiveView.Refresh
    
    
End Sub



'** Subroutine to read optimization table values
Public Function ReadOptimizationParametersForExistingCond(pBMPID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    'Get landuse reclassification table: LUReclass, Create new if not found
     Dim pOptimizationTable As iTable
     Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
     If (pOptimizationTable Is Nothing) Then
        Exit Function
     End If
     
    'Query in the table for existing records
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'Option'"
    Dim pCursor As ICursor
    Set pCursor = pOptimizationTable.Search(pQueryFilter, True)
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    'get the optimization option value
    Dim pOptimizeOption As Integer
    pOptimizeOption = -1
    If Not (pRow Is Nothing) Then
        pOptimizeOption = pRow.value(iPropValue)
    End If
     
    'Query in the table for existing records
    pQueryFilter.WhereClause = "ID = " & pBMPID
    Set pCursor = pOptimizationTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    
    Dim pRowString As String
    Dim pSplittedString
    Dim pFactorName As String
    Dim pFactorGroup As String
    Dim pFactorType As String
    Dim pCalcDays As Double
    Dim pCalcMode As String
    Dim pTargetValue As Double
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = CreateObject("Scripting.Dictionary")
    Do While Not pRow Is Nothing
        pRowString = pRow.value(iPropValue)
        pSplittedString = Split(pRowString, ",")
        pFactorGroup = pSplittedString(0)  'FACTOR-GROUP: 1,2,3 for pollutants, -1 for flow
        pFactorType = pSplittedString(1)   'FACTOR-TYPE: -1, -2, -3 for AAFV, PDF, FEF, 1,2,3 for AAL, AAC, MAC
        pCalcDays = CDbl(pSplittedString(2))
        pCalcMode = pSplittedString(3)
        pTargetValue = CDbl(pSplittedString(4))
        'pFactorName = pSplittedString(5)
        
        If (pFactorGroup = -1) Then     'FOR FLOW
            Select Case pFactorType
                Case "-1":
                    pValueDictionary.add "AAFV", vbChecked
                Case "-2":
                    pValueDictionary.add "PDF", vbChecked
                Case "-3":
                    pValueDictionary.add "FEF", vbChecked
                    pValueDictionary.add "FEF_CalcDays", pCalcDays
            End Select
        Else    '*** FOR POLLUTANTS
            Select Case pFactorType
                Case "1":   'AAL
                    pValueDictionary.add "AAL_Pollutant" & pFactorGroup, vbChecked
                Case "2":   'AAC
                    pValueDictionary.add "AAC_Pollutant" & pFactorGroup, vbChecked
                Case "3":   'MAC
                    pValueDictionary.add "MAC_Pollutant" & pFactorGroup, vbChecked
                    pValueDictionary.add "MAC_CalcDays" & pFactorGroup, pCalcDays
            End Select
        End If
        
        Set pRow = pCursor.NextRow
    Loop
    
    'Set the memory release
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pOptimizationTable = Nothing
    
    'Return the value dictionary
    Set ReadOptimizationParametersForExistingCond = pValueDictionary
    
    Exit Function
ShowError:
    MsgBox "ReadOptimizationParametersForExistingCond: " & Err.description
End Function



'** Subroutine to read optimization table values
Public Function ReadOptimizationParametersForMaximizeBenefit(pBMPID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    'Get landuse reclassification table: LUReclass, Create new if not found
     Dim pOptimizationTable As iTable
     Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
     If (pOptimizationTable Is Nothing) Then
        Exit Function
     End If
    
    
    'Query in the table for existing records
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pBMPID
    Dim pCursor As ICursor
    Set pCursor = pOptimizationTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    
    Dim pRowString As String
    Dim pSplittedString
    Dim pFactorName As String
    Dim pFactorGroup As String
    Dim pFactorType As String
    Dim pCalcDays As Double
    Dim pCalcMode As String
    Dim pTargetValue As Double
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = CreateObject("Scripting.Dictionary")
    Do While Not pRow Is Nothing
        pRowString = pRow.value(iPropValue)
        pSplittedString = Split(pRowString, ",")
        pFactorGroup = pSplittedString(0)  'FACTOR-GROUP: 1,2,3 for pollutants, -1 for flow
        pFactorType = pSplittedString(1)   'FACTOR-TYPE: -1, -2, -3 for AAFV, PDF, FEF, 1,2,3 for AAL, AAC, MAC
        pCalcDays = CDbl(pSplittedString(2))
        pCalcMode = pSplittedString(3)
        pTargetValue = CDbl(pSplittedString(4))
        
        If (pFactorGroup = -1) Then     'FOR FLOW
            Select Case pFactorType
                Case "-1":
                    pValueDictionary.add "flowannual", vbChecked
                    pValueDictionary.add "flowannualPrty", pTargetValue
                Case "-2":
                    pValueDictionary.add "flowstormpeak", vbChecked
                    pValueDictionary.add "flowstormpeakPrty", pTargetValue
                Case "-3":
                    pValueDictionary.add "flowfrequency", vbChecked
                    pValueDictionary.add "flowfrequencyPrty", pTargetValue
                    pValueDictionary.add "flowCFS", pCalcDays
            End Select
            
        Else    '*** FOR POLLUTANTS: FACTOR-GROUP: 1,2,3 for pollutants
            Select Case pFactorType
                Case "1":   'AAL
                    pValueDictionary.add "Pollutant" & pFactorGroup & "Load", vbChecked
                    pValueDictionary.add "Pollutant" & pFactorGroup & "LoadPrty", pTargetValue
                Case "2":   'AAC
                    pValueDictionary.add "Pollutant" & pFactorGroup & "Concentration", vbChecked
                    pValueDictionary.add "Pollutant" & pFactorGroup & "ConcentrationPrty", pTargetValue
                Case "3":   'MAC
                    pValueDictionary.add "Pollutant" & pFactorGroup & "MaxDaily", vbChecked
                    pValueDictionary.add "Pollutant" & pFactorGroup & "MaxDailyPrty", pTargetValue
                    pValueDictionary.add "Pollutant" & pFactorGroup & "Days", pCalcDays
            End Select
        End If
        
        Set pRow = pCursor.NextRow
    Loop
    
    'Set the memory release
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pOptimizationTable = Nothing
    
    'Return the value dictionary
    Set ReadOptimizationParametersForMaximizeBenefit = pValueDictionary
    
    Exit Function
ShowError:
    MsgBox "ReadOptimizationParametersForMaximizeBenefit: " & Err.description
End Function


'** Subroutine to read optimization table values
Public Function ReadOptimizationParametersForMinimizeCost(pBMPID As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    'Get landuse reclassification table: LUReclass, Create new if not found
     Dim pOptimizationTable As iTable
     Set pOptimizationTable = GetInputDataTable("OptimizationDetail")
     If (pOptimizationTable Is Nothing) Then
        Exit Function
     End If
     
    'Query in the table for existing records
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = " & pBMPID
    Dim pCursor As ICursor
    Set pCursor = pOptimizationTable.Search(pQueryFilter, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    
    Dim pRowString As String
    Dim pSplittedString
    Dim pFactorName As String
    Dim pFactorGroup As String
    Dim pFactorType As String
    Dim pCalcDays As Double
    Dim pCalcMode As String
    Dim pTargetValue As Double
    Dim pValueDictionary As Scripting.Dictionary
    Set pValueDictionary = CreateObject("Scripting.Dictionary")
    Do While Not pRow Is Nothing
        pRowString = pRow.value(iPropValue)
        pSplittedString = Split(pRowString, ",")
        pFactorGroup = pSplittedString(0)  'FACTOR-GROUP: 1,2,3 for pollutants, -1 for flow
        pFactorType = pSplittedString(1)   'FACTOR-TYPE: -1, -2, -3 for AAFV, PDF, FEF, 1,2,3 for AAL, AAC, MAC
        pCalcDays = CDbl(pSplittedString(2))
        pCalcMode = pSplittedString(3)
        pTargetValue = CDbl(pSplittedString(4))
        
        
        If (pFactorGroup = -1) Then     'FOR FLOW
            Select Case pFactorType
                Case "-1":
                    pValueDictionary.add "_Flow_AAFV", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Flow_AAFV", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Flow_AAFV", pTargetValue
                Case "-2":
                    pValueDictionary.add "_Flow_PDF", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Flow_PDF", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Flow_PDF", pTargetValue
                Case "-3":
                    pValueDictionary.add "_Flow_FEF", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Flow_FEF", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Flow_FEF", pTargetValue
                    pValueDictionary.add "CalcDays" & "_Flow_FEF", pCalcDays
            End Select
            
        Else    '*** FOR POLLUTANTS: FACTOR-GROUP: 1,2,3 for pollutants
            Select Case pFactorType
                Case "1":   'AAL
                    pValueDictionary.add "_Pollutant" & pFactorGroup & "_AAL", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Pollutant" & pFactorGroup & "_AAL", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Pollutant" & pFactorGroup & "_AAL", pTargetValue
                Case "2":   'AAC
                    pValueDictionary.add "_Pollutant" & pFactorGroup & "_AAC", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Pollutant" & pFactorGroup & "_AAC", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Pollutant" & pFactorGroup & "_AAC", pTargetValue
                Case "3":   'MAC
                    pValueDictionary.add "_Pollutant" & pFactorGroup & "_MAC", vbChecked
                    pValueDictionary.add "Option" & pCalcMode & "_Pollutant" & pFactorGroup & "_MAC", vbChecked
                    pValueDictionary.add "TextBox" & pCalcMode & "_Pollutant" & pFactorGroup & "_MAC", pTargetValue
                    pValueDictionary.add "CalcDays" & "_Pollutant" & pFactorGroup & "_MAC", pCalcDays
            End Select
        End If
        
        Set pRow = pCursor.NextRow
    Loop
    
    'Set the memory release
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pQueryFilter = Nothing
    Set pOptimizationTable = Nothing
    
    'Return the value dictionary
    Set ReadOptimizationParametersForMinimizeCost = pValueDictionary
    
    Exit Function
ShowError:
    MsgBox "ReadOptimizationParametersForMinimizeCost: " & Err.description
End Function


Public Sub ActivateAssessmentTool()
    Dim pbars As ICommandBars
    Set pbars = gApplication.Document.CommandBars
    
    'show the Edit BMP toolbar
    Dim pUID As UID
    Set pUID = New UID
    'Below is a toolbar in this project (BMPTool..."
    pUID.value = "SUSTAIN.EditBMPToolbar"
    Dim pbar As ICommandBar
    Set pbar = pbars.Find(pUID)
    pbar.Dock esriDockShow
        
    'Make the define assessment point tool as active tool
    Dim pUID1 As UID
    Set pUID1 = New UID
    'Below is a toolbar in this project (BMPTool..."
    pUID1.value = "SUSTAIN.DefineAssessPoint"
    
    Dim pSelectTool As ICommandItem
    Set pSelectTool = pbars.Find(pUID1)
     
    'Set the current tool of the application to be the Select Graphics Tool
    Set gApplication.CurrentTool = pSelectTool

End Sub
