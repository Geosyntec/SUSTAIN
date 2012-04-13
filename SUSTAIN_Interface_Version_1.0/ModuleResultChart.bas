Attribute VB_Name = "ModuleResultChart"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleResultChart
'   Purpose:     Plots the optimization results for the selected assessment points
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  03/../2005 - Sabu Paul
'                Modified: 03/../2005 -
'
'******************************************************************************


Option Explicit
Option Base 0
'*******************************************************************************
'Subroutine : InitEvaluationChart
'Purpose    : Initialize optimization result chart
'Arguments  : Id of the BMP site, path of the simulation output directory
'Author     : Sabu Paul
'History    :
'*******************************************************************************

Public Function InitEvaluationChart(pBMPID As Integer, pOutputDirName As String) As Boolean
On Error GoTo ErrorHandler:
  
    FrmResultChart.bmpId.Text = pBMPID
    FrmResultChart.OutputDir.Text = pOutputDirName
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
      
    Dim evalFactorCodeArray
    Dim evalFactorCodeCount As Integer
    evalFactorCodeArray = GetEvalFactorList(pBMPID, pOutputDirName)
    If (evalFactorCodeArray(0) = "") Then
        InitEvaluationChart = False
        GoTo CleanUp
    End If
    evalFactorCodeCount = UBound(evalFactorCodeArray) + 1
        
    Dim pCurrentCost As Double
    Dim pCurrentTarget As Double
    
    Dim costs
    costs = GetCostValues(pOutputDirName)
    
    Dim initEvalFactorArray
    initEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "Init_Eval.out", evalFactorCodeArray)
    If (initEvalFactorArray(0) = -9999.9) Then
       InitEvaluationChart = False
       GoTo CleanUp
    End If
    
    Dim preDevEvalFactorArray
    preDevEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "PreDev_Eval.out", evalFactorCodeArray)
    If (preDevEvalFactorArray(0) = -9999.9) Then
       InitEvaluationChart = False
       GoTo CleanUp
    End If
       
    Dim postDevEvalFactorArray
    postDevEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "PostDev_Eval.out", evalFactorCodeArray)
    If (postDevEvalFactorArray(0) = -9999.9) Then
       InitEvaluationChart = False
       GoTo CleanUp
    End If
    
    Dim bestEvalFactorArray
    Dim testDict As Scripting.Dictionary
    Dim costIncr As Integer
    Dim tempCount As Integer
    
    Dim targets
    targets = GetTargetValues(pBMPID, pOutputDirName, evalFactorCodeArray, postDevEvalFactorArray, preDevEvalFactorArray)
    
    If (costs(0) <> -9999.9) Then 'only if the bestsolutions files exists
        
        ReDim bestEvalFactorArray(0 To UBound(costs), (evalFactorCodeCount - 1))
        
        'Loop through the number of best solutions
        'Total number of best solutions will be equal
        'to the number of elements in the costs array
                
        Dim tempArray
        For costIncr = 0 To UBound(costs)
            tempArray = GetEvalValues(pBMPID, pOutputDirName, "Best" & costIncr + 1 & "_Eval.out", evalFactorCodeArray, postDevEvalFactorArray, preDevEvalFactorArray)
            If (tempArray(0) <> -9999.9) Then
                For tempCount = 0 To UBound(tempArray)
                    bestEvalFactorArray(costIncr, tempCount) = tempArray(tempCount)
                Next tempCount
            Else
                bestEvalFactorArray(costIncr, 0) = -9999.99
            End If
        Next costIncr
    End If
    
    
    'Label the chart tabs
    Dim tabNames As Scripting.Dictionary
    Set tabNames = GetEvalNameDict(evalFactorCodeArray)
    
    For tempCount = 0 To evalFactorCodeCount - 1
        FrmResultChart.TabCharts.TabCaption(tempCount) = tabNames.Item(evalFactorCodeArray(tempCount))
    Next tempCount
    
    
    Dim yAxisTitle As String
    Dim curFactName As String
    'Set the chart properties
    For tempCount = 0 To evalFactorCodeCount - 1
    'For tempCount = 0 To 0
        Set testDict = New Scripting.Dictionary
        testDict.add "PreDev", preDevEvalFactorArray(tempCount)
        testDict.add "PostDev", postDevEvalFactorArray(tempCount)
        testDict.add "Existing", initEvalFactorArray(tempCount)
        
        curFactName = tabNames.Item(evalFactorCodeArray(tempCount))
        
        If (curFactName = "Flow Volume") Then
                yAxisTitle = "Flow Volume (ft3/yr)"
        ElseIf (curFactName = "Flow Rate") Then
                yAxisTitle = "Flow Rate (cfs)"
        ElseIf (curFactName = "Flow Exceedence Frequency") Then
                yAxisTitle = "Flow Exceedence Frequency (/yr)"
        ElseIf StringContains(curFactName, "Load") Then
                yAxisTitle = "Pollutant Load (lb/yr)"
        ElseIf StringContains(curFactName, "Mean Conc.") Then
                yAxisTitle = "Pollutant Conc. (mg/L)"
        ElseIf StringContains(curFactName, "Conc.") Then
                yAxisTitle = "Pollutant Conc. (mg/L)"
        End If
        
        If (costs(0) <> -9999.9) Then
            For costIncr = 0 To UBound(costs)
                If (bestEvalFactorArray(costIncr, 0) <> -9999.99) Then
                    testDict.add "Best" & costIncr + 1, bestEvalFactorArray(costIncr, tempCount)
                End If
            Next costIncr
        End If
        
        
        'Targets are defined, only if costs are defined
        pCurrentTarget = -99
        If (targets(0) <> -9999.9) Then
            pCurrentTarget = targets(tempCount)
        End If
    
        '** plot the chart
        Call FrmResultChart.PlotEvaluationChart(testDict, tempCount + 1, pCurrentTarget, , yAxisTitle)
              
    Next tempCount

    'Display all costs values if present
    Call FrmResultChart.DisplayCostValuesOnChartFrame(costs)

    InitEvaluationChart = True
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error in Initializing Evaluation Chart " & Err.description
CleanUp:
End Function

'*******************************************************************************
'Subroutine : GetEvalFactorList
'Purpose    : Creates a list of evaluation factors based on the initial result file
'Arguments  : Id of the BMP site, path of the simulation output directory
'Author     : Sabu Paul
'History    :
'*******************************************************************************

Public Function GetEvalFactorList(pBMPID As Integer, pOutputDirName As String) As Variant
On Error GoTo ErrorHandler:

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim initResFileName As String
    initResFileName = pOutputDirName & "\Init_Eval.out"
    
    Dim evalFactorCodeArray()
    Dim evalFactorCodeCount As Integer
    
    If (Not fso.FileExists(initResFileName)) Then
        ReDim Preserve evalFactorCodeArray(0)
        evalFactorCodeArray(0) = ""
        GetEvalFactorList = evalFactorCodeArray
        GoTo CleanUp
    End If
    Dim pTS As TextStream
    Set pTS = fso.OpenTextFile(initResFileName, ForReading, False, TristateUseDefault)
    
    Dim doContinue As Boolean
    Dim pString As String
    Dim pWords

    doContinue = True
    Do While doContinue
        pString = pTS.ReadLine
        If StringContains(pString, "Assessment Point (ID)     Factor Name     Factor Value") Then
            'pString = pTS.ReadLine
            doContinue = False
        End If
    Loop
    
    Do While Not pTS.AtEndOfStream
        pString = pTS.ReadLine
        pWords = Split(pString, vbTab, , vbTextCompare)
        If (CInt(pWords(0)) = pBMPID) Then
             ReDim Preserve evalFactorCodeArray(evalFactorCodeCount)
             evalFactorCodeArray(evalFactorCodeCount) = CStr(pWords(1))
             evalFactorCodeCount = evalFactorCodeCount + 1
        End If
    Loop
    pTS.Close
    
    If (evalFactorCodeCount > 0) Then
        GetEvalFactorList = evalFactorCodeArray
        GoTo CleanUp
    Else
        ReDim Preserve evalFactorCodeArray(0)
        evalFactorCodeArray(0) = ""
        GetEvalFactorList = evalFactorCodeArray
        GoTo CleanUp
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error Reading Initial Evaluation Values " & Err.description
CleanUp:
    Set pTS = Nothing
    Set fso = Nothing
End Function

'*******************************************************************************
'Subroutine : GetEvalValues
'Purpose    : Gets evaluation factor values for a selected results (initial, pre-developed, or best solutions)
'Arguments  : Id of the BMP site, path of the simulation output directory, result file name,
'             evaluation factor list,initial result array,pre-developed result array
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetEvalValues(pBMPID As Integer, pOutputDirName As String, pFileName As String, evalFactorCodeArray As Variant, _
                Optional postDevEvalFactorArray As Variant, Optional preDevEvalFactorArray As Variant) As Variant
On Error GoTo ErrorHandler:

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim resFileName As String
    resFileName = pOutputDirName & "\" & pFileName
       
    Dim evalValuesArray()
    Dim pWords
    
    'if the file is missing return nothing
    If (Not fso.FileExists(resFileName)) Then
        ReDim Preserve evalValuesArray(0)
        evalValuesArray(0) = -9999.9
        GetEvalValues = evalValuesArray
    End If
    
    Dim pTS As TextStream
    
    Set pTS = fso.OpenTextFile(resFileName, ForReading, False, TristateUseDefault)
    Dim doContinue As Boolean
    Dim pString As String
    
    doContinue = True
    Do While doContinue
        pString = pTS.ReadLine
        If StringContains(pString, "Assessment Point (ID)     Factor Name     Factor Value") Then
            'pString = pTS.ReadLine
            doContinue = False
        End If
    Loop
    
    Dim tempCount As Integer
    tempCount = 0
    
    Do While Not pTS.AtEndOfStream
        pString = pTS.ReadLine
        pWords = Split(pString, vbTab, , vbTextCompare)
        If (UBound(pWords) = 0) Then
            pWords = CustomSplit(pString)
        End If
        If (CInt(pWords(0)) = pBMPID) Then
            'If (evalFactorCodeArray(tempCount) = pWords(1)) Then
            If (StringContains(CStr(pWords(1)), CStr(evalFactorCodeArray(tempCount)))) Then
                ReDim Preserve evalValuesArray(tempCount)
                'evalValuesArray(tempCount) = CDbl(pWords(2))
                
                If ((Not IsMissing(postDevEvalFactorArray)) And (Not IsMissing(preDevEvalFactorArray))) Then
                    If (StringContains(CStr(pWords(1)), "_%")) Then
                        evalValuesArray(tempCount) = CDbl(pWords(2)) * postDevEvalFactorArray(tempCount) / 100 'precent of existing
                    ElseIf (StringContains(CStr(pWords(1)), "_S")) Then
                        evalValuesArray(tempCount) = preDevEvalFactorArray(tempCount) + (CDbl(pWords(2)) * (postDevEvalFactorArray(tempCount) - preDevEvalFactorArray(tempCount))) 'scale between predeveloped and existing
                    Else
                        evalValuesArray(tempCount) = CDbl(pWords(2))
                    End If
                
                Else
                    evalValuesArray(tempCount) = CDbl(pWords(2))
                End If
                tempCount = tempCount + 1
            End If
        End If
        
    Loop
    pTS.Close
    
    If (tempCount > 0) Then
        GetEvalValues = evalValuesArray
    Else
        ReDim Preserve evalValuesArray(0)
        evalValuesArray(0) = -9999.9
        GetEvalValues = evalValuesArray
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error Reading Evaluation Values for " & pFileName & "  " & Err.description
CleanUp:
    Set pTS = Nothing
    Set fso = Nothing
End Function


'*******************************************************************************
'Subroutine : GetTargetValues
'Purpose    : Gets target values from the first best solution file
'Arguments  : Id of the BMP site, path of the simulation output directory
'             evaluation factor list,initial result array,pre-developed result array
'Author     : Sabu Paul
'History    :
'*******************************************************************************

Public Function GetTargetValues(pBMPID As Integer, pOutputDirName As String, evalFactorCodeArray As Variant, _
                        postDevEvalFactorArray As Variant, preDevEvalFactorArray As Variant) As Variant
On Error GoTo ErrorHandler:

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim resFileName As String
    resFileName = pOutputDirName & "\Best1_Eval.out"
       
    'if the file is missing return -9999.9 array
    
    Dim targetValuesArray()
    If (Not fso.FileExists(resFileName)) Then
        ReDim Preserve targetValuesArray(0)
        targetValuesArray(0) = -9999.9
        GetTargetValues = targetValuesArray
        GoTo CleanUp
    End If
    
    Dim pTS As TextStream
    Dim pString As String
    Dim pWords

    Set pTS = fso.OpenTextFile(resFileName, ForReading, False, TristateUseDefault)
    Dim doContinue As Boolean
    doContinue = True
    Do While doContinue
        pString = pTS.ReadLine
        If StringContains(pString, "Assessment Point (ID)     Factor Name     Factor Value") Then
            'pString = pTS.ReadLine
            doContinue = False
        End If
    Loop
    
    Dim tempCount As Integer
    tempCount = 0
    
    
    Do While Not pTS.AtEndOfStream
        pString = pTS.ReadLine
        pWords = Split(pString, vbTab, , vbTextCompare)
        If (UBound(pWords) = 0) Then
            pWords = CustomSplit(pString)
        End If
        If (CInt(pWords(0)) = pBMPID) Then
            'If (evalFactorCodeArray(tempCount) = pWords(1)) Then
            If (StringContains(CStr(pWords(1)), CStr(evalFactorCodeArray(tempCount)))) Then
                If (UBound(pWords) >= 3) Then
                    ReDim Preserve targetValuesArray(tempCount)
                    
                    'targetValuesArray(tempCount) = CDbl(pWords(3))
                    
                    If (StringContains(CStr(pWords(1)), "_%")) Then
                        targetValuesArray(tempCount) = CDbl(pWords(3)) * postDevEvalFactorArray(tempCount) / 100 'precent of existing
                    ElseIf (StringContains(CStr(pWords(1)), "_S")) Then
                        targetValuesArray(tempCount) = preDevEvalFactorArray(tempCount) + (CDbl(pWords(3)) * (postDevEvalFactorArray(tempCount) - preDevEvalFactorArray(tempCount))) 'scale between predeveloped and existing
                    Else
                        targetValuesArray(tempCount) = CDbl(pWords(3))
                    End If
                    
                    tempCount = tempCount + 1
                End If
            End If
        End If
    Loop
    pTS.Close
    
    If (tempCount > 0) Then
        GetTargetValues = targetValuesArray
        GoTo CleanUp
    Else
        ReDim Preserve targetValuesArray(0)
        targetValuesArray(0) = -9999.9
        GetTargetValues = targetValuesArray
        GoTo CleanUp
    End If
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error Reading Evaluation Values for" & resFileName & "  " & Err.description
CleanUp:
    Set pTS = Nothing
    Set fso = Nothing
End Function
'*******************************************************************************
'Subroutine : GetCostValues
'Purpose    : Gets costs for the best BMP placement options
'Arguments  : Path of the simulation output directory
'
'Author     : Sabu Paul
'History    :
'*******************************************************************************

Public Function GetCostValues(pOutputDirName As String) As Variant
On Error GoTo ErrorHandler:

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim bestSolutionsFileName As String
    bestSolutionsFileName = pOutputDirName & "\BestSolutions.out"
    
    Dim tempCount As Integer
    tempCount = 0
    
    Dim costs()
    
    Dim pTS As TextStream
    Dim pString As String
    Dim pWords

    If (fso.FileExists(bestSolutionsFileName)) Then
        Set pTS = fso.OpenTextFile(bestSolutionsFileName, ForReading, False, TristateUseDefault)
        pTS.SkipLine ' Skip first three lines
        pTS.SkipLine
        'pTS.SkipLine
        
        Do While Not pTS.AtEndOfStream
            pString = pTS.ReadLine
            ReDim Preserve costs(tempCount)
            pWords = Split(pString, vbTab, , vbTextCompare)
            costs(tempCount) = CDbl(pWords(1))
            tempCount = tempCount + 1
        Loop
        pTS.Close
    'even if bestsolutions file is missing, the chart should be plotted
    Else
        ReDim Preserve costs(0)
        costs(0) = -9999.9
        GetCostValues = costs
    End If
    
    If (tempCount > 0) Then
        GetCostValues = costs
    Else
        ReDim Preserve costs(0)
        costs(0) = -9999.9
        GetCostValues = costs
    End If
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error Reading Evaluation Values for" & bestSolutionsFileName & "  " & Err.description
CleanUp:
    Set pTS = Nothing
    Set fso = Nothing
End Function


'*******************************************************************************
'Subroutine : GetEvalNameDict
'Purpose    : Gets the name list corresponding to the evaluation funtions
'Arguments  : Evaluation factor code list
'
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetEvalNameDict(ByVal pFactorCodeList) As Scripting.Dictionary
    Dim pFactorNameDict As Scripting.Dictionary
    Set pFactorNameDict = New Scripting.Dictionary
    
    Dim fIncr As Long
    Dim curCode As String
    Dim curName As String
    
    'Create the list of pollutants
    Call CreatePollutantList
    
    Dim pollutantIndex As Integer
    For fIncr = 0 To UBound(pFactorCodeList)
        curCode = Trim(pFactorCodeList(fIncr))
        curName = ""
        If (StringContains(curCode, "AAFV")) Then
           curName = "Flow Volume" ' (ft3/yr)
        ElseIf (StringContains(curCode, "PDF")) Then
            curName = "Flow Rate" ' (cfs)
        ElseIf StringContains(curCode, "FEF") Then
            curName = "Flow Exceedence Frequency" ' (cfs)
        ElseIf StringContains(curCode, "AAL") Then
            curName = Replace(curCode, "AAL", "- Load") ' (lb/yr)
        ElseIf StringContains(curCode, "AAC") Then
            curName = Replace(curCode, "AAC", "- Mean Conc.")
        ElseIf StringContains(curCode, "MAC") Then
            curName = Replace(curCode, "MAC", "- Max Day Conc.")
        End If
        
        If (Not pFactorNameDict.Exists(curCode)) Then
            pFactorNameDict.add curCode, curName
        End If
    Next fIncr
    Set GetEvalNameDict = pFactorNameDict
End Function


Public Sub SetGASolution(pOutputDirName As String, ByRef xValues() As Double, ByRef yValues() As Double, pBMPID As Integer, iCalcMode As Integer)
On Error GoTo ErrorHandler

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim finalPopFileName As String
    finalPopFileName = pOutputDirName & "\BestSolutions.out"
        
    Dim evalFactorCodeArray
    Dim evalFactorCodeCount As Integer
    evalFactorCodeArray = GetEvalFactorList(pBMPID, pOutputDirName)
    If (evalFactorCodeArray(0) = "") Then
        GoTo CleanUp
    End If
    
    evalFactorCodeCount = UBound(evalFactorCodeArray) + 1
                
    Dim initEvalFactorArray
    initEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "Init_Eval.out", evalFactorCodeArray)
    If (initEvalFactorArray(0) = -9999.9) Then
       GoTo CleanUp
    End If
    
    Dim preDevEvalFactorArray
    preDevEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "PreDev_Eval.out", evalFactorCodeArray)
    If (preDevEvalFactorArray(0) = -9999.9) Then
       GoTo CleanUp
    End If
       
    Dim postDevEvalFactorArray
    postDevEvalFactorArray = GetEvalValues(pBMPID, pOutputDirName, "PostDev_Eval.out", evalFactorCodeArray)
    If (postDevEvalFactorArray(0) = -9999.9) Then
       GoTo CleanUp
    End If
        
    Dim pTS As TextStream
    Dim pString As String
    Dim pWords

    Dim tempCount As Long
    tempCount = 0

    If (fso.FileExists(finalPopFileName)) Then
        Set pTS = fso.OpenTextFile(finalPopFileName, ForReading, False, TristateUseDefault)
        pTS.SkipLine ' Skip first two lines
        pTS.SkipLine
        
        Do While Not pTS.AtEndOfStream
            pString = pTS.ReadLine
            ReDim Preserve xValues(tempCount)
            ReDim Preserve yValues(tempCount)
            pWords = Split(pString, vbTab, , vbTextCompare)
            'xValues(tempCount) = 100 - CDbl(pWords(1))
            xValues(tempCount) = CDbl(pWords(1))
            'For now plot the results as it comes.
'            Select Case iCalcMode
'                Case 1
'                    yValues(tempCount) = CDbl(pWords(7)) * postDevEvalFactorArray(0) / 100 'precent of existing
'                Case 2
'                    yValues(tempCount) = preDevEvalFactorArray(0) + (CDbl(pWords(7)) * (postDevEvalFactorArray(0) - preDevEvalFactorArray(0))) 'scale between predeveloped and existing
'                Case 3
'                    yValues(tempCount) = CDbl(pWords(7))
'            End Select
            yValues(tempCount) = CDbl(pWords(7))
            
            tempCount = tempCount + 1
        Loop
        pTS.Close
    End If
    
    GoTo CleanUp
ErrorHandler:
    MsgBox "Error in SetGASolution: " & Err.description
CleanUp:
    Set fso = Nothing
End Sub
