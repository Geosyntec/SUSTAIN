Attribute VB_Name = "ModuleResults"
Option Explicit
Public gWeatherInputFile As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
           "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
           ByVal lpFile As String, ByVal lpParameters As String, _
           ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
           

        
Public Sub OpenDoc(pExcelDir As String, pExcelFile As String)
    'open a Word document
    Call ShellExecute(0, "open", pExcelFile, vbNullString, pExcelDir, 1)
End Sub


Public Sub ViewSUSTAINResults(pBMPID As Integer, pBMPType As String)

On Error GoTo ShowError
    
    Dim sFile As String
    
    Dim pOptimizationDetail As iTable
    Set pOptimizationDetail = GetInputDataTable("OptimizationDetail")
    If (pOptimizationDetail Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'OutputFolder'"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    Dim pOutputFolder As String
    pOutputFolder = ""
    If Not (pRow Is Nothing) Then
        pOutputFolder = Trim(pRow.value(iPropValue))
    End If
    
    If (pOutputFolder = "") Then
        MsgBox "Output folder " & pOutputFolder & " not found."
        Exit Sub
    End If
    
    'Get BMP Name
    Dim BMPName As String
    BMPName = pBMPType & "_" & pBMPID
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim excelFileName As String
    excelFileName = "SUSTAIN_PostProcessor.xls"
    'Check if the file is saved properly
    Dim pModelFolder As String
    pModelFolder = ""
  
    pModelFolder = ModuleUtility.GetApplicationPath & "\etc\"
    'name of the file changed to SUSTAIN_PostProcessor.xls (SUSTAINAnalysis.xls)
'    If (fso.FileExists(pModelFolder & "SUSTAIN_PostProcessor.xls")) Then
'        OpenDoc pModelFolder, excelFileName ' "SUSTAIN_PostProcessor.xls"
'    Else
    
    If Not fso.FileExists(pModelFolder & excelFileName) Then
        Dim pattern As String
        pattern = "Excel File (*.xls)|*.xls"
        With FrmFileEditor.CommonDialog
            .DialogTitle = "Select Excel post processor file"
            .Filter = pattern
            .CancelError = True
            .ShowOpen
            If (Err <> cdlCancel) Then
                sFile = .FileName
                pModelFolder = fso.GetParentFolderName(sFile)
                excelFileName = fso.GetBaseName(sFile) & ".xls"
                'OpenDoc pModelFolder, fso.GetBaseName(sFile) & ".xls"
                pModelFolder = pModelFolder & "\"
            End If
        End With
    End If
    
    'Find the weather file
''    If Trim(gWeatherInputFile) = "" Then
''        If gInternalSimulation Then
''            Dim pClimatologyDict As Scripting.Dictionary
''            Set pClimatologyDict = LoadSWMMClimatologyDataToDictionary
''            If Not (pClimatologyDict Is Nothing) Then
''                If (StringContains(pClimatologyDict.Item("Temperature"), "FILE")) Then
''                    gWeatherInputFile = Trim(Replace(Replace(pClimatologyDict.Item("Temperature"), "FILE:", ""), """", ""))
''                    If Not fso.FileExists(gWeatherInputFile) Then gWeatherInputFile = ""
''                End If
''            End If
''        End If
''        If Trim(gWeatherInputFile) = "" Then
''            FrmWeatherData.Show vbModal
''        End If
''    End If
    
    FrmWeatherData.Show vbModal
    
    If gWeatherInputFile = "" Then
        MsgBox "Needed weather data file is missing. Cannot perform storm analysis. Please try after adding weather data", vbExclamation
    End If
    
    Dim configFileName As String
    configFileName = pModelFolder & "Sustain_PP_Config.csv"
    
    Dim configTextStream  As TextStream
    Set configTextStream = fso.CreateTextFile(configFileName, True)
    
    'Write weather data file name
    configTextStream.WriteLine ("Weather File Name, " & gWeatherInputFile)
    
    'Find start and end date
    Dim startDate As String
    Dim endDate As String
    
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'StartDate'"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    If Not (pRow Is Nothing) Then
        startDate = Trim(pRow.value(iPropValue))
    End If
    
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'EndDate'"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    If Not (pRow Is Nothing) Then
        endDate = Trim(pRow.value(iPropValue))
    End If
    
    configTextStream.WriteLine ("StartDate, " & startDate)
    If (startDate = "") Then
        MsgBox "Simulation start date is not defined", vbExclamation
    End If
    
    configTextStream.WriteLine ("EndDate, " & endDate)
    If (endDate = "") Then
        MsgBox "Simulation end date is not defined", vbExclamation
    End If
    
    'Set init, post-, and pre-dev file names
    If (fso.FileExists(pOutputFolder & "\Init_" & BMPName & ".out")) Then
        configTextStream.WriteLine ("Init File Path," & pOutputFolder & "\Init_" & BMPName & ".out")
    Else
        configTextStream.WriteLine ("Init File Path,")
        MsgBox "File " & pOutputFolder & "\Init_" & BMPName & ".out is missing"
    End If
    
    If (fso.FileExists(pOutputFolder & "\PostDev_" & BMPName & ".out")) Then
        configTextStream.WriteLine ("BMPScenario File Path," & pOutputFolder & "\PostDev_" & BMPName & ".out")
    Else
        configTextStream.WriteLine ("BMPScenario File Path,")
        MsgBox "File " & pOutputFolder & "\PostDev_" & BMPName & ".out is missing"
    End If
    
    If (fso.FileExists(pOutputFolder & "\PreDev_" & BMPName & ".out")) Then
        configTextStream.WriteLine ("Predevelopment File Path," & pOutputFolder & "\PreDev_" & BMPName & ".out")
    Else
        configTextStream.WriteLine ("Predevelopment File Path,")
        MsgBox "File " & pOutputFolder & "\PreDev_" & BMPName & ".out is missing"
    End If
    
    'Add input file name
    Dim inputFileName As String
    pQueryFilter.WhereClause = "ID = 0 AND PropName = 'InputFile'"
    Set pCursor = pOptimizationDetail.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    If Not (pRow Is Nothing) Then
        inputFileName = Trim(pRow.value(iPropValue))
    End If
    If (fso.FileExists(inputFileName)) Then
        configTextStream.WriteLine ("Input File Path," & inputFileName)
    Else
        configTextStream.WriteLine ("Input File Path,")
        MsgBox "Input File " & inputFileName & " is missing"
    End If
    
    Call CreatePollutantList
    Dim i As Integer
    configTextStream.WriteLine ("Flow (cfs)")
    For i = 0 To UBound(gPollutants)
        configTextStream.WriteLine (gPollutants(i) & " (lbs/hr)")
    Next
    configTextStream.Close
    
    Set configTextStream = Nothing
    Set fso = Nothing
    
    OpenDoc pModelFolder, excelFileName
    
    Exit Sub
    
    
ShowError:
    MsgBox "ViewSUSTAINResults: " & Err.description
End Sub



