Attribute VB_Name = "ModuleFile"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleFile
'   Purpose:     This module contains all functions and utilities to write the input
'                text file.
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 08/19/2004 - Mira Chokshi
'
'******************************************************************************

Option Explicit
Option Base 0

Public gHasInFileError As Boolean

'Define all private variables
Private fso As Scripting.FileSystemObject
Private pFile As TextStream
Private pInputFilePath As String
Private StrCard715BMPTypes As String
Private StrCard720PointSources As String 'New point source card
Private StrCard725BMPClassA As String
Private StrCard730BMPClassA As String
Private StrCard735BMPClassB As String
Private StrCard740BMPSoilIndex As String
Private StrCard745BMPGrowthIndex As String
Private StrCard750ConduitDimensions As String
Private StrCard755ConduitCrossSections As String
Private StrCard1190ConduitLosses As String
'Private StrCard1200VFSParameters As String ' VFS is handled differently - June 18, 2007
'Private StrCard1210VFSParameters As String
Private StrCard765DecayFactors As String
Private StrCard770PercentRemoval As String
Private StrCard810AdjustParameter As String
Private StrCard805BMPCost As String
Private StrCard800OptimizationControls As String
Private StrCard815Assess As String
Private pBmpDetailDict As Scripting.Dictionary
Private pBasinToBMPDict As Scripting.Dictionary
Private pMaxLandTypeGroupID As Integer
Public pInputFileName As String
Public strTimeStepLine As String
Public gStrStartDate As String
Public gStrEndDate As String
Public pInputFolder As String
Public pOutputFolder As String
Public pPredevelopedLanduse As String
Public pLanduseSimulationOption As Integer
Public pSWMMLanduseOutflowFile As String
Public pSWMMPreDevOutflowFile As String

Private StrCard766KFactors As String
Private StrCard767CValues As String

Private StrCard775Sediment As String
Private StrCard780SandTransport As String
Private StrCard785SiltTransport As String
Private StrCard786ClayTransport As String

Private StrCard901_VFS_Dim As String
Private StrCard902_VFS_SegDetails As String
Private StrCard903_VFS_SoilProps As String
Private StrCard904_VFS_Buf_Sed As String
Private StrCard905_VFS_Sed_Filt As String
Private StrCard906_VFS_Sed_Fracion As String
Private StrCard907_VFS_FO_Adsorbed As String
Private StrCard908_VFS_FO_Dissolved As String
Private StrCard909_VFS_TC_Adsorbed As String
Private StrCard910_VFS_TC_Dissolved As String

Private strETOptions As String
Private strMonETCoeffs As String
    
Private pTotalPollutantCount As Integer

Private gBMPDrainAreaDict As Scripting.Dictionary
Private StrCard790LandTypeRouting As String
'Const c_sModuleFileName As String = "D:\SUSTAIN\VBProject\ModuleFile.bas"

Private gBmpTypeClassDict As Scripting.Dictionary


    
'******************************************************************************
'Subroutine: WriteInputTextFile
'Author:     Mira Chokshi
'Purpose:    Main subroutine to write the output file. Asks the user about
'            the input file name. Calls subroutines to write each card.
'******************************************************************************
Public Function WriteInputTextFile() As Boolean
On Error GoTo ShowError
    
''    BMP Type    BMP Type ID Class
''BioRetentionBasin   BIORETENTION   A
''WetPond WETPOND A
''Cistern CISTERN A
''DryPond DRYPOND A
''InfiltrationTrench  INFILTRATIONTRENCH  A
''GreenRoof   GREENROOF   A
''PorousPavement  POROUSPAVEMENT  A
''RainBarrel  RAINBARREL  A
''VegetativeSwale SWALE   B
''VFS BUFFERSTRIP D
''Conduit CONDUIT C
''Regulator   REGULATOR   X
''VirtualOutlet   VIRTUALOUTLET   X

    Set gBmpTypeClassDict = New Scripting.Dictionary
    gBmpTypeClassDict.Item("BioRetentionBasin") = "BIORETENTION"
    gBmpTypeClassDict.Item("WetPond") = "WETPOND"
    gBmpTypeClassDict.Item("Cistern") = "CISTERN"
    gBmpTypeClassDict.Item("DryPond") = "DRYPOND"
    gBmpTypeClassDict.Item("InfiltrationTrench") = "INFILTRATIONTRENCH"
    gBmpTypeClassDict.Item("GreenRoof") = "GREENROOF"
    gBmpTypeClassDict.Item("PorousPavement") = "POROUSPAVEMENT"
    gBmpTypeClassDict.Item("RainBarrel") = "RAINBARREL"
    gBmpTypeClassDict.Item("VegetativeSwale") = "SWALE"
    gBmpTypeClassDict.Item("VFS") = "BUFFERSTRIP"
    gBmpTypeClassDict.Item("Conduit") = "CONDUIT"
    gBmpTypeClassDict.Item("Regulator") = "REGULATOR"
    gBmpTypeClassDict.Item("VirtualOutlet") = "VIRTUALOUTLET"
    gBmpTypeClassDict.Item("Junction") = "JUNCTION"
    
    gHasInFileError = False
    
    If (CheckInputFileDataRequirements = False) Then
        WriteInputTextFile = False
        Exit Function
    End If
    
    pInputFileName = ""
    FrmSimulationPeriod.Show vbModal
        
    If (pInputFileName = "") And (Not gHasInFileError) Then WriteInputTextFile = False: Exit Function
    
    If gHasInFileError Then Err.Raise 5001, "Error in setting simulation options"
    
'    If (pInputFileName = "") Then
'        WriteInputTextFile = False
'        Exit Function
'    End If
    Set pBmpDetailDict = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
     'Clear memory
     StrCard715BMPTypes = ""
     StrCard720PointSources = ""
     StrCard725BMPClassA = ""
     StrCard730BMPClassA = ""
     StrCard735BMPClassB = ""
     StrCard740BMPSoilIndex = ""
     StrCard745BMPGrowthIndex = ""
     StrCard750ConduitDimensions = ""
     StrCard755ConduitCrossSections = ""
     StrCard1190ConduitLosses = ""
     StrCard765DecayFactors = ""
     StrCard770PercentRemoval = ""
     StrCard810AdjustParameter = ""
     StrCard805BMPCost = ""
     StrCard800OptimizationControls = ""
     StrCard815Assess = ""
     
     StrCard766KFactors = ""
     StrCard767CValues = ""
     
     StrCard775Sediment = ""
     StrCard780SandTransport = ""
     StrCard785SiltTransport = ""
     StrCard786ClayTransport = ""

     StrCard901_VFS_Dim = ""
     StrCard902_VFS_SegDetails = ""
     StrCard903_VFS_SoilProps = ""
     StrCard904_VFS_Buf_Sed = ""
     StrCard905_VFS_Sed_Filt = ""
     StrCard906_VFS_Sed_Fracion = ""
     StrCard907_VFS_FO_Adsorbed = ""
     StrCard908_VFS_FO_Dissolved = ""
     StrCard909_VFS_TC_Adsorbed = ""
     StrCard910_VFS_TC_Dissolved = ""
    
    strETOptions = ""
    strMonETCoeffs = ""
    
    Dim totalNumCards As Integer
    totalNumCards = 26
    
    Dim pVFSFLayer As IFeatureLayer
    Set pVFSFLayer = GetInputFeatureLayer("VFS")
    If Not pVFSFLayer Is Nothing Then totalNumCards = totalNumCards + 11
    
    frmSplash.Show vbModeless
    frmSplash.Refresh
    AlwaysOnTop frmSplash, -1
    
    frmSplash.lblProductName.Caption = "Writing input file. Please wait!!!"
        
    Set pFile = fso.OpenTextFile(pInputFileName, ForWriting, True, TristateUseDefault)
    'Call subroutines to write card data to input file
    frmSplash.lblProcess.Caption = "1 of " & totalNumCards & ". Writing card 700 - "
    frmSplash.Refresh
    WriteCard700 'WriteCard1000
    
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 700"
    
    frmSplash.lblProcess.Caption = "2 of " & totalNumCards & ". Writing card 705 - "
    frmSplash.Refresh
    WriteCard705
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 705"
    
    frmSplash.lblProcess.Caption = "3 of " & totalNumCards & ". Writing card 710 - "
    frmSplash.Refresh
    WriteCard710
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 710"
    
    Call GetLandTypeRouting
    If gHasInFileError Then Err.Raise 5001, "Error in calculating land use distribution"
    
    frmSplash.lblProcess.Caption = "4 of " & totalNumCards & ". Writing card 715 - "
    frmSplash.Refresh
    'WriteCard1020 ' the information in this card is included in c700
    WriteCard715
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 715"
    
    frmSplash.lblProcess.Caption = "5 of " & totalNumCards & ". Writing card 720 - "
    frmSplash.Refresh
    WriteCard720
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 720"
    
    frmSplash.lblProcess.Caption = "6 of " & totalNumCards & ". Writing card 725 - "
    frmSplash.Refresh
    WriteCard725
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 725"
    
    frmSplash.lblProcess.Caption = "7 of " & totalNumCards & ". Writing card 730 - "
    frmSplash.Refresh
    WriteCard730
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 730"
    
    frmSplash.lblProcess.Caption = "8 of " & totalNumCards & ". Writing card 735 - "
    frmSplash.Refresh
    WriteCard735
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 735"
    
    frmSplash.lblProcess.Caption = "9 of " & totalNumCards & ". Writing card 740 - "
    frmSplash.Refresh
    
    WriteCard740
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 740"
    
    frmSplash.lblProcess.Caption = "10 of " & totalNumCards & ". Writing card 745 - "
    frmSplash.Refresh
    
    WriteCard745
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 745"
    
    frmSplash.lblProcess.Caption = "11 of " & totalNumCards & ". Writing card 750 - "
    frmSplash.Refresh
    
    WriteCard750
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 750"
    
    frmSplash.lblProcess.Caption = "12 of " & totalNumCards & ". Writing card 755 - "
    frmSplash.Refresh
    
    WriteCard755
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 755"
    'WriteCard1190  'added the card details to c755 - June 18, 2007
'    WriteCard1200 ' VFS is handled differently - June 18, 2007
'    WriteCard1210
    
    frmSplash.lblProcess.Caption = "13 of " & totalNumCards & ". Writing card 765 - "
    frmSplash.Refresh
    
    WriteCard765
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 765"
    
    frmSplash.lblProcess.Caption = "14 of " & totalNumCards & ". Writing card 766 - "
    frmSplash.Refresh
    
    WriteCard766
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 766"
    
    frmSplash.lblProcess.Caption = "15 of " & totalNumCards & ". Writing card 767 - "
    frmSplash.Refresh
    
    WriteCard767
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 767"
    
    frmSplash.lblProcess.Caption = "16 of " & totalNumCards & ". Writing card 770 - "
    frmSplash.Refresh
    
    WriteCard770
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 770"
    
    frmSplash.lblProcess.Caption = "17 of " & totalNumCards & ". Writing card 775 - "
    frmSplash.Refresh
    
    WriteCard775 ' New sediment card
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 775"
    
    frmSplash.lblProcess.Caption = "18 of " & totalNumCards & ". Writing card 780 - "
    frmSplash.Refresh
    
    WriteCard780
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 780"
    
    frmSplash.lblProcess.Caption = "19 of " & totalNumCards & ". Writing card 785 - "
    frmSplash.Refresh
    
    WriteCard785
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 785"
    
    frmSplash.lblProcess.Caption = "20 of " & totalNumCards & ". Writing card 790 - "
    frmSplash.Refresh
    
    WriteCard786
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 786"
    
    frmSplash.lblProcess.Caption = "21 of " & totalNumCards & ". Writing card 786 - "
    frmSplash.Refresh
    
    WriteCard790
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 790"
    
    frmSplash.lblProcess.Caption = "22 of " & totalNumCards & ". Writing card 795 - "
    frmSplash.Refresh
    
    WriteCard795
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 795"
    
    frmSplash.lblProcess.Caption = "23 of " & totalNumCards & ". Writing card 800 - "
    frmSplash.Refresh
    
    WriteCard800
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 800"
    
    frmSplash.lblProcess.Caption = "24 of " & totalNumCards & ". Writing card 805 - "
    frmSplash.Refresh
    
    WriteCard805
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 805"
    
    frmSplash.lblProcess.Caption = "25 of " & totalNumCards & ". Writing card 810 - "
    frmSplash.Refresh
    
    WriteCard810
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 810"
    
    frmSplash.lblProcess.Caption = "26 of " & totalNumCards & ". Writing card 815 - "
    frmSplash.Refresh
    
    WriteCard815
    If gHasInFileError Then Err.Raise 5001, "Error in writing card 815"
    
    'Write the VFSMOD parameters if there is any filter strip in the project
    
    'If filter strip is present in the study area then write the following cards
    If StrCard901_VFS_Dim <> "" Then
        'FrmVFSSimOptions.Show vbModal
        
        frmSplash.lblProcess.Caption = "27 of " & totalNumCards & ". Writing card 900 - "
        frmSplash.Refresh
        
        WriteCard900
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 900"
        
        frmSplash.lblProcess.Caption = "28 of " & totalNumCards & ". Writing card 901 - "
        frmSplash.Refresh
        
        WriteCard901
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 901"
        
        frmSplash.lblProcess.Caption = "29 of " & totalNumCards & ". Writing card 902 - "
        frmSplash.Refresh
        
        WriteCard902
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 902"
        
        frmSplash.lblProcess.Caption = "30 of " & totalNumCards & ". Writing card 903 - "
        frmSplash.Refresh
        
        WriteCard903
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 903"
        
        frmSplash.lblProcess.Caption = "31 of " & totalNumCards & ". Writing card 904 - "
        frmSplash.Refresh
        
        WriteCard904
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 904"
        
        frmSplash.lblProcess.Caption = "32 of " & totalNumCards & ". Writing card 905 - "
        frmSplash.Refresh
        
        WriteCard905
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 905"
        
        frmSplash.lblProcess.Caption = "33 of " & totalNumCards & ". Writing card 906 - "
        frmSplash.Refresh
        
        WriteCard906
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 906"
        
        frmSplash.lblProcess.Caption = "34 of " & totalNumCards & ". Writing card 907 - "
        frmSplash.Refresh
        
        WriteCard907
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 907"
        
        frmSplash.lblProcess.Caption = "35 of " & totalNumCards & ". Writing card 908 - "
        frmSplash.Refresh
        
        WriteCard908
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 908"
        
        frmSplash.lblProcess.Caption = "36 of " & totalNumCards & ". Writing card 909 - "
        frmSplash.Refresh
        
        WriteCard909
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 909"
        
        frmSplash.lblProcess.Caption = "37 of " & totalNumCards & ". Writing card 910 - "
        frmSplash.Refresh
        WriteCard910
        If gHasInFileError Then Err.Raise 5001, "Error in writing card 910"
    End If
    'Close the input file
    pFile.Close
    Unload frmSplash
    
    MsgBox "Input file created successfully", vbInformation
    Set pFile = Nothing
    Set fso = Nothing
   
    WriteInputTextFile = True
    Exit Function
ShowError:
    Unload frmSplash
    WriteInputTextFile = False
    MsgBox "WriteInputTextFile: " & Err.description
End Function


'******************************************************************************
'Subroutine: CheckInputFileDataRequirements
'Author:     Mira Chokshi
'Purpose:    This subroutine checks for data required for writing input file.
'******************************************************************************
Public Function CheckInputFileDataRequirements() As Boolean
On Error GoTo ShowError

    
    '*** Check for Basin to BMP routing connection
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If (pWatershedFLayer Is Nothing) Then
        CheckInputFileDataRequirements = False
    End If
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "BMPID = 0"
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pWatershedFClass.Search(pQueryFilter, True)
    Dim pFeature As IFeature
    Set pFeature = pFeatureCursor.NextFeature
    Dim iID As Long
    iID = pFeatureCursor.FindField("ID")
    Dim strMessage As String
    strMessage = ""
    Do While Not (pFeature Is Nothing)
        If (strMessage <> "") Then
            strMessage = strMessage & ","
        End If
        strMessage = strMessage & pFeature.value(iID)
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    If (strMessage <> "") Then
        MsgBox "Watershed(s) " & strMessage & " are not routed to BMPs. Please use Basin-BMP routing tool to define the network.", vbExclamation
        CheckInputFileDataRequirements = False
        GoTo CleanUp
    End If
    
    '*** Check for BMP to BMP network connection
    Dim pBMPNetwork As iTable
    Set pBMPNetwork = GetInputDataTable("BMPNetwork")
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim pCollection As Collection
    Set pCollection = New Collection
    pQueryFilter.WhereClause = "DSID = 0"
    Set pCursor = pBMPNetwork.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim iIDFld As Long
    iIDFld = pBMPNetwork.FindField("ID")
    'Loop over all records with DSID = 0, i.e. no downstream defined
    Do While Not (pRow Is Nothing)
        pCollection.add pRow.value(iIDFld)
        Set pRow = pCursor.NextRow
    Loop
    'Loop over this collection to check which ID's don't have upstream
    Dim pID As Integer
    Dim pCount As Long
    Dim i As Integer
    strMessage = ""
    For i = 1 To pCollection.Count
        pID = pCollection.Item(i)
        pQueryFilter.WhereClause = "DSID = " & pID
        pCount = pBMPNetwork.RowCount(pQueryFilter)
        If (pCount = 0) Then
            If (strMessage <> "") Then
                strMessage = strMessage & ","
            End If
            strMessage = strMessage & pID
        End If
    Next
    If (strMessage <> "") Then
        MsgBox "BMPs " & strMessage & " are not connected to other BMPs. Please use BMP network tool to define this network.", vbExclamation
        CheckInputFileDataRequirements = False
        GoTo CleanUp
    End If
        
    'All checks successful, return TRUE
    CheckInputFileDataRequirements = True
    GoTo CleanUp
ShowError:
    MsgBox "CheckInputFileDataRequirements: " & Err.description
CleanUp:

End Function

'******************************************************************************
'Subroutine: WriteCard700
'Author:     Sabu Paul, Mira Chokshi
'Purpose:    Write INPUT/OUTPUT FILE DIRECTORIES
'            Get the start and end date of simulation and simulation time-step.
'Modified:   Mira Chokshi modified to update the simulation parameters from
'            dialog box
'******************************************************************************

Public Sub WriteCard700()
On Error GoTo ShowError
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If (Right(pOutputFolder, 1) <> "\") Then
        pOutputFolder = pOutputFolder & "\"
    End If
    pInputFolder = pOutputFolder & "In\"
    
    If Not (fso.FolderExists(pOutputFolder)) Then
        fso.CreateFolder pOutputFolder
    End If
    If Not (fso.FolderExists(pInputFolder)) Then
        fso.CreateFolder pInputFolder
    End If
    pFile.WriteLine ("c700 Model Controls")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c LINE1 = Land simulation control (0-external,1-internal),")
    pFile.WriteLine ("c         Land output directory (containing land output timeseries),")
    pFile.WriteLine ("c         Mixed landuse output file name (for internal control),")
    pFile.WriteLine ("c         PreDeveloped landuse output file name (for internal control)")
    pFile.WriteLine ("c LINE2 = Start date of simulation (Year Month Day)")
    pFile.WriteLine ("c LINE3 = End date of simulation (Year Month Day)")
    pFile.WriteLine ("c LINE4 = BMP simulation timestep (Min),")
    pFile.WriteLine ("c         Model output control (0-daily,1-hourly),")
    pFile.WriteLine ("c         Model output directory")
    pFile.WriteLine ("c LINE5 = ET Flag (0-constant monthly ET,1-daily ET from the timeseries,2-calculate daily ET from the daily temperature data),")
    pFile.WriteLine ("c         Climate time series file path (required if ET flag is 1 or 2),")
    pFile.WriteLine ("c         Latitude (Decimal degrees) required if ET flag is 2")
    pFile.WriteLine ("c LINE6 = Monthly ET rate (in/day) if ET flag is 0   OR")
    pFile.WriteLine ("c         Monthly pan coefficient (multiplier to ET value) if ET flag is 1   OR")
    pFile.WriteLine ("c         Monthly variable coefficient to calculate ET values")
    pFile.WriteLine ("c")
    If pLanduseSimulationOption = 0 Then
        pFile.WriteLine ("0" & vbTab & pInputFolder)
    Else
        'pFile.WriteLine ("1" & vbTab & pInputFolder & vbTab & pSWMMLanduseOutflowFile & vbTab & pSWMMPreDevOutflowFile)
        pFile.WriteLine ("1" & vbTab & pInputFolder & vbTab & Replace(pSWMMLanduseOutflowFile, """", "") & vbTab & Replace(pSWMMPreDevOutflowFile, """", ""))
    End If
    
    pFile.WriteLine (gStrStartDate)
    pFile.WriteLine (gStrEndDate)
    pFile.WriteLine (strTimeStepLine)
      
    Call SetETOptionString
    
    pFile.WriteLine (strETOptions)
    pFile.WriteLine (strMonETCoeffs)
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 700:", Err.description
End Sub

Private Sub SetETOptionString()
On Error GoTo ErrorHandler
    ReadLayerTagDictionaryToSRCFile

    Dim etOption As Integer
    etOption = gLayerNameDictionary.Item("ETOPTION")
    
    Dim months
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    Dim i As Integer
    For i = 0 To 11
        strMonETCoeffs = strMonETCoeffs & gLayerNameDictionary.Item("MonET" & months(i)) & vbTab
    Next
        
    If etOption = 0 Then
        strETOptions = etOption
    ElseIf etOption = 1 Then
        strETOptions = etOption & vbTab & gLayerNameDictionary.Item("TSFILE")
    Else
        strETOptions = etOption & vbTab & gLayerNameDictionary.Item("TSFILE") & gLayerNameDictionary.Item("LATITUDE")
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in SetETOptionString: " & Err.description
End Sub
'******************************************************************************
'Subroutine: WriteCard1000
'Author:     Mira Chokshi
'Purpose:    Write INPUT/OUTPUT FILE DIRECTORIES
'            Get the start and end date of simulation and simulation time-step.
'Modified:   Mira Chokshi modified to update the simulation parameters from
'            dialog box
'******************************************************************************
''Private Sub WriteCard1000()
''On Error GoTo ShowError
''
''    Dim fso As Scripting.FileSystemObject
''    Set fso = CreateObject("Scripting.FileSystemObject")
''    'pInputFolder = pInputFilePath & "In\"
''    If (Right(pOutputFolder, 1) <> "\") Then
''        pOutputFolder = pOutputFolder & "\"
''    End If
''    pInputFolder = pOutputFolder & "In\"
''    'pOutputFolder = pInputFilePath & "Out\"
''    If Not (fso.FolderExists(pOutputFolder)) Then
''        fso.CreateFolder pOutputFolder
''    End If
''    If Not (fso.FolderExists(pInputFolder)) Then
''        fso.CreateFolder pInputFolder
''    End If
''
''    pFile.WriteLine ("c1000 INPUT/OUTPUT FILE DIRECTORIES")
''    pFile.WriteLine ("c LINE1 = Input Directory")
''    pFile.WriteLine ("c LINE2 = Output Directory")
''    pFile.WriteLine ("c LINE3 = Start Date of Simulation(Year Month Day)")
''    pFile.WriteLine ("c LINE4 = End Date of Simulation(Year Month Day)")
''    pFile.WriteLine ("c LINE5 = Land Simulation (0-External, 1-Internal) BMP Simulation timestep(Min) Output Simulation timestep(0-Daily,1-Hourly) ")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    'Modified for new input file format -- Sabu Paul ; Aug 24, 2004
''    pFile.WriteLine ("1" & vbTab & pInputFolder) 'to avoid confusion between data lines and comment lines
''    pFile.WriteLine ("2" & vbTab & pOutputFolder)
''    pFile.WriteLine (gStrStartDate)
''    pFile.WriteLine (strEndDate)
''    pFile.WriteLine (strTimeStepLine)
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c")
''    Exit Sub
''ShowError:
''    MsgBox "Card 1000:", err.description
''End Sub


'******************************************************************************
'Subroutine: WriteCard705
'Author:     Sabu Paul
'Purpose:    Write pollutant identification card
'
'Modified:
'******************************************************************************
Private Sub WriteCard705()
On Error GoTo ShowError
    pFile.WriteLine ("c705 Pollutant Definition")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c POLLUT_ID   = Unique pollutant identifier (Sequence number same as in land output time series)")
    pFile.WriteLine ("c POLLUT_NAME = Unique pollutant name")
    pFile.WriteLine ("c MULTIPLIER  = Multiplying factor used to convert the pollutant load to lbs (external control)")
    pFile.WriteLine ("c               or the pollutant conc to lb/ft3 (internal control)")
    pFile.WriteLine ("c SED_FLAG    = The sediment flag (0-not sediment,1-sand,2-silt,3-clay,4-total sediment)")
    pFile.WriteLine ("c               if = 4 SEDIMENT will be splitted into sand, silt,and clay based on the fractions defined in card 710.")
    pFile.WriteLine ("c SED_QUAL     = The sediment-associated pollutant flag (0-no, 1-yes)")
    pFile.WriteLine ("c                      if = 1 then SEDIMENT is required in the pollutant list")
    pFile.WriteLine ("c SAND_QFRAC  = The sediment-associated qual-fraction on sand (0-1), only required if SED_QUAL = 1")
    pFile.WriteLine ("c SILT_QFRAC  = The sediment-associated qual-fraction on silt (0-1), only required if SED_QUAL = 1")
    pFile.WriteLine ("c CLAY_QFRAC  = The sediment-associated qual-fraction on clay (0-1), only required if SED_QUAL = 1")

    pFile.WriteLine ("c")
    pFile.WriteLine ("c  POLLUT_ID    POLLUT_NAME    MULTIPLIER     SED_FLAG    SED_QUAL     SAND_QFRAC     SILT_QFRAC     CLAY_QFRAC")
    'Add the pollutant lines.
    pFile.Write GetPollutantDetails
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 705:", Err.description
End Sub

'******************************************************************************
'Subroutine: WriteCard1010
'Author:     Mira Chokshi
'Purpose:    Write LAND DEFINITION -- LEGEND
'******************************************************************************
''Private Sub WriteCard1010()
''    '** Write this information only if external landuse types are selected.
''    If (pLanduseSimulationOption = 0) Then
''        pFile.WriteLine ("c1010 LANDUSE DEFINITION")
''        pFile.WriteLine ("c")
''        pFile.WriteLine ("c LANDTYPE       = Unique landuse definition identifier")
''        pFile.WriteLine ("c LANDNAME       = Landuse name")
''        pFile.WriteLine ("c IMPERVIOUS     = Distinguishes pervious/impervious land unit")
''        pFile.WriteLine ("c TIMESERIESFILE = File name containing input timeseries")
''        pFile.WriteLine ("c")
''        pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''        pFile.WriteLine ("c  LANDTYPE       LANDNAME   IMPERVIOUS     TIMESERIESFILE ")
''        pFile.Write (GetLandTypeDetails)
''        pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    End If
''End Sub

Private Sub WriteCard710()
On Error GoTo ShowError
    '** Write this information only if external landuse types are selected.
    If (pLanduseSimulationOption = 0) Then
        pFile.WriteLine ("c710 LANDUSE DEFINITION (required if land simulation control is external)")
        pFile.WriteLine ("c")
        pFile.WriteLine ("c LANDTYPE       = Unique landuse definition identifier")
        pFile.WriteLine ("c LANDNAME       = Landuse name")
        pFile.WriteLine ("c IMPERVIOUS     = Distinguishes pervious/impervious land unit")
        pFile.WriteLine ("c TIMESERIESFILE = File name containing input timeseries")
        pFile.WriteLine ("c SAND_FRAC      = The fraction of total sediment from the land which is sand (0-1)")
        pFile.WriteLine ("c SILT_FRAC      = The fraction of total sediment from the land which is silt (0-1)")
        pFile.WriteLine ("c CLAY_FRAC      = The fraction of total sediment from the land which is clay (0-1)")
        pFile.WriteLine ("c")
        pFile.WriteLine ("c  LANDTYPE       LANDNAME              IMPERVIOUS      TIMESERIESFILE     SAND_FRAC       SILT_FRAC       CLAY_FRAC")
        pFile.Write (GetLandTypeDetails)
        pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    End If
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub



'******************************************************************************
'Subroutine: WriteCard1020
'Author:     Mira Chokshi
'Purpose:    Write SWMM landuse & predeveloped landuse outflow file paths.
'******************************************************************************
'No need of this card - June 14, 2007
''Private Sub WriteCard1020()
''    '** Write this information only if internal swmm simulation is selected.
''    If (pLanduseSimulationOption = 1) Then
''        pFile.WriteLine ("c1020 SWMM LAND OUTFLOWS")
''        pFile.WriteLine ("c")
''        pFile.WriteLine ("c LAND_OUTFLOW = Land simulation output file")
''        pFile.WriteLine ("c PREDEV_OUTFLOW  = Pre-developed land simulation output file")
''        pFile.WriteLine ("c")
''        pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''        pFile.WriteLine ("c  LAND_OUTFLOW       PREDEV_OUTFLOW")
''        pFile.WriteLine (pSWMMLanduseOutflowFile & vbTab & pSWMMPreDevOutflowFile)
''        pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    End If
''End Sub


'''******************************************************************************
'''Subroutine: WriteCard1100
'''Author:     Mira Chokshi
'''Purpose:    Write BMP DEFINITION -- LEGEND
'''******************************************************************************
''Private Sub WriteCard1100()
''    Call CreatePollutantList
''    pTotalPollutantCount = UBound(gPollutants)  'Get total pollutants
''    'Get total BMP,Conduits count
''    Dim pBMPFeatureLayer As IFeatureLayer
''    Set pBMPFeatureLayer = GetInputFeatureLayer("BMPs")
''    If (pBMPFeatureLayer Is Nothing) Then
''        MsgBox "BMPs feature layer not found."
''        Exit Sub
''    End If
''    Dim pConduitFeatureLayer As IFeatureLayer
''    Set pConduitFeatureLayer = GetInputFeatureLayer("Conduits")
''    If (pConduitFeatureLayer Is Nothing) Then
''        MsgBox "Conduits feature layer not found."
''        Exit Sub
''    End If
''    Dim pVFSFeatureLayer As IFeatureLayer
''    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")
''
''    '** Sum up total BMPs
''    Dim pTotalBMPCount As Integer
''    pTotalBMPCount = pBMPFeatureLayer.FeatureClass.FeatureCount(Nothing)
''    pTotalBMPCount = pTotalBMPCount + pConduitFeatureLayer.FeatureClass.FeatureCount(Nothing)
''
''    '** ADD VFS to total bmps + conduits
''    If (Not pVFSFeatureLayer Is Nothing) Then
''        pTotalBMPCount = pTotalBMPCount + pVFSFeatureLayer.FeatureClass.FeatureCount(Nothing)
''    End If
''
''    Set pBMPFeatureLayer = Nothing
''    Set pConduitFeatureLayer = Nothing
''    'Loop through each bmp sequentially now, modified 04/05/2005, mira chokshi.
''    Dim iBMP As Integer
''    For iBMP = 1 To pTotalBMPCount
''        BMPCardsDetail (iBMP)
''        ConduitCardsDetails (iBMP)
''        VFSCardsDetail (iBMP)
''    Next
''    'Write header information
''    pFile.WriteLine ("c1100 BMP SITE INFORMATION")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE         = Unique BMP site identifier")
''    pFile.WriteLine ("c BMPNAME         = BMP template name or site name, not needed for the model")
''    pFile.WriteLine ("c BMPCLASS        = Distinguishes BMP Class (A, B, ..., X)")
''    pFile.WriteLine ("c DArea           = Total Drainage Area in acre")
''    pFile.WriteLine ("c PreLUType       = Predevelopment Landuse type")
''    pFile.WriteLine ("c X               = X Coords (if blank, means virtual assessment point without drainage area)")
''    pFile.WriteLine ("c Y               = Y Coords (if blank, means virtual assessment point without drainage area)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE         BMPNAME       BMPCLASS      DArea       PreLUType       X       Y")
''    pFile.Write StrCard1100BMPTypes
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''
''End Sub


'******************************************************************************
'Subroutine: WriteCard1100
'Author:     Mira Chokshi
'Purpose:    Write BMP DEFINITION -- LEGEND
'******************************************************************************
Private Sub WriteCard715()
On Error GoTo ShowError
    Call CreatePollutantList
    pTotalPollutantCount = UBound(gPollutants) + 1 'Get total pollutants
    'Get total BMP,Conduits count
    Dim pBMPFeatureLayer As IFeatureLayer
    Set pBMPFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pBMPFeatureLayer Is Nothing) Then
        MsgBox "BMPs feature layer not found."
        Exit Sub
    End If
    Dim pConduitFeatureLayer As IFeatureLayer
    Set pConduitFeatureLayer = GetInputFeatureLayer("Conduits")
    If (pConduitFeatureLayer Is Nothing) Then
        MsgBox "Conduits feature layer not found."
        Exit Sub
    End If
    Dim pVFSFeatureLayer As IFeatureLayer
    Set pVFSFeatureLayer = GetInputFeatureLayer("VFS")

    '** Sum up total BMPs
    Dim pTotalBMPCount As Integer
    pTotalBMPCount = pBMPFeatureLayer.FeatureClass.FeatureCount(Nothing)
    pTotalBMPCount = pTotalBMPCount + pConduitFeatureLayer.FeatureClass.FeatureCount(Nothing)
    
    '** ADD VFS to total bmps + conduits
    If (Not pVFSFeatureLayer Is Nothing) Then
        pTotalBMPCount = pTotalBMPCount + pVFSFeatureLayer.FeatureClass.FeatureCount(Nothing)
    End If
    
    Set pBMPFeatureLayer = Nothing
    Set pConduitFeatureLayer = Nothing
    'Loop through each bmp sequentially now, modified 04/05/2005, mira chokshi.
    Dim iBMP As Integer
    For iBMP = 1 To pTotalBMPCount
        BMPCardsDetail (iBMP)
        ConduitCardsDetails (iBMP)
        VFSCardsDetail (iBMP)
    Next
    'Write header information
    pFile.WriteLine ("c715 BMP SITE INFORMATION")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE         = Unique BMP site identifier")
    pFile.WriteLine ("c BMPNAME         = BMP template name or site name, not needed for the model")
    'pFile.WriteLine ("c BMPCLASS        = Distinguishes BMP Class (A, B, ..., X)")
    pFile.WriteLine ("c BMPTYPE         = Distinguishes BMP Types (Bioretention, Rainbarrel, etc.)")
    pFile.WriteLine ("c DArea           = Total Drainage Area in acre")
    pFile.WriteLine ("c NUMUNIT = Number of BMP structures")
    pFile.WriteLine ("c DDAREA  = Design drainage area of the BMP structure")
    pFile.WriteLine ("c PreLUType       = Predevelopment Landuse type")
'    pFile.WriteLine ("c X               = X Coords (if blank, means virtual assessment point without drainage area)")
'    pFile.WriteLine ("c Y               = Y Coords (if blank, means virtual assessment point without drainage area)")
    pFile.WriteLine ("c")
'    pFile.WriteLine ("c BMPSITE         BMPNAME       BMPCLASS      DArea       PreLUType       X       Y")
    'pFile.WriteLine ("c BMPSITE         BMPNAME       BMPCLASS      DArea       NUMUNIT     DDAREA      PreLUType")
    pFile.WriteLine ("c BMPSITE         BMPNAME       BMPTYPE      DArea       NUMUNIT     DDAREA      PreLUType")
    pFile.Write StrCard715BMPTypes
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

Private Sub WriteCard720()
On Error GoTo ShowError
    Dim pTable As iTable
    Set pTable = GetInputDataTable("ExternalTS")
    
    pFile.WriteLine ("c720 Point Source Definition")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c POINTSOURCE    = Unique point source identifier")
    pFile.WriteLine ("c BMPSITE        = BMP site identifier in card 715")
    pFile.WriteLine ("c MULTIPLIER     = Multiplier applied to the timeseries file")
    pFile.WriteLine ("c TIMESERIESFILE = File name containing input timeseries")
    pFile.WriteLine ("c SAND_FRAC      = The fraction of total sediment which is sand (0-1)")
    pFile.WriteLine ("c SILT_FRAC      = The fraction of total sediment which is silt (0-1)")
    pFile.WriteLine ("c CLAY_FRAC      = The fraction of total sediment which is clay (0-1)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c  POINTSOURCE  BMPSITE   MULTIPLIER     TIMESERIESFILE   SAND_FRAC       SILT_FRAC       CLAY_FRAC")
    If Not pTable Is Nothing Then
        pFile.Write GetPointSourceDetails
    End If
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'''******************************************************************************
'''Subroutine: WriteCard1110
'''Author:     Mira Chokshi
'''Purpose:    Write CLASS A DIMENSION GROUPS
'''Modified:   Jenny Z. required additional information added in card 1110.
'''            Added RELTP, PEOPLE, DDAYS - Mira Chokshi: modified on 08/20/04
'''******************************************************************************
''Private Sub WriteCard1110()
''    pFile.WriteLine ("c1110 CLASS A BMP Site DIMENSION GROUPS")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE = Class A BMP dimension group identifier in card 715")
''    pFile.WriteLine ("c WIDTH   = Basin bottom width (ft) / no of units used for rain barrel or cistern")
''    pFile.WriteLine ("c LENGTH  = Basin bottom length (ft) / diameter (ft) for rain barrel or cistern")
''    pFile.WriteLine ("c OHEIGHT = Orifice Height (ft)")
''    pFile.WriteLine ("c DIAM    = Orifice Diameter (in)")
''    pFile.WriteLine ("c EXTP    = Exit Type   (1 for C=1,2 for C=0.61, 3 for C=0.61, 4 for C=0.5)")
''    pFile.WriteLine ("c RELTP   = Release Type   (1-Cistern, 2-Rain barrel, 3-others)")
''    pFile.WriteLine ("c PEOPLE  = Number of persons (Cistern Option)")
''    pFile.WriteLine ("c DDAYS   = Number of dry days (Rain Barrel Option)")
''    pFile.WriteLine ("c WEIRTP  = Weir Type   (1-Rectangular,2-Triangular)")
''    pFile.WriteLine ("c WEIRH   = Weir Height (ft)")
''    pFile.WriteLine ("c WEIRW   = (weir type 1) Weir width  (ft)")
''    pFile.WriteLine ("c THETA   = (weir type 2) Weir angle  (degrees)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  WIDTH  LENGTH  OHEIGHT    DIAM     EXITYPE   RELEASETYPE   PEOPLE    DDAYS    WEIRTYPE    WEIRH    WEIRW   THETA")
''    pFile.Write (StrCard1110BMPClassA)
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''
''End Sub

'******************************************************************************
'Subroutine: WriteCard725
'Author:     Mira Chokshi
'Purpose:    Write CLASS A DIMENSION GROUPS
'Modified:   Jenny Z. required additional information added in card 1110.
'            Added RELTP, PEOPLE, DDAYS - Mira Chokshi: modified on 08/20/04
'******************************************************************************
Private Sub WriteCard725()
On Error GoTo ShowError
    pFile.WriteLine ("c725 CLASS-A BMP Site Parameters (required if BMPSITE is CLASS-A in card 715)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE = Class A BMP dimension group identifier in card 715")
    pFile.WriteLine ("c WIDTH   = Basin bottom width (ft) / no of units used for rain barrel or cistern")
    pFile.WriteLine ("c LENGTH  = Basin bottom length (ft) / diameter (ft) for rain barrel or cistern")
    pFile.WriteLine ("c OHEIGHT = Orifice Height (ft)")
    pFile.WriteLine ("c DIAM    = Orifice Diameter (in)")
    pFile.WriteLine ("c EXTP    = Exit Type   (1 for C=1,2 for C=0.61, 3 for C=0.61, 4 for C=0.5)")
    pFile.WriteLine ("c RELTP   = Release Type   (1-Cistern, 2-Rain barrel, 3-others)")
    pFile.WriteLine ("c PEOPLE  = Number of persons (Cistern Option)")
    pFile.WriteLine ("c DDAYS   = Number of dry days (Rain Barrel Option)")
    pFile.WriteLine ("c WEIRTP  = Weir Type   (1-Rectangular,2-Triangular)")
    pFile.WriteLine ("c WEIRH   = Weir Height (ft)")
    pFile.WriteLine ("c WEIRW   = (weir type 1) Weir width  (ft)")
    pFile.WriteLine ("c THETA   = (weir type 2) Weir angle  (degrees)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  WIDTH  LENGTH  OHEIGHT    DIAM     EXITYPE   RELEASETYPE   PEOPLE    DDAYS    WEIRTYPE    WEIRH    WEIRW   THETA")
    pFile.Write (StrCard725BMPClassA)
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
'''******************************************************************************
'''Subroutine: WriteCard1120
'''Author:     Mira Chokshi
'''Purpose:    Write CLASS A Cistern Control Water Release Curve
'''Modified:   Mira Chokshi modified on 08/20/04 to add hourly water release
'''            for cistern type of BMP.
'''******************************************************************************
''Private Sub WriteCard1120()
''    pFile.WriteLine ("c1120 CLASS A Cistern Control Water Release Curve")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE = Class A BMP dimension group identifier in card 715")
''    pFile.WriteLine ("c Flow    = Hourly water release per capita from the Cistern Control (ft3/hr/capita)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  FLOW")
''    pFile.Write (StrCard1120BMPClassA)
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard730
'Author:     Mira Chokshi
'Purpose:    Write CLASS A Cistern Control Water Release Curve
'Modified:   Mira Chokshi modified on 08/20/04 to add hourly water release
'            for cistern type of BMP.
'******************************************************************************
Private Sub WriteCard730()
On Error GoTo ShowError
    pFile.WriteLine ("c730  Cistern Control Water Release Curve (applies if release type is cistern in card 720)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE = Class A BMP dimension group identifier in card 715")
    pFile.WriteLine ("c Flow    = Hourly water release per capita from the Cistern Control (ft3/hr/capita)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  FLOW")
    pFile.Write (StrCard730BMPClassA)
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
''
'''******************************************************************************
'''Subroutine: WriteCard1130
'''Author:     Mira Chokshi
'''Purpose:    Write CLASS B DIMENSION GROUPS
'''******************************************************************************
''Private Sub WriteCard1130()
''    pFile.WriteLine ("c1130 CLASS B BMP Site DIMENSION GROUPS")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE = BMP Site identifier in card 715")
''    pFile.WriteLine ("c WIDTH    = basin bottom width (ft)")
''    pFile.WriteLine ("c LENGTH   = basin bottom Length (ft)")
''    pFile.WriteLine ("c MAXDEPTH = Maximum depth of channel (ft)")
''    pFile.WriteLine ("c SLOPE1   = Side slope 1 (ft/ft)")
''    pFile.WriteLine ("c SLOPE2   = Side slope 2 (ft/ft)   (1-4)")
''    pFile.WriteLine ("c SLOPE3   = Side slope 3 (ft/ft)")
''    pFile.WriteLine ("c MANN_N = Manning  's roughness coefficient")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  WIDTH  LENGTH  MAXDEPTH   SLOPE1   SLOPE2   SLOPE3    MANN_N")
''    pFile.Write StrCard1130BMPClassB
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard735
'Author:     Mira Chokshi
'Purpose:    Write CLASS B DIMENSION GROUPS
'******************************************************************************
Private Sub WriteCard735()
On Error GoTo ShowError
    pFile.WriteLine ("c735 CLASS B BMP Site DIMENSION GROUPS")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE = BMP Site identifier in card 715")
    pFile.WriteLine ("c WIDTH    = basin bottom width (ft)")
    pFile.WriteLine ("c LENGTH   = basin bottom Length (ft)")
    pFile.WriteLine ("c MAXDEPTH = Maximum depth of channel (ft)")
    pFile.WriteLine ("c SLOPE1   = Side slope 1 (ft/ft)")
    pFile.WriteLine ("c SLOPE2   = Side slope 2 (ft/ft)   (1-4)")
    pFile.WriteLine ("c SLOPE3   = Side slope 3 (ft/ft)")
    pFile.WriteLine ("c MANN_N = Manning  's roughness coefficient")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  WIDTH  LENGTH  MAXDEPTH   SLOPE1   SLOPE2   SLOPE3    MANN_N")
    pFile.Write StrCard735BMPClassB
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub


'''******************************************************************************
'''Subroutine: WriteCard1140
'''Author:     Mira Chokshi
'''Purpose:    Write BOTTOM SOIL/VEGITATION CHARACTERISTICS FOR HOLTAN EQUATION AND UNDERDRAIN STRUCTURE
'''******************************************************************************
''Private Sub WriteCard1140()
''    pFile.WriteLine ("c1140 BMP Site BOTTOM SOIL/VEGITATION CHARACTERISTICS FOR HOLTAN EQUATION AND UNDERDRAIN STRUCTURE")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c HOLTAN EQUATION:    F = GI * AVEG * (Computed Available Soil Storage)^1.4 + FINFILT")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in c715")
''    pFile.WriteLine ("c SDEPTH    =  Soil Depth (ft)")
''    pFile.WriteLine ("c POROSITY  =  Soil Porosity (0-1)")
''    pFile.WriteLine ("c AVEG      =  Vegitative Parameter A (0.1-1.0) (Empirical)")
''    pFile.WriteLine ("c FINFILT   =  Soil layer infiltration rate (in/hr)")
''    pFile.WriteLine ("c UNDSWITCH =  Consider underdrain (1), Do not consider underdrain (0)")
''    pFile.WriteLine ("c UNDDEPTH  =  Depth of storage media below underdrain")
''    pFile.WriteLine ("c UNDVOID   =  Fraction of underdrain storage depth that is void space (0-1)")
''    pFile.WriteLine ("c UNDINFILT =  Background infiltration rate, below underdrain (in/hr)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE   SDEPTH   POROSITY   AVEG   FINFILT   UNDSWITCH   UNDDEPTH   UNDVOID   UNDINFILT")
''    pFile.Write StrCard1140BMPSoilIndex
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub
'******************************************************************************
'Subroutine: WriteCard740
'Author:     Mira Chokshi
'Purpose:    Write BOTTOM SOIL/VEGITATION CHARACTERISTICS FOR HOLTAN EQUATION AND UNDERDRAIN STRUCTURE
'******************************************************************************
Private Sub WriteCard740()
On Error GoTo ShowError
    pFile.WriteLine ("c740 BMP Site BOTTOM SOIL/VEGITATION CHARACTERISTICS FOR HOLTAN EQUATION AND UNDERDRAIN STRUCTURE")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c HOLTAN EQUATION:    F = GI * AVEG * (Computed Available Soil Storage)^1.4 + FINFILT")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in c715")
    pFile.WriteLine ("c INFILTM   =  Infiltration Method (0-Holtan, 1-Green Ampt)")
    'pFile.WriteLine ("c POLROTM   =  Pollutant Routing Method (0-Plug Flow,1-Completely mixed, >1-number of CSTRs in series)")
    pFile.WriteLine ("c POLROTM   =  Pollutant Routing Method (1-Completely mixed, >1-number of CSTRs in series)")
    pFile.WriteLine ("c POLREMM   =  Pollutant Removal Method (0-1st order decay, 1-kadlec and knight method )")
    pFile.WriteLine ("c SDEPTH    =  Soil Depth (ft)")
    pFile.WriteLine ("c POROSITY  =  Soil Porosity (0-1)")
    pFile.WriteLine ("c FCAPACITY =  Soil Field Capacity (ft/ft)")
    pFile.WriteLine ("c WPOINT    =  Soil Wilting Point (ft/ft)")
    pFile.WriteLine ("c AVEG      =  Vegitative Parameter A (0.1-1.0) (Empirical)")
    pFile.WriteLine ("c FINFILT   =  Soil layer infiltration rate (in/hr)")
    pFile.WriteLine ("c UNDSWITCH =  Consider underdrain (1), Do not consider underdrain (0)")
    pFile.WriteLine ("c UNDDEPTH  =  Depth of storage media below underdrain")
    pFile.WriteLine ("c UNDVOID   =  Fraction of underdrain storage depth that is void space (0-1)")
    pFile.WriteLine ("c UNDINFILT =  Background infiltration rate, below underdrain (in/hr)")
    pFile.WriteLine ("c SUCTION   =  Average value of soil capillary suction along the wetting front, value must be greater than zero (in)")
    pFile.WriteLine ("c HYDCON    =  Soil saturated hydraulic conductivity, value must be greater than zero (in/hr)")
    pFile.WriteLine ("c IMDMAX    =  Difference between soil porosity and initial moisture content, value must be greater than or equal to zero (a fraction)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   INFILTM   POLROTM POLREMM SDEPTH   POROSITY   FCAPACITY   WPOINT  AVEG   FINFILT   UNDSWITCH   UNDDEPTH   UNDVOID   UNDINFILT   SUCTION HYDCON  IMDMAX")
    pFile.Write StrCard740BMPSoilIndex
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
'''******************************************************************************
'''Subroutine: WriteCard1150
'''Author:     Mira Chokshi
'''Purpose:    Write HOLTAN GROWTH INDEX
'''******************************************************************************
''Private Sub WriteCard1150()
''    pFile.WriteLine ("c1150 BMP Site HOLTAN GROWTH INDEX")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c HOLTAN EQUATION:    F = GI * AVEG * (Computed Available Soil Storage)^1.4 + FINFILT")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in card 715")
''    pFile.WriteLine ("c GIi       =  12 monthly values for GI in HOLTAN equation")
''    pFile.WriteLine ("c              Where i = jan, feb, mar ... dec")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c  BMPSITE   jan   feb   mar   apr    may    jun    jul    aug    sep    oct    nov    dec")
''    pFile.Write StrCard1150BMPGrowthIndex
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard745
'Author:     Mira Chokshi
'Purpose:    Write HOLTAN GROWTH INDEX
'******************************************************************************
Private Sub WriteCard745()
On Error GoTo ShowError
    pFile.WriteLine ("c745 BMP Site HOLTAN GROWTH INDEX")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c HOLTAN EQUATION:    F = GI * AVEG * (Computed Available Soil Storage)^1.4 + FINFILT")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in card 715")
    pFile.WriteLine ("c GIi       =  12 monthly values for GI in HOLTAN equation")
    pFile.WriteLine ("c              Where i = jan, feb, mar ... dec")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c  BMPSITE   jan   feb   mar   apr    may    jun    jul    aug    sep    oct    nov    dec")
    pFile.Write StrCard745BMPGrowthIndex
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

''
'''******************************************************************************
'''Subroutine: WriteCard1160
'''Author:     Mira Chokshi
'''Purpose:    Write Conduit Dimension Groups
'''******************************************************************************
''Private Sub WriteCard1160()
''    'Call conduit cards details to get Class C- conduit parameters
''    pFile.WriteLine ("c1160 Class C Conduit Dimension Groups")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c NAME = BMP site identifier in card 715")
''    pFile.WriteLine ("c INLET_NODE = BMP Id at the entrance of the conduit")
''    pFile.WriteLine ("c OUTLET_NODE = BMP Id at the exit of the conduit")
''    pFile.WriteLine ("c LENGTH = Conduit length")
''    pFile.WriteLine ("c MANNING_N = Manning's roughness coefficient")
''    pFile.WriteLine ("c INLET_IEL = Invert Elevation at the entrance of the conduit")
''    pFile.WriteLine ("c OUTLET_IEL = Invert Elevation at the exit of the conduit")
''    pFile.WriteLine ("c INIT_FLOW = Initial flow in the conduit (cfs)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c  NAME   INLET_NODE   OUTLET_NODE  LENGTH   MANNING_N    INLET_IEL    OUTLET_IEL    INIT_FLOW")
''    pFile.Write StrCard1160ConduitDimensions
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub



'******************************************************************************
'Subroutine: WriteCard750
'Author:     Mira Chokshi
'Purpose:    Write Conduit Dimension Groups
'******************************************************************************
Private Sub WriteCard750()
On Error GoTo ShowError
    'Call conduit cards details to get Class C- conduit parameters
    pFile.WriteLine ("c750 Class-C Conduit Parameters (required if BMPSITE is CLASS-C in card 715)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE = BMP site identifier in card 715")
    pFile.WriteLine ("c INLET_NODE = BMP Id at the entrance of the conduit")
    pFile.WriteLine ("c OUTLET_NODE = BMP Id at the exit of the conduit")
    pFile.WriteLine ("c LENGTH = Conduit length")
    pFile.WriteLine ("c MANNING_N = Manning's roughness coefficient")
    pFile.WriteLine ("c INLET_IEL = Invert Elevation at the entrance of the conduit")
    pFile.WriteLine ("c OUTLET_IEL = Invert Elevation at the exit of the conduit")
    pFile.WriteLine ("c INIT_FLOW = Initial flow in the conduit (cfs)")
    pFile.WriteLine ("c INLET_HL    = Head loss coefficient at the entrance of the conduit")
    pFile.WriteLine ("c OUTLET_HL   = Head loss coefficient at the exit of the conduit")
    pFile.WriteLine ("c AVERAGE_HL  = Head loss coefficient along the length of the conduit")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c  BMPSITE   INLET_NODE   OUTLET_NODE  LENGTH   MANNING_N    INLET_IEL    OUTLET_IEL    INIT_FLOW")
    pFile.Write StrCard750ConduitDimensions
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub



'''******************************************************************************
'''Subroutine: WriteCard1170
'''Author:     Mira Chokshi
'''Purpose:    Write Conduit Cross-Section
'''******************************************************************************
''Private Sub WriteCard1170()
''    pFile.WriteLine ("c1170 Class C Conduit Cross Sections")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c LINK = BMP site identifier in card 715")
''    pFile.WriteLine ("c TYPE = Conduit Type (rectangular, circular...)")
''    pFile.WriteLine ("c GEOM1 = Geometric cross-sectional property of the conduit")
''    pFile.WriteLine ("c GEOM2 = Geometric cross-sectional property of the conduit")
''    pFile.WriteLine ("c GEOM3 = Geometric cross-sectional property of the conduit")
''    pFile.WriteLine ("c GEOM4 = Geometric cross-sectional property of the conduit")
''    pFile.WriteLine ("c BARRELS = Number of Barrels in the conduit")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c  LINK   TYPE   GEOM1   GEOM2   GEOM3    GEOM4    BARRELS")
''    pFile.Write StrCard1170ConduitCrossSections
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub
'******************************************************************************
'Subroutine: WriteCard755
'Author:     Mira Chokshi
'Purpose:    Write Conduit Cross-Section
'******************************************************************************
Private Sub WriteCard755()
On Error GoTo ShowError
    pFile.WriteLine ("c755 Class C Conduit Cross Sections")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c LINK = BMP site identifier in card 715")
    pFile.WriteLine ("c TYPE = Conduit Type (rectangular, circular...)")
    pFile.WriteLine ("c GEOM1 = Geometric cross-sectional property of the conduit")
    pFile.WriteLine ("c GEOM2 = Geometric cross-sectional property of the conduit")
    pFile.WriteLine ("c GEOM3 = Geometric cross-sectional property of the conduit")
    pFile.WriteLine ("c GEOM4 = Geometric cross-sectional property of the conduit")
    pFile.WriteLine ("c BARRELS = Number of Barrels in the conduit")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c  LINK   TYPE   GEOM1   GEOM2   GEOM3    GEOM4    BARRELS")
    pFile.Write StrCard755ConduitCrossSections
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'''******************************************************************************
'''Subroutine: WriteCard1190
'''Author:     Mira Chokshi
'''Purpose:    Write Conduit Dimension Groups
'''******************************************************************************
''Private Sub WriteCard1190()
''    'Call conduit cards details to get Class C- conduit parameters
''    pFile.WriteLine ("c1190 Class C Conduit Dimension Groups")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c NAME = BMP site identifier in card 715")
''    pFile.WriteLine ("c INLET = Head loss coefficient at the entrance of the conduit")
''    pFile.WriteLine ("c OUTLET = Head loss coefficient at the exit of the conduit")
''    pFile.WriteLine ("c AVERAGE = Head loss coefficient along the length of the conduit")
''    pFile.WriteLine ("c FLAP_GATE = Flag Gate present to stop conduit backflow")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c  NAME   INLET   OUTLET  AVERAGE   FLAP_GATE")
''    pFile.Write StrCard1190ConduitLosses
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'**************************************
'Commented the following two subs - VFS is handled differently - June 18, 2007
'**************************************

'******************************************************************************
'Subroutine: WriteCard1200
'Author:     Mira Chokshi
'Purpose:    Write Conduit Dimension Groups
'******************************************************************************
''Private Sub WriteCard1200()
''    'Call conduit cards details to get Class D - Buffer Strip parameters
''    pFile.WriteLine ("c1200 CLASS D (Buffer Strip) Site GROUPS")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c   BMPSITE = Class D BMP dimension group identifier in card 715")
''    pFile.WriteLine ("c   Width  = Buffer strip width along the stream (ft)")
''    pFile.WriteLine ("c   INFILT1 =")
''    pFile.WriteLine ("c   INFILT2 =")
''    pFile.WriteLine ("c   INFILT3 =")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  WIDTH  INFILT1  INFILT2  INFILT3")
''    pFile.Write StrCard1200VFSParameters
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub


'******************************************************************************
'Subroutine: WriteCard1210
'Author:     Mira Chokshi
'Purpose:    Write Conduit Dimension Groups
'******************************************************************************
''Private Sub WriteCard1210()
''    'Call conduit cards details to get Class D - Buffer Strip parameters
''    pFile.WriteLine ("c1210 CLASS D (Buffer Strip) Site Segments")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE = Class D BMP dimension group identifier in card 715")
''    pFile.WriteLine ("c SEGMENT   = Segment ID")
''    pFile.WriteLine ("c LENGTH    = Buffer strip length along the flow direction (perpendicular to stream) for each segment (ft)")
''    pFile.WriteLine ("c SLOPE     = Slope (vertical drop / longitudinal distance)")
''    pFile.WriteLine ("c Manning 's n   =")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  SEGMENT  LENGTH  SLOPE   MANNING'S")
''    pFile.Write StrCard1210VFSParameters
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'''******************************************************************************
'''Subroutine: WriteCard1300
'''Author:     Mira Chokshi
'''Purpose:    Write LAND ROUTING NETWORK
'''******************************************************************************
''Private Sub WriteCard1300()
''    pFile.WriteLine ("c1300 LAND ROUTING NETWORK")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c UniqueID   = Identifies an instance of LANDTYPE in SCHEMATIC")
''    pFile.WriteLine ("c LANDTYPE   = Corresponds to LANDTYPE in c710")
''    pFile.WriteLine ("c AREA       = Area of LANDTYPE in ACRES")
''    pFile.WriteLine ("c DS         = UNIQUE ID of DS BMP (0 - no BMP, add to end)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c UniqueID    LANDTYPE     AREA      DS")
''    pFile.Write GetLandTypeRouting
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard790
'Author:     Mira Chokshi
'Purpose:    Write LAND ROUTING NETWORK
'******************************************************************************
Private Sub WriteCard790()
On Error GoTo ShowError
    pFile.WriteLine ("c790 LAND TO BMP ROUTING NETWORK (required for external land simulation control in card 700)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c UniqueID   = Identifies an instance of LANDTYPE in SCHEMATIC")
    pFile.WriteLine ("c LANDTYPE   = Corresponds to LANDTYPE in c710")
    pFile.WriteLine ("c AREA       = Area of LANDTYPE in ACRES")
    pFile.WriteLine ("c DS         = UNIQUE ID of DS BMP (0 - no BMP, add to end)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c UniqueID    LANDTYPE     AREA      DS")
    pFile.Write StrCard790LandTypeRouting 'GetLandTypeRouting
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
'''******************************************************************************
'''Subroutine: WriteCard1310
'''Author:     Mira Chokshi
'''Purpose:    Write BMP ROUTING NETWORK
'''Modified:   Mira Chokshi modified on 08/19/04. The new changes allow easy
'''            implementation of a splitter. Outlet type = "1" indicates no splitter
'''******************************************************************************
''Private Sub WriteCard1310()
''    pFile.WriteLine ("c1310 BMP Site ROUTING NETWORK")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in card 715")
''    pFile.WriteLine ("c OUTLET_TYPE   =  Outlet type (1-total, 2-weir, 3-orifice or channel, 4-underdrain)")
''    pFile.WriteLine ("c DS            =  Downstrem BMP site identifier in card 715 (0 - no BMP, add to end)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE      OUTLET_TYPE      DS")
''    pFile.Write GetBMPNetworkRouting
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard795
'Author:     Mira Chokshi
'Purpose:    Write BMP ROUTING NETWORK
'Modified:   Mira Chokshi modified on 08/19/04. The new changes allow easy
'            implementation of a splitter. Outlet type = "1" indicates no splitter
'******************************************************************************
Private Sub WriteCard795()
On Error GoTo ShowError
    pFile.WriteLine ("c795 BMP Site ROUTING NETWORK")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   =  BMPSITE identifier in card 715")
    pFile.WriteLine ("c OUTLET_TYPE   =  Outlet type (1-total, 2-weir, 3-orifice or channel, 4-underdrain)")
    pFile.WriteLine ("c DS            =  Downstrem BMP site identifier in card 715 (0 - no BMP, add to end)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE      OUTLET_TYPE      DS")
    pFile.Write GetBMPNetworkRouting
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'''******************************************************************************
'''Subroutine: WriteCard1320
'''Author:     Mira Chokshi
'''Purpose:    Write BMP Pollutant Decay/Loss rates
'''******************************************************************************
''Private Sub WriteCard1320()
''    pFile.WriteLine ("c1320 BMP SITE Pollutant Decay/Loss rates")
''    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
''    pFile.WriteLine ("c QUALDECAYi = First-order decay rate for pollutant i (hr^-1)")
''    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  QUALDECAY1  QUALDECAY2 ... QUALDECAYN")
''    pFile.Write StrCard1320DecayFactors
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub
'******************************************************************************
'Subroutine: WriteCard765
'Author:     Mira Chokshi
'Purpose:    Write BMP Pollutant Decay/Loss rates
'******************************************************************************
Private Sub WriteCard765()
On Error GoTo ShowError
    pFile.WriteLine ("c765 BMP SITE Pollutant Decay/Loss rates")
    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
    pFile.WriteLine ("c QUALDECAYi = First-order decay rate for pollutant i (hr^-1)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  QUALDECAY1  QUALDECAY2 ... QUALDECAYN")
    pFile.Write StrCard765DecayFactors
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
Private Sub WriteCard766()
On Error GoTo ShowError
    pFile.WriteLine ("c766 Pollutant K' values (applies when pollutant removal method is kadlec and knight method in card 740)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
    pFile.WriteLine ("C K 'i        = Constant rate for pollutant i (ft/yr)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  QUALK'1  QUALK'2 ... QUALK'N")
    pFile.Write StrCard766KFactors
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
Private Sub WriteCard767()
On Error GoTo ShowError
    pFile.WriteLine ("c767 Pollutant C* values (applies when pollutant removal method is kadlec and knight method in card 740)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
    pFile.WriteLine ("c C*i        = Background concentration for pollutant i (mg/l)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  QUALC*1  QUALC*2 ... QUALC*N")
    pFile.Write StrCard767CValues
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub


'******************************************************************************
'Subroutine: WriteCard775
'Author:     Sabu Paul
'Purpose:    Write sediment parameters
'******************************************************************************
Private Sub WriteCard775()
On Error GoTo ShowError
    pFile.WriteLine ("c775 Sediment General Parameters (required if pollutant type is sediment in card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
    pFile.WriteLine ("c BEDWID     = Bed width (ft) - this is constant for the entire simulation period")
    pFile.WriteLine ("c BEDDEP     = Initial bed depth (ft)")
    pFile.WriteLine ("c BEDPOR     = Bed sediment porosity")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   BEDWID   BEDDEP   BEDPOR")
    pFile.Write StrCard775Sediment
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'******************************************************************************
'Subroutine: WriteCard780
'Author:     Sabu Paul
'Purpose:    Write sediment parameters
'******************************************************************************
Private Sub WriteCard780()
On Error GoTo ShowError
    pFile.WriteLine ("c780 Sand Transport Parameters (required if pollutant type is sediment in card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   = BMP site identifier in card 715")
    pFile.WriteLine ("c D         = Effective diameter of the transported sand particles (in)")
    pFile.WriteLine ("c W         = The corresponding fall velocity in still water (in/sec)")
    pFile.WriteLine ("c RHO       = The density of the sand particles (lb/ft3)")
    pFile.WriteLine ("c KSAND     = The coefficient in the sandload power function formula")
    pFile.WriteLine ("c EXPSND    = The exponent in the sandload power function formula")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   D        W        RHO      KSAND    EXPSND")
    pFile.Write StrCard780SandTransport
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'******************************************************************************
'Subroutine: WriteCard785
'Author:     Sabu Paul
'Purpose:    Write sediment parameters
'******************************************************************************
Private Sub WriteCard785()
On Error GoTo ShowError
    pFile.WriteLine ("c785 Silt Transport Parameters (required if pollutant type is sediment in card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   = BMP site identifier in card 715")
    pFile.WriteLine ("c D         = Effective diameter of the transported silt particles (in)")
    pFile.WriteLine ("c W         = The corresponding fall velocity in still water (in/sec)")
    pFile.WriteLine ("c RHO       = The density of the silt particles (lb/ft3)")
    pFile.WriteLine ("c TAUCD     = The critical bed shear stress for deposition (lb/ft2)")
    pFile.WriteLine ("c TAUCS     = The critical bed shear stress for scour (lb/ft2)")
    pFile.WriteLine ("c M         = The erodibility coefficient of the silt particles (lb/ft2/day)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   D        W        RHO      TAUCD    TAUCS    M")
    pFile.Write StrCard785SiltTransport
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'******************************************************************************
'Subroutine: WriteCard786
'Author:     Sabu Paul
'Purpose:    Write sediment parameters
'******************************************************************************
Private Sub WriteCard786()
On Error GoTo ShowError
    pFile.WriteLine ("c786 Clay Transport Parameters (required if pollutant type is sediment in card 705)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   = BMP site identifier in card 715")
    pFile.WriteLine ("c D         = Effective diameter of the transported clay particles (in)")
    pFile.WriteLine ("c W         = The corresponding fall velocity in still water (in/sec)")
    pFile.WriteLine ("c RHO       = The density of the silt/clay particles (lb/ft3)")
    pFile.WriteLine ("c TAUCD     = The critical bed shear stress for deposition (lb/ft2)")
    pFile.WriteLine ("c TAUCS     = The critical bed shear stress for scour (lb/ft2)")
    pFile.WriteLine ("c M         = The erodibility coefficient of the clay particles (lb/ft2/day)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE   D        W        RHO      TAUCD    TAUCS    M")
    pFile.Write StrCard786ClayTransport
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub


''******************************************************************************
''Subroutine: WriteCard1330
''Author:     Mira Chokshi
''Purpose:    Write BMP Underdrain Pollutant Percent Removal (applies when underdrain is on in card 740)
''******************************************************************************
'Private Sub WriteCard1330()
'    pFile.WriteLine ("c1330 BMP Underdrain Pollutant Percent Removal (applies when underdrain is on in card 740)")
'    pFile.WriteLine ("c BMPSITE     = BMPSITE identifier in card 715")
'    pFile.WriteLine ("c QUALPCTREMi = Perecent Removal for pollutant i through underdrain (0-1)")
'    pFile.WriteLine ("c               Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
'    pFile.WriteLine ("c")
'    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
'    pFile.WriteLine ("c BMPSITE  QUALPCTREM1  QUALPCTREM2 ... QUALPCTREMN")
'    pFile.Write StrCard1330PercentRemoval
'    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
'End Sub

'******************************************************************************
'Subroutine: WriteCard770
'Author:     Mira Chokshi
'Purpose:    Write BMP Underdrain Pollutant Percent Removal (applies when underdrain is on in card 740)
'******************************************************************************
Private Sub WriteCard770()
On Error GoTo ShowError
    pFile.WriteLine ("c770 BMP Underdrain Pollutant Percent Removal (applies when underdrain is on in card 740)")
    pFile.WriteLine ("c BMPSITE     = BMPSITE identifier in card 715")
    pFile.WriteLine ("c QUALPCTREMi = Perecent Removal for pollutant i through underdrain (0-1)")
    pFile.WriteLine ("c               Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  QUALPCTREM1  QUALPCTREM2 ... QUALPCTREMN")
    pFile.Write StrCard770PercentRemoval
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub
'''******************************************************************************
'''Subroutine: WriteCard1400
'''Author:     Mira Chokshi
'''Purpose:    Write Adjustable BMP parameters (applies when parameter is selected for optimization)
'''******************************************************************************
''Private Sub WriteCard1400()
''    pFile.WriteLine ("c1400 BMP SITE Adjustable Parameters")
''    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
''    pFile.WriteLine ("c VARIABLE   = Variable name")
''    pFile.WriteLine ("c FROM       = From value in the range")
''    pFile.WriteLine ("c TO         = To value in the range")
''    pFile.WriteLine ("c STEP       = Increment step")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE  VARIABLE  FROM   TO   STEP")
''    pFile.Write StrCard1400AdjustParameter
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard810
'Author:     Mira Chokshi
'Purpose:    Write Adjustable BMP parameters (applies when parameter is selected for optimization)
'******************************************************************************
Private Sub WriteCard810()
  On Error GoTo ErrorHandler

    pFile.WriteLine ("c810 BMP SITE Adjustable Parameters")
    pFile.WriteLine ("c BMPSITE    = BMP site identifier in card 715")
    pFile.WriteLine ("c VARIABLE   = Variable name")
    pFile.WriteLine ("c FROM       = From value in the range")
    pFile.WriteLine ("c TO         = To value in the range")
    pFile.WriteLine ("c STEP       = Increment step")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE  VARIABLE  FROM   TO   STEP")
    pFile.Write StrCard810AdjustParameter
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")

  Exit Sub
ErrorHandler:
    gHasInFileError = True
End Sub
'''******************************************************************************
'''Subroutine: WriteCard1410
'''Author:     Mira Chokshi
'''Purpose:    BMP Cost Functions
'''Modified:   Sabu Paul added this new card. The BMP cost is a function of Depth,
'''            Area and other constants.
'''******************************************************************************
''Private Sub WriteCard1410()
''    pFile.WriteLine ("c1410 BMP Cost Functions")
''    pFile.WriteLine ("c Cost ($) = (Aa) Area^(Ab)x(Da)Depth^(Db) + (LdCost)Area + (ConstCost)")
''    pFile.WriteLine ("c Depth = WeirHeight + SoilDepth + UnderdrainDepth")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c BMPSITE = BMP site identifier in card 715")
''    pFile.WriteLine ("c")
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''    pFile.WriteLine ("c BMPSITE       Aa      Ab      Da      Db      LdCost  ConstCost")
''    pFile.Write StrCard1410BMPCost
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'******************************************************************************
'Subroutine: WriteCard805
'Author:     Mira Chokshi
'Purpose:    BMP Cost Functions
'Modified:   Sabu Paul added this new card. The BMP cost is a function of Depth,
'            Area and other constants.
'******************************************************************************
Private Sub WriteCard805()
  On Error GoTo ErrorHandler

'    pFile.WriteLine ("c805 BMP Cost Functions")
'    pFile.WriteLine ("c Cost ($) = (Aa) Area^(Ab)x(Da)Depth^(Db) + (LdCost)Area + (ConstCost)")
'    pFile.WriteLine ("c Depth = WeirHeight + SoilDepth + UnderdrainDepth")
'    pFile.WriteLine ("c")
'    pFile.WriteLine ("c BMPSITE = BMP site identifier in card 715")
'    pFile.WriteLine ("c")
'    pFile.WriteLine ("c BMPSITE       Aa      Ab      Da      Db      LdCost  ConstCost")
'    pFile.Write StrCard805BMPCost
'    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c805 BMP Cost Functions")
    'pFile.WriteLine ("c Cost ($) = ((LinearCost) Length + (AreaCost)Area + (TotalVolumeCost)TotalVolume + (MediaVolumeCost)SoilMediaVolume + (UnderDrainVolumeCost)UnderDrainVolume+ (ConstantCost)) * (1+PercentCost/100)")
    pFile.WriteLine ("c Cost ($) = ((LinearCost)Length^(LengthExp)  + (AreaCost)Area^(AreaExp)  + (TotalVolumeCost)TotalVolume^(TotalVolExp) + (MediaVolumeCost)SoilMediaVolume^(MediaVolExp) + (UnderDrainVolumeCost)UnderDrainVolume^(UDVolExp) + (Unitcost) +  (ConstantCost)) * (1+PercentCost/100)")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE                 = BMP site identifier in card 715")
    pFile.WriteLine ("c LinearCost              = Cost per unit length of the BMP structure ($/ft)")
    pFile.WriteLine ("c AreaCost                = Cost per unit area of the BMP structure ($/ft^2)")
    pFile.WriteLine ("c TotalVolumeCost         = Cost per unit total volume of the BMP structure ($/ft^3)")
    pFile.WriteLine ("c MediaVolumeCost         = Cost per unit volume of the soil media ($/ft^3)")
    pFile.WriteLine ("c UnderDrainVolumeCost    = Cost per unit volume of the under drain structure ($/ft^3)")
    'pFile.WriteLine ("c UnitCost                = Cost per unit of the functional component ($/num)")
    pFile.WriteLine ("c ConstantCost            = Constant cost ($)")
    pFile.WriteLine ("c PercentCost             = Cost in percentage of all other cost (%)")
    pFile.WriteLine ("c LengthExp               = Exponent for linear unit")
    pFile.WriteLine ("c AreaExp                 = Exponent for area unit")
    pFile.WriteLine ("c TotalVolExp             = Exponent for total volume unit")
    pFile.WriteLine ("c MediaVolExp             = Exponent for soil media volume unit")
    pFile.WriteLine ("c UDVolExp                = Exponent for underdrain volume unit")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE       LinearCost      AreaCost      TotalVolumeCost MediaVolumeCost UnderDrainVolumeCost    ConstantCost  PercentCost   LengthExp   AreaExp     TotalVolExp   MediaVolExp    UDVolExp")
    pFile.Write StrCard805BMPCost
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")

  Exit Sub
ErrorHandler:
    gHasInFileError = True
  'HandleError False, "WriteCard805 " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub

'''******************************************************************************
'''Subroutine: WriteCard1420
'''Author:     Mira Chokshi
'''Purpose:    Assessment point details
'''Modified:   Sabu Paul added new card details to enter additional information
'''            for BMPs added as assessment points.
'''******************************************************************************
''Private Sub WriteCard1420()
''    GetAssessPointOptimizationDetails
''    pFile.WriteLine ("c1420 Assessment Point and Evaluation Factor")
''    pFile.WriteLine ("c Option -- Optimization options")
''    pFile.WriteLine ("c     0 = no optimization")
''    pFile.WriteLine ("c     1 = specific control target and minimize cost")
''    pFile.WriteLine ("c     2 = cost limit and maximize control")
''    pFile.WriteLine ("c     3 = generate trade-off curve")
''    pFile.WriteLine ("c StopDelta -- Criteria for stopping the optimization iteration")
''    pFile.WriteLine ("c              in $ for option 1, meaning if the cost not improved by this criteria, stop the search")
''    pFile.WriteLine ("c              in % for option 2, meaning if the Evaluation Factor not improved by this criteria, stop the search")
''    pFile.WriteLine ("c MaxRunTime(hr) -- Maximum search time (in Hour) allowed")
''    pFile.WriteLine ("c NumBest -- Number of best solutions for output")
''    pFile.WriteLine ("c BMPSITE -- BMP site identifier in card 715 if it is an assessment point")
''    pFile.WriteLine ("c FactorGroup -- -1 for flow related, positive number for pollutant column order in card 765 and 770")
''    pFile.WriteLine ("c FactorType -- Evaluation Factor Type (negative number for flow related and positive number for pollutant related)")
''    pFile.WriteLine ("c    -1 = AAFV Annual Average Flow Volume (ft3/yr)")
''    pFile.WriteLine ("c    -2 = PDF  Peak Discharge Flow (cfs)")
''    pFile.WriteLine ("c    -3 = FEF  Flow Exceeding frequency (cfs)")
''    pFile.WriteLine ("c     1 = AAL  Annual Average Load (kg/yr)")
''    pFile.WriteLine ("c     2 = AAC  Annual Average Concentration (mg/L)")
''    pFile.WriteLine ("c     3 = MAC  Maximum #days Average Concentraion (mg/L)")
''    pFile.WriteLine ("c CalcDays --  FactorType 3 (MAC): Maxmimum #Days;")
''    pFile.WriteLine ("c       FactorType -3 (FEF): Threshold (cfs)")
''    pFile.WriteLine ("c CalcMode -- Evaluation Factor Calculation Mode")
''    pFile.WriteLine ("c   -99 =        Option 0: no optimizaiton, only calculate EF")
''    pFile.WriteLine ("c     0 =        Option 2: Maximize Control")
''    pFile.WriteLine ("c     1 = %      percent of value under existing condition (0-100)")
''    pFile.WriteLine ("c     2 = S      scale between pre-develop and existing condition (0-1)")
''    pFile.WriteLine ("c     3 = Value  absolute value in the unit as shown in Factor_ID block")
''    pFile.WriteLine ("c Target_Value -- Target value for Option 1 and Priority Factor (0-10) for Option 2")
''    pFile.WriteLine ("c TargetVal_Low and TargetVal_Up -- Target Value Lower and Upper Limit for Option 3")
''    pFile.WriteLine ("c Factor_Name = Evaluation factor name, e.g. FlowVolume or SEDIMENT")
''    pFile.WriteLine ("c---------------------------------------------------------------------------------------------")
''    pFile.Write StrCard1420Assess
''    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
''End Sub

'New card for optimization controls
'******************************************************************************
'Subroutine: WriteCard800
'Author:     Mira Chokshi
'Purpose:    Assessment point details
'Modified:   Sabu Paul added new card details to enter additional information
'            for BMPs added as assessment points.
'******************************************************************************
Private Sub WriteCard800()
On Error GoTo ShowError
    GetAssessPointOptimizationDetails
    
    pFile.WriteLine ("c800 Optimization Controls")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c Technique -- Optimization Techniques")
    pFile.WriteLine ("c     0 = no optimization")
    pFile.WriteLine ("c     1 = Scatter Search")
    pFile.WriteLine ("c     2 = NSGAII")
    pFile.WriteLine ("c Option -- Optimization options")
    pFile.WriteLine ("c     0 = no optimization")
    pFile.WriteLine ("c     1 = specific control target and minimize cost")
    'pFile.WriteLine ("c     3 = cost limit and maximize control")
    pFile.WriteLine ("c     2 = generate cost effectiveness curve")
    pFile.WriteLine ("c StopDelta -- Criteria for stopping the optimization iteration")
    pFile.WriteLine ("c              in $ for option 1, meaning if the cost not improved by this criteria, stop the search")
    pFile.WriteLine ("c              in % for option 2, meaning if the Evaluation Factor not improved by this criteria, stop the search")
    pFile.WriteLine ("c MaxRunTime(hr) -- Maximum search time (in Hour) allowed")
    pFile.WriteLine ("c NumBest -- Number of best solutions for output")
    'pFile.WriteLine ("c NumBreak – Number of break points between the lower and upper target values of trade-off curve.")
    pFile.WriteLine ("c")
    pFile.Write StrCard800OptimizationControls
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
End Sub

'******************************************************************************
'Subroutine: WriteCard815
'Author:     Mira Chokshi
'Purpose:    Assessment point details
'Modified:   Sabu Paul added new card details to enter additional information
'            for BMPs added as assessment points.
'******************************************************************************
Private Sub WriteCard815()
  On Error GoTo ErrorHandler

    GetAssessPointOptimizationDetails
    pFile.WriteLine ("c815 Assessment Point and Evaluation Factor")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c BMPSITE -- BMP site identifier in card 715 if it is an assessment point")
    pFile.WriteLine ("c FactorGroup -- -1 for flow related, positive number for pollutant column order in card 765 and 770")
    pFile.WriteLine ("c FactorType -- Evaluation Factor Type (negative number for flow related and positive number for pollutant related)")
    pFile.WriteLine ("c    -1 = AAFV Annual Average Flow Volume (ft3/yr)")
    pFile.WriteLine ("c    -2 = PDF  Peak Discharge Flow (cfs)")
    pFile.WriteLine ("c    -3 = FEF  Flow Exceeding frequency (cfs)")
    pFile.WriteLine ("c     1 = AAL  Annual Average Load (kg/yr)")
    pFile.WriteLine ("c     2 = AAC  Annual Average Concentration (mg/L)")
    pFile.WriteLine ("c     3 = MAC  Maximum #days Average Concentraion (mg/L)")
    pFile.WriteLine ("c CalcDays --  FactorType 3 (MAC): Maxmimum #Days;")
    pFile.WriteLine ("c       FactorType -3 (FEF): Threshold (cfs)")
    pFile.WriteLine ("c CalcMode -- Evaluation Factor Calculation Mode")
    pFile.WriteLine ("c   -99 =        Option 0: no optimizaiton, only calculate EF")
    pFile.WriteLine ("c     0 =        Option 2: Maximize Control")
    pFile.WriteLine ("c     1 = %      percent of value under existing condition (0-100)")
    pFile.WriteLine ("c     2 = S      scale between pre-develop and existing condition (0-1)")
    pFile.WriteLine ("c     3 = Value  absolute value in the unit as shown in Factor_ID block")
    pFile.WriteLine ("c Target_Value -- Target value for Option 1 and Priority Factor (0-10) for Option 2")
    pFile.WriteLine ("c TargetVal_Low and TargetVal_Up -- Target Value Lower and Upper Limit for Option 3")
    pFile.WriteLine ("c Factor_Name = Evaluation factor name, e.g. FlowVolume or SEDIMENT")
    pFile.WriteLine ("c")
    pFile.Write StrCard815Assess
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")

  Exit Sub
ErrorHandler:
    gHasInFileError = True
  'HandleError False, "WriteCard815 " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub
'*******************************************************************************
'Subroutine : GetPollutantDetails
'Purpose    : Read Pollutants table and print the list of pollutants
'Note       :
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetPollutantDetails() As String
On Error GoTo ShowError
    Dim result As String
    result = ""
    
    Dim pTable As iTable
    Set pTable = GetInputDataTable("Pollutants")
    If (pTable Is Nothing) Then
        GetPollutantDetails = result
        MsgBox "Missing pollutants table: Define pollutants first"
        Exit Function
    End If
    
    Dim iNameFld As Integer
    iNameFld = pTable.FindField("Name")
    
    Dim iMultiplierFld As Integer
    iMultiplierFld = pTable.FindField("Multiplier")
    
    Dim iSedFlagFld As Integer
    iSedFlagFld = pTable.FindField("Sediment")
    
    Dim iSedAssoc As Integer
    iSedAssoc = pTable.FindField("SedAssoc")
    Dim iSandFrac As Integer
    iSandFrac = pTable.FindField("SandFrac")
    Dim iSiltFrac As Integer
    iSiltFrac = pTable.FindField("SiltFrac")
    Dim iClayFrac As Integer
    iClayFrac = pTable.FindField("ClayFrac")
       
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID > 0 Order by ID"
    
    Set pCursor = pTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow
    
    Dim sedFlag As Integer
    Dim sedAssoc As Integer
    Dim sandFrac As Double
    Dim siltFrac As Double
    Dim clayFrac As Double
        
    If gInternalSimulation Then
        sedAssoc = 0
        sandFrac = 0
        siltFrac = 0
        clayFrac = 0
    End If
        
    Dim rCount As Integer
    rCount = 1
    Do Until pRow Is Nothing
        Select Case UCase(pRow.value(iSedFlagFld))
            Case "SEDIMENT":
                sedFlag = 4
            Case "SAND":
                sedFlag = 1
            Case "SILT":
                sedFlag = 2
            Case "CLAY":
                sedFlag = 3
            Case Else
                sedFlag = 0
        End Select
        If gExternalSimulation Then
            sedAssoc = pRow.value(iSedAssoc)
            If sedAssoc = 1 Then
                sandFrac = pRow.value(iSandFrac)
                siltFrac = pRow.value(iSiltFrac)
                clayFrac = pRow.value(iClayFrac)
            Else
                sandFrac = 0
                siltFrac = 0
                clayFrac = 0
            End If
        End If
        'result = result & vbTab & rCount & vbTab & pRow.value(iNameFld) & vbTab & pRow.value(iMultiplierFld) & vbTab & pRow.value(iSedFlagFld) & vbNewLine
        result = result & rCount & vbTab & pRow.value(iNameFld) & vbTab & pRow.value(iMultiplierFld) & vbTab & sedFlag & vbTab & sedAssoc & vbTab & FormatNumber(sandFrac, 3) & vbTab & FormatNumber(siltFrac, 3) & vbTab & FormatNumber(clayFrac, 3) & vbNewLine
        Set pRow = pCursor.NextRow
        rCount = rCount + 1
    Loop
    
    GetPollutantDetails = result
    GoTo CleanUp
ShowError:
    MsgBox "Error in GetPollutantDetails:" & Err.description
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pQueryFilter = Nothing
End Function
'******************************************************************************
'Subroutine: GetLandTypeDetails
'Author:     Mira Chokshi
'Purpose:    Read the lureclass table, gets all parameters for landuse types.
'*****************************************************************************
Private Function GetLandTypeDetails() As String
On Error GoTo ShowError

    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim pTable As iTable
    Dim pTableName As String
    
    If gExternalSimulation Then
        pTableName = "TSAssigns"
    Else
       pTableName = "LUReclass"
    End If
    
    Set pTable = GetInputDataTable(pTableName)
    If (pTable Is Nothing) Then
        MsgBox pTableName & " table not found."
        Exit Function
    End If
        
    Dim iLuGroupID As Long
    iLuGroupID = pTable.FindField("LUGroupID")
    Dim iLuGroup As Long
    iLuGroup = pTable.FindField("LUGroup")
    Dim iLUType As Long
    iLUType = pTable.FindField("Impervious")
    Dim iTimeSeries As Long
    If gExternalSimulation Then iTimeSeries = pTable.FindField("TimeSeries")
    
    Dim iSand As Long
    iSand = pTable.FindField("SandFrac")
    Dim iSilt As Long
    iSilt = pTable.FindField("SiltFrac")
    Dim iClay As Long
    iClay = pTable.FindField("ClayFrac")
    
    Dim pTimeSeriesFile As String
    pMaxLandTypeGroupID = 0
    Dim pReturnString As String
    Dim pImpervious As Integer
    Dim LUgroupCount As Integer
    LUgroupCount = 0
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "LUGroupID > 0"
    
    Dim pTableSort As ITableSort
    Set pTableSort = New esriGeoDatabase.TableSort
    
    With pTableSort
      .Fields = "LUGroupID, LuGroup"
      .Ascending("LUGroupID") = True
      .Ascending("LuGroup ") = True
      Set .QueryFilter = pQueryFilter
      Set .Table = pTable
    End With

    pTableSort.Sort Nothing
           
    Dim pCursor As ICursor
    'Set pCursor = pTable.Search(Nothing, True)
    Dim pRow As iRow
    'Set pRow = pCursor.NextRow
    
    Set pCursor = pTableSort.Rows
    Set pRow = pCursor.NextRow
    
    Do While Not pRow Is Nothing
        If (pMaxLandTypeGroupID <> pRow.value(iLuGroupID)) Then
            pMaxLandTypeGroupID = pRow.value(iLuGroupID)
            LUgroupCount = LUgroupCount + 1
            pTimeSeriesFile = pRow.value(iTimeSeries)
            fso.CopyFile pTimeSeriesFile, pInputFolder
            If gExternalSimulation Then
                pTimeSeriesFile = fso.GetFileName(pTimeSeriesFile)
            Else
                pTimeSeriesFile = "Internal Simulation"
            End If
            'Modified to add sand, silt and clay fraction - June 14, 2007
'            pReturnString = pReturnString & _
'                            pRow.value(iLuGroupID) & vbTab & _
'                            pRow.value(iLuGroup) & vbTab & _
'                            pRow.value(iLUType) & vbTab & _
'                            pTimeSeriesFile & vbNewLine
            'SAND_FRAC       SILT_FRAC       CLAY_FRAC set to 0.33 each
                        
            pReturnString = pReturnString & _
                            pRow.value(iLuGroupID) & vbTab & _
                            Replace(pRow.value(iLuGroup), " ", "_") & vbTab & _
                            pRow.value(iLUType) & vbTab & _
                            pTimeSeriesFile & vbTab & _
                            pRow.value(iSand) & vbTab & _
                            pRow.value(iSilt) & vbTab & _
                            pRow.value(iClay) & vbNewLine
        End If
        Set pRow = pCursor.NextRow
    Loop
    
    '*** Get details of external TS
''    Set pTable = GetInputDataTable("ExternalTS")
''    If Not (pTable Is Nothing) Then
''        Set pCursor = pTable.Search(Nothing, True)
''        Set pRow = pCursor.NextRow
''        Dim iLUDesc As Long
''        iLUDesc = pCursor.FindField("LUDescrip")
''        Dim iExtTimeSeries As Long
''        iExtTimeSeries = pCursor.FindField("TimeSeries")
''        Do While Not pRow Is Nothing
''            LUgroupCount = LUgroupCount + 1
''            pTimeSeriesFile = pRow.value(iExtTimeSeries)
''            fso.CopyFile pTimeSeriesFile, pInputFolder
''            pTimeSeriesFile = fso.GetFileName(pTimeSeriesFile)
''            'Modified to add sand, silt and clay fraction - June 14, 2007
''''            pReturnString = pReturnString & _
''''                            LUgroupCount & vbTab & _
''''                            pRow.value(iLUDesc) & vbTab & _
''''                            "1" & vbTab & _
''''                            pTimeSeriesFile & vbNewLine
''            pReturnString = pReturnString & _
''                            LUgroupCount & vbTab & _
''                            pRow.value(iLUDesc) & vbTab & _
''                            "1" & vbTab & _
''                            pTimeSeriesFile & vbTab & _
''                            "0.33" & vbTab & _
''                            "0.33" & vbTab & _
''                            "0.33" & vbNewLine
''
''            Set pRow = pCursor.NextRow
''        Loop
''    End If
    
    'Return the LandType Information
    GetLandTypeDetails = pReturnString
    GoTo CleanUp
ShowError:
    MsgBox "GetLandTypeDetails: " & Err.description
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set fso = Nothing
End Function

Private Function GetPointSourceDetails() As String
On Error GoTo ShowError
    Dim pReturnString As String
    pReturnString = ""
    
    Dim pTable As iTable
    Dim pCursor As ICursor
    Dim pRow As iRow
    
    Set pTable = GetInputDataTable("ExternalTS")
    If pTable Is Nothing Then
        GetPointSourceDetails = pReturnString
        Exit Function
    Else
        Set pCursor = pTable.Search(Nothing, True)
        
        Dim iBMPIndex As Long
        iBMPIndex = pTable.FindField("BMPID")
        Dim iLUDesc As Long
        iLUDesc = pTable.FindField("LUDescrip")
        Dim iExtTimeSeries As Long
        iExtTimeSeries = pTable.FindField("TimeSeries")
        'Added three new fields - June 18, 2007
        Dim iSandFracIndex As Long
        iSandFracIndex = pTable.FindField("SandFrac")
        Dim iSiltFracIndex As Long
        iSiltFracIndex = pTable.FindField("SiltFrac")
        Dim iClayFracIndex As Long
        iClayFracIndex = pTable.FindField("ClayFrac")
        
        Dim pRowCount As Integer
        Dim pBMPID As String
        Dim pMultiplier As Double
        Dim pTimeSeriesFile As String
        Dim pSandFrac As Double
        Dim pSiltFrac As Double
        Dim pClayFrac As Double
        
        pRowCount = 1
        
        Set pRow = pCursor.NextRow
        Do Until pRow Is Nothing
            pBMPID = pRow.value(iBMPIndex)
            pTimeSeriesFile = pRow.value(iExtTimeSeries)
            fso.CopyFile pTimeSeriesFile, pInputFolder
            pTimeSeriesFile = fso.GetFileName(pTimeSeriesFile)
            'Modified to add sand, silt and clay fraction - June 14, 2007
            pSandFrac = pRow.value(iSandFracIndex)
            pSiltFrac = pRow.value(iSiltFracIndex)
            pClayFrac = pRow.value(iClayFracIndex)
            
''            pReturnString = pReturnString & _
''                            LUgroupCount & vbTab & _
''                            pRow.value(iLUDesc) & vbTab & _
''                            "1" & vbTab & _
''                            pTimeSeriesFile & vbNewLine
                            
            pReturnString = pReturnString & _
                            pRowCount & vbTab & _
                            pBMPID & vbTab & _
                            pMultiplier & vbTab & _
                            pTimeSeriesFile & vbTab & _
                            pSandFrac & vbTab & _
                            pSiltFrac & vbTab & _
                            pClayFrac & vbNewLine
                            
            Set pRow = pCursor.NextRow
            pRowCount = pRowCount + 1
        Loop
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "Error in GetPointSourceDetails :" & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
End Function

'******************************************************************************
'Subroutine: CustomSplit
'Author:     Sabu Paul
'Purpose:    Removes all spaces in a string.
'*****************************************************************************
Public Function CustomSplit(pString As String) As String()
  On Error GoTo ErrorHandler
    Dim res() As String
    If Trim(pString) = "" Then
        CustomSplit = Split(Trim(pString))
        Exit Function
    End If
    pString = Replace(pString, " ", vbTab)
    Dim tmpRes
    'tmpRes = Split(pString, " ", -1, vbTextCompare)
    tmpRes = Split(pString, vbTab, , vbTextCompare)
    Dim resWords As Long
    resWords = 0
    Dim incr As Integer
    For incr = 0 To UBound(tmpRes)
        If Trim(tmpRes(incr)) <> "" Then
            ReDim Preserve res(resWords)
            res(resWords) = tmpRes(incr)
            resWords = resWords + 1
        End If
    Next incr
    CustomSplit = res

  Exit Function
ErrorHandler:
  'HandleError True, "CustomSplit " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Function


'******************************************************************************
'Subroutine: BMPCardsDetail
'Author:     Mira Chokshi
'Purpose:    Get the dimension properties of BMPs. Also get the soil properties,
'            cost values, assessment point evaluation factors, growth index.
'*****************************************************************************
Public Sub BMPCardsDetail(bmpId As Integer)
On Error GoTo ShowError
    'MsgBox " BMPCardsDetail 1 for " & bmpId
    Dim pReturnString As String
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("BMPs")
    If (pFeatureLayer Is Nothing) Then
        MsgBox "BMPs featurelayer not found."
        Exit Sub
    End If
    Dim pTotalBMPs As Integer
    pTotalBMPs = pFeatureLayer.FeatureClass.FeatureCount(Nothing)
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPDetail")
    If (pTable Is Nothing) Then
        MsgBox "BMPDetail table not found."
        Exit Sub
    End If
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iPropName As Long
    iPropName = pTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    
    Dim iBMP As Integer
    Dim strBmpFeatType As String
    Dim pBMPType As String
    Dim pBMPClass As String
    Dim isAssessPoint As Boolean
    
    'Modified so that bmp ids are read from BMPs feature table
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pFeature As IFeature
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeatIDIndex As Long
    pFeatIDIndex = pFeatureclass.FindField("ID")
    Dim pFeatTypeIndex As Long
    pFeatTypeIndex = pFeatureclass.FindField("Type")

    'Define watershed feature layer
    Dim pTotalDrainageArea As Double
        
    'set whereclause of the bmp feature layer
    pQueryFilter.WhereClause = "ID = " & bmpId
    Set pFeatureCursor = pFeatureLayer.Search(pQueryFilter, False)
    Set pFeature = pFeatureCursor.NextFeature
    Dim pPoint As IPoint
    Dim pXpos As Double
    Dim pYpos As Double
    
    Dim pAgg_BMP_Cat_Dict As Scripting.Dictionary
    
    Dim bmpCardStr As String
    While Not pFeature Is Nothing
        'MsgBox " BMPCardsDetail 2 for " & bmpId
        iBMP = pFeature.value(pFeatIDIndex)
        strBmpFeatType = pFeature.value(pFeatTypeIndex)
        'Set the point to find X and Y coordinates
        Set pPoint = pFeature.Shape
        pPoint.QueryCoords pXpos, pYpos
        
        'Format pXpos, pYpos
        pXpos = FormatNumber(pXpos, 1)
        pYpos = FormatNumber(pYpos, 1)
        
        If UCase(strBmpFeatType) = "AGGREGATE" Then
            'Set pAgg_BMP_Cat_Dict = Get_OnMap_BMP_Categories(iBMP)
            Set pAgg_BMP_Cat_Dict = GetAggBMPTypes(iBMP)
            
            'Get individual BMP properties
            Dim catIndex As Integer
            If Not pAgg_BMP_Cat_Dict Is Nothing Then
                For catIndex = 0 To pAgg_BMP_Cat_Dict.Count - 1
                    If pAgg_BMP_Cat_Dict.Exists(pAgg_BMP_Cat_Dict.keys(catIndex)) Then
                        bmpCardStr = iBMP & "_" & pAgg_BMP_Cat_Dict.keys(catIndex)
                        pBmpDetailDict.RemoveAll
                        If gBMPDrainAreaDict.Exists(bmpCardStr) Then
                            pTotalDrainageArea = gBMPDrainAreaDict.Item(bmpCardStr)
                        Else
                            pTotalDrainageArea = 0
                        End If
                        Set pBmpDetailDict = GetBMPDetailDict(CInt(pAgg_BMP_Cat_Dict.keys(catIndex)), "AgBMPDetail")
                        Call GetBMPDetails_Classes(iBMP, pBMPType, pXpos, pYpos, pTotalDrainageArea, bmpCardStr)
                    End If
                Next
            End If

            'bmpCardStr = iBMP & "_E"
            bmpCardStr = iBMP '& "_0"
            pBmpDetailDict.RemoveAll
            pBmpDetailDict.Item("BMPType") = "Junction"
            pBmpDetailDict.Item("BMPClass") = "X"
            If gBMPDrainAreaDict.Exists(bmpCardStr) Then
                pTotalDrainageArea = gBMPDrainAreaDict.Item(bmpCardStr)
            Else
                pTotalDrainageArea = 0
            End If
            GetDummyBMPParameters bmpCardStr, pXpos, pYpos, pTotalDrainageArea
        
        Else
            pQueryFilter.WhereClause = "ID = " & iBMP
            Set pCursor = pTable.Search(pQueryFilter, True)
            Set pRow = pCursor.NextRow
            'Remove all the parameters present in the dictionary
            pBmpDetailDict.RemoveAll
            
            Do While Not pRow Is Nothing
                pBmpDetailDict.Item(pRow.value(iPropName)) = pRow.value(iPropValue)
                Set pRow = pCursor.NextRow
            Loop
            pBMPType = Trim(pBmpDetailDict.Item("BMPType"))
            pBMPClass = Trim(pBmpDetailDict.Item("BMPClass"))
            If pBMPType = "" Then
                pBMPType = pFeature.value(pFeatureclass.FindField("TYPE"))
            End If
            
            If gBMPDrainAreaDict.Exists(iBMP) Then
                pTotalDrainageArea = gBMPDrainAreaDict.Item(iBMP)
            Else
                pTotalDrainageArea = 0
            End If
'            If pBMPClass = "X" Then
'                GetDummyBMPParameters CStr(iBMP), pXpos, pYpos, pTotalDrainageArea
'            Else
'                Call GetBMPDetails_Classes(iBMP, pBMPType, pXpos, pYpos, pTotalDrainageArea)
'            End If
            Call GetBMPDetails_Classes(iBMP, pBMPType, pXpos, pYpos, pTotalDrainageArea)
        End If
        Set pFeature = pFeatureCursor.NextFeature
    Wend
    GoTo CleanUp
ShowError:
    MsgBox "BMPCardsDetail: " & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pFeatureclass = Nothing
    Set pFeature = Nothing
    Set pFeatureCursor = Nothing
    Set pPoint = Nothing
'    Set pWatershedLayer = Nothing
'    Set pWatershedClass = Nothing
'    Set pWatershedCursor = Nothing
'    Set pWatershedFeature = Nothing
End Sub


'******************************************************************************
'Subroutine: ConduitCardsDetails
'Author:     Mira Chokshi   02/08/2005
'Purpose:    Get the dimension and cross section information for conduits
'*****************************************************************************
Private Sub ConduitCardsDetails(bmpId As Integer)
On Error GoTo ShowError
    'Get the Conduits feature layer
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Conduits")
    If (pFeatureLayer Is Nothing) Then
        MsgBox "Conduits feature layer not found."
        Exit Sub
    End If
    'Get the BMPDetail table
    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPDetail")
    If (pTable Is Nothing) Then
        MsgBox "BMPDetail table not found."
        Exit Sub
    End If
    'Define the query, cursor, row for conduits detail table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iPropName As Long
    iPropName = pTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    'Define the feature class, feature cursor, feature for conduit feature layer
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim iIDFld As Long
    Dim pConduitID As Integer
    iIDFld = pFeatureclass.FindField("ID")
    Dim iFROMFld As Long
    Dim iTOFld As Long
    iFROMFld = pFeatureclass.FindField("CFROM")
    iTOFld = pFeatureclass.FindField("CTO")
    pQueryFilter.WhereClause = "ID = " & bmpId
    Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    'Get the ID of each conduit feature, get details for each feature from conduit details table
    Do While Not pFeature Is Nothing
        'Get the Conduit ID value
        pConduitID = pFeature.value(iIDFld)
        pQueryFilter.WhereClause = "ID = " & pConduitID
        Set pCursor = pTable.Search(pQueryFilter, True)
        Set pRow = pCursor.NextRow
        'Remove all parameters from bmpdetaildict
        pBmpDetailDict.RemoveAll
        pBmpDetailDict.add "INLETNODE", pFeature.value(iFROMFld)
        pBmpDetailDict.add "OUTLETNODE", pFeature.value(iTOFld)
        'Get each property name
        Do While Not pRow Is Nothing
            pBmpDetailDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow  'Next row
        Loop
        'Get Decay Factors for conduits
        GetBMPDecayFactors CStr(pConduitID)
        
        GetBMP_K_C_factors CStr(pConduitID)
        'Get Class C - Conduits parameters
        GetConduitParameters pConduitID
        'Iterate to next feature
        Set pFeature = pFeatureCursor.NextFeature
    Loop

    GoTo CleanUp
ShowError:
    MsgBox "ConduitCardsDetail: " & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pFeatureclass = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
End Sub



'******************************************************************************
'Subroutine: VFSCardsDetail
'Author:     Mira Chokshi   02/08/2005
'Purpose:    Get the dimension and cross section information for conduits
'*****************************************************************************
Private Sub VFSCardsDetail(bmpId As Integer)
On Error GoTo ShowError
    'Get the Conduits feature layer
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("VFS")
    If (pFeatureLayer Is Nothing) Then
        Exit Sub
    End If
    'Get the BMPDetail table
    Dim pTable As iTable
    Set pTable = GetInputDataTable("VFSDetail")
    If (pTable Is Nothing) Then
        MsgBox "VFSDetail table not found."
        Exit Sub
    End If
    'Define the query, cursor, row for conduits detail table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim iPropName As Long
    iPropName = pTable.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pTable.FindField("PropValue")
    'Define the feature class, feature cursor, feature for conduit feature layer
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim iIDFld As Long
    iIDFld = pFeatureclass.FindField("ID")
    Dim pVFSID As Integer
    pQueryFilter.WhereClause = "ID = " & bmpId
    Set pFeatureCursor = pFeatureclass.Search(pQueryFilter, True)
    Set pFeature = pFeatureCursor.NextFeature
    'Get the ID of each conduit feature, get details for each feature from conduit details table
    Do While Not pFeature Is Nothing
        'Get the Conduit ID value
        pVFSID = pFeature.value(iIDFld)
        pQueryFilter.WhereClause = "ID = " & pVFSID
        Set pCursor = pTable.Search(pQueryFilter, pVFSID)
        Set pRow = pCursor.NextRow
        'Remove all parameters from bmpdetaildict
        pBmpDetailDict.RemoveAll
        'Get each property name
        Do While Not pRow Is Nothing
            pBmpDetailDict.add pRow.value(iPropName), pRow.value(iPropValue)
            Set pRow = pCursor.NextRow  'Next row
        Loop
        'Get Class D - Buffer Strip parameters
        GetVFSParameters pVFSID
        'Iterate to next feature
        Set pFeature = pFeatureCursor.NextFeature
    Loop

    GoTo CleanUp
ShowError:
    MsgBox "VFSCardsDetail: " & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pFeatureclass = Nothing
    Set pFeatureCursor = Nothing
    Set pFeature = Nothing
End Sub

'******************************************************************************
'Subroutine: GetBMPParametersA
'Author:     Mira Chokshi
'Purpose:    Get the dimention details of BMP Type A. BMP Type A could be a
'            rain barrel or water cistern. In case of cistern, also get the
'            24 hour, hourly water release information.
'*****************************************************************************
'Private Sub GetBMPParametersA(iBMPindex As Integer, pXpos As Double, pYpos As Double, pDrainageArea As Double)
Private Sub GetBMPParametersA(strBMPindex As String, pXpos As Double, pYpos As Double, pDrainageArea As Double)
On Error GoTo ShowError
    Dim BMPName As String
    Dim BMPType As String
    Dim bmpClass As String
    Dim basinlength As Double
    Dim basinwidth As Double
    Dim basinsurf_area As Double
    Dim orificediam As Double
    Dim orificeheight As Double
    Dim orifice_area As Double
    Dim orificecoef As Double
    Dim weirheight As Double
    Dim weirwidth As Double
    Dim exittype As Integer
    Dim WeirType As Integer
    Dim weirangle As Double
    Dim releaseoption As String
    Dim releasetype As Integer
    Dim NumDays As Integer
    Dim NumPeople As Integer
    Dim strCisternFlow As String
    Dim pCisternFlow
    Dim NumUnits As Integer
    Dim ddArea As Double
        
    BMPName = pBmpDetailDict.Item("BMPName")
    BMPType = pBmpDetailDict.Item("BMPType")
    bmpClass = pBmpDetailDict.Item("BMPClass")
    basinlength = CDbl(pBmpDetailDict.Item("BMPLength"))
    basinwidth = CDbl(pBmpDetailDict.Item("BMPWidth"))
    orificeheight = CDbl(pBmpDetailDict.Item("BMPOrificeHeight"))
    orificediam = CDbl(pBmpDetailDict.Item("BMPOrificeDiameter"))
    weirheight = CDbl(pBmpDetailDict.Item("BMPWeirHeight"))
    weirwidth = CDbl(pBmpDetailDict.Item("BMPRectWeirWidth"))
    exittype = CInt(pBmpDetailDict.Item("ExitType"))
    WeirType = CInt(pBmpDetailDict.Item("WeirType"))
    releaseoption = CStr(pBmpDetailDict.Item("ReleaseOption"))
    NumDays = CInt(pBmpDetailDict.Item("NumDays"))
    NumPeople = CInt(pBmpDetailDict.Item("NumPeople"))
    weirangle = CDbl(pBmpDetailDict.Item("BMPTriangularWeirAngle"))
    orificecoef = CDbl(pBmpDetailDict.Item("OrificeCoef"))
    NumUnits = CInt(pBmpDetailDict.Item("NumUnits"))
    ddArea = CDbl(pBmpDetailDict.Item("DrainArea"))
    
    Select Case releaseoption
        Case "Cistern":
            releasetype = 1
            strCisternFlow = pBmpDetailDict.Item("CisternFlow")
            pCisternFlow = Split(strCisternFlow, ";")
        Case "RainBarrel":
            releasetype = 2
        Case "None":
            releasetype = 3
    End Select

    Dim bmpTypeID As String
    If Not gBmpTypeClassDict Is Nothing Then
        If gBmpTypeClassDict.Exists(BMPType) Then
            bmpTypeID = gBmpTypeClassDict.Item(BMPType)
        End If
    End If
    'Define card 715
'    StrCard715BMPTypes = StrCard715BMPTypes & _
'                         iBMPindex & vbTab & _
'                         BMPName & vbTab & _
'                         bmpClass & vbTab & _
'                         pDrainageArea & vbTab & _
'                         pPredevelopedLanduse & vbTab & _
'                         pXpos & vbTab & _
'                         pYpos & vbNewLine
    StrCard715BMPTypes = StrCard715BMPTypes & _
                         strBMPindex & vbTab & _
                         BMPName & vbTab & _
                         bmpTypeID & vbTab & _
                         FormatNumber(pDrainageArea, 2) & vbTab & _
                         NumUnits & vbTab & _
                         ddArea & vbTab & _
                         pPredevelopedLanduse & vbNewLine
    
    'Add Class A Parameters
    StrCard725BMPClassA = StrCard725BMPClassA & _
                          strBMPindex & vbTab & _
                          basinwidth & vbTab & _
                          basinlength & vbTab & _
                          orificeheight & vbTab & _
                          orificediam & vbTab & _
                          exittype & vbTab & _
                          releasetype & vbTab & _
                          NumPeople & vbTab & _
                          NumDays & vbTab & _
                          WeirType & vbTab & _
                          weirheight & vbTab & _
                          weirwidth & vbTab & _
                          weirangle & vbNewLine
                          
    'Add Cistern Flow
    StrCard730BMPClassA = StrCard730BMPClassA & _
                          strBMPindex
    Dim f As Integer
    If (Trim(strCisternFlow) = "") Then
        ReDim pCisternFlow(0 To 23) As Double
    End If
    For f = LBound(pCisternFlow) To UBound(pCisternFlow)
         StrCard730BMPClassA = StrCard730BMPClassA & vbTab & CDbl(pCisternFlow(f))
    Next
    
    StrCard730BMPClassA = StrCard730BMPClassA & vbNewLine
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPParametersA: " & Err.description
CleanUp:
    
End Sub


'******************************************************************************
'Subroutine: GetBMPParametersB
'Author:     Mira Chokshi
'Purpose:    Get the spatial and dimensional parameters for BMP Type B.
'*****************************************************************************
'Private Sub GetBMPParametersB(iBMPindex As Integer, pXpos As Double, pYpos As Double, pDrainageArea As Double)
Private Sub GetBMPParametersB(strBMPindex As String, pXpos As Double, pYpos As Double, pDrainageArea As Double)
On Error GoTo ShowError

    Dim BMPName As String
    Dim BMPType As String
    Dim bmpClass As String
    Dim channellength As Double 'BMP Properties
    Dim channelwidth As Double
    Dim channeldepth As Double
    Dim slope1 As Double
    Dim slope2 As Double
    Dim slope3 As Double
    Dim manncoeff As Double
    
    BMPName = pBmpDetailDict.Item("BMPName")
    BMPType = pBmpDetailDict.Item("BMPType")
    bmpClass = pBmpDetailDict.Item("BMPClass")
    channellength = CDbl(pBmpDetailDict.Item("BMPLength"))  'BMP Properties
    channelwidth = CDbl(pBmpDetailDict.Item("BMPWidth"))
    channeldepth = CDbl(pBmpDetailDict.Item("BMPMaxDepth"))
    slope1 = CDbl(pBmpDetailDict.Item("BMPSlope1"))
    slope2 = CDbl(pBmpDetailDict.Item("BMPSlope2"))
    slope3 = CDbl(pBmpDetailDict.Item("BMPSlope3"))
    manncoeff = CDbl(pBmpDetailDict.Item("BMPManningsN"))
    
    Dim NumUnits As Integer
    Dim ddArea As Double
    
    NumUnits = CInt(pBmpDetailDict.Item("NumUnits"))
    ddArea = CDbl(pBmpDetailDict.Item("DrainArea"))
    
    Dim bmpTypeID As String
    If Not gBmpTypeClassDict Is Nothing Then
        If gBmpTypeClassDict.Exists(BMPType) Then
            bmpTypeID = gBmpTypeClassDict.Item(BMPType)
        End If
    End If
    
    'Define card 715 string
'    StrCard715BMPTypes = StrCard715BMPTypes & _
'                         iBMPindex & vbTab & _
'                         BMPName & vbTab & _
'                         bmpClass & vbTab & _
'                         pDrainageArea & vbTab & _
'                         pPredevelopedLanduse & vbTab & _
'                         pXpos & vbTab & _
'                         pYpos & vbNewLine
    
    'X, y coods not needed
    StrCard715BMPTypes = StrCard715BMPTypes & _
                         strBMPindex & vbTab & _
                         BMPName & vbTab & _
                         bmpTypeID & vbTab & _
                         FormatNumber(pDrainageArea, 2) & vbTab & _
                         NumUnits & vbTab & _
                         ddArea & vbTab & _
                         pPredevelopedLanduse & vbNewLine
                             
    'Define card 735 string
    StrCard735BMPClassB = StrCard735BMPClassB & _
                          strBMPindex & vbTab & _
                          channelwidth & vbTab & _
                          channellength & vbTab & _
                          channeldepth & vbTab & _
                          slope1 & vbTab & _
                          slope2 & vbTab & _
                          slope3 & vbTab & _
                          manncoeff & vbNewLine
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPParametersB: " & Err.description
CleanUp:
End Sub




'******************************************************************************
'Subroutine: GetConduitParameters
'Author:     Mira Chokshi
'Purpose:    Get the dimensional parameters for BMP Type C - Conduit.
'*****************************************************************************
Private Sub GetConduitParameters(iConduitIndex As Integer)
On Error GoTo ShowError

    '*** define all conduit parameters
    Dim bmpClass As String
    Dim conduitShape As String
    Dim conduitInletNode As Integer
    Dim conduitOutletNode As Integer
    Dim conduitBarrels As Integer
    Dim conduitGeom1 As Double  'Max. depth
    Dim conduitGeom2 As Double
    Dim conduitGeom3 As Double
    Dim conduitGeom4 As Double
    Dim conduitLength As Double
    Dim conduitManning As Double
    Dim conduitInitFlow As Double
    Dim conduitHeadLossEnt As Double
    Dim conduitHeadLossExit As Double
    Dim conduitHeadLossAvg As Double
    Dim conduitSlopeEnt As Double
    Dim conduitSlopeExit As Double
    Dim conduitSlopeAvg As Double
    
    'Get all BMP type C - Conduit parameter names
    bmpClass = pBmpDetailDict.Item("BMPClass")
    conduitShape = pBmpDetailDict.Item("TYPE")
    conduitInletNode = pBmpDetailDict.Item("INLETNODE")
    conduitOutletNode = pBmpDetailDict.Item("OUTLETNODE")
    conduitBarrels = pBmpDetailDict.Item("BARRELS")
    conduitGeom1 = pBmpDetailDict.Item("GEOM1")
    conduitGeom2 = pBmpDetailDict.Item("GEOM2")
    conduitGeom3 = pBmpDetailDict.Item("GEOM3")
    conduitGeom4 = pBmpDetailDict.Item("GEOM4")
    conduitLength = pBmpDetailDict.Item("LENGTH")
    conduitManning = pBmpDetailDict.Item("MANN_N")
    conduitInitFlow = pBmpDetailDict.Item("INIFLOW")
    conduitHeadLossEnt = pBmpDetailDict.Item("ENTLOSS")
    conduitHeadLossExit = pBmpDetailDict.Item("EXTLOSS")
    conduitHeadLossAvg = pBmpDetailDict.Item("AVGLOSS")
    conduitSlopeEnt = pBmpDetailDict.Item("ENTINVERTLEV")
    conduitSlopeExit = pBmpDetailDict.Item("EXTINVERTLEV")
    
    Dim bedWidth As Double
    Dim bedDepth As Double
    Dim bedPorosity As Double
    Dim SAND_FRAC As Double
    Dim SILT_FRAC As Double
    Dim CLAY_FRAC As Double
    Dim sandDiameter As Double
    Dim sandVelocity As Double
    Dim sandDensity As Double
    Dim sandCoeff As Double
    Dim sandExponent As Double
    Dim siltDiameter As Double
    Dim siltVelocity As Double
    Dim siltDensity As Double
    Dim siltTaucd As Double
    Dim siltTaucs As Double
    Dim siltM As Double
    
    bedWidth = pBmpDetailDict.Item("Bed width")
    bedDepth = pBmpDetailDict.Item("Bed depth")
    bedPorosity = pBmpDetailDict.Item("Porosity")
    SAND_FRAC = pBmpDetailDict.Item("Sand fraction")
    SILT_FRAC = pBmpDetailDict.Item("Silt fraction")
    CLAY_FRAC = pBmpDetailDict.Item("Clay fraction")
    sandDiameter = pBmpDetailDict.Item("Sand effective diameter")
    sandVelocity = pBmpDetailDict.Item("Sand velocity")
    sandDensity = pBmpDetailDict.Item("Sand density")
    sandCoeff = pBmpDetailDict.Item("Sand coefficient")
    sandExponent = pBmpDetailDict.Item("Sand exponent")
    siltDiameter = pBmpDetailDict.Item("Silt effective diameter")
    siltVelocity = pBmpDetailDict.Item("Silt velocity")
    siltDensity = pBmpDetailDict.Item("Silt density")
    siltTaucd = pBmpDetailDict.Item("Silt Deposition stress")
    siltTaucs = pBmpDetailDict.Item("Silt Scour stress")
    siltM = pBmpDetailDict.Item("Silt Erodibility")
    
    Dim clayDiameter As Double
    Dim clayVelocity As Double
    Dim clayDensity As Double
    Dim clayTaucd As Double
    Dim clayTaucs As Double
    Dim clayM As Double
    
    clayDiameter = pBmpDetailDict.Item("Clay effective diameter")
    clayVelocity = pBmpDetailDict.Item("Clay velocity")
    clayDensity = pBmpDetailDict.Item("Clay density")
    clayTaucd = pBmpDetailDict.Item("Clay Deposition stress")
    clayTaucs = pBmpDetailDict.Item("Clay Scour stress")
    clayM = pBmpDetailDict.Item("Clay Erodibility")
    
    Dim bmpTypeID As String
    If Not gBmpTypeClassDict Is Nothing Then
        If gBmpTypeClassDict.Exists("Conduit") Then
            bmpTypeID = gBmpTypeClassDict.Item("Conduit")
        End If
    End If
    
    'Define card 715 string
'    StrCard715BMPTypes = StrCard715BMPTypes & _
'                         iConduitIndex & vbTab & _
'                         "Conduit" & vbTab & _
'                         bmpClass & vbTab & _
'                         0 & vbTab & _
'                         pPredevelopedLanduse & vbTab & _
'                         0 & vbTab & _
'                         0 & vbNewLine
    StrCard715BMPTypes = StrCard715BMPTypes & _
                         iConduitIndex & vbTab & _
                         "Conduit" & vbTab & _
                         bmpTypeID & vbTab & _
                         0 & vbTab & _
                         1 & vbTab & _
                         0 & vbTab & _
                         pPredevelopedLanduse & vbNewLine
                         
    'Define card 750 string
    StrCard750ConduitDimensions = StrCard750ConduitDimensions & _
                         iConduitIndex & vbTab & _
                         conduitInletNode & vbTab & _
                         conduitOutletNode & vbTab & _
                         conduitLength & vbTab & _
                         conduitManning & vbTab & _
                         conduitSlopeEnt & vbTab & _
                         conduitSlopeExit & vbTab & _
                         conduitInitFlow & vbTab & _
                         conduitHeadLossEnt & vbTab & _
                         conduitHeadLossExit & vbTab & _
                         conduitHeadLossAvg & vbNewLine
                             
    'Define card 755 string
    StrCard755ConduitCrossSections = StrCard755ConduitCrossSections & _
                          iConduitIndex & vbTab & _
                          conduitShape & vbTab & _
                          conduitGeom1 & vbTab & _
                          conduitGeom2 & vbTab & _
                          conduitGeom3 & vbTab & _
                          conduitGeom4 & vbTab & _
                          conduitBarrels & vbNewLine
    StrCard775Sediment = StrCard775Sediment & _
                        iConduitIndex & vbTab & _
                        bedWidth & vbTab & _
                        bedDepth & vbTab & _
                        bedPorosity & vbTab & _
                        SAND_FRAC & vbTab & _
                        SILT_FRAC & vbTab & _
                        CLAY_FRAC & vbNewLine
                        
    StrCard780SandTransport = StrCard780SandTransport & _
                        iConduitIndex & vbTab & _
                        sandDiameter & vbTab & _
                        sandVelocity & vbTab & _
                        sandDensity & vbTab & _
                        sandCoeff & vbTab & _
                        sandExponent & vbNewLine
    StrCard785SiltTransport = StrCard785SiltTransport & _
                        iConduitIndex & vbTab & _
                        siltDiameter & vbTab & _
                        siltVelocity & vbTab & _
                        siltDensity & vbTab & _
                        siltTaucd & vbTab & _
                        siltTaucs & vbTab & _
                        siltM & vbNewLine
    StrCard786ClayTransport = StrCard786ClayTransport & _
                        iConduitIndex & vbTab & _
                        clayDiameter & vbTab & _
                        clayVelocity & vbTab & _
                        clayDensity & vbTab & _
                        clayTaucd & vbTab & _
                        clayTaucs & vbTab & _
                        clayM & vbNewLine
   'Commented the following - June 18, 2007
''    'Define card 1190 string
''    StrCard1190ConduitLosses = StrCard1190ConduitLosses & _
''                          iConduitIndex & vbTab & _
''                          conduitHeadLossEnt & vbTab & _
''                          conduitHeadLossExit & vbTab & _
''                          conduitHeadLossAvg & vbTab & _
''                          "NO" & vbNewLine
                          
    GoTo CleanUp
ShowError:
    MsgBox "GetConduitParameters: " & Err.description
CleanUp:
End Sub

'******************************************************************************
'Subroutine: GetVFSParameters
'Author:     Mira Chokshi
'Purpose:    Get the dimensional parameters for BMP Type D - Buffer Strip
'*****************************************************************************
Private Sub GetVFSParameters(iVFSIndex As Integer)
On Error GoTo ShowError
    
    Dim i As Integer
        
    '*** define all conduit parameters
    Dim bmpClass As String
    Dim vfsLength As Double
    Dim vfsWidth As Double
        
    'Get all BMP type D - Buffer Strip parameter names
    bmpClass = "D"  'pBMPDetailDict.Item("BMPClass")
'    vfsLength = pBMPDetailDict.Item("Length")
'    vfsWidth = pBMPDetailDict.Item("Width")
    
    'Change in control name - Sabu Paul, June 7 2007
    vfsLength = pBmpDetailDict.Item("BufferLength")
    vfsWidth = pBmpDetailDict.Item("BufferWidth")

    Dim NPROP As Integer
    NPROP = CInt(pBmpDetailDict.Item("NPROP"))
    
    Dim bmpTypeID As String
    If Not gBmpTypeClassDict Is Nothing Then
        If gBmpTypeClassDict.Exists("VFS") Then
            bmpTypeID = gBmpTypeClassDict.Item("VFS")
        End If
    End If
    
''    'Define card 1100 string
''    StrCard1100BMPTypes = StrCard1100BMPTypes & _
''                         iVFSIndex & vbTab & _
''                         "VFS" & vbTab & _
''                         bmpClass & vbTab & _
''                         0 & vbTab & _
''                         pPredevelopedLanduse & vbTab & _
''                         0 & vbTab & _
''                         0 & vbNewLine
''
''    'Define card 1200 string
''    StrCard1200VFSParameters = StrCard1200VFSParameters & _
''                         iVFSIndex & vbTab & _
''                         vfsWidth & vbTab & _
''                         "0" & vbTab & _
''                         "0" & vbTab & _
''                         "0" & vbNewLine
''
''    'Define card 1210 string
''    StrCard1210VFSParameters = StrCard1210VFSParameters & _
''                         iVFSIndex & vbTab & _
''                         "1" & vbTab & _
''                         vfsLength & vbTab & _
''                         "0.01" & vbTab & _
''                         "0.14" & vbNewLine

    'Define card 715 string
'    StrCard715BMPTypes = StrCard715BMPTypes & _
'                         iVFSIndex & vbTab & _
'                         "VFS" & vbTab & _
'                         bmpClass & vbTab & _
'                         0 & vbTab & _
'                         pPredevelopedLanduse & vbTab & _
'                         0 & vbTab & _
'                         0 & vbNewLine

    Dim vfsDA As Double
    vfsDA = 0
    If gBMPDrainAreaDict.Exists(iVFSIndex) Then vfsDA = gBMPDrainAreaDict.Item(iVFSIndex)
    
    
    StrCard715BMPTypes = StrCard715BMPTypes & _
                         iVFSIndex & vbTab & _
                         pBmpDetailDict.Item("Name") & vbTab & _
                         bmpTypeID & vbTab & _
                         FormatNumber(vfsDA, 2) & vbTab & _
                         1 & vbTab & _
                         0 & vbTab & _
                         pPredevelopedLanduse & vbNewLine

    'Define card 901 string
    StrCard901_VFS_Dim = StrCard901_VFS_Dim & _
                         iVFSIndex & vbTab & _
                         pBmpDetailDict.Item("Name") & vbTab & _
                         vfsWidth & vbTab & _
                         vfsLength & vbTab & _
                         NPROP & vbNewLine
                      
    'Define card 902 string
    For i = 1 To NPROP
        StrCard902_VFS_SegDetails = StrCard902_VFS_SegDetails & _
                    iVFSIndex & vbTab & _
                    i & vbTab & _
                    pBmpDetailDict.Item("SX" & i) & vbTab & _
                    pBmpDetailDict.Item("RNA" & i) & vbTab & _
                    pBmpDetailDict.Item("SOA" & i) & vbNewLine
                         
    Next
    
    'Define card 903 string
    StrCard903_VFS_SoilProps = StrCard903_VFS_SoilProps & _
                    iVFSIndex & vbTab & _
                    pBmpDetailDict.Item("VKS") & vbTab & _
                    pBmpDetailDict.Item("Sav") & vbTab & _
                    pBmpDetailDict.Item("OS") & vbTab & _
                    pBmpDetailDict.Item("OI") & vbTab & _
                    pBmpDetailDict.Item("SM") & vbTab & _
                    pBmpDetailDict.Item("SCHK") & vbNewLine
    
    'Define card 904 string
    StrCard904_VFS_Buf_Sed = StrCard904_VFS_Buf_Sed & _
                    iVFSIndex & vbTab & _
                    pBmpDetailDict.Item("SS") & vbTab & _
                    pBmpDetailDict.Item("VN") & vbTab & _
                    pBmpDetailDict.Item("H") & vbTab & _
                    pBmpDetailDict.Item("Vn2") & vbTab & _
                    0 & vbNewLine 'card ICO - no feed back
    
     'Define card 905 string
    StrCard905_VFS_Sed_Filt = StrCard905_VFS_Sed_Filt & _
                    iVFSIndex & vbTab & _
                    pBmpDetailDict.Item("NPARTSand") & vbTab & _
                    pBmpDetailDict.Item("NPARTSilt") & vbTab & _
                    pBmpDetailDict.Item("NPARTClay") & vbTab & _
                    pBmpDetailDict.Item("COARSESand") & vbTab & _
                    pBmpDetailDict.Item("COARSESilt") & vbTab & _
                    pBmpDetailDict.Item("COARSEClay") & vbTab & _
                    pBmpDetailDict.Item("PORSand") & vbTab & _
                    pBmpDetailDict.Item("PORSilt") & vbTab & _
                    pBmpDetailDict.Item("PORClay") & vbTab & _
                    pBmpDetailDict.Item("DPSand") & vbTab & _
                    pBmpDetailDict.Item("DPSilt") & vbTab & _
                    pBmpDetailDict.Item("DPClay") & vbTab & _
                    pBmpDetailDict.Item("SGSand") & vbTab & _
                    pBmpDetailDict.Item("SGSilt") & vbTab & _
                    pBmpDetailDict.Item("SGClay") & vbNewLine
    
    Call CreatePollutantList
    Dim pTotalPollutants As Integer
    pTotalPollutants = UBound(gPollutants) + 1
    
    'Check to see whether DBF file exists
'''    Dim pMultiplierTable As iTable
'''    Set pMultiplierTable = GetInputDataTable("TSMultipliers")
'''
'''
'''    Dim sedPollInd As Integer
'''
'''    Dim pRow As iRow
'''    Dim pCursor As esriGeoDatabase.ICursor
'''    Dim iR As Integer
'''
'''    If pMultiplierTable Is Nothing Then
'''        MsgBox "TSMultipliers table is missing. Can not identify sediment"
'''        Exit Sub
'''    Else
'''
'''        Set pCursor = pMultiplierTable.Search(Nothing, False)
'''        Dim pSedFlagInd As Integer
'''        pSedFlagInd = pMultiplierTable.FindField("SedFlag")
'''
'''        Set pRow = pCursor.NextRow
'''        iR = 1
'''        Do While Not pRow Is Nothing
'''            If pRow.value(pSedFlagInd) = 1 Then
'''                sedPollInd = iR 'Set the pollutant index for sediment
'''                Exit Do
'''            End If
'''            Set pRow = pCursor.NextRow
'''            iR = iR + 1
'''        Loop
'''    End If
    
    StrCard906_VFS_Sed_Fracion = StrCard906_VFS_Sed_Fracion & iVFSIndex
    StrCard907_VFS_FO_Adsorbed = StrCard907_VFS_FO_Adsorbed & iVFSIndex
    StrCard908_VFS_FO_Dissolved = StrCard908_VFS_FO_Dissolved & iVFSIndex
    StrCard909_VFS_TC_Adsorbed = StrCard909_VFS_TC_Adsorbed & iVFSIndex
    StrCard910_VFS_TC_Dissolved = StrCard910_VFS_TC_Dissolved & iVFSIndex
    
    Dim iR As Integer
    For iR = 1 To pTotalPollutants
        'If iR <> sedPollInd Then
        If pBmpDetailDict.Exists("SedFrac" & iR) Then
            StrCard906_VFS_Sed_Fracion = StrCard906_VFS_Sed_Fracion & vbTab & pBmpDetailDict.Item("SedFrac" & iR)
            StrCard907_VFS_FO_Adsorbed = StrCard907_VFS_FO_Adsorbed & vbTab & pBmpDetailDict.Item("SedDec" & iR)
            StrCard908_VFS_FO_Dissolved = StrCard908_VFS_FO_Dissolved & vbTab & pBmpDetailDict.Item("WatDec" & iR)
            StrCard909_VFS_TC_Adsorbed = StrCard909_VFS_TC_Adsorbed & vbTab & pBmpDetailDict.Item("SedCorr" & iR)
            StrCard910_VFS_TC_Dissolved = StrCard910_VFS_TC_Dissolved & vbTab & pBmpDetailDict.Item("WatCorr" & iR)
        End If
    Next
    
    StrCard906_VFS_Sed_Fracion = StrCard906_VFS_Sed_Fracion & vbNewLine
    StrCard907_VFS_FO_Adsorbed = StrCard907_VFS_FO_Adsorbed & vbNewLine
    StrCard908_VFS_FO_Dissolved = StrCard908_VFS_FO_Dissolved & vbNewLine
    StrCard909_VFS_TC_Adsorbed = StrCard909_VFS_TC_Adsorbed & vbNewLine
    StrCard910_VFS_TC_Dissolved = StrCard910_VFS_TC_Dissolved & vbNewLine

    Exit Sub
ShowError:
    MsgBox "Error in GetVFSParameters: " & Err.description
CleanUp:
End Sub


'******************************************************************************
'Subroutine: GetDummyBMPParameters
'Author:     Mira Chokshi
'Purpose:    Get the spatial and dimensional parameters for DUMMY BMP.
'*****************************************************************************
'Private Sub GetDummyBMPParameters(iBMPindex As Integer, pXpos As Double, pYpos As Double, pDrainageArea As Double)
Private Sub GetDummyBMPParameters(strBMPindex As String, pXpos As Double, pYpos As Double, pDrainageArea As Double)
On Error GoTo ShowError
    
    Dim BMPType As String
    Dim bmpClass As String
    
    BMPType = pBmpDetailDict.Item("BMPType")
    bmpClass = pBmpDetailDict.Item("BMPClass")
    
    Dim bmpTypeID As String
    If Not gBmpTypeClassDict Is Nothing Then
        If gBmpTypeClassDict.Exists(BMPType) Then
            bmpTypeID = gBmpTypeClassDict.Item(BMPType)
        End If
    End If
    
    'Define card 715 string
'    StrCard715BMPTypes = StrCard715BMPTypes & _
'                         iBMPindex & vbTab & _
'                         bmpType & vbTab & _
'                         bmpClass & vbTab & _
'                         pDrainageArea & vbTab & _
'                         pPredevelopedLanduse & vbTab & _
'                         pXpos & vbTab & _
'                         pYpos & vbNewLine
    StrCard715BMPTypes = StrCard715BMPTypes & _
                         strBMPindex & vbTab & _
                         BMPType & vbTab & _
                         bmpTypeID & vbTab & _
                         FormatNumber(pDrainageArea, 2) & vbTab & _
                         1 & vbTab & _
                         0 & vbTab & _
                         pPredevelopedLanduse & vbNewLine
                             
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPParametersB: " & Err.description
CleanUp:
End Sub

'''******************************************************************************
'''Subroutine: GetAssessPointDetails
'''Author:     Mira Chokshi
'''Purpose:    Assessment point is assumed to be a class A BMP with zero dimensions
'''            Get the spatial information for assessment point and evaluation
'''            factor.
'''            Mira Chokshi, Sabu Paul modified 08/23/04 to change output format
'''*****************************************************************************
''Private Sub GetAssessPointDetails(iBMPindex As Integer, pXpos As Double, pYpos As Double, pDrainageArea As Double)
''On Error GoTo ShowError
''
''    Dim BMPName As String
''    Dim bmpType As String
''    Dim bmpClass As String
''
''    Dim pBMPKeys
''    pBMPKeys = pBMPDetailDict.Keys
''
''    BMPName = pBMPDetailDict.Item("BMPName")
''    bmpType = pBMPDetailDict.Item("BMPType")
''    bmpClass = pBMPDetailDict.Item("BMPClass")
''
''    If bmpClass = "X" Then
''        StrCard1100BMPTypes = StrCard1100BMPTypes & _
''                     iBMPindex & vbTab & _
''                     "AssessmentPoint" & vbTab & _
''                     bmpClass & vbTab & _
''                     pDrainageArea & vbTab & _
''                     pPredevelopedLanduse & vbTab & _
''                     pXpos & vbTab & _
''                     pYpos & vbNewLine
''    End If
''
''    GoTo CleanUp
''ShowError:
''    MsgBox "GetAssessPointDetails: " & Err.description
''CleanUp:
''    Set pBMPKeys = Nothing
''End Sub

'******************************************************************************
'Subroutine: GetSoilIndex
'Author:     Mira Chokshi
'Purpose:    Get the spatial and dimensional parameters for BMP Type B.
'*****************************************************************************
Private Sub GetSoilIndex(strBMPindex As String) '(iBMPindex As Integer)
On Error GoTo ShowError
    Dim SoilDepth As Double     'Soil Growth Properties
    Dim SoilPorosity As Double
    Dim udsoildepth As Double
    Dim udsoilporosity As Double
    Dim udfinalf As Double
    Dim finalf As Double
    Dim vegparma As Double
    Dim underdrain_on As Integer
    
    Dim infiltm As Integer
    Dim polRoutM As Integer
    Dim polRemovM As Integer
    Dim soilFC As Double
    Dim soilWP As Double
    Dim Suction As Double
    Dim HydCon  As Double
    Dim IMDmax As Double

    SoilDepth = CDbl(pBmpDetailDict.Item("SoilDepth"))
    SoilPorosity = CDbl(pBmpDetailDict.Item("SoilPorosity"))
    udsoildepth = CDbl(pBmpDetailDict.Item("StorageDepth"))
    udsoilporosity = CDbl(pBmpDetailDict.Item("VoidFraction"))
    udfinalf = CDbl(pBmpDetailDict.Item("BackgroundInfiltration"))
    vegparma = CDbl(pBmpDetailDict.Item("VegetativeParam"))
    finalf = CDbl(pBmpDetailDict.Item("SoilLayerInfiltration"))
    underdrain_on = 0
    If (CBool(pBmpDetailDict.Item("UnderDrainON")) = True) Then
        underdrain_on = 1
    End If
    
    infiltm = pBmpDetailDict.Item("Infiltration Method")

    polRoutM = pBmpDetailDict.Item("Pollutant Routing Method")
    polRemovM = pBmpDetailDict.Item("Pollutant Removal Method")
    soilFC = CDbl(pBmpDetailDict.Item("SoilFieldCapacity"))
    soilWP = CDbl(pBmpDetailDict.Item("SoilWiltingPoint"))
    Suction = CDbl(pBmpDetailDict.Item("SuctionHead"))
    HydCon = CDbl(pBmpDetailDict.Item("Conductivity"))
    IMDmax = CDbl(pBmpDetailDict.Item("InitialDeficit"))

'    StrCard740BMPSoilIndex = StrCard740BMPSoilIndex & _
'                             iBMPindex & vbTab & _
'                             SoilDepth & vbTab & _
'                             SoilPorosity & vbTab & _
'                             vegparma & vbTab & _
'                             finalf & vbTab & _
'                             underdrain_on & vbTab & _
'                             udsoildepth & vbTab & _
'                             udsoilporosity & vbTab & _
'                             udfinalf & vbNewLine
    StrCard740BMPSoilIndex = StrCard740BMPSoilIndex & _
                             strBMPindex & vbTab & _
                             infiltm & vbTab & _
                             polRoutM & vbTab & _
                             polRemovM & vbTab & _
                             SoilDepth & vbTab & _
                             SoilPorosity & vbTab & _
                             soilFC & vbTab & _
                             soilWP & vbTab & _
                             vegparma & vbTab & _
                             finalf & vbTab & _
                             underdrain_on & vbTab & _
                             udsoildepth & vbTab & _
                             udsoilporosity & vbTab & _
                             udfinalf & vbTab & _
                             Suction & vbTab & _
                             HydCon & vbTab & _
                             IMDmax & vbNewLine
    'if pollutant removal is kadlec and knight method
    GetBMP_K_C_factors strBMPindex  'iBMPindex
''    If polRemovM = 1 Then
''        GetBMP_K_C_factors iBMPindex
''    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "GetSoilIndex: " & Err.description
CleanUp:
End Sub


'******************************************************************************
'Subroutine: GetBMPDecayFactors
'Author:     Mira Chokshi
'Purpose:    Get the decay factors for each bmp/nutrient and write to a string.
'*****************************************************************************
Public Sub GetBMPDecayFactors(pBMPIdStr As String) '(pBMPId As Integer)
On Error GoTo ShowError
    'pBmpDetailDict
    
    '*** BMPDetail table
'''    Dim pTable As iTable
'''    Set pTable = GetInputDataTable("BMPDetail")
'''    If (pTable Is Nothing) Then
'''        MsgBox ("BMPDetail table not found.")
'''        Exit Sub
'''    End If
'''
'''    Dim pQueryFilter As IQueryFilter
'''    Set pQueryFilter = New QueryFilter
'''    Dim pCursor As ICursor
'''    Dim pRow As iRow
'''    Dim iPropValue As Long
'''    iPropValue = pTable.FindField("PropValue")
'''
'''    Dim pStrCard765ForBMP As String
'''    pStrCard765ForBMP = ""
'''    Dim pDecayFactorValue As Double
'''    Dim i As Integer
'''    For i = 1 To pTotalPollutantCount
'''        pQueryFilter.WhereClause = "ID = " & pBMPId & " AND PropName = 'Decay" & i & "'"
'''        Set pCursor = pTable.Search(pQueryFilter, True)
'''        Set pRow = pCursor.NextRow
'''        pDecayFactorValue = 0
'''        If Not (pRow Is Nothing) Then
'''             pDecayFactorValue = pRow.value(iPropValue)
'''        End If
'''        pStrCard765ForBMP = pStrCard765ForBMP & pDecayFactorValue & vbTab
'''    Next
'''    If (pStrCard765ForBMP <> "") Then
'''         StrCard765DecayFactors = StrCard765DecayFactors & pBMPId & vbTab & pStrCard765ForBMP & vbNewLine
'''    End If

    Dim pStrCard765ForBMP As String
    pStrCard765ForBMP = ""
    Dim pDecayFactorValue As Double
    Dim i As Integer
    For i = 1 To pTotalPollutantCount
        pDecayFactorValue = 0
        If pBmpDetailDict.Exists("Decay" & i) Then pDecayFactorValue = pBmpDetailDict.Item("Decay" & i)
        'pStrCard765ForBMP = pStrCard765ForBMP & Format(pDecayFactorValue / 24, "#.####") & vbTab
        pStrCard765ForBMP = pStrCard765ForBMP & FormatNumber(pDecayFactorValue / 24, 4) & vbTab
    Next
    If (pStrCard765ForBMP <> "") Then
         StrCard765DecayFactors = StrCard765DecayFactors & pBMPIdStr & vbTab & pStrCard765ForBMP & vbNewLine
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPDecayFactors: " & Err.description
CleanUp:
'    Set pTable = Nothing
'    Set pCursor = Nothing
'    Set pRow = Nothing
End Sub


'******************************************************************************
'Subroutine: GetBMPPercentRemovalFactors
'Author:     Mira Chokshi
'Purpose:    Get the percent removal factors for each bmp/nutrient
'            and write to a string.
'*****************************************************************************
Private Sub GetBMPPercentRemovalFactors(pBMPIdStr As String) '(pBMPId As Integer)
On Error GoTo ShowError
    '***
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("BMPDetail")
''    If (pTable Is Nothing) Then
''        MsgBox ("BMPDetail table not found.")
''        Exit Sub
''    End If
''
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''    Dim pCursor As ICursor
''    Dim pRow As iRow
''    Dim iPropValue As Long
''    iPropValue = pTable.FindField("PropValue")
''
''    Dim pStrCard770ForBMP As String
''    pStrCard770ForBMP = ""
''    Dim i As Integer
''    Dim pPercentRemovalValue As Double
''    For i = 1 To pTotalPollutantCount
''        pQueryFilter.WhereClause = "ID = " & pBMPId & " AND PropName = 'PctRem" & i & "'"
''        Set pCursor = pTable.Search(pQueryFilter, True)
''        Set pRow = pCursor.NextRow
''        pPercentRemovalValue = 0
''        If Not (pRow Is Nothing) Then
''              pPercentRemovalValue = pRow.value(iPropValue)
''        End If
''        pStrCard770ForBMP = pStrCard770ForBMP & pPercentRemovalValue & vbTab
''    Next
''    If (pStrCard770ForBMP <> "") Then
''        StrCard770PercentRemoval = StrCard770PercentRemoval & pBMPId & vbTab & pStrCard770ForBMP & vbNewLine
''    End If
    Dim pStrCard770ForBMP As String
    pStrCard770ForBMP = ""
    Dim i As Integer
    Dim pPercentRemovalValue As Double
    For i = 1 To pTotalPollutantCount
        pPercentRemovalValue = 0
        If pBmpDetailDict.Exists("PctRem" & i) Then pPercentRemovalValue = pBmpDetailDict.Item("PctRem" & i)
        pStrCard770ForBMP = pStrCard770ForBMP & pPercentRemovalValue & vbTab
    Next
    If (pStrCard770ForBMP <> "") Then
        StrCard770PercentRemoval = StrCard770PercentRemoval & pBMPIdStr & vbTab & pStrCard770ForBMP & vbNewLine
    End If
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPPercentRemovalFactors: " & Err.description
CleanUp:
'    Set pRow = Nothing
'    Set pCursor = Nothing
'    Set pTable = Nothing
'    Set pQueryFilter = Nothing
End Sub

Private Sub GetBMP_K_C_factors(pBMPID As String) 'pBMPId As Integer)
On Error GoTo ShowError
    '***
''    Dim pTable As iTable
''    Set pTable = GetInputDataTable("BMPDetail")
''    If (pTable Is Nothing) Then
''        MsgBox ("BMPDetail table not found.")
''        Exit Sub
''    End If
''
''    Dim pQueryFilter As IQueryFilter
''    Set pQueryFilter = New QueryFilter
''    Dim pCursor As ICursor
''    Dim pRow As iRow
''    Dim iPropValue As Long
''    iPropValue = pTable.FindField("PropValue")
''
''    Dim i As Integer
''    Dim kValue As Double
''    Dim cValue As Double
''
''    StrCard766KFactors = StrCard766KFactors & pBMPId
''    StrCard767CValues = StrCard767CValues & pBMPId
''
''    For i = 1 To pTotalPollutantCount
''        pQueryFilter.WhereClause = "ID = " & pBMPId & " AND PropName = 'K" & i & "'"
''        Set pCursor = pTable.Search(pQueryFilter, True)
''        Set pRow = pCursor.NextRow
''        kValue = 0
''        If Not (pRow Is Nothing) Then
''              kValue = pRow.value(iPropValue)
''        End If
''        StrCard766KFactors = StrCard766KFactors & vbTab & kValue
''
''        pQueryFilter.WhereClause = "ID = " & pBMPId & " AND PropName = 'C" & i & "'"
''        Set pCursor = pTable.Search(pQueryFilter, True)
''        Set pRow = pCursor.NextRow
''        cValue = 0
''        If Not (pRow Is Nothing) Then
''              cValue = pRow.value(iPropValue)
''        End If
''        StrCard767CValues = StrCard767CValues & vbTab & cValue
''    Next
''    StrCard766KFactors = StrCard766KFactors & vbNewLine
''    StrCard767CValues = StrCard767CValues & vbNewLine
   
    Dim i As Integer
    Dim kValue As Double
    Dim cValue As Double
    
    StrCard766KFactors = StrCard766KFactors & pBMPID
    StrCard767CValues = StrCard767CValues & pBMPID
    
    For i = 1 To pTotalPollutantCount
        kValue = 0
        If pBmpDetailDict.Exists("K" & i) Then kValue = pBmpDetailDict.Item("K" & i)
        StrCard766KFactors = StrCard766KFactors & vbTab & kValue
        
        cValue = 0
        If pBmpDetailDict.Exists("C" & i) Then cValue = pBmpDetailDict.Item("C" & i)
        StrCard767CValues = StrCard767CValues & vbTab & cValue
    Next
    StrCard766KFactors = StrCard766KFactors & vbNewLine
    StrCard767CValues = StrCard767CValues & vbNewLine
    GoTo CleanUp
ShowError:
    MsgBox "GetBMP_K_C_factors: " & Err.description
CleanUp:
'    Set pRow = Nothing
'    Set pCursor = Nothing
'    Set pTable = Nothing
'    Set pQueryFilter = Nothing
End Sub

'******************************************************************************
'Subroutine: GetBMPCosts
'Author:     Mira Chokshi
'Purpose:    Get the cost function parameters for each bmp and write to a string.
'*****************************************************************************
Private Sub GetBMPCosts(strBMPindex As String) 'iBMPindex As Integer
On Error GoTo ShowError
''    Dim bmpAa As Double
''    Dim bmpAb As Double
''    Dim bmpDa As Double
''    Dim bmpDb As Double
''    Dim bmpLdCost As Double
''    Dim bmpConstCost As Double
''
''    bmpAa = CDbl(pBmpDetailDict.Item("Aa"))
''    bmpAb = CDbl(pBmpDetailDict.Item("Ab"))
''    bmpDa = CDbl(pBmpDetailDict.Item("Da"))
''    bmpDb = CDbl(pBmpDetailDict.Item("Db"))
''    bmpLdCost = CDbl(pBmpDetailDict.Item("LdCost"))
''    bmpConstCost = CDbl(pBmpDetailDict.Item("ConstCost"))
''
''    StrCard805BMPCost = StrCard805BMPCost & _
''                            iBMPindex & vbTab & _
''                            bmpAa & vbTab & _
''                            bmpAb & vbTab & _
''                            bmpDa & vbTab & _
''                            bmpDb & vbTab & _
''                            bmpLdCost & vbTab & _
''                            bmpConstCost & vbNewLine

    Dim costUnits
    Dim costAdjUnitCosts
    Dim costVolTypes
    Dim costNumUnits
    
    Dim volCostTotal As Double: volCostTotal = 0#
    Dim areaCost As Double: areaCost = 0#
    Dim linearCost As Double: linearCost = 0#
    Dim numCost As Double: numCost = 0#
    Dim consCost As Double: consCost = 0#
    Dim percCost As Double: percCost = 0#
    Dim volCostMedia As Double: volCostMedia = 0#
    Dim volCostUnderDrain As Double: volCostUnderDrain = 0#
           
    Dim adjCost As Double
    Dim i As Integer
           
    Dim costExps
    Dim costExpDict As Scripting.Dictionary
    Set costExpDict = New Scripting.Dictionary
    
    If pBmpDetailDict.Exists("CostComponents") Then
        costUnits = Split(pBmpDetailDict.Item("CostUnits"), ";", , vbTextCompare)
        costAdjUnitCosts = Split(pBmpDetailDict.Item("CostAdjUnitCosts"), ";")
        
        costVolTypes = Split(pBmpDetailDict.Item("CostVolTypes"), ";", , vbTextCompare)
        costNumUnits = Split(pBmpDetailDict.Item("CostNumUnits"), ";", , vbTextCompare)
        costExps = Split(pBmpDetailDict.Item("costExponents"), ";", , vbTextCompare)
        
        For i = 0 To UBound(costUnits)
            adjCost = CDbl(costAdjUnitCosts(i))
            costExpDict.Item(UCase(costUnits(i)) & "_" & CInt(costVolTypes(i))) = costExps(i)
            
            Select Case UCase(costUnits(i))
                Case "FEET"
                    linearCost = linearCost + adjCost
                Case "SQUARE FEET"
                    areaCost = areaCost + adjCost
                Case "CUBIC FEET"
                    If CInt(costVolTypes(i)) = COST_VOLUME_TYPE_TOTAL Then
                        volCostTotal = volCostTotal + adjCost
                    ElseIf CInt(costVolTypes(i)) = COST_VOLUME_TYPE_MEDIA Then
                        volCostMedia = volCostMedia + adjCost
                    Else
                        volCostUnderDrain = volCostUnderDrain + adjCost
                    End If
                Case "PERCENTAGE"
                    percCost = percCost + adjCost
                Case "CONSTANT"
                    consCost = consCost + adjCost
                Case "PER UNIT"
                    'numCost = numCost + (adjCost * CInt(costNumUnits(i)))
                    consCost = consCost + (adjCost * CInt(costNumUnits(i)))
            End Select
        Next
    End If
        
    
'    StrCard805BMPCost = StrCard805BMPCost & _
'                            strBMPindex & vbTab & _
'                            linearCost & vbTab & _
'                            areaCost & vbTab & _
'                            volCostTotal & vbTab & _
'                            volCostMedia & vbTab & _
'                            volCostUnderDrain & vbTab & _
'                            numCost & vbTab & _
'                            consCost & vbTab & _
'                            percCost & vbNewLine
'    StrCard805BMPCost = StrCard805BMPCost & _
'                            strBMPindex & vbTab & _
'                            linearCost & vbTab & _
'                            areaCost & vbTab & _
'                            volCostTotal & vbTab & _
'                            volCostMedia & vbTab & _
'                            volCostUnderDrain & vbTab & _
'                            consCost & vbTab & _
'                            percCost & vbNewLine
    Dim ftExp As String
    ftExp = 1
    If costExpDict.Exists("FEET_1") Then ftExp = costExpDict.Item("FEET_1")
    Dim sqftExp As String
    sqftExp = 1
    If costExpDict.Exists("SQUARE FEET_1") Then sqftExp = costExpDict.Item("SQUARE FEET_1")
    Dim cubftExp1 As String, cubftExp2 As String, cubftExp3 As String
    cubftExp1 = 1: cubftExp2 = 1: cubftExp3 = 1
    If costExpDict.Exists("CUBIC FEET_" & COST_VOLUME_TYPE_TOTAL) Then cubftExp1 = costExpDict.Item("CUBIC FEET_" & COST_VOLUME_TYPE_TOTAL)
    If costExpDict.Exists("CUBIC FEET_" & COST_VOLUME_TYPE_MEDIA) Then cubftExp2 = costExpDict.Item("CUBIC FEET_" & COST_VOLUME_TYPE_MEDIA)
    If costExpDict.Exists("CUBIC FEET_" & COST_VOLUME_TYPE_UNDERDRAIN) Then cubftExp3 = costExpDict.Item("CUBIC FEET_" & COST_VOLUME_TYPE_UNDERDRAIN)
    
    StrCard805BMPCost = StrCard805BMPCost & _
                            strBMPindex & vbTab & _
                            linearCost & vbTab & _
                            areaCost & vbTab & _
                            volCostTotal & vbTab & _
                            volCostMedia & vbTab & _
                            volCostUnderDrain & vbTab & _
                            consCost & vbTab & _
                            percCost & vbTab & _
                            ftExp & vbTab & _
                            sqftExp & vbTab & _
                            cubftExp1 & vbTab & _
                            cubftExp2 & vbTab & _
                            cubftExp3 & vbNewLine
      
    GoTo CleanUp
ShowError:
    MsgBox "GetBMPCosts: " & Err.description
CleanUp:
End Sub


'******************************************************************************
'Subroutine: GetOptimizationParameters
'Author:     Mira Chokshi
'Purpose:    Get the parameters that will be used for optimization.
'            The optimized parameter will have lower value, higher value and
'            increment value.
'*****************************************************************************
Private Sub GetOptimizationParameters(strBMPindex As String) 'iBMPindex As Integer)
On Error GoTo ShowError

    Dim lengthMin As Double     'Adjustable parameters
    Dim lengthMax As Double
    Dim lengthStep As Double
    Dim widthMin As Double
    Dim widthMax As Double
    Dim widthStep As Double
    Dim weirHMin As Double
    Dim weirHMax As Double
    Dim weirHStep As Double
    
    Dim soilDMin As Double
    Dim soilDMax As Double
    Dim soilDStep As Double
    
        
    If pBmpDetailDict.Exists("BLengthOptimized") Then
        If (CBool(pBmpDetailDict.Item("BLengthOptimized"))) Then
            lengthMin = CDbl(pBmpDetailDict.Item("MinBasinLength"))
            lengthMax = CDbl(pBmpDetailDict.Item("MaxBasinLength"))
            lengthStep = CDbl(pBmpDetailDict.Item("BasinLengthIncr"))
           
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "Length" & vbTab & _
                                            lengthMin & vbTab & _
                                            lengthMax & vbTab & _
                                            lengthStep & vbNewLine
        End If
    End If
    
    If pBmpDetailDict.Exists("BWidthOptimized") Then
        If (CBool(pBmpDetailDict.Item("BWidthOptimized"))) Then
            widthMin = CDbl(pBmpDetailDict.Item("MinBasinWidth"))
            widthMax = CDbl(pBmpDetailDict.Item("MaxBasinWidth"))
            widthStep = CDbl(pBmpDetailDict.Item("BasinWidthIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "Width" & vbTab & _
                                            widthMin & vbTab & _
                                            widthMax & vbTab & _
                                            widthStep & vbNewLine
        End If
    End If
    
    If pBmpDetailDict.Exists("WHeightOptimized") Then
        If (CBool(pBmpDetailDict.Item("WHeightOptimized"))) Then
            weirHMin = CDbl(pBmpDetailDict.Item("MinWeirHeight"))
            weirHMax = CDbl(pBmpDetailDict.Item("MaxWeirHeight"))
            weirHStep = CDbl(pBmpDetailDict.Item("WeirHeightIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "WeirH" & vbTab & _
                                            weirHMin & vbTab & _
                                            weirHMax & vbTab & _
                                            weirHStep & vbNewLine
        End If
    End If
    
    If pBmpDetailDict.Exists("SoilDOptimized") Then
        If (CBool(pBmpDetailDict.Item("SoilDOptimized"))) Then
            soilDMin = CDbl(pBmpDetailDict.Item("MinSoilDepth"))
            soilDMax = CDbl(pBmpDetailDict.Item("MaxSoilDepth"))
            soilDStep = CDbl(pBmpDetailDict.Item("SoilDepthIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "SDepth" & vbTab & _
                                            soilDMin & vbTab & _
                                            soilDMax & vbTab & _
                                            soilDStep & vbNewLine
        End If
     End If
     
    'Additional parameters -- Sabu Paul
    Dim depthMin As Double
    Dim depthMax As Double
    Dim depthStep As Double
    If pBmpDetailDict.Exists("BLengthBOptimized") Then
        If (CBool(pBmpDetailDict.Item("BLengthBOptimized"))) Then
            lengthMin = CDbl(pBmpDetailDict.Item("MinBasinBLength"))
            lengthMax = CDbl(pBmpDetailDict.Item("MaxBasinBLength"))
            lengthStep = CDbl(pBmpDetailDict.Item("BasinBLengthIncr"))
           
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "Length" & vbTab & _
                                            lengthMin & vbTab & _
                                            lengthMax & vbTab & _
                                            lengthStep & vbNewLine
        End If
    End If
    
    If pBmpDetailDict.Exists("BWidthBOptimized") Then
        If (CBool(pBmpDetailDict.Item("BWidthBOptimized"))) Then
            widthMin = CDbl(pBmpDetailDict.Item("MinBasinBWidth"))
            widthMax = CDbl(pBmpDetailDict.Item("MaxBasinBWidth"))
            widthStep = CDbl(pBmpDetailDict.Item("BasinBWidthIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "Width" & vbTab & _
                                            widthMin & vbTab & _
                                            widthMax & vbTab & _
                                            widthStep & vbNewLine
        End If
    End If
    If pBmpDetailDict.Exists("BDepthBOptimized") Then
        If (CBool(pBmpDetailDict.Item("BDepthBOptimized"))) Then
            depthMin = CDbl(pBmpDetailDict.Item("MinBasinBDepth"))
            depthMax = CDbl(pBmpDetailDict.Item("MaxBasinBDepth"))
            depthStep = CDbl(pBmpDetailDict.Item("BasinBDepthIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "Width" & vbTab & _
                                            depthMin & vbTab & _
                                            depthMax & vbTab & _
                                            depthStep & vbNewLine
        End If
    End If
    
    Dim numMin As Double
    Dim numMax As Double
    Dim numStep As Double
    
    If pBmpDetailDict.Exists("NumUnitsOptimized") Then
        If (CBool(pBmpDetailDict.Item("NumUnitsOptimized"))) Then
            numMin = CDbl(pBmpDetailDict.Item("MinNumUnits"))
            numMax = CDbl(pBmpDetailDict.Item("MaxNumUnits"))
            numStep = CDbl(pBmpDetailDict.Item("NumUnitsIncr"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "NUMUNIT" & vbTab & _
                                            numMin & vbTab & _
                                            numMax & vbTab & _
                                            numStep & vbNewLine
        End If
    End If
    
    If pBmpDetailDict.Exists("NumUnitsOptimizedB") Then
        If (CBool(pBmpDetailDict.Item("NumUnitsOptimizedB"))) Then
            numMin = CDbl(pBmpDetailDict.Item("MinNumUnitsB"))
            numMax = CDbl(pBmpDetailDict.Item("MaxNumUnitsB"))
            numStep = CDbl(pBmpDetailDict.Item("NumUnitsIncrB"))
            
            StrCard810AdjustParameter = StrCard810AdjustParameter & _
                                            strBMPindex & vbTab & _
                                            "NUMUNIT" & vbTab & _
                                            numMin & vbTab & _
                                            numMax & vbTab & _
                                            numStep & vbNewLine
        End If
    End If
               
     GoTo CleanUp
ShowError:
    MsgBox "GetOptimizationParameters: " & Err.description
CleanUp:
End Sub


'******************************************************************************
'Subroutine: GetGrowthIndex
'Author:     Mira Chokshi
'Purpose:    Get the growth index parameters for each bmp. Growth index varies
'            each month, hence has monthly values.
'*****************************************************************************
Private Sub GetGrowthIndex(strBMPindex As String) 'iBMPindex As Integer)
On Error GoTo ShowError
    Dim strGrowthIndex As String    'Growth Index
    Dim iGrowthIndex() As String
    Dim i As Integer
    
    strGrowthIndex = CStr(pBmpDetailDict.Item("GrowthIndex"))
    iGrowthIndex = Split(strGrowthIndex, ";")
    
    If (UBound(iGrowthIndex) > LBound(iGrowthIndex)) Then
        StrCard745BMPGrowthIndex = StrCard745BMPGrowthIndex & strBMPindex & vbTab
        For i = LBound(iGrowthIndex) To UBound(iGrowthIndex)
            StrCard745BMPGrowthIndex = StrCard745BMPGrowthIndex & iGrowthIndex(i) & vbTab
        Next
        StrCard745BMPGrowthIndex = StrCard745BMPGrowthIndex & vbNewLine
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "GetGrowthIndex: " & Err.description
CleanUp:
    
End Sub



'******************************************************************************
'Subroutine: GetLandTypeRouting
'Author:     Mira Chokshi
'Purpose:    Call the subroutine to summarize the landuse area over the subbasin.
'            It returns a dictionary with keys as subbasin id, and values as
'            dictionary of landuse code and area mapping. Determine the bmps each
'            landuse flows into and output to a string variable.
'*****************************************************************************
Private Sub GetLandTypeRouting()
On Error GoTo ShowError
    StrCard790LandTypeRouting = ""
    Set gBMPDrainAreaDict = New Scripting.Dictionary
    
    'Call the function to do landuse reclassification for subwatershed
    Call ComputeLanduseAreaForEachSubBasin
    
    'Call the function to get the basin to bmp routing and save into a dictionary
    Call DefineBasinToBMPDictionary
    
    Dim pTable As iTable
    Dim pTableName As String
    If gExternalSimulation Then
        pTableName = "TSAssigns"
    Else
        pTableName = "LUReclass"
    End If
    
    Set pTable = GetInputDataTable(pTableName) '"LUReclass")
    If (pTable Is Nothing) Then
        MsgBox pTableName & " table not found."
        Exit Sub
    End If
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    Dim pLUCode As Integer
    Dim iLuCode As Long
    iLuCode = pTable.FindField("LUCode")
    
    Dim pLanduseDict As Scripting.Dictionary
    Set pLanduseDict = CreateObject("Scripting.Dictionary")
    'Get subwater and landtype dictionary
    Dim pSubWaterKeys
    pSubWaterKeys = gSubWaterLandUseDict.keys
    Dim pLandTypeKeys
    Dim iGroupArea As Double
    Dim iXpos As Long
    Dim iYpos As Long
    Dim idsBMP As Integer
    Dim iCount As Integer
    iCount = 1
    Dim iGroup, iSubWater, iLandType As Integer
    Dim iDrainageBasin As Integer
    Dim pLandUseAreaDict As Scripting.Dictionary
    
    'Get the pervious/impervious percentage - Sabu Paul, Sep 17, 2004
    Dim iPercentageCode As Long
    iPercentageCode = pTable.FindField("Percentage")
    Dim pPercentage As Double
    Dim pLUAreaPerGroup As Double
    
    Dim pAggBmpLuDict As Scripting.Dictionary
    Dim pAggBMPArea As Double
    Dim pAggLuArea As Double
'    Dim agg_bmp_cat_Ids, catIndex As Integer
'    agg_bmp_cat_Ids = Array("A", "B", "C", "D", "E")
    Dim bmpCardStr As String
    Dim catIndex As Integer

''    'Loop over each landtype group
''    For iGroup = 1 To pMaxLandTypeGroupID
''        'Get each subwater, find the area for each land type
''        For iSubWater = 0 To gSubWaterLandUseDict.Count - 1
''            iDrainageBasin = pSubWaterKeys(iSubWater)
''
''            Set pLandUseAreaDict = Nothing
''            Set pLandUseAreaDict = gSubWaterLandUseDict.Item(iDrainageBasin)
''
''            'Get the draining BMP from the basin to bmp dictionary
''            idsBMP = pBasinToBMPDict.Item(iDrainageBasin)
''
''            'Query the subwatershed for same group landuse type
''            pQueryFilter.WhereClause = "LUGroupID = " & iGroup
''            Set pCursor = pTable.Search(pQueryFilter, True)
''            Set pRow = pCursor.NextRow
''            'Find all landtype subgroup and sum their area
''            iGroupArea = 0
''            Do While Not pRow Is Nothing
''                pLUCode = pRow.value(iLUCode)
''                pPercentage = pRow.value(iPercentageCode)
''                pLandTypeKeys = pLandUseAreaDict.keys
''                For iLandType = 0 To pLandUseAreaDict.Count - 1
''                    If (pLandTypeKeys(iLandType) = pLUCode) Then
''                        pLUAreaPerGroup = 0
''                        If (pLandUseAreaDict.Exists(pLandTypeKeys(iLandType))) Then
''                             pLUAreaPerGroup = pLandUseAreaDict.Item(pLandTypeKeys(iLandType))
''                        End If
''                        iGroupArea = iGroupArea + ((pLUAreaPerGroup * pPercentage) / 4046.856) ' area in acres -- Sabu Paul, Aug 24, 2004
''                    End If
''                Next
''                'Go to next landtype
''                Set pRow = pCursor.NextRow
''            Loop
''            Set pRow = Nothing
''            Set pCursor = Nothing
''            If (iGroupArea > 0) Then
''                StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
''                                            Format(iGroupArea, "#.##") & vbTab & idsBMP & vbNewLine
''                iCount = iCount + 1
''            End If
''
''            'Check to see if it is aggregage BMP
''
''            'Check for Aggregate BMP distribution
''            Set pAggBmpLuDict = Get_Agg_BMP_Lu_Distrib(idsBMP)
''
''            gBMPDrainAreaDict.Item(idsBMP) = gBMPDrainAreaDict.Item(idsBMP) + CDbl(pHruDictTmp.Item(iHru))
''
''        Next
''    Next
    
    Dim pSQAcreFactor As Double
    If gMetersPerUnit = 0# Then GetMetersPerLinearUnit
    pSQAcreFactor = gMetersPerUnit * gMetersPerUnit * 0.0002471044       'sq meter to acre conversion

    Dim strAggBMPIds As String
    Dim listAggBmpIds
    Dim strLuDist As String, listLuDist
    
    Dim lineNumber As String
    
    'Get each subwater, find the area for each land type
    For iSubWater = 0 To gSubWaterLandUseDict.Count - 1
        lineNumber = "SubWater = " & iSubWater & " 1 "
        
        iDrainageBasin = pSubWaterKeys(iSubWater)
            
        Set pLandUseAreaDict = Nothing
        Set pLandUseAreaDict = gSubWaterLandUseDict.Item(iDrainageBasin)
        
        'Get the draining BMP from the basin to bmp dictionary
        idsBMP = pBasinToBMPDict.Item(iDrainageBasin)
        
        Set pAggBmpLuDict = Nothing
        'Check to see if it is aggregage BMP
        If gAggBMPFlagDict.Exists(idsBMP) Then
            If gAggBMPFlagDict.Item(idsBMP) Then
                'Check for Aggregate BMP distribution
                Set pAggBmpLuDict = Get_Agg_BMP_Lu_Distrib(idsBMP)
                If pAggBmpLuDict.Exists(-98) Then
                    strAggBMPIds = pAggBmpLuDict.Item(-98)(1)
                    listAggBmpIds = Split(strAggBMPIds, ",")
                End If
            End If
        End If
        
        lineNumber = "SubWater = " & iSubWater & " 2 "
        'Loop over each landtype group
        For iGroup = 1 To pMaxLandTypeGroupID
                                
            'Query the subwatershed for same group landuse type
            pQueryFilter.WhereClause = "LUGroupID = " & iGroup
            Set pCursor = pTable.Search(pQueryFilter, True)
            Set pRow = pCursor.NextRow
            'Find all landtype subgroup and sum their area
            iGroupArea = 0
            Do While Not pRow Is Nothing
                pLUCode = pRow.value(iLuCode)
                pPercentage = pRow.value(iPercentageCode)
                pLandTypeKeys = pLandUseAreaDict.keys
                For iLandType = 0 To pLandUseAreaDict.Count - 1
                    If (pLandTypeKeys(iLandType) = pLUCode) Then
                        pLUAreaPerGroup = 0
                        If (pLandUseAreaDict.Exists(pLandTypeKeys(iLandType))) Then
                             pLUAreaPerGroup = pLandUseAreaDict.Item(pLandTypeKeys(iLandType))
                        End If
                        iGroupArea = iGroupArea + ((pLUAreaPerGroup * pPercentage) * pSQAcreFactor) ' area in acres -- Sabu Paul, Aug 24, 2004
                    End If
                Next
                'Go to next landtype
                Set pRow = pCursor.NextRow
            Loop
            Set pRow = Nothing
            Set pCursor = Nothing
            lineNumber = "SubWater = " & iSubWater & " 3 "
            If (iGroupArea > 0) Then
                If pAggBmpLuDict Is Nothing Then
                    lineNumber = "SubWater = " & iSubWater & " 4 Group = " & iGroup
                    'StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
                    '                            Format(iGroupArea, "#.##") & vbTab & idsBMP & vbNewLine
                    StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
                                                FormatNumber(iGroupArea, 2) & vbTab & idsBMP & vbNewLine
                    gBMPDrainAreaDict.Item(idsBMP) = gBMPDrainAreaDict.Item(idsBMP) + iGroupArea
                    lineNumber = "SubWater = " & iSubWater & " 4_1 Group = " & iGroup
                    iCount = iCount + 1
                Else
                    lineNumber = "SubWater = " & iSubWater & " 5 Group = " & iGroup
                    If pAggBmpLuDict.Exists(CInt(iGroup)) Then
                        lineNumber = "SubWater = " & iSubWater & " 5_1 Group = " & iGroup
                        'pAggBMPArea = pAggBmpLuDict.Item(CInt(iGroup))(0)
'                        For catIndex = 0 To UBound(agg_bmp_cat_Ids)
'
'                            lineNumber = "SubWater = " & iSubWater & " 5_2 Group = " & iGroup
'                            bmpCardStr = idsBMP & "_" & agg_bmp_cat_Ids(catIndex)
'                            pAggLuArea = CDbl(pAggBmpLuDict.Item(CInt(iGroup))(catIndex + 1)) / 100 * iGroupArea
'                            lineNumber = "SubWater = " & iSubWater & " 5_3 Group = " & iGroup
'                            If pAggLuArea > 0 Then
'                                lineNumber = "SubWater = " & iSubWater & " 5_4 Group = " & iGroup
'                                StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
'                                                    Format(pAggLuArea, "#.##") & vbTab & bmpCardStr & vbNewLine
'                                gBMPDrainAreaDict.Item(bmpCardStr) = gBMPDrainAreaDict.Item(bmpCardStr) + pAggLuArea
'                                lineNumber = "SubWater = " & iSubWater & " 5_5 Group = " & iGroup
'                                iCount = iCount + 1
'                            End If
'
'                        Next
                        
                        If pAggBmpLuDict.Exists(CInt(iGroup)) Then
                            strLuDist = pAggBmpLuDict.Item(CInt(iGroup))(1)
                            listLuDist = Split(strLuDist, ",")
                            For catIndex = 0 To UBound(listLuDist)
                                If CInt(listAggBmpIds(catIndex)) = 0 Then
                                    bmpCardStr = idsBMP
                                Else
                                    bmpCardStr = idsBMP & "_" & listAggBmpIds(catIndex)
                                End If
                                pAggLuArea = CDbl(listLuDist(catIndex)) / 100 * iGroupArea
                                If pAggLuArea > 0 Then
                                    lineNumber = "SubWater = " & iSubWater & " 5_4 Group = " & iGroup
'                                    StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
'                                                        Format(pAggLuArea, "#.##") & vbTab & bmpCardStr & vbNewLine
                                    StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & iGroup & vbTab & _
                                                        FormatNumber(pAggLuArea, 2) & vbTab & bmpCardStr & vbNewLine
                                    gBMPDrainAreaDict.Item(bmpCardStr) = gBMPDrainAreaDict.Item(bmpCardStr) + pAggLuArea
                                    lineNumber = "SubWater = " & iSubWater & " 5_5 Group = " & iGroup
                                    iCount = iCount + 1
                                End If
                            Next
                        End If
                    End If
                End If
                
            End If
            
            

        Next
    Next
    
    lineNumber = "SubWater = " & iSubWater & " 6 = "
    
    '*** Add additional details of external timeseries file
    Dim LUgroupCount As Integer
    LUgroupCount = pMaxLandTypeGroupID
    Set pTable = GetInputDataTable("ExternalTS")
    If Not (pTable Is Nothing) Then
        Set pCursor = pTable.Search(Nothing, True)
        Set pRow = pCursor.NextRow
        Dim iBMPID As Long
        iBMPID = pCursor.FindField("BMPID")
        Dim iMultiplier As Long
        iMultiplier = pCursor.FindField("Multiplier")
        Do While Not pRow Is Nothing
                LUgroupCount = LUgroupCount + 1
'                StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & LUgroupCount & vbTab & _
'                                            FormatNumber(0, 2) & vbTab & pRow.value(iBMPID) & vbNewLine
                StrCard790LandTypeRouting = StrCard790LandTypeRouting & iCount & vbTab & LUgroupCount & vbTab & _
                                            "0.00" & vbTab & pRow.value(iBMPID) & vbNewLine
                iCount = iCount + 1
            Set pRow = pCursor.NextRow
        Loop
    End If
    
    'GetLandTypeRouting = StrCard790LandTypeRouting
    
    GoTo CleanUp
    
ShowError:
    'MsgBox lineNumber
    gHasInFileError = True
    MsgBox "GetLandTypeRouting: " & Err.description
CleanUp:
    Set pTable = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set pQueryFilter = Nothing
    Set pLanduseDict = Nothing
    Set pSubWaterKeys = Nothing
    Set pLandTypeKeys = Nothing
    Set pLandUseAreaDict = Nothing
End Sub


'******************************************************************************
'Subroutine: DefineBasinToBMPDictionary
'Author:     Mira Chokshi
'Purpose:    Get the BMP Identifier into which a watershed flows into.
'*****************************************************************************
Private Sub DefineBasinToBMPDictionary()
On Error GoTo ShowError
    'Define variables for watershed feature layer
    Dim pWatershedFLayer As IFeatureLayer
    Set pWatershedFLayer = GetInputFeatureLayer("Watershed")
    If (pWatershedFLayer Is Nothing) Then
        MsgBox "Watershed feature layer not found."
        Exit Sub
    End If
    Dim pWatershedFClass As IFeatureClass
    Set pWatershedFClass = pWatershedFLayer.FeatureClass
    Dim pWatershedFCursor As IFeatureCursor
    Dim pWatershedFeature As IFeature
    'Define variables to access watershed feature layer fields
    Dim iBasinIDFld As Long
    iBasinIDFld = pWatershedFClass.FindField("ID")
    Dim iBmpIdFld As Long
    iBmpIdFld = pWatershedFClass.FindField("BMPID")
    'Iterate over different watershed feature layer features
    Set pWatershedFCursor = pWatershedFClass.Search(Nothing, True)
    Set pWatershedFeature = pWatershedFCursor.NextFeature
    
    'Initialize the dictionary for subbasin to bmp routing
    Set pBasinToBMPDict = CreateObject("Scripting.Dictionary")
    Do While Not pWatershedFeature Is Nothing
        'KEY = Basin ID, VALUE = Draining BMP
        pBasinToBMPDict.add pWatershedFeature.value(iBasinIDFld), pWatershedFeature.value(iBmpIdFld)
        Set pWatershedFeature = pWatershedFCursor.NextFeature
    Loop
    GoTo CleanUp

ShowError:
    MsgBox "DefineBasinToBMPDictionary: " & Err.description
CleanUp:
    Set pWatershedFLayer = Nothing
    Set pWatershedFClass = Nothing
    Set pWatershedFCursor = Nothing
    Set pWatershedFeature = Nothing
End Sub

Private Function SetAggBMPNetwork(curBmpID As Integer, usBmpId As Integer, _
                    dsBmpId As Integer, pBMPOutletType As Integer, Optional bUsOnly As Boolean) As String
On Error GoTo ShowError
    Dim pAgg_BMP_Cat_Dict As Scripting.Dictionary
'    Dim agg_bmp_cats
'    agg_bmp_cats = Array("On-Site Interception", "On-Site Treatment", "Routing Attenuation", "Regional Storage/Treatment")
'    Dim agg_bmp_cat_Ids
'    agg_bmp_cat_Ids = Array("A", "B", "C", "D", "E")
    
    
    Dim catIndex As Integer, catIndex2 As Integer
    Dim bmpCardStr As String, bmpCardStr2 As String
    Dim bDsFound As Boolean, bUsBmpSet As Boolean
    Set pAgg_BMP_Cat_Dict = Nothing
    bUsBmpSet = False
    
    Set pAgg_BMP_Cat_Dict = GetAggBMPTypes(curBmpID) 'Get_OnMap_BMP_Categories(curBmpID)
    If pAgg_BMP_Cat_Dict Is Nothing Then Exit Function
    
    Dim strAggBMPIds As String
    Dim strAggBMPDsIds As String
    Dim listAggBmpIds, listAggDsIds
    Dim strLuDist As String, listLuDist
    Dim pAggBmpLuDict As Scripting.Dictionary
    Set pAggBmpLuDict = Get_Agg_BMP_Lu_Distrib(curBmpID)
    If pAggBmpLuDict.Exists(-98) And pAggBmpLuDict.Exists(-99) Then
        strAggBMPIds = pAggBmpLuDict.Item(-98)(1)
        listAggBmpIds = Split(strAggBMPIds, ",")
        strAggBMPDsIds = pAggBmpLuDict.Item(-99)(1)
        listAggDsIds = Split(strAggBMPDsIds, ",")
    End If
        
'    Dim upstreamBMP As String
'    If pAggBmpLuDict.Exists(-97) Then
'        upstreamBMP = pAggBmpLuDict.Item(-97)(1)
'    End If
    Dim resultString As String
    resultString = ""
    'Set the routing network for aggregate BMP A-B-C-D
''    For catIndex = 0 To UBound(agg_bmp_cats)
''        bDsFound = False
''        If pAgg_BMP_Cat_Dict.Exists(CStr(agg_bmp_cats(catIndex))) Then
''            bmpCardStr = curBmpID & "_" & agg_bmp_cat_Ids(catIndex)
''
''            If Not bUsBmpSet Then
''                bUsBmpSet = True
''                If usBmpId <> 0 Then
''                    resultString = resultString & _
''                                  usBmpId & vbTab & _
''                                  pBMPOutletType & vbTab & _
''                                  bmpCardStr & vbNewLine
''                End If
''            End If
''
''            If Not bUsOnly Then
''                For catIndex2 = catIndex + 1 To UBound(agg_bmp_cats)
''                    If pAgg_BMP_Cat_Dict.Exists(CStr(agg_bmp_cats(catIndex2))) Then
''                        bmpCardStr2 = curBmpID & "_" & agg_bmp_cat_Ids(catIndex2)
''                        resultString = resultString & _
''                                          bmpCardStr & vbTab & _
''                                          pBMPOutletType & vbTab & _
''                                          bmpCardStr2 & vbNewLine
''                        bDsFound = True
''                        Exit For
''                    End If
''                Next
''
''                If Not bDsFound Then
''                    resultString = resultString & _
''                                      bmpCardStr & vbTab & _
''                                      pBMPOutletType & vbTab & _
''                                      curBmpID & "_E" & vbNewLine
''                End If
''            End If
''        End If
''    Next
    
    'if this bmp is the most upstream
'    If Not bUsBmpSet Then
'        bUsBmpSet = True
        If usBmpId <> 0 Then
            resultString = resultString & _
                          usBmpId & vbTab & _
                          pBMPOutletType & vbTab & _
                          curBmpID & vbNewLine  'curBmpID & "_0" & vbNewLine
        End If
'    End If
    
    For catIndex = 0 To UBound(listAggBmpIds) - 1
        'bmpCardStr = curBmpID & "_" & listAggBmpIds(catIndex)
        
        If Not bUsOnly Then
            resultString = resultString & _
                        curBmpID & "_" & listAggBmpIds(catIndex) & vbTab & _
                        pBMPOutletType & vbTab & _
                        curBmpID & "_" & listAggDsIds(catIndex) & vbNewLine
                        
            If CInt(listAggBmpIds(catIndex)) = 0 Then
                resultString = resultString & curBmpID
            Else
                resultString = resultString & curBmpID & "_" & listAggBmpIds(catIndex)
            End If
            
            If CInt(listAggDsIds(catIndex)) = 0 Then
                resultString = resultString & vbTab & _
                        pBMPOutletType & vbTab & _
                        curBmpID & vbNewLine
            Else
                resultString = resultString & vbTab & _
                        pBMPOutletType & vbTab & _
                        curBmpID & "_" & listAggDsIds(catIndex) & vbNewLine
            End If
            
        End If
    Next

    'bmpCardStr = curBmpID & "_E"
    bmpCardStr = curBmpID '& "_0"
'    If Not bUsBmpSet Then
'        If usBmpId <> 0 Then
'            resultString = resultString & _
'                              usBmpId & vbTab & _
'                              pBMPOutletType & vbTab & _
'                              bmpCardStr & vbNewLine
'        End If
'    End If
    If Not bUsOnly Then
        'Set the routing network for aggregate BMP E- outside
        resultString = resultString & _
                        bmpCardStr & vbTab & _
                        pBMPOutletType & vbTab & _
                        dsBmpId & vbNewLine
    End If
    
    SetAggBMPNetwork = resultString
    Exit Function
ShowError:
    MsgBox "Error in SetAggBMPNetwork:" & Err.description
End Function
'******************************************************************************
'Subroutine: GetBMPNetworkRouting
'Author:     Mira Chokshi
'Purpose:    Get the bmp network. Each bmp has three outlets. If all three flow
'            flow to same d/s bmp, the outlet_type is 1, else outlet_type is
'            2, 3, 4 with individual d/s bmp specified. This information is
'            already present in BMPNetwork table.
'*****************************************************************************
Private Function GetBMPNetworkRouting()
On Error GoTo ShowError

    Dim pTable As iTable
    Set pTable = GetInputDataTable("BMPNetwork")
    If (pTable Is Nothing) Then
        MsgBox "BMPNetwork table not found."
        Exit Function
    End If
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("Conduits")
    If (pFeatureLayer Is Nothing) Then
        MsgBox "Conduits feature layer not found."
        Exit Function
    End If
    
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(Nothing, True)
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    'Get field index
    Dim iID As Long
    iID = pCursor.FindField("ID")
    Dim iOUTLETTYPE As Long
    iOUTLETTYPE = pCursor.FindField("OutletType")
    Dim iDSID As Long
    iDSID = pCursor.FindField("DSID")
    'Get field values
    Dim StrCard795BMPNetworkRouting As String
    Dim pBMPID As Integer
    Dim pBMPdsId As Integer
    Dim pBMPOutletType As Integer
    'Iterate over bmpnetwork table to get values for bmp's flowing to downstream
    
    Dim bIsAggBMP As Boolean
    Do While Not pRow Is Nothing
        pBMPID = pRow.value(iID)
        pBMPOutletType = pRow.value(iOUTLETTYPE)
        pBMPdsId = pRow.value(iDSID)
        bIsAggBMP = False
        If gAggBMPFlagDict.Exists(pBMPID) Then
            If gAggBMPFlagDict.Item(pBMPID) Then bIsAggBMP = True
        End If
            
        If (pBMPdsId = 0) Then
            If Not bIsAggBMP Then
                StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & _
                                              pBMPID & vbTab & _
                                              pBMPOutletType & vbTab & _
                                              pBMPdsId & vbNewLine
            Else
                StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & SetAggBMPNetwork(pBMPID, 0, pBMPdsId, pBMPOutletType)
            End If
        End If
        'Move to next table record
        Set pRow = pCursor.NextRow
    Loop
    
    'Iterate over conduits feature table to get bmp-->conduit-->bmp network
    Dim pQueryFilter As IQueryFilter
    Dim pFeatureclass As IFeatureClass
    Set pFeatureclass = pFeatureLayer.FeatureClass
    Dim pFeatureCursor As IFeatureCursor
    Set pFeatureCursor = pFeatureclass.Search(Nothing, True)
    Dim pFeature As IFeature
    Dim iFROM As Long
    iFROM = pFeatureclass.FindField("CFROM")
    Dim iTO As Long
    iTO = pFeatureclass.FindField("CTO")
    Dim iCID As Long
    iCID = pFeatureclass.FindField("ID")
    Dim iOUTTYPE As Long
    iOUTTYPE = pFeatureclass.FindField("OUTLETTYPE")
    Dim pFROM As Integer
    Dim pTO As Integer
    Dim pCID As Integer
    Dim pOUTLETType As Integer
    
    'Loop over all features
    Set pFeature = pFeatureCursor.NextFeature
    Do While Not pFeature Is Nothing
        pFROM = pFeature.value(iFROM)
        pTO = pFeature.value(iTO)
        pCID = pFeature.value(iCID)
        pOUTLETType = pFeature.value(iOUTTYPE)
        
        bIsAggBMP = False
        If gAggBMPFlagDict.Exists(pFROM) Then
            If gAggBMPFlagDict.Item(pFROM) Then bIsAggBMP = True
        End If
        
        'First line for u/s bmp to conduit, next line for conduit to d/s bmp
        If Not bIsAggBMP Then
            StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & _
                                          pFROM & vbTab & _
                                          pOUTLETType & vbTab & _
                                          pCID & vbNewLine
        Else
            StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & SetAggBMPNetwork(pFROM, 0, pCID, pOUTLETType)
        End If
        
        bIsAggBMP = False
        If gAggBMPFlagDict.Exists(pTO) Then
            If gAggBMPFlagDict.Item(pTO) Then bIsAggBMP = True
        End If
        
        pOUTLETType = 3  'Channel, since conduit is a channel
        If Not bIsAggBMP Then
            'Line for conduit to d/s bmp
            StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & _
                                          pCID & vbTab & _
                                          pOUTLETType & vbTab & _
                                          pTO & vbNewLine
        Else
            StrCard795BMPNetworkRouting = StrCard795BMPNetworkRouting & SetAggBMPNetwork(pTO, pCID, 0, pBMPOutletType, True)
        End If
        'Move to next conduit
        Set pFeature = pFeatureCursor.NextFeature
    Loop
    
    
    'Return the bmp network routing
    GetBMPNetworkRouting = StrCard795BMPNetworkRouting
GoTo CleanUp
    
ShowError:
    MsgBox "GetBMPNetworkRouting: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
End Function


'******************************************************************************
'Subroutine: GetAssessPointOptimizationDetails
'Author:     Mira Chokshi
'Purpose:    Get details for assessment point - optimization parameters
'*****************************************************************************
Private Function GetAssessPointOptimizationDetails()
On Error GoTo ShowError

    Dim pTable As iTable
    Set pTable = GetInputDataTable("OptimizationDetail")
    If (pTable Is Nothing) Then
        MsgBox "OptimizationDetail table not found."
        Exit Function
    End If
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "ID = 0"
    Dim pCursor As ICursor
    Set pCursor = pTable.Search(pQueryFilter, True)

    'Get field index
    Dim iID As Long
    iID = pCursor.FindField("ID")
    Dim iPropName As Long
    iPropName = pCursor.FindField("PropName")
    Dim iPropValue As Long
    iPropValue = pCursor.FindField("PropValue")
    Dim pRow As iRow
    Set pRow = pCursor.NextRow
    
    '*** Option header
    'StrCard800OptimizationControls = "c Technique Option    Cost Limit($)   StopDelta   MaxRunTime(hr)  NumBest      NumBreak" & vbNewLine
    StrCard800OptimizationControls = "c Technique Option    StopDelta   MaxRunTime(hr)  NumBest" & vbNewLine
    'StrCard815Assess = "c Option    Cost Limit($)   StopDelta   MaxRunTime(hr)  NumBest" & vbNewLine
    Dim pTechniqueVal As Integer 'New parameter - June 18, 2007
    Dim pOptionVal As Integer
'    Dim pCostLimit As Double
    Dim pStopDelta As Double
    Dim pMaxRunTime As Double
    Dim pNumBest As Integer
'    Dim pNumBreak As Integer
'    pNumBreak = 0
    
    Do While Not pRow Is Nothing
        Select Case (pRow.value(iPropName))
            Case "Technique":
                    pTechniqueVal = CDbl(pRow.value(iPropValue))
            Case "Option":
                    pOptionVal = CInt(pRow.value(iPropValue))
'            Case "CostLimit":
'                    pCostLimit = CDbl(pRow.value(iPropValue))
            Case "StopDelta":
                    pStopDelta = CDbl(pRow.value(iPropValue))
            Case "MaxRunTime":
                    pMaxRunTime = CDbl(pRow.value(iPropValue))
            Case "NumBest":
                    pNumBest = CDbl(pRow.value(iPropValue))
'            Case "NumBreak":
'                    pNumBreak = CDbl(pRow.value(iPropValue))
        End Select
        Set pRow = pCursor.NextRow
    Loop
'    StrCard800OptimizationControls = StrCard800OptimizationControls & pTechniqueVal & vbTab & _
'                                        pOptionVal & vbTab & _
'                                        pCostLimit & vbTab & _
'                                        pStopDelta & vbTab & _
'                                        pMaxRunTime & vbTab & _
'                                        pNumBest & vbTab & _
'                                        pNumBreak & vbNewLine
    
    StrCard800OptimizationControls = StrCard800OptimizationControls & pTechniqueVal & vbTab & _
                                        pOptionVal & vbTab & _
                                        pStopDelta & vbTab & _
                                        pMaxRunTime & vbTab & _
                                        pNumBest & vbNewLine
    '*** BMP Assessment optimization header
    'StrCard815Assess = StrCard815Assess & "c   BMPSITE     FactorGroup     FactorType      CalcDays    CalcMode    TargetValue    FactorName" & vbNewLine
    If pOptionVal = 1 Then
        StrCard815Assess = "c   BMPSITE     FactorGroup     FactorType      CalcDays    CalcMode    TargetValue    FactorName" & vbNewLine
    Else
        StrCard815Assess = "c   BMPSITE     FactorGroup     FactorType      CalcDays    CalcMode    TargetVal_Low TargetVal_Up    FactorName" & vbNewLine
    End If
    
    pQueryFilter.WhereClause = "ID <> 0"
    Set pCursor = pTable.Search(pQueryFilter, True)
    Set pRow = pCursor.NextRow
    Dim pIDString As String
    Dim pParamString As String
    Do While Not (pRow Is Nothing)
        pIDString = pRow.value(iID) & vbTab
        pParamString = pRow.value(iPropValue)
        pParamString = Replace(pParamString, ",", vbTab)
        StrCard815Assess = StrCard815Assess & pIDString & pParamString & vbNewLine
        Set pRow = pCursor.NextRow
    Loop
GoTo CleanUp
    
ShowError:
    MsgBox "GetAssessPointOptimizationDetails: " & Err.description
CleanUp:
    Set pRow = Nothing
    Set pCursor = Nothing
    Set pTable = Nothing
End Function

Private Sub WriteCard900()
On Error GoTo ShowError
    pFile.WriteLine ("c900 Vegetative Filter Strip Simulation control card")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c N             = Number of nodes in the domain(must be an odd number for a quadratic finite element solution")
    pFile.WriteLine ("c THETAW        = Time-weight factor for the Crank-Nicholson solution (0.5 recommended)")
    pFile.WriteLine ("c CR            = Courant number for the calculation of time step from 0.5 - 0.8 (recommended).")
    pFile.WriteLine ("c MAXITER       = Maximum number of iterations alowed in the Picard loop (integer).")
    pFile.WriteLine ("c NPOL          = Number of nodal points over each element")
    pFile.WriteLine ("c IELOUT        = Flag to output elemental information - (1-Feeeback, 0-No Feedback - SUSTAIN Always -0 )")
    pFile.WriteLine ("c KPG           = Flag to choose the Petrov-Galerkin solution (1) or regular finite element (0)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c N THETAW CR MAXITER NPOL IELOUT KPG")
    pFile.Write GetVFSSimulationOptions
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 900:" & Err.description
End Sub

Private Sub WriteCard901()
On Error GoTo ShowError
    pFile.WriteLine ("c901 Vegetative Filter Strip (VFS) Overland Flow Solution")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier")
    pFile.WriteLine ("c LABEL         = Filter strip name for identification")
    pFile.WriteLine ("c FWIDTH        = Width of the filter strip (ft)")
    pFile.WriteLine ("c VL            = Lenght of the filter strip (ft)")
    pFile.WriteLine ("c NPROP         = Number of segments with different surface properties (slope or roughness)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID LABEL FWIDTH VL NPROP")
    pFile.Write StrCard901_VFS_Dim
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 901:" & Err.description
End Sub

Private Sub WriteCard902()
On Error GoTo ShowError
    pFile.WriteLine ("c902 VFS Segment Details")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID           = Unique Filter strip identifier")
    pFile.WriteLine ("c SegID           = Segment identifier")
    pFile.WriteLine ("c SX(SegID)       = Horizontal distance from the beginning on the filter (ft)")
    pFile.WriteLine ("c RNA(SegID)      = Manning’s roughness for each segment (s.ft^(-1/3))")
    pFile.WriteLine ("c SOA(SegID)      = Slope at each segment")
    pFile.WriteLine ("c                  SX, RNA, SOA will be for SegID = 1 to NPROP")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID SEGID SX(SegID) RNA(SegID) SOA(SegID)")
    pFile.Write StrCard902_VFS_SegDetails
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 902:" & Err.description
End Sub

Private Sub WriteCard903()
On Error GoTo ShowError
    pFile.WriteLine ("c903 VFS Soil Properties")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID           = Unique Filter strip identifier ")
    pFile.WriteLine ("c VKS             = Saturated hydraulic conductivity (in/hr)")
    pFile.WriteLine ("c SAV             = Green-Ampt’s average suction at wet front(ft)")
    pFile.WriteLine ("c OS              = Saturated soil-water content(ft^3/ft^3)")
    pFile.WriteLine ("c OI              = Initial soil-water content (ft^3/ft^3)")
    pFile.WriteLine ("c SM              = Maximum surface storage (ft)")
    pFile.WriteLine ("c SCHK            = Relative distance from de upper filter edge where the check for ponding conditions is made (i.e. 1= end filter, 0.5= mid point, 0= beginning)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID VKS SAV OS OI SM SCHK ")
    pFile.Write StrCard903_VFS_SoilProps
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 903:" & Err.description
End Sub

Private Sub WriteCard904()
On Error GoTo ShowError
    pFile.WriteLine ("c904 VFS Buffer Properties for Sediment Filtration Model")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier ")
    pFile.WriteLine ("c SS            = Spacing of the filter media elements (in)")
    pFile.WriteLine ("c VN            = Filter media Manning's n (0.012 for cylindrical media) (s.in^(-1/3))")
    pFile.WriteLine ("c H             = Filter media height (ft)")
    pFile.WriteLine ("c VN2           = Bare surface Manning's n for sediment inundated area and overland flow (s.ft^(-1/3))")
    pFile.WriteLine ("c ICO           = Flag to feedback the change in slope and surface roughness at the sediment wedge for each time step (0= no feedback; 1= feedback - SUSTAIN always 0)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.Write StrCard904_VFS_Buf_Sed
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 904:" & Err.description
End Sub

Private Sub WriteCard905()
On Error GoTo ShowError
    pFile.WriteLine ("c905 VFS Sediment Properties for Sediment Filtration Model")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier ")
    pFile.WriteLine ("c NPARTi        = Incoming sediment particle class according to the USDA (1975) particle classes for 3 soil size classes(where i=1,2,3)")
    pFile.WriteLine ("c COARSEi       = % of particles from incoming sediment with diameter > 0.0037 cm (coarse fraction that will be routed through wedge) (unit fraction, i.e. 100% = 1.0) for 3 soil size classes(where i=1,2,3)")
    pFile.WriteLine ("c PORi          = Porosity of deposited sediment (unit fraction, i.e. 43.4% = 0.434) for 3 soil size classes(where i=1,2,3)")
    pFile.WriteLine ("c DPi           = Sediment particle size, diameter(in), required only if NPART=7 for 3 soil size classes(where i=1,2,3)")
    pFile.WriteLine ("c SGi           = Sediment particle density (lb/ft^3), required only if NPART=7 for 3 soil size classes(where i=1,2,3)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.Write StrCard905_VFS_Sed_Filt
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 905:" & Err.description
End Sub

Private Sub WriteCard906()
On Error GoTo ShowError

    pFile.WriteLine ("c906 VFS Pollutant Sediment Fraction for the pollutants")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier")
    pFile.WriteLine ("c QUALSEDFRACi  = Sediment fraction for pollutant i (hr^-1)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID  QUALSEDFRAC1  QUALSEDFRAC2 ... QUALSEDFRACN")
    pFile.Write StrCard906_VFS_Sed_Fracion
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 906:" & Err.description
End Sub

Private Sub WriteCard907()
On Error GoTo ShowError
    pFile.WriteLine ("c907 VFS Pollutant First Order Decay Factors in the adsorbed fraction")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier ")
    pFile.WriteLine ("c QUALDECAYi    = First-order decay rate for pollutant i (hr^-1)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID  QUALDECAY1  QUALDECAY2 ... QUALDECAYN")
    pFile.Write StrCard907_VFS_FO_Adsorbed
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 907:" & Err.description
End Sub
Private Sub WriteCard908()
On Error GoTo ShowError
    pFile.WriteLine ("c908 VFS Pollutant First Order Decay Factors in the dissolved fraction")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID         = Unique Filter strip identifier ")
    pFile.WriteLine ("c QUALDECAYi    = First-order decay rate for pollutant i (hr^-1)")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID  QUALDECAY1  QUALDECAY2 ... QUALDECAYN")
    pFile.Write StrCard908_VFS_FO_Dissolved
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 908:" & Err.description
End Sub
Private Sub WriteCard909()
On Error GoTo ShowError
    pFile.WriteLine ("c909 VFS Pollutant Temperature Correction Factors in the adsorbed fraction")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID          = Unique Filter strip identifier ")
    pFile.WriteLine ("c TEMPCORRi      = Temperature correction for pollutant i ")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID  TEMPCORR1  TEMPCORR2 ... TEMPCORRN")
    pFile.Write StrCard909_VFS_TC_Adsorbed
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 909:" & Err.description
End Sub
Private Sub WriteCard910()
On Error GoTo ShowError
    pFile.WriteLine ("c910 VFS Pollutant Temperature Correction Factors in the dissolved fraction")
    pFile.WriteLine ("c")
    pFile.WriteLine ("c VFSID           = Unique Filter strip identifier ")
    pFile.WriteLine ("c TEMPCORRi       = Temperature correction for pollutant i ")
    pFile.WriteLine ("c              Where i = 1 to N (N = Number of QUAL from TIMESERIES FILES)")
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    pFile.WriteLine ("c VFSID  TEMPCORR1  TEMPCORR2 ... TEMPCORRN")
  
    pFile.Write StrCard910_VFS_TC_Dissolved
    pFile.WriteLine ("c--------------------------------------------------------------------------------------------")
    Exit Sub
ShowError:
    gHasInFileError = True
    MsgBox "Card 910:" & Err.description
End Sub

Public Sub GetBMPDetails_Classes(ByVal iBMP As Integer, pBMPType As String, _
        pXpos As Double, pYpos As Double, pTotalDrainageArea As Double, Optional strBmpCardId As String)
    
On Error GoTo ShowError
    'ByVal pBmpDetailDict As Scripting.Dictionary, _

    If strBmpCardId = "" Then strBmpCardId = CStr(iBMP)
    
    Dim pBMPClass As String
    pBMPClass = Trim(pBmpDetailDict.Item("BMPClass"))
    Dim isAssessPoint As Boolean
    isAssessPoint = CBool(pBmpDetailDict.Item("isAssessmentPoint"))
    'Get BMP A/B Dimension Parameters
    If (pBMPClass = "A") Then
        GetBMPParametersA strBmpCardId, pXpos, pYpos, pTotalDrainageArea
        'Get BMP Decay Factors for bmp's
        GetBMPDecayFactors strBmpCardId
        'Get BMP Percent Removal Factor
        GetBMPPercentRemovalFactors strBmpCardId
        
    ElseIf (pBMPClass = "B") Then
        GetBMPParametersB strBmpCardId, pXpos, pYpos, pTotalDrainageArea
        'Get BMP Decay Factors for bmp's
        GetBMPDecayFactors strBmpCardId
        'Get BMP Percent Removal Factor
        GetBMPPercentRemovalFactors strBmpCardId
    'Get Dummy BMP Dimension Parameters
    ElseIf (pBMPClass = "X") Then
        GetDummyBMPParameters strBmpCardId, pXpos, pYpos, pTotalDrainageArea
    End If
    If (pBMPClass <> "X" And pBMPType <> "VirtualOutlet") Then
        GetSoilIndex strBmpCardId     'Get soil properties
        GetGrowthIndex strBmpCardId  'Get growth index
        GetBMPCosts strBmpCardId      'Get the Cost parameters
        GetOptimizationParameters strBmpCardId 'iBMP 'Get Adjustable parameter range and increment
        GetBMPSedimentParameters strBmpCardId
    End If
    Exit Sub
    
ShowError:
    gHasInFileError = True
    MsgBox "Error in GetBMPDetails_Classes: " & Err.description
End Sub


Public Sub GetBMPSedimentParameters(iBMPIndex As String)
On Error GoTo ShowError
    Dim bedWidth As Double
    Dim bedDepth As Double
    Dim bedPorosity As Double
    Dim SAND_FRAC As Double
    Dim SILT_FRAC As Double
    Dim CLAY_FRAC As Double
    Dim sandDiameter As Double
    Dim sandVelocity As Double
    Dim sandDensity As Double
    Dim sandCoeff As Double
    Dim sandExponent As Double
    Dim siltDiameter As Double
    Dim siltVelocity As Double
    Dim siltDensity As Double
    Dim siltTaucd As Double
    Dim siltTaucs As Double
    Dim siltM As Double
   
    bedWidth = pBmpDetailDict.Item("Bed width")
    bedDepth = pBmpDetailDict.Item("Bed depth")
    bedPorosity = pBmpDetailDict.Item("Porosity")
    SAND_FRAC = pBmpDetailDict.Item("Sand fraction")
    SILT_FRAC = pBmpDetailDict.Item("Silt fraction")
    CLAY_FRAC = pBmpDetailDict.Item("Clay fraction")
    sandDiameter = pBmpDetailDict.Item("Sand effective diameter")
    sandVelocity = pBmpDetailDict.Item("Sand velocity")
    sandDensity = pBmpDetailDict.Item("Sand density")
    sandCoeff = pBmpDetailDict.Item("Sand coefficient")
    sandExponent = pBmpDetailDict.Item("Sand exponent")
    siltDiameter = pBmpDetailDict.Item("Silt effective diameter")
    siltVelocity = pBmpDetailDict.Item("Silt velocity")
    siltDensity = pBmpDetailDict.Item("Silt density")
    siltTaucd = pBmpDetailDict.Item("Silt Deposition stress")
    siltTaucs = pBmpDetailDict.Item("Silt Scour stress")
    siltM = pBmpDetailDict.Item("Silt Erodibility")
    
    Dim clayDiameter As Double
    Dim clayVelocity As Double
    Dim clayDensity As Double
    Dim clayTaucd As Double
    Dim clayTaucs As Double
    Dim clayM As Double
    
    clayDiameter = pBmpDetailDict.Item("Clay effective diameter")
    clayVelocity = pBmpDetailDict.Item("Clay velocity")
    clayDensity = pBmpDetailDict.Item("Clay density")
    clayTaucd = pBmpDetailDict.Item("Clay Deposition stress")
    clayTaucs = pBmpDetailDict.Item("Clay Scour stress")
    clayM = pBmpDetailDict.Item("Clay Erodibility")
    
    StrCard775Sediment = StrCard775Sediment & _
                        iBMPIndex & vbTab & _
                        bedWidth & vbTab & _
                        bedDepth & vbTab & _
                        bedPorosity & vbTab & _
                        SAND_FRAC & vbTab & _
                        SILT_FRAC & vbTab & _
                        CLAY_FRAC & vbNewLine
                        
    StrCard780SandTransport = StrCard780SandTransport & _
                        iBMPIndex & vbTab & _
                        sandDiameter & vbTab & _
                        sandVelocity & vbTab & _
                        sandDensity & vbTab & _
                        sandCoeff & vbTab & _
                        sandExponent & vbNewLine
    StrCard785SiltTransport = StrCard785SiltTransport & _
                        iBMPIndex & vbTab & _
                        siltDiameter & vbTab & _
                        siltVelocity & vbTab & _
                        siltDensity & vbTab & _
                        siltTaucd & vbTab & _
                        siltTaucs & vbTab & _
                        siltM & vbNewLine
    StrCard786ClayTransport = StrCard786ClayTransport & _
                        iBMPIndex & vbTab & _
                        clayDiameter & vbTab & _
                        clayVelocity & vbTab & _
                        clayDensity & vbTab & _
                        clayTaucd & vbTab & _
                        clayTaucs & vbTab & _
                        clayM & vbNewLine
    Exit Sub
ShowError:
    MsgBox "Error in GetBMPSedimentParameters: " & Err.description
End Sub
