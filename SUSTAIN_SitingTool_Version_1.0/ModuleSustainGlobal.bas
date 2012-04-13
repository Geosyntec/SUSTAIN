Attribute VB_Name = "ModuleSustainGlobal"
'******************************************************************************
'   Application: Sustain - BMP Siting Tool
'   Company:     Tetra Tech, Inc
'******************************************************************************


Option Explicit
Option Base 0


'*** Global Variables
Public gMxDoc As IMxDocument
Public gMap As IMap
Public gApplication As IApplication
Public gApplicationPath As String

'
Public gDataValid As Boolean
Public gBMPCriteriaDictionary As Scripting.Dictionary
Public gLayerNameDictionary As Scripting.Dictionary
Public gBMPtypeDict As Scripting.Dictionary
Public gBMPSelDict As Scripting.Dictionary
Public gWorkingfolder As String
Public gRasterfolder As String
Public gDACriteria As String

'
Public gDEMdata As String
Public gSoildata As String
Public gImperviousdata As String
Public gLandusedata As String
Public gStreamdata As String
Public gRoaddata As String
Public gMRLCdata As String
Public gWTdata As String
Public gSoilTable As String
Public gMrlcTable As String

'
Public gHydrologyOp As IHydrologyOp
Public gMapAlgebraOp As IMapAlgebraOp
Public gCellSize As Double

' Always on Top.............
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\ModuleSustainGlobal.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

Public Function CheckMapDocumentSavedStatus_ST() As Boolean
  On Error GoTo ErrorHandler


    CheckMapDocumentSavedStatus_ST = False
    
    'Create a file for writing the datasources -- Arun Raj; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = DefineApplicationPath_ST
   
    Dim dataSrcFN As String 'Arun Raj -- October 2004
    dataSrcFN = gApplication.Document
    
    'Check if the .mxd is saved, if not force the user to save the .mxd
    If (Replace(dataSrcFN, ".mxd", "") = "Untitled") Then
        MsgBox "Please save .mxd file to continue. ", vbExclamation
        Exit Function
    End If
        
    CheckMapDocumentSavedStatus_ST = True
    

  Exit Function
ErrorHandler:
  HandleError True, "CheckMapDocumentSavedStatus_ST " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function


'******************************************************************************
'Subroutine: DefineApplicationPath_ST
'Author:     Mira Chokshi
'Purpose:    Gets the path of the map and stores it in global variable
'******************************************************************************
Public Function DefineApplicationPath_ST() As String
  On Error GoTo ErrorHandler
  
  gApplicationPath = ""
  
  Dim pTemplates As ITemplates
  Dim lTempCount As Long
  Dim strDocPath As String
  
  If gApplication Is Nothing Then Exit Function
  
  Set pTemplates = gApplication.Templates
  
  Dim pDoc As IDocument
  Set pDoc = gApplication.Document
  
  If pDoc Is Nothing Then Exit Function
  
  lTempCount = pTemplates.Count
    
  'The document is always the last item
  If lTempCount > 0 Then
    strDocPath = pTemplates.Item(lTempCount - 1)
  Else
    Exit Function
    'strDocPath = pTemplates.Item(lTempCount)
  End If
 
  strDocPath = Replace(strDocPath, pDoc.Title, "")
  
  'Added the following section to make sure the path contains no .mxd - Arun Raj, August 31, 2005
  If (StringContains(strDocPath, ".mxd")) Then
    strDocPath = Replace(strDocPath, ".mxd", "")
  End If
  
  gApplicationPath = strDocPath
  
  DefineApplicationPath_ST = gApplicationPath
  Exit Function
  
ErrorHandler:

  HandleError True, "DefineApplicationPath_ST " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'*******************************************************************************
'Subroutine : StringContains
'Purpose    : Checks whether a given string is contained in another one
'Note       :
'Arguments  :
'Author     : Arun Raj
'History    :
'*******************************************************************************
Public Function StringContains(FindString As String, SearchString As String) As Boolean
  On Error GoTo ErrorHandler

    Dim TempString As String
    TempString = Replace(FindString, SearchString, "")
    If (FindString <> TempString) Then
        StringContains = True
    Else
        StringContains = False
    End If

  Exit Function
ErrorHandler:
  HandleError True, "StringContains " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'******************************************************************************
'Subroutine: SetDataDirectory_ST
'Author:     Arun Raj
'Purpose:    Check the data directory path and sets
'******************************************************************************

Public Sub SetDataDirectory_ST()
  On Error GoTo ErrorHandler

    'Create a file for writing the datasources -- Arun Raj; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = DefineApplicationPath_ST
    
    Dim fso As FileSystemObject
    ' Create a Folder for analysis data ....
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not (fso.FolderExists(pAppPath & "SitingTool")) Then
        fso.CreateFolder pAppPath & "SitingTool"
    End If
    gWorkingfolder = pAppPath & "SitingTool" ' Store into Global....
    gRasterfolder = pAppPath & "Cache" ' Store into Global....
       
    Dim dataSrcFN As String 'Arun Raj -- October 2004
    dataSrcFN = gApplication.Document
       
    '*** Get the complete path of the application
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_Siting.src"
    dataSrcFN = pAppPath & dataSrcFN
 
    Set gLayerNameDictionary = CreateObject("Scripting.Dictionary")
    gLayerNameDictionary.RemoveAll
    
    If Not (fso.FileExists(dataSrcFN)) Then
        Exit Sub
    End If
    
    Dim pDataSrcFile
    Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForReading)
    
    Dim pStrDataSrc As String
    pStrDataSrc = pDataSrcFile.ReadAll
    Dim pDataLines
    pDataLines = Split(pStrDataSrc, vbNewLine, , vbTextCompare)
    Dim lIncr As Integer
    Dim lWords
    For lIncr = 0 To UBound(pDataLines) - 1
        lWords = Split(pDataLines(lIncr), vbTab, , vbTextCompare)
        If Not gLayerNameDictionary.Exists(lWords(0)) Then
            gLayerNameDictionary.Add lWords(0), lWords(1)
        End If
    Next lIncr
    pDataSrcFile.Close
    
    '*** Get the complete path of the application
    dataSrcFN = gApplication.Document
    'dataSrcFN = Replace(dataSrcFN, ".mxd", "_criteria.src")
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_criteria.src"
    dataSrcFN = pAppPath & dataSrcFN
    ' **** Initialize the Dictionaries..........
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set gBMPCriteriaDictionary = Nothing
    Set gBMPSelDict = Nothing
    Set gBMPtypeDict = Nothing
    
    If fso.FileExists(dataSrcFN) Then
    
        Set gBMPCriteriaDictionary = New Scripting.Dictionary
        gBMPCriteriaDictionary.RemoveAll
        Set gBMPSelDict = New Scripting.Dictionary
        gBMPSelDict.RemoveAll
        
        Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForReading)
          Dim oBMP As BMPobj
          Set oBMP = New BMPobj
        
          pStrDataSrc = pDataSrcFile.ReadAll
          pDataLines = Split(pStrDataSrc, vbNewLine, , vbTextCompare)
          
          ' First get the BMP Selection list.....
          Dim strBMPSel As String
          strBMPSel = pDataLines(0)
          Dim iCnt As Integer
          iCnt = 0
          
          Do While strBMPSel <> ""
                gBMPSelDict.Add strBMPSel, strBMPSel
                iCnt = iCnt + 1
                strBMPSel = pDataLines(iCnt)
          Loop
          
          For lIncr = iCnt + 1 To UBound(pDataLines) - 1
              If pDataLines(lIncr) = "" Then
                  gBMPCriteriaDictionary.Add oBMP.BMPName, oBMP
                  iCnt = iCnt + 1
                  Set oBMP = New BMPobj
              Else
                  lWords = Split(pDataLines(lIncr), vbTab, , vbTextCompare)
                  If lWords(0) = "BMPName" Then oBMP.BMPName = lWords(1)
                  If lWords(0) = "BMPType" Then oBMP.BMPType = lWords(1)
                  If lWords(0) = "DC_BB" Then oBMP.DC_BB = lWords(1): oBMP.DC_BB_State = lWords(2)
                  If lWords(0) = "DC_DA" Then oBMP.DC_DA = lWords(1): oBMP.DC_DA_State = lWords(2)
                  If lWords(0) = "DC_DS" Then oBMP.DC_DS = lWords(1): oBMP.DC_DS_State = lWords(2)
                  If lWords(0) = "DC_HG" Then oBMP.DC_HG = lWords(1): oBMP.DC_HG_State = lWords(2)
                  If lWords(0) = "DC_IMP" Then oBMP.DC_IMP = lWords(1): oBMP.DC_IMP_State = lWords(2)
                  If lWords(0) = "DC_RB" Then oBMP.DC_RB = lWords(1): oBMP.DC_RB_State = lWords(2)
                  If lWords(0) = "DC_SB" Then oBMP.DC_SB = lWords(1): oBMP.DC_SB_State = lWords(2)
                  If lWords(0) = "DC_WT" Then oBMP.DC_WT = lWords(1): oBMP.DC_WT_State = lWords(2)
                      
              End If
          Next lIncr
          If oBMP.BMPName <> "" Then gBMPCriteriaDictionary.Add oBMP.BMPName, oBMP
          pDataSrcFile.Close
    End If
    
    'Assign the Layers
    gDEMdata = gLayerNameDictionary.Item("gDEMdata")
    gMRLCdata = gLayerNameDictionary.Item("gMRLCdata")
    gLandusedata = gLayerNameDictionary.Item("gLandusedata")
    gRoaddata = gLayerNameDictionary.Item("gRoaddata")
    gSoildata = gLayerNameDictionary.Item("gSoildata")
    gStreamdata = gLayerNameDictionary.Item("gStreamdata")
    gImperviousdata = gLayerNameDictionary.Item("gImperviousdata")
    gWTdata = gLayerNameDictionary.Item("gWTdata")
    gSoilTable = gLayerNameDictionary.Item("gSoilTable")
    gMrlcTable = gLayerNameDictionary.Item("gMrlcTable")

  Exit Sub
ErrorHandler:
  HandleError True, "SetDataDirectory_ST " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
  
End Sub


'******************************************************************************
'Subroutine: InitializeMapDocument
'Author:     Mira Chokshi
'Purpose:    Initializes Current Map, Application Path and Map Algebra, Hydrology,
'            Neighborhood and Reclassification Operators
'******************************************************************************
Public Sub InitializeMapDocument()
  On Error GoTo ErrorHandler

    Set gMxDoc = gApplication.Document
    Set gMap = gMxDoc.FocusMap
        
    'Read data layer information from src file
    Call SetDataDirectory_ST
    
  Exit Sub
ErrorHandler:
  HandleError True, "InitializeMapDocument " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


'******************************************************************************
'Subroutine: GetInputFeatureLayer
'Purpose:    Function to get feature layer from map. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetInputFeatureLayer(FLayerName As String) As ILayer
On Error GoTo ShowError
    
    Dim i As Integer
    Dim pFLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pFLayer = gMap.Layer(i)
        If ((pFLayer.Name = FLayerName) And (TypeOf pFLayer Is IFeatureLayer Or TypeOf pFLayer Is IRasterLayer) And pFLayer.Valid) Then
            Set GetInputFeatureLayer = pFLayer     'Get the plot layer
            GoTo Cleanup
        End If
    Next
    GoTo Cleanup
ShowError:
    MsgBox "GetInputFeatureLayer: " & Err.Description
Cleanup:
    Set pFLayer = Nothing
End Function

'******************************************************************************
'Subroutine: GetInputFeatureLayer
'Purpose:    Function to get feature layer from map. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetInputTable(TableName As String) As ITable
On Error GoTo ShowError
    
    Dim pStandCol As IStandaloneTableCollection
    Set pStandCol = gMap
    Dim i As Integer
    For i = 0 To pStandCol.StandaloneTableCount - 1
      If UCase(pStandCol.StandaloneTable(i).Name) = UCase(TableName) Then
          Set GetInputTable = pStandCol.StandaloneTable(i).Table
          GoTo Cleanup
      End If
    Next
    
    GoTo Cleanup
ShowError:
    MsgBox "GetInputTable: " & Err.Description
Cleanup:
    Set pStandCol = Nothing
End Function

'******************************************************************************
'Subroutine: GetFeatureLayer
'Purpose:    Function to get feature layer from Workspace. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetFeatureLayer(ByVal strWorkspace As String, ByVal FLayerName As String) As ILayer

    On Error GoTo ShowError
    
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = OpenShapeFile(strWorkspace, FLayerName)
    If pFeatClass Is Nothing Then GoTo Cleanup
    Dim pFeatLyr As IFeatureLayer
    Set pFeatLyr = New FeatureLayer
    Set pFeatLyr.FeatureClass = pFeatClass
    pFeatLyr.Name = pFeatClass.AliasName
    pFeatLyr.Visible = False
    
    Set GetFeatureLayer = pFeatLyr
    
    Exit Function
ShowError:
    MsgBox "GetFeatureLayer: " & Err.Description
Cleanup:
    Set GetFeatureLayer = Nothing
End Function


'******************************************************************************
'Subroutine: Open_Toolbox
'Author: Arun Raj
'Purpose:    General Function to get open a tool from Toolbox and execute.
'Arguments : Should provide the Tool Name, category & parameter collection.
'******************************************************************************

Public Function Open_Toolbox(ToolName As String, ToolBox As String, oColParams As Collection) As Boolean

    On Error GoTo ErrorHandler
        
        Open_Toolbox = False
        
        ' Get the ARC installation path....
        Dim strInstallPath As String
        strInstallPath = GetArcGISPath
        
        'Create a toolbox workspace factory
        Dim pToolboxWorkspaceFactory As IWorkspaceFactory
        Set pToolboxWorkspaceFactory = New ToolboxWorkspaceFactory

        'Open a toolbox workspace
        Dim pToolboxWorkspace As IToolboxWorkspace
        Set pToolboxWorkspace = pToolboxWorkspaceFactory.OpenFromFile(strInstallPath & "ArcToolbox\Toolboxes", 0)

        'Open a toolbox by Name
        Dim pGPToolbox As IGPToolbox
        Set pGPToolbox = pToolboxWorkspace.OpenToolbox(ToolBox)
        
        ' To Display the Tools in ToolBox....
        Dim pTools As IEnumGPTool
        Dim pGPTool As IGPTool
        Dim pTool As String
        Set pTools = pGPToolbox.Tools
        pTools.Reset
        Set pGPTool = pTools.Next
        Do While Not pGPTool Is Nothing
            pTool = pGPTool.DisplayName
            If pTool = ToolName Then
                Exit Do
            End If
            Set pGPTool = pTools.Next
        Loop
        
        If pGPTool Is Nothing Then Exit Function
        
        ' ************************************************
        ' ************************************************
        
        ' Now Construct the Parameters......
        Dim pParams As IArray
        Set pParams = pGPTool.ParameterInfo
        
        Dim pParameter As IGPParameter
        Dim pParamEdit As IGPParameterEdit
        Dim pDataType As IGPDataType
        Dim sValue As String
        
        ' Check if the passed parameters length is same as actual parameters.....
        If oColParams.Count <> pParams.Count Then
            MsgBox "Mismatch between parameters provided & required.", vbCritical, "BMP Siting Tool"
            Exit Function
        End If
        
        ' Construct all Parameters.....
        Dim iParams As Integer
        For iParams = 0 To pParams.Count - 1
            Set pParameter = pParams.Element(iParams)
            Set pParamEdit = pParameter
            Set pDataType = pParameter.DataType
            sValue = oColParams.Item(iParams + 1)
            Set pParamEdit.Value = pDataType.CreateValue(sValue)
        Next

        'Validate the parameters and create a IGPMessages object containing IGPMessage objects.
        Dim pGPMessages As IGPMessages
        Set pGPMessages = pGPTool.Validate(pParams, True, Nothing)

        'Create a IGPMessage object and get the error code and description.
        Dim pGPMessage As IGPMessage
        Set pGPMessage = pGPMessages.GetMessage(0)
        If pGPMessage.ErrorCode <> 0 Then
            MsgBox (pGPMessage.ErrorCode & " :" & pGPMessage.Description)
            Exit Function
        End If

        ' Execute the Tool......
        pGPTool.Execute pParams, Nothing, Nothing, pGPMessages

        Set pGPMessage = pGPMessages.GetMessage(0)
        If pGPMessage.ErrorCode <> 0 Then
            MsgBox (pGPMessage.ErrorCode & " :" & pGPMessage.Description)
        End If

        Open_Toolbox = True
        
Exit Function
ErrorHandler:
  HandleError True, "Open_Toolbox " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Function


Public Function GetArcGISPath() As String
  On Error GoTo ErrorHandler

'*********************************************************************
'***The module passes the path information for the ArcView Software***
'*********************************************************************

Dim strReadValue As String
Dim objShell As Object
Set objShell = CreateObject("WScript.Shell")
Const Reg_Key = "HKEY_LOCAL_MACHINE\SOFTWARE\ESRI\ArcGIS\InstallDir"
strReadValue = objShell.RegRead(Reg_Key)

GetArcGISPath = strReadValue
    

  Exit Function
ErrorHandler:
  HandleError True, "GetArcGISPath " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

'******************************************************************************
'Subroutine: CleanUpMemory
'Author:     Mira Chokshi
'Purpose:    Clears the memory by setting all public variables to nothing
'******************************************************************************
Public Sub CleanUpMemory()
  On Error GoTo ErrorHandler

    Set gMxDoc = Nothing
    Set gMap = Nothing

  Exit Sub
ErrorHandler:
  HandleError True, "CleanUpMemory " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Public Sub Delete_Dataset_ST(ByVal strWorkspace As String, ByVal strDataset As String)
    
    On Error GoTo ErrorHandler
    
      'Delete the FeatureClass.....
      Dim pFClass As IFeatureClass
      Set pFClass = OpenShapeFile(strWorkspace, strDataset)
      If Not pFClass Is Nothing Then
        Dim pDataset As IDataset
        Set pDataset = pFClass
        If pDataset.CanDelete Then pDataset.Delete
      End If
    
  Exit Sub
ErrorHandler:
  HandleError True, "Delete_Dataset_ST " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND


End Sub

'******************************************************************************
'Subroutine: OpenShapeFile
'Purpose:    Open feature dataset (.shp) from disk and return the featureclass.
'******************************************************************************
Public Function OpenShapeFile(dir As String, Name As String) As IFeatureClass
    
  On Error GoTo ErrorHandler
  Dim pWSFact As IWorkspaceFactory
  Dim ConnectionProperties As IPropertySet
  Dim pShapeWS As IFeatureWorkspace
  Dim isShapeWS As Boolean
  Set OpenShapeFile = Nothing
  Set pWSFact = New ShapefileWorkspaceFactory
  isShapeWS = pWSFact.IsWorkspace(dir)
  If (isShapeWS) Then
    Set ConnectionProperties = New PropertySet
    ConnectionProperties.SetProperty "DATABASE", dir
    Set pShapeWS = pWSFact.Open(ConnectionProperties, 0)
    Dim pFClass As IFeatureClass
    Set pFClass = pShapeWS.OpenFeatureClass(Name)
    Set OpenShapeFile = pFClass
    Set pFClass = Nothing
  End If
  
GoTo Cleanup



Cleanup:
    Set pWSFact = Nothing
    Set ConnectionProperties = Nothing
    Set pShapeWS = Nothing
    Set pFClass = Nothing
    
    Exit Function
ErrorHandler:
  GoTo Cleanup
    
End Function

'******************************************************************************
'Subroutine: CheckInputDataProjection_ST
'Purpose:    Checks the input projection of input layers. Returns a FALSE if
'            all input layers are not in same projection
'******************************************************************************
Public Function CheckInputDataProjection_ST() As Boolean

    On Error GoTo ShowError
    CheckInputDataProjection_ST = False
    
    Dim pLayer As ILayer
    Dim pSpatialReference As ISpatialReference
    'check the projection of only dem, STREAM & landuse
    Dim i As Integer
    Dim pSpatialRefDict As Scripting.Dictionary
    Dim pSpatialReferenceName As String
    
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        Set pSpatialReference = GetSpatialReferenceForLayer(pLayer)
        If (pSpatialReference Is Nothing) Then
            pSpatialReferenceName = "Undefined"
        Else
            pSpatialReferenceName = pSpatialReference.Name
        End If
        
        ' now Check the Projection...........
        If pSpatialRefDict Is Nothing Then
            Set pSpatialRefDict = New Scripting.Dictionary
            pSpatialRefDict.Add pSpatialReferenceName, pSpatialReferenceName
        End If
        If Not pSpatialRefDict.Exists(pSpatialReferenceName) Then
            MsgBox "Spatial Refenece mismatch between input data. Please correct.", vbCritical, "BMP Siting Tool"
            Exit Function
        End If
     Next
     
     CheckInputDataProjection_ST = True
    
    GoTo Cleanup
ShowError:
    MsgBox "CheckInputDataProjection_ST: " & Err.Description
Cleanup:
    Set pLayer = Nothing
    Set pSpatialReference = Nothing
    
End Function

'******************************************************************************
'Subroutine: ValidateDatasets_ST
'Purpose:    Checks the input Datasets for specific fields availablility
'                    that are used for Analysis.......
'******************************************************************************

Public Function ValidateDatasets_ST(ByVal CreateJoin As Boolean) As Boolean
        
    On Error GoTo ShowError:
    ValidateDatasets_ST = False
    Dim pFLayer As IFeatureLayer
    Dim pFTable As ITable
    Dim pStandCol As IStandaloneTableCollection
    Dim pStTab As IStandaloneTable
    
    ' ******************************
    ' Nothing to Validate on Road Data.......
    
    ' ******************************
    ' Nothing to Validate on Stream Data....
    
    ' ******************************
    ' Validate Water Table Data..................
    Set pFTable = GetInputFeatureLayer(gWTdata)
    If Not pFTable Is Nothing Then
        If pFTable.Fields.FindField("GWdep_ft") = -1 Then
            MsgBox "Water table depth field is missing in Water table data." & vbCrLf & "'GWdep_ft' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
        End If
     End If
    
    ' ******************************
    ' Validate Soil Data................................
    
    ' First delete any Joins the Layer has..........
    ' Clear all joins on the layer.
    Set pFTable = GetFeatureLayer(gWorkingfolder, gSoildata) ' Get the Layer from Working folder....
    If Not pFTable Is Nothing Then
        If pFTable.Fields.FindField("MUKEY") = -1 Then
            MsgBox "Relational field is missing in Soil data." & vbCrLf & "'MUKEY' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
        End If
    End If
    Set pStandCol = gMap
    Dim i As Integer
    For i = 0 To pStandCol.StandaloneTableCount - 1
      If UCase(pStandCol.StandaloneTable(i).Name) = UCase(gSoilTable) Then
          Set pStTab = pStandCol.StandaloneTable(i)
          If pStTab.Table.Fields.FindField("MUKEY") = -1 Then
            MsgBox "Relational field is missing in Soil table." & vbCrLf & "'MUKEY' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
          End If
          If pStTab.Table.Fields.FindField("HYDGRP") = -1 Then
            MsgBox "Hydrogroup field is missing in Soil table." & vbCrLf & "'HYDGRP' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
          End If
          Exit For
      End If
    Next
    ' ******************************
    ' Now Create Soil relation......................
    ' ******************************
    If CreateJoin Then
        Set pFLayer = GetFeatureLayer(gWorkingfolder, gSoildata) ' Get the Layer from Working folder....
        If Not pFLayer Is Nothing Then
            Dim pTable As ITable
            Dim pFCLayer As IFeatureClass
            Dim pDispTable As IDisplayTable
            Set pDispTable = pFLayer
            Set pTable = pDispTable.DisplayTable
            
            Dim pDispTable2 As IDisplayTable
            Dim pReltable As ITable
            Set pDispTable2 = pStTab
            Set pReltable = pDispTable2.DisplayTable
            gSoildata = Create_Join(pFLayer, pTable, "MUKEY", pReltable, "MUKEY")
        End If
    End If
        
    
    ' ******************************
    ' Validate Landuse Data........................
    Set pFTable = GetInputFeatureLayer(gLandusedata)
    If Not pFTable Is Nothing Then
        If pFTable.Fields.FindField("LU_DESC") = -1 Then
            MsgBox "Landuse field is missing in Landuse data." & vbCrLf & "'LU_DESC' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
        End If
    End If

    ' ******************************
    ' Validate Mrlc Table Data......................
    For i = 0 To pStandCol.StandaloneTableCount - 1
      If UCase(pStandCol.StandaloneTable(i).Name) = UCase(gMrlcTable) Then
          Set pStTab = pStandCol.StandaloneTable(i)
          If pStTab.Table.Fields.FindField("LUCODE") = -1 Then
            MsgBox "Relational field is missing in MRLC table." & vbCrLf & "'LUCODE' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
          End If
          If pStTab.Table.Fields.FindField("LUNAME") = -1 Then
            MsgBox "Hydrogroup field is missing in MRLC table." & vbCrLf & "'LUNAME' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
          End If
          If pStTab.Table.Fields.FindField("SUITABLE") = -1 Then
            MsgBox "SUITABLE field is missing in MRLC table." & vbCrLf & "'SUITABLE' field should be available for analysis.", vbCritical, "BMP Siting Tool"
            Exit Function
          End If
          Exit For
      End If
    Next
    ' ******************************
    ' Now Create Mrlc relation.....................
    ' ******************************
    If CreateJoin Then
        Dim pLayer As ILayer
        Set pLayer = GetInputFeatureLayer(gMRLCdata)  ' Get the Layer from Working folder....
        If Not pLayer Is Nothing Then
            Dim pFClass As IFeatureClass
            Set pFClass = ConvertRastertoFeature(gRasterfolder, gMRLCdata, False, False)
            Dim pFeatlayer As IFeatureLayer
            Set pFeatlayer = New FeatureLayer
            Set pFeatlayer.FeatureClass = pFClass
        
            Set pDispTable = pFeatlayer
            Set pTable = pDispTable.DisplayTable
            Set pDispTable2 = pStTab
            Set pReltable = pDispTable2.DisplayTable
            gMRLCdata = Create_Join(pFeatlayer, pTable, "GRIDCODE", pReltable, "LUCODE")
        End If
    End If
    
    ValidateDatasets_ST = True
    GoTo Cleanup
    
ShowError:
    MsgBox "ValidateDatasets_ST: " & Err.Description
Cleanup:
    Set pFTable = Nothing
    Set pFLayer = Nothing
    Set pStandCol = Nothing
    Set pStTab = Nothing

End Function

Private Function Create_Join(ByVal pLayer As IFeatureLayer, pTable As ITable, pFld1 As String, pStTab As ITable, pFld2 As String) As String

    ' Create virtual relate
    Dim pMemRelFact As IMemoryRelationshipClassFactory
    Dim pRelClass As IRelationshipClass
    Set pMemRelFact = New MemoryRelationshipClassFactory
    Set pRelClass = pMemRelFact.Open(pLayer.Name, pStTab, pFld2, pTable, _
                            pFld1, "forward", "backward", esriRelCardinalityOneToOne)
    
    ' use Relate to perform a join
    Dim pDispRC As IDisplayRelationshipClass
    Set pDispRC = pLayer
    pDispRC.DisplayRelationshipClass pRelClass, esriLeftInnerJoin
    
    'QI

      Dim pGLayer As IGeoFeatureLayer
      Set pGLayer = pLayer
      
      Dim pFeatureLayerDef As IFeatureLayerDefinition
      Set pFeatureLayerDef = pGLayer
    
      Dim pFeatureClass As IFeatureClass
      Set pFeatureClass = pGLayer.DisplayFeatureClass
    
      Dim pQueryFilter As IQueryFilter
      Set pQueryFilter = New QueryFilter
      pQueryFilter.WhereClause = pFeatureLayerDef.DefinitionExpression
    
      Dim pSelFeatLayer As IFeatureLayer
      Set pSelFeatLayer = pFeatureLayerDef.CreateSelectionLayer(pLayer.Name, True, "", "")
    
      Dim pSelFeatureClass As IFeatureClass
      Set pSelFeatureClass = pSelFeatLayer.FeatureClass
    
      Dim pDataset As IDataset
      Set pDataset = pFeatureClass
    
      Dim pInDSName As IDatasetName
      Set pInDSName = pDataset.FullName
      
      Dim pFeatureClassName As IFeatureClassName
      Set pFeatureClassName = New FeatureClassName
      
      Dim pOutDatasetName As IDatasetName
      Set pOutDatasetName = pFeatureClassName
      pOutDatasetName.Name = pSelFeatureClass.AliasName & "_Join"
      
      Dim pWorkspaceName As IWorkspaceName
      Set pWorkspaceName = New WorkspaceName
      pWorkspaceName.PathName = gWorkingfolder
      pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"
      Set pOutDatasetName.WorkspaceName = pWorkspaceName
      'Give the output shapefile the same props as the input dataset
      pFeatureClassName.FeatureType = pFeatureClass.FeatureType
      pFeatureClassName.ShapeType = pFeatureClass.ShapeType
      pFeatureClassName.ShapeFieldName = pFeatureClass.ShapeFieldName
    
      'Export selected features
      Dim pExportOp As IExportOperation
      Set pExportOp = New ExportOperation
      pExportOp.ExportFeatureClass pInDSName, Nothing, Nothing, Nothing, pOutDatasetName, 0
      
'      ' Now Add the Layer to the Map....
'      Dim pOutputFeatLayer As IFeatureLayer
'      Set pOutputFeatLayer = New FeatureLayer
'      Set pOutputFeatLayer.FeatureClass = OpenShapeFile(gWorkingfolder, pSelFeatureClass.AliasName & "_Join")
'      pOutputFeatLayer.Name = pSelFeatureClass.AliasName & "_Join"
'      pOutputFeatLayer.Visible = True
'      gMap.AddLayer pOutputFeatLayer ' Add the Layer to the Map.........
      
      Create_Join = pSelFeatureClass.AliasName & "_Join"
      

End Function

'******************************************************************************
'Subroutine: GetSpatialReferenceForLayer
'Purpose:    Checks the input projection of an input layer. Checks the type of
'            input layer (feature/raster) and gets its spatial reference
'******************************************************************************
Public Function GetSpatialReferenceForLayer(pLayer As ILayer) As ISpatialReference
On Error GoTo ShowError:

    Dim pFeatureLayer As IFeatureLayer
    Dim pRasterLayer As IRasterLayer
    If (TypeOf pLayer Is IFeatureLayer) Then
        Set pFeatureLayer = pLayer
        Dim pGeoDataSet As IGeoDataset
        Set pGeoDataSet = pFeatureLayer.FeatureClass
        Set GetSpatialReferenceForLayer = pGeoDataSet.SpatialReference
    ElseIf (TypeOf pLayer Is IRasterLayer) Then
        Set pRasterLayer = pLayer
        Dim pRasterProps As IRasterProps
        Set pRasterProps = pRasterLayer.Raster
        Set GetSpatialReferenceForLayer = pRasterProps.SpatialReference
    End If
    GoTo Cleanup
ShowError:
    MsgBox "GetSpatialReferenceForLayer: " & Err.Description
Cleanup:
    Set pFeatureLayer = Nothing
    Set pRasterLayer = Nothing
    Set pGeoDataSet = Nothing
    Set pRasterProps = Nothing
End Function


Public Sub RenderUniqueValueFillSymbol_ST(pFeatureLayer As IFeatureLayer, pFieldNames As String, pHeading As String)

    On Error GoTo ErrorHandler
    
    Dim pLyr As IGeoFeatureLayer
    Set pLyr = pFeatureLayer
    
    ' **************************************************************
    ' Seems to be problem deletin/adding fields with result.........
    ' Copy the result Featureclass and Clean the featureClass.......
    ' **************************************************************
    Dim pFeatClass As IFeatureClass
    Set pFeatClass = pFeatureLayer.FeatureClass
    
    ' Now add the Field.....
    Dim pField As iField
    Dim pFieldEdit As IFieldEdit
    Set pField = New Field
    Set pFieldEdit = pField
    pFieldEdit.Name = "BMP_Combin"
    pFieldEdit.AliasName = "BMP_Combin"
    pFieldEdit.Type = esriFieldType.esriFieldTypeString
    pFieldEdit.Length = 255
    pFeatClass.AddField pField
    
    Dim strFields
    Dim strExp As String
    strFields = Split(pFieldNames, ";")
    Dim iCnt As Integer
    strExp = "Dim Output as String" & vbNewLine & "Output = "
    For iCnt = 0 To UBound(strFields)
        strExp = strExp & "[" & strFields(iCnt) & "]" & " & "","" & "
    Next
    strExp = Mid(strExp, 1, Len(strExp) - 9) & vbNewLine & "Output = Replace(Output, "" "", """")"
    strExp = strExp & vbNewLine & "Do While InStr(1, Output, "",,"", vbBinaryCompare) > 0 " & vbNewLine & "Output = Replace(Trim(Output), "",,"", "","")" & vbNewLine & "Loop"
    strExp = strExp & vbNewLine & "Output = Trim(Output)"
    strExp = strExp & vbNewLine & "if Mid(Output, 1, 1) = "","" Then Output = Mid(Output, 2)"
    strExp = strExp & vbNewLine & "if Mid(Output, Len(Output), 1) = "","" Then Output = Mid(Output, 1, Len(Output)-1)"

    
    ' Now Calulate the new field value.............
    Dim pCalc As ICalculator
    Set pCalc = New Calculator
    Dim pCursor As ICursor
    Set pCursor = pFeatClass.Update(Nothing, True)
    With pCalc
      Set .Cursor = pCursor
        .PreExpression = strExp
        .Expression = "Output"
        .Field = "BMP_Combin"
    End With
    pCalc.Calculate ' Execute the Query.......
        
    '** Creates a UniqueValuesRenderer and applies it to first layer in the map.
     Dim pFeatCls As IFeatureClass
     Set pFeatCls = pFeatureLayer.FeatureClass
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter 'empty supports: SELECT *
     Dim pFeatCursor As IFeatureCursor
     Set pFeatCursor = pFeatCls.Search(pQueryFilter, False)
 
     '** Make the color ramp we will use for the symbols in the renderer
     Dim pFromColor As IHsvColor
     Set pFromColor = New HsvColor
     pFromColor.Hue = 120  ' green
     pFromColor.Saturation = 100
     pFromColor.Value = 70

     Dim pToColor As IHsvColor
     Set pToColor = New HsvColor
     pToColor.Hue = 65        ' light green
     pToColor.Saturation = 100
     pToColor.Value = 100
    
     Dim rx As IAlgorithmicColorRamp
     Set rx = New AlgorithmicColorRamp
     rx.Algorithm = esriHSVAlgorithm
     rx.FromColor = pFromColor
     rx.ToColor = pToColor
     
     '** Make the renderer
     Dim pRender As IUniqueValueRenderer, n As Long
     Set pRender = New UniqueValueRenderer
     
     Dim symd As ISimpleFillSymbol
     Set symd = New SimpleFillSymbol
     symd.Style = esriSFSSolid
     symd.Outline.Width = 0.4
     
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "BMP_Combin"
     pRender.FieldDelimiter = ","
     pRender.DefaultSymbol = symd
     pRender.UseDefaultSymbol = True
          
     Dim pFeat As IFeature
     n = pFeatCls.FeatureCount(pQueryFilter)
    
    Dim i As Integer
     i = 0
     Dim ValFound As Boolean
     Dim NoValFound As Boolean
     Dim uh As Integer
     Dim pFields As IFields
     Dim iField As Integer

     Set pFields = pFeatCursor.Fields
     Do Until i = n
         Dim symx As ISimpleFillSymbol
         Set symx = New SimpleFillSymbol
         symx.Style = esriSFSSolid
         symx.Outline.Width = 0.4
         Set pFeat = pFeatCursor.NextFeature
         Dim x As String
         x = ""
         iField = pFields.FindField("BMP_Combin")
         x = x & "," & pFeat.Value(iField)  '*new Cory*
         x = Mid(x, 2)
        '** Test to see if we've already added this value
        '** to the renderer, if not, then add it.
        ValFound = False
        For uh = 0 To (pRender.ValueCount - 1)
          If pRender.Value(uh) = x Then
            NoValFound = True
            Exit For
          End If
        Next uh
        If Not ValFound Then
            pRender.AddValue x, pHeading, symx
            pRender.Label(x) = x
            pRender.Symbol(x) = symx
        End If
         
         i = i + 1
     Loop
     
     '** now that we know how many unique values there are
     '** we can size the color ramp and assign the colors.
     rx.Size = pRender.ValueCount
     rx.CreateRamp (True)
     Dim RColors As IEnumColors, ny As Long
     Set RColors = rx.Colors
     RColors.Reset
     For ny = 0 To (pRender.ValueCount - 1)
         Dim xv As String
         xv = pRender.Value(ny)
         If xv <> "" Then
             Dim jsy As ISimpleFillSymbol
             Set jsy = pRender.Symbol(xv)
             jsy.Color = RColors.Next
             pRender.Symbol(xv) = jsy
         End If
     Next ny
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     pRender.ColorScheme = "Custom"
     pRender.FieldType(0) = True
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = pHeading
 
     '** This makes the layer properties symbology tab show
     '** show the correct interface.
     Dim hx As IRendererPropertyPage
     Set hx = New UniqueValuePropertyPage
     pLyr.RendererPropertyPageClassID = hx.ClassID
 
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     
     '** Draw the map
     gMxDoc.ActiveView.Refresh
    
  
  GoTo Cleanup
ErrorHandler:
    MsgBox "RenderUniqueValueFillSymbol_ST: " & Err.Description
Cleanup:
End Sub

Public Sub RenderUniqueValueonFields_ST(pFeatureLayer As IFeatureLayer, pFieldNames As String, pHeading As String)

    On Error GoTo ErrorHandler
    
    Dim pLyr As IGeoFeatureLayer
    Set pLyr = pFeatureLayer
    
    ' First Count the Fields.....
    Dim strFields
    strFields = Split(pFieldNames, ";")
        
    '** Creates a UniqueValuesRenderer and applies it to first layer in the map.
     Dim pFeatCls As IFeatureClass
     Set pFeatCls = pFeatureLayer.FeatureClass
     Dim pQueryFilter As IQueryFilter
     Set pQueryFilter = New QueryFilter 'empty supports: SELECT *
     Dim pFeatCursor As IFeatureCursor
     Set pFeatCursor = pFeatCls.Search(pQueryFilter, False)
 
     '** Make the color ramp we will use for the symbols in the renderer
     Dim pFromColor As IHsvColor
     Set pFromColor = New HsvColor
     pFromColor.Hue = 120  ' green
     pFromColor.Saturation = 100
     pFromColor.Value = 70

     Dim pToColor As IHsvColor
     Set pToColor = New HsvColor
     pToColor.Hue = 65        ' light green
     pToColor.Saturation = 100
     pToColor.Value = 100
    
     Dim rx As IAlgorithmicColorRamp
     Set rx = New AlgorithmicColorRamp
     rx.Algorithm = esriHSVAlgorithm
     rx.FromColor = pFromColor
     rx.ToColor = pToColor
     
     '** Make the renderer
     Dim pRender As IUniqueValueRenderer, n As Long
     Set pRender = New UniqueValueRenderer
     
     Dim symd As ISimpleFillSymbol
     Set symd = New SimpleFillSymbol
     symd.Style = esriSFSSolid
     symd.Outline.Width = 0.4
     
     '** These properties should be set prior to adding values
     pRender.FieldCount = 1
     pRender.Field(0) = "BMP_Combinations"
     pRender.FieldDelimiter = ","
     pRender.DefaultSymbol = symd
     pRender.UseDefaultSymbol = True
          
     Dim pFeat As IFeature
     n = pFeatCls.FeatureCount(pQueryFilter)
    
    Dim i As Integer
     i = 0
     Dim ValFound As Boolean
     Dim NoValFound As Boolean
     Dim uh As Integer
     Dim pFields As IFields
     Dim iField As Integer

     Set pFields = pFeatCursor.Fields
     Do Until i = n
         Dim symx As ISimpleFillSymbol
         Set symx = New SimpleFillSymbol
         symx.Style = esriSFSSolid
         symx.Outline.Width = 0.4
         Set pFeat = pFeatCursor.NextFeature
         Dim x As String
         x = ""
         iField = pFields.FindField(pRender.Field(0))
        x = x & "," & pFeat.Value(iField)  '*new Cory*
        x = Mid(x, 2)
         '** Test to see if we've already added this value
         '** to the renderer, if not, then add it.
         ValFound = False
         For uh = 0 To (pRender.ValueCount - 1)
           If pRender.Value(uh) = x Then
             NoValFound = True
             Exit For
           End If
         Next uh
         If Not ValFound Then
             pRender.AddValue x, pHeading, symx
             pRender.Label(x) = x
             pRender.Symbol(x) = symx
         End If
         i = i + 1
     Loop
     
     '** now that we know how many unique values there are
     '** we can size the color ramp and assign the colors.
     rx.Size = pRender.ValueCount
     rx.CreateRamp (True)
     Dim RColors As IEnumColors, ny As Long
     Set RColors = rx.Colors
     RColors.Reset
     For ny = 0 To (pRender.ValueCount - 1)
         Dim xv As String
         xv = pRender.Value(ny)
         If xv <> "" Then
             Dim jsy As ISimpleFillSymbol
             Set jsy = pRender.Symbol(xv)
             jsy.Color = RColors.Next
             pRender.Symbol(xv) = jsy
         End If
     Next ny
 
     '** If you didn't use a color ramp that was predefined
     '** in a style, you need to use "Custom" here, otherwise
     '** use the name of the color ramp you chose.
     pRender.ColorScheme = "Custom"
     pRender.FieldType(0) = True
     Set pLyr.Renderer = pRender
     pLyr.DisplayField = pHeading
 
     '** This makes the layer properties symbology tab show
     '** show the correct interface.
     Dim hx As IRendererPropertyPage
     Set hx = New UniqueValuePropertyPage
     pLyr.RendererPropertyPageClassID = hx.ClassID
 
     '** Refresh the TOC
     gMxDoc.ActiveView.ContentsChanged
     gMxDoc.UpdateContents
     
     '** Draw the map
     gMxDoc.ActiveView.Refresh
    
  
  GoTo Cleanup
ErrorHandler:
    MsgBox "RenderUniqueValueonFields_ST: " & Err.Description
Cleanup:
End Sub

Public Function GetWorkspace(sPath As String) As IWorkspace

    On Error GoTo ErrorHandler
    ' This function returns a shapefile workspace object  given the path.
    ' The path needs to contain shape files
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New ShapefileWorkspaceFactory
    Set GetWorkspace = pWSF.OpenFromFile(sPath, 0)
    
    Exit Function
ErrorHandler:
  HandleError True, "GetWorkspace " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function


Public Function Parse_Expression(ByVal strExp As String, ByVal pParse As Boolean) As String
    
    On Error GoTo ErrorHandler
    Dim strTmp As String
    Dim iCnt As Integer
    iCnt = 1
    
    strTmp = Mid(strExp, iCnt, 1)
    Do While Not IsNumeric(strTmp)
        Parse_Expression = Parse_Expression & strTmp
        iCnt = iCnt + 1
        strTmp = Mid(strExp, iCnt, 1)
    Loop
    
    If pParse Then
        Parse_Expression = Parse_Expression & " " & Trim(Mid(strExp, iCnt)) / (0.00002295675 * gCellSize * gCellSize)
    Else
        Parse_Expression = Parse_Expression & " " & Trim(Mid(strExp, iCnt))
    End If

Exit Function
ErrorHandler:

  HandleError True, "Parse_Expression " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND

End Function

Public Function Generic_Trim(ByVal strSearch As String, strFind As String, strReplace As String) As String

    On Error GoTo ErrorHandler
    
    'Trim at the Start....
    Do While Mid(strSearch, 1, 1) = strFind
        strSearch = strReplace & Mid(strSearch, 2)
    Loop
    'Trim at the End....
    Do While Mid(strSearch, Len(strSearch), 1) = strFind
        strSearch = Mid(strSearch, 1, Len(strSearch) - 1) & strReplace
    Loop
    
    Generic_Trim = strSearch

  Exit Function
ErrorHandler:
  HandleError True, "Generic_Trim" & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function

Public Sub AlwaysOnTop_ST(FrmID As Form, OnTop As Integer)
    ' ===========================================
    ' Requires the following declaration
    ' For VB4:
    ' Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    ' ===========================================
    ' Usage:
    ' AlwaysOnTop Me, -1  ' To make always on top
    ' AlwaysOnTop Me, -2  ' To make NOT always on top
    ' ===========================================
    If OnTop = -1 Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOOWNERZORDER)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOOWNERZORDER)
    End If
End Sub

