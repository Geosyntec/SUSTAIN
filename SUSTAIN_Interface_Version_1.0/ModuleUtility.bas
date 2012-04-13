Attribute VB_Name = "ModuleUtility"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleUtility
'   Purpose:     Contains All helper functions to open feature, raster layer from map,
'                disk, convert between feature and raster functions, data management
'                helper functions, etc
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:
'                Modified: 08/19/2004 - mira chokshi added comments to project
'
'******************************************************************************
Option Explicit
Option Base 0

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*** Global Variables
Public gCommandExtension As IExtensionConfig
Public gMxDoc As IMxDocument
Public gMap As IMap
Public gApplicationPath As String
Public gMapTempFolder As String
Public gMetersPerUnit As Double
Public gLinearUnitName As String
Public bType As String
Public bSplitter As Boolean
Public bRegulator As Boolean
Public gForestTimeSeries As String
Public gAgriTimeSeries As String
Public gUrbanTimeSeries As String
Public gApplication As IApplication
Public gManualDelineationFlag As Boolean    'Flag to set manual delineation tools active

Public gDEMRaster As IRaster
Public gCellSize As Double
Public gReclassOp As IReclassOp
Public gNeighborhoodOp As INeighborhoodOp
Public gAlgebraOp As IMapAlgebraOp
Public gHydrologyOp As IHydrologyOp
Public gRasterDistanceOp As IDistanceOp

'Define variables to store different layers, tables and database, field names
Public gLayerNameDictionary As Scripting.Dictionary
Public gSUBBASINFieldName As String
Public gSUBBASINRFieldName As String
Public gIsStreamAlongFlowDir As Boolean 'Set true if the streams are along flow direction

'Public variables for ADO connection
Public gAdoConn As ADODB.Connection

'Define variables to save bmp details temporary
Public gBMPDetailDict As Scripting.Dictionary
Public gBMPDictionary As Scripting.Dictionary
Public gSubWaterLandUseDict As Dictionary
Public gBufferStripDetailDict As Scripting.Dictionary

'Public BMPCalculator directory names
Public gBmpInputDir As String
Public gBmpOutputDir As String

'Variables to identify definition of BMPTypes or creation of BMP instance
Public gNewBMPType As String
Public gNewBMPId As Integer
Public gNewBMPName As String

'Variables to identify definition of VFSTypes or creation of VFS instance
Public gNewVFSType As String
Public gNewVFSId As Integer
Public gNewVFSName As String

'Variable to identify which parameter is selected to be optimized
Public gCurOptParam As String

' Define Global variables for Geodatabase & Images....
Public gGDBpath As String
Public gCostDBpath As String
Public gDEMLayer As String
Public gStreamLayer As String

' Global for Dataloaded flag.....
Public gDataLoad As Boolean
Public gGDBFlag As Boolean
Public gDefLayers As Boolean
Public gDefPollutants As Boolean
Public gInternalSimulation As Boolean
Public gExternalSimulation As Boolean
Public gFeatClassDictionary As Scripting.Dictionary
Public gColorDict As Scripting.Dictionary
Public gPostDevfile As String
Public gPreDevfile As String
Public gBMPTypeDict As Scripting.Dictionary
Public gBMPCatDict As Scripting.Dictionary
Public gBMPOptionsDict As Scripting.Dictionary
Public gBMPDefTab As Integer
Public gBMPEditMode As Boolean
Public gBMPPlacedDict As Scripting.Dictionary

Public Type gParamInfo
    name As String
    Decay As Double
    PctRem As Double
    K As Double
    C As Double
End Type

Public gMaxPollutants As Integer
Public gParamInfos() As gParamInfo
Public gPollutants() As String

'Global variable for the FrmOutlet type
Public gBMPOutletType As Integer

'Data Struct to hold assessment info
Public Type gAssessInfo
    Factor As String
    Unit As String
    isTargetEval As Boolean
    Target As Double
    isRedEval As Boolean
    Reduction As Double
End Type
Public gAssessInfos() As gAssessInfo

'Variables for the browse for folder function
Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Const dhcErrorExtendedError = 1208&
Const dhcNoError = 0&

'specify root dir for browse for folder by constants
'you can also specify values by constants for searhcable folders and options.
Const dhcCSIdlDesktop = &H0
Const dhcCSIdlPrograms = &H2
Const dhcCSIdlControlPanel = &H3
Const dhcCSIdlInstalledPrinters = &H4
Const dhcCSIdlPersonal = &H5
Const dhcCSIdlFavorites = &H6
Const dhcCSIdlStartupPmGroup = &H7
Const dhcCSIdlRecentDocDir = &H8
Const dhcCSIdlSendToItemsDir = &H9
Const dhcCSIdlRecycleBin = &HA
Const dhcCSIdlStartMenu = &HB
Const dhcCSIdlDesktopDirectory = &H10
Const dhcCSIdlMyComputer = &H11
Const dhcCSIdlNetworkNeighborhood = &H12
Const dhcCSIdlNetHoodFileSystemDir = &H13
Const dhcCSIdlFonts = &H14
Const dhcCSIdlTemplates = &H15
'constants for limiting choices for BrowseForFolder Dialog
Const dhcBifReturnAll = &H0
Const dhcBifReturnOnlyFileSystemDirs = &H1
Const dhcBifDontGoBelowDomain = &H2
Const dhcBifIncludeStatusText = &H4
Const dhcBifSystemAncestors = &H8
Const dhcBifBrowseForComputer = &H1000
Const dhcBifBrowseForPrinter = &H2000

' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Const REG_SZ = 1
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE _
    Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))


Public Declare Function GetWindowsDirectory Lib "kernel32" _
   Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal _
   lpString1 As String, ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As _
   BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal _
   pidList As Long, ByVal lpBuffer As String) As Long
   
'corrected
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
(ByVal hWndOwner As Long, ByVal nFolder As Long, ByRef pidl As Long) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long

Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

    
Public Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
        
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
         (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As _
         Integer, ByVal lParam As Any) As Long

'constants for searching the ListBox
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2


' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "..\ModuleUtility.bas"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

' Always on Top.............
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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

'******************************************************************************
'Subroutine: InitializeMapDocument
'Author:     Mira Chokshi
'Purpose:    Initializes Current Map, Application Path and Map Algebra, Hydrology,
'            Neighborhood and Reclassification Operators
'******************************************************************************
Public Sub InitializeMapDocument()
    Set gMxDoc = gApplication.Document
    Set gMap = gMxDoc.FocusMap
    
    'Check for extension
    Dim isExtensionON As Boolean
    isExtensionON = GetSUSTAINExtension(gApplication)
    
    'Define the application path
    Call DefineApplicationPath
        
    'Read data layer information from src file
    Call ReadLayerTagDictionaryToSRCFile

End Sub

 

'******************************************************************************
'Subroutine: GetSUSTAINExtension
'Author:     Mira Chokshi
'Purpose:    Initializes Current Map, Application Path and Map Algebra, Hydrology,
'            Neighborhood and Reclassification Operators
'******************************************************************************
 Public Function GetSUSTAINExtension(pApplication As IApplication) As Boolean
 On Error GoTo ErrorHandler
   
  Set gCommandExtension = pApplication.FindExtensionByName("SUSTAIN Extension")
 
  If (Not gCommandExtension Is Nothing) Then
    If (gCommandExtension.State = esriESEnabled) Then
     GetSUSTAINExtension = True
     Else
      GetSUSTAINExtension = False
    End If
  Else
     GetSUSTAINExtension = False
  End If
 
  Exit Function
ErrorHandler:
  MsgBox "GetSUSTAINExtension: " & Err.description

 End Function
     

'******************************************************************************
'Subroutine: InitializeOperators
'Author:     Mira Chokshi
'Purpose:    Initializes Global Map Algebra, Hydrology, Neighborhood, Reclass
'            Operators.
'            Set extent, cell size to dem raster's extent and cell size
'******************************************************************************
Public Function InitializeOperators(Optional isLanduse As Boolean) As Boolean
    On Error GoTo EH
    
    If Not CheckSpatialAnalystLicense() Then
        InitializeOperators = False
        Exit Function
    End If
    
    Dim pLayerName As String
    
    If isLanduse Then
        pLayerName = "Landuse"
    Else
        pLayerName = "DEM"
    End If
    
    Set gMxDoc = gApplication.Document
    Set gMap = gMxDoc.FocusMap
  
    Dim pDEMRasterLayer As IRasterLayer
    Set pDEMRasterLayer = GetInputRasterLayer(pLayerName) ' "Landuse
    If pDEMRasterLayer Is Nothing Then Exit Function
    Set gDEMRaster = pDEMRasterLayer.Raster
        
    'Get the raster props
    Dim pDEMRasterProps As IRasterProps
    Set pDEMRasterProps = gDEMRaster
    
    'Get the raster cell size
    gCellSize = (pDEMRasterProps.MeanCellSize.X + pDEMRasterProps.MeanCellSize.Y) / 2
        
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Dim pRWS As IRasterWorkspace2
    Set pRWS = pWSF.OpenFromFile(gMapTempFolder, 0)
    Dim pRAEnv As IRasterAnalysisEnvironment
    
    ' Create the global gReclassOp object
    Set gReclassOp = New RasterReclassOp
    Set pRAEnv = gReclassOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = gDEMRaster
    
    ' Create the global gNeighborhoodOp object
    Set gNeighborhoodOp = New RasterNeighborhoodOp
    Set pRAEnv = gNeighborhoodOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = gDEMRaster
   
    ' Create the global gAlgebraOp object
    Set gAlgebraOp = New RasterMapAlgebraOp
    Set pRAEnv = gAlgebraOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = gDEMRaster

    ' Create the global gHydrologyOp object
    Set gHydrologyOp = New RasterHydrologyOp
    Set pRAEnv = gHydrologyOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = gDEMRaster
    
    ' Create the global raster distance object
    Set gRasterDistanceOp = New RasterDistanceOp
    Set pRAEnv = gRasterDistanceOp
    pRAEnv.SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
    pRAEnv.SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
    Set pRAEnv.OutSpatialReference = pDEMRasterProps.SpatialReference
    Set pRAEnv.OutWorkspace = pRWS
    Set pRAEnv.Mask = gDEMRaster
    
    ReDim Preserve gPollutants(0) As String
    InitializeOperators = True
    GoTo CleanUp
    
EH:
    MsgBox "Failed in Initializing the system - " & Err.description

CleanUp:
    Set pDEMRasterLayer = Nothing
    Set pDEMRasterProps = Nothing
    Set pRWS = Nothing
    Set pWSF = Nothing
    Set pRAEnv = Nothing
End Function


'******************************************************************************
'Subroutine: DefineApplicationPath
'Author:     Mira Chokshi
'Purpose:    Gets the path of the map and stores it in global variable
'******************************************************************************
Public Sub DefineApplicationPath()
  
  Dim pTemplates As ITemplates
  Dim lTempCount As Long
  Dim strDocPath As String
  Set pTemplates = gApplication.Templates
  
  Dim pDoc As IDocument
  Set pDoc = gApplication.Document

  lTempCount = pTemplates.Count
  ' The document is always the last item
  strDocPath = pTemplates.Item(lTempCount - 1)
  Dim strMatch As String
  strMatch = pDoc.Title
  If Replace(UCase(pDoc.Title), ".MXD", "") = UCase(pDoc.Title) Then strMatch = pDoc.Title & ".mxd"
  'strDocPath = Replace(strDocPath, pDoc.Title, "")
  strDocPath = Replace(strDocPath, strMatch, "")
 
  gApplicationPath = strDocPath
End Sub


'******************************************************************************
'Subroutine: CheckSpatialAnalystLicense
'Author:     Mira Chokshi
'Purpose:    Check the availability of Spatial Analyst license, returns
'            TRUE if SA license found, else returns FALSE. This subroutine
'            should be called at the beginning of any process using SA.
'******************************************************************************
Public Function CheckSpatialAnalystLicense() As Boolean
On Error GoTo ShowError
    
    CheckSpatialAnalystLicense = False

    Dim pLicManager As IExtensionManager
    Set pLicManager = New ExtensionManager
    
    Dim pLicAdmin As IExtensionManagerAdmin
    Set pLicAdmin = pLicManager
    
    Dim saUID As Variant
    saUID = "esriSpatialAnalystUI.SAExtension.1"
    
    Dim pUID As New UID
    pUID.value = saUID
    
    Dim v As Variant
    Call pLicAdmin.AddExtension(pUID, v)
    
    Dim pExtension As IExtension
    Set pExtension = pLicManager.FindExtension(pUID)
    
    Dim pExtensionConfig As IExtensionConfig
    Set pExtensionConfig = pExtension
    pExtensionConfig.State = esriESEnabled
    
    CheckSpatialAnalystLicense = True
    GoTo CleanUp

ShowError:
    MsgBox "Failed in checking Spatial Analyst License."
CleanUp:
    Set pLicManager = Nothing
    Set pLicAdmin = Nothing
    Set saUID = Nothing
    Set pUID = Nothing
    Set v = Nothing
    Set pExtension = Nothing
    Set pExtensionConfig = Nothing
End Function


'******************************************************************************
'Subroutine: CheckSpatialAnalystLicense
'Author:     Mira Chokshi
'Purpose:    Change pixeltype for those not supported by VB. This function
'            is called during cell based computation of rasters.
'******************************************************************************
Public Function GetVBSupportedPixelType(iPixeltype As Integer)
On Error GoTo ShowError

    If iPixeltype <= 4 Then
        GetVBSupportedPixelType = 3 ' PT_UCHAR
    ElseIf iPixeltype <= 6 Then
        GetVBSupportedPixelType = 6 ' PT_SHORT
    ElseIf iPixeltype <= 8 Then
        GetVBSupportedPixelType = 8 ' PT_LONG
    ElseIf iPixeltype >= 9 Then
        GetVBSupportedPixelType = 9 ' PT_FLOAT
    End If
    Exit Function
ShowError:
    MsgBox "GetVBSupportedPixelType: " & Err.description
End Function


'******************************************************************************
'Subroutine: AddRasterToMap
'Author:     Mira Chokshi
'Purpose:    Creates a raster layer from a raster and adds it to the map.
'******************************************************************************
Public Sub AddRasterToMap(ByRef pRaster As IRaster, pName As String, bVisible As Boolean)
On Error GoTo ShowError
    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = New RasterLayer
    pRasterLayer.CreateFromRaster pRaster
    pRasterLayer.Visible = bVisible
    AddLayerToMap pRasterLayer, pName
    GoTo CleanUp
ShowError:
    MsgBox "AddRasterToMap: " & Err.description
CleanUp:
    Set pRasterLayer = Nothing
End Sub


'******************************************************************************
'Subroutine: AddLayerToMap
'Author:     Mira Chokshi
'Purpose:    General Function to Add a layer(feature/raster) to map.
'******************************************************************************
Public Sub AddLayerToMap(ByRef pLayer As ILayer, pLayerName As String)
On Error GoTo ShowError
        
    Dim pLegendInfo As ILegendInfo
    Dim pLegendGroup As ILegendGroup
    ' Set the name, make it invisible
    pLayer.name = pLayerName
    pLayer.Visible = True
    ' Expand the legend group for this layer.
    Set pLegendInfo = pLayer
    Set pLegendGroup = pLegendInfo.LegendGroup(0)
    pLegendGroup.Visible = False ' set to False to hide.
    ' Add the legend to GIS group layer - mira chokshi added on 08/26/04
    gMap.AddLayer pLayer
    gMxDoc.ActiveView.Refresh
    gMxDoc.UpdateContents
    GoTo CleanUp
    
ShowError:
    MsgBox "AddLayerToMap: " & Err.description
CleanUp:
    Set pLegendGroup = Nothing
    Set pLegendInfo = Nothing
End Sub



'******************************************************************************
'Subroutine: AddTableToMap
'Author:     Mira Chokshi
'Purpose:    General Function to Add a table to map.
'******************************************************************************
Public Sub AddTableToMap(pTable As iTable)
On Error GoTo ShowError
    Dim pStandAloneTableColl As IStandaloneTableCollection
    Set pStandAloneTableColl = gMap
    Dim pStandAloneTable As IStandaloneTable
    Set pStandAloneTable = New StandaloneTable
    Set pStandAloneTable.Table = pTable
    pStandAloneTableColl.AddStandaloneTable pStandAloneTable
    gMxDoc.ActiveView.Refresh
    gMxDoc.UpdateContents
    GoTo CleanUp
ShowError:
    MsgBox "AddTableToMap: " & Err.description
CleanUp:
    Set pStandAloneTableColl = Nothing
    Set pStandAloneTable = Nothing
End Sub


'******************************************************************************
'Subroutine: DeleteLayerFromMap
'Author:     Mira Chokshi
'Purpose:    General Function to Delete a layer(feature/raster/etc) from map.
'******************************************************************************
Public Sub DeleteLayerFromMap(pLayerName As String)
On Error GoTo ShowError
    'If the map has subwatershed layer, remove it
    Dim i As Integer
    Dim pLayer As ILayer
    Dim pdatalayer2 As IDataLayer2

    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If (UCase(pLayer.name) = UCase(pLayerName)) Then
            Set pdatalayer2 = pLayer
            gMap.DeleteLayer pLayer
            pdatalayer2.Disconnect
            Exit For
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "DeleteLayerFromMap: " & Err.description
CleanUp:
    Set pLayer = Nothing
End Sub

'******************************************************************************
'Subroutine: ExpandPointEnvelope
'Author:     Mira Chokshi
'Purpose:    Function to get point on the map and expand it
'******************************************************************************
Function ExpandPointEnvelope(pPointEnvelope As IEnvelope) As IEnvelope
    'Get the active view and get 1/100 of its extent
    Dim pActiveView As IActiveView
    Set pActiveView = gMxDoc.ActiveView
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pActiveView.Extent
    pEnvelope.Expand 0.01, 0.01, True
    
    'Expand the point envelop to the size of 1/100 of the map extent
    pPointEnvelope.Expand pEnvelope.Width, pEnvelope.Height, False
    Set pEnvelope = Nothing
    Set pActiveView = Nothing
    
    'Return the expanded envelope
    Set ExpandPointEnvelope = pPointEnvelope
End Function

'******************************************************************************
'Subroutine: GetInputGroupLayer
'Author:     Mira Chokshi
'Purpose:    Function to get feature layer from map. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetInputGroupLayer(keylayername As String) As IGroupLayer
On Error GoTo ShowError
    Dim GLayerName As String
    GLayerName = keylayername
    Dim i As Integer
    Dim pGLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pGLayer = gMap.Layer(i)
        If ((pGLayer.name = GLayerName) And (TypeOf pGLayer Is IGroupLayer) And pGLayer.Valid) Then
            Set GetInputGroupLayer = pGLayer     'Get the plot layer
            GoTo CleanUp
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "GetInputGroupLayer: " & Err.description
CleanUp:
    Set pGLayer = Nothing
End Function


'******************************************************************************
'Subroutine: GetInputFeatureLayer
'Author:     Mira Chokshi
'Purpose:    Function to get feature layer from map. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetInputFeatureLayer(keylayername As String) As IFeatureLayer
On Error GoTo ShowError
    
    If (gLayerNameDictionary Is Nothing) Then
        Exit Function
    End If
    
    Dim FLayerName As String
    If (gLayerNameDictionary.Exists(keylayername)) Then
        FLayerName = gLayerNameDictionary.Item(keylayername)
    Else
        FLayerName = keylayername
    End If
    
    Dim i As Integer
    Dim pFlayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pFlayer = gMap.Layer(i)
        If ((UCase(pFlayer.name) = UCase(FLayerName)) And (TypeOf pFlayer Is IFeatureLayer) And pFlayer.Valid) Then
            Set GetInputFeatureLayer = pFlayer     'Get the plot layer
            GoTo CleanUp
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "GetInputFeatureLayer: " & Err.description
CleanUp:
    Set pFlayer = Nothing
End Function


'******************************************************************************
'Subroutine: GetInputLayerIndex
'Author:     Mira Chokshi
'Purpose:    Function to get feature layer from map. If featurelayer does not
'            contain feature class, no feature layer is returned.
'******************************************************************************
Public Function GetInputLayerIndex(keylayername As String) As Integer
On Error GoTo ShowError
    
    If (gLayerNameDictionary Is Nothing) Then
        Exit Function
    End If
    
    Dim LayerName As String
    If (gLayerNameDictionary.Exists(keylayername)) Then
        LayerName = gLayerNameDictionary.Item(keylayername)
    Else
        LayerName = keylayername
    End If
    
    Dim i As Integer
    Dim pLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If ((pLayer.name = LayerName) And pLayer.Valid) Then
            GetInputLayerIndex = i     'Get the layer index
            GoTo CleanUp
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "GetInputLayerIndex: " & Err.description
CleanUp:
    Set pLayer = Nothing
End Function


'******************************************************************************
'Subroutine: MoveLayerToIndex
'Author:     Mira Chokshi
'Purpose:    Function to move a particular layer to a given index
'******************************************************************************
Public Function MoveLayerToIndex(keylayername As String, MoveIndex As Integer)
On Error GoTo ShowError
    
    If (gLayerNameDictionary Is Nothing) Then
        Exit Function
    End If
    
    Dim LayerName As String
    If (gLayerNameDictionary.Exists(keylayername)) Then
        LayerName = gLayerNameDictionary.Item(keylayername)
    Else
        LayerName = keylayername
    End If
    
    Dim i As Integer
    Dim pLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If ((pLayer.name = LayerName) And pLayer.Valid) Then
            gMap.MoveLayer pLayer, MoveIndex    'move layer to a given index
        End If
    Next
    GoTo CleanUp

ShowError:
    MsgBox "GetInputLayerIndex: " & Err.description
CleanUp:
    Set pLayer = Nothing
End Function


'******************************************************************************
'Subroutine: GetInputRasterLayer
'Author:     Mira Chokshi
'Purpose:    Function to get raster layer from map. If rasterlayer does not
'            contain raster dataset, no raster layer is returned.
'******************************************************************************
Public Function GetInputRasterLayer(keylayername As String) As IRasterLayer
On Error GoTo ShowError
    
    If (gLayerNameDictionary Is Nothing) Then
        Exit Function
    End If
    
    Dim RLayerName As String
    If (gLayerNameDictionary.Exists(keylayername)) Then
        RLayerName = gLayerNameDictionary.Item(keylayername)
    Else
        RLayerName = keylayername
    End If
    Dim i As Integer
    Dim pRLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pRLayer = gMap.Layer(i)
        If ((UCase(pRLayer.name) = UCase(RLayerName)) And (TypeOf pRLayer Is IRasterLayer) And pRLayer.Valid) Then
            Set GetInputRasterLayer = pRLayer     'Get the plot layer
            GoTo CleanUp
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "GetInputRasterLayer: " & Err.description
CleanUp:
    Set pRLayer = Nothing
End Function


'******************************************************************************
'Subroutine: GetInputDataTable
'Author:     Mira Chokshi
'Purpose:    Function to get input data table. This function returns a table
'            added on the map interface.
'******************************************************************************
Public Function GetInputDataTable(tablename As String) As iTable
On Error GoTo ShowError
   
   If (gLayerNameDictionary Is Nothing) Then
        Exit Function
   End If
    
   Dim pTableName As String
   If (gLayerNameDictionary.Exists(tablename)) Then
    pTableName = gLayerNameDictionary.Item(tablename)
   Else
    pTableName = tablename
   End If
   
   Dim pTabCollection As IStandaloneTableCollection
   Dim pStTable As IStandaloneTable
   Dim pDispTable As IDisplayTable
   Set pTabCollection = gMap
   Dim i As Integer
   For i = 0 To (pTabCollection.StandaloneTableCount - 1)
      Set pStTable = pTabCollection.StandaloneTable(i)
      If (UCase(pStTable.name) = UCase(pTableName)) Then
            Set GetInputDataTable = pStTable.Table
      End If
   Next
   GoTo CleanUp
 
ShowError:
   MsgBox "GetInputDataTable: " & Err.description
CleanUp:
   Set pTabCollection = Nothing
   Set pStTable = Nothing
   Set pDispTable = Nothing
End Function


'******************************************************************************
'Subroutine: DeleteDataTable
'Author:     Mira Chokshi
'Purpose:    Remove the data table from the map, and delete from disk.
'******************************************************************************
Public Sub DeleteDataTable(tableDir As String, tablename As String)
On Error GoTo ShowError
 
  'Define temporary variables
  Dim pTabCollection As IStandaloneTableCollection
  Set pTabCollection = gMap
  
  Dim pStTable As IStandaloneTable
  Dim i As Integer
  
  For i = 0 To (pTabCollection.StandaloneTableCount - 1)
      Set pStTable = pTabCollection.StandaloneTable(i)
      If (pStTable.name = tablename) Then
            pTabCollection.RemoveStandaloneTable pStTable
            Exit For ' After deleting the table -- Sabu Paul, Aug 24, 2004
      End If
  Next
  
  Dim fso As Scripting.FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(tableDir & "\" & tablename & ".dbf")) Then
    fso.DeleteFile (tableDir & "\" & tablename & ".dbf")
  End If
  GoTo CleanUp
 
ShowError:
   MsgBox "DeleteDataTable: " & Err.description
CleanUp:
    Set pTabCollection = Nothing
    Set pStTable = Nothing
    Set fso = Nothing
End Sub

'******************************************************************************
'Subroutine: CreateUniqueTableName
'Author:     Mira Chokshi
'Purpose:    Creates a unique table name by adding a number at end.
'******************************************************************************
Public Function CreateUniqueTableName(iDir As String, iTable As String) As String
On Error GoTo ShowError
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Dim FileName As String
    Dim fcounter As Integer
    fcounter = 1
    FileName = iDir & "\" & iTable & CStr(fcounter) & ".dbf"
    Do While fsObj.FileExists(FileName)
        fcounter = fcounter + 1
        FileName = iDir & "\" & iTable & CStr(fcounter) & ".dbf"
    Loop
    
    'Return the unique table name
    CreateUniqueTableName = iTable & CStr(fcounter)
    GoTo CleanUp
 
ShowError:
   MsgBox "CreateUniqueTableName: " & Err.description
CleanUp:
    Set fsObj = Nothing
End Function


'******************************************************************************
'Subroutine: ConvertFeatureToRaster
'Author:     Mira Chokshi
'Purpose:    Converts feature class to raster dataset. Required parameters
'            include featureclass, name of the field used for conversion,
'            name of the raster file name. Returns a RasterDataset.
'******************************************************************************
Public Function ConvertFeatureToRaster(pFeatureclass As IFeatureClass, pFieldName As String, pFileName As String, pQueryFilter As IQueryFilter) As IRasterDataset
   
On Error GoTo ShowError
    
    Dim pDEMRLayer As IRasterLayer
    Set pDEMRLayer = GetInputRasterLayer("DEM")
    
    'Create a workspace
    Dim pWSF As IWorkspaceFactory
    Dim pWs As IWorkspace
    Set pWSF = New RasterWorkspaceFactory
    Set pWs = pWSF.OpenFromFile(gMapTempFolder, 0)
    
    'Select all features of the feature class
    Dim pSelectionSet As ISelectionSet
    ' Use the query filter to select features from STREAM feature layer
    Set pSelectionSet = pFeatureclass.Select(pQueryFilter, esriSelectionTypeIDSet, esriSelectionOptionNormal, Nothing)
    
    ' Define the featureclassdescriptor
    Dim pGeoDataDescriptor As IFeatureClassDescriptor
    Set pGeoDataDescriptor = New FeatureClassDescriptor
    ' Get the selection set
    pGeoDataDescriptor.CreateFromSelectionSet pSelectionSet, Nothing, pFieldName
    
    Dim pGeoDS As IGeoDataset
    Set pGeoDS = pGeoDataDescriptor
    ' Delete Old Files
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    If fsObj.FolderExists(gMapTempFolder & "\" & pFileName) Then
        fsObj.DeleteFolder gMapTempFolder & "\" & pFileName
    End If
    
    
    Dim pRasterPropsDEM As IRasterProps
    If Not pDEMRLayer Is Nothing Then
        Set pRasterPropsDEM = pDEMRLayer.Raster
    End If
    
    Dim pRasterLUProps As IRasterProps
    Set pRasterLUProps = GetInputRasterLayer("Landuse").Raster
    
    '*** Create the conversion object
    Dim pConvert As IConversionOp
    Set pConvert = New RasterConversionOp
    
    'Get the conversion environment
    Dim pEnv As IRasterAnalysisEnvironment
    Dim pCellSize As Double
    Set pEnv = pConvert
    If Not pDEMRLayer Is Nothing Then
        Set pEnv.OutSpatialReference = pRasterPropsDEM.SpatialReference
        pEnv.SetExtent esriRasterEnvValue, pRasterPropsDEM.Extent
        pCellSize = pRasterPropsDEM.MeanCellSize.X
        pEnv.SetCellSize esriRasterEnvValue, pCellSize
        Set pEnv.Mask = pDEMRLayer.Raster
    Else
        Set pEnv.OutSpatialReference = pRasterLUProps.SpatialReference
        pEnv.SetExtent esriRasterEnvValue, pRasterLUProps.Extent
        pCellSize = pRasterLUProps.MeanCellSize.X
        pEnv.SetCellSize esriRasterEnvValue, pCellSize
        Set pEnv.Mask = GetInputRasterLayer("Landuse").Raster
    End If
    
    'Create a new raster dataset
    Dim pConRaster As IRasterDataset
    Set pConRaster = pConvert.ToRasterDataset(pGeoDS, "GRID", pWs, pFileName)
    
    'Return the value
    Set ConvertFeatureToRaster = pConRaster
    GoTo CleanUp

ShowError:
    MsgBox "ConvertFeatureToRaster: " & Err.description & "Feature " & pFeatureclass.AliasName
CleanUp:
    Set pDEMRLayer = Nothing
    Set pWSF = Nothing
    Set pWs = Nothing
    Set pSelectionSet = Nothing
    Set pGeoDataDescriptor = Nothing
    Set pGeoDS = Nothing
    Set fsObj = Nothing
    Set pRasterPropsDEM = Nothing
    Set pRasterLUProps = Nothing
    Set pConvert = Nothing
    Set pEnv = Nothing
    Set pConRaster = Nothing
End Function


'******************************************************************************
'Subroutine: OpenShapeFile
'Author:     Mira Chokshi
'Purpose:    Open feature dataset (.shp) from disk and return the featureclass.
'******************************************************************************
Public Function OpenShapeFile(dir As String, name As String) As IFeatureClass
On Error GoTo ErrHandler

  Dim pWsFact As IWorkspaceFactory
  Dim ConnectionProperties As IPropertySet
  Dim pShapeWS As IFeatureWorkspace
  Dim isShapeWS As Boolean
  Set OpenShapeFile = Nothing
  Set pWsFact = New ShapefileWorkspaceFactory
  isShapeWS = pWsFact.IsWorkspace(dir)
  If (isShapeWS) Then
    Set ConnectionProperties = New PropertySet
    ConnectionProperties.SetProperty "DATABASE", dir
    Set pShapeWS = pWsFact.Open(ConnectionProperties, 0)
    Dim pFClass As IFeatureClass
    Set pFClass = pShapeWS.OpenFeatureClass(name)
    Set OpenShapeFile = pFClass
    Set pFClass = Nothing
  End If
  
GoTo CleanUp
ErrHandler:
    'This errhandler purposely has no error message.  Mira Chokshi 10/08/04
CleanUp:
    Set pWsFact = Nothing
    Set ConnectionProperties = Nothing
    Set pShapeWS = Nothing
    Set pFClass = Nothing
End Function


'******************************************************************************
'Subroutine: ConvertRasterToFeature
'Author:     Mira Chokshi
'Purpose:    Convert raster to feature dataset. Requires Raster object (memory
'            representation), field name used for converting and file name
'            of output. This functions creates a feature layer from the converted
'            feature class and returns a feature layer.
'******************************************************************************
Public Function ConvertRasterToFeature(pRaster As IRaster, pFieldName As String, pFileName As String, pFeatType As String) As IFeatureLayer
On Error GoTo ShowError

    'Create a workspace
    Dim pWSF As IWorkspaceFactory
    Dim pWs As IWorkspace
    Set pWSF = New ShapefileWorkspaceFactory
    Set pWs = pWSF.OpenFromFile(gMapTempFolder, 0)
    Dim pDS As IDataset
    Dim pFClass As IFeatureClass
    Set pFClass = OpenShapeFile(gMapTempFolder, pFileName)
    If (Not pFClass Is Nothing) Then
      Set pDS = pFClass
      pDS.Delete
    End If
    
    ' Delete Old Files
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    If fsObj.FileExists(gMapTempFolder & "\" & pFileName) Then
        fsObj.DeleteFile gMapTempFolder & "\" & Replace(pFileName, "shp", "*")
    End If
    
    ' Create RasterDecriptor
    Dim pRDescr As IRasterDescriptor
    Set pRDescr = New RasterDescriptor
    pRDescr.Create pRaster, Nothing, pFieldName
     ' Create ConversionOp
    Dim pConversionOp As IConversionOp
    Set pConversionOp = New RasterConversionOp
     ' Perform conversion
    Dim pOutFClass As IFeatureClass
    Select Case pFeatType
        Case "Point"
            Set pOutFClass = pConversionOp.RasterDataToPointFeatureData(pRDescr, pWs, pFileName)
        Case "Polygon"
            Set pOutFClass = pConversionOp.RasterDataToPolygonFeatureData(pRDescr, pWs, pFileName, True)
    End Select
    
    ' Create a feature layer
    Dim pOutFLayer As IFeatureLayer
    Set pOutFLayer = New FeatureLayer
    Set pOutFLayer.FeatureClass = pOutFClass
    ' Return the feature layer
    Set ConvertRasterToFeature = pOutFLayer
    GoTo CleanUp
    
ShowError:
    MsgBox "ConvertRasterToFeature: " & Err.description
CleanUp:
    Set pWSF = Nothing
    Set pWs = Nothing
    Set pDS = Nothing
    Set pFClass = Nothing
    Set fsObj = Nothing
    Set pRDescr = Nothing
    Set pConversionOp = Nothing
    Set pOutFClass = Nothing
    Set pOutFLayer = Nothing
End Function


'******************************************************************************
'Subroutine: OpenRasterDatasetFromDisk
'Author:     Mira Chokshi
'Purpose:    Opens raster dataset from disk. Requires the name of the raster
'            dataset. This function does not require the directory path.
'            It assums the directory path as TEMP directory.
'******************************************************************************
Public Function OpenRasterDatasetFromDisk(pRasterName As String) As IRaster
On Error GoTo ShowError
    ' check if raster dataset exist
    Dim fsObj As Scripting.FileSystemObject
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    If Not fsObj.FolderExists(gMapTempFolder & "\" & pRasterName) Then
        Set OpenRasterDatasetFromDisk = Nothing
        GoTo CleanUp
    End If
    Set fsObj = Nothing
          
    'Open workspace
    Dim pWF As IWorkspaceFactory
    Set pWF = New RasterWorkspaceFactory
    Dim pRW As IRasterWorkspace
    Set pRW = pWF.OpenFromFile(gMapTempFolder, 0)
    Dim pRDS As IRasterDataset
    If (pRW.IsWorkspace(gMapTempFolder)) Then
      Set pRDS = pRW.OpenRasterDataset(LCase(pRasterName))
    End If
    If pRDS Is Nothing Then
      GoTo CleanUp
    End If
    'Get Raster from the raster dataset
    Dim pRaster As IRaster
    Set pRaster = pRDS.CreateDefaultRaster
    'Return raster
    Set OpenRasterDatasetFromDisk = pRaster
    GoTo CleanUp
ShowError:
    MsgBox "OpenRasterDatasetFromDisk: " & Err.description
CleanUp:
    Set fsObj = Nothing
    Set pWF = Nothing
    Set pRW = Nothing
    Set pRDS = Nothing
    Set pRaster = Nothing
End Function


'******************************************************************************
'Subroutine: WriteRasterDatasetToDisk
'Author:     Mira Chokshi
'Purpose:    Writes the temporary raster (in memory) to the disk. This function
'            requires the name of raster file, assumes the output directory as
'            TEMP directory.
'******************************************************************************
Public Sub WriteRasterDatasetToDisk(ByRef pRaster As IRaster, pOutName As String)
On Error GoTo ShowError
    ' Create a raster workspace
    Dim pRWS As IRasterWorkspace
    Dim pWSF As IWorkspaceFactory
    Set pWSF = New RasterWorkspaceFactory
    Set pRWS = pWSF.OpenFromFile(gMapTempFolder, 0)
    'Delete the raster dataset if present on disk
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(gMapTempFolder & "\" & pOutName)) Then
        Dim pRasterDataset As IRasterDataset
        Set pRasterDataset = pRWS.OpenRasterDataset(pOutName)
        If Not (pRasterDataset Is Nothing) Then
            Dim pDataset As IDataset
            Set pDataset = pRasterDataset
            If pDataset.CanDelete Then pDataset.Delete
            Set pDataset = Nothing
            Set pRasterDataset = Nothing
        End If
    End If
    ' SaveAs the projected raster
    Dim pDS As IDataset
    Dim pRasBandCol As IRasterBandCollection
    Set pRasBandCol = pRaster
    Set pDS = pRasBandCol.SaveAs(pOutName, pRWS, "GRID")
    GoTo CleanUp
ShowError:
    MsgBox "WriteRasterDatasetToDisk: " & Err.description, vbExclamation, pOutName
CleanUp:
    Set pRWS = Nothing
    Set pWSF = Nothing
    Set fso = Nothing
    Set pRasterDataset = Nothing
    Set pDataset = Nothing
    Set pDS = Nothing
    Set pRasBandCol = Nothing
End Sub


'******************************************************************************
'Subroutine: SelectLayerOrTableData
'Author:     Mira Chokshi
'Purpose:    This is a helper function to let the user select input data
'            using dialog interface. It provides filters for Feature, Raster,
'            Table(.dbf) Dataset. This function requires the input specifying
'            the type of filter.
'******************************************************************************
Public Function SelectLayerOrTableData(datasetType As String) As String

On Error GoTo ErrorHandler

  Dim pDlg As IGxDialog
  Dim pGXSelect As IEnumGxObject
  Dim pGxObject As IGxObject
  Dim pGXDataset As IGxDataset
  Dim pGXDatabase As IGxDatabase
  Dim pEnumDataset As IEnumDataset
  Dim pFeatWS As IFeatureWorkspace
  Dim pFeatCls As IFeatureClass
  Dim pFeatLyr As IFeatureLayer
  Dim className As String
  Dim pFeatClsCont As IFeatureClassContainer
  Dim pObjectFilter As IGxObjectFilter
  Dim i As Long
  Dim pActiveView As IActiveView
  Dim pFDOGraphicsLayerFactory As IFDOGraphicsLayerFactory

  Set pActiveView = gMap
  ' set up filters on the files that will be browsed
  Set pDlg = New GxDialog
  If (datasetType = "Feature") Then
    Set pObjectFilter = New GxFilterFeatureClasses
  ElseIf (datasetType = "GeoTable") Then
    Set pObjectFilter = New GxFilterDatasetsAndLayers
  ElseIf (datasetType = "GeoDB") Then
    Set pObjectFilter = New GxFilterGeoDatasets
    'if the dataset is of raster type
    '-- Sabu Paul, July 14, 2004
  ElseIf (datasetType = "Raster") Then
    Set pObjectFilter = New GxFilterRasterDatasets
  End If

  pDlg.AllowMultiSelect = False
  pDlg.Title = "Select Data"
  Set pDlg.ObjectFilter = pObjectFilter

  If (pDlg.DoModalOpen(pActiveView.ScreenDisplay.hWnd, pGXSelect) = False) Then Exit Function

    ' got a valid selection from the GX Dialog, now extract the feature classes datasets etc.
    ' loop through the selection enumeration
    pGXSelect.Reset
    Set pGxObject = pGXSelect.Next
    If (Not pGxObject Is Nothing) Then
        ' We could be handed objects of various types, work out what types we have been handed and then open
        ' them up and add a feature layer to handle them
        Set pGXDataset = pGxObject
        If (TypeOf pGxObject Is IGxDataset) Then
          If (pGXDataset.Type = esriDTFeatureClass) Then
                Set pFeatCls = pGXDataset.Dataset
                If pFeatCls.FeatureType = esriFTAnnotation Then
                  Set pFDOGraphicsLayerFactory = New FDOGraphicsLayerFactory
                  Set pFeatLyr = pFDOGraphicsLayerFactory.OpenGraphicsLayer(pFeatCls.FeatureDataset.Workspace, pFeatCls.FeatureDataset, pFeatCls.AliasName)
                Else
                  Set pFeatLyr = New FeatureLayer
                  Set pFeatLyr.FeatureClass = pFeatCls
                  pFeatLyr.name = pFeatCls.AliasName
                  pFeatLyr.Visible = False
                  gMap.AddLayer pFeatLyr
                End If
                SelectLayerOrTableData = pFeatCls.AliasName
        ElseIf (pGXDataset.Type = esriDTTable) Then
                Dim pTable As iTable
                Set pTable = pGXDataset.Dataset
                Dim pTableColl As ITableCollection
                Set pTableColl = gMap
                pTableColl.AddTable pTable
                SelectLayerOrTableData = pGXDataset.datasetname.name
        ElseIf (pGXDataset.Type = esriDTGeo) Then
                SelectLayerOrTableData = pGXDataset.datasetname.name
        'if the dataset is of raster type
        '-- Sabu Paul, July 14, 2004
        ElseIf (pGXDataset.Type = esriDTRasterDataset) Then
            Dim pRasterLayer As IRasterLayer
            Set pRasterLayer = New RasterLayer
            pRasterLayer.CreateFromDataset pGXDataset.Dataset
            gMap.AddLayer pRasterLayer
            gMxDoc.ActiveView.Refresh
            SelectLayerOrTableData = pGXDataset.datasetname.name
        End If
     End If
     Set pGxObject = pGXSelect.Next
  End If

  Exit Function
ErrorHandler:
  MsgBox "SelectLayerOrTableData: " & Err.Number & "  " & Err.Source & "  " & Err.description

End Function



'******************************************************************************
'Subroutine: OpenAccessWorkspace
'Author:     Sabu Paul
'Purpose:    Open microsoft access (.mdb) workspace and returns it. Requires
'            the user the specify the file path (path of mdb file)
'******************************************************************************
Public Function OpenAccessWorkspace(ConnString As String) As IWorkspace
On Error GoTo ShowError
    Dim pWs As IWorkspace
    Dim pWorkspaceFactory As IWorkspaceFactory
    
    Set OpenAccessWorkspace = Nothing
    Set pWorkspaceFactory = New AccessWorkspaceFactory
    Set pWs = pWorkspaceFactory.OpenFromFile(ConnString, 0)
    Set OpenAccessWorkspace = pWs
    GoTo CleanUp
ShowError:
    MsgBox "OpenAccessWorkspace: " & Err.description
CleanUp:
    Set pWs = Nothing
    Set pWorkspaceFactory = Nothing
End Function

'******************************************************************************
'Subroutine: SetAdoConn
'Author:     Sabu Paul
'Purpose:    Sets the Ado Connection to the specified access database.
'******************************************************************************
Public Sub SetAdoConn(ConnStr As String)
On Error GoTo ShowError
    Dim pWs As IWorkspace
    Set gAdoConn = New ADODB.Connection
    gAdoConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnStr & ";"
    Set pWs = OpenAccessWorkspace(ConnStr)
    GoTo CleanUp
ShowError:
    MsgBox "SetAdoConn: " & Err.description
CleanUp:
    Set pWs = Nothing
    Set gAdoConn = Nothing
End Sub


'******************************************************************************
'Subroutine: CleanUpMemory
'Author:     Mira Chokshi
'Purpose:    Clears the memory by setting all public variables to nothing
'******************************************************************************
Public Sub CleanUpMemory()
    Set gMxDoc = Nothing
    Set gMap = Nothing
    Set gDEMRaster = Nothing
    Set gReclassOp = Nothing
    Set gNeighborhoodOp = Nothing
    Set gAlgebraOp = Nothing
    Set gHydrologyOp = Nothing
    Set gRasterDistanceOp = Nothing
    Set gAdoConn = Nothing
    Set gBMPDetailDict = Nothing
    Set gBMPDictionary = Nothing
    Set gSubWaterLandUseDict = Nothing
End Sub


'******************************************************************************
'Subroutine: BrowseForFolder
'Author:     Sabu Paul
'Purpose:    Function calling system functions to browse for directory folder
'******************************************************************************
Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String

    Dim nNull As Integer
    Dim lpIDList As Long
    Dim nResult As Long
    Dim sPath As String
    Dim bi As BrowseInfo

    bi.hWndOwner = hWndOwner
    bi.lpszTitle = lstrcat(sPrompt, "")
    bi.ulFlags = BIF_RETURNONLYFSDIRS

    lpIDList = SHBrowseForFolder(bi)
    
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        nResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        nNull = InStr(sPath, vbNullChar)
        If nNull Then
            sPath = Left$(sPath, nNull - 1)
        End If
    End If

    BrowseForFolder = sPath

End Function


'******************************************************************************
'Subroutine: ValidateDataSource
'Author:     Sabu Paul
'Purpose:    Define Input Data sources
'******************************************************************************
Public Function ValidateDataSource() As Boolean
On Error GoTo ShowError
    'Set it to FALSE initially
    ValidateDataSource = False
    If (gLayerNameDictionary Is Nothing) Then
        Exit Function
    End If
    
    Dim pRasterLayer As IRasterLayer
    Set pRasterLayer = GetInputRasterLayer("Landuse")
    If (pRasterLayer Is Nothing) Then
        Exit Function
    End If
    Dim pDataTable As iTable
    Set pDataTable = GetInputDataTable("lulookup")
    If (pDataTable Is Nothing) Then
        Exit Function
    End If
    
    'gMapTempFolder = gLayerNameDictionary.Item("TEMP")
    If (gLayerNameDictionary.Item("Landuse") = "" Or gLayerNameDictionary.Item("lulookup") = "" Or gMapTempFolder = "") Then
        Exit Function
    Else
        'Checks for projection of all input layers
        ValidateDataSource = CheckInputDataProjection
        'Get meters per unit factor
        GetMetersPerLinearUnit
    End If
    
    'Added deflayers to true - Sabu Paul, Dec 5, 2008
    gDefLayers = True
    ValidateDataSource = True
    Exit Function
ShowError:
    MsgBox "ValidateDataSource: " & Err.description
End Function



'******************************************************************************
'Subroutine: CheckInputDataProjection
'Author:     Mira Chokshi
'Purpose:    Checks the input projection of input layers. Returns a FALSE if
'            all input layers are not in same projection
'******************************************************************************
Public Function CheckInputDataProjection() As Boolean
On Error GoTo ShowError
    Dim pLayer As ILayer
    Dim pSpatialReference1 As ISpatialReference
    Dim pSpatialReference2 As ISpatialReference
    Dim pSpatialReference3 As ISpatialReference
    'check the projection of only dem, STREAM & landuse
    Dim i As Integer
    Dim pLayerName As String
    FrmErrors.ListProjections.Clear
    
    Dim pSpatialReferenceName As String
    
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        Select Case pLayer.name
            Case gLayerNameDictionary.Item("DEM"):
                Set pSpatialReference1 = GetSpatialReferenceForLayer(pLayer)
                If (pSpatialReference1 Is Nothing) Then
                    pSpatialReferenceName = "Undefined"
                Else
                    pSpatialReferenceName = pSpatialReference1.name
                End If
                FrmErrors.ListProjections.AddItem "Layer: " & gLayerNameDictionary.Item("DEM") & vbTab & " Projection: " & pSpatialReferenceName
            Case gLayerNameDictionary.Item("STREAM"):
                Set pSpatialReference2 = GetSpatialReferenceForLayer(pLayer)
                If (pSpatialReference2 Is Nothing) Then
                    pSpatialReferenceName = "Undefined"
                Else
                    pSpatialReferenceName = pSpatialReference2.name
                End If
                FrmErrors.ListProjections.AddItem "Layer: " & gLayerNameDictionary.Item("STREAM") & vbTab & " Projection: " & pSpatialReferenceName
            Case gLayerNameDictionary.Item("Landuse"):
                Set pSpatialReference3 = GetSpatialReferenceForLayer(pLayer)
                If (pSpatialReference3 Is Nothing) Then
                    pSpatialReferenceName = "Undefined"
                Else
                    pSpatialReferenceName = pSpatialReference3.name
                End If
                FrmErrors.ListProjections.AddItem "Layer: " & gLayerNameDictionary.Item("Landuse") & vbTab & " Projection: " & pSpatialReferenceName
        End Select
     Next
     
     CheckInputDataProjection = True
     '** Check landuse projection, if nothing, set false
     '** If DEM is present, check project, if nothing, set false, else compare its proj to landuse proj
     '** If STREAM is present, check project, if nothing, set false, else compare its proj to landuse proj
     '** modified by Mira Chokshi 04/13/2005

    If Not GetInputRasterLayer(gLayerNameDictionary.Item("Landuse")) Is Nothing Then
            If (pSpatialReference3 Is Nothing) Then
                CheckInputDataProjection = False
            End If
            
            If Not GetInputRasterLayer(gLayerNameDictionary.Item("DEM")) Is Nothing Then
                If (pSpatialReference1 Is Nothing) Then
                    CheckInputDataProjection = False
                End If
                
                If (CheckInputDataProjection = True) Then
                    If (pSpatialReference3.FactoryCode <> pSpatialReference1.FactoryCode) Then
                        CheckInputDataProjection = False
                    End If
                End If
            End If
                
            If Not GetInputRasterLayer(gLayerNameDictionary.Item("STREAM")) Is Nothing Then
                If (pSpatialReference2 Is Nothing) Then
                    CheckInputDataProjection = False
                End If
                
                If (CheckInputDataProjection = True) Then
                    If (pSpatialReference3.FactoryCode <> pSpatialReference2.FactoryCode) Then
                        CheckInputDataProjection = False
                    End If
                End If
            End If
    Else
        CheckInputDataProjection = False
    End If
    
    GoTo CleanUp
ShowError:
    MsgBox "CheckInputDataProjection: " & Err.description
CleanUp:
    Set pLayer = Nothing
    Set pSpatialReference1 = Nothing
    Set pSpatialReference2 = Nothing
    Set pSpatialReference3 = Nothing
End Function


'******************************************************************************
'Subroutine: GetSpatialReferenceForLayer
'Author:     Mira Chokshi
'Purpose:    Checks the input projection of an input layer. Checks the type of
'            input layer (feature/raster) and gets its spatial reference
'******************************************************************************
Public Function GetSpatialReferenceForLayer(pLayer As ILayer) As ISpatialReference
On Error GoTo ShowError

    Dim pFeatureLayer As IFeatureLayer
    Dim pRasterLayer As IRasterLayer
    If (TypeOf pLayer Is IFeatureLayer) Then
        Set pFeatureLayer = pLayer
        Dim pGeoDataset As IGeoDataset
        Set pGeoDataset = pFeatureLayer.FeatureClass
        Set GetSpatialReferenceForLayer = pGeoDataset.SpatialReference
    ElseIf (TypeOf pLayer Is IRasterLayer) Then
        Set pRasterLayer = pLayer
        Dim pRasterProps As IRasterProps
        Set pRasterProps = pRasterLayer.Raster
        Set GetSpatialReferenceForLayer = pRasterProps.SpatialReference
    End If
    GoTo CleanUp
ShowError:
    MsgBox "GetSpatialReferenceForLayer: " & Err.description
CleanUp:
    Set pFeatureLayer = Nothing
    Set pRasterLayer = Nothing
    Set pGeoDataset = Nothing
    Set pRasterProps = Nothing
End Function

Public Sub GetMetersPerLinearUnit()
On Error GoTo ShowError
    Dim pLandUseRLayer As IRasterLayer
    Set pLandUseRLayer = GetInputRasterLayer("Landuse")
    If (pLandUseRLayer Is Nothing) Then
        GoTo CleanUp
    End If
    Dim pSpatialReference As ISpatialReference
    Set pSpatialReference = GetSpatialReferenceForLayer(pLandUseRLayer)
    
    If (pSpatialReference Is Nothing) Then
        Exit Sub
    End If
    
    Dim pGeoCoordSys As IGeographicCoordinateSystem
    Dim pProjectedCoordSys As IProjectedCoordinateSystem
    Dim pLinearUnit As ILinearUnit
    If (TypeOf pSpatialReference Is IGeographicCoordinateSystem) Then
        Set pGeoCoordSys = pSpatialReference
        Set pLinearUnit = pGeoCoordSys.CoordinateUnit
    ElseIf (TypeOf pSpatialReference Is IProjectedCoordinateSystem) Then
        Set pProjectedCoordSys = pSpatialReference
        Set pLinearUnit = pProjectedCoordSys.CoordinateUnit
    End If
 
    If (Not pLinearUnit Is Nothing) Then
        gMetersPerUnit = pLinearUnit.MetersPerUnit
        gLinearUnitName = LCase(pLinearUnit.name)
    Else
        gMetersPerUnit = 1
        gLinearUnitName = "units"
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "GetMetersPerLinearUnit: " & Err.description
CleanUp:
    Set pLinearUnit = Nothing
    Set pSpatialReference = Nothing
    Set pLandUseRLayer = Nothing
End Sub


'******************************************************************************
'Subroutine: CheckInputTableFieldFormats
'Author:     Mira Chokshi
'Purpose:    Checks the input projection of an input layer. Checks the type of
'            input layer (feature/raster) and gets its spatial reference
'******************************************************************************
Public Function CheckInputTableFieldFormats() As Boolean
On Error GoTo ShowError
    'Only one table is required right now, landuse lookup table with
    'landuse code and landuse name
    Dim StrMsgError As String
    Dim pInputTable1 As iTable
    Set pInputTable1 = GetInputDataTable("lulookup")
    Dim iLuCode As Long
    iLuCode = pInputTable1.FindField("LUCODE")
    FrmErrors.ListProjections.Clear
    If (iLuCode < 0) Then
        StrMsgError = "Landuse lookup table requires a LUCODE field."
        FrmErrors.ListProjections.AddItem StrMsgError
    End If
    Dim iLUName As Long
    iLUName = pInputTable1.FindField("LUNAME")
    If (iLUName < 0) Then
        StrMsgError = "Landuse lookup table requires a LUNAME field."
        FrmErrors.ListProjections.AddItem StrMsgError
    End If
    
    '*** Check for FID field in STREAM feature layer
    Dim pFeatureLayer As IFeatureLayer
    Set pFeatureLayer = GetInputFeatureLayer("STREAM")
    If Not (pFeatureLayer Is Nothing) Then
        Dim iFIDLong As Long
        iFIDLong = pFeatureLayer.FeatureClass.FindField("FID")
        If (iFIDLong < 0) Then
            iFIDLong = pFeatureLayer.FeatureClass.FindField("OBJECTID")
            If (iFIDLong < 0) Then
                StrMsgError = "STREAM feature layer requires a FID or ObjectID field."
                FrmErrors.ListProjections.AddItem StrMsgError
            End If
        End If
    End If
    
    
    If (FrmErrors.ListProjections.ListCount > 0) Then
        CheckInputTableFieldFormats = False
    Else
        CheckInputTableFieldFormats = True
    End If
    GoTo CleanUp
ShowError:
    MsgBox "CheckInputTableFieldFormats: " & Err.description
CleanUp:
    Set pInputTable1 = Nothing
End Function

'******************************************************************************
'Subroutine: BubbleSort
'Author:     Sabu Paul
'Purpose:    Sorts the array of integers
'******************************************************************************

Public Sub BubbleSort(arr As Variant, Optional numEls As Variant, Optional descending As Boolean)
    
    Dim value As Variant
    Dim Index As Long
    Dim firstItem As Long
    Dim indexLimit As Long, lastSwap As Long
    
    ' account for optional arguments
    If IsMissing(numEls) Then
        numEls = UBound(arr)
    End If
    firstItem = LBound(arr)
    lastSwap = numEls
    
    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For Index = firstItem To indexLimit
            value = arr(Index)
            If (value > arr(Index + 1)) Xor descending Then
                ' if the items are not in order, swap them
                arr(Index) = arr(Index + 1)
                arr(Index + 1) = value
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
End Sub


'******************************************************************************
'Subroutine: Pause
'Author:     Sabu Paul
'Purpose:    Pauses the process for nSeconds
'******************************************************************************
Public Sub Pause(ByVal nSecond As Single)
   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer
      dummy = DoEvents()
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop
End Sub



'******************************************************************************
'Subroutine: WriteLayerTagDictionaryToSRCFile
'Author:     Mira Chokshi
'Purpose:    Write all the layers in layer name dictionary to file
'******************************************************************************
Public Sub WriteLayerTagDictionaryToSRCFile()
    
    'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = gApplicationPath
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not (fso.FolderExists(gMapTempFolder)) Then
        fso.CreateFolder gMapTempFolder
    End If
    
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = gApplication.Document
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & ".src"
    dataSrcFN = pAppPath & dataSrcFN
    
    Dim pDataSrcFile As TextStream
    Set pDataSrcFile = fso.OpenTextFile(dataSrcFN, ForWriting, True, TristateUseDefault)

    Dim pKeys
    pKeys = gLayerNameDictionary.keys
    Dim pkey As String
    Dim ikey As Integer
    For ikey = 0 To gLayerNameDictionary.Count - 1
        pkey = pKeys(ikey)
        pDataSrcFile.WriteLine pkey & vbTab & gLayerNameDictionary.Item(pkey)
    Next
    
    pDataSrcFile.Close
    Set pDataSrcFile = Nothing
    Set fso = Nothing
    
    
End Sub



'******************************************************************************
'Subroutine: ReadLayerTagDictionaryToSRCFile
'Author:     Mira Chokshi
'Purpose:    Read all the layers from input src file to layer name dictionary
'******************************************************************************
Public Sub ReadLayerTagDictionaryToSRCFile()
    
    'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = gApplicationPath
   
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = gApplication.Document
       
    '*** Get the complete path of the application
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & ".src"
    dataSrcFN = pAppPath & dataSrcFN
    
    'Create a layer name dictionary
    Set gLayerNameDictionary = CreateObject("Scripting.Dictionary")
    gLayerNameDictionary.RemoveAll
    
    Call SetDataDirectory
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
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
        If UBound(lWords) > 0 Then gLayerNameDictionary.add lWords(0), lWords(1)
    Next lIncr
    pDataSrcFile.Close
    
    'Assign the temp folder
    'gMapTempFolder = gLayerNameDictionary.Item("TEMP")
    
    If (gLayerNameDictionary.Exists("SUBBASIN")) Then
        gSUBBASINFieldName = gLayerNameDictionary.Item("SUBBASIN")
    End If
    If (gLayerNameDictionary.Exists("SUBBASINR")) Then
        gSUBBASINRFieldName = gLayerNameDictionary.Item("SUBBASINR")
    End If
        
    'Check if watershed feature layer is present, and set gManualDelineationFlag = True
    If (gLayerNameDictionary.Exists("Watershed")) Then
        gManualDelineationFlag = True
    End If
        
    'Set the simulation options - Sabu Paul- June 14, 2007
    If gLayerNameDictionary.Exists("SimulationOption") Then
        If gLayerNameDictionary.Item("SimulationOption") = "External" Then
            gExternalSimulation = True
            gInternalSimulation = False
        Else
            gExternalSimulation = False
            gInternalSimulation = True
        End If
    End If
    
    
    Set pDataSrcFile = Nothing
    Set fso = Nothing
    
End Sub


Public Function CheckMapDocumentSavedStatus() As Boolean

    CheckMapDocumentSavedStatus = False
    
    'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Dim pAppPath As String
    pAppPath = gApplicationPath
   
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = gApplication.Document
    
    'Check if the .mxd is saved, if not force the user to save the .mxd
    If (Replace(dataSrcFN, ".mxd", "") = "Untitled") Then
        MsgBox "Please save .mxd file to continue. ", vbExclamation
        Exit Function
    End If
        
    CheckMapDocumentSavedStatus = True
    
End Function

Public Sub FlashLine(pDisplay As IScreenDisplay, pGeometry As IGeometry)
  Dim pLineSymbol As ISimpleLineSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pLineSymbol = New SimpleLineSymbol
  pLineSymbol.Width = 4
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pLineSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pLineSymbol
  pDisplay.DrawPolyline pGeometry
  Sleep 300
  pDisplay.DrawPolyline pGeometry
End Sub

Public Sub FlashPolygon(pDisplay As IScreenDisplay, pGeometry As IGeometry)
  Dim pFillSymbol As ISimpleFillSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pFillSymbol = New SimpleFillSymbol
  pFillSymbol.Outline = Nothing
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pFillSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pFillSymbol
  pDisplay.DrawPolygon pGeometry
  Sleep 300
  pDisplay.DrawPolygon pGeometry
End Sub

Public Sub FlashPoint(pDisplay As IScreenDisplay, pGeometry As IGeometry)
  Dim pMarkerSymbol As ISimpleMarkerSymbol
  Dim pSymbol As ISymbol
  Dim pRGBColor As IRgbColor
  
  Set pMarkerSymbol = New SimpleMarkerSymbol
  pMarkerSymbol.Style = esriSMSSquare
  pMarkerSymbol.Size = 12
  pMarkerSymbol.XOffset = 0
  pMarkerSymbol.YOffset = 0
  
  Set pRGBColor = New RgbColor
  pRGBColor.Green = 128
  
  Set pSymbol = pMarkerSymbol
  pSymbol.ROP2 = esriROPNotXOrPen
  
  pDisplay.SetSymbol pMarkerSymbol
  pDisplay.DrawPoint pGeometry
  Sleep 300
  pDisplay.DrawPoint pGeometry
End Sub

Public Sub FlashSelectedFeature(pFeature As IFeature)
  
  ' Start Drawing on screen
  Dim pActiveView As IActiveView
  Set pActiveView = gMxDoc.ActiveView
  pActiveView.ScreenDisplay.StartDrawing 0, esriNoScreenCache
  
  ' Switch functions based on Geomtry type
  Select Case pFeature.Shape.GeometryType
    Case esriGeometryPolyline
      FlashLine pActiveView.ScreenDisplay, pFeature.Shape
    Case esriGeometryPolygon
      FlashPolygon pActiveView.ScreenDisplay, pFeature.Shape
    Case esriGeometryPoint
      FlashPoint pActiveView.ScreenDisplay, pFeature.Shape
  End Select
  
  ' Finish drawing on screen
  pActiveView.ScreenDisplay.FinishDrawing

  Set pActiveView = Nothing
End Sub


Public Sub DeactivateCurrentTool()

    ' This example makes the builtin Select Graphics Tool the active tool.
    Dim pSelectTool As ICommandItem
    Dim pCommandBars As ICommandBars
    ' The identifier for the Select Graphics Tool
    'Find the Select Graphics Tool
    Set pCommandBars = gApplication.Document.CommandBars
    Set pSelectTool = pCommandBars.Find("esriSurveyExt.SelectTool")
    'Set the current tool of the application to be the Select Graphics Tool
    Set gApplication.CurrentTool = pSelectTool
End Sub

'******************************************************************************
'Subroutine: GetApplicationPath
'Author:     Sabu Paul
'Purpose:    Reads the registry to get the SUSTAIN application path
'******************************************************************************

Public Function GetApplicationPath() As String
'*********************************************************************
'***The module passes the path information for the ArcView Software***
'*********************************************************************

Dim hKey As Long  ' receives a handle to the newly created or opened registry key
    
    Dim subkey As String  ' name of the subkey to open
    Dim stringbuffer As String  ' receives data read from the registry
    Dim datatype As Long  ' receives data type of read value
    Dim slength As Long  ' receives length of returned data
    Dim retval As Long  ' return value

    ' Set the name of the new key and the default security settings
    subkey = "Software\Tetra Tech\SUSTAIN"

    ' Create or open the registry key
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, subkey, 0, KEY_READ, hKey)
    If retval <> 0 Then
        Debug.Print "ERROR: Unable to open registry key!"
        Exit Function
    End If

    ' Make room in the buffer to receive the incoming data.
    stringbuffer = Space(255)
    slength = 255
    ' Read the "InstallPath" value from the registry key.
    retval = RegQueryValueEx(hKey, "InstallPath", 0, datatype, ByVal stringbuffer, slength)
    ' Only attempt to display the data if it is in fact a string.
    If datatype = REG_SZ Then
        ' Remove empty space from the buffer and display the result.
        stringbuffer = Left(stringbuffer, slength - 1)
    Else
        ' Don't bother trying to read any other data types.
        Debug.Print "Data not in string format.  Unable to interpret data."
    End If

    ' Close the registry key.
    retval = RegCloseKey(hKey)
    
    GetApplicationPath = stringbuffer
    
End Function

Public Function Delete_Raster(dir As String, name As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim pRDS As IRasterDataset
    Set pRDS = OpenRasterDataset(dir, name)
    If pRDS Is Nothing Then Exit Function
    
    Dim pDS As IDataset
    Set pDS = pRDS
    If (pDS.CanDelete) Then
        pDS.Delete
        Delete_Raster = True
    Else
        Delete_Raster = False
    End If

Exit Function
ErrorHandler:
  MsgBox "Delete_Raster Dataset Error :" + Err.description
End Function

Public Function OpenRasterDataset(sDir As String, sRasterDs As String) As IRasterDataset
' Open raster dataset in a workspace
On Error GoTo er
    Dim pWsFact As IWorkspaceFactory
    Dim pWs As IRasterWorkspace
    
    Set pWsFact = New RasterWorkspaceFactory
    Set pWs = pWsFact.OpenFromFile(sDir, 0)
    Dim pRaster As IRaster
    Set pRaster = OpenRasterDatasetFromDisk(sRasterDs)
    If pRaster Is Nothing Then Exit Function
    Set OpenRasterDataset = pWs.OpenRasterDataset(sRasterDs)
    
    Set pWsFact = Nothing
    Set pWs = Nothing
    Exit Function
er:
MsgBox "Open Raster Dataset Error :" + Err.description

End Function


' ########################################################################

'******************************************************************************
'Subroutine: SetDataDirectory
'Author:     Sabu Paul
'Purpose:    Check the data directory path and sets
'******************************************************************************

Public Sub SetDataDirectory()
  On Error GoTo ErrorHandler

    'Create a file for writing the datasources -- Sabu Paul; Aug 24, 2004
    Call DefineApplicationPath
    Dim pAppPath As String
    pAppPath = gApplicationPath
    
    'Create a layer name dictionary
'    Set gLayerNameDictionary = CreateObject("Scripting.Dictionary")
'    gLayerNameDictionary.RemoveAll
    
    Dim dataSrcFN As String 'Sabu Paul -- October 2004
    dataSrcFN = gApplication.Document.Title
       
    '*** Get the complete path of the application
    dataSrcFN = Replace(dataSrcFN, ".mxd", "") & "_data.src"
    dataSrcFN = pAppPath & dataSrcFN
 
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
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
            'gLayerNameDictionary.add lWords(0), lWords(1)
            If UBound(lWords) > 0 Then gLayerNameDictionary.add lWords(0), lWords(1)
        End If
    Next lIncr
    pDataSrcFile.Close
    
    'Assign the temp folder
    gCostDBpath = gLayerNameDictionary.Item("gCostDBpath")
    gGDBpath = gLayerNameDictionary.Item("gGDBpath")
    gMapTempFolder = gLayerNameDictionary.Item("gMapTempFolder")
    
    gLayerNameDictionary.RemoveAll
    
    If Trim(gGDBpath) <> "" Then gGDBFlag = True

  Exit Sub
ErrorHandler:
  HandleError True, "SetDataDirectory " & c_sModuleFileName, Err.Number, Err.Source, Err.description, 1, m_ParentHWND
End Sub

Public Function SetComboItemIndex(cbxComponent As ComboBox, itemString As String) As Integer

    On Error GoTo ErrorHandler
    Dim pCompIndex As Integer
    For pCompIndex = cbxComponent.ListCount To 1 Step -1
        If cbxComponent.List(pCompIndex - 1) = itemString Then
            cbxComponent.ListIndex = pCompIndex - 1
            Exit For
        End If
    Next
    
'    If pCompIndex = 0 Then  'set the string as text for this combobox
'        cbxComponent.Text = itemString
'    End If
    
    SetComboItemIndex = pCompIndex
    Exit Function
ErrorHandler:
    HandleError True, "SetComboItemIndex " & c_sModuleFileName, Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Function
Public Function LoadData_FromPGDB(GDBPath As String) As Boolean

    On Error GoTo ErrorHandler
    
    LoadData_FromPGDB = False
    Dim pWorkspaceFactory As IWorkspaceFactory
    'Set pWorkspaceFactory = New AccessWorkspaceFactory
    Set pWorkspaceFactory = New FileGDBWorkspaceFactory
    Dim pWorkspace As IWorkspace
    Set pWorkspace = pWorkspaceFactory.OpenFromFile(GDBPath, 0)
    Dim pFeatureClassContainer As IFeatureClassContainer
    Dim pDataset As IDataset, pDataset2 As IDataset, pEnumDataset As IEnumDataset, pEnumDataset2 As IEnumDataset
        
    Set pEnumDataset = pWorkspace.Datasets(esriDTAny)
    Set pDataset = pEnumDataset.Next
    
    Dim pFeatCls As IFeatureClass
    Dim pFeatLyr As IFeatureLayer
    Dim pTable As iTable
    Dim pStTab As IStandaloneTable
    Dim pStTabColl As IStandaloneTableCollection
    Dim pchkLayer As ILayer
    
    Dim pRaster As IRaster
    Dim pRasterLyr As IRasterLayer
    'Clear all Layers.....
    'gMap.ClearLayers
    'Set pStTabColl = gMap
    'pStTabColl.RemoveAllStandaloneTables
    
    Set gFeatClassDictionary = CreateObject("Scripting.Dictionary") ' to store the FeatureClass Names.....
    gFeatClassDictionary.RemoveAll
    Dim pNewLyr As IFeatureLayer
    Dim pDefLyr As IFeatureLayerDefinition

    
        '
        Do Until pDataset Is Nothing
            If TypeOf pDataset Is IFeatureDataset Then
                Set pFeatureClassContainer = pDataset
                Set pEnumDataset2 = pFeatureClassContainer.Classes
                Set pDataset2 = pEnumDataset2.Next
                Do Until pDataset2 Is Nothing
                    Set pFeatCls = pDataset2
                    Set pFeatLyr = New FeatureLayer
                    Set pFeatLyr.FeatureClass = pFeatCls
                    pFeatLyr.name = pFeatCls.AliasName
                    Set pchkLayer = GetInputFeatureLayer(pFeatLyr.name)
                    If Not pchkLayer Is Nothing Then gMap.DeleteLayer pchkLayer
                    pFeatLyr.Visible = False
                    gMap.AddLayer pFeatLyr
                    gFeatClassDictionary.add pFeatLyr.name, "FeatureClass"
                    Set pDataset2 = pEnumDataset2.Next
                Loop
            ElseIf TypeOf pDataset Is IFeatureClass Then
                    Set pFeatCls = pDataset
                    Set pFeatLyr = New FeatureLayer
                    Set pFeatLyr.FeatureClass = pFeatCls
                    pFeatLyr.name = pFeatCls.AliasName
                    pFeatLyr.Visible = False
                    Set pchkLayer = GetInputFeatureLayer(pFeatLyr.name)
                    If Not pchkLayer Is Nothing Then gMap.DeleteLayer pchkLayer
                    gMap.AddLayer pFeatLyr
                    gFeatClassDictionary.add pFeatLyr.name, "FeatureClass"
            ElseIf TypeOf pDataset Is iTable Then
                    Set pStTab = New StandaloneTable
                    Set pStTabColl = gMap
                    Set pTable = GetInputDataTable(pDataset.name)
                    If pTable Is Nothing Then
                        Set pTable = pDataset
                        Set pStTab.Table = pTable
                        pStTabColl.AddStandaloneTable pStTab
                    End If
                    gFeatClassDictionary.add pDataset.name, "Table"
            'Add the raster datasets
            ElseIf TypeOf pDataset Is IRasterDataset Then
                'Set pRaster = pDataset
                Set pRasterLyr = New RasterLayer
                pRasterLyr.CreateFromDataset pDataset
                pRasterLyr.name = pDataset.name
                pRasterLyr.Visible = False
                Set pchkLayer = GetInputRasterLayer(pDataset.name)
                If Not pchkLayer Is Nothing Then gMap.DeleteLayer pchkLayer
                gMap.AddLayer pRasterLyr
                gFeatClassDictionary.add pRasterLyr.name, "Raster"
            End If
            Set pDataset = pEnumDataset.Next
        Loop
        
    ' Reorder the Layers....
    ReorderLayers
    LoadData_FromPGDB = True

Exit Function
ErrorHandler:
  HandleError True, "LoadData_FromPGDB " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND


End Function

Public Function CreateList_FromGDB(GDBPath As String) As Boolean

    On Error GoTo ErrorHandler
    
    CreateList_FromGDB = False
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New FileGDBWorkspaceFactory
    Dim pWorkspace As IWorkspace
    Set pWorkspace = pWorkspaceFactory.OpenFromFile(GDBPath, 0)
    Dim pFeatureClassContainer As IFeatureClassContainer
    Dim pDataset As IDataset, pDataset2 As IDataset, pEnumDataset As IEnumDataset, pEnumDataset2 As IEnumDataset
        
    Set pEnumDataset = pWorkspace.Datasets(esriDTAny)
    Set pDataset = pEnumDataset.Next
    
    Dim pFeatCls As IFeatureClass
'    Dim pFeatLyr As IFeatureLayer
'    Dim pTable As iTable
'    Dim pStTab As IStandaloneTable
'    Dim pStTabColl As IStandaloneTableCollection
'    Dim pchkLayer As ILayer
    
    Set gFeatClassDictionary = CreateObject("Scripting.Dictionary") ' to store the FeatureClass Names.....
    gFeatClassDictionary.RemoveAll
'    Dim pNewLyr As IFeatureLayer
'    Dim pDefLyr As IFeatureLayerDefinition

    
        '
        Do Until pDataset Is Nothing
            If TypeOf pDataset Is IFeatureDataset Then
                Set pFeatureClassContainer = pDataset
                Set pEnumDataset2 = pFeatureClassContainer.Classes
                Set pDataset2 = pEnumDataset2.Next
                Do Until pDataset2 Is Nothing
                    Set pFeatCls = pDataset2
'                    Set pFeatLyr = New FeatureLayer
'                    Set pFeatLyr.FeatureClass = pFeatCls
'                    pFeatLyr.name = pFeatCls.AliasName
'                    Set pchkLayer = GetInputFeatureLayer(pFeatLyr.name)
'                    If Not pchkLayer Is Nothing Then gMap.DeleteLayer pchkLayer
'                    pFeatLyr.Visible = False
'                    gMap.AddLayer pFeatLyr
                    gFeatClassDictionary.add pFeatCls.AliasName, "FeatureClass"
                    Set pDataset2 = pEnumDataset2.Next
                Loop
            ElseIf TypeOf pDataset Is IFeatureClass Then
                    Set pFeatCls = pDataset
'                    Set pFeatLyr = New FeatureLayer
'                    Set pFeatLyr.FeatureClass = pFeatCls
'                    pFeatLyr.name = pFeatCls.AliasName
'                    pFeatLyr.Visible = False
'                    Set pchkLayer = GetInputFeatureLayer(pFeatLyr.name)
'                    If Not pchkLayer Is Nothing Then gMap.DeleteLayer pchkLayer
'                    gMap.AddLayer pFeatLyr
                    gFeatClassDictionary.add pFeatCls.AliasName, "FeatureClass"
            ElseIf TypeOf pDataset Is iTable Then
'                    Set pStTab = New StandaloneTable
'                    Set pStTabColl = gMap
'                    Set pTable = GetInputDataTable(pDataset.name)
'                    If pTable Is Nothing Then
'                        Set pTable = pDataset
'                        Set pStTab.Table = pTable
'                        pStTabColl.AddStandaloneTable pStTab
'                    End If
                    gFeatClassDictionary.add pDataset.name, "Table"
            End If
            Set pDataset = pEnumDataset.Next
        Loop
        
    CreateList_FromGDB = True

Exit Function
ErrorHandler:
  HandleError True, "CreateList_FromGDB " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND


End Function

Public Sub Load_ShapeFile_to_PGDB(GDBPath As String)

    On Error GoTo ErrorHandler
    Dim pDlg As IGxDialog
      Dim pGXSelect As IEnumGxObject
      Dim pGxObject As IGxObject
      Dim pGXDataset As IGxDataset
      Dim pObjectFilter As IGxObjectFilter
      Dim i As Long
      Dim pActiveView As IActiveView
    
      Set pActiveView = gMap
      ' set up filters on the files that will be browsed
      Set pDlg = New GxDialog
      Set pObjectFilter = New GxFilterTablesAndFeatureClasses
    
      pDlg.AllowMultiSelect = False
      pDlg.Title = "Select Data"
      Set pDlg.ObjectFilter = pObjectFilter
    
      If (pDlg.DoModalOpen(pActiveView.ScreenDisplay.hWnd, pGXSelect) = False) Then Exit Sub
      If (MsgBox("The selected data will be added to the database. Are you sure you want to add the selected data?", vbQuestion + vbYesNo) = vbYes) Then
      
        ' got a valid selection from the GX Dialog, now extract the feature classes datasets etc.
        ' loop through the selection enumeration
        pGXSelect.Reset
        Set pGxObject = pGXSelect.Next
        Do While Not pGxObject Is Nothing
            If (Not pGxObject Is Nothing) Then
                ' We could be handed objects of various types, work out what types we have been handed and then open
                ' them up and add a feature layer to handle them
                Set pGXDataset = pGxObject
                If (TypeOf pGxObject Is IGxDataset) Then
                  If (pGXDataset.Type = esriDTFeatureClass) Or (pGXDataset.Type = esriDTTable) Then
                        Dim fso As New FileSystemObject
                        SUSTAIN.frmSplash.Show vbModeless
                        SUSTAIN.frmSplash.Refresh
                        If (pGXDataset.Type = esriDTFeatureClass) Then
                            'Convert the Shape fle to Geodatabase featureClass
                            Call Import_Shape_To_GDB(fso.GetParentFolderName(pGxObject.FullName), pGxObject.name, GDBPath, esriDTFeatureClass)
                        Else
                            Call Import_Shape_To_GDB(fso.GetParentFolderName(pGxObject.FullName), pGxObject.name, GDBPath, esriDTTable)
                        End If
                        Unload SUSTAIN.frmSplash
                 End If
                End If
             End If
         Set pGxObject = pGXSelect.Next
         Loop
         
        End If

Exit Sub
ErrorHandler:
  HandleError True, "Load_ShapeFile_to_PGDB " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

' *********************************
' Shape file Conversion...............................
' *********************************

Public Sub Import_Shape_To_GDB(ByVal Shp_Path As String, ByVal Shp_name As String, ByVal GDB_Path As String, ByVal pMode As esriDatasetType)

    On Error GoTo ErrorHandler
        Dim sSourceName As String
        Dim sTargetFCName As String
        Dim pInPropertySet As IPropertySet
        Dim pOutPropertySet As IPropertySet

        Set pInPropertySet = New PropertySet
        pInPropertySet.SetProperty "DATABASE", Shp_Path
        sSourceName = Shp_Path

        ' Set up output property set.
        Dim pWorkspace As IWorkspace
        Dim pWorkspaceFactory As IWorkspaceFactory
        'Set pWorkspaceFactory = New AccessWorkspaceFactory
        Set pWorkspaceFactory = New FileGDBWorkspaceFactory
        Set pWorkspace = pWorkspaceFactory.OpenFromFile(gGDBpath, 0)
        Set pOutPropertySet = pWorkspace.ConnectionProperties

        sTargetFCName = Replace(Shp_name, ".shp", "")
        sTargetFCName = Replace(sTargetFCName, ".dbf", "")

        ' ***********************************************************
        ' The first routine creates a featureClass and Loads the data...(No Append)
        ' The second routine Creates/Append a featureClass
        ' ***********************************************************
        
        FCLoader Shp_name, pInPropertySet, pOutPropertySet, sTargetFCName, pMode
        'ConvertShp_GDB sSourceName, sTargetFCName, sTargetFCName, pWorkspace

Exit Sub
ErrorHandler:
  HandleError True, "Import_Shape_To_GDB " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND


End Sub
    
Private Sub FCLoader(ByVal sInName As String, _
                        ByVal pInPropertySet As IPropertySet, _
                        ByVal pOutPropertySet As IPropertySet, _
                        ByVal sOutName As String, ByVal pMode As esriDatasetType)
                        
        On Error GoTo ErrorHandler
        
        Dim pErrInfo As IInvalidObjectInfo
        Dim pEnumErrors As IEnumInvalidObject
        Dim pOutFC As IFeatureClass
        Dim pOutFCFields As esriGeoDatabase.IFields
        Dim pTable As iTable
        
        Dim pWorkspaceFactory As IWorkspaceFactory
        Set pWorkspaceFactory = New FileGDBWorkspaceFactory
        
         ' check if the FeatureClass already exists.....
        Set pOutFC = GetFeatureClass(pOutPropertySet.GetProperty("DATABASE"), sOutName)
        If Not pOutFC Is Nothing Then
            MsgBox "The Featureclass already exists. Please delete and then try.", vbCritical
            Exit Sub
        End If
        Set pTable = GetTable(pOutPropertySet.GetProperty("DATABASE"), sOutName)
        If Not pTable Is Nothing Then
            Dim pWorkspace As IWorkspace
            'Set pWorkspace = ModuleUtility.OpenAccessWorkspace(gGDBpath)
            Set pWorkspace = pWorkspaceFactory.OpenFromFile(gGDBpath, 0)
            DeleteGDBData pWorkspace, sOutName, esriDTTable
        End If
                
        ' Setup output workspace.
        Dim pOutWorkspaceName As IWorkspaceName
        Set pOutWorkspaceName = New WorkspaceName
        pOutWorkspaceName.ConnectionProperties = pOutPropertySet
        'pOutWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesGDB.AccessWorkspaceFactory.1"
        pOutWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesGDB.FileGDBWorkspaceFactory"
        
        ' Set up for open.
        Dim pInWorkspaceName As IWorkspaceName
        Set pInWorkspaceName = New WorkspaceName
        pInWorkspaceName.ConnectionProperties = pInPropertySet
        pInWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
        
        If pMode = esriDTFeatureClass Then
            ' Set in dataset and table names.
            Dim pInFCName As IFeatureClassName
            Set pInFCName = New FeatureClassName
            Dim pInDatasetName As IDatasetName
            Set pInDatasetName = pInFCName
            pInDatasetName.name = sInName
            Set pInDatasetName.WorkspaceName = pInWorkspaceName
    
            ' Set out dataset and table names.
            Dim pOutDatasetName As IDatasetName
            Dim pOutFCName As IFeatureClassName
            Set pOutFCName = New FeatureClassName
            Set pOutDatasetName = pOutFCName
            Set pOutDatasetName.WorkspaceName = pOutWorkspaceName
            pOutDatasetName.name = sOutName
    
            ' Open input Featureclass to get field definitions.
            Dim pName As IName
            Dim pInFC As IFeatureClass
            Set pName = pInFCName
            Set pInFC = pName.Open
            Set pOutFCFields = pInFC.Fields
            
            ' Validate the field names.
            Dim i As Long
            ' +++ Loop through the output fields to find the geometry field
            Dim pGeoField As esriGeoDatabase.IField
            For i = 0 To pOutFCFields.FieldCount - 1
                If pOutFCFields.Field(i).Type = esriFieldType.esriFieldTypeGeometry Then
                    Set pGeoField = pOutFCFields.Field(i)
                    Exit For
                End If
            Next i
    
            ' +++ Get the geometry field's geometry defenition
            Dim pOutFCGeoDef As IGeometryDef
            Set pOutFCGeoDef = pGeoField.GeometryDef
    
            ' +++ Give the geometry definition a spatial index grid count and grid size
            Dim pOutFCGeoDefEdit As IGeometryDefEdit
            Set pOutFCGeoDefEdit = pOutFCGeoDef
            pOutFCGeoDefEdit.GridCount = 1
            pOutFCGeoDefEdit.GridSize(0) = DefaultIndexGrid(pInFC)
            Set pOutFCGeoDefEdit.SpatialReference = pGeoField.GeometryDef.SpatialReference
        Else
            '###############################################################
            ' Set in dataset and table names.
            Dim pInTableName As ITableName
            Set pInTableName = New tablename
            Set pInDatasetName = pInTableName
            pInDatasetName.name = sInName
            Set pInDatasetName.WorkspaceName = pInWorkspaceName
            
            ' Set out dataset and table names.
            Dim pOutTableName As ITableName
            Set pOutTableName = New tablename
            Set pOutDatasetName = pOutTableName
            Set pOutDatasetName.WorkspaceName = pOutWorkspaceName
            pOutDatasetName.name = sOutName
            
            ' Open input table to get field definitions.
            Set pName = pInTableName
            Set pTable = pName.Open
            Set pOutFCFields = pTable.Fields
        End If
               

        ' Load the table.
        Dim pFCToFC As IFeatureDataConverter
        Set pFCToFC = New FeatureDataConverter
        If pMode = esriDTFeatureClass Then
            Set pEnumErrors = pFCToFC.ConvertFeatureClass(pInFCName, Nothing, Nothing, pOutFCName, pOutFCGeoDef, pOutFCFields, "", 1000, 0)
        Else
            Set pEnumErrors = pFCToFC.ConvertTable(pInTableName, Nothing, pOutTableName, pOutFCFields, "", 1000, 0)
        End If

        ' If some of the records do not load, report to report window.
        Set pErrInfo = pEnumErrors.Next
        Set pEnumErrors = pEnumErrors
        If Not pErrInfo Is Nothing Then
            MsgBox (pErrInfo.InvalidObjectID & vbTab & pErrInfo.ErrorDescription)
            Do
                Set pErrInfo = pEnumErrors.Next
                If pErrInfo Is Nothing Then Exit Do
                MsgBox (pErrInfo.InvalidObjectID & vbTab & pErrInfo.ErrorDescription)
            Loop
            pEnumErrors.Reset
            MsgBox (sInName & " data load completed with errors")
        Else
            'MsgBox (sInName & " data load completed")
        End If
        
        If pMode = esriDTFeatureClass Then
            Set pName = pOutFCName
            Set pOutFC = pName.Open
            Dim pFeatLyr As IFeatureLayer
            Set pFeatLyr = New FeatureLayer
            Set pFeatLyr.FeatureClass = pOutFC
            pFeatLyr.name = pOutFC.AliasName
            pFeatLyr.Visible = False
            gFeatClassDictionary.add pFeatLyr.name, "FeatureClass" ' Add to the Dictionary....
            ' Prompt the user to add to the Map...
            If (MsgBox("Do you want to add the featureclass to the current Map?", vbQuestion + vbYesNo) = vbYes) Then
                ' Open input Featureclass to get field definitions.
                gMap.AddLayer pFeatLyr
            End If
        Else
            Set pTable = GetInputDataTable(sOutName)
            If pTable Is Nothing Then
                Set pName = pOutTableName
                Set pTable = pName.Open
                If (MsgBox("Do you want to add the table to the current Map?", vbQuestion + vbYesNo) = vbYes) Then
                    AddTableToMap pTable
                End If
            End If
            
            If (gFeatClassDictionary Is Nothing) Then
                Set gFeatClassDictionary = CreateObject("Scripting.Dictionary")
            End If
            gFeatClassDictionary.add sOutName, "Table"    ' Add to the Dictionary....
        End If
        


        Exit Sub

ErrorHandler:
  HandleError True, "FCLoader " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub
Public Function GetFeatureClass(ByVal GDBPath As String, ByVal strFeatureClass As String) As IFeatureClass
        On Error GoTo ErrorHandler

        Set GetFeatureClass = Nothing
        Dim pWorkspaceFactory As IWorkspaceFactory
        'Set pWorkspaceFactory = New AccessWorkspaceFactory
        Set pWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace
        Set pWorkspace = pWorkspaceFactory.OpenFromFile(gGDBpath, 0)
        Dim pFeatureClassContainer As IFeatureClassContainer
        Dim pDataset As IDataset, pDataset2 As IDataset, pEnumDataset As IEnumDataset, pEnumDataset2 As IEnumDataset
        
        Set pEnumDataset = pWorkspace.Datasets(esriDTAny)
        Set pDataset = pEnumDataset.Next
        Do Until pDataset Is Nothing
            If TypeOf pDataset Is IFeatureDataset Then
                Set pFeatureClassContainer = pDataset
                Set pEnumDataset2 = pFeatureClassContainer.Classes
                Set pDataset2 = pEnumDataset2.Next
                Do Until pDataset2 Is Nothing
                    If UCase(pDataset2.name) = UCase(strFeatureClass) Then
                        Set GetFeatureClass = pDataset2
                        Exit Function
                    End If
                    Set pDataset2 = pEnumDataset2.Next
                Loop
            ElseIf TypeOf pDataset Is IFeatureClass Then
                    If UCase(pDataset.name) = UCase(strFeatureClass) Then
                        Set GetFeatureClass = pDataset
                        Exit Function
                    End If
            End If
            Set pDataset = pEnumDataset.Next
        Loop
        
        Set pWorkspace = Nothing
        Set pWorkspaceFactory = Nothing
        Set pDataset = Nothing
        Set pDataset2 = Nothing

        Exit Function

ErrorHandler:

         HandleError True, "GetFeatureClass  " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

    End Function
    
Public Function GetTable(ByVal GDBPath As String, ByVal strFeatureClass As String) As iTable
        On Error GoTo ErrorHandler

        Set GetTable = Nothing
        Dim pWorkspaceFactory As IWorkspaceFactory
        'Set pWorkspaceFactory = New AccessWorkspaceFactory
        Set pWorkspaceFactory = New FileGDBWorkspaceFactory
        Dim pWorkspace As IWorkspace
        Set pWorkspace = pWorkspaceFactory.OpenFromFile(GDBPath, 0)
        Dim pFeatureClassContainer As IFeatureClassContainer
        Dim pDataset As IDataset, pDataset2 As IDataset, pEnumDataset As IEnumDataset, pEnumDataset2 As IEnumDataset
        
        Set pEnumDataset = pWorkspace.Datasets(esriDTTable)
        Set pDataset = pEnumDataset.Next
        Do Until pDataset Is Nothing
            If UCase(pDataset.name) = UCase(strFeatureClass) Then
                Set GetTable = pDataset
                Exit Function
            End If
            Set pDataset = pEnumDataset.Next
        Loop
        
        Set pWorkspace = Nothing
        Set pWorkspaceFactory = Nothing
        Set pDataset = Nothing
        Set pDataset2 = Nothing

        Exit Function

ErrorHandler:

         HandleError True, "GetTable  " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

    End Function

Public Function DeleteGDBData(ByVal pWorkspace As IWorkspace, ByVal sDatasetName As String, ByVal eESRIObjectType As esriDatasetType) As Boolean

        On Error GoTo ErrorHandler
        DeleteGDBData = False
        ' Qualify the name
        sDatasetName = gdbGetQualifiedName(pWorkspace, sDatasetName)

        ' Set up Workspace name
        Dim pWorkspaceName As IWorkspaceName
        Set pWorkspaceName = New WorkspaceName

        pWorkspaceName.ConnectionProperties = pWorkspace.ConnectionProperties

        ' Set WorkspaceFactoryProgID based on the type of the workspace
        If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
            pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesGDB.SdeWorkspaceFactory.1"
        ElseIf pWorkspace.Type = esriWorkspaceType.esriLocalDatabaseWorkspace Then
            'pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesGDB.AccessWorkspaceFactory.1"
            pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesGDB.FileGDBWorkspaceFactory"
        ElseIf pWorkspace.Type = esriWorkspaceType.esriFileSystemWorkspace Then
            pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
        Else
            MsgBox ("Workspace type not supported.")
            Exit Function
        End If

        Dim pDatasetName As IDatasetName
        Select Case eESRIObjectType
            Case esriDatasetType.esriDTFeatureDataset
                Dim pInFDName As IFeatureDatasetName
                Set pInFDName = New FeatureDatasetName
                Set pDatasetName = pInFDName
            Case esriDatasetType.esriDTFeatureClass
                Dim pInFCName As IFeatureClassName
                Set pInFCName = New FeatureClassName
                Set pDatasetName = pInFCName
            Case esriDatasetType.esriDTTable
                Dim pInTableName As ITableName
                Set pInTableName = New tablename
                Set pDatasetName = pInTableName
            Case esriDatasetType.esriDTGeometricNetwork
                Dim pInNetworkName As IGeometricNetworkName
                Set pInNetworkName = New GeometricNetworkName
                Set pDatasetName = pInNetworkName
            Case esriDatasetType.esriDTRelationshipClass
                Dim pInRelationshipClassName As IRelationshipClassName
                Set pInRelationshipClassName = New RelationshipClassName
                Set pDatasetName = pInRelationshipClassName
            Case Else
                MsgBox ("Dataset Type not supported.")
                Exit Function
        End Select

        ' Set the name of the object to be deleted
        Set pDatasetName.WorkspaceName = pWorkspaceName
        pDatasetName.name = sDatasetName

        Dim pFeatureWorkspaceManage As IFeatureWorkspaceManage
        Set pFeatureWorkspaceManage = pWorkspace


        pFeatureWorkspaceManage.DeleteByName pDatasetName
        DeleteGDBData = True

        Exit Function

ErrorHandler:
  HandleError True, "DeleteGDBData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND
  DeleteGDBData = False

    End Function
    
    Private Function DefaultIndexGrid(ByVal InFC As IFeatureClass) As Double
        ' Calculate approximate first grid
        ' based on the average of a random sample of feature extents times five
        Dim lngNumFeat As Long
        Dim lngSampleSize As Long
        Dim pFields As esriGeoDatabase.IFields
        Dim pField As esriGeoDatabase.IField
        Dim strFIDName As String
        Dim strWhereClause As String
        Dim lngCurrFID As Long
        Dim pFeat As IFeature
        Dim pFeatCursor As IFeatureCursor
        Dim pFeatEnv As IEnvelope
        Dim pQueryFilter As IQueryFilter
        Dim pNewCol As New Collection
        Dim lngKMax As Long

        Dim dblMaxDelta As Double
        dblMaxDelta = 0
        Dim dblMinDelta As Double
        dblMinDelta = 1000000000000#
        Dim dblSquareness As Double
        dblSquareness = 1

        Dim i As Long
        Dim j As Long
        Dim K As Long


        Const SampleSize = 1
        Const Factor = 1

        ' Create a recordset

        Dim ColInfo(0), c0(3)

        c0(0) = "minext"
        c0(1) = CInt(5)
        c0(2) = CInt(-1)
        c0(3) = False

        ColInfo(0) = c0

        lngNumFeat = InFC.FeatureCount(Nothing) - 1
        If lngNumFeat <= 0 Then
            DefaultIndexGrid = 1000
            Exit Function
        End If

        'if the feature type is points use the density function
        If InFC.ShapeType = esriGeometryType.esriGeometryMultipoint Or InFC.ShapeType = esriGeometryType.esriGeometryPoint Then
            DefaultIndexGrid = DefaultIndexGridPoint(InFC)
            Exit Function
        End If

        ' Get the sample size
        lngSampleSize = lngNumFeat * SampleSize
        ' Don't allow too large a sample size to speed
        If lngSampleSize > 1000 Then lngSampleSize = 1000

        ' Get the ObjectID Fieldname of the feature class
        Set pFields = InFC.Fields
        ' FID is always the first field
        Set pField = pFields.Field(0)
        strFIDName = pField.name

        ' Add every nth feature to the collection of FIDs
        For i = 1 To lngNumFeat Step CLng(lngNumFeat / lngSampleSize)
            pNewCol.add (i)
        Next i

        For j = 0 To pNewCol.Count - 1 Step 250
            ' Will we top out the features before the next 250 chunk?
            lngKMax = Min(pNewCol.Count - j, 250)
            strWhereClause = strFIDName + " IN("
            For K = 1 To lngKMax
                strWhereClause = strWhereClause + CStr(pNewCol.Item(j + K)) + ","
            Next K
            ' Remove last comma and add close parenthesis
            strWhereClause = Mid(strWhereClause, 1, Len(strWhereClause) - 1) + ")"
            Set pQueryFilter = New QueryFilter
            pQueryFilter.WhereClause = strWhereClause
            Set pFeatCursor = InFC.Search(pQueryFilter, True)
            Set pFeat = pFeatCursor.NextFeature
            Do While Not pFeat Is Nothing
                ' Get the extent of the current feature
                Set pFeatEnv = pFeat.Extent
                ' Find the min, max side of all extents. The "Squareness", a measure
                ' of how close the extent is to a square, is accumulated for later
                ' average calculation.
                dblMaxDelta = Max(dblMaxDelta, Max(pFeatEnv.Width, pFeatEnv.Height))
                dblMinDelta = Min(dblMinDelta, Min(pFeatEnv.Width, pFeatEnv.Height))
                '  lstSort.AddItem Max(pFeatEnv.Width, pFeatEnv.Height)
                If dblMinDelta <> 0 Then
                    dblSquareness = dblSquareness + ((Min(pFeatEnv.Width, pFeatEnv.Height) / (Max(pFeatEnv.Width, pFeatEnv.Height))))
                Else
                    dblSquareness = dblSquareness + 0.0001
                End If
                Set pFeat = pFeatCursor.NextFeature
            Loop
        Next j



        ' If the average envelope approximates a square set the grid size half
        ' way between the min and max sides. If the envelope is more rectangular,
        ' then set the grid size to half of the max.
        If ((dblSquareness / lngSampleSize) > 0.5) Then
            DefaultIndexGrid = (dblMinDelta + ((dblMaxDelta - dblMinDelta) / 2)) * Factor
        Else
            DefaultIndexGrid = (dblMaxDelta / 2) * Factor
        End If
    End Function
    
    Private Function DefaultIndexGridPoint(ByVal InFC As IFeatureClass) As Double

        ' Calculates the Index grid based on input feature class

        ' Get the dataset
        Dim pGeoDataset As IGeoDataset
        Set pGeoDataset = InFC

        ' Get the envelope of the input dataset
        Dim pEnvelope As IEnvelope
        Set pEnvelope = pGeoDataset.Extent

        'Calculate approximate first grid
        Dim lngNumFeat As Long
        Dim dblArea As Double
        lngNumFeat = InFC.FeatureCount(Nothing)

        If lngNumFeat = 0 Or pEnvelope.IsEmpty Then
            ' when there are no features or an empty bnd - return 1000
            DefaultIndexGridPoint = 1000
        Else
            dblArea = pEnvelope.Height * pEnvelope.Width
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            ' approximate grid size is the square root of area over the number of features
            'DefaultIndexGridPoint = Math.Sqrt(dblArea / lngNumFeat)
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            ' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        End If

        Set pGeoDataset = Nothing
        Set pEnvelope = Nothing

    End Function
    
    Private Function Min(ByVal v1 As Variant, ByVal v2 As Variant) As Variant
        Min = IIf(v1 < v2, v1, v2)
    End Function
    Private Function Max(ByVal v1 As Variant, ByVal v2 As Variant) As Variant
        Max = IIf(v1 > v2, v1, v2)
    End Function
    
    Private Function gdbGetQualifiedName(ByVal pWorkspace As IWorkspace, ByVal sName As String) As String

    On Error GoTo ErrorHandler
        ' *** Returns the input name qualified as required by the workspace type.

        ' If a string containing a "." is passed to this function, assume it
        ' is already qualified and return the string unchanged.

        ' Qualify only remote (SDE) Geodatabases
        If pWorkspace.Type = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
            If InStr(sName, ".") = 0 Then
                Dim pDatabaseConnectionInfo As IDatabaseConnectionInfo
                Set pDatabaseConnectionInfo = pWorkspace

                Dim pSQLSyntax As ISQLSyntax
                Set pSQLSyntax = pWorkspace

                gdbGetQualifiedName = pSQLSyntax.QualifyTableName(pDatabaseConnectionInfo.ConnectedDatabase, pDatabaseConnectionInfo.ConnectedUser, sName)
            Else
                gdbGetQualifiedName = sName
            End If
        Else
            ' Strip off any existing qualification as it is not needed for Access
            If InStr(sName, ".") > 0 Then
                sName = Right(sName, (Len(sName) - InStrRev(sName, ".")))
            End If
            gdbGetQualifiedName = sName
        End If
        
    Exit Function
ErrorHandler:
  HandleError True, "gdbGetQualifiedName " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND


    End Function
    
'******************************************************************************
'Subroutine: GetLayerFromMap
'Author:
'Purpose:    General Function to get a layer(feature/raster/etc) from map.
'******************************************************************************
Public Function GetLayerFromMap(pLayerName As String) As ILayer
On Error GoTo ShowError
    'If the map has subwatershed layer, remove it
    Dim i As Integer
    Dim pLayer As ILayer
    For i = 0 To (gMap.LayerCount - 1)
        Set pLayer = gMap.Layer(i)
        If (UCase(pLayer.name) = UCase(pLayerName)) Then
            Set GetLayerFromMap = pLayer
            Exit For
        End If
    Next
    GoTo CleanUp
ShowError:
    MsgBox "GetLayerFromMap: " & Err.description
CleanUp:
    Set pLayer = Nothing
End Function
    
Public Sub ReorderLayers()
        
    On Error GoTo ErrorHandler
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim iCnt As Integer
    Dim iRas As Integer
    iRas = gMap.LayerCount
    Dim oColLay As Collection
    Set oColLay = New Collection
 
    For iCnt = 0 To gMap.LayerCount - 1
      Set pLayer = gMap.Layer(iCnt)
      oColLay.add pLayer
    Next iCnt
    
    For iCnt = 1 To oColLay.Count
      Set pLayer = oColLay.Item(iCnt)
      If TypeOf pLayer Is IRasterLayer Then
        gMap.MoveLayer pLayer, gMap.LayerCount
        iRas = iRas - 1
      End If
      If TypeOf pLayer Is IFeatureLayer Then
        Set pFlayer = pLayer
        If pFlayer.FeatureClass.ShapeType = esriGeometryPolygon Then
            gMap.MoveLayer pLayer, iRas - 1
        End If
      End If
    Next iCnt
        
    ' Zoom to extents and save the Image.....
    Dim pActView As IActiveView
    Set pActView = gMap
    pActView.Extent = pActView.FullExtent
    'Refresh the active view
    pActView.Refresh


Exit Sub
ErrorHandler:
  HandleError True, "ReorderLayers " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 1, m_ParentHWND

End Sub

Public Function SelectByLocationIN(ByVal pMap As IMap, ByVal pFLayer_poly As IFeatureLayer, ByVal pGeometry As IGeometry, ByVal pShapeFieldName As String) As IFeatureCursor
    
    On Error GoTo ErrorHandler:
    
    Dim pAView As IActiveView
    Dim pFSelection_poly As IFeatureSelection
    Dim pSelectionSet As ISelectionSet
    Dim pFCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pSpatialFilter As ISpatialFilter
    

    Set pAView = pMap
    Set pSpatialFilter = New SpatialFilter
    Set pSpatialFilter.Geometry = pGeometry
    pSpatialFilter.GeometryField = pShapeFieldName
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    Set pFSelection_poly = pFLayer_poly
    pFSelection_poly.SelectFeatures pSpatialFilter, esriSelectionResultNew, True
    
    Set pSelectionSet = pFSelection_poly.SelectionSet
    pSelectionSet.Search Nothing, False, pFCursor
    Set SelectByLocationIN = pFCursor
    pFSelection_poly.Clear

    Exit Function

ErrorHandler:
  MsgBox Err.Number & Err.description & "In SelectbyLocationCompIN"
    
    End Function

'function to get find an item in the Listbox
Public Function GetListBoxIndex(objX As Object, sStr As String, Optional sExact As Boolean = False, Optional sStart As Long = -1) As Long

   If TypeOf objX Is ListBox Then
      If sExact = True Then
         GetListBoxIndex = SendMessage(objX.hWnd, LB_FINDSTRINGEXACT, sStart, ByVal sStr)
      Else
         GetListBoxIndex = SendMessage(objX.hWnd, LB_FINDSTRING, sStart, ByVal sStr)
      End If
   Else
      If sExact = True Then
         GetListBoxIndex = SendMessage(objX.hWnd, CB_FINDSTRINGEXACT, sStart, ByVal sStr)
      Else
         GetListBoxIndex = SendMessage(objX.hWnd, CB_FINDSTRING, sStart, ByVal sStr)
      End If
   End If

    
End Function

Public Function Get_BMP_MappingName(ByVal strBMP As String) As String

    Select Case strBMP
        Case "DryPond"
            Get_BMP_MappingName = "Dry_Pond"
        Case "WetPond"
            Get_BMP_MappingName = "Wet_Pond"
        Case "Infiltrationbasin"
            Get_BMP_MappingName = "Infiltration_basin"
        Case "InfiltrationTrench"
            Get_BMP_MappingName = "Infiltration_trench"
        Case "BioRetentionBasin"
            Get_BMP_MappingName = "Bioretention"
        Case "Sand filter (surface)"
            Get_BMP_MappingName = "Sand_filter_(surface)"
        Case "Sand filter (non-surface)"
            Get_BMP_MappingName = "Sand_filter_(non-surface)"
        Case "Constructed wetland"
            Get_BMP_MappingName = "Constructed_wetland"
        Case "PorousPavement"
            Get_BMP_MappingName = "Porous_Pavement"
        Case "VegetativeSwale"
            Get_BMP_MappingName = "Grassed_swales"
        Case "Buffer Strip"
            Get_BMP_MappingName = "Vegetated_filterstrip"
        Case "RainBarrel"
            Get_BMP_MappingName = "Rain_barrel"
        Case "Cistern"
            Get_BMP_MappingName = "Cistern"
        Case "GreenRoof"
            Get_BMP_MappingName = "Green_roof"
    End Select
            
    
End Function

Public Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
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
        OnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOOWNERZORDER)
    Else
        OnTop = SetWindowPos(FrmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOOWNERZORDER)
    End If
End Sub

Public Function Remove_Numbers(strReplace)

 Dim oReg As RegExp
 Set oReg = New RegExp
 
 oReg.IgnoreCase = True
 oReg.pattern = "[0-9]"
 oReg.Global = True

 Remove_Numbers = oReg.Replace(strReplace, "")
 
End Function
