Attribute VB_Name = "ModuleBMPData"
'******************************************************************************
'   Application: SUSTAIN - Best Management Practice Decision Support System
'   Company:     Tetra Tech, Inc
'   File Name:   ModuleBMPTypes
'   Purpose:     Functions and subroutine to
'                and specific BMP sites
'
'   Designer:    Leslie Shoemaker, Ting Dai, Khalid Alvi, Jenny Zhen, John Riverson, Sabu Paul
'   Developer:   Haihong Yang, Sabu Paul, Mira Chokshi
'   History:     Created:  08/../2004 - Sabu Paul
'                Modified: 08/19/2004 - Sabu Paul added comments to project
'
'******************************************************************************

Option Explicit
Option Base 0

Public Enum COST_VOLUME_TYPE
  COST_VOLUME_TYPE_TOTAL = 1
  COST_VOLUME_TYPE_MEDIA = 2
  COST_VOLUME_TYPE_UNDERDRAIN = 3
End Enum
Public gBMPTypeTag As String 'To differentiate BMP Template and on Map
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "C:\projects\sustain\SUSTAIN_9_3\ModuleBMPData.bas"


'*******************************************************************************
'Subroutine : CallInitRoutines
'Purpose    : Based on the type of the BMP, calls different subroutines to initialize
'             the appropriate forms
'Note       :
'Arguments  : BMP Type
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub CallInitRoutines(pBMPType As String, pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    
'    'show BMPTypeDef dialog box
'    If gBMPEditMode = True Then
'        SetupFrmBmpTypeDef pBMPType, pBmpDetailDict
'    End If
    
    'Set the underdrain parameter on in the infiltration method is GA
''    If Not pBmpDetailDict Is Nothing Then
''        If pBmpDetailDict.Exists("Infiltration Method") Then
''            If pBmpDetailDict.Item("Infiltration Method") = 1 Then
''                pBmpDetailDict.Item("UnderDrainON") = True
''            End If
''        ElseIf Not gBMPOptionsDict Is Nothing Then
''            If gBMPOptionsDict.Exists("Infiltration Method") Then
''                If gBMPOptionsDict.Item("Infiltration Method") = 1 Then
''                    pBmpDetailDict.Item("UnderDrainON") = True
''                End If
''            End If
''        End If
''    End If
    
    If (pBMPType <> "Regulator") Then
        'Initialize cost tab
        FrmBMPData.InitCostFromDB pBMPType
        
        'Initialize new cost tab
        FrmBMPData.Update_Component_List pBmpDetailDict
        'FrmBMPData.Refresh
        
        'Based on the BMP type call different initialization routines
        With FrmBMPData
            .FrameExitType.Visible = True
            .FrameReleaseOption.Visible = True
            .LabelOrificeDiameter.Visible = True
            .LabelOrificeHeight.Visible = True
            .BMPOrificeDiameter.Visible = True
            .BMPOrificeHeight.Visible = True
            .imgBmpa1.Picture = .ImgLstDgm.ListImages(1).Picture
            .imgBmpa1.Left = 120
'            .imgGreenRoof.Visible = False
'            .imgPorousPave.Visible = False
            .labelLength.Caption = "Length (ft)"
            .labelWidth.Caption = "             Width (ft)"
            
            'Hide the Width optimization option
            .BWidthOptimized2.Visible = False
            .BWidthOptimized.Visible = False
            .BWidthBOptimized.Visible = False
            .BWidthBOptimized2.Visible = False
        
            'Set the underdrain parameter on in the infiltration method is GA
            If Not gBMPOptionsDict Is Nothing Then
                If gBMPOptionsDict.Exists("Infiltration Method") Then
                    If gBMPOptionsDict.Item("Infiltration Method") = 1 Then
                        .GreenAmptON.value = vbChecked
                    Else
                        .GreenAmptON.value = vbUnchecked
                    End If
                End If
            End If
            
            'Call initialize sediment properties
            .InitializeSedimentParameters pBmpDetailDict
        End With
    End If
                
    Select Case pBMPType
        Case "BioRetentionBasin":
            Call InitForBioRB(pBmpDetailDict)
        Case "WetPond":
            Call InitForWetPond(pBmpDetailDict)
        Case "Cistern":
            Call InitForCistern(pBmpDetailDict)
        Case "DryPond":
            Call InitForDryPond(pBmpDetailDict)
        Case "InfiltrationTrench":
            Call InitForInfiltTrench(pBmpDetailDict)
        Case "GreenRoof":
            Call InitForGreenRoof(pBmpDetailDict)
        Case "PorousPavement":
            Call InitForPorousPavement(pBmpDetailDict)
        Case "RainBarrel":
            Call InitForRainB(pBmpDetailDict)
        Case "VegetativeSwale":
            Call InitForVegSwale(pBmpDetailDict)
        Case "Conduit":
            Call InitForVegSwale(pBmpDetailDict)
        Case "Regulator":
            Call InitForRegulator(pBmpDetailDict)
        Case Else
            Exit Sub 'MsgBox "This BMP type is not yet implemented", vbExclamation
    End Select
            
    GoTo CleanUp
    
ShowError:
    MsgBox "CallInitRoutines :", Err.description
CleanUp:
    
End Sub

Public Function Get_BMP_Name(ByVal strBMP As String) As String
  On Error GoTo ErrorHandler


    ' Remove any numerics.....
    strBMP = Remove_Numbers(strBMP)
    
    Select Case strBMP
        Case "InfiltrationTrench"
            Get_BMP_Name = "Infiltration Trench"
        Case "VegetativeSwale"
            Get_BMP_Name = "Vegetative Swale"
        Case "WetPond"
            Get_BMP_Name = "Wet Pond"
        Case "DryPond"
            Get_BMP_Name = "Dry Pond"
        Case "BioRetentionBasin"
            Get_BMP_Name = "Bioretention"
        Case "RainBarrel"
            Get_BMP_Name = "Rain Barrel"
        Case "Cistern"
            Get_BMP_Name = "Cistern"
        Case "PorousPavement"
            Get_BMP_Name = "Porous Pavement"
        Case "GreenRoof"
            Get_BMP_Name = "Green Roof"
        Case "Conduit"
            Get_BMP_Name = "Conduit"
    End Select
    

  Exit Function
ErrorHandler:
  HandleError True, "Get_BMP_Name " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Function

Public Sub SetupFrmBmpTypeDef(pBMPType As String, pBmpDetailDict As Scripting.Dictionary, Optional bmpId As Integer, Optional strFrmTag As String)
On Error GoTo ShowError
    
    'frmBMPDef.Form_Initialize

    If (gBMPCatDict Is Nothing) Or (gBMPTypeDict Is Nothing) Then
        Call InitBmpTypeCatDict
    End If
    'If strFrmTag <> "" Then frmBMPDef.Tag = strFrmTag
    
    Dim BMPName As String
    BMPName = Get_BMP_Name(pBMPType)
    'go to the tab, disabled other tabs
    Dim ikey As Integer
    For ikey = 0 To frmBMPDef.TabBMPType.Tabs - 1
        If ikey = Get_Tab_Index(gBMPTypeDict.Item(BMPName)) Then
            frmBMPDef.TabBMPType.Tab = ikey
        Else
            frmBMPDef.TabBMPType.TabEnabled(ikey) = False
        End If
    Next

    'set bmp category and then disable it (not editable)
    For ikey = 0 To frmBMPDef.cmbBMPCategory.ListCount
        
        If frmBMPDef.cmbBMPCategory.List(ikey) = gBMPCatDict.Item(BMPName) Then
            frmBMPDef.cmbBMPCategory.ListIndex = ikey
            Exit For
        End If
    Next
    frmBMPDef.cmbBMPCategory.Enabled = False

    'set bmp type and then disable it (not editable)
    For ikey = 0 To frmBMPDef.cmbBmpType.ListCount
        
        If frmBMPDef.cmbBmpType.List(ikey) = BMPName Then
            frmBMPDef.cmbBmpType.ListIndex = ikey
            Exit For
        End If
    Next
    frmBMPDef.cmbBmpType.Enabled = False
    
    frmBMPDef.BMPNameA.Text = pBmpDetailDict.Item("BMPName")
    frmBMPDef.BMPNameA.Enabled = False
    
    'populate selections
    If pBmpDetailDict.Item("Infiltration Method") = 0 Then
        frmBMPDef.optHal.value = True
        frmBMPDef.optGreen.value = False
    Else
        frmBMPDef.optHal.value = False
        frmBMPDef.optGreen.value = True
    End If
    
    If pBmpDetailDict.Item("Pollutant Removal Method") = 0 Then
        frmBMPDef.optDecay.value = True
        frmBMPDef.optKadlac.value = False
    Else
        frmBMPDef.optDecay.value = False
        frmBMPDef.optKadlac.value = True
    End If
    
    Select Case pBmpDetailDict.Item("Pollutant Routing Method")
        Case 0:
            frmBMPDef.optPlug.value = True
        Case 1:
            frmBMPDef.optMixed.value = True
        Case Else:
            frmBMPDef.optSeries.value = True
            frmBMPDef.txtCSTR.Text = pBmpDetailDict.Item("Pollutant Routing Method")
    End Select
    GoTo CleanUp
    
ShowError:
    MsgBox "SetupFrmBmpTypeDef :", Err.description
CleanUp:
End Sub
'*******************************************************************************
'Subroutine : createNewBmpType
'Purpose    : Create a new BMP template type and calls the subroutine
'             to store the detail information in BMPDefaults table
'Note       :
'Arguments  : Type of BMP
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub CreateNewBmpType(BMPType As String, BMPName As String) ', Optional myTag As String
On Error GoTo ShowError
    Dim pBMPTypesTable As iTable
    Set pBMPTypesTable = GetInputDataTable("BMPTypes")
        
    Dim pNewID As Integer
    Dim pNewName As String
    Dim pNewClass As String
    
    pNewID = pBMPTypesTable.RowCount(Nothing) + 1
    pNewName = BMPName
    
    Select Case BMPType
        Case "BioRetentionBasin":
            pNewClass = "A"
        Case "WetPond":
            pNewClass = "A"
        Case "DryPond":
            pNewClass = "A"
        Case "RainBarrel":
            pNewClass = "A"
        Case "Cistern":
            pNewClass = "A"
        Case "InfiltrationTrench":
            pNewClass = "A"
        Case "GreenRoof":
            pNewClass = "A"
        Case "PorousPavement":
            pNewClass = "A"
        Case "VegetativeSwale":
            pNewClass = "B"
        Case Else
            pNewClass = "A"
    End Select
   
    Dim pIDindex As Long
    pIDindex = pBMPTypesTable.FindField("ID")
    Dim pNameIndex As Long
    pNameIndex = pBMPTypesTable.FindField("Name")
    Dim pTypeIndex As Long
    pTypeIndex = pBMPTypesTable.FindField("Type")
    Dim pClassIndex As Long
    pClassIndex = pBMPTypesTable.FindField("Class")
    Dim pRow As iRow
    
    gNewBMPId = pNewID
    'gNewBMPType = BMPType
    gNewBMPName = pNewName

    '***creating the dictionary to hold control names and corresponging values
    Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
    '
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = GetBMPPropDict(gNewBMPId)
    
    'Initialize the decay rates and underdrain percentage removals
    InitPollutantData gNewBMPType

    'Call initialization routine based on the BMP type
    CallInitRoutines gNewBMPType, pBmpDetailDict
    
    'If myTag <> "BMPOnMap" Then
    If gBMPTypeTag <> "BMPOnMap" Then
        If Not (gBMPDetailDict Is Nothing) Then
            'Insert another row into the BMPTypes table
            Set pRow = pBMPTypesTable.CreateRow
            'Get the name of the BMP from the Dictionary -- Sabu paul, August 30, 2004
            pNewName = gBMPDetailDict.Item("BMPName")
            pRow.value(pIDindex) = pNewID
            pRow.value(pNameIndex) = pNewName
            pRow.value(pTypeIndex) = BMPType
            pRow.value(pClassIndex) = pNewClass
            pRow.Store
            'Insert the BMP properties into BMPDefaults table
            Call InsertNewBmpTypeDetails(gBMPDetailDict)
        End If
    End If
    GoTo CleanUp
ShowError:
    MsgBox "CreateNewBmpType :", Err.description
CleanUp:
    Set pBMPTypesTable = Nothing
    Set pRow = Nothing
End Sub
'*******************************************************************************
'Subroutine : getBMPPropDict
'Purpose    : Gets the BMP details for a BMP template from BMPDefaults
'Note       :
'Arguments  : Id of BMP template
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Function GetBMPPropDict(bmpId As Integer) As Scripting.Dictionary
On Error GoTo ShowError

    Dim pBMPDefaultTable As iTable
    Set pBMPDefaultTable = GetInputDataTable("BMPDefaults")
    
    If (pBMPDefaultTable Is Nothing) Then
        Set pBMPDefaultTable = CreateBMPDefaultsDBF("BMPDefaults")
        AddTableToMap pBMPDefaultTable
    End If
    
    Dim pIDindex As Long
    pIDindex = pBMPDefaultTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDefaultTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDefaultTable.FindField("PropValue")
                  
    Dim pPropertyName As String
    Dim pPropertyValue As String
    Dim pTmpBMPName As String
    Dim pTmpBMPID As Integer
    
    'Check the existence of BioRetBasin in the table
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "ID = " & bmpId

    Dim pCursor As ICursor
    
    Set pCursor = pBMPDefaultTable.Search(pQueryFilter, False)
    
    Dim pSelRowCount As Long
    pSelRowCount = pBMPDefaultTable.RowCount(pQueryFilter)
    
    Dim pRow As iRow
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = CreateObject("Scripting.Dictionary")
    
    If pSelRowCount > 0 Then
        Do
            Set pRow = pCursor.NextRow
            If Not (pRow Is Nothing) Then
                pTmpBMPID = pRow.value(pIDindex)
                pPropertyName = pRow.value(pPropNameIndex)
                pPropertyValue = pRow.value(pPropValueIndex)
                Debug.Print pPropertyName & "," & pPropertyValue
                If pPropertyName <> "ID" Then
                    pBmpDetailDict.add pPropertyName, pPropertyValue
                End If
            End If
        Loop Until (pRow Is Nothing)
    Set GetBMPPropDict = pBmpDetailDict
    Else
        Set GetBMPPropDict = Nothing
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "GetBMPPropDict :", Err.description
    
CleanUp:
    Set pBMPDefaultTable = Nothing
    Set pQueryFilter = Nothing
    Set pBmpDetailDict = Nothing
    Set pCursor = Nothing
End Function
'*******************************************************************************
'Subroutine : GetAggBMPPropDict
'Purpose    : Gets the BMP details for a Aggregate BMP template from BMPDefaults
'Note       :
'Arguments  : Aggregate BMP Name
'Author     : Sabu Paul
'*******************************************************************************
Public Function GetAggBMPPropDict(AgBMPName As String) As Scripting.Dictionary
On Error GoTo ShowError
    Dim resDict As Scripting.Dictionary
    Dim bmpDict As Scripting.Dictionary
    
    Set resDict = New Scripting.Dictionary
    
    Dim pBMPDefaultTable As iTable
    Set pBMPDefaultTable = GetInputDataTable("BMPDefaults")
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    pQueryFilter.WhereClause = "PropName='Type' And PropValue='" & AgBMPName & "'"
    
    Dim pIDindex As Long
    pIDindex = pBMPDefaultTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDefaultTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDefaultTable.FindField("PropValue")
    
    Dim pCursor As ICursor
    Dim pRow As iRow
    Set pCursor = pBMPDefaultTable.Search(pQueryFilter, False)
    Set pRow = pCursor.NextRow

    Dim pQueryFilter2 As IQueryFilter
    Dim pCursor2 As ICursor
    Dim pRow2 As iRow
    
''    Dim pCategories
''    pCategories = Array("On-Site Interception", "On-Site Treatment", "Routing Attenuation", "Regional Storage/Treatment")
''
''    Dim catInd As Integer
''    Dim catId As Integer
''
''    Set pQueryFilter2 = New QueryFilter
''
''    Do While Not pRow Is Nothing
''        For catInd = 0 To UBound(pCategories)
''            pQueryFilter2.WhereClause = "PropName='Category' And ID = " & pRow.value(pIDindex) & " And PropValue='" & CStr(pCategories(catInd)) & "'"
''
''            Set pCursor2 = pBMPDefaultTable.Search(pQueryFilter2, False)
''            Set pRow2 = pCursor2.NextRow
''            If Not pRow2 Is Nothing Then
''                Set bmpDict = New Scripting.Dictionary
''                catId = pRow2.value(pIDindex)
''                pQueryFilter2.WhereClause = " ID = " & catId
''                Set pCursor2 = Nothing
''                Set pCursor2 = pBMPDefaultTable.Search(pQueryFilter2, False)
''                Set pRow2 = pCursor2.NextRow
''                Do Until pRow2 Is Nothing
''                    bmpDict.add pRow2.value(pPropNameIndex), pRow2.value(pPropValueIndex)
''                    Set pRow2 = pCursor2.NextRow
''                Loop
''                'bmpDict.Item("ID") = catInd + 1
''                Set resDict.Item(pCategories(catInd)) = bmpDict
''            End If
''        Next
''        Set pRow = pCursor.NextRow
''    Loop
''    Set GetAggBMPPropDict = resDict

    Set pQueryFilter2 = New QueryFilter
    
    Dim typeId As Integer
    Dim strType As String
    Do While Not pRow Is Nothing
        pQueryFilter2.WhereClause = "ID = " & pRow.value(pIDindex) & " AND PropName='BMPType'"
        Set pCursor2 = pBMPDefaultTable.Search(pQueryFilter2, False)
        Set pRow2 = pCursor2.NextRow
        If Not pRow2 Is Nothing Then
            Set bmpDict = New Scripting.Dictionary
            typeId = pRow2.value(pIDindex)
            strType = pRow2.value(pPropValueIndex)
            pQueryFilter2.WhereClause = " ID = " & typeId
            Set pCursor2 = Nothing
            Set pCursor2 = pBMPDefaultTable.Search(pQueryFilter2, False)
            Set pRow2 = pCursor2.NextRow
            Do Until pRow2 Is Nothing
                bmpDict.add pRow2.value(pPropNameIndex), pRow2.value(pPropValueIndex)
                Set pRow2 = pCursor2.NextRow
            Loop
            Set resDict.Item(strType) = bmpDict
        End If
        Set pRow = pCursor.NextRow
    Loop
    Set GetAggBMPPropDict = resDict
    GoTo CleanUp
    
ShowError:
    MsgBox "GetAggBMPPropDict :", Err.description
CleanUp:
    Set pBMPDefaultTable = Nothing
    Set pQueryFilter = Nothing
    Set pCursor = Nothing
    Set pRow = Nothing
    Set resDict = Nothing
    Set bmpDict = Nothing
    Set pQueryFilter2 = Nothing
    Set pCursor2 = Nothing
    Set pRow2 = Nothing
End Function
'*******************************************************************************
'Subroutine : initForBioRB
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for bioretention (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForBioRB(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        .SSTabBMP.TabVisible(1) = False 'All tab indices from 2 are modified -- Sabu Paul March 31, 2005
        .SSTabBMP.TabVisible(4) = False
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
    End With
   
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        
        'Set the default values for Bio Retention BMP -- Sabu Paul, Aug 25 2004
        With FrmBMPData
            .BMPWidthA.Text = 10
            .BMPLengthA.Text = 10
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 0.5
            .BMPRectWeirWidth.Text = 1#
            .WeirType1.value = True
            .BMPTriangularWeirAngle.Enabled = False
            .BMPTriangularWeirAngle.BackColor = &H80000016
            .SoilDepth.Text = 4#
            .SoilPorosity.Text = 0.4
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.5
            .StorageDepth.Text = 0#
            .VoidFraction.Text = 0#
            .BackgroundInfiltration.Text = 0#
            .Month1.Text = 0.55
            .Month2.Text = 0.6
            .Month3.Text = 0.65
            .Month4.Text = 0.85
            .Month5.Text = 0.95
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 0.95
            .Month11.Text = 0.75
            .Month12.Text = 0.6
            ' Changed the design of cost module - Sabu Paul, September 2007
'            .Aa.Text = 5.3
'            .Ab.Text = 1
'            .Da.Text = 1
'            .Db.Text = 1
'            .LdCost.Text = 10
'            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
    End If
    GoTo CleanUp
    


    
ShowError:
    MsgBox "initForBioRB :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub
'*******************************************************************************
'Subroutine : initForCistern
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for cistern (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForCistern(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class B Dimension Tab and enable Cistern Tab
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(2) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(3) = False
        .SSTabBMP.TabVisible(4) = True
        .SSTabBMP.TabVisible(0) = True
        .OptionRelNone.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .OptionRelCistern.value = True
        .WeirType1.value = True
        .BMPTriangularWeirAngle.Enabled = False
        .BMPTriangularWeirAngle.BackColor = &H80000016
'        .imgBmpa1.Visible = False
'        .imgCistDia.Visible = True
        .imgBmpa1.Picture = .ImgLstDgm.ListImages(2).Picture
        .labelLength.Caption = "Diameter (ft)"
        '.labelWidth.Caption = "No. of Units used"
        .BMPWidthA.Enabled = False
    End With
    Load FrmBMPData
    'Create a dictionary to store the property names and property values
    'Set gBMPDetailDict = getBMPPropDict("Cistern")
    
    
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        FrmBMPData.Cistern_Initialize
        With FrmBMPData
            '.BMPWidthA.Text = 10
            '.BMPUnitsA = 1
            .BMPLengthA.Text = 10
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 0.5
            .BMPRectWeirWidth.Text = 1#
            .SoilDepth.Text = 4#
            .SoilPorosity.Text = 0.4
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.5
            .StorageDepth.Text = 0#
            .VoidFraction.Text = 0#
            .BackgroundInfiltration.Text = 0#
            .Month1.Text = 0.55
            .Month2.Text = 0.6
            .Month3.Text = 0.65
            .Month4.Text = 0.85
            .Month5.Text = 0.95
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 0.95
            .Month11.Text = 0.75
            .Month12.Text = 0.6
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 5.3
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "initForCistern :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub
'*******************************************************************************
'Subroutine : initForDryPond
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for dry pond (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForDryPond(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class B Dimension Tab and Cistern Tab
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        .WeirType1.value = True
        .BMPTriangularWeirAngle.Enabled = False
        .BMPTriangularWeirAngle.BackColor = &H80000016
    End With
    Load FrmBMPData
    
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        With FrmBMPData
            .BMPWidthA.Text = 15
            .BMPLengthA.Text = 45
            .BMPOrificeHeight.Text = 0.5
            .BMPOrificeDiameter.Text = 5
            .BMPWeirHeight.Text = 5
            .BMPRectWeirWidth.Text = 2
            .SoilDepth.Text = 1
            .SoilPorosity.Text = 0.2
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.3
            .StorageDepth.Text = 0
            .VoidFraction.Text = 0
            .BackgroundInfiltration.Text = 0
            .Month1.Text = 1
            .Month2.Text = 1
            .Month3.Text = 1
            .Month4.Text = 1
            .Month5.Text = 1
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 1
            .Month11.Text = 1
            .Month12.Text = 1
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 0.65
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal

    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "initForDryPond :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
    
End Sub

'*******************************************************************************
'Subroutine : initForInfiltTrench
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for infiltration trench (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForInfiltTrench(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class B Dimension Tab and Cistern Tab
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        .OptionRelNone.value = True
        .BMPTriangularWeirAngle.Enabled = False
        .BMPTriangularWeirAngle.BackColor = &H80000016
    End With
    Load FrmBMPData

    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        With FrmBMPData
            .BMPWidthA.Text = 5
            .BMPLengthA.Text = 10
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 0.5
            .BMPRectWeirWidth.Text = 1#
            .SoilDepth.Text = 8#
            .SoilPorosity.Text = 0.5
            .VegetativeParam.Text = 1
            .SoilLayerInfiltration.Text = 5
            .StorageDepth.Text = 0.6
            .VoidFraction.Text = 0.5
            .BackgroundInfiltration.Text = 0.5
            .Month1.Text = 1
            .Month2.Text = 1
            .Month3.Text = 1
            .Month4.Text = 1
            .Month5.Text = 1
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 1
            .Month11.Text = 1
            .Month12.Text = 1
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 5#
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
        
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "initForInfiltTrench :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub


'*******************************************************************************
'Subroutine : InitForGreenRoof
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for infiltration trench (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForGreenRoof(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        '.imgBmpa1.Visible = False
        'Parameters switched off for green roof
        .FrameExitType.Visible = False
        .FrameReleaseOption.Visible = False
        .LabelOrificeDiameter.Visible = False
        .LabelOrificeHeight.Visible = False
        .BMPOrificeDiameter.Visible = False
        .BMPOrificeHeight.Visible = False
'        .imgGreenRoof.Visible = True
'        .imgBmpa1.Visible = False
'        .imgCistDia.Visible = False
'        .imgPorousPave.Visible = False
        .imgBmpa1.Picture = .ImgLstDgm.ListImages(3).Picture
        .imgBmpa1.Left = 2160
    End With
   
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        
        'Set the default values for Bio Retention BMP -- Sabu Paul, Aug 25 2004
        With FrmBMPData
            .BMPWidthA.Text = 10
            .BMPLengthA.Text = 10
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 0.5
            .BMPRectWeirWidth.Text = 1#
            .WeirType1.value = True
            .BMPTriangularWeirAngle.Enabled = False
            .BMPTriangularWeirAngle.BackColor = &H80000016
            .SoilDepth.Text = 4#
            .SoilPorosity.Text = 0.4
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.5
            .StorageDepth.Text = 0#
            .VoidFraction.Text = 0#
            .BackgroundInfiltration.Text = 0#
            .Month1.Text = 0.55
            .Month2.Text = 0.6
            .Month3.Text = 0.65
            .Month4.Text = 0.85
            .Month5.Text = 0.95
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 0.95
            .Month11.Text = 0.75
            .Month12.Text = 0.6
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 5.3
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
    End If
    GoTo CleanUp

ShowError:
    MsgBox "InitForGreenRoof :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub



'*******************************************************************************
'Subroutine : InitForPorousPavement
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for infiltration trench (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForPorousPavement(pBmpDetailDict As Scripting.Dictionary)

On Error GoTo ShowError
    With FrmBMPData
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        '.imgBmpa1.Visible = False
        'Parameters switched off for green roof
        .FrameExitType.Visible = False
        .FrameReleaseOption.Visible = False
        .LabelOrificeDiameter.Visible = False
        .LabelOrificeHeight.Visible = False
        .BMPOrificeDiameter.Visible = False
        .BMPOrificeHeight.Visible = False
'        .imgGreenRoof.Visible = False
'        .imgBmpa1.Visible = False
'        .imgCistDia.Visible = False
'        .imgPorousPave.Visible = True
        .imgBmpa1.Picture = .ImgLstDgm.ListImages(3).Picture
        .imgBmpa1.Left = 2160
    End With
   
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        
        'Set the default values for Bio Retention BMP -- Sabu Paul, Aug 25 2004
        With FrmBMPData
            .BMPWidthA.Text = 10
            .BMPLengthA.Text = 10
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 0.5
            .BMPRectWeirWidth.Text = 1#
            .WeirType1.value = True
            .BMPTriangularWeirAngle.Enabled = False
            .BMPTriangularWeirAngle.BackColor = &H80000016
            .SoilDepth.Text = 4#
            .SoilPorosity.Text = 0.4
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.5
            .StorageDepth.Text = 0#
            .VoidFraction.Text = 0#
            .BackgroundInfiltration.Text = 0#
            .Month1.Text = 0.55
            .Month2.Text = 0.6
            .Month3.Text = 0.65
            .Month4.Text = 0.85
            .Month5.Text = 0.95
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 0.95
            .Month11.Text = 0.75
            .Month12.Text = 0.6
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 5.3
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
    End If
    GoTo CleanUp

ShowError:
    MsgBox "InitForPorousPavement :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub


'*******************************************************************************
'Subroutine : initForRainB
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for rain barrel (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForRainB(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class A Dimension Tab and Cistern Tab
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(2) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(3) = False
        .SSTabBMP.TabVisible(4) = False
        .SSTabBMP.TabVisible(0) = True
        .OptionRelNone.Enabled = False
        .OptionRelCistern.Enabled = False
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        .OptionRelRainB.value = True
        .NumDays.BackColor = vbWhite
        '.imgBmpa1.Visible = False
'        .imgCistDia.Visible = True
        .imgBmpa1.Picture = .ImgLstDgm.ListImages(2).Picture
        .labelLength.Caption = "Diameter (ft)"
        '.labelWidth.Caption = "No. of Units used"
        .BMPWidthA.Enabled = False
    End With
    Load FrmBMPData
   
   
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        FrmBMPData.BMPNameA = gNewBMPName
        With FrmBMPData
            '.BMPWidthA.Text = 2
            .BMPUnitsA.Text = 2
            .BMPLengthA.Text = 1.5
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0.8
            .NumDays.Text = 1
            .BMPWeirHeight.Text = 3
            .WeirType2.value = True
            .BMPTriangularWeirAngle.Text = 100
            .BMPTriangularWeirAngle.Enabled = False
            .BMPRectWeirWidth.BackColor = &H80000016 ' Background color
            .SoilDepth.Text = 0
            .SoilPorosity.Text = 0
            .VegetativeParam.Text = 0
            .SoilLayerInfiltration.Text = 0
            .StorageDepth.Text = 0
            .VoidFraction.Text = 0
            .BackgroundInfiltration.Text = 0
            .Month1.Text = 0
            .Month2.Text = 0
            .Month3.Text = 0
            .Month4.Text = 0
            .Month5.Text = 0
            .Month6.Text = 0
            .Month7.Text = 0
            .Month8.Text = 0
            .Month9.Text = 0
            .Month10.Text = 0
            .Month11.Text = 0
            .Month12.Text = 0
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 1#
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "initForRainB :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub
'*******************************************************************************
'Subroutine : initForVegSwale
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for vegetative swale (Class B) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForVegSwale(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class A Dimension Tab and Cistern Tab
        .SSTabBMP.TabVisible(0) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
    End With
    Load FrmBMPData
    
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBMP2Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        FrmBMPData.BMPNameB = gNewBMPName
        With FrmBMPData
            .BMPWidthB.Text = 3
            .BMPLengthB.Text = 40
            .BMPMaxDepth.Text = 0.33
            .BMPSlope1.Text = 1
            .BMPSlope2.Text = 1
            .BMPSlope3.Text = 0.04
            .BMPManningsN.Text = 0.15
            .SoilDepth.Text = 1
            .SoilPorosity.Text = 0.3
            .VegetativeParam.Text = 0.6
            .SoilLayerInfiltration.Text = 0.3
            .StorageDepth.Text = 0
            .VoidFraction.Text = 0
            .BackgroundInfiltration.Text = 0
            .Month1.Text = 0.6
            .Month2.Text = 0.65
            .Month3.Text = 0.7
            .Month4.Text = 0.85
            .Month5.Text = 0.95
            .Month6.Text = 1
            .Month7.Text = 1
            .Month8.Text = 1
            .Month9.Text = 1
            .Month10.Text = 0.85
            .Month11.Text = 0.75
            .Month12.Text = 0.62
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 0.5
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal

    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "initForVegSwale :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
End Sub

'*******************************************************************************
'Subroutine : initForWetPond
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for wet pond (Class A) type BMP
'Arguments  :
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InitForWetPond(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    With FrmBMPData
        'Disable Class B Dimension Tab and Cistern Tab
        .SSTabBMP.TabVisible(1) = False
        .SSTabBMP.TabVisible(4) = False 'Sabu Paul -- March 31, 2005 Tab index modified
        .SSTabBMP.TabVisible(0) = True
        .OptionRelCistern.Enabled = False
        .OptionRelRainB.Enabled = False
        .NumDays.Enabled = False
        .NumDays.BackColor = &H80000016
        .NumPeople.Enabled = False
        .NumPeople.BackColor = &H80000016
        .BMPTriangularWeirAngle.Enabled = False
        .BMPTriangularWeirAngle.BackColor = &H80000016
    End With
    Load FrmBMPData
    

    
    If Not (pBmpDetailDict Is Nothing) Then
        Call ModuleBMPData.SetBmp1Props(gNewBMPId, pBmpDetailDict)
        FrmBMPData.Show vbModal
    Else
        FrmBMPData.BMPNameA = gNewBMPName
        Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
        With FrmBMPData
            .BMPWidthA.Text = 25
            .BMPLengthA.Text = 80
            .BMPOrificeHeight.Text = 0
            .BMPOrificeDiameter.Text = 0
            .BMPWeirHeight.Text = 5
            .BMPRectWeirWidth.Text = 3
            .SoilDepth.Text = 0
            .SoilPorosity.Text = 0
            .VegetativeParam.Text = 0
            .SoilLayerInfiltration.Text = 0
            .StorageDepth.Text = 0
            .VoidFraction.Text = 0
            .BackgroundInfiltration.Text = 0
            .Month1.Text = 0
            .Month2.Text = 0
            .Month3.Text = 0
            .Month4.Text = 0
            .Month5.Text = 0
            .Month6.Text = 0
            .Month7.Text = 0
            .Month8.Text = 0
            .Month9.Text = 0
            .Month10.Text = 0
            .Month11.Text = 0
            .Month12.Text = 0
            ' Changed the design of cost module - Sabu Paul, September 2007
''            .Aa.Text = 1#
''            .Ab.Text = 1
''            .Da.Text = 1
''            .Db.Text = 1
''            .LdCost.Text = 10
''            .ConstCost.Text = 1000
        End With
        FrmBMPData.Show vbModal
        
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "initForWetPond :", Err.description
    
CleanUp:
    Set pBmpDetailDict = Nothing
    
End Sub

'*******************************************************************************
'Subroutine : initForRegulator
'Purpose    : Calls the appropriate subroutines to set the controls on the forms
'             for Regulator type BMP
'Arguments  :
'Author     : Ying Cao
'History    : 12/23/2008
'*******************************************************************************
Public Sub InitForRegulator(pBmpDetailDict As Scripting.Dictionary)
On Error GoTo ShowError
    Dim exittype As Integer  '*** type index for ORIFICE COEFFICIENT
    Dim pWeirType As Integer
    
    With FrmRegulator:
        .BMPName.Text = pBmpDetailDict.Item("BMPName")
        .BMPLength.Text = pBmpDetailDict.Item("BMPLength")
        .BMPWidth.Text = pBmpDetailDict.Item("BMPWidth")
        exittype = CInt(pBmpDetailDict.Item("OrificeExitType"))
        .OrificeExitType.Item(exittype).value = True
        .BMPOrificeDiameter.Text = pBmpDetailDict.Item("BMPOrificeDiameter")
        .BMPOrificeHeight.Text = pBmpDetailDict.Item("BMPOrificeHeight")
        .BMPWeirHeight.Text = pBmpDetailDict.Item("BMPWeirHeight")
        pWeirType = CInt(pBmpDetailDict.Item("WeirType"))
        .WeirType.Item(pWeirType).value = True
        If (pWeirType = 1) Then
            .BMPRectWeirWidth.Text = pBmpDetailDict.Item("BMPRectWeirWidth")
        Else
            .BMPTriangularWeirAngle.Text = pBmpDetailDict.Item("BMPTriangularWeirAngle")
        End If
        
    End With

    FrmRegulator.Show vbModal
    GoTo CleanUp
    
ShowError:
    MsgBox "initForRegulator :", Err.description
    
CleanUp:
    'Set pBmpDetailDict = Nothing

End Sub


'*******************************************************************************
'Subroutine : insertNewBmpTypeDetails
'Purpose    : Store the detail information of a BMP type in BMPDefaults table
'Note       :
'Arguments  : Dictionary containing the BMP details
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub InsertNewBmpTypeDetails(BMPDetailDict As Dictionary)
On Error GoTo ShowError

    InitializeMapDocument
    '**** This inserts the dimensions of a new BMP type into BMP table
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDefaults")
    
    If (pBMPDetailTable Is Nothing) Then
        Set pBMPDetailTable = CreatePropertiesTableDBF("BMPDefaults")
        AddTableToMap pBMPDetailTable
    End If
    Dim pIDindex As Long
    pIDindex = pBMPDetailTable.FindField("ID")
    Dim pPropNameIndex As Long
    pPropNameIndex = pBMPDetailTable.FindField("PropName")
    Dim pPropValueIndex As Long
    pPropValueIndex = pBMPDetailTable.FindField("PropValue")
  
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    'Iterate over the entire dictionary
    Dim pRow As iRow

    Dim i As Integer
'    If BMPDetailDict.Exists("BMPName") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPName"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
'    If BMPDetailDict.Exists("BMPType") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPType"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
'    If BMPDetailDict.Exists("BMPClass") Then
'        Set pRow = pBMPDetailTable.CreateRow
'        pRow.value(pIDindex) = gNewBMPId
'        pPropertyName = "BMPClass"
'        pPropertyValue = BMPDetailDict.Item(pPropertyName)
'        pRow.value(pPropNameIndex) = pPropertyName
'        pRow.value(pPropValueIndex) = pPropertyValue
'        pRow.Store
'        BMPDetailDict.Remove (pPropertyName)
'    End If
    
    Dim pBMPKeys
    If Not BMPDetailDict Is Nothing Then
        pBMPKeys = BMPDetailDict.keys
            
        For i = 0 To (BMPDetailDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            pPropertyValue = BMPDetailDict.Item(pPropertyName)
            Set pRow = pBMPDetailTable.CreateRow
            pRow.value(pIDindex) = gNewBMPId
            pRow.value(pPropNameIndex) = pPropertyName
            pRow.value(pPropValueIndex) = pPropertyValue
            pRow.Store
        Next
    End If
    
    ' ***************************************
    ' Now Insert the BMP option Parameters....
    ' ***************************************
    If Not gBMPOptionsDict Is Nothing Then
        pBMPKeys = gBMPOptionsDict.keys
        For i = 0 To (gBMPOptionsDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            'Modified on Feb 23, 2009 - Sabu Paul - No need to insert the entry if it is already in the BMPDetailDict.
            'If pPropertyName <> "BMPType" Then
            If Not BMPDetailDict.Exists(pPropertyName) Then
                pPropertyValue = gBMPOptionsDict.Item(pPropertyName)
                Set pRow = pBMPDetailTable.CreateRow
                pRow.value(pIDindex) = gNewBMPId
                pRow.value(pPropNameIndex) = pPropertyName
                pRow.value(pPropValueIndex) = pPropertyValue
                pRow.Store
            End If
        Next
        ' Store the Placed BMP to dict.....
        'If gBMPOptionsDict.Exists("Category") Then gBMPPlacedDict.add gBMPOptionsDict.Item("Category"), gBMPOptionsDict.Item("Category")
        If gBMPOptionsDict.Exists("BMPType") Then gBMPPlacedDict.add gBMPOptionsDict.Item("BMPType"), gBMPOptionsDict.Item("BMPType")
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "insertNewBmpTypeDetails :", Err.description
    
CleanUp:
    Set pBMPDetailTable = Nothing
    Set pRow = Nothing
    Set pBMPKeys = Nothing
End Sub


'*******************************************************************************
'Subroutine : modifyBmpTypeDetails
'Purpose    : Modify the detail information of an existing BMP type in BMPDefaults table
'Note       :
'Arguments  : Id and name of BMP Type
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub ModifyBmpTypeDetails(pBMPID As Integer, pBMPName As String)
On Error GoTo ShowError
    
    gNewBMPId = pBMPID
    gNewBMPName = pBMPName
        
    '***creating the dictionary to hold control names and corresponging values
    Set gBMPDetailDict = CreateObject("Scripting.Dictionary")
    
    Dim pBmpDetailDict As Scripting.Dictionary
    Set pBmpDetailDict = GetBMPPropDict(gNewBMPId)
    'Set the decay rates and underdrain percentage removals
    LoadPollutantData pBmpDetailDict

    Call CallInitRoutines(gNewBMPType, pBmpDetailDict)
       
    'First delete the rows with specific ID
    Dim pBMPDetailTable As iTable
    Set pBMPDetailTable = GetInputDataTable("BMPDefaults")
    
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.WhereClause = "ID = " & pBMPID
    If Not (gBMPDetailDict Is Nothing) Then
          If Not (pBMPDetailTable Is Nothing) Then
              pBMPDetailTable.DeleteSearchedRows pQueryFilter
          Else
              MsgBox "Missing 'BMPDefaults' Table"
          End If
        
          'Then insert the new rows into the table
          Call InsertNewBmpTypeDetails(gBMPDetailDict)
    End If
    GoTo CleanUp
    
ShowError:
    MsgBox "modifyBmpTypeDetails :", Err.description
    
CleanUp:
    Set pBMPDetailTable = Nothing
    Set pQueryFilter = Nothing
End Sub
'*******************************************************************************
'Subroutine : setBmp1Props
'Purpose    : Sets the values of form controls based on previously saved table
'             for Class A BMP's
'Note       :
'Arguments  : Id of BMP, Type of BMP, Dictionary containing BMP details
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub SetBmp1Props(bmpId As Integer, valueDict As Dictionary)
On Error GoTo ShowError
    Dim pBMPName As String
    pBMPName = valueDict.Item("BMPName")
    
    Dim tmpBmpDict As Scripting.Dictionary
    Set tmpBmpDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim pBMPKeys
    pBMPKeys = valueDict.keys
    
    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    Dim pControlName As String
    Dim pControl As Control
    
    With FrmBMPData
        .BMPNameA = pBMPName
        .BMPNameA.Enabled = False
        For i = 0 To (valueDict.Count - 1)
           pPropertyName = pBMPKeys(i)
           pPropertyValue = valueDict.Item(pPropertyName)
           pControlName = pPropertyName
           
            Select Case pPropertyName
                Case "SoilDOptimized":
                    If CBool(pPropertyValue) Then
                        'Based on whether soil depth is selected for optimization
                        'make the appropriate image visible
                        .SoilDOptimized.Visible = False
                        .SoilDOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinSoilDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxSoilDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "SoilDepthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BLengthOptimized":
                    If CBool(pPropertyValue) Then
                        .BLengthOptimized.Visible = False
                        .BLengthOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinBasinLength":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxBasinLength":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BasinLengthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BWidthOptimized":
                    If CBool(pPropertyValue) Then
                        .BWidthOptimized.Visible = False
                        .BWidthOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinBasinWidth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxBasinWidth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BasinWidthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "WHeightOptimized":
                    If CBool(pPropertyValue) Then
                        .WHeightOptimized.Visible = False
                        .WHeightOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinWeirHeight":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxWeirHeight":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "WeirHeightIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                    
                Case "OrificeCoef":
                    Select Case pPropertyValue
                        Case 1:
                            pControlName = "ExitType1"
                        Case 0.61:
                            pControlName = "ExitType2"
                        Case 0.5:
                            pControlName = "ExitType4"
                        Case Else:
                            pControlName = "ExitType1"
                    End Select
                Case "WeirType":
                    If pPropertyValue = "1" Then
                        pControlName = "WeirType1"
                        .WeirType1.value = True
                        .BMPTriangularWeirAngle.Enabled = False
                        .BMPTriangularWeirAngle.BackColor = &H80000016 ' Background color
                    Else
                        pControlName = "WeirType2"
                        .WeirType2.value = True
                        .BMPRectWeirWidth.Enabled = False
                        .BMPRectWeirWidth.BackColor = &H80000016 ' Background color
            
                    End If
                Case "GrowthIndex":
                    'Initialize GrowthIndex Form
                    Dim pMonthlyGIs
                    pMonthlyGIs = Split(pPropertyValue, ";")
                    Dim pMonIndex As Long
                    For pMonIndex = 0 To 11
                        For Each pControl In .Controls
                            If pControl.name = "Month" & (pMonIndex + 1) Then
                                pControl.Text = pMonthlyGIs(pMonIndex)
                            End If
                        Next pControl
                    Next pMonIndex
                Case "CisternFlow":
                    'Initialize CisternFlow Form
                    Dim pHourlyFlows
                    pHourlyFlows = Split(pPropertyValue, ";")
                    Dim pHrIndex As Long
                    For pHrIndex = 0 To 23
                        For Each pControl In .Controls
                            If pControl.name = "txtHr" & (pHrIndex + 1) Then
                                pControl.Text = pHourlyFlows(pHrIndex)
                            End If
                        Next pControl
                    Next pHrIndex
                Case "BMPLength":
                    .BMPLengthA.Text = pPropertyValue
                Case "BMPWidth":
                    .BMPWidthA.Text = pPropertyValue
                Case "SoilPorosity"
                    .SoilPorosity.Text = pPropertyValue
                Case "SoilWiltingPoint"
                    .txtWilting.Text = pPropertyValue
                Case "SoilFieldCapacity"
                    .txtCapacity.Text = pPropertyValue
                Case "SuctionHead"
                    .txtSuction.Text = pPropertyValue
                Case "Conductivity"
                    .txtConduct.Text = pPropertyValue
                Case "InitialDeficit"
                    .txtDeficit.Text = pPropertyValue
                Case "NumUnits"
                    .BMPUnitsA.Text = pPropertyValue
                Case "DrainArea"
                    .BMPDrainAreaA.Text = pPropertyValue
                Case "NumUnitsOptimized":
                    If CBool(pPropertyValue) Then
                        .NumUnitsOptimized.Visible = False
                        .NumUnitsOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinNumUnits":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxNumUnits":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "NumUnitsIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case Else
                    If pPropertyName <> "BMPName" And pPropertyName <> "BMPID" Then
                         For Each pControl In .Controls
                            If pControl.name = pPropertyName Then
                                If ((TypeOf pControl Is TextBox) Or (TypeOf pControl Is ComboBox)) Then
                                    pControl.Text = pPropertyValue
                                ElseIf (TypeOf pControl Is OptionButton) Then
                                    pControl.value = True
                                ElseIf (TypeOf pControl Is CheckBox) Then
                                    If CBool(pPropertyValue) Then
                                        pControl.value = 1
                                    Else
                                        pControl.value = 0
                                    End If
                                End If
                            End If
                        Next pControl
                    
                    End If
            End Select

        Next i
    End With
    
'        pNumUnits = CInt(Trim(BMPUnitsA.Text))
'        pDrainArea = CDbl(Trim(BMPDrainAreaA.Text))
        
    For i = 0 To (tmpBmpDict.Count - 1)
        pPropertyName = tmpBmpDict.keys(i)
        pPropertyValue = tmpBmpDict.Item(pPropertyName)
        gBMPDetailDict.add pPropertyName, pPropertyValue
    Next i
    
    GoTo CleanUp
    
ShowError:
    MsgBox "setBmp1Props :", Err.description
    
CleanUp:
    Set pBMPKeys = Nothing
    Set pControl = Nothing
    Set tmpBmpDict = Nothing
End Sub




'*******************************************************************************
'Subroutine : setBMP2Props
'Purpose    : Sets the values of form controls based on previously saved table
'             for Class B BMP's
'Note       :
'Arguments  : Id of BMP, Type of BMP, Dictionary containing BMP details
'Author     : Sabu Paul
'History    :
'*******************************************************************************
Public Sub SetBMP2Props(bmpId As Integer, valueDict As Dictionary)
On Error GoTo ShowError

    Dim pBMPName As String
    pBMPName = valueDict.Item("BMPName")
    
    Dim tmpBmpDict As Scripting.Dictionary
    Set tmpBmpDict = CreateObject("Scripting.Dictionary")
    
    
    Dim i As Integer
    Dim pBMPKeys
    pBMPKeys = valueDict.keys
    
    Dim pControlName As String
    Dim pControl As Control

    Dim pPropertyName As String
    Dim pPropertyValue As String
    
    With FrmBMPData
        .BMPNameB = pBMPName
        .BMPNameB.Enabled = False
        For i = 0 To (valueDict.Count - 1)
            pPropertyName = pBMPKeys(i)
            pPropertyValue = valueDict.Item(pPropertyName)
            
            Select Case pPropertyName
                Case "SoilDOptimized":
                    If CBool(pPropertyValue) Then
                        .SoilDOptimized.Visible = False
                        .SoilDOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinSoilDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxSoilDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "SoilDepthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "GrowthIndex":
                    'Initialize GrowthIndex Form
                    Dim pMonthlyGIs
                    pMonthlyGIs = Split(pPropertyValue, ";")
                    Dim pMonIndex As Long
                    For pMonIndex = 0 To 11
                        For Each pControl In .Controls
                            If pControl.name = "Month" & (pMonIndex + 1) Then
                                pControl.Text = pMonthlyGIs(pMonIndex)
                            End If
                        Next pControl
                    Next pMonIndex
                Case "BMPLength":
                    .BMPLengthB.Text = pPropertyValue
                Case "BMPWidth":
                    .BMPWidthB.Text = pPropertyValue
                Case "BLengthBOptimized":
                    If CBool(pPropertyValue) Then
                        .BLengthBOptimized.Visible = False
                        .BLengthBOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinBasinBLength":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxBasinBLength":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BasinBLengthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                    
                Case "BWidthBOptimized":
                    If CBool(pPropertyValue) Then
                        .BWidthBOptimized.Visible = False
                        .BWidthBOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinBasinBWidth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxBasinBWidth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BasinBWidthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                    
                Case "BDepthBOptimized":
                    If CBool(pPropertyValue) Then
                        .BDepthBOptimized.Visible = False
                        .BDepthBOptimized2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinBasinBDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxBasinBDepth":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "BasinBDepthIncr":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                    
                Case "NumUnits"
                    .BMPUnitsB.Text = pPropertyValue
                Case "DrainArea"
                    .BMPDrainAreaB.Text = pPropertyValue
                    
                Case "NumUnitsOptimizedB":
                    If CBool(pPropertyValue) Then
                        .NumUnitsOptimizedB.Visible = False
                        .NumUnitsOptimizedB2.Visible = True
                        tmpBmpDict.add pPropertyName, pPropertyValue
                    End If
                Case "MinNumUnitsB":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "MaxNumUnitsB":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case "NumUnitsIncrB":
                    tmpBmpDict.add pPropertyName, pPropertyValue
                Case Else
                    If pPropertyName <> "BMPName" And pPropertyName <> "BMPID" Then
                         For Each pControl In .Controls
                            If pControl.name = pPropertyName Then
                                If ((TypeOf pControl Is TextBox) Or (TypeOf pControl Is ComboBox)) Then
                                    pControl.Text = pPropertyValue
                                ElseIf (TypeOf pControl Is OptionButton) Then
                                    pControl.value = True
                                ElseIf (TypeOf pControl Is CheckBox) Then
                                    If CBool(pPropertyValue) Then
                                        pControl.value = 1
                                    Else
                                        pControl.value = 0
                                    End If
                                End If
                            End If
                        Next pControl
                    End If
            End Select
        Next i
    End With
    For i = 0 To (tmpBmpDict.Count - 1)
        pPropertyName = tmpBmpDict.keys(i)
        pPropertyValue = tmpBmpDict.Item(pPropertyName)
        gBMPDetailDict.add pPropertyName, pPropertyValue
    Next i
    
    GoTo CleanUp
    
ShowError:
    MsgBox "setBmp2Props :", Err.description
    
CleanUp:
    Set pBMPKeys = Nothing
    Set pControl = Nothing
    Set tmpBmpDict = Nothing
End Sub

'Initialize the Pollutant decay and removal rates
Public Sub InitPollutantData(Optional pBMPType As String)
On Error GoTo ShowError

    'If (IsEmpty(gPollutants)) Then
        ModuleDecayFact.CreatePollutantList
    'End If

    Dim i As Integer
    gMaxPollutants = UBound(gPollutants) + 1
    ReDim gParamInfos(0 To gMaxPollutants - 1)
    
    For i = 0 To (gMaxPollutants - 1)
        With gParamInfos(i)
            .name = gPollutants(i)
            .Decay = 0.2
            .PctRem = 0.1
            
            If (pBMPType <> "") Then
                 Select Case pBMPType
                     Case "BioRetentionBasin":
                         If (StringContains(gPollutants(i), "SEDIMENT")) Then
                            .Decay = 1.2
                            .PctRem = 0.9
                         ElseIf (StringContains(gPollutants(i), "BOD")) Then
                            .Decay = 1#
                            .PctRem = 0.8
                         ElseIf (StringContains(gPollutants(i), "NITROGEN")) Then
                            .Decay = 0.8
                            .PctRem = 0.4
                         ElseIf (StringContains(gPollutants(i), "PHOSPHOR")) Then
                            .Decay = 0.8
                            .PctRem = 0.7
                         ElseIf (StringContains(gPollutants(i), "ZINC")) Then
                            .Decay = 0.8
                            .PctRem = 0.95
                         End If
                     Case "WetPond":
                         If (StringContains(gPollutants(i), "SEDIMENT")) Then
                            .Decay = 2#
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "BOD")) Then
                            .Decay = 1.5
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "NITROGEN")) Then
                            .Decay = 1
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "PHOSPHOR")) Then
                            .Decay = 1
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "ZINC")) Then
                            .Decay = 1
                            .PctRem = 0
                         End If

                     Case "DryPond":
                         If (StringContains(gPollutants(i), "SEDIMENT")) Then
                            .Decay = 1.2
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "BOD")) Then
                            .Decay = 1#
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "NITROGEN")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "PHOSPHOR")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "ZINC")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         End If
                     Case "GreenRoof":
                         If (StringContains(gPollutants(i), "SEDIMENT")) Then
                            .Decay = 1.2
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "BOD")) Then
                            .Decay = 1#
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "NITROGEN")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "PHOSPHOR")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         ElseIf (StringContains(gPollutants(i), "ZINC")) Then
                            .Decay = 0.8
                            .PctRem = 0
                         End If
                     Case "PorousPavement":
                        
                         If (StringContains(gPollutants(i), "SEDIMENT")) Then
                            .Decay = 1.2
                            .PctRem = 0.9
                         ElseIf (StringContains(gPollutants(i), "BOD")) Then
                            .Decay = 1#
                            .PctRem = 0.8
                         ElseIf (StringContains(gPollutants(i), "NITROGEN")) Then
                            .Decay = 0.8
                            .PctRem = 0.4
                         ElseIf (StringContains(gPollutants(i), "PHOSPHOR")) Then
                            .Decay = 0.8
                            .PctRem = 0.7
                         ElseIf (StringContains(gPollutants(i), "ZINC")) Then
                            .Decay = 0.8
                            .PctRem = 0.95
                         End If
                     Case Else

                 End Select
             End If
        End With
    Next i

   Exit Sub
    
ShowError:
    MsgBox "InitPollutantData: " & Err.description
End Sub

'Initialize the Pollutant decay and removal rates for
'an existing BMP
Public Sub LoadPollutantData(pBmpDetailDict As Scripting.Dictionary)
  On Error GoTo ErrorHandler

    If pBmpDetailDict Is Nothing Then Exit Sub
    
    'If UBound(gPollutants) <= 0 Then
        ModuleDecayFact.CreatePollutantList
    'End If
    Dim i As Integer
    gMaxPollutants = UBound(gPollutants) + 1
    ReDim gParamInfos(0 To gMaxPollutants - 1)
    For i = 0 To gMaxPollutants - 1
        With gParamInfos(i)
            .name = gPollutants(i)
            If pBmpDetailDict.Exists("Decay" & i + 1) Then
                .Decay = pBmpDetailDict.Item("Decay" & i + 1)
            Else
                .Decay = 0.2
            End If
            If pBmpDetailDict.Exists("PctRem" & i + 1) Then
                .PctRem = pBmpDetailDict.Item("PctRem" & i + 1)
            Else
                .PctRem = 0.1
            End If
            If pBmpDetailDict.Exists("K" & i + 1) Then
                .K = pBmpDetailDict.Item("K" & i + 1)
            Else
                .K = 0.1
            End If
            If pBmpDetailDict.Exists("C" & i + 1) Then
                .C = pBmpDetailDict.Item("C" & i + 1)
            Else
                .C = 0.1
            End If
        End With
    Next i
    
    'New default Decay & percent removal factors
    

  Exit Sub
ErrorHandler:
  HandleError True, "LoadPollutantData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub



'Initialize the Evaluation factors for a new assessment point
'''Public Sub InitEvaluationFactors()
'''    Call InitializeMapDocument
'''    Call InitializeOperators
'''    If UBound(gPollutants) <= 0 Then
'''        ModuleDecayFact.CreatePollutantList
'''    End If
'''    Dim i As Integer
'''    gMaxPollutants = UBound(gPollutants) + 1
'''    ReDim gAssessInfos(0 To gMaxPollutants)
'''    'First parameter is flow
'''    gAssessInfos(0).Factor = "Flow"
'''    gAssessInfos(0).Unit = "ft3/yr"
'''    gAssessInfos(0).isRedEval = True
'''    gAssessInfos(0).Target = 10
'''    gAssessInfos(0).isTargetEval = True
'''    gAssessInfos(0).Reduction = 0
'''
'''    For i = 1 To gMaxPollutants
'''        With gAssessInfos(i)
'''            .Factor = gPollutants(i - 1)
'''            .Unit = "ton/yr"
'''            .isRedEval = False
'''            .Target = 10
'''            .isTargetEval = False
'''            .Reduction = 0
'''        End With
'''    Next i
'''End Sub


'''Load the Evaluation factors for an existing assessment point
''Public Sub LoadEvaluationFactors(pBMPDetailDict As Scripting.Dictionary)
''    Call InitializeMapDocument
''    Call InitializeOperators
''    If UBound(gPollutants) <= 0 Then
''        ModuleDecayFact.CreatePollutantList
''    End If
''
''    Dim i As Integer
''    gMaxPollutants = UBound(gPollutants) + 1
''    ReDim gAssessInfos(0 To gMaxPollutants)
''
''    With gAssessInfos(0)
''        .Factor = "Flow"
''        .Unit = "ft3/yr"
''        If pBMPDetailDict.Exists("isFlowVolEval") Then
''            .isTargetEval = pBMPDetailDict.Item("isFlowVolEval")
''            If pBMPDetailDict.Exists("FlowVol") Then
''                .Target = pBMPDetailDict.Item("FlowVol")
''            End If
''        Else
''            .isTargetEval = False
''        End If
''        If pBMPDetailDict.Exists("isFlowRednEval") Then
''            .isRedEval = pBMPDetailDict.Item("isFlowRednEval")
''            If pBMPDetailDict.Exists("FlowRedn") Then
''                .Reduction = pBMPDetailDict.Item("FlowRedn")
''            End If
''        Else
''            .isRedEval = False
''        End If
''    End With
''
''
''    For i = 1 To gMaxPollutants
''        With gAssessInfos(i)
''            .Factor = gPollutants(i - 1)
''            .Unit = "ton/yr"
''            If pBMPDetailDict.Exists("isParam" & i & "LoadEval") Then
''                .isTargetEval = pBMPDetailDict.Item("isParam" & i & "LoadEval")
''                If pBMPDetailDict.Exists("Param" & i & "Load") Then
''                    .Target = pBMPDetailDict.Item("Param" & i & "Load")
''                End If
''            Else
''                .isTargetEval = False
''            End If
''            If pBMPDetailDict.Exists("isParam" & i & "RednEval") Then
''                .isRedEval = pBMPDetailDict.Item("isParam" & i & "RednEval")
''                If pBMPDetailDict.Exists("Param" & i & "Redn") Then
''                    .Reduction = pBMPDetailDict.Item("Param" & i & "Redn")
''                End If
''            Else
''                .isRedEval = False
''            End If
''        End With
''    Next i
''End Sub

Public Sub InitializeAggBMPTypes()
  On Error GoTo ErrorHandler

    ' Load the BMPS with their Types.......
    Set gBMPTypeDict = New Scripting.Dictionary
    gBMPTypeDict.RemoveAll
    gBMPTypeDict.add "Bioretention", "Aggregate"
    gBMPTypeDict.add "Dry Pond", "Aggregate"
    gBMPTypeDict.add "Wet Pond", "Aggregate"
    gBMPTypeDict.add "Rain Barrel", "Aggregate"
    gBMPTypeDict.add "Cistern", "Aggregate"
    gBMPTypeDict.add "Porous Pavement", "Aggregate"
    gBMPTypeDict.add "Green Roof", "Aggregate"
    gBMPTypeDict.add "Infiltration Trench", "Aggregate"
    gBMPTypeDict.add "Vegetative Swale", "Aggregate"
    gBMPTypeDict.add "Conduit", "Aggregate"
    ' Load the BMPS with their Category.......
    Set gBMPCatDict = New Scripting.Dictionary
    gBMPCatDict.RemoveAll
    gBMPCatDict.add "Bioretention", "On-Site Treatment"
    gBMPCatDict.add "Dry Pond", "Regional Storage/Treatment"
    gBMPCatDict.add "Wet Pond", "Regional Storage/Treatment"
    gBMPCatDict.add "Rain Barrel", "On-Site Interception"
    gBMPCatDict.add "Cistern", "On-Site Interception"
    gBMPCatDict.add "Porous Pavement", "On-Site Treatment"
    gBMPCatDict.add "Green Roof", "On-Site Interception"
    gBMPCatDict.add "Infiltration Trench", "On-Site Treatment"
    gBMPCatDict.add "Vegetative Swale", "Routing Attenuation"
    gBMPCatDict.add "Conduit", "Routing Attenuation"
            
    ' Initialize the Dict.....
    Set gBMPPlacedDict = CreateObject("Scripting.Dictionary")
    

  Exit Sub
ErrorHandler:
  HandleError True, "InitializeAggBMPTypes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub


Public Sub SetCostAdjDict(costAdjDict As Scripting.Dictionary)
On Error GoTo ShowError
    'Set costAdjDict = New Scripting.Dictionary

    Dim ConnStr As String
    'ConnStr = "D:\SUSTAIN\CostDB\BMPCosts.mdb"
    If Trim(gCostDBpath) = "" Then Exit Sub
    ConnStr = Trim(gCostDBpath)
    
    Dim pAdoConn As ADODB.Connection
    Set pAdoConn = New ADODB.Connection
    pAdoConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ConnStr & ";"
    
    Dim curYear As Integer
    Dim costAdj As Double
    costAdj = 1

    Dim curIndex As Double
    curIndex = 0#
    Dim maxIndex As Double

    Dim pRs As ADODB.Recordset
    Set pRs = New ADODB.Recordset
    
    Dim strSql As String
    strSql = "SELECT Year, Dec_CCI" & _
            " From ConsolidatedDecCCI " & _
            " WHERE Year =( Select Max(Year) from ConsolidatedDecCCI) " & _
            " AND ucase(Location)='NATIONAL'"
    pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
    
    pRs.MoveFirst
    Dim maxCCIYear As Integer
    If Not pRs.EOF Then
        maxCCIYear = pRs("Year")
        maxIndex = pRs("Dec_CCI")
    End If
    pRs.Close
    
'    strSql = "SELECT Year, Dec_CCI" & _
'            " From ConsolidatedDecCCI " & _
'            " WHERE Year = " & curYear & _
'            " AND ucase(Location)='NATIONAL'"
    strSql = "SELECT Year, Dec_CCI" & _
            " From ConsolidatedDecCCI " & _
            " WHERE ucase(Location)='NATIONAL'"
    pRs.Open strSql, pAdoConn, adOpenDynamic, adLockOptimistic
    
    pRs.MoveFirst
    If Not pRs.EOF Then
        Do Until pRs.EOF
            curYear = pRs("Year")
            curIndex = pRs("Dec_CCI")
            If curIndex > 0 Then
                costAdj = maxIndex / curIndex
                costAdjDict.Item(curYear) = costAdj
            End If
            pRs.MoveNext
            curIndex = 0#
        Loop
    End If
    
    GoTo CleanUp
    
ShowError:
    MsgBox "Error in SetCostAdjDict 1:" & Err.description
CleanUp:
    Set pRs = Nothing
    pAdoConn.Close
    Set pAdoConn = Nothing
End Sub


Public Sub InitBmpTypeCatDict()
  On Error GoTo ErrorHandler


    ' Load the BMPS with their Types.......
    Set gBMPTypeDict = New Scripting.Dictionary
    gBMPTypeDict.add "Bioretention", "Point"
    gBMPTypeDict.add "Dry Pond", "Point"
    gBMPTypeDict.add "Wet Pond", "Point"
    gBMPTypeDict.add "Rain Barrel", "Point"
    gBMPTypeDict.add "Cistern", "Point"
    gBMPTypeDict.add "Porous Pavement", "Area"
    gBMPTypeDict.add "Green Roof", "Area"
    gBMPTypeDict.add "Infiltration Trench", "Line"
    gBMPTypeDict.add "Vegetative Swale", "Line"
    gBMPTypeDict.add "Buffer Strip", "Line"
    ' Load the BMPS with their Category.......
    Set gBMPCatDict = New Scripting.Dictionary
    gBMPCatDict.add "Bioretention", "Low-Impact Development Practices"
    gBMPCatDict.add "Dry Pond", "Conventional Practices"
    gBMPCatDict.add "Wet Pond", "Conventional Practices"
    gBMPCatDict.add "Rain Barrel", "Low-Impact Development Practices"
    gBMPCatDict.add "Cistern", "Low-Impact Development Practices"
    gBMPCatDict.add "Porous Pavement", "Low-Impact Development Practices"
    gBMPCatDict.add "Green Roof", "Low-Impact Development Practices"
    gBMPCatDict.add "Infiltration Trench", "Conventional Practices"
    gBMPCatDict.add "Vegetative Swale", "Conventional Practices"
    gBMPCatDict.add "Buffer Strip", "Generalized Practices"


  Exit Sub
ErrorHandler:
  HandleError True, "InitBmpTypeCatDict " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Sub
Public Function Get_Tab_Index(ByVal strTab As String) As Integer
  On Error GoTo ErrorHandler


    Select Case strTab
        Case "Point"
            Get_Tab_Index = 0
        Case "Line"
            Get_Tab_Index = 1
        Case "Area"
            Get_Tab_Index = 2
        Case "Aggregate"
            Get_Tab_Index = 3
    End Select


  Exit Function
ErrorHandler:
  HandleError True, "Get_Tab_Index " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.description, 4
End Function

