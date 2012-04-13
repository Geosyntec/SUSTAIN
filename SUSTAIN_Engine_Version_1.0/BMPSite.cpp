// BMPSite.cpp: implementation of the CBMPSite class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BMPSite.h"
#include "StringToken.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBMPSite::CBMPSite()
{
	m_strID = "";
	m_strName = "";
	m_strType = "";
	m_bChecked = false;
	m_lfBMPUnit = 1.0;
	m_nBMPClass = 0;
	m_lfDDarea = 0.0;
	m_lfAccDArea = 0.0;
	m_lfCost = 0.0;

	m_lfSurfaceArea = 0.0;
	m_lfExcavatnVol = 0.0;
	m_lfSurfStorVol = 0.0;
	m_lfSoilStorVol = 0.0;
	m_lfUdrnStorVol	= 0.0;

	m_lfThreshFlow = 0.0;
	m_lfSiteDArea = 0.0;
	m_pDecay = NULL;
	m_pK = NULL;
	m_pCstar = NULL;
	m_pConc = NULL;
	m_pUndRemoval = NULL;
	m_pSiteProp = NULL;
	m_preLU = NULL;
	m_RAConc = NULL;
	m_fileOut = NULL;

	m_nQualNum = 0;
	m_nTSNum = 0;
	m_strCostFile = "";
	m_nBreakPoints = 0;
	m_lfBreakPtID = 0.0;
	m_TradeOff = NULL;
	m_pDataMixLU = NULL;
	m_pDataPreLU = NULL;
	m_lfROpeakMixLU = 0.0;
	m_lfROpeakPreLU = 0.0;
	m_tmStart = COleDateTime(1890,1,1,0,0,0);

	m_sediment.m_lfBEDWID = 0.0;
	m_sediment.m_lfBEDDEP = 0.0;
	m_sediment.m_lfBEDPOR = 0.0;
	m_sediment.m_lfSAND_FRAC = 0.0;
	m_sediment.m_lfSILT_FRAC = 0.0;
	m_sediment.m_lfCLAY_FRAC = 0.0;

	m_sand.m_lfD = 0.0;
	m_sand.m_lfW = 0.0;
	m_sand.m_lfRHO = 0.0;
	m_sand.m_lfKSAND = 0.0;
	m_sand.m_lfEXPSND = 0.0;

	m_silt.m_lfD = 0.0;
	m_silt.m_lfW = 0.0;
	m_silt.m_lfRHO = 0.0;
	m_silt.m_lfTAUCD = 0.0;
	m_silt.m_lfTAUCS = 0.0;
	m_silt.m_lfM = 0.0;

	m_clay.m_lfD = 0.0;
	m_clay.m_lfW = 0.0;
	m_clay.m_lfRHO = 0.0;
	m_clay.m_lfTAUCD = 0.0;
	m_clay.m_lfTAUCS = 0.0;
	m_clay.m_lfM = 0.0;

	m_nInfiltMethod = 0;
	m_nPolRotMethod = 1;
	m_nPolRemMethod = 0;
	m_nGAInfil_Index = 0;
	m_bUndSwitch = false;
	m_lfSoilDepth = 0.0;
	m_lfPorosity = 0.0;
	m_lfFCapacity = 0.3;
	m_lfWPoint = 0.15;
	m_lfUndDepth = 0.0;
	m_lfUndVoid = 0.0;
	m_lfUndInfilt = 0.0;
	m_holtanParam.m_lfVegA = 0.0;
	m_holtanParam.m_lfFInfilt = 0.0;
	for (int i=0; i<12; i++)
		m_holtanParam.m_lfGrowth[i] = 0.0;

	m_costParam.m_lfLinearCost = 0.0;			
	m_costParam.m_lfAreaCost = 0.0;            				
	m_costParam.m_lfTotalVolumeCost = 0.0;     				
	m_costParam.m_lfMediaVolumeCost = 0.0;     				
	m_costParam.m_lfUnderDrainVolumeCost = 0.0;				
	m_costParam.m_lfConstantCost = 0.0;        				
	m_costParam.m_lfPercentCost = 0.0;         
	m_costParam.m_lfLengthExp = 0.0;         
	m_costParam.m_lfAreaExp = 0.0;         
	m_costParam.m_lfTotalVolExp = 0.0;         
	m_costParam.m_lfMediaVolExp = 0.0;         
	m_costParam.m_lfUDVolExp = 0.0;         
	
	while (qFlow.size() != 24)
		qFlow.push(0.0);
}

CBMPSite::CBMPSite(const CString& strID, const CString& strName, const CString& strType, int bmpClass)
{
	m_strID = strID;
	m_strName = strName;
	m_strType = strType;
	m_bChecked = false;
	m_lfBMPUnit = 1.0;
	m_nBMPClass = bmpClass;
	m_lfDDarea = 0.0;
	m_lfAccDArea = 0.0;
	m_lfCost = 0.0;		

	m_lfSurfaceArea = 0.0;
	m_lfExcavatnVol = 0.0;
	m_lfSurfStorVol = 0.0;
	m_lfSoilStorVol = 0.0;
	m_lfUdrnStorVol	= 0.0;

	m_lfThreshFlow = 0.0;
	m_lfSiteDArea = 0.0;
	m_pDecay = NULL;
	m_pK = NULL;
	m_pCstar = NULL;
	m_pConc = NULL;
	m_pUndRemoval = NULL;

	switch (m_nBMPClass)
	{
		case CLASS_A:
		{
			BMP_A* pBmpA = new BMP_A;
			pBmpA->m_lfBasinWidth = 0.0;
			pBmpA->m_lfBasinLength = 0.0;
			pBmpA->m_lfOrificeHeight = 0.0;
			pBmpA->m_lfOrificeDiameter = 0.0;
			pBmpA->m_nExitType = -1;
			pBmpA->m_nWeirType = -1;
			pBmpA->m_lfWeirHeight = 0.0;
			pBmpA->m_lfWeirWidth = 0.0;
			pBmpA->m_lfWeirAngle = 0.0;
			pBmpA->m_nORelease = -1;
			pBmpA->m_nPeople = -1;
			pBmpA->m_nDays = -1;
			for(int i=0; i<24; i++)
				pBmpA->m_lfRelease[i] = 0.0;

			// green-ampt parameters
			pBmpA->m_pGAInfil.F = 0.0;
			pBmpA->m_pGAInfil.FU = 0.0;
			pBmpA->m_pGAInfil.FUmax = 0.0;
			pBmpA->m_pGAInfil.IMD = 0.0;
			pBmpA->m_pGAInfil.IMDmax = 0.0;
			pBmpA->m_pGAInfil.Ks = 0.0;
			pBmpA->m_pGAInfil.L = 0.0;
			pBmpA->m_pGAInfil.S = 0.0;
			pBmpA->m_pGAInfil.Sat = FALSE;
			pBmpA->m_pGAInfil.T = 0.0;

			m_pSiteProp = (void*) pBmpA;
			break;
		}
		case CLASS_B:
		{
			BMP_B* pBmpB = new BMP_B;
			pBmpB->m_lfBasinWidth = 0.0;
			pBmpB->m_lfBasinLength = 0.0;
			pBmpB->m_lfMaximumDepth = 0.0;
			pBmpB->m_lfSideSlope1 = 0.0;
			pBmpB->m_lfSideSlope2 = 0.0;
			pBmpB->m_lfSideSlope3 = 0.0;
			pBmpB->m_lfManning = 0.0;

			// green-ampt parameters
			pBmpB->m_pGAInfil.F = 0.0;
			pBmpB->m_pGAInfil.FU = 0.0;
			pBmpB->m_pGAInfil.FUmax = 0.0;
			pBmpB->m_pGAInfil.IMD = 0.0;
			pBmpB->m_pGAInfil.IMDmax = 0.0;
			pBmpB->m_pGAInfil.Ks = 0.0;
			pBmpB->m_pGAInfil.L = 0.0;
			pBmpB->m_pGAInfil.S = 0.0;
			pBmpB->m_pGAInfil.Sat = FALSE;
			pBmpB->m_pGAInfil.T = 0.0;

			m_pSiteProp = (void*) pBmpB;
			break;
		}
		case CLASS_C:	// conduit (01-2005)
		{
			BMP_C* pBmpC = new BMP_C;
			pBmpC->m_strCondType = "DUMMY";
			pBmpC->m_strID = "";
			pBmpC->m_nIndex = -1;
	
			// --- initialize link properties
			pBmpC->m_pTLink.ID = "";				    // link ID
			pBmpC->m_pTLink.type = CONDUIT;				// link type code
			pBmpC->m_pTLink.subIndex = -1;			    // index of link's sub-category
			pBmpC->m_pTLink.rptFlag = FALSE;		    // reporting flag
			pBmpC->m_pTLink.node1 = 0;				    // start node index
			pBmpC->m_pTLink.node2 = 0;				    // end node index
			pBmpC->m_pTLink.z1 = 0.;				    // upstrm invert ht. above node invert (ft)
			pBmpC->m_pTLink.z2 = 0.;				    // downstrm invert ht. above node invert (ft)
			pBmpC->m_pTLink.q0 = 0.;				    // initial flow (cfs)
			pBmpC->m_pTLink.qLimit = 0.;			    // constraint on max. flow (cfs)
			pBmpC->m_pTLink.cLossInlet = 0.;		    // inlet loss coeff.
			pBmpC->m_pTLink.cLossOutlet = 0.;		    // outlet loss coeff.
			pBmpC->m_pTLink.cLossAvg = 0.;			    // avg. loss coeff.
			pBmpC->m_pTLink.hasFlapGate = FALSE;	    // true if flap gate present
			pBmpC->m_pTLink.oldFlow = 0.;			    // previous flow rate (cfs)
			pBmpC->m_pTLink.newFlow = 0.;			    // current flow rate (cfs)
			pBmpC->m_pTLink.oldDepth = 0.;			    // previous flow depth (ft)
			pBmpC->m_pTLink.newDepth = 0.;			    // current flow depth (ft)
			pBmpC->m_pTLink.oldVolume = 0.;				// previous flow volume (ft3)
			pBmpC->m_pTLink.newVolume = 0.;				// current flow volume (ft3)
			pBmpC->m_pTLink.qFull = 0.;					// flow when full (cfs)
			pBmpC->m_pTLink.setting = 0.;			    // control setting
			pBmpC->m_pTLink.froude = 0.;			    // Froude number
			pBmpC->m_pTLink.oldQual = NULL;				// previous quality state
			pBmpC->m_pTLink.newQual = NULL;				// current quality state
			pBmpC->m_pTLink.flowClass = -1;				// flow classification
			pBmpC->m_pTLink.dqdh = 0.;				    // change in flow w.r.t. head (ft2/sec)
			pBmpC->m_pTLink.direction = -1;				// flow direction flag
			pBmpC->m_pTLink.isClosed = FALSE;		    // flap gate closed flag
			pBmpC->m_pTLink.xsect.type = -1;		    // type code of cross section shape
			pBmpC->m_pTLink.xsect.transect = -1;	    // index of transect (if applicable)
			pBmpC->m_pTLink.xsect.yFull = 0.;		    // depth when full (ft)
			pBmpC->m_pTLink.xsect.wMax = 0.;		    // width at widest point (ft)
			pBmpC->m_pTLink.xsect.aFull = 0.;		    // area when full (ft2)
			pBmpC->m_pTLink.xsect.rFull = 0.;		    // hyd. radius when full (ft)
			pBmpC->m_pTLink.xsect.sFull = 0.;		    // section factor when full (ft^4/3)
			pBmpC->m_pTLink.xsect.sMax = 0.;		    // section factor at max. flow (ft^4/3)
			pBmpC->m_pTLink.xsect.yBot = 0.;		    // depth of bottom section
			pBmpC->m_pTLink.xsect.aBot = 0.;		    // area of bottom section
			pBmpC->m_pTLink.xsect.sBot = 0.;		    // slope of bottom section
			pBmpC->m_pTLink.xsect.rBot = 0.;			// radius of bottom section
			
			// --- initialize conduit properties
			pBmpC->m_pTConduit.length = 0.;				// conduit length (ft)
			pBmpC->m_pTConduit.roughness = 0.;		    // Manning's n
			pBmpC->m_pTConduit.barrels = 1;				// number of barrels
			pBmpC->m_pTConduit.modLength = 0.;		    // modified conduit length (ft)
			pBmpC->m_pTConduit.roughFactor = 0.;	    // roughness factor for DW routing
			pBmpC->m_pTConduit.slope = 0.;			    // slope
			pBmpC->m_pTConduit.beta = 0.;	    		// discharge factor
			pBmpC->m_pTConduit.qMax = 0.;	    		// max. flow (cfs)
			pBmpC->m_pTConduit.a1 = 0.;		    	    // upstream areas (ft2)
			pBmpC->m_pTConduit.a2 = 0.;		    	    // downstream areas (ft2)
			pBmpC->m_pTConduit.q1 = 0.;		    	    // upstream flows per barrel (cfs)
			pBmpC->m_pTConduit.q2 = 0.;					// downstream flows per barrel (cfs)
			pBmpC->m_pTConduit.q1Old = 0.;			    // previous values of q1 & q2 (cfs)
			pBmpC->m_pTConduit.q2Old = 0.;			    // previous values of q1 & q2 (cfs)
			pBmpC->m_pTConduit.superCritical = FALSE;	// super-critical flow flag
			pBmpC->m_pTConduit.hasLosses = FALSE;	    // local losses flag
			
			// --- initialize transect properties
			pBmpC->m_pTTransect.ID = "";                // section ID
			pBmpC->m_pTTransect.yFull = 0.;             // depth when full (ft)
			pBmpC->m_pTTransect.aFull = 0.;             // area when full (ft2)
			pBmpC->m_pTTransect.rFull = 0.;             // hyd. radius when full (ft)
			pBmpC->m_pTTransect.wMax = 0.;              // width at widest point (ft)
			pBmpC->m_pTTransect.sMax = 0.;              // section factor at max. flow (ft^4/3)
			pBmpC->m_pTTransect.aMax = 0.;              // area at max. flow (ft2)
			pBmpC->m_pTTransect.roughness = 0.;         // Manning's n
			for (int i=0; i<N_TRANSECT_TBL; i++)
			{
				pBmpC->m_pTTransect.areaTbl[i] = 0.0;	// table of area v. depth
				pBmpC->m_pTTransect.hradTbl[i] = 0.0;	// table of hyd. radius v. depth
				pBmpC->m_pTTransect.widthTbl[i] = 0.0;	// table of top width v. depth
			}
			pBmpC->m_pTTransect.nTbl = N_TRANSECT_TBL;	// size of geometry tablesnTbl = N_TRANSECT_TBL;

			m_pSiteProp = (void*) pBmpC;
			break;
		}
		case CLASS_D:
		{
			BMP_D* pBmpD = new BMP_D;
			pBmpD->m_strName = "";
			pBmpD->m_strID = "";			
			pBmpD->m_nSegments = 1;			
			pBmpD->m_lfWidth = 0.0;
			pBmpD->m_lfLength = 0.0;
			pBmpD->m_lfVKS = 0.0;
			pBmpD->m_lfSAV = 0.0;
			pBmpD->m_lfOS = 0.0;
			pBmpD->m_lfOI = 0.0;
			pBmpD->m_lfSM = 0.0;
			pBmpD->m_lfSCHK = 0.0;
			pBmpD->m_lfSS = 0.0;
			pBmpD->m_lfVN = 0.0;
			pBmpD->m_lfH = 0.0;
			pBmpD->m_lfVN2 = 0.0;
			pBmpD->m_nICO = 0.0;
			for (int i=0; i<3; i++)
			{
				pBmpD->m_nNPART[i] = 0;
				pBmpD->m_lfCOARSE[i] = 0.0;
				pBmpD->m_lfPOR[i] = 0.0;
				pBmpD->m_lfDP[i] = 0.0;
				pBmpD->m_lfSG[i] = 0.0;
			}
			pBmpD->m_pSEGMENT_D = NULL;		
			pBmpD->m_pPOLLUTANT_D = NULL;
			m_pSiteProp = (void*) pBmpD;
			break;
		}
		default:
			m_pSiteProp = NULL;
			break;
	}

	m_preLU = NULL;
	m_RAConc = NULL;
	m_fileOut = NULL;

	m_nQualNum = 0;
	m_nTSNum = 0;
	m_strCostFile = "";
	m_nBreakPoints = 0;
	m_lfBreakPtID = 0.0;
	m_TradeOff = NULL;
	m_pDataMixLU = NULL;
	m_pDataPreLU = NULL;
	m_lfROpeakMixLU = 0.0;
	m_lfROpeakPreLU = 0.0;
	m_tmStart = COleDateTime(1890,1,1,0,0,0);

	m_sediment.m_lfBEDWID = 0.0;
	m_sediment.m_lfBEDDEP = 0.0;
	m_sediment.m_lfBEDPOR = 0.0;
	m_sediment.m_lfSAND_FRAC = 0.0;
	m_sediment.m_lfSILT_FRAC = 0.0;
	m_sediment.m_lfCLAY_FRAC = 0.0;

	m_sand.m_lfD = 0.0;
	m_sand.m_lfW = 0.0;
	m_sand.m_lfRHO = 0.0;
	m_sand.m_lfKSAND = 0.0;
	m_sand.m_lfEXPSND = 0.0;

	m_silt.m_lfD = 0.0;
	m_silt.m_lfW = 0.0;
	m_silt.m_lfRHO = 0.0;
	m_silt.m_lfTAUCD = 0.0;
	m_silt.m_lfTAUCS = 0.0;
	m_silt.m_lfM = 0.0;

	m_clay.m_lfD = 0.0;
	m_clay.m_lfW = 0.0;
	m_clay.m_lfRHO = 0.0;
	m_clay.m_lfTAUCD = 0.0;
	m_clay.m_lfTAUCS = 0.0;
	m_clay.m_lfM = 0.0;

	m_nInfiltMethod = 0;
	m_nPolRotMethod = 1;
	m_nPolRemMethod = 0;
	m_nGAInfil_Index = 0;
	m_bUndSwitch = false;
	m_lfSoilDepth = 0.0;
	m_lfPorosity = 0.0;
	m_lfFCapacity = 0.3;
	m_lfWPoint = 0.15;
	m_lfUndDepth = 0.0;
	m_lfUndVoid = 0.0;
	m_lfUndInfilt = 0.0;
	m_holtanParam.m_lfVegA = 0.0;
	m_holtanParam.m_lfFInfilt = 0.0;
	for (int i=0; i<12; i++)
		m_holtanParam.m_lfGrowth[i] = 0.0;

	m_costParam.m_lfLinearCost = 0.0;			
	m_costParam.m_lfAreaCost = 0.0;            				
	m_costParam.m_lfTotalVolumeCost = 0.0;     				
	m_costParam.m_lfMediaVolumeCost = 0.0;     				
	m_costParam.m_lfUnderDrainVolumeCost = 0.0;				
	m_costParam.m_lfConstantCost = 0.0;        				
	m_costParam.m_lfPercentCost = 0.0;         
	m_costParam.m_lfLengthExp = 0.0;         
	m_costParam.m_lfAreaExp = 0.0;         
	m_costParam.m_lfTotalVolExp = 0.0;         
	m_costParam.m_lfMediaVolExp = 0.0;         
	m_costParam.m_lfUDVolExp = 0.0;         
	
	while (qFlow.size() != 24)
		qFlow.push(0.0);
}

CBMPSite::~CBMPSite()
{
	if (m_pSiteProp != NULL)
	{
		switch (m_nBMPClass)
		{
			case CLASS_A:
				{
					delete (BMP_A*)m_pSiteProp;
					break;
				}
			case CLASS_B:
				{
					delete (BMP_B*)m_pSiteProp;
					break;
				}
			case CLASS_C: 
				{
					BMP_C* pBmpC = (BMP_C*)m_pSiteProp;
					if (pBmpC->m_pTLink.newQual != NULL)
						delete []pBmpC->m_pTLink.newQual;
					if (pBmpC->m_pTLink.oldQual != NULL)
						delete []pBmpC->m_pTLink.oldQual;
					delete (BMP_C*)m_pSiteProp;
					break;
				}
			case CLASS_D:
				{
					BMP_D* pBmpD = (BMP_D*)m_pSiteProp;
					if (pBmpD->m_pSEGMENT_D != NULL)
						delete []pBmpD->m_pSEGMENT_D;
					if (pBmpD->m_pPOLLUTANT_D != NULL)
						delete []pBmpD->m_pPOLLUTANT_D;
					delete (BMP_D*)m_pSiteProp;
					break;
				}
			default:
				break;
		}
	}
	
	if (m_TradeOff != NULL)
		delete []m_TradeOff;
	if (m_pDecay != NULL)
		delete []m_pDecay;
	if (m_pK != NULL)
		delete []m_pK;
	if (m_pCstar != NULL)
		delete []m_pCstar;
	if (m_pConc != NULL)
		delete []m_pConc;
	if (m_pUndRemoval != NULL)
		delete []m_pUndRemoval;
	if (m_RAConc != NULL)
		delete []m_RAConc;

	if (m_pDataMixLU != NULL)
		delete []m_pDataMixLU;
	if (m_pDataPreLU != NULL)
		delete []m_pDataPreLU;

	m_siteluList.RemoveAll();
	m_sitepsList.RemoveAll();

	while (!qFlow.empty())
		qFlow.pop();
	while (!m_dsbmpsiteList.IsEmpty())
		delete (DS_BMPSITE*) m_dsbmpsiteList.RemoveTail();
	while (!m_usbmpsiteList.IsEmpty())
		delete (US_BMPSITE*) m_usbmpsiteList.RemoveTail();
	while (!m_adjustList.IsEmpty())
		delete (ADJUSTABLE_PARAM*) m_adjustList.RemoveTail();
	while (!m_factorList.IsEmpty())
		delete (EVALUATION_FACTOR*) m_factorList.RemoveTail();
}

double* CBMPSite::GetVariablePointer(CString strVarName)
{
	if (m_nBMPClass == CLASS_A)
	{
		BMP_A* pBmpA = (BMP_A*) m_pSiteProp;
		if (strVarName.CompareNoCase("Width") == 0)
			return &pBmpA->m_lfBasinWidth;
		else if (strVarName.CompareNoCase("Length") == 0)
			return &pBmpA->m_lfBasinLength;
		else if (strVarName.CompareNoCase("OrificeH") == 0)
			return &pBmpA->m_lfOrificeHeight;
		else if (strVarName.CompareNoCase("OrificeD") == 0)
			return &pBmpA->m_lfOrificeDiameter;
		else if (strVarName.CompareNoCase("WeirH") == 0)
			return &pBmpA->m_lfWeirHeight;
		else if (strVarName.CompareNoCase("SDepth") == 0)
			return &m_lfSoilDepth;
		else if (strVarName.CompareNoCase("CECurve") == 0)
			return &m_lfBreakPtID;
		else if (strVarName.CompareNoCase("NUMUNIT") == 0)
			return &m_lfBMPUnit;
		else
			return NULL;
	}
	else if (m_nBMPClass == CLASS_B)
	{
		BMP_B* pBmpB = (BMP_B*) m_pSiteProp;
		if (strVarName.CompareNoCase("Width") == 0)
			return &pBmpB->m_lfBasinWidth;
		else if (strVarName.CompareNoCase("Length") == 0)
			return &pBmpB->m_lfBasinLength;
		else if (strVarName.CompareNoCase("MaxDepth") == 0)
			return &pBmpB->m_lfMaximumDepth;
		else if (strVarName.CompareNoCase("SDepth") == 0)
			return &m_lfSoilDepth;
		else if (strVarName.CompareNoCase("CECurve") == 0)
			return &m_lfBreakPtID;
		else if (strVarName.CompareNoCase("NUMUNIT") == 0)
			return &m_lfBMPUnit;
		else
			return NULL;
	}
	else
	{
		if (strVarName.CompareNoCase("CECurve") == 0)
			return &m_lfBreakPtID;
	}
	
	return NULL;
}

double CBMPSite::GetBMPArea()
{
	if (m_nBMPClass == CLASS_A)
	{
		BMP_A* pBmpA = (BMP_A*) m_pSiteProp;
		return pBmpA->m_lfBasinWidth*pBmpA->m_lfBasinLength;
	}
	else if (m_nBMPClass == CLASS_B)
	{
		BMP_B* pBmpB = (BMP_B*) m_pSiteProp;
		//calculate the top width (ft)
		double top_width = (pBmpB->m_lfMaximumDepth/pBmpB->m_lfSideSlope1 
			+ pBmpB->m_lfMaximumDepth/pBmpB->m_lfSideSlope1 + pBmpB->m_lfBasinWidth); 
		//return pBmpB->m_lfBasinWidth*pBmpB->m_lfBasinLength;
		return top_width*pBmpB->m_lfBasinLength;
	}
	return NULL;
}

// load the time series data
bool CBMPSite::LoadWatershedTSData(CString BMPSiteID, COleDateTime startDate, 
	 COleDateTime endDate,CString strFileName, double* multiplier, int landfg)
{
	int i, j;
	char strLine[MAXLINE];
	CString str;

	FILE *fpin = NULL;
	// open the file for reading
	fpin = fopen (strFileName, "rt");
	if(fpin == NULL)
		return false;

	// skip first two lines
	i = 2;
	while(i-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	//read the reporting timestep in seconds
	fgets(strLine, MAXLINE, fpin);
	str = strLine;
	CStringToken strToken1(str);
	double delts = atof((LPCSTR)strToken1.NextToken());	//sec/ivl

	//read the number of constituents
	fgets(strLine, MAXLINE, fpin);
	str = strLine;
	CStringToken strToken2(str);
	m_nQualNum = atoi((LPCSTR)strToken2.NextToken());	// including flow

	// skip constituents names
	i = m_nQualNum;
	while(i-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	//read the number of nodes
	fgets(strLine, MAXLINE, fpin);
	str = strLine;
	CStringToken strToken(str);
	int nNodes = atoi((LPCSTR)strToken.NextToken());
	//int nID = -999;
	CString strID;

	// check if the BMPSITE exist in the file
	i = nNodes;
	while(i-- > 0 && !feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		str = strLine;
		CStringToken strToken(str);
		//nID = atoi((LPCSTR)strToken.NextToken()); // node ID
		strID = strToken.NextToken(); // node ID
		//if (nID == BMPSiteID)
		if (strID.CompareNoCase(BMPSiteID) == 0)
			break;
	}

	//if (nID != BMPSiteID)
	if (strID.CompareNoCase(BMPSiteID) != 0)
		return true;

	// skip lines until find the key word "node"
	CString strFind = "node";
	while (!feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		str = strLine;
		str.MakeLower();
		if(str.Find(strFind) != -1)
			break;
	}

	// count the time series numbers
	COleDateTimeSpan tsSpan = endDate - startDate;
	m_nTSNum = (long)tsSpan.GetTotalHours() + 24;
	
    // read first data line for starting date of the time series data
	long nStart = ftell (fpin);
	fgets(strLine, MAXLINE, fpin);
    fseek(fpin, nStart, SEEK_SET);

	int year, month, day, hour, min, sec;
	str = strLine;
	CStringToken strToken3(str);
	strToken3.NextToken(); // skip the node number
	str = strToken3.NextToken();
	year = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	month = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	day = atoi((LPCSTR)str);
	
	str = strToken3.NextToken();
	hour = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	min = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	sec = atoi((LPCSTR)str);

	m_tmStart = COleDateTime(year, month, day, 0, 0, 0);
	tsSpan = startDate - m_tmStart;

	// calculate how many lines we need to skip
	long nSkipLineNum = (long)tsSpan.GetTotalHours();

    // skip all lines the time stamp is before the specified start date
	while (nSkipLineNum-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	double* pData = NULL;
	if (landfg == 0)
	{
		if(m_pDataPreLU != NULL)
			delete []m_pDataPreLU;
		m_pDataPreLU = new double[m_nTSNum*m_nQualNum];
		pData = m_pDataPreLU;
	}
	else
	{
		if(m_pDataMixLU != NULL)
			delete []m_pDataMixLU;
		m_pDataMixLU = new double[m_nTSNum*m_nQualNum];
		pData = m_pDataMixLU;
	}
	
    // read the data
	i = 0;

    while (!feof(fpin))
    {
		// read one line
		fgets(strLine, MAXLINE, fpin);

		// get the data
		str = strLine;
		CStringToken strToken(str);
		//int nID = atoi((LPCSTR)strToken.NextToken()); // node ID
		strID = strToken.NextToken(); // node ID
		//if (nID != BMPSiteID)
		if (strID.CompareNoCase(BMPSiteID) != 0)
			continue;
		strToken.NextToken(); // year
		strToken.NextToken(); // month
		strToken.NextToken(); // day
		strToken.NextToken(); // hour
		strToken.NextToken(); // minute
		strToken.NextToken(); // second
		double flow = 0.0;

		for(j=0; j<m_nQualNum; j++)
		{
			str = strToken.NextToken();
			if (j == 0)	// flow (cfs)
			{
				flow = atof((LPCSTR)str) * delts;	//ft3/sec * sec/ivl
				*(pData++) = flow;					// ft3/hr (ivl = hr)

				// find the peak runoff (ft3/hr) 
				if (landfg == 0)
				{
					if (m_lfROpeakPreLU < flow)
						m_lfROpeakPreLU = flow;
				}
				else
				{
					if (m_lfROpeakMixLU < flow)
						m_lfROpeakMixLU = flow;
				}
			}
			else // quals (conc)
			{
				*(pData++) = atof((LPCSTR)str) * multiplier[j-1] * flow;	//lb/hr
			}
		}

		i++;
		if (i == m_nTSNum)
			break;
	}

	fclose(fpin);
	return (i == m_nTSNum);
}
// read the tradeoff curve solution file
bool CBMPSite::ReadTradeOffCurveCosts()
{
	int j, k;
	char strLine[MAXLINE];
	CString str;

	if (m_strCostFile == "")
		return true;

	FILE *fpin = NULL;
	// open the file for reading
	fpin = fopen (m_strCostFile, "rt");
	if(fpin == NULL)
		return false;

	// skip first three line
	j = 3;
	while(j-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);
	
	// read the data
	k = 1;
	while (!feof(fpin))
	{
		// read one line
		fgets(strLine, MAXLINE, fpin);

		// get the data
		str = strLine;
		CStringToken strToken1(str);
		int nBreak = atoi((LPCSTR)strToken1.NextToken());		// Target Break
		int nSolution = atoi((LPCSTR)strToken1.NextToken());	// Solution #
		double lfCost = atof((LPCSTR)strToken1.NextToken());	// total cost ($)

		if (nBreak == k && nSolution == 1)
		{
			m_TradeOff[k+2].m_lfCost = lfCost;	//0,1,2 for init, pre, and postdev conditions
			k++;
			if (k > m_nBreakPoints)
				break;
		}
	}

	fclose(fpin);

	return true;
}
    
// load the time series data
bool CBMPSite::LoadTradeOffCurveData(int BPindex,COleDateTime startDate,COleDateTime endDate)
{
	if (m_TradeOff == NULL)
		return true;

	int i = BPindex, j, k;
	FILE *fpin = NULL;
	char strLine[MAXLINE];
	CString str;

	// open the file for reading
	fpin = fopen (m_TradeOff[i].m_strBrPtFile, "rt");
	if(fpin == NULL)
		return false;

	// find the flag "date/time"
	while (!feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		str = strLine;
		str.MakeLower();

		// skip no line once found "date/time"
		if(str.Find("date/time") != -1)
		{
			j = 0;
			while(j-- > 0 && !feof(fpin))
				fgets(strLine, MAXLINE, fpin);
			break;
		}
	}
	
	// count the time series numbers
	COleDateTimeSpan tsSpan = endDate - startDate;
	m_TradeOff[i].m_nTSNum = (long)tsSpan.GetTotalHours() + 24;
	// read first data line for starting date of the time series data
	long nStart = ftell (fpin);
	fgets(strLine, MAXLINE, fpin);
	fseek(fpin, nStart, SEEK_SET);

	int year, month, day, hour, min;
	str = strLine;
	CStringToken strToken(str);
	strToken.NextToken(); // skip the dummy number
	str = strToken.NextToken();
	year = atoi((LPCSTR)str);

	str = strToken.NextToken();
	month = atoi((LPCSTR)str);

	str = strToken.NextToken();
	day = atoi((LPCSTR)str);
	
	str = strToken.NextToken();
	hour = atoi((LPCSTR)str);

	str = strToken.NextToken();
	min = atoi((LPCSTR)str);

	if(hour == 24)																	
	{                          
		hour = 0;                  
		m_TradeOff[i].m_tmStart = COleDateTime(year, month, day, hour, min, 0) + COleDateTimeSpan(1,0,0,0);
	}
	else
	{
		m_TradeOff[i].m_tmStart = COleDateTime(year, month, day, hour, min, 0);
	}

	tsSpan = startDate - m_TradeOff[i].m_tmStart;

	// calculate how many lines we need to skip
	long nSkipLineNum = (long)tsSpan.GetTotalHours();

	// skip all lines the time stamp is before the specified start date
	while (nSkipLineNum-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	if(m_TradeOff[i].m_pDataBrPt != NULL)
		delete []m_TradeOff[i].m_pDataBrPt;
	m_TradeOff[i].m_pDataBrPt = new double[m_TradeOff[i].m_nTSNum*m_TradeOff[i].m_nQualNum];

	// read the data
	k = 0;
	double *pData = m_TradeOff[i].m_pDataBrPt;
	
	while (!feof(fpin))
	{
		// read one line
		fgets(strLine, MAXLINE, fpin);

		// get the data
		str = strLine;
		CStringToken strToken1(str);
		strToken1.NextToken(); // serial number
		strToken1.NextToken(); // year
		strToken1.NextToken(); // month
		strToken1.NextToken(); // day
		strToken1.NextToken(); // hour
		strToken1.NextToken(); // minute
		strToken1.NextToken(); // volume
		strToken1.NextToken(); // stage     
		strToken1.NextToken(); // inflow_t  
		strToken1.NextToken(); // outflow_w 
		strToken1.NextToken(); // outflow_o 
		strToken1.NextToken(); // outflow_ud
		strToken1.NextToken(); // outflow_ut
		*(pData++) = atof((LPCSTR)strToken1.NextToken())*3600/3630;//cfs to in-ac/hr
		strToken1.NextToken(); // infiltration 
		strToken1.NextToken(); // percolation 
		strToken1.NextToken(); // evapotranspiration 
		strToken1.NextToken(); // seepage 

		for(j=1; j<m_TradeOff[i].m_nQualNum; j++)
		{
			strToken1.NextToken(); // inflow mass
			strToken1.NextToken(); // mass outflow from weir
			strToken1.NextToken(); // mass outflow from orifice
			strToken1.NextToken(); // mass outflow from underdrain
			strToken1.NextToken(); // mass outflow from untreated
			*(pData++) = atof((LPCSTR)strToken1.NextToken());// lb
			strToken1.NextToken(); // outflow concentration 
		}

		k++;
		if (k == m_TradeOff[i].m_nTSNum)
			break;
	}

	fclose(fpin);

	return true;
}
    
bool CBMPSite::UnLoadTradeOffCurveData(int nBrPtIndex)
{

	if (m_TradeOff == NULL)
		return true;

	if (m_nBreakPoints > 0)
	{
		if(m_TradeOff[nBrPtIndex].m_pDataBrPt != NULL)
		{
			delete []m_TradeOff[nBrPtIndex].m_pDataBrPt;
			m_TradeOff[nBrPtIndex].m_pDataBrPt = NULL;
		}
	}

	return true;
}



