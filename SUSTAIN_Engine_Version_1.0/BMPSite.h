// BMPSite.h: interface for the CBMPSite class.
//
//////////////////////////////////////////////////////////////////////
#include <queue>
using namespace std;

#if !defined(AFX_BMPSITE_H__D3CC98A1_98BD_4E1A_BDD9_42084729FD84__INCLUDED_)
#define AFX_BMPSITE_H__D3CC98A1_98BD_4E1A_BDD9_42084729FD84__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define	CLASS_A 1
#define	CLASS_B 2
#define	CLASS_C 3		
#define	CLASS_D 4		
#define	CLASS_X 100

//SWMM5
#include "SWMM5/headers.h"

typedef struct tag_ADJUSTABLE_PARAM {
	CString m_strVariable;
	double	m_lfFrom;
	double	m_lfTo;
	double	m_lfStep;
} ADJUSTABLE_PARAM;

typedef struct tag_COST_PARAM {
	double  m_lfLinearCost;				//Cost per unit length of the BMP structure ($/ft)
	double  m_lfAreaCost;               //Cost per unit area of the BMP structure ($/ft^2)
	double  m_lfTotalVolumeCost;        //Cost per unit total volume of the BMP structure ($/ft^3)
	double  m_lfMediaVolumeCost;        //Cost per unit volume of the soil media ($/ft^3)
	double  m_lfUnderDrainVolumeCost;   //Cost per unit volume of the under drain structure ($/ft^3)
	double  m_lfConstantCost;           //Constant cost ($)
	double  m_lfPercentCost;            //Cost in percentage of all other cost (%)
	double  m_lfLengthExp;
	double  m_lfAreaExp;
	double  m_lfTotalVolExp;
	double  m_lfMediaVolExp;
	double  m_lfUDVolExp;
} COST_PARAM;

typedef struct tag_EVALUATION_FACTOR {
	CString m_strFactor;	// Evaluation factor name, e.g. FlowVolume or TSSLoad
	int		m_nFactorGroup;	// -1 for flow, positive number for pollutant column order
	int		m_nFactorType;	// -1 -- AAFV, -2 -- PDF, -3 -- FEF, 1 -- AAL, 2 -- AAC, 3 -- MAC, 4 -- CEF
	int		m_nCalcMode;	// 1 -- %, 2 -- Scale, 3 -- Value
	int		m_nCalcDays;	// For factor type MAC only, Maxmimum #Days
	double  m_lfThreshold;	// For factor type FEF only, Flow Threshold
	double  m_lfConcThreshold;	// For factor type CEF only, Conc Threshold (optional)
	double	m_lfTarget;		// Target value of evaluation factor
	double	m_lfNextTarget;	// Target value of evaluation factor for next target break
	double	m_lfPriorFactor;// Priority factor for maximize control option
	double	m_lfLowerTarget;// Lower target value (for Option Trade-Off Curve only)
	double	m_lfUpperTarget;// Upper target value (for Option Trade-Off Curve only)
	double	m_lfInit;		// Initial value for evaluation factor for the first run
	double	m_lfCurrent;	// Current value for evaluation factor for the optimization run
	double	m_lfPreDev;		// PreDeveloped value for evaluation factor (without BMP)
	double	m_lfPostDev;	// PostDeveloped value for evaluation factor (without BMP)
} EVALUATION_FACTOR;

typedef struct tag_HOLTAN_PARAM {
	double	m_lfVegA;
	double	m_lfFInfilt;
	double	m_lfGrowth[12];		// growth index
} HOLTAN_PARAM;

typedef struct tag_BMP_A {
	double	 m_lfBasinWidth;
	double	 m_lfBasinLength;//bmp length (ft)
	double	 m_lfOrificeHeight;
	double	 m_lfOrificeDiameter;
	int		 m_nExitType;
	int		 m_nWeirType;
	double	 m_lfWeirHeight;
	double	 m_lfWeirWidth;
	double	 m_lfWeirAngle;
	int		 m_nORelease;		//   (07-14-04)
	int		 m_nPeople;			//   (07-14-04)
	int		 m_nDays;			//   (07-14-04)
	double	 m_lfRelease[24];	// hourly water release in a day	//   (07-14-04)
	TGrnAmpt m_pGAInfil;		// manage parameters for green ampt equation
} BMP_A;

typedef struct tag_BMP_B {
	double	 m_lfBasinWidth;
	double	 m_lfBasinLength;
	double	 m_lfMaximumDepth;
	double	 m_lfSideSlope1;
	double	 m_lfSideSlope2;
	double	 m_lfSideSlope3;
	double	 m_lfManning;
	TGrnAmpt m_pGAInfil;		// manage parameters for green ampt equation
} BMP_B;

typedef struct tag_Conduit {	// conduit input parameters (01-2005)
	int			m_nIndex;		// sequence number	
	//int			m_nID;			// BMP site ID
	CString		m_strID;		// BMP site ID
	CString		m_strCondType;	// Type of cross-section
	CString		m_strCondName;	// for irregular shape only
	TLink		m_pTLink;		// link parameters
	TConduit	m_pTConduit;	// conduit parameters
	TTransect	m_pTTransect;	// transect parameters
} BMP_C;

typedef struct tag_SEGMENT_D {
	int			m_nSegmentID;	
	double		m_lfSX;	
	double		m_lfRNA;	
	double		m_lfSOA;	
} SEGMENT_D;

typedef struct tag_POLLUTT_D {
	double		m_lfQUALSED_frac;	
	double		m_lfQUALDECAY_ads;	
	double		m_lfQUALDECAY_dis;
	double		m_lfTEMPCORR_ads;
	double		m_lfTEMPCORR_dis;
} POLLUTT_D;

typedef struct tag_BMP_D {
	CString		m_strName;			
	CString		m_strID;			
	int			m_nSegments;			
	int			m_nICO;
	int			m_nNPART[3];
	double		m_lfWidth;		//width of the buffer strip
	double		m_lfLength;
	double		m_lfVKS;
	double		m_lfSAV;
	double		m_lfOS;
	double		m_lfOI;
	double		m_lfSM;
	double		m_lfSCHK;
	double		m_lfSS;
	double		m_lfVN;
	double		m_lfH;
	double		m_lfVN2;
	double		m_lfCOARSE[3];
	double		m_lfPOR[3];
	double		m_lfDP[3];
	double		m_lfSG[3];
	SEGMENT_D*	m_pSEGMENT_D;		
	POLLUTT_D*  m_pPOLLUTANT_D;
} BMP_D;

typedef struct tag_BMP_GROUP {
	int			m_nGroupID;		// group ID for BMPs in a group
	double		m_lfTotalArea;	// limit of total area of all BMPs in this group
	CPtrList	m_bmpList;		// list of all BMPs belonging to a group
} BMP_GROUP;

typedef struct tag_SEDIMENT 
{
	double m_lfBEDWID;		// Bed width (ft)
	double m_lfBEDDEP;		// Initial bed depth (ft)
	double m_lfBEDPOR;		// Bed sediment porosity
	double m_lfSAND_FRAC;	// Bed sediment sand fraction
	double m_lfSILT_FRAC;	// Bed sediment silt fraction
	double m_lfCLAY_FRAC;	// Bed sediment clay fraction
} SEDIMENT;

typedef struct tag_SAND 
{
	double m_lfD;		// Effective diameter of the transported sand particles (in)
	double m_lfW;		// The corresponding fall velocity in still water (in/sec)
	double m_lfRHO;		// The density of the sand particles (lb/ft3)
	double m_lfKSAND;	// The coefficient in the sandload power function formula
	double m_lfEXPSND;	// The exponent in the sandload power function formula
} SAND;

typedef struct tag_SILTCLAY 
{
	double m_lfD;		// Effective diameter of the transported silt/clay particles (in)
	double m_lfW;		// The corresponding fall velocity in still water (in/sec)
	double m_lfRHO;		// The density of the silt/clay particles (lb/ft3)
	double m_lfTAUCD;	// The critical bed shear stress for deposition (lb/ft2)
	double m_lfTAUCS;	// The critical bed shear stress for scour (lb/ft2)
	double m_lfM;		// The erodibility coefficient of the sediment (lb/ft2/day)
} SILTCLAY;

class TradeOffCurve 
{
public:
	int m_nID;				// Break Point ID
	int m_nQualNum;
	long m_nTSNum;
	double m_lfMult;		// The multiplier to the timeseries file
	double m_lfCost;		// The cost associated for the timeseries file
	double m_lfSand;		// The fraction of total sediemnt which is sand
	double m_lfSilt;		// The fraction of total sediemnt which is silt
	double m_lfClay;		// The fraction of total sediemnt which is clay
	double* m_pDataBrPt;	// pointer to the point source timeseries data 
	CString m_strBrPtFile;	// The timeseries file name
	COleDateTime m_tmStart;

	TradeOffCurve()
	{
		m_nID  = 0;
		m_nQualNum = 0;
		m_nTSNum = 0;
		m_lfMult = 0.0;
		m_lfCost = 0.0;
		m_lfSand = 0.0;
		m_lfSilt = 0.0;
		m_lfClay = 0.0;
		m_strBrPtFile = "";
		m_tmStart = COleDateTime(1890,1,1,0,0,0);
		m_pDataBrPt = NULL;
	}

	virtual ~TradeOffCurve()
	{
		if(m_pDataBrPt != NULL)
			delete []m_pDataBrPt;
	}
};

class POLLUT_RAConc
{
public:
	int 	m_nRDays;				// maximum number of running days
	double*	m_lfRFlow;				// running flow
	double*	m_lfRLoad;				// running load

	//optional
	double  m_lfThreshConc;			// threshold conc. (Count/L for bacteria, lb/L for other pollutants)
	queue<double> qMass;			// holding 24 previous values for mass (lb)

	POLLUT_RAConc()
	{
		m_nRDays  = 0;
		m_lfRFlow = NULL;
		m_lfRLoad = NULL;
		m_lfThreshConc = 0.0;
		while (qMass.size() != 24)
			qMass.push(0.0);
	}

	virtual ~POLLUT_RAConc()
	{
		if(m_lfRFlow != NULL)
			delete []m_lfRFlow;
		if(m_lfRLoad != NULL)
			delete []m_lfRLoad;
		while (!qMass.empty())
			qMass.pop();
	}
};

class CLandUse;	
	
class CBMPSite : public CObject  
{
public:
	CBMPSite();
	CBMPSite(const CString& strID,const CString& strName,const CString& strType,int bmpClass);
	virtual ~CBMPSite();

public:
	CString	m_strID;			// BMP id
	CString	m_strName;			// BMP name
	CString	m_strType;			// BMP type
	bool    m_bChecked;			// used for cycle checking and routing list creation
	bool	m_bUndSwitch;		// switch to turn on the underdrain option (0-off, 1-on)
	int		m_nInfiltMethod;	// Infiltration Method (0-Holtan, 1-Green-Ampt)	
	int		m_nPolRotMethod;	// Pollutant Routing Method (1-Completely mixed, >1-number of CSTRs in series)	
	int		m_nPolRemMethod;	// Pollutant Removal Method (0-1st order decay, 1-kadlec and knight method )	
	int		m_nGAInfil_Index;	// Index for the Green-Ampt Infiltration array (positive number)
	int		m_nBMPClass;		// 1-class A; 2-class B; 3-class C; 4-class D; 100-Assessment Point
	double	m_lfBMPUnit;		// number of BMP units	
	double	m_lfDDarea;			// BMP design drainage area (acre)
	double	m_lfAccDArea;		// accumulative design drainage area for this site (acre)
	double	m_lfSoilDepth;		// BMP substrate soil depth (ft)
	double	m_lfPorosity;		// Soil porosity (fraction)
	double	m_lfFCapacity;		// Soil field capacity (fraction)
	double	m_lfWPoint;			// Soil wilting point (fraction)
	double	m_lfUndDepth;		// Depth of filter media for underdrain option (ft)
	double	m_lfUndVoid;		// Filter media voids for the underdrain option (fraction)
	double	m_lfUndInfilt;		// Soil infiltration rate underneath the filter media (in/hr)
	double	m_lfCost;			// cost ($)				

	double	m_lfSurfaceArea;	// BMP surface area (acre)				
	double	m_lfExcavatnVol;	// BMP excavation volume (acre-ft)				
	double	m_lfSurfStorVol;	// BMP surface storage volume (acre-ft)				
	double	m_lfSoilStorVol;	// BMP soil storage volume (acre-ft)				
	double	m_lfUdrnStorVol;	// BMP surface storage volume (acre-ft)				

	double  m_lfThreshFlow;		// user defined threshold flow (cfs)
	double	m_lfSiteDArea;		// drainage area for this site
	double*	m_pDecay;			// decay/loss rate for pollutants
	double*	m_pK;				// Constant rate for pollutants (ft/hr)
	double*	m_pCstar;			// Background concentration for pollutants (lb/ft3)
	double*	m_pConc;			// CSTR concentration for pollutants (lb/ft3)
	double*	m_pUndRemoval;		// removal rate for underdrain pollutants
	void*	m_pSiteProp;		// the actual pointer type could be BMP_A, BMP_B, ..., BMP_X
	CLandUse* m_preLU;			// pre-developed landuse, the actual pointer type is CLanduse
	POLLUT_RAConc* m_RAConc;	// structure of running average concentration
	FILE*    m_fileOut;			// used for outputing simulation result
	HOLTAN_PARAM m_holtanParam;	// manage parameters for holtan equation
	COST_PARAM   m_costParam;	// manage parameters for cost functions
	CPtrList m_dsbmpsiteList;	// element type in the list is DS_BMPSITE
	CPtrList m_adjustList;		// element type in the list is ADJUSTABLE_PARAM
	CPtrList m_factorList;		// element type in the list is EVALUATION_FACTOR
	CPtrList m_usbmpsiteList;	// list for all upstream BMP sites which flows into this BMP site directly, element type in the list is DS_BMPSITE
	CObList  m_siteluList;		// list for all site land use associated with this BMP site
	CObList  m_sitepsList;		// list for all site point sources associated with this BMP site
	int m_nBreakPoints;			// total number of break points on the cost effectiveness curve
	double m_lfBreakPtID;		// break point ID for the timeseries
	CString  m_strCostFile;		// tradeoff curve solution file
	TradeOffCurve* m_TradeOff;	// array of Break point timeseries files
	int m_nQualNum;				// total number of constituents in the time series including flow and water quality
	long m_nTSNum;				// total number of records for the simulation duration 
	double* m_pDataMixLU;		// pointer to the timeseries data for internal land simulation (mixed landuses)
	double* m_pDataPreLU;		// pointer to the timeseries data for internal land simulation (predeveloped landuses)
	double m_lfROpeakMixLU;		// peak land runoff for the simulation duration (mixed landuses) ft3/hr 
	double m_lfROpeakPreLU;		// peak land runoff for the simulation duration (predeveloped landuses) ft3/hr 
	COleDateTime m_tmStart;		// start date and time of the time series data 
	SEDIMENT m_sediment;		// general parameters for sediment transport
	SAND     m_sand;			// sand parameters for sediment transportt
	SILTCLAY m_silt;			// silt/clay parameters for sediment transport
	SILTCLAY m_clay;			// clay parameters for sediment transport
	queue<double> qFlow;		// holding 24 previous values for flow (cfs)

public:
	double* GetVariablePointer(CString strVarName);
	double GetBMPArea();
	bool ReadTradeOffCurveCosts();
	bool LoadTradeOffCurveData(int BPindex,COleDateTime startDate, COleDateTime endDate);
	bool UnLoadTradeOffCurveData(int nBrPtIndex);
	bool LoadWatershedTSData(CString BMPSiteID, COleDateTime startDate, COleDateTime endDate, 
		 CString strFileName, double* multiplier, int landfg);
};

typedef struct tag_DS_BMPSITE 
{
	int		  m_nOutletType;	// 1-total 2-weir 3-orifica or channel 4-underdrain
	CBMPSite* m_pDSBMPSite;		// pointer to a downstream BMP site
} DS_BMPSITE;

typedef struct tag_US_BMPSITE 
{
	int		  m_nOutletType;	// 1-total 2-weir 3-orifica or channel 4-underdrain
	CBMPSite* m_pUSBMPSite;		// pointer to a upstream BMP site
} US_BMPSITE;

#endif // !defined(AFX_BMPSITE_H__D3CC98A1_98BD_4E1A_BDD9_42084729FD84__INCLUDED_)
