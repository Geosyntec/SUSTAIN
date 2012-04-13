// BMPData.h : interface of the CBMPData class
//
/////////////////////////////////////////////////////////////////////////////

#include <afxtempl.h>
#include <queue>
using namespace std;

#if !defined(AFX_BMPDATA_H__FFC3A527_E06A_4F89_83DA_449E3C851C65__INCLUDED_)
#define AFX_BMPDATA_H__FFC3A527_E06A_4F89_83DA_449E3C851C65__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRATEGY_SCATTER_SEARCH		1
#define STRATEGY_GENETIC_ALGORITHM	2

int FindObIndexFromList(CObList& list, CObject* ob);

class CBMPSite;
class CLandUse;

typedef struct tag_WEATHERDATA {
	COleDateTime tDATE;	// date for the weather file data
	bool	bWetInt;	// wet interval (wet day or with in 72 hrs after wet day)
	double	lfPrec;		// precip (in/timestep)
	double	lfDailyPrec;// Daily precip (in/day)
} WEATHERDATA;

typedef struct tag_POLLUTANT {
	int	m_nID;				//unique id
	CString	m_sName;		//name description
	int	m_nSedfg;			//sediment class (sand, silt, clay or total)
	int m_nSedQual;			//sediment associated qual flag (0=no, 1=yes)
	double	m_lfMult;		//time series multiplier
	double m_lfsand_qfr;	//pollutant fraction associated with sand
	double m_lfsilt_qfr;	//pollutant fraction associated with silt
	double m_lfclay_qfr;	//pollutant fraction associated with clay
} POLLUTANT;

typedef struct tag_BMPCOST {
	int	m_nBMPClass;
	CString	m_strBMPType;
	double	m_lfCost;
} BMPCOST;

class CBMPData
{
public:
	CBMPData();
	virtual ~CBMPData();

// Attributes
public:
	int nStrategy;			// Optimization strategy -- Scatter Search or Genetic Algorithm
	CString strInputDir;
	CString strOutputDir;
	CString strMixLUFileName;
	CString strPreLUFileName;
	CString strError;
	int nLandSimulation;	// Land Simulation Option (0-External,1-Internal)
	int nWeatherFile;
	int nBMPTimeStep;		// BMP Simulation Time Step (1 - 60 minutes)
	int nOutputTimeStep;	// Output Time Step -- 0 for daily, 1 for hourly
	int nBIORETENTION;		// total number of BIORETENTION in class A
	int nWETPOND;			// total number of WETPOND in class A
	int nCISTERN;			// total number of CISTERN in class A
	int nDRYPOND;			// total number of DRYPOND in class A
	int nINFILTRATIONTRENCH;// total number of INFILTRATIONTRENCH in class A
	int nGREENROOF;			// total number of GREENROOF in class A
	int nPOROUSPAVEMENT;	// total number of POROUSPAVEMENT in class A
	int nRAINBARREL;		// total number of RAINBARREL in class A
	int nREGULATOR;			// total number of REGULATOR in class A
	int nSWALE;				// total number of SWALE in class B
	int nBMPtype;			// total unique bmp types (A or B)
	int nBMPA;				// total number of BMPs in class A
	int nBMPB;				// total number of BMPs in class B
	int nBMPC;				// total number of BMPs in class C	- conduit (01-2005)
	int nBMPD;				// total number of BMPs in class D	- bufferstrip (06-2007)
	int nAdjVariable;		// total number of adjustable variables
	int nEvalFactor;		// total number of evaluation factors
	int nRunOption;			// 0 -- No Optimization, 1 -- Minimize Cost, 2 -- Maximize Control, 3 -- Generate Trade-off Curve
	int nSolution;			// Number of best solutions output for postprocessor
	int nTargetBreak;		// Number of break points between the lower and upper target values of trade-off curve 
	double lfCostLimit;		// cost limit for running option 2 (Maximize Control)
	double lfStopDelta;		// Stop criteria, $ for Minimize Cost, % for Maximize Control
	double lfMaxRunTime;	// Maximum run time (specified in the input file)
	COleDateTime startDate;
	COleDateTime endDate;

	int* nSedflag;			// array of sediment flag
	double* polmultiplier;	// array of pollutant multiplier
	POLLUTANT* m_pPollutant;
	BMPCOST* m_pBMPcost;	// array for unique bmp type in the project
	int nPollutant;			// number of pollutants in the time series
	int nNWQ;				// number of pollutants (TSS splits into sand, silt, and clay)

	int m_nNum;
	int nETflag;			// ET calculation flag (0-constant monthly ET,1-daily evaporation rate from the timeseries data,2-calculated from the daily temperature data)
	double lfLatitude;		// Latitude (Degree,Minute,Second)
	double lfmonET[12];		// Constant monthly ET rate (in/day) if ET flag is 0 otherwise monthly coefficient to calculate ET values
	double* m_pDataClimate;	// pointer to the timeseries data for climate file
	CString strClimateFileName;	// climate file path
	int m_nSedQualFlag;		// sediment-associated qual flag

	//bufferstrip parameters
	int nN, nMAXITER, nNPOL, nIELOUT, nKPG;
	double lfTHETAW, lfCR; 

	CObList luList;
	CObList siteluList;
	CObList bmpsiteList;
	CObList routeList;		// routing list
	CObList sitepsList;		// site point source list

	CPtrList bmpGroupList;	// list of bmp groups

	//optional weather data
	long lRecords,lStartIndex,lEndIndex;
	int nWetPeriod;
	int nWetInt;
	double lfWetDays;
	COleDateTime *pWetPeriod;
	WEATHERDATA	*pWEATHERDATA;	// array of original weather data

// Operations
public:
	CBMPSite* FindBMPSite(const CString& strID);
	CLandUse* FindLandUse(int nLuID);

	// functions for loading BMP data
	bool ReadInputFile(CString strFileName);
	bool ReadFileSection(FILE *fp, int nSection);
	void SkipCommentLine(FILE *fp);
	bool ReadDataLine(FILE *fp, CString& strData);
	bool ReadBestPopFile(int nBestPopId);
	void OutputFileHeaderForTradeOffCurve(FILE* fp);

	bool ReadWeatherFile(CString strFileName);
	long GetNumberOfRecords(FILE *fp);
	long FindDataIndex(COleDateTime tCurrent);
	bool MarkWetIntervals(COleDateTime tStart,COleDateTime tEnd);

	// functions for associate landuse with BMP site, creating routing network, and load time series data
	void ClearCheckedFlag();
	bool RoutingCycleExist(CBMPSite* pBMPSite);
	void AddRouteNode(CBMPSite* pBMPSite);
	bool PrepareDataForModel();

	// functions for outputing calculation results
	bool OpenOutputFiles(const CString& runID);
	void WriteFileHeader(FILE *fp, int NWQ);
	bool CloseOutputFiles();
	bool ProcessPollutantData();
	bool ProcessTransportData();

	bool LoadClimateTSData(COleDateTime startDate,COleDateTime endDate);
	// helper function for testing and debuging
	CString GetRoutingOrder();

	// functions for running VFSMOD for buffer strips
	bool WriteVFSMODFiles(int nRunMode,CString strID,CString& strVFSCall);
	double GetPeakFlowRate(int nRunMode,long nNBCROFF,CString strID);
	bool ReadVFSMODFiles(int nRunMode,CString strID);
	bool RunVFSMOD(int nRunMode);
};

//SWMM5
void InitializeGAInfil(int nNum);
void CopyGAInfil(TGrnAmpt* Source, TGrnAmpt* Target);
void InitializeLinkConduitTransect(int nNum);
void CopyLink(TLink* Source, TLink* Target, int nPollutant);
void CopyConduit(TConduit* Source, TConduit* Target);
void CopyTransect(TTransect* Source, TTransect* Target);
void Validate_Conduit(int nNum);
/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BMPDATA_H__FFC3A527_E06A_4F89_83DA_449E3C851C65__INCLUDED_)
