// BMPRunner.h: interface for the CBMPRunner class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BMPRUNNER_H__B87012A4_1DFE_4E79_9DF3_D61E429AFFAD__INCLUDED_)
#define AFX_BMPRUNNER_H__B87012A4_1DFE_4E79_9DF3_D61E429AFFAD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// definition for running mode
#define RUN_INIT		0
#define RUN_OPTIMIZE	1
#define RUN_OUTPUT		2
#define RUN_PREDEV		3
#define RUN_POSTDEV		4		

// definition for outlet type
#define TOTAL			1
#define WEIR_			2
#define ORIFICE_CHANNEL	3
#define UNDERDRAIN		4

// definition for simulation option
#define OPTION_NO_OPTIMIZATION		0
#define OPTION_MIMIMIZE_COST		1
#define OPTION_TRADE_OFF_CURVE		2
#define OPTION_MAXIMIZE_CONTROL		3

// definition for factor type
#define AAFV			-1
#define PDF				-2
#define FEF				-3
#define AAL				1
#define AAC				2
#define MAC				3
#define CEF				4	//optional

// definition for calculation mode
#define CALC_PERCENT	1
#define CALC_SCALE		2
#define CALC_VALUE		3

// definition for sediment classes
#define SAND			1
#define SILT			2
#define CLAY			3
#define TSS				4

//definition for the unit conversions
#define POUND2GRAM		453.5924	//lb to gram
#define LBpCFT2MGpL		16018.46	//lb/ft3 to mg/l
#define CFS2CMS			0.0283		//ft3/s to m3/s
#define CF2CM			0.0283		//ft3 to m3
#define FpS2MpS			0.3048		//ft/s to m/s
#define FOOT2METER		0.3048		//ft to m
#define fThreshold		1.0e-7		//flow threshold (cfs)

class CBMPOptimizer;		
class CBMPData;
class CProgressWnd;			

typedef INT (CALLBACK* GETDLLVIEW)(LPCSTR, INT);

class CBMPRunner  
{
public:
	CBMPRunner();
	CBMPRunner(CBMPData* bmpData);
	virtual ~CBMPRunner();

public:
	int optcounter;
	int outcounter;
	double lInitRunTime;			// Init run time (in millisecond)
	long nMaxRun;
	CBMPData* pBMPData;
	CProgressWnd* pWndProgress;		
	COleDateTime time_i;			// time at the beginning of the simulation
	FILE *fp;

public:
	void advect(double imat,double svol,double sro,double evol,double ero,double delts,
		 double crrat,double& conc,double& romat);
	void bmp_a(int nInfiltMethod,int nGAindex, bool underdrain_on,int timestep,
		 int npeople,int ddays,int releasetype,int weirtype,int& counter,double oinflow,
		 double BMParea,double orifice_area,double orificeheight,double orificecoef,
		 double weirwidth,double weirheight,double weirangle,double cisternoutflow,
		 double soildepth,double soilporosity,double finalf,double vegparma,
		 double holtpar,double udfinalf,double udsoildepth,double udsoilporosity,
		 double FC,double WP,double ETrate,double& AET,double& perc,double& ovolume,
		 double& ostage,double& infilt,double& orifice,double& weir,double& osa,
		 double& ostorage,double& udout,double& seepage);
	void bmp_b(int nInfiltMethod,int nGAindex, bool underdrain_on,int timestep,
		 double oinflow,double BMPdepth,double BMPwidth,double BMPlength,double slope1,
		 double slope2,double slope3,double man_n,double soildepth,double soilporosity,
		 double finalf,double vegparma,double holtpar,double udfinalf,double udsoildepth,
		 double udsoilporosity,double FC,double WP,double ETrate,double& AET,
		 double& perc,double& ovolume,double& ostage,double& infilt,double& channel,
		 double& weir,double& osa,double& ostorage,double& udout,double& seepage);
	void UpdateXareaStageSarea(double nvolume,double vol_max,double s_area_max,
		 double BMPdepth,double BMPwidth,double BMPlength,double slope1,
		 double slope2,double& x_area,double& nstage,double& sur_area);
	void RunModel(int nRunMode);
	bool OpenOutputFiles(const CString& runID, int nRunOption, int nRunMode);
	bool CloseOutputFiles();
	void WriteFileHeader(int nRunOption, int nRunMode);
};

//SWMM5
float getHydRad(TXsect* xsect, float y);
void  findLinkQual2(int i,float tStep,double wAdded,double kDecay,double& c);

#endif // !defined(AFX_BMPRUNNER_H__B87012A4_1DFE_4E79_9DF3_D61E429AFFAD__INCLUDED_)
