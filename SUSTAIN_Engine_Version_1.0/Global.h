// Global.h : main header file for the SUSTAIN DLL
//

#if !defined(AFX_GLOBAL_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_)
#define AFX_GLOBAL_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

# define INF 1.0e14
# define EPS 1.0e-14
# define E  2.71828182845905
# define pi 3.14159265358979

#define SUSTAIN_VERSION    "Version 1.0 - March 27, 2009"

#define BMPCALC_EXPORTS

#ifdef BMPCALC_EXPORTS
#define BMPCALC_API __declspec(dllexport)
#else
#define BMPCALC_API __declspec(dllimport)
#endif

typedef struct tag_SCATTER_SEARCH {
	int		n_var;
	int		b1;
	int		b2;
	int		PSize;
	bool	LS;		  // =1 LocalSearch ON, 0 OFF
	int		iter;
	double	digits;

	double	*high;
	double	*low;
	double	*inc;
	int		**ranges; // Diversification Generator

	double	**refSet1;// Solutions 
	double	*value1;  // Objective value
	int		*order1;  // Order of solutions
	int		*iter1;   // Number of iter of each sol.

	double	**refSet2;// Diversification Elements
	double	*value2;  // Dissim value
	int		*order2;  // Order of Maximum min-distance
	int		*iter2;

	int		last_combine;  //Number of iter of last solution combination
	int		new_elements;  //True if new element is added since last combine

	// added variables for TradeOff Curve option
	double	**evaSolutions;
//	double	**evaOutputs;
	double	*evaValues;
	int		*evaOrders;
} SCATTER_SEARCH;

typedef struct tag_GA_PROBLEM {
	int popsize;
	int ngen;
	int nobj;
	int ncon;
	int nBMPtype;

	int nreal;
	double pcross_real;
	double pmut_real;
	int nrealmut;
	int nrealcross;
	double *min_realvar;
	double *max_realvar;
	double *inc_realvar;
	double eta_c;
	double eta_m;
}
GA_PROBLEM;

// for testing start
void randomize(double seed);
void warmup_random(double seed);
void advance_random();
// for testing end

double random_perc();
int random_int(int low, int high);
double random_real(double low, double high);
double random_real_with_inc(double low, double high, double inc);
double pet_Hamon(double lat, double cts, double tavc, double day);

class CProgressWnd;
int LandSimulation(int landfg,char* strInputFilePath,CProgressWnd* pwndProgress);
int CallVFSMOD(LPCSTR strVfsProjFile);

extern "C" BOOL PASCAL EXPORT StartLandSimulation(char* strLandPreDevFilePath,
												  char* strLandPostDevFilePath);
extern "C" BOOL PASCAL EXPORT StartSimulation(char* strInputFilePath,char* strBestPopRun);

#endif // !defined(AFX_GLOBAL_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_)
