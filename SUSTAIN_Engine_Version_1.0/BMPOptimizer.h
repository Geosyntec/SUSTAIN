// BMPOptimizer.h: interface for the CBMPOptimizer class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BMPOPTIMIZER_H__CEB24D45_C523_4016_A3AC_B5C7324376DF__INCLUDED_)
#define AFX_BMPOPTIMIZER_H__CEB24D45_C523_4016_A3AC_B5C7324376DF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBMPRunner;

class CBMPOptimizer  
{
public:
	CBMPOptimizer();
	CBMPOptimizer(CBMPRunner* pBMPRunner);
	virtual ~CBMPOptimizer();

public:
	int  nMaxIter;			// Number of global iterations
	int  nRunCounter;		// Counter of runs
	int  nMaxRun;			// Number of maximum runs for evaluation (calculated by dividing maximum run time by init run time)
	double m_lfPrevResult;	// Previous total cost (Min_Cost) or value (Max_Ctrl) for every PSize times run
	double m_lfPrevValue;	// Previous total value (Min_Cost) for every PSize times run
	double** m_pVariables;
	CBMPRunner* m_pBMPRunner;
	FILE* m_pAllSolutions;
//	FILE* m_pDebug; // For debugging
	SCATTER_SEARCH problem;
	int pull_index;

public:
	void InitProblem(int nVar, int b1, int b2, int pSize, bool localSearch);
	void InitRefSet();
	void ResetRefSet();
	void CombineRefSet();
	void Combine(double *x, double *y, double **offsprings, int number);
	void Combine_inc(double *x, double *y, double **offsprings, int number);
	void TryAddRefSet1(double *sol);
	void TryAddRefSet2(double *sol);
	void UpdateRefSet2();
	void TryAddEvaluation(double *sol, double current_value);

	double GenerateValue(int a);
	double Evaluate(double* sol);
	double Evaluate_MinCost(double* sol);
	double Evaluate_MaxCtrl(double* sol);
	double Evaluate_TradeOff(double* sol);
	void ImproveSolution(double* sol, double* value);
	bool IsNewSolution(double **solutions, int dim, double *sol);
	void GetOrderIndices(int* indices, double *pesos, int num, int tipo);
	double DistanceToRefSet1(double *sol);
	double DistanceToRefSet(double *sol);
	int amoeba(double **p, double *y, int ndim, double ftol, int *nfunk);
	double amotry(double **p, double *y, double *psum, int ndim, int ihi, int *nfunk, double fac);
	void PerformSearch();
	void OutputDebugFileHeader(FILE* fp);
	void OutputDebugInformation(FILE* fp);
	void OutputFileHeader(CString header, FILE* fp);
	void OutputFileHeaderForTradeOffCurve(FILE* fp);
	void OutputBestSolutions();
	void OutputBestSolutionsForTradeOffCurve(int breakNum, FILE* fp);
};

#endif // !defined(AFX_BMPOPTIMIZER_H__CEB24D45_C523_4016_A3AC_B5C7324376DF__INCLUDED_)
