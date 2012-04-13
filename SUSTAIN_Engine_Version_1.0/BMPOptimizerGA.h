// BMPOptimizerGA.h: interface for the BMPOptimizerGA class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BMPOPTIMIZERGA_H__EF832A11_19CE_44C3_B4C9_D2E5AF1D8D90__INCLUDED_)
#define AFX_BMPOPTIMIZERGA_H__EF832A11_19CE_44C3_B4C9_D2E5AF1D8D90__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CIndividual;
class CPopulation;
class CBMPRunner;

class CBMPOptimizerGA  
{
public:
	CBMPOptimizerGA();
	CBMPOptimizerGA(CBMPRunner* pBMPRunner);
	virtual ~CBMPOptimizerGA();

public:
	GA_PROBLEM problem;

    CPopulation *parent_pop;
    CPopulation *child_pop;
    CPopulation *mixed_pop;

	FILE *fpBestPop;

	int  nRunCounter;		// Counter of runs
	int  nMaxRun;			// Number of maximum runs for evaluation (calculated by dividing maximum run time by init run time)
	CBMPRunner* m_pBMPRunner;
	double** m_pVariables;
	FILE* m_pAllSolutions;

public:
	bool OpenOutputFiles();
	bool CloseOutputFiles();

	bool LoadData();
	bool ValidateParams();
	bool InitProblem();
	bool PerformSearch();
	void Selection(CPopulation *old_pop, CPopulation *new_pop);
	CIndividual* Tournament(CIndividual *ind1, CIndividual *ind2);
	void CrossOver(CIndividual *parent1, CIndividual *parent2, CIndividual *child1, CIndividual *child2);
	void CrossOverReal(CIndividual *parent1, CIndividual *parent2, CIndividual *child1, CIndividual *child2);
	void Merge(CPopulation *pop1, CPopulation *pop2, CPopulation *pop3);
	void FillNondominatedSort(CPopulation *mixed_pop, CPopulation *new_pop);
	void CrowdingFill(CPopulation *mixed_pop, CPopulation *new_pop, int count, int front_size, void *list);
	bool EvaluateSolution(double *xreal, double *obj, double *bmpcost, double *constr);
	bool Evaluate_MinCost(double *xreal, double *obj, double *bmpcost, double *constr);
	bool Evaluate_MaxCtrl(double *xreal, double *obj, double *bmpcost, double *constr);
	bool Evaluate_TradeOff(double *xreal, double *obj, double *bmpcost, double *constr);
	void OutputBestPopulation();
//	void OutputBestSolutions();
	void OutputFileHeader(CString header, FILE* fp);
};

#endif // !defined(AFX_BMPOPTIMIZERGA_H__EF832A11_19CE_44C3_B4C9_D2E5AF1D8D90__INCLUDED_)
