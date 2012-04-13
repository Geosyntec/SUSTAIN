// Population.h: interface for the CPopulation class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_POPULATION_H__36F56525_4CD4_4C2E_8B6E_D1D1A3097808__INCLUDED_)
#define AFX_POPULATION_H__36F56525_4CD4_4C2E_8B6E_D1D1A3097808__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CIndividual;
class CBMPOptimizerGA;

class CPopulation  
{
public:
	CPopulation();
	CPopulation(int nsize, GA_PROBLEM *pProblem, CBMPOptimizerGA *pOptimizer);
	virtual ~CPopulation();

public:
	GA_PROBLEM *pProblem;
	CBMPOptimizerGA *pGAOptimizer;
	int nSize;
	CIndividual *individuals;

public:
	void Initialize();
	void Evaluate();
	void Mutate();
	void AssignRankAndCrowdingDistance();
	void AssignCrowdingDistanceList(void *list, int front_size);
	void AssignCrowdingDistance(int *dist, int **obj_array, int front_size);
	void AssignCrowdingDistanceIndices(int c1, int c2);
	void QuickSortFront(int objcount, int *obj_array, int obj_array_size);
	void QuickSortFrontImpl(int objcount, int *obj_array, int left, int right);
	void QuickSortDist(int *dist, int front_size);
	void QuickSortDistImpl(int *dist, int left, int right);
	void ReportIndividualToFile(FILE *fp, int index);
	void ReportAllToFile(FILE *fp);
	void ReportBestToFile(FILE *fp);
	int  GetBestSolutionIndex();
	int  GetNextBestSolutionIndex(double prevMinCost);
};

#endif // !defined(AFX_POPULATION_H__36F56525_4CD4_4C2E_8B6E_D1D1A3097808__INCLUDED_)
