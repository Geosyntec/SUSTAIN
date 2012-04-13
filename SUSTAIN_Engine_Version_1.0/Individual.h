// Individual.h: interface for the CIndividual class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_INDIVIDUAL_H__9D3833E9_748E_4E06_9520_955235EC87C8__INCLUDED_)
#define AFX_INDIVIDUAL_H__9D3833E9_748E_4E06_9520_955235EC87C8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBMPOptimizerGA;

class CIndividual  
{
public:
	CIndividual();
	CIndividual(GA_PROBLEM *pProb, CBMPOptimizerGA *pOptimizer);
	virtual ~CIndividual();

public:
	GA_PROBLEM *pProblem;
	CBMPOptimizerGA *pGAOptimizer;

    bool validSolution;
	int rank;
    double constr_violation;
    double *xreal;
    double *obj;
    double *BMPcost;
    double *constr;
    double crowd_dist;

public:
	void Allocate(GA_PROBLEM *pProb, CBMPOptimizerGA *pOptimizer);
	void Initialize();
//	void TestProblem(double *xreal, double *obj, double *constr);
	void Evaluate();
	void Mutate();
	void CopyFrom(const CIndividual& ind);
	int CheckDominance(const CIndividual& ind);
};

#endif // !defined(AFX_INDIVIDUAL_H__9D3833E9_748E_4E06_9520_955235EC87C8__INCLUDED_)
