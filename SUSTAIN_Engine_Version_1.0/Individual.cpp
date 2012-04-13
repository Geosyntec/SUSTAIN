// Individual.cpp: implementation of the CIndividual class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "Individual.h"
#include "BMPOptimizerGA.h"
#include <math.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

extern double random_perc();
extern double random_real(double low, double high);
extern double random_real_with_inc(double low, double high, double inc);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CIndividual::CIndividual()
{
	pProblem = NULL;
	pGAOptimizer = NULL;

	validSolution = false;
    rank = 0;
    constr_violation = 0;
    xreal = NULL;
    obj = NULL;
	BMPcost = NULL;
    constr = NULL;
    crowd_dist = 0;
}

CIndividual::CIndividual(GA_PROBLEM *pProb, CBMPOptimizerGA *pOptimizer)
{
	Allocate(pProb, pOptimizer);
}

CIndividual::~CIndividual()
{
	if (pProblem == NULL)
		return;

	if (xreal != NULL)
		delete []xreal;

	if (obj != NULL)
		delete []obj;

	if (BMPcost != NULL)
		delete []BMPcost;

	if (constr != NULL)
		delete []constr;
}

void CIndividual::Allocate(GA_PROBLEM *pProb, CBMPOptimizerGA *pOptimizer)
{
	ASSERT(pProb != NULL);
	ASSERT(pOptimizer != NULL);

	validSolution = false;
    rank = 0;
    constr_violation = 0;
    crowd_dist = 0;

	pProblem = pProb;
	pGAOptimizer = pOptimizer;

    if (pProblem->nreal > 0)
        xreal = new double[pProblem->nreal];
	else
		xreal = NULL;

	if (pProblem->nobj > 0)
		obj = new double[pProblem->nobj];
	else
		obj = NULL;

	if (pProblem->nBMPtype > 0)
		BMPcost = new double[pProblem->nBMPtype];
	else
		BMPcost = NULL;

    if (pProblem->ncon > 0)
        constr = new double[pProblem->ncon];
	else
		constr = NULL;
}

void CIndividual::Initialize()
{
	ASSERT(pProblem != NULL);

    int i;
    for (i=0; i<pProblem->nreal; i++)
    {
        xreal[i] = random_real_with_inc(pProblem->min_realvar[i], pProblem->max_realvar[i], pProblem->inc_realvar[i]);
    }

    for (i=0; i<pProblem->nBMPtype; i++)
    {
        BMPcost[i] = 0.0;
    }
}

/*
// zdt6 problem for testing
void CIndividual::TestProblem(double *xreal, double *obj, double *constr)
{
    double f1 = 1.0 - exp(-4.0*xreal[0]) * pow(sin(4.0*pi*xreal[0]), 6.0);
    
	double g = 0.0;
    for (int i=1; i<10; i++)
    {
        g += xreal[i];
    }
	g = g/9.0;
    g = pow(g, 0.25);
    g = 1.0 + 9.0*g;
    
	double h = 1.0 - pow(f1/g, 2.0);
    double f2 = g*h;

    obj[0] = f1;
    obj[1] = f2;
}

// zdt4 problem for testing
void CIndividual::TestProblem(double *xreal, double *obj, double *constr)
{
    double f1, f2, g, h;
    int i;
    f1 = xreal[0];
    g = 0.0;
    for (i=1; i<10; i++)
    {
        g += xreal[i]*xreal[i] - 10.0*cos(4.0*pi*xreal[i]);
    }
    g += 91.0;
    h = 1.0 - sqrt(f1/g);
    f2 = g*h;
    obj[0] = f1;
    obj[1] = f2;
}
*/

void CIndividual::Evaluate()
{
	ASSERT(pProblem != NULL);
	ASSERT(pGAOptimizer != NULL);

//	TestProblem(xreal, obj, constr);
	validSolution = pGAOptimizer->EvaluateSolution(xreal, obj, BMPcost, constr);

	constr_violation = 0.0;
	for (int i=0; i<pProblem->ncon; i++)
	{
		if (constr[i] < 0.0)
			constr_violation += constr[i];
	}
}

void CIndividual::Mutate()
{
	ASSERT(pProblem != NULL);

	int i;

    if (pProblem->nreal <= 0)
		return;

	double rnd_perc, delta1, delta2, mut_pow, deltaq;
	double y, yl, yu, val, xy;

	for (i=0; i<pProblem->nreal; i++)
	{
		if (random_perc() <= pProblem->pmut_real)
		{
			y = xreal[i];
			yl = pProblem->min_realvar[i];
			yu = pProblem->max_realvar[i];
			delta1 = (y-yl)/(yu-yl);
			delta2 = (yu-y)/(yu-yl);

			mut_pow = 1.0/(pProblem->eta_m+1.0);
			rnd_perc = random_perc();

			if (rnd_perc <= 0.5)
			{
				xy = 1.0-delta1;
				val = 2.0*rnd_perc+(1.0-2.0*rnd_perc)*pow(xy, pProblem->eta_m+1.0);
				deltaq = pow(val, mut_pow) - 1.0;
			}
			else
			{
				xy = 1.0-delta2;
				val = 2.0*(1.0-rnd_perc)+2.0*(rnd_perc-0.5)*pow(xy, pProblem->eta_m+1.0);
				deltaq = 1.0 - pow(val, mut_pow);
			}
			y += deltaq*(yu-yl);
			
			y = yl+int((y-yl)/pProblem->inc_realvar[i]+0.5)*pProblem->inc_realvar[i];
			if (y < yl)
				y = yl;
			if (y > yu)
				y = yu;
			xreal[i] = y;
			pProblem->nrealmut++;
		}
	}
}

void CIndividual::CopyFrom(const CIndividual& ind)
{
    int i;

	validSolution = ind.validSolution;
    rank = ind.rank;
    constr_violation = ind.constr_violation;
    crowd_dist = ind.crowd_dist;

    for (i=0; i<pProblem->nreal; i++)
        xreal[i] = ind.xreal[i];

    for (i=0; i<pProblem->nobj; i++)
        obj[i] = ind.obj[i];

    for (i=0; i<pProblem->nBMPtype; i++)
        BMPcost[i] = ind.BMPcost[i];

    for (i=0; i<pProblem->ncon; i++)
        constr[i] = ind.constr[i];
}

/*
Routine for usual non-domination checking.
It will return the following values:
   1  if a dominates b
   -1 if b dominates a
   0  if both a and b are non-dominated
*/
int CIndividual::CheckDominance(const CIndividual& ind)
{
	ASSERT(pProblem != NULL);

    if (constr_violation < 0 && ind.constr_violation < 0)
    {
        if (constr_violation > ind.constr_violation)
            return 1;
        else if (constr_violation < ind.constr_violation)
            return -1;
        else
            return 0;
    }
    else
    {
        if (constr_violation < 0 && ind.constr_violation == 0)
        {
            return -1;
        }
        else
        {
            if (constr_violation == 0 && ind.constr_violation < 0)
			{
                return 1;
			}
            else
            {
				int flag1 = 0;
				int flag2 = 0;

                for (int i=0; i<pProblem->nobj; i++)
                {
                    if (obj[i] < ind.obj[i])
                        flag1 = 1;
                    else if (obj[i] > ind.obj[i])
						flag2 = 1;
                }

                if (flag1 == 1 && flag2 == 0)
                    return 1;
                else if (flag1 == 0 && flag2 == 1)
                    return -1;
                else
                    return 0;
            }
        }
    }
}

