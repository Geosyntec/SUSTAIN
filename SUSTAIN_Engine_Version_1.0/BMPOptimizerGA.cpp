// BMPOptimizerGA.cpp: implementation of the BMPOptimizerGA class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "Individual.h"
#include "Population.h"
#include "BMPSite.h"
#include "BMPData.h"
#include "BMPRunner.h"
#include "BMPOptimizerGA.h"
#include <math.h>
#include <float.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBMPOptimizerGA::CBMPOptimizerGA()
{
	problem.nreal = 0;
	problem.nobj = 0;
	problem.nBMPtype = 0;
	problem.popsize = 0;
	problem.pcross_real = 0.0;
	problem.pmut_real = 0.0;
	problem.eta_c = 0.0;
	problem.eta_m = 0.0;
	problem.ngen = 0;
	problem.nrealmut = 0;
	problem.nrealcross = 0;
	problem.min_realvar = NULL;
	problem.max_realvar = NULL;
	problem.inc_realvar = NULL;

	parent_pop = NULL;
	child_pop = NULL;
	mixed_pop = NULL;
	fpBestPop = NULL;

	m_pBMPRunner = NULL;
	m_pVariables = NULL;

	nRunCounter = 0;
	nMaxRun = 0;
	m_pAllSolutions = NULL;
}

CBMPOptimizerGA::CBMPOptimizerGA(CBMPRunner* pBMPRunner)
{
	problem.nreal = 0;
	problem.nobj = 0;
	problem.nBMPtype = 0;
	problem.popsize = 0;
	problem.pcross_real = 0.0;
	problem.pmut_real = 0.0;
	problem.eta_c = 0.0;
	problem.eta_m = 0.0;
	problem.ngen = 0;
	problem.nrealmut = 0;
	problem.nrealcross = 0;
	problem.min_realvar = NULL;
	problem.max_realvar = NULL;
	problem.inc_realvar = NULL;

	parent_pop = NULL;
	child_pop = NULL;
	mixed_pop = NULL;
	fpBestPop = NULL;

	m_pBMPRunner = pBMPRunner;
	m_pVariables = NULL;

	nRunCounter = 0;
	nMaxRun = 0;
	m_pAllSolutions = NULL;
}

CBMPOptimizerGA::~CBMPOptimizerGA()
{
	CloseOutputFiles();

	if (problem.min_realvar != NULL)
		delete []problem.min_realvar;
	if (problem.max_realvar != NULL)
		delete []problem.max_realvar;
	if (problem.inc_realvar != NULL)
		delete []problem.inc_realvar;

	if (parent_pop != NULL)
		delete parent_pop;
	if (child_pop != NULL)
		delete child_pop;
	if (mixed_pop != NULL)
		delete mixed_pop;

	if (m_pVariables != NULL)
		delete []m_pVariables;
}

bool CBMPOptimizerGA::OpenOutputFiles()
{
	if (m_pBMPRunner == NULL)
		return false;

	int i;
	CString strFilePath;

	// open file for all solutions
	strFilePath = m_pBMPRunner->pBMPData->strOutputDir + "AllSolutions.out";
	m_pAllSolutions = fopen(LPCSTR(strFilePath), "wt");
	if (m_pAllSolutions == NULL)
		goto FILE_OPEN_ERROR;
	else
	    //fprintf(m_pAllSolutions, "All solutions\n");
		OutputFileHeader("NSGA-II - All solutions",m_pAllSolutions);

	// open file for best population
	strFilePath = m_pBMPRunner->pBMPData->strOutputDir + "BestSolutions.out";
	fpBestPop = fopen(LPCSTR(strFilePath), "wt");
	if (fpBestPop == NULL)
		goto FILE_OPEN_ERROR;
	else
		OutputFileHeader("NSGA-II - Best population",fpBestPop);

	return true;

FILE_OPEN_ERROR:
	m_pBMPRunner->pBMPData->strError = "Cannot open file " + strFilePath + " for writing.";
	CloseOutputFiles();
	return false;
}

bool CBMPOptimizerGA::CloseOutputFiles()
{
	if (m_pAllSolutions != NULL)
	{
		fclose(m_pAllSolutions);
		m_pAllSolutions = NULL;
	}
	if (fpBestPop != NULL)
	{
		fclose(fpBestPop);
		fpBestPop = NULL;
	}

	return true;
}

bool CBMPOptimizerGA::LoadData()
{
	if (m_pBMPRunner == NULL)
		return false;

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;
	if (pBMPData == NULL)
		return false;

	// Simulation parameters
	//problem.popsize = 100;
	problem.popsize = 4 * pBMPData->nAdjVariable;
	problem.ngen = 20;

	double lfdivider = m_pBMPRunner->lInitRunTime * problem.popsize;
	if (lfdivider > 0)
		problem.ngen = int((pBMPData->lfMaxRunTime*3600000)/lfdivider) + 1;
	problem.nobj = 1; // cost is the first objective by default

	problem.nBMPtype = pBMPData->nBMPtype;//unique bmp types

	// If using real variables
	problem.nreal = pBMPData->nAdjVariable;
	problem.pcross_real = 1.0;
	problem.pmut_real = 0.0333;
	problem.nrealmut = 0;
	problem.nrealcross = 0;
	problem.eta_c = 15;
	problem.eta_m = 20;
	problem.min_realvar = new double[problem.nreal];
	problem.max_realvar = new double[problem.nreal];
	problem.inc_realvar = new double[problem.nreal];

	int nIndex = 0;
	ADJUSTABLE_PARAM* pAP;
	CBMPSite* pBMPSite;
	POSITION pos, pos1;

	m_pVariables = new double*[problem.nreal];

	pos = pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);
		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			problem.min_realvar[nIndex] = pAP->m_lfFrom;
			problem.max_realvar[nIndex] = pAP->m_lfTo;
			problem.inc_realvar[nIndex] = pAP->m_lfStep;
			m_pVariables[nIndex] = pBMPSite->GetVariablePointer(pAP->m_strVariable);
			nIndex++;
		}


		problem.nobj += pBMPSite->m_factorList.GetCount();
	}

	problem.ncon = problem.nobj;

/*
	// for testing start
	// testing problem zdt6 from nsga2r
	problem.popsize = 100;
	problem.ngen = 200;
	problem.nobj = 2;
	problem.ncon = 0;

	problem.nreal = 10;
	if (problem.nreal > 0)
	{
		problem.min_realvar = new double[problem.nreal];
		problem.max_realvar = new double[problem.nreal];
		problem.inc_realvar = NULL;
		for (int i=0; i<problem.nreal; i++)
		{
			problem.min_realvar[i] = 0;
			problem.max_realvar[i] = 1;
		}
		problem.pcross_real = 0.9;
		problem.pmut_real = 0.1;
		problem.eta_c = 15;
		problem.eta_m = 20;
	}
	else
	{
		problem.min_realvar = NULL;
		problem.max_realvar = NULL;
		problem.inc_realvar = NULL;
	}

	// testing problem zdt4 from nsga2r
	problem.popsize = 100;
	problem.ngen = 200;
	problem.nobj = 2;
	problem.ncon = 0;

	problem.nreal = 10;
	if (problem.nreal > 0)
	{
		problem.min_realvar = new double[problem.nreal];
		problem.max_realvar = new double[problem.nreal];
		problem.inc_realvar = NULL;
		problem.min_realvar[0] = 0;
		problem.max_realvar[0] = 1;
		for (int i=1; i<problem.nreal; i++)
		{
			problem.min_realvar[i] = -5;
			problem.max_realvar[i] = 5;
		}
		problem.pcross_real = 0.9;
		problem.pmut_real = 0.1;
		problem.eta_c = 15;
		problem.eta_m = 20;
	}
	else
	{
		problem.min_realvar = NULL;
		problem.max_realvar = NULL;
		problem.inc_realvar = NULL;
	}
*/
	// for testing end

	return true;
}

bool CBMPOptimizerGA::ValidateParams()
{
	if (m_pBMPRunner == NULL)
		return false;

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	// validate population size, must be a positive multiplier of 4
    if (problem.popsize < 4 || (problem.popsize%4) != 0)
	{
		pBMPData->strError.Format("Wrong population size %d.\n", problem.popsize);
		return false;
	}

	// validate generation number, must be at least 1 generation
    if (problem.ngen < 1)
	{
		pBMPData->strError.Format("Wrong number of generations %d.\n", problem.ngen);
		return false;
	}

	// validate objectives number, must be at least 1 objective
	if (problem.nobj < 1)
	{
		pBMPData->strError.Format("Wrong number of objectives %d.\n", problem.nobj);
		return false;
	}

	// validate constraints number, must be a non-negative integer
	if (problem.ncon < 0)
	{
		pBMPData->strError.Format("Wrong number of constraints %d.\n", problem.ncon);
		return false;
	}

	// validate real variables number, must be a non-negative integer
	if (problem.nreal < 0)
	{
		pBMPData->strError.Format("Wrong number of real variables %d.\n", problem.nreal);
		return false;
	}
    else if (problem.nreal > 0)
    {
        for (int i=0; i<problem.nreal; i++)
        {
            if (problem.max_realvar[i] < problem.min_realvar[i])
            {
				pBMPData->strError.Format("Wrong limits entered for the min and max bounds of real variable %d: min - %lf max - %lf.\n", i+1, problem.min_realvar[i], problem.max_realvar[i]);
				return false;
            }
        }
        if (problem.pcross_real < 0.0 || problem.pcross_real > 1.0)
        {
			pBMPData->strError.Format("Entered value of probability of crossover of real variables is out of bounds: %lf.\n", problem.pcross_real);
			return false;
        }
        if (problem.pmut_real < 0.0 || problem.pmut_real > 1.0)
        {
			pBMPData->strError.Format("Entered value of probability of mutation of real variables is out of bounds: %lf.\n", problem.pmut_real);
			return false;
        }
        if (problem.eta_c <= 0)
        {
			pBMPData->strError.Format("Wrong value of distribution index for crossover entered: %lf.\n", problem.eta_c);
			return false;
        }
        if (problem.eta_m <= 0)
        {
			pBMPData->strError.Format("Wrong value of distribution index for mutation entered: %lf.\n", problem.eta_m);
			return false;
        }
    }

    if (problem.nreal == 0)
    {
		pBMPData->strError.Format("Both number of real variables and number of binary variables are zero.\n");
		return false;
    }

	return true;
}

bool CBMPOptimizerGA::InitProblem()
{
	// output population file headers for bookkeeping
//    fprintf(fpBestPop, "# of objectives = %d, # of constraints = %d, # of real_var = %d, constr_violation, rank, crowding_distance\n", problem.nobj, problem.ncon, problem.nreal);

//    fprintf(fpInitialPop, "# of objectives = %d, # of constraints = %d, # of real_var = %d, constr_violation, rank, crowding_distance\n", problem.nobj, problem.ncon, problem.nreal);
//    fprintf(fpFinalPop, "# of objectives = %d, # of constraints = %d, # of real_var = %d, constr_violation, rank, crowding_distance\n", problem.nobj, problem.ncon, problem.nreal);
//    fprintf(fpAllPop, "# of objectives = %d, # of constraints = %d, # of real_var = %d, constr_violation, rank, crowding_distance\n", problem.nobj, problem.ncon, problem.nreal);
//    fprintf(fpDebug, "# of objectives = %d, # of constraints = %d, # of real_var = %d, constr_violation, rank, crowding_distance\n", problem.nobj, problem.ncon, problem.nreal);

	// output parameters for monitoring
//    fprintf(fpParams, "Population size = %d\n", problem.popsize);
//    fprintf(fpParams, "Number of generations = %d\n", problem.ngen);
//    fprintf(fpParams, "Number of objective functions = %d\n", problem.nobj);
//    fprintf(fpParams, "Number of constraints = %d\n", problem.ncon);
//	  fprintf(fpParams, "Number of real variables = %d\n", problem.nreal);
//    if (problem.nreal != 0)
//    {
//        for (int i=0; i<problem.nreal; i++)
//        {
//            fprintf(fpParams, "Lower limit of real variable %d = %e\n", i+1, problem.min_realvar[i]);
//            fprintf(fpParams, "Upper limit of real variable %d = %e\n", i+1, problem.max_realvar[i]);
//        }
//        fprintf(fpParams, "Probability of crossover of real variable = %e\n", problem.pcross_real);
//        fprintf(fpParams, "Probability of mutation of real variable = %e\n", problem.pmut_real);
//        fprintf(fpParams, "Distribution index for crossover = %e\n", problem.eta_c);
//        fprintf(fpParams, "Distribution index for mutation = %e\n", problem.eta_m);
//    }

	// initialize GA-related variables
    problem.nrealmut = 0;
    problem.nrealcross = 0;

	parent_pop = new CPopulation(problem.popsize, &problem, this);
	child_pop = new CPopulation(problem.popsize, &problem, this);
	mixed_pop = new CPopulation(2*problem.popsize, &problem, this);

	parent_pop->Initialize();
    
    fflush(fpBestPop);
//	  fflush(fpInitialPop);
//    fflush(fpFinalPop);
//    fflush(fpAllPop);
//    fflush(fpParams);
//    fflush(fpDebug);

	// Initialization done. Now start performing first generation ...
	return true;
}

bool CBMPOptimizerGA::PerformSearch()
{
	int bestSolIndex;
	double prevValue, curValue;
	ASSERT(parent_pop != NULL);
	ASSERT(child_pop != NULL);
	ASSERT(mixed_pop != NULL);

	parent_pop->Evaluate();
	parent_pop->AssignRankAndCrowdingDistance();
//	parent_pop->ReportAllToFile(fpInitialPop);
//	fprintf(fpAllPop, "# gen = 1\n");
//    parent_pop->ReportAllToFile(fpAllPop);
//    fflush(fpInitialPop);
//    fflush(fpAllPop);

	bestSolIndex = parent_pop->GetBestSolutionIndex();
	if (bestSolIndex != -1)
		prevValue = parent_pop->individuals[bestSolIndex].obj[0];
	else
		prevValue = DBL_MAX;

    for (int i=2; i<=problem.ngen; i++)
    {
		Selection(parent_pop, child_pop);
		child_pop->Mutate();
        child_pop->Evaluate();

		Merge(parent_pop, child_pop, mixed_pop);
		FillNondominatedSort(mixed_pop, parent_pop);

        // Comment following three lines if information for all
        // generations is not desired. It will speed up the execution
//        fprintf(fpAllPop, "# gen = %d\n", i);
//        parent_pop->ReportAllToFile(fpAllPop);
//        fflush(fpAllPop);

		bestSolIndex = parent_pop->GetBestSolutionIndex();
		if (bestSolIndex != -1)
			curValue = parent_pop->individuals[bestSolIndex].obj[0];
		else
			curValue = DBL_MAX;

		if (prevValue != DBL_MAX && curValue != DBL_MAX) {
//			double curDelta = prevValue-curValue;
			double curDelta = fabs(prevValue-curValue);
			if (curDelta < m_pBMPRunner->pBMPData->lfStopDelta)
			{
				CString strMsg;
				strMsg.Format("Cost of the best solution has been reduced by $%.1lf. The cost reduction is within the stopping delta range. Do you want to continue?", curDelta);
				if (AfxMessageBox(strMsg, MB_YESNO|MB_ICONINFORMATION) != IDYES)
					return true;
			}
		}

		prevValue = curValue;
    }

	return true;
}

/* Routine for tournament selection, it creates a new_pop from old_pop by performing tournament selection and the crossover */
void CBMPOptimizerGA::Selection(CPopulation *old_pop, CPopulation *new_pop)
{
    int i;
    int *a1 = new int[problem.popsize];
    int *a2 = new int[problem.popsize];

    for (i=0; i<problem.popsize; i++)
    {
        a1[i] = i;
		a2[i] = i;
    }

    int temp, rnd_index;
    for (i=0; i<problem.popsize; i++)
    {
        rnd_index = random_int(i, problem.popsize-1);
        temp = a1[rnd_index];
        a1[rnd_index] = a1[i];
        a1[i] = temp;
        
		rnd_index = random_int(i, problem.popsize-1);
        temp = a2[rnd_index];
        a2[rnd_index] = a2[i];
        a2[i] = temp;
    }

    CIndividual *parent1, *parent2;
    for (i=0; i<problem.popsize; i+=4)
    {
        parent1 = Tournament(&old_pop->individuals[a1[i]], &old_pop->individuals[a1[i+1]]);
        parent2 = Tournament(&old_pop->individuals[a1[i+2]], &old_pop->individuals[a1[i+3]]);
        CrossOver(parent1, parent2, &new_pop->individuals[i], &new_pop->individuals[i+1]);

        parent1 = Tournament(&old_pop->individuals[a2[i]], &old_pop->individuals[a2[i+1]]);
        parent2 = Tournament(&old_pop->individuals[a2[i+2]], &old_pop->individuals[a2[i+3]]);
        CrossOver(parent1, parent2, &new_pop->individuals[i+2], &new_pop->individuals[i+3]);
    }
    
	delete []a1;
    delete []a2;
}

/* Routine for binary tournament */
CIndividual* CBMPOptimizerGA::Tournament(CIndividual *ind1, CIndividual *ind2)
{
    int flag = ind1->CheckDominance(*ind2);

    if (flag == 1)
        return ind1;
	
	if (flag == -1)
        return ind2;
    
	if (ind1->crowd_dist > ind2->crowd_dist)
        return ind1;

    if (ind2->crowd_dist > ind1->crowd_dist)
        return ind2;

    if (random_perc() <= 0.5)
        return ind1;
    else
        return ind2;
}

/* Function to cross two individuals */
void CBMPOptimizerGA::CrossOver(CIndividual *parent1, CIndividual *parent2, CIndividual *child1, CIndividual *child2)
{
    if (problem.nreal > 0)
        CrossOverReal(parent1, parent2, child1, child2);
}

/* Routine for real variable SBX crossover */
void CBMPOptimizerGA::CrossOverReal(CIndividual *parent1, CIndividual *parent2, CIndividual *child1, CIndividual *child2)
{
    int i;
    double rnd_perc;
    double y1, y2, yl, yu;
    double c1, c2;
    double alpha, beta, betaq;

    if (random_perc() <= problem.pcross_real)
    {
        problem.nrealcross++;
        for (i=0; i<problem.nreal; i++)
        {
            if (random_perc() <= 0.5)
            {
                if (fabs(parent1->xreal[i]-parent2->xreal[i]) > EPS)
                {
                    if (parent1->xreal[i] < parent2->xreal[i])
                    {
                        y1 = parent1->xreal[i];
                        y2 = parent2->xreal[i];
                    }
                    else
                    {
                        y1 = parent2->xreal[i];
                        y2 = parent1->xreal[i];
                    }

                    yl = problem.min_realvar[i];
                    yu = problem.max_realvar[i];

                    rnd_perc = random_perc();

                    beta = 1.0 + (2.0*(y1-yl)/(y2-y1));
                    alpha = 2.0 - pow(beta, -(problem.eta_c+1.0));
                    if (rnd_perc <= 1.0/alpha)
                        betaq = pow(rnd_perc*alpha, 1.0/(problem.eta_c+1.0));
                    else
                        betaq = pow(1.0/(2.0-rnd_perc*alpha), 1.0/(problem.eta_c+1.0));
                    c1 = (y1+y2-(betaq*(y2-y1)))/2.0;
					c1 = yl+int((c1-yl)/problem.inc_realvar[i]+0.5)*problem.inc_realvar[i];
                    if (c1 < yl)
                        c1 = yl;
                    if (c1 > yu)
                        c1 = yu;

                    beta = 1.0 + (2.0*(yu-y2)/(y2-y1));
                    alpha = 2.0 - pow(beta, -(problem.eta_c+1.0));
                    if (rnd_perc <= 1.0/alpha)
                        betaq = pow(rnd_perc*alpha, 1.0/(problem.eta_c+1.0));
                    else
                        betaq = pow(1.0/(2.0-rnd_perc*alpha), 1.0/(problem.eta_c+1.0));
                    c2 = 0.5*(y1+y2+betaq*(y2-y1));
					c2 = yl+int((c2-yl)/problem.inc_realvar[i]+0.5)*problem.inc_realvar[i];
                    if (c2 < yl)
                        c2 = yl;
                    if (c2 > yu)
                        c2 = yu;

                    if (random_perc() <= 0.5)
                    {
                        child1->xreal[i] = c2;
                        child2->xreal[i] = c1;
                    }
                    else
                    {
                        child1->xreal[i] = c1;
                        child2->xreal[i] = c2;
                    }
                }
                else
                {
                    child1->xreal[i] = parent1->xreal[i];
                    child2->xreal[i] = parent2->xreal[i];
                }
            }
            else
            {
                child1->xreal[i] = parent1->xreal[i];
                child2->xreal[i] = parent2->xreal[i];
            }
        }
    }
    else
    {
        for (i=0; i<problem.nreal; i++)
        {
            child1->xreal[i] = parent1->xreal[i];
            child2->xreal[i] = parent2->xreal[i];
        }
    }
}

/* Routine to merge two populations into one */
void CBMPOptimizerGA::Merge(CPopulation *pop1, CPopulation *pop2, CPopulation *pop3)
{
	int i, j;

    for (i=0; i<problem.popsize; i++)
		pop3->individuals[i].CopyFrom(pop1->individuals[i]);

    for (i=0, j=problem.popsize; i<problem.popsize; i++, j++)
		pop3->individuals[j].CopyFrom(pop2->individuals[i]);
}

/* Routine to perform non-dominated sorting */
void CBMPOptimizerGA::FillNondominatedSort(CPopulation *mixed_pop, CPopulation *new_pop)
{
	int i, j;

	CList<int, int> pool;
	CList<int, int> elite;
	POSITION pos1, pos2, tmp_pos;

	int archive_size = 0;
	int front_size = 0;
	int rank = 1;
	int pool_index, elite_index;

	for (i=0; i<2*problem.popsize; i++)
		pool.AddTail(i);

    i=0;
	while (archive_size < problem.popsize)
	{
		pool_index = pool.GetHead();
		pool.RemoveHead();
		pos1 = pool.GetHeadPosition();

		elite.AddHead(pool_index);
        front_size = 1;

        while (pos1 != NULL)
        {
			int flag = -1;
			pool_index = pool.GetAt(pos1);
			pos2 = elite.GetHeadPosition();

            while (pos2 != NULL)
            {
				elite_index = elite.GetAt(pos2);
                flag = mixed_pop->individuals[pool_index].CheckDominance(mixed_pop->individuals[elite_index]);

                if (flag == 1)
                {
					pool.AddHead(elite_index);
                    front_size--;
					tmp_pos = pos2;
					elite.GetNext(pos2);
					elite.RemoveAt(tmp_pos);
                }
                else if (flag == 0)
                {
					elite.GetNext(pos2);
                }
				else if (flag == -1)
				{
					break;
				}
            }
            
			if (flag != -1)
            {
				elite.AddHead(pool_index);
                front_size++;

				tmp_pos = pos1;
				pool.GetNext(pos1);
				pool.RemoveAt(tmp_pos);
            }
			else
			{
				pool.GetNext(pos1);
			}
        }
        
        if ((archive_size+front_size) <= problem.popsize)
        {
	        j = i;
            
			pos2 = elite.GetHeadPosition();
			while (pos2 != NULL)
            {
				elite_index = elite.GetNext(pos2);
                new_pop->individuals[i].CopyFrom(mixed_pop->individuals[elite_index]);
                new_pop->individuals[i].rank = rank;
                archive_size++;
                i++;
            }

            new_pop->AssignCrowdingDistanceIndices(j, i-1);
            rank++;
        }
        else
        {
            CrowdingFill(mixed_pop, new_pop, i, front_size, (void*) &elite);
            archive_size = problem.popsize;
            for (j=i; j<problem.popsize; j++)
                new_pop->individuals[j].rank = rank;
        }
        
		elite.RemoveAll();
    }

	pool.RemoveAll();
}

/* Routine to fill a population with individuals in the decreasing order of crowding distance */
void CBMPOptimizerGA::CrowdingFill(CPopulation *mixed_pop, CPopulation *new_pop, int count, int front_size, void *list)
{
	ASSERT(mixed_pop != NULL);
	ASSERT(new_pop != NULL);
	ASSERT(list != NULL);

    mixed_pop->AssignCrowdingDistanceList(list, front_size);

    CList<int, int> *p_list = (CList<int, int> *) list;
    int *dist = new int[front_size];

    int i, j;

    POSITION pos = p_list->GetHeadPosition();
    for (i=0; i<front_size; i++)
        dist[i] = p_list->GetNext(pos);

	mixed_pop->QuickSortDist(dist, front_size);
    for (i=count, j=front_size-1; i<problem.popsize; i++, j--)
        new_pop->individuals[i].CopyFrom(mixed_pop->individuals[dist[j]]);
    
	delete []dist;
}

bool CBMPOptimizerGA::EvaluateSolution(double *xreal, double *obj, double *bmpcost, double *constr)
{
	bool bValid = false;
	nRunCounter++;

	switch (m_pBMPRunner->pBMPData->nRunOption)
	{
		case OPTION_MIMIMIZE_COST:
			bValid = Evaluate_MinCost(xreal, obj, bmpcost, constr);
			break;
		case OPTION_MAXIMIZE_CONTROL:
			bValid = Evaluate_MaxCtrl(xreal, obj, bmpcost, constr);
			break;
		case OPTION_TRADE_OFF_CURVE:
			bValid = Evaluate_TradeOff(xreal, obj, bmpcost, constr);
			break;
	}

	return bValid;
}

bool CBMPOptimizerGA::Evaluate_MinCost(double *xreal, double *obj, double *bmpcost, double *constr)
{
	int i;
	bool bValid = true;
	double output;
	double totalCost = 0.0;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;

	// mapping solution to ajustable variables
	for(i=0; i<problem.nreal; i++)
		*m_pVariables[i] = xreal[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
	int nIndex = 0;
	POSITION pos, pos1;
	pos = pBMPData->routeList.GetHeadPosition();
	double delta;

	strEF = "";
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

		totalCost += pBMPSite->m_lfCost;
		totalSurfaceArea += pBMPSite->m_lfSurfaceArea;
		totalExcavatnVol += pBMPSite->m_lfExcavatnVol;
		totalSurfStorVol += pBMPSite->m_lfSurfStorVol;
		totalSoilStorVol += pBMPSite->m_lfSoilStorVol;
		totalUdrnStorVol += pBMPSite->m_lfUdrnStorVol;

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			nIndex++;
			obj[nIndex] = DBL_MAX;

			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);

			if (pEF->m_lfTarget < 0 )
			{
				AfxMessageBox("Target value can not be negative under minimize cost option");
				continue;
			}

			double output1 = 0.0;

			if (pEF->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
			{
				delta = 25/pEF->m_lfTarget;
				if (pEF->m_lfPostDev > 0)
					output1 = pEF->m_lfCurrent/pEF->m_lfPostDev*100;
//				else
//					output1 = pEF->m_lfCurrent;
			}
			else if (pEF->m_nCalcMode == CALC_VALUE) // if the calculation mode is value
			{
				delta = 0.25;
				output1 = pEF->m_lfCurrent;
			}
			else // if the calculation mode is scale
			{
				delta = 0.25/pEF->m_lfTarget;
				if (pEF->m_lfPostDev - pEF->m_lfPreDev > 0)
					output1 = (pEF->m_lfCurrent - pEF->m_lfPreDev) / (pEF->m_lfPostDev-pEF->m_lfPreDev);
//				else
//					output1 = pEF->m_lfCurrent;
			}

			strValue.Format("\t%lf", output1);
			strEF += strValue;

			if (pEF->m_lfTarget == 0.0)
			{
				obj[nIndex] = output1;

				if (output1 > 0) 
				{
					constr[nIndex] = -output1;
					bValid = false;
				}
				else
					constr[nIndex] = 0;
			}
			else
			{
				output = output1/pEF->m_lfTarget; // normalize Evaluation Factor using Target
			
				obj[nIndex] = output;

				if (output > 1+delta)
					constr[nIndex] = 1-output;
				else if (output < 1-delta)
					constr[nIndex] = output-1;
				else
					constr[nIndex] = 0;

				if (output > 1)
					bValid = false;
			}
		}
	}

	constr[0] = 0;
	obj[0] = totalCost;
	
	//strLine.Format("%d\t%lf\t%s", nRunCounter, totalCost, strEF);
	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		bmpcost[i] = pBMPData->m_pBMPcost[i].m_lfCost;
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.nreal; i++)
	{
		strValue.Format("\t%lf", xreal[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);

	return bValid;
}

bool CBMPOptimizerGA::Evaluate_MaxCtrl(double *xreal, double *obj, double *bmpcost, double *constr)
{
	bool bValid = true;
	if (m_pBMPRunner->pBMPData->lfCostLimit <= 0.0)
	{
		AfxMessageBox("Cost limit value can not be negative or zero under maximize control option");
		return false;
	}

	int i;
	double output = 0.0;
	double value = 0.0; // value = f(BMPOutputForEF, target, cost)
	double totalCost = 0.0, costWeight = 50.0;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;

	// mapping solution to ajustable variables
	for(i=0; i<problem.nreal; i++)
		*m_pVariables[i] = xreal[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
	int nIndex = 0;
	POSITION pos, pos1;
	pos = pBMPData->routeList.GetHeadPosition();

	strEF = "";
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

		totalCost += pBMPSite->m_lfCost;
		totalSurfaceArea += pBMPSite->m_lfSurfaceArea;
		totalExcavatnVol += pBMPSite->m_lfExcavatnVol;
		totalSurfStorVol += pBMPSite->m_lfSurfStorVol;
		totalSoilStorVol += pBMPSite->m_lfSoilStorVol;
		totalUdrnStorVol += pBMPSite->m_lfUdrnStorVol;

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			nIndex++;
			obj[nIndex] = DBL_MAX;
			constr[nIndex] = 0;

			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);

			double output1 = 0.0;

			if (pEF->m_lfPostDev > 0)
				//if (pEF->m_nCalcMode == CALC_PERCENT)  MAX_CTRL ALWAYS USE CALC_% // if the calculation mode is percentage
				// Incorporate priority facor here (JZ)
				output1 = pEF->m_lfPriorFactor * pEF->m_lfCurrent/pEF->m_lfPostDev*100;

			strValue.Format("\t%lf", output1);
			strEF += strValue;

			obj[nIndex] = output1;

			if (output1 > 1)
				bValid = false;
		}
	}
	
	obj[0] = totalCost;

/*
	BMP_GROUP* pBMPGroup;
	double totalArea;

	pos = pBMPData->bmpGroupList.GetHeadPosition();
	while (pos != NULL)
	{
		totalArea = 0.0;
		pBMPGroup = (BMP_GROUP*) pBMPData->bmpGroupList.GetNext(pos);

		if (pBMPGroup->m_lfTotalArea > 0.0)
		{
			pos1 = pBMPGroup->m_bmpList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pBMPSite = (CBMPSite*) pBMPGroup->m_bmpList.GetNext(pos);
				totalArea += pBMPSite->GetBMPArea();
			}

			if (totalArea > pBMPGroup->m_lfTotalArea)
				value *= (totalArea/pBMPGroup->m_lfTotalArea)*10;
		}
	}
*/

	
	output = totalCost/m_pBMPRunner->pBMPData->lfCostLimit;
	if (output > 1.1)
		constr[0] = -output;
	else
		constr[0] = 0.0;

	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		bmpcost[i] = pBMPData->m_pBMPcost[i].m_lfCost;
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.nreal; i++)
	{
		strValue.Format("\t%lf", xreal[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);

	return bValid;
}

bool CBMPOptimizerGA::Evaluate_TradeOff(double *xreal, double *obj, double *bmpcost, double *constr)
{
	bool bValid = true;
	int i;
	double totalCost = 0.0;
	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;

	// mapping solution to ajustable variables
	for(i=0; i<problem.nreal; i++)
		*m_pVariables[i] = xreal[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
	int nIndex = 0;
	POSITION pos, pos1;
	pos = pBMPData->routeList.GetHeadPosition();

	strEF = "";
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

		totalCost += pBMPSite->m_lfCost;
		totalSurfaceArea += pBMPSite->m_lfSurfaceArea;
		totalExcavatnVol += pBMPSite->m_lfExcavatnVol;
		totalSurfStorVol += pBMPSite->m_lfSurfStorVol;
		totalSoilStorVol += pBMPSite->m_lfSoilStorVol;
		totalUdrnStorVol += pBMPSite->m_lfUdrnStorVol;

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			nIndex++;
			obj[nIndex] = DBL_MAX;
			constr[nIndex] = DBL_MAX;

			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);

			if (pEF->m_lfTarget < 0 )
			{
				AfxMessageBox("Target value can not be negative under minimize cost option");
				continue;
			}

			double output1 = 0.0;

			if (pEF->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
			{
				if (pEF->m_lfPostDev > 0.0)
					output1 = pEF->m_lfCurrent/pEF->m_lfPostDev*100;
//				else
//					output1 = pEF->m_lfCurrent;
			}
			else if (pEF->m_nCalcMode == CALC_VALUE) // if the calculation mode is value
			{
				output1 = pEF->m_lfCurrent;
			}
			else // if the calculation mode is scale
			{
				if (pEF->m_lfPostDev - pEF->m_lfPreDev > 0)
					output1 = (pEF->m_lfCurrent - pEF->m_lfPreDev) / (pEF->m_lfPostDev-pEF->m_lfPreDev);
//				else
//					output1 = pEF->m_lfCurrent;
			}

			strValue.Format("\t%lf", output1);
			strEF += strValue;

			obj[nIndex] = output1;

			if (output1 > pEF->m_lfUpperTarget+(pEF->m_lfUpperTarget+pEF->m_lfLowerTarget)/20)
				constr[nIndex] = pEF->m_lfUpperTarget-output1;
			else if (output1 < pEF->m_lfLowerTarget-(pEF->m_lfUpperTarget+pEF->m_lfLowerTarget)/20)
				constr[nIndex] = output1-pEF->m_lfLowerTarget;
			else
				constr[nIndex] = 0;
		}
	}

	constr[0] = 0;
	obj[0] = totalCost;
	
	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		bmpcost[i] = pBMPData->m_pBMPcost[i].m_lfCost;
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.nreal; i++)
	{
		strValue.Format("\t%lf", xreal[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);

	return bValid;
}

void CBMPOptimizerGA::OutputFileHeader(CString header, FILE* fp)
{
	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	ADJUSTABLE_PARAM* pAP;
	EVALUATION_FACTOR* pEF;
	CString strLine, strValue, strValue2;

	int nEF = 0;
	strValue2 = "";
	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
		nEF += pBMPSite->m_factorList.GetCount();

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
			strValue.Format("\t%s_%s_%d", pBMPSite->m_strID, pEF->m_strFactor, pEF->m_nCalcMode);
			strValue2 += strValue;
		}
	}

	strLine.Format("%d\t%d\t" + header + "\n",nEF,m_pBMPRunner->pBMPData->nBMPtype);
	strLine += "NO.\tTotalCost($)\tTotalSurfaceArea(ac)\tTotalExcavatnVol(ac-ft)\tTotalSurfStorVol(ac-ft)\tTotalSoilStorVol(ac-ft)\tTotalUdrnStorVol(ac-ft)";
	strLine += strValue2;

	//add cost for each unique bmp type
	for(int i=0; i<m_pBMPRunner->pBMPData->nBMPtype; i++)
	{
		strValue.Format("\t%s", m_pBMPRunner->pBMPData->m_pBMPcost[i].m_strBMPType);
		strLine += strValue;
	}

	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);

		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			strValue.Format("\t%s_%s", pBMPSite->m_strID, pAP->m_strVariable);
			strLine += strValue;
		}
	}

	strLine += "\n";
	fputs(strLine, fp);
	fflush(fp);
}

void CBMPOptimizerGA::OutputBestPopulation()
{
	int i, j;
	if (m_pBMPRunner == NULL)
		return;

	if (m_pBMPRunner->pBMPData == NULL)
		return;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;
	CString strLine, strValue;

	CBMPSite* pBMPSite;
	CBMPData* pBMPData = m_pBMPRunner->pBMPData;
	POSITION pos;
	
	for (i=0; i<parent_pop->pProblem->popsize; i++)
	{
		// mapping solution to ajustable variables
		for (j=0; j<parent_pop->pProblem->nreal; j++)
			*m_pVariables[j] = parent_pop->individuals[i].xreal[j];

		totalSurfaceArea = 0.0;
		totalExcavatnVol = 0.0;
		totalSurfStorVol = 0.0;
		totalSoilStorVol = 0.0;
		totalUdrnStorVol = 0.0;

		pos = pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

			//calculate the BMP area, and storages
			pBMPSite->m_lfSurfaceArea = 0.0;
			pBMPSite->m_lfExcavatnVol = 0.0;
			pBMPSite->m_lfSurfStorVol = 0.0;
			pBMPSite->m_lfSoilStorVol = 0.0;
			pBMPSite->m_lfUdrnStorVol = 0.0;

			double soilporosity   = pBMPSite->m_lfPorosity;
			double udsoilporosity = pBMPSite->m_lfUndVoid;	
			double sqft2acre = 2.295675e-005;
			double cuft2acft = 2.295675e-005;

			if (pBMPSite->m_nBMPClass == CLASS_A)
			{
				BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;

				int    releasetype    = pBMP->m_nORelease;		    
				double basinlength    = pBMP->m_lfBasinLength;			//ft
				double basinwidth     = pBMP->m_lfBasinWidth;			//ft
				double weirheight     = pBMP->m_lfWeirHeight;			//ft
				double soildepth      = pBMPSite->m_lfSoilDepth;		//ft	
				double udsoildepth    = pBMPSite->m_lfUndDepth;			//ft

				double BMParea = basinlength * basinwidth;				//ft^2	 
				
				// check if this BMP is cistern or rainbarrel, if so then 
				if (releasetype == 1 || releasetype == 2)
					BMParea = 3.142857/4.0*pow(basinlength,2);	// ft2 
				
				double BMPdepth2 = weirheight + soildepth + udsoildepth;	// ft	 
				
				pBMPSite->m_lfSurfaceArea = BMParea * pBMPSite->m_lfBMPUnit * sqft2acre;//acre
				pBMPSite->m_lfExcavatnVol = BMParea * BMPdepth2 * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfSurfStorVol = BMParea * weirheight * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfSoilStorVol = BMParea * soildepth * soilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfUdrnStorVol = BMParea * udsoildepth * udsoilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			}
			else if (pBMPSite->m_nBMPClass == CLASS_B)
			{
				BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;

				double BMPlength   = pBMP->m_lfBasinLength;
				double BMPwidth    = pBMP->m_lfBasinWidth;
				double BMPdepth    = pBMP->m_lfMaximumDepth;
				double soildepth   = pBMPSite->m_lfSoilDepth;	
				double udsoildepth = pBMPSite->m_lfUndDepth;	
				double slope1      = pBMP->m_lfSideSlope1;
				double slope2      = pBMP->m_lfSideSlope2;

				//calculate the top width (ft)
				double top_width = (BMPdepth/slope1 + BMPdepth/slope2 + BMPwidth); 

				double BMParea     = BMPlength * top_width;						// ft^2		 
				double BMPdepth2   = BMPdepth + soildepth + udsoildepth;		// ft	 

				pBMPSite->m_lfSurfaceArea = BMParea * pBMPSite->m_lfBMPUnit * sqft2acre;//acre
				pBMPSite->m_lfExcavatnVol = BMParea * BMPdepth2 * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfSurfStorVol = BMParea * BMPdepth * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfSoilStorVol = BMParea * soildepth * soilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
				pBMPSite->m_lfUdrnStorVol = BMParea * udsoildepth * udsoilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			}

			totalSurfaceArea += pBMPSite->m_lfSurfaceArea;
			totalExcavatnVol += pBMPSite->m_lfExcavatnVol;
			totalSurfStorVol += pBMPSite->m_lfSurfStorVol;
			totalSoilStorVol += pBMPSite->m_lfSoilStorVol;
			totalUdrnStorVol += pBMPSite->m_lfUdrnStorVol;
		}

		strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", i+1, parent_pop->individuals[i].obj[0], 
						totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
						totalSoilStorVol, totalUdrnStorVol);

		for (j=1; j<parent_pop->pProblem->nobj; j++)
		{
			strValue.Format("\t%lf", parent_pop->individuals[i].obj[j]);//evaluation factors
			strLine += strValue;
		}

		//add cost for each unique bmp type
		for(j=0; j<pBMPData->nBMPtype; j++)
		{
			strValue.Format("\t%lf", parent_pop->individuals[i].BMPcost[j]);
			strLine += strValue;
		}

		for (j=0; j<parent_pop->pProblem->nreal; j++)
		{
			strValue.Format("\t%lf", parent_pop->individuals[i].xreal[j]);//decision variables
			strLine += strValue;
		}
		strLine += "\n";

		fputs(strLine, fpBestPop);
		fflush(fpBestPop);	
	}

	return;
}
/*
void CBMPOptimizerGA::OutputBestSolutions()
{
	int i, j;
	if (m_pBMPRunner == NULL)
		return;

	if (m_pBMPRunner->pBMPData == NULL)
		return;

	int bestSolIndex = -1;
	double minCost = DBL_MAX;
	double totalCost = 0.0;
	CString strLine, strValue, strFilePath;

	POSITION pos, pos1;
	FILE *fp = NULL;
	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
		
	strFilePath = m_pBMPRunner->pBMPData->strOutputDir + "\\BestSolutions.out";
	fp = fopen(LPCSTR(strFilePath), "wt");
	if(fp == NULL)
		return;

	OutputFileHeader("NSGA-II - Best Solutions", fp);

	for(i=0; i<m_pBMPRunner->pBMPData->nSolution;i++)
	{
		if (i==0)
		{
			//get the best solution
			bestSolIndex = parent_pop->GetBestSolutionIndex();
			if (bestSolIndex != -1)
				minCost = parent_pop->individuals[bestSolIndex].obj[0];
			else
				break;
		}
		else
		{
			//get the next best solution
			bestSolIndex = parent_pop->GetNextBestSolutionIndex(minCost);
			if (bestSolIndex != -1)
				minCost = parent_pop->individuals[bestSolIndex].obj[0];
			else
				break;
		}

		// mapping solution to ajustable variables
		for(j=0; j<parent_pop->pProblem->nreal; j++)
			*m_pVariables[j] = parent_pop->individuals[bestSolIndex].xreal[j];
		
		strValue.Format("Best%d", i+1);
		if (!m_pBMPRunner->pBMPData->OpenOutputFiles(strValue))	// time series for the best solution
			return;
		if (!m_pBMPRunner->OpenOutputFiles(strValue, m_pBMPRunner->pBMPData->nRunOption, RUN_OUTPUT))
			return;
		m_pBMPRunner->RunModel(RUN_OUTPUT);

		if (!m_pBMPRunner->pBMPData->CloseOutputFiles())
			return;
		if (!m_pBMPRunner->CloseOutputFiles())
			return;

		totalCost = 0.0;
		pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
			totalCost += pBMPSite->m_lfCost;
		}

		strLine.Format("%d\t%lf\t%d", i+1, totalCost, parent_pop->individuals[bestSolIndex].validSolution);
		for(j=0; j<parent_pop->pProblem->nreal; j++)
		{
			strValue.Format("\t%lf", parent_pop->individuals[bestSolIndex].xreal[j]);
			strLine += strValue;
		}

		pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
			pos1 = pBMPSite->m_factorList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
				strValue.Format("\t%lf", pEF->m_lfCurrent); 
				strLine += strValue;
			}
		}
		strLine += "\n";
		fputs(strLine, fp);
	}
	
	fclose(fp);
}
*/