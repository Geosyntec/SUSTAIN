// Population.cpp: implementation of the CPopulation class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "Individual.h"
#include "Population.h"
#include <afxtempl.h>
#include <math.h>
#include <float.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

extern int random_int(int low, int high);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPopulation::CPopulation()
{
	nSize = 0;
	pProblem = NULL;
	pGAOptimizer = NULL;
	individuals = NULL;
}

CPopulation::CPopulation(int nsize, GA_PROBLEM *pProb, CBMPOptimizerGA *pOptimizer)
{
	ASSERT(pProb != NULL);
	ASSERT(pOptimizer != NULL);

	nSize = nsize;
	pProblem = pProb;
	pGAOptimizer = pOptimizer;

	individuals = new CIndividual[nSize];
	for (int i=0; i<nSize; i++)
		individuals[i].Allocate(pProblem, pOptimizer);
}

CPopulation::~CPopulation()
{
	if (individuals != NULL)
		delete []individuals;
}

void CPopulation::Initialize()
{
	ASSERT(individuals != NULL);
	ASSERT(pProblem != NULL);

	for (int i=0; i<nSize; i++)
		individuals[i].Initialize();
}

void CPopulation::Evaluate()
{
	ASSERT(individuals != NULL);
	ASSERT(pProblem != NULL);
	ASSERT(pGAOptimizer != NULL);

	for (int i=0; i<nSize; i++)
		individuals[i].Evaluate();
}

void CPopulation::Mutate()
{
	ASSERT(individuals != NULL);
	ASSERT(pProblem != NULL);

	for (int i=0; i<nSize; i++)
		individuals[i].Mutate();
}

void CPopulation::AssignRankAndCrowdingDistance()
{
	ASSERT(individuals != NULL);
	ASSERT(pProblem != NULL);

    int front_size = 0;
    int rank = 1;
	int orig_index, cur_index;

    POSITION pos1, pos2, tmp_pos;
	CList<int, int> orig;
    CList<int, int> cur;
    
    for (int i=0; i<pProblem->popsize; i++)
		orig.AddTail(i);

    while (!orig.IsEmpty())
    {
		pos1 = orig.GetHeadPosition();
		orig_index = orig.GetNext(pos1);

        if (pos1 == NULL)
        {
            individuals[orig_index].rank = rank;
            individuals[orig_index].crowd_dist = INF;
            break;
        }

		cur.AddHead(orig_index);
        front_size = 1;

		orig.RemoveHead();
		pos1 = orig.GetHeadPosition();
		
        while (pos1 != NULL)
        {
			int flag = -1;
        
			orig_index = orig.GetAt(pos1);
			pos2 = cur.GetHeadPosition();

            while (pos2 != NULL)
            {
				cur_index = cur.GetAt(pos2);
                flag = individuals[orig_index].CheckDominance(individuals[cur_index]);

                if (flag == 1)
                {
					orig.AddHead(cur_index);
                    front_size--;

					tmp_pos = pos2;
                    cur.GetNext(pos2);
					cur.RemoveAt(tmp_pos);
                }
                else if (flag == 0)
                {
                    cur.GetNext(pos2);
                }
                else if (flag == -1)
				{
                    break;
				}
            }

            if (flag != -1)
            {
				cur.AddHead(orig_index);
                front_size++;
				
				tmp_pos = pos1;
				orig.GetNext(pos1);
				orig.RemoveAt(tmp_pos);
            }
			else
			{
				orig.GetNext(pos1);
			}
        }

        pos2 = cur.GetHeadPosition();
        while (pos2 != NULL)
        {
			cur_index = cur.GetNext(pos2);
            individuals[cur_index].rank = rank;
        }

        AssignCrowdingDistanceList((void*) &cur, front_size);
        
		cur.RemoveAll();

        rank++;
    }
}

void CPopulation::AssignCrowdingDistanceList(void* list, int front_size)
{
	ASSERT(list != NULL);

    CList<int, int> *p_list = (CList<int, int> *) list;

	POSITION pos = p_list->GetHeadPosition();
	if (pos == NULL)
		return;

	int index = p_list->GetNext(pos);

    if (front_size == 1)
    {
        individuals[index].crowd_dist = INF;
        return;
    }
	else if (front_size == 2)
    {
        individuals[index].crowd_dist = INF;
        individuals[p_list->GetAt(pos)].crowd_dist = INF;
        return;
    }

    int i;
    int **obj_array = new int*[pProblem->nobj];
    for (i=0; i<pProblem->nobj; i++)
        obj_array[i] = new int[front_size];

	pos = p_list->GetHeadPosition();
    int *dist = new int[front_size];
    for (i=0; i<front_size; i++)
        dist[i] = p_list->GetNext(pos);

    AssignCrowdingDistance(dist, obj_array, front_size);
    
	delete []dist;
    for (i=0; i<pProblem->nobj; i++)
        delete []obj_array[i];
    delete []obj_array;
}

void CPopulation::AssignCrowdingDistance(int *dist, int **obj_array, int front_size)
{
    int i, j;

    for (i=0; i<pProblem->nobj; i++)
    {
        for (j=0; j<front_size; j++)
            obj_array[i][j] = dist[j];
        QuickSortFront(i, obj_array[i], front_size);
    }

    for (j=0; j<front_size; j++)
        individuals[dist[j]].crowd_dist = 0.0;

    for (i=0; i<pProblem->nobj; i++)
        individuals[obj_array[i][0]].crowd_dist = INF;

    for (i=0; i<pProblem->nobj; i++)
	{
        for (j=1; j<front_size-1; j++)
        {
            if (individuals[obj_array[i][j]].crowd_dist != INF)
            {
                if (individuals[obj_array[i][front_size-1]].obj[i] == individuals[obj_array[i][0]].obj[i])
                    individuals[obj_array[i][j]].crowd_dist += 0.0;
                else
                    individuals[obj_array[i][j]].crowd_dist += (individuals[obj_array[i][j+1]].obj[i] - individuals[obj_array[i][j-1]].obj[i])/(individuals[obj_array[i][front_size-1]].obj[i] - individuals[obj_array[i][0]].obj[i]);
            }
        }
	}

    for (j=0; j<front_size; j++)
    {
        if (individuals[dist[j]].crowd_dist != INF)
        {
            individuals[dist[j]].crowd_dist = individuals[dist[j]].crowd_dist/pProblem->nobj;
        }
    }
}

/* 
   Routine to compute crowding distance based on objective function values 
   when the population is in the form of an array
 */
void CPopulation::AssignCrowdingDistanceIndices(int c1, int c2)
{
    int i;
    
	int front_size = c2-c1+1;

    if (front_size == 1)
    {
        individuals[c1].crowd_dist = INF;
        return;
    }

    if (front_size == 2)
    {
        individuals[c1].crowd_dist = INF;
        individuals[c2].crowd_dist = INF;
        return;
    }

    int **obj_array = new int*[pProblem->nobj];
    for (i=0; i<pProblem->nobj; i++)
        obj_array[i] = new int[front_size];

    int *dist = new int[front_size];
    for (i=0; i<front_size; i++)
        dist[i] = c1++;

    AssignCrowdingDistance(dist, obj_array, front_size);

    delete []dist;
    for (i=0; i<pProblem->nobj; i++)
        delete []obj_array[i];
    delete []obj_array;
}

/* Randomized quick sort routine to sort a population based on a particular objective chosen */
void CPopulation::QuickSortFront(int objcount, int *obj_array, int obj_array_size)
{
	QuickSortFrontImpl(objcount, obj_array, 0, obj_array_size-1);
}

/* Actual implementation of the randomized quick sort used to sort a population based on a particular objective chosen */
void CPopulation::QuickSortFrontImpl(int objcount, int *obj_array, int left, int right)
{
    if (left < right)
    {
        int index = random_int(left, right);
        int temp = obj_array[right];
        obj_array[right] = obj_array[index];
        obj_array[index] = temp;

        int tmp_index = left-1;
		double pivot = individuals[obj_array[right]].obj[objcount];

        for (int i=left; i<right; i++)
        {
            if (individuals[obj_array[i]].obj[objcount] <= pivot)
            {
                tmp_index++;
                temp = obj_array[i];
                obj_array[i] = obj_array[tmp_index];
                obj_array[tmp_index] = temp;
            }
        }
        
		index = tmp_index + 1;
        temp = obj_array[index];
        obj_array[index] = obj_array[right];
        obj_array[right] = temp;

        QuickSortFrontImpl(objcount, obj_array, left, index-1);
        QuickSortFrontImpl(objcount, obj_array, index+1, right);
    }
}

/* Randomized quick sort routine to sort a population based on crowding distance */
void CPopulation::QuickSortDist(int *dist, int front_size)
{
    QuickSortDistImpl(dist, 0, front_size-1);
}

/* Actual implementation of the randomized quick sort used to sort a population based on crowding distance */
void CPopulation::QuickSortDistImpl(int *dist, int left, int right)
{
    if (left < right)
    {
        int index = random_int(left, right);
        int temp = dist[right];
        dist[right] = dist[index];
        dist[index] = temp;

        int tmp_index = left-1;
        double pivot = individuals[dist[right]].crowd_dist;

        for (int i=left; i<right; i++)
        {
            if (individuals[dist[i]].crowd_dist <= pivot)
            {
                tmp_index++;
                temp = dist[i];
                dist[i] = dist[tmp_index];
                dist[tmp_index] = temp;
            }
        }
        
		index = tmp_index + 1;
        temp = dist[index];
        dist[index] = dist[right];
        dist[right] = temp;

        QuickSortDistImpl(dist, left, index-1);
        QuickSortDistImpl(dist, index+1, right);
    }
}

/* Function to print the information of an individual solution into a file */
void CPopulation::ReportIndividualToFile(FILE *fp, int index)
{
    int i;

    fprintf(fp, "%d\t", index+1);

    for (i=0; i<pProblem->nobj; i++)
        fprintf(fp, "%lf\t", individuals[index].obj[i]);

    for (i=0; i<pProblem->nreal; i++)
    {
        fprintf(fp, "%lf\t", individuals[index].xreal[i]);
    }

    fprintf(fp, "\n");

//    fprintf(fp, "%lf\t", individuals[index].constr_violation);
//    fprintf(fp, "%d\t", individuals[index].rank);
//    fprintf(fp, "%lf\n", individuals[index].crowd_dist);
}

/* Function to print the information of a population into a file */
void CPopulation::ReportAllToFile(FILE *fp)
{
    for (int i=0; i<pProblem->popsize; i++)
		ReportIndividualToFile(fp, i);
}

/* Function to print the information of feasible and non-dominated population into a file */
void CPopulation::ReportBestToFile(FILE *fp)
{
    for (int i=0; i<pProblem->popsize; i++)
	{
        if (individuals[i].constr_violation == 0.0 && individuals[i].rank == 1)
		{
			ReportIndividualToFile(fp, i);
			break;
		}
    }
}

int CPopulation::GetBestSolutionIndex()
{
	int index = -1;
	double minCost = DBL_MAX;
    for (int i=0; i<pProblem->popsize; i++)
	{
        if (individuals[i].validSolution && individuals[i].obj[0] < minCost)
		{
			minCost = individuals[i].obj[0];
			index = i;
		}
    }

	return index;
}

int CPopulation::GetNextBestSolutionIndex(double prevMinCost)
{
	int index = -1;
	double minCost = DBL_MAX;
    for (int i=0; i<pProblem->popsize; i++)
	{
        if (individuals[i].validSolution && 
			individuals[i].obj[0] < minCost &&
			individuals[i].obj[0] > prevMinCost)
		{
			minCost = individuals[i].obj[0];
			index = i;
		}
    }

	return index;
}
