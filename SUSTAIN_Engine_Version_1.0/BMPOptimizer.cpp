// BMPOptimizer.cpp: implementation of the CBMPOptimizer class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "BMPSite.h"
#include "BMPData.h"
#include "BMPRunner.h"
#include "BMPOptimizer.h"
#include "ProgressWnd.h" 
#include <math.h>
#include <float.h>


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

// definition for ameoba algorithm
#define NMAX	5000
#define ALPHA	1.0
#define BETA	0.5
#define GAMMA	2.0

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBMPOptimizer::CBMPOptimizer()
{
	nMaxIter     = 0;
	nRunCounter  = 0;
	nMaxRun		 = 0;
	m_lfPrevResult = 0.0;
	m_lfPrevValue = 0.0;
	m_pVariables = NULL;
	m_pBMPRunner = NULL;
	m_pAllSolutions = NULL;
//	m_pDebug = NULL;

	problem.high    = NULL;
	problem.low     = NULL;
	problem.inc     = NULL;
	problem.ranges  = NULL;
	problem.value1  = NULL;
	problem.refSet1 = NULL;
	problem.order1  = NULL;
	problem.iter1   = NULL;
	problem.value2  = NULL;
	problem.refSet2 = NULL;
	problem.order2  = NULL;
	problem.iter2   = NULL;
}

CBMPOptimizer::CBMPOptimizer(CBMPRunner* pBMPRunner)
{
	nMaxIter     = 5;	// Number of global iterations
	nRunCounter  = 0;
	nMaxRun		 = 10000;
	m_lfPrevResult = 0.0;
	m_lfPrevValue = 0.0;
	m_pVariables = NULL;
	m_pBMPRunner = pBMPRunner;
	m_pAllSolutions = NULL;
//	m_pDebug = NULL;

	problem.high    = NULL;
	problem.low     = NULL;
	problem.inc     = NULL;
	problem.ranges  = NULL;
	problem.value1  = NULL;
	problem.refSet1 = NULL;
	problem.order1  = NULL;
	problem.iter1   = NULL;
	problem.value2  = NULL;
	problem.refSet2 = NULL;
	problem.order2  = NULL;
	problem.iter2   = NULL;
}

CBMPOptimizer::~CBMPOptimizer()
{
	if (m_pVariables != NULL)
		delete []m_pVariables;

	if (problem.high)
		delete []problem.high;
	if (problem.low)
		delete []problem.low;
	if (problem.inc)
		delete []problem.inc;
	if (problem.ranges)
	{
		for (int i=0; i<problem.n_var; i++)
			delete []problem.ranges[i];
		delete []problem.ranges;
	}
	if (problem.value1)
		delete []problem.value1;
	if (problem.refSet1)
	{
		for (int i=0; i<problem.b1; i++)
			delete []problem.refSet1[i];
		delete []problem.refSet1;
	}
	if (problem.order1)
		delete []problem.order1;
	if (problem.iter1)
		delete []problem.iter1;
	if (problem.value2)
		delete []problem.value2;
	if (problem.refSet2)
	{
		for (int i=0; i<problem.b2; i++)
			delete []problem.refSet2[i];
		delete []problem.refSet2;
	}
	if (problem.order2)
		delete []problem.order2;
	if (problem.iter2)
		delete []problem.iter2;

	if (problem.evaSolutions)
	{
		for (int i=0; i<problem.PSize; i++)
			delete []problem.evaSolutions[i];
		delete []problem.evaSolutions;
	}

/*
	if (problem.evaOutputs)
	{
		for (int i=0; i<problem.PSize; i++)
			delete []problem.evaOutputs[i];
		delete []problem.evaOutputs;
	}
*/
	delete []problem.evaValues;
	delete []problem.evaOrders;
}


void CBMPOptimizer::InitProblem(int nvar, int b1, int b2, int pSize, bool localSearch)
{
	if (m_pBMPRunner == NULL)
		return;

	int i, j;

	problem.n_var  = nvar;
	problem.b1	   = b1;
	problem.b2	   = b2;
	problem.PSize  = pSize;
	problem.LS     = localSearch;
	problem.digits = 0;
	problem.last_combine = 0;
	problem.iter   = 0;

	problem.high   = new double[nvar];
	problem.low    = new double[nvar];
	problem.inc    = new double[nvar];
	problem.ranges = new int*[nvar];
	for (i=0; i<nvar; i++)
	{
		problem.ranges[i] = new int[5];
		for (j=0; j<5; j++)
			problem.ranges[i][j] = 0;
	}

	problem.value1  = new double[b1];
	problem.refSet1	= new double*[b1];
	for (i=0; i<b1; i++)
		problem.refSet1[i] = new double[nvar];
	problem.order1  = new int[b1];
	problem.iter1   = new int[b1];
	
	problem.value2  = new double[b2];
	problem.refSet2 = new double*[b2];
	for (i=0; i<b2; i++)
		problem.refSet2[i] = new double[nvar];
	problem.order2  = new int[b2];
	problem.iter2   = new int[b2];

	m_pVariables = new double*[nvar];

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;
	if (pBMPData == NULL)
		return;

	ADJUSTABLE_PARAM* pAP;
	CBMPSite* pBMPSite;
	int nIndex = 0;
	POSITION pos, pos1;

	pos = pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);
		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			problem.high[nIndex] = pAP->m_lfTo;
			problem.low[nIndex]  = pAP->m_lfFrom;
			problem.inc[nIndex]  = pAP->m_lfStep;
			m_pVariables[nIndex] = pBMPSite->GetVariablePointer(pAP->m_strVariable);
			nIndex++;
		}
	}

	// initialize addtional variables for TradeOff Curve - March 5, 2007
	problem.evaSolutions = new double*[pSize];
	for (i=0; i<pSize; i++)
		problem.evaSolutions[i] = new double[nvar];

	/*
	problem.evaOutputs = new double*[pSize];
	for (i=0; i<pSize; i++)
		problem.evaOutputs[i] = new double[nvar];
	*/

	problem.evaValues = new double[pSize];
	problem.evaOrders = new int[pSize];
	for (i=0; i<pSize; i++)
		problem.evaOrders[i] = -1; // -1 means that this order has not been taken (No enough PSize solutions have been evaluated)
}

void CBMPOptimizer::InitRefSet()
{
	double *current, *min_dist, *value, **solutions;
	int i, j, k, a, *index, *index2, cont=0;
	double d, dmax, current_value;

	problem.iter = 0;
	problem.new_elements = 1;

	current   = new double[problem.n_var];		// array storing the current solution
	min_dist  = new double[problem.PSize];		// array storing the distance value for all solutions (PSize)
	value     = new double[problem.PSize];		// array storing the obj value for all solutions (PSize)
	solutions = new double*[problem.PSize];		// matrix storing all solutions (PSize x number of variables)
	for (i=0; i<problem.PSize; i++)
		solutions[i] = new double[problem.n_var];

	for (i=0; i<problem.PSize; i++)
	{
		// Generate new solution
		for(j=0; j<problem.n_var; j++)	
			current[j] = GenerateValue(j);		// call GenerateValue "Diversification Generation Method"

/*
		// Evaluate Solution
		if(nRunCounter > nMaxRun)
			return;
*/
		
		if(problem.LS)
		{
			current_value = Evaluate(current);
			ImproveSolution(current, &current_value);
		}

		if(IsNewSolution(solutions, i-1, current))
		{
			// Store solution in matrix "solutions"
			for(j=0; j<problem.n_var; j++)
				solutions[i][j] = current[j];
		
			value[i] = Evaluate(current);
		}
		else 
		{
			i--;
			cont++;
		}

		if(cont > problem.PSize/2)
		{
			problem.digits++;
			cont = 0;
		}
	}
	
	index = new int[problem.PSize];
	GetOrderIndices(index, value, problem.PSize, -1);

	// Add the best b1 to RefSet1
	for(i=0; i<problem.b1; i++)
	{
		for(j=0; j<problem.n_var; j++)
			problem.refSet1[i][j] = solutions[index[i]][j];
		
		problem.value1[i] = value[index[i]];
		problem.order1[i] = i;
		problem.iter1[i]  = 0;
	}

	// Compute minimum distances
	for(i=0; i<problem.PSize; i++)
		min_dist[i] = DistanceToRefSet1(solutions[i]);

	// Add the second b2 to RefSet2
	for(i=0; i<problem.b2; i++)
	{
		// Select the solution with maximum minimum-distance
		dmax = -1.0;
		for(j=0; j<problem.PSize; j++)
			if(min_dist[j] > dmax)
			{
				dmax = min_dist[j];
				a = j;
			}

		for(j=0; j<problem.n_var; j++)
			problem.refSet2[i][j] = solutions[a][j];

		problem.value2[i] = min_dist[a];

		// Update minimum distances
		for(k=0; k<problem.PSize; k++)
		{
			d = 0;
//			for(j=0; j<problem.n_var; j++)
//				d += pow(solutions[k][j] - solutions[a][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
			for(j=0; j<problem.n_var; j++)
				d += pow((solutions[k][j] - solutions[a][j])/(problem.high[j] - problem.low[j]), 2);
			if(d < min_dist[k])
				min_dist[k] = d;
		}
	}

	// Update minimum distances in RefSet2
	for(i=0; i<problem.b2; i++)
	//	for(k=0; k<problem.b2 && k!=i; k++) // correction made on March, 12, 2007
		for(k=i+1; k<problem.b2; k++)
		{
			d = 0;
//			for(j=0; j<problem.n_var; j++)
//				d += pow(problem.refSet2[i][j] - problem.refSet2[k][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
			for(j=0; j<problem.n_var; j++)
				d += pow((problem.refSet2[i][j] - problem.refSet2[k][j])/(problem.high[j] - problem.low[j]), 2);

			if(problem.value2[i] > d)
				problem.value2[i] = d;
		}

	index2 = new int[problem.b2];
	GetOrderIndices(index2, problem.value2, problem.b2, 1);
	
	for(i=0; i<problem.b2; i++)
	{
		problem.order2[i] = index2[i];
		problem.iter2[i]  = 0;
	}

	delete []current;
	delete []min_dist;
	delete []value;
	for (i=0; i<problem.PSize; i++)
		delete []solutions[i];
	delete []solutions;
	delete []index;
	delete []index2;
}

void CBMPOptimizer::ResetRefSet()
{
	double *current, *min_dist;
	int i, j, k, a, *index2;
	double d, dmax;
//	double current_value;

	problem.iter = 0;
	problem.new_elements = 1;
	problem.digits = 0;

	current   = new double[problem.n_var];		// array storing the current solution
	min_dist  = new double[problem.PSize];		// array storing the distance value for all solutions (PSize)

/*
	for (i=0; i<problem.PSize; i++)
	{
		// Get a solution from the cached array
		for(j=0; j<problem.n_var; j++)	
			current[j] = problem.evaSolutions[i][j];

		if(problem.LS)
		{
			current_value = Evaluate(current);
			ImproveSolution(current, &current_value);
		}

		// Store solution in matrix "solutions"
		for(j=0; j<problem.n_var; j++)
			problem.evaSolutions[i][j] = current[j];
	
		problem.evaValue[i] = Evaluate(current);
	}
	
	GetOrderIndices(problem.evaOrders, problem.evaValues, problem.PSize, -1);
*/

	// Add the best b1 to RefSet1
	for(i=0; i<problem.b1; i++)
	{
		for(j=0; j<problem.n_var; j++)
			problem.refSet1[i][j] = problem.evaSolutions[problem.evaOrders[i]][j];
		
		problem.value1[i] = problem.evaValues[problem.evaOrders[i]];
		problem.order1[i] = i;
		problem.iter1[i]  = 0;
	}

	// Compute minimum distances
	for(i=0; i<problem.PSize; i++)
	{
		int orderIndex = problem.evaOrders[i];
		if(i<problem.b1)
			min_dist[orderIndex] = 0.0;
		else
			min_dist[orderIndex] = DistanceToRefSet1(problem.evaSolutions[orderIndex]);
	}

	// Add the second b2 to RefSet2
	for(i=0; i<problem.b2; i++)
	{
		// Select the solution with maximum minimum-distance
		dmax = -1.0;
		for(j=0; j<problem.PSize; j++)
			if(min_dist[j] > dmax)
			{
				dmax = min_dist[j];
				a = j;
			}

		for(j=0; j<problem.n_var; j++)
			problem.refSet2[i][j] = problem.evaSolutions[a][j];
		problem.value2[i] = min_dist[a];

		// Update minimum distances
		for(k=0; k<problem.PSize; k++)
		{
			d = 0;
			for(j=0; j<problem.n_var; j++)
				d += pow((problem.evaSolutions[k][j] - problem.evaSolutions[a][j])/(problem.high[j] - problem.low[j]), 2);
			if(d < min_dist[k])
				min_dist[k] = d;
		}
	}

	// Update minimum distances in RefSet2
	for(i=0; i<problem.b2; i++)
		for(k=i+1; k<problem.b2; k++)
		{
			d = 0;
			for(j=0; j<problem.n_var; j++)
				d += pow((problem.refSet2[i][j] - problem.refSet2[k][j])/(problem.high[j] - problem.low[j]), 2);

			if(problem.value2[i] > d)
				problem.value2[i] = d;
		}

	index2 = new int[problem.b2];
	GetOrderIndices(index2, problem.value2, problem.b2, 1);
	
	for(i=0; i<problem.b2; i++)
	{
		problem.order2[i] = index2[i];
		problem.iter2[i]  = 0;
	}

	delete []current;
	delete []min_dist;
	delete []index2;

	for (i=0; i<problem.PSize; i++)
		problem.evaOrders[i] = -1; // -1 means that this order has not been taken (No enough PSize solutions have been evaluated)
}

// Diversification Generation Method
double CBMPOptimizer::GenerateValue(int a)
{
	int i, j;

	double low   = problem.low[a];
	double inc   = problem.inc[a];
	double range = problem.high[a] - problem.low[a];
	int*   frec  = problem.ranges[a]; // frequency count
	int*   rfrec = new int[5]; // reverse frec to penalize high frecs
	for(i=0;i<5;i++)
		rfrec[i] = 0;

	for(i=1; i<5; i++)
	{
		rfrec[i]  = frec[0] - frec[i];
		rfrec[0] += rfrec[i];
	}

	if(rfrec[0] == 0)
		i = rand()%4 + 1; // return a random number in the range from 1 to 4
	else
	{
		// select a subrange (from 1 to 4) according to rfrec
		j = rand()%rfrec[0] + 1;
		i = 1;
		while(j > rfrec[i])
		{
			j -= rfrec[i];
			i++;
		}
		if(i > 4)
		{
			delete []rfrec;
			// Abort with error message ("Problems generating values");
		}
	}
	delete []rfrec;

	// i is the selected subrange
	frec[0]++;
	frec[i]++;

	// randomly select an element in subrange i
	double r = rand()%10001/10000.0;
    
	int Ninc  = int(range/inc);
	double value = low + int((i-1+r)*Ninc/4 + 0.5) * inc;
	return value;
}

double CBMPOptimizer::Evaluate(double* sol)
{
	double value;
	nRunCounter++;

	switch (m_pBMPRunner->pBMPData->nRunOption)
	{
		case OPTION_MIMIMIZE_COST:
			value = Evaluate_MinCost(sol);
			break;
		case OPTION_MAXIMIZE_CONTROL:
			value = Evaluate_MaxCtrl(sol);
			break;
		case OPTION_TRADE_OFF_CURVE:
			value = Evaluate_TradeOff(sol);
			break;
		default:
			value = 0.0;
	}

	return value;
}

double CBMPOptimizer::Evaluate_MinCost(double* sol)    //Evaluation Function for Minimizing Cost Given Control Target
{
	int i;
	double output, value = 0.0; // value = f(BMPOutputForEF, target, cost)
	double totalCost = 0.0, costWeight = 50.0;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;

	// mapping solution to ajustable variables
	for(i=0; i<problem.n_var; i++)
		*m_pVariables[i] = sol[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
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
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);

			if (pEF->m_lfTarget < 0 )
			{
				AfxMessageBox("Target value can not be negative under minimize cost option");
				continue;
			}
	
			double output1 = 0.0; 

			if (pEF->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
			{
				if (pEF->m_lfPostDev > 0)
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

			if (pEF->m_lfTarget == 0.0)
			{
				if (output1 > pEF->m_lfTarget)
					output = output1 * 1E5;
				else
					output = 0.0;
			}
			else
				output = output1/pEF->m_lfTarget; // normalize Evaluation Factor using Target
			
			if (output > 1)
				output *= 1E5;	// penalize if the constraint is NOT met

			value += output;	// add up all output
		}
	}
	
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

	// convert cost to an equivalent value of less magnitude and add COST factor into value calculation
	if (totalCost > 1)
		value += costWeight * log(totalCost);

	//strLine.Format("%d\t%lf\t%lf", nRunCounter, totalCost, value);
	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.n_var; i++)
	{
		strValue.Format("\t%lf", sol[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);

	return value;
}

double CBMPOptimizer::Evaluate_MaxCtrl(double* sol)  //Evaluation Function for Maximizing Control Given Cost Limit
{
	int i;
	double output, value = 0.0; // value = f(BMPOutputForEF, target, cost)
	double totalCost = 0.0, costWeight = 50.0;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;

	// mapping solution to ajustable variables
	for(i=0; i<problem.n_var; i++)
		*m_pVariables[i] = sol[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
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
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
			output = 0.0;

			if (pEF->m_lfPostDev > 0)
				//if (pEF->m_nCalcMode == CALC_PERCENT)  MAX_CTRL ALWAYS USE CALC_% // if the calculation mode is percentage
				// Incorporate priority facor here (JZ)
				output = pEF->m_lfPriorFactor * pEF->m_lfCurrent/pEF->m_lfPostDev*100;

			strValue.Format("\t%lf", output);
			strEF += strValue;

			if (output > 1)
				output *= 1000;	// penalize if current value is greater than init value

			value += output;		// add up all output
		}
	}
	
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

	
	if (m_pBMPRunner->pBMPData->lfCostLimit == 0.0)
		output = totalCost;
	else
	{
		output = totalCost/m_pBMPRunner->pBMPData->lfCostLimit;
		if (output > 1)
			output *= 1E5;	// penalize if the constraint is NOT met
	}

	value += output;

	//strLine.Format("%d\t%lf\t%lf", nRunCounter, totalCost, value);
	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.n_var; i++)
	{
		strValue.Format("\t%lf", sol[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);
	return value;
}

double CBMPOptimizer::Evaluate_TradeOff(double* sol)    //Evaluation Function for Minimizing Cost Given Control Target Range
{
	int i;
	double output, value = 0.0; // value = f(BMPOutputForEF, target, cost)
	double totalCost = 0.0, costWeight = 50.0;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString strLine, strValue, strEF;
	double valueForNextTarget = 0.0;

	// mapping solution to ajustable variables
	for(i=0; i<problem.n_var; i++)
		*m_pVariables[i] = sol[i];

	// run model with the new solution in optimize mode
	m_pBMPRunner->RunModel(RUN_OPTIMIZE);

	CBMPData* pBMPData = m_pBMPRunner->pBMPData;

	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	
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
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);

			if (pEF->m_lfTarget < 0 )
			{
				AfxMessageBox("Target value can not be negative under minimize cost option");
				continue;
			}
	
			double output1 = 0.0; 

			if (pEF->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
			{
				if (pEF->m_lfPostDev > 0)
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

			if (pEF->m_lfTarget == 0.0)
			{
				if (output1 > pEF->m_lfTarget)
					output = output1 * 1E5;
				else
					output = 0.0;
			}
			else
				output = output1/pEF->m_lfTarget; // normalize Evaluation Factor using Target
			
			if (output > 1)
				output *= 1E5;	// penalize if the constraint is NOT met

			value += output;	// add up all output

			// added below for preparing for next run for new target
			if (pEF->m_lfNextTarget == 0.0)
			{
				if (output1 > pEF->m_lfNextTarget)
					output = output1 * 1E5;
				else
					output = 0.0;
			}
			else
				output = output1/pEF->m_lfNextTarget; // normalize Evaluation Factor using Target
			
			if (output > 1)
				output *= 1E5;	// penalize if the constraint is NOT met

			valueForNextTarget += output;	// add up all output
		}
	}
	
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
			{
				value *= (totalArea/pBMPGroup->m_lfTotalArea)*10;
				valueForNextTarget *= (totalArea/pBMPGroup->m_lfTotalArea)*10;
			}
		}
	}

	// convert cost to an equivalent value of less magnitude and add COST factor into value calculation
	if (totalCost > 1)
	{
		value += costWeight * log(totalCost);
		valueForNextTarget += costWeight * log(totalCost);
	}

	strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", nRunCounter, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);
	strLine += strEF;

	//add cost for each unique bmp type
	for(i=0; i<pBMPData->nBMPtype; i++)
	{
		strValue.Format("\t%lf", pBMPData->m_pBMPcost[i].m_lfCost);
		strLine += strValue;
	}

	for(i=0; i<problem.n_var; i++)
	{
		strValue.Format("\t%lf", sol[i]);
		strLine += strValue;
	}
	strLine += "\n";

	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);
	TryAddEvaluation(sol, valueForNextTarget);
	return value;
}

/*
double CBMPOptimizer::Evaluate(double* sol)  
{
	int i;
	nRunCounter++;

	double value = 0.0;
	for(i=0; i<5; i++)
		value += sol[i];
	value -= pow(sol[5], 2);
	value += pow(sol[6], 2);
	value -= pow(sol[7], 2);
	value += pow(sol[8], 2);
	value -= pow(sol[9], 2);
	
	CString strLine, strValue;

	strLine.Format("%ld\t%lf", nRunCounter, value);
	for(i=0; i<problem.n_var; i++)
	{
		strValue.Format("\t%lf", sol[i]);
		strLine += strValue;
	}
	strLine += "\n";
	fputs(strLine, m_pAllSolutions);
	fflush(m_pAllSolutions);
	return value;
}
*/

void CBMPOptimizer::ImproveSolution(double *sol, double *value)
{
	int i, j;
	double range, perturb;

	double** p = new double*[problem.n_var+1];
	for(i=0; i<=problem.n_var; i++)
		p[i] = new double[problem.n_var];
	for(i=0; i<problem.n_var; i++)
		p[0][i] = sol[i];

	double* y = new double[problem.n_var+1];
	y[0]= *value;

	for(i=0; i<problem.n_var; i++)
	{
		range   = problem.high[i] - problem.low[i];
	//	perturb = 0.1 * range;		//JZ_inc
		perturb = problem.inc[i];
		sol[i] += perturb;

		if(sol[i] > problem.high[i])
			sol[i] = problem.high[i];
	
		if(sol[i] < problem.low[i])
			sol[i] = problem.low[i];

		for(j=0; j<problem.n_var; j++)
			p[i+1][j] = sol[j];
		
//		if (nRunCounter > nMaxRun)
//			return;
		y[i+1]  = Evaluate(sol);
		sol[i] -= perturb;
	}

	int nfunk;
	// Call Nelder and Mead's Simplex method
	amoeba(p, y, problem.n_var, 0.1, &nfunk);

	int best_sol = 0;
	for(i=0; i<=problem.n_var; i++)
	{
		if(*value > y[i])
		{
			*value   = y[i];
			best_sol = i;
		}
	}

	if(best_sol > 0)
	{
		for(i=0; i<problem.n_var; i++)
		{
			sol[i] = p[best_sol][i];
													
			if(sol[i] < problem.low[i])
				sol[i] = problem.low[i];
			if(sol[i] > problem.high[i])
				sol[i] = problem.high[i];

			if(((sol[i]-problem.low[i])/problem.inc[i] - int((sol[i]-problem.low[i])/problem.inc[i])) != 0)  //JZ_inc
		       sol[i] = problem.low[i] + int((sol[i]-problem.low[i])/problem.inc[i] + 0.5) * problem.inc[i]; //JZ_inc
		}
//		if (nRunCounter > nMaxRun)
//			return;
		*value = Evaluate(sol);
	}

	delete []y;
	for(i=0; i<=problem.n_var; i++)
		delete []p[i];
	delete []p;
}


/*
double CBMPOptimizer::GetRandomNum()
{
	int i;
	long M  = 714025;
	long IA = 1366;
	long IC = 150889;

	if(problem.seed_reset == 1)
	{
		problem.seed_reset = 0;
		problem.iff = 0;
	}

	if(problem.idum < 0 || problem.iff == 0)
	{
		problem.iff  = 1;
		problem.idum = (IC - problem.idum) % M;
		if(problem.idum < 0)
			problem.idum = -problem.idum;

		for (i=1; i<=97; i++)
		{
			problem.idum  = (IA*problem.idum + IC) % M;
			problem.ir[i] = problem.idum;
		}

		problem.idum = (IA*problem.idum + IC) % M;
		problem.iy   = problem.idum;
	}

	i = (int)(1 + 97.0*problem.iy/M);
	if (i > 97 || i < 1)
		SSabort("Failure in random number generator.");

	problem.iy    = problem.ir[i];
	problem.idum  = (IA*problem.idum + IC) % M;
	problem.ir[i] = problem.idum;

	return ((double)problem.iy)/M;
}
*/

bool CBMPOptimizer::IsNewSolution(double **solutions, int dim, double *sol)
{
	int i, j;
	bool is_new;
	double precision = 1/pow(10, problem.digits);

	for(i=0; i<dim; i++)
	{
		is_new = false;
		for(j=0; j<problem.n_var; j++)
		{
			if(fabs(solutions[i][j] - sol[j]) >= precision)
			{
				is_new = true;
				break;
			}
		}

		if(!is_new)
			return false;
	}
	
	return true;
}

void CBMPOptimizer::GetOrderIndices(int* indices, double *pesos, int num, int tipo)
{
	int i, b, t, tempi;
	double temp;
	
	double* coste = new double[num];

	for(i=0; i<num; i++)
	{
		indices[i] = i;
		coste[i]   = pesos[i];
	}
	
	b = num;
	while(b != 0)
	{   
		t = 0;
		for(i=0; i<b-1; i++)
		{
			if( (tipo == 1  && coste[i] < coste[i+1]) ||
				(tipo == -1 && coste[i] > coste[i+1]))
			{
				temp = coste[i+1];
				coste[i+1] = coste[i];
				coste[i] = temp;
				
				tempi = indices[i+1];
				indices[i+1] = indices[i];
				indices[i] = tempi;
				
				t = i;
			}		
		}
		b = t;
	}

	delete []coste;
}

double CBMPOptimizer::DistanceToRefSet1(double *sol)
{
	double d, min_dist = DBL_MAX;
	int i, j;

	for(i=0; i<problem.b1; i++)
	{
		d = 0;
//		for(j=0; j<problem.n_var; j++)
//			d += pow(sol[j] - problem.refSet1[i][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
		for(j=0; j<problem.n_var; j++)
			d += pow((sol[j] - problem.refSet1[i][j])/(problem.high[j] - problem.low[j]), 2);
		if(d == 0.0)
			return 0.0;
		if(min_dist > d)
			min_dist = d;
	}

	return min_dist;
}

int CBMPOptimizer::amoeba(double **p, double *y, int ndim, double ftol, int *nfunk)
{
	int i, j, ilo, ihi, inhi, mpts=ndim+1;
	double ytry, ysave, sum, rtol, *psum;

	*nfunk = 0;
	psum = new double[ndim+1];

	for (i=0; i<ndim; i++)
	{
		sum = 0.0;
		for (j=0; j<mpts; j++)
			sum += p[j][i];
		psum[i] = sum;
	}

	for (;;)
	{
		ilo  = 0;
		ihi  = (y[0] > y[1])?0:1;
		inhi = (y[0] > y[1])?1:0;

		for (i=0; i<mpts; i++)
		{
			if (y[i] < y[ilo])
				ilo = i;

			if (y[i] > y[ihi])
			{
				inhi = ihi;
				ihi  = i;
			}
			else if (y[i] > y[inhi])
			{
				if (i != ihi)
					inhi = i;
			}
		}

		rtol = 2.0*fabs(y[ihi]-y[ilo])/(fabs(y[ihi])+fabs(y[ilo]));
		if (rtol < ftol)
			break;
		if (*nfunk >= NMAX) 
			return -1;

		ytry = amotry(p, y, psum, ndim, ihi, nfunk, -ALPHA);

		if (ytry <= y[ilo])
			ytry = amotry(p, y, psum, ndim, ihi, nfunk, GAMMA);
		else if (ytry >= y[inhi])
		{
			ysave = y[ihi];
			ytry = amotry(p, y, psum, ndim, ihi, nfunk, BETA);
			if (ytry >= ysave)
			{
				for (i=0; i<mpts; i++)
				{
					if (i != ilo)
					{
						for (j=0; j<ndim; j++)
						{
							psum[j] = 0.5*(p[i][j]+p[ilo][j]);
							p[i][j] = psum[j];
						}
//						if (nRunCounter > nMaxRun)
//							return -1;
						y[i] = Evaluate(psum);
					}
				}
				*nfunk = *nfunk + ndim;

				for (i=0; i<ndim; i++)
				{
					sum = 0.0;
					for (j=0; j<mpts; j++)
						sum += p[j][i];
					psum[i] = sum;
				}
			}
		}
	}
	
	delete []psum;
	return 0;
}

double CBMPOptimizer::amotry(double **p, double *y, double *psum, int ndim, int ihi, int *nfunk, double fac)
{
	int i;
	double fac1, fac2, ytry;
	double* ptry = new double[ndim];

	fac1 = (1.0-fac)/ndim;
	fac2 = fac1 - fac;

	for (i=0; i<ndim; i++)
		ptry[i] = psum[i]*fac1 - p[ihi][i]*fac2;

	ytry = Evaluate(ptry);
	(*nfunk)++;

	if (ytry < y[ihi]) 
	{
		y[ihi] = ytry;
		for (i=0; i<ndim; i++)
		{
			psum[i]  += ptry[i] - p[ihi][i];
			p[ihi][i] = ptry[i];
		}
	}

	delete []ptry;
	return ytry;
}

void CBMPOptimizer::CombineRefSet()
{
	int i, j, k, l, pull_size, total_size;

	problem.new_elements = 0;
	double** offsprings = new double*[4];
	for (i=0; i<4; i++)
		offsprings[i] = new double[problem.n_var];

	// New solutions are temporarily stored in a pull
	pull_size  = 0;
	total_size = (2*problem.b1*problem.b1) + (3*problem.b1*problem.b2) + (problem.b2*problem.b2);
	double** pull = new double*[total_size];
	for (i=0; i<total_size; i++)
		pull[i] = new double[problem.n_var];

	// Combine elements in RefSet1
	for(i=0; i<problem.b1-1; i++)
		for(j=i+1; j<problem.b1; j++)
		{
			// Combine solutions not combined in the past
			if(problem.iter1[i] > problem.last_combine ||
				problem.iter1[j] > problem.last_combine)
			{
				Combine_inc(problem.refSet1[i], problem.refSet1[j], offsprings, 4);	// Combine_inc  JZ_inc

				for(k=0; k<4; k++)
				{
					for(l=0; l<problem.n_var; l++)
						pull[pull_size][l] = offsprings[k][l];
					pull_size++;  
				}
			}
		}


	// Combine RefSet1 with RefSet2
	for(i=0; i<problem.b1; i++)
		for(j=0; j<problem.b2; j++)
		{
			if (problem.iter1[i] > problem.last_combine ||
				problem.iter2[j] > problem.last_combine)
			{
				Combine_inc(problem.refSet1[i], problem.refSet2[j], offsprings, 3);	// Combine_inc  JZ_inc

				for(k=0; k<3; k++)
				{
					for(l=0; l<problem.n_var; l++)
						pull[pull_size][l] = offsprings[k][l];
					pull_size++;   
				}
			}
		}

	// Combine elements in Refset2
	for(i=0; i<problem.b2-1; i++)
		for(j=i+1; j<problem.b2; j++)
		{
			if (problem.iter2[i] > problem.last_combine ||
				problem.iter2[j] > problem.last_combine)
			{
				Combine_inc(problem.refSet2[i], problem.refSet2[j], offsprings, 2);	// Combine_inc  JZ_inc

				for(k=0; k<2; k++)
				{
					for(l=0; l<problem.n_var; l++)
						pull[pull_size][l] = offsprings[k][l];
					pull_size++;
				}
			}
		}

	// Update, if necessary, Reference Set
	problem.last_combine = problem.iter;
	problem.iter++;

	for(i=0; i<pull_size; i++)
	{
		pull_index = i;

		bool isDuplicate = false;
		for(j=0; j<i; j++)
		{
			bool sameAsCurrent = true;
			for(k=0; k<problem.n_var; k++)
			{
				if (pull[i][k] != pull[j][k])
				{
					sameAsCurrent = false;
					break;
				}
			}

			if (sameAsCurrent)
			{
				isDuplicate = true;
				break;
			}
		}


		if (isDuplicate)
			continue;

/*
		fprintf(m_pDebug, "pull%d", i+1);
		for(j=0; j<problem.n_var; j++)
			fprintf(m_pDebug, "\t%lf", pull[i][j]);
		fprintf(m_pDebug, "\n");
*/

		TryAddRefSet1(pull[i]);
		TryAddRefSet2(pull[i]);
	}

	for (i=0; i<4; i++)
		delete []offsprings[i];
	delete []offsprings;

	for (i=0; i<total_size; i++)
		delete []pull[i];
	delete []pull;
}

void CBMPOptimizer::Combine(double *x, double *y, double **offsprings, int number)
{
	int i;
	double a;

	double* d = new double[problem.n_var];
	for(i=0; i<problem.n_var; i++)
		d[i] = (y[i] - x[i]) / 2;

	double r = rand()%10001/10000.0;

	// Generate C2
	for(i=0; i<problem.n_var; i++)
	{
		offsprings[0][i] = x[i] + r*d[i];
		if (offsprings[0][i] > problem.high[i])
			offsprings[0][i] = problem.high[i];
		if (offsprings[0][i] < problem.low[i])
			offsprings[0][i] = problem.low[i];
	}

	// Generate C1 or C3
	if(number >= 2)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if (a <=0.5)
				offsprings[1][i] = x[i] - r*d[i];
			else
				offsprings[1][i] = y[i] + r*d[i];

			if (offsprings[1][i] > problem.high[i])
				offsprings[1][i] = problem.high[i];
			if (offsprings[1][i] < problem.low[i])
				offsprings[1][i] = problem.low[i];
		}
	}

	// Generate the other one (C1 or C3)
	if(number >= 3)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if (a > 0.5)
				offsprings[2][i] = x[i] - r*d[i];
			else
				offsprings[2][i] = y[i] + r*d[i];

			if (offsprings[2][i] > problem.high[i])
				offsprings[2][i] = problem.high[i];
			if (offsprings[2][i] < problem.low[i])
				offsprings[2][i] = problem.low[i];
		}
	}

	// Generate another C2
	if(number == 4)
	{
		r = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			offsprings[3][i] = x[i] + r*d[i];

			if (offsprings[3][i] > problem.high[i])
				offsprings[3][i] = problem.high[i];
			if (offsprings[3][i] < problem.low[i])
				offsprings[3][i] = problem.low[i];
		}
	}

	delete []d;
}

/*
//JZ_inc
void CBMPOptimizer::Combine_inc(double *x, double *y, double **offsprings, int number)  //JZ_inc
{
	int i;
	double a;

	int* dInc = new int[problem.n_var];
	int* dInc_max = new int[problem.n_var];

	for(i=0; i<problem.n_var; i++)
	{
		dInc_max[i] = int((y[i] - x[i])/(problem.inc[i] * 2));
		// generate random numbers between 0 to dInc_max
		if (dInc_max[i] < 0)
			dInc[i] = -(rand() % (-dInc_max[i]+1));
		else
			dInc[i] = rand() % (dInc_max[i]+1);
	}

	// Generate C2
	for(i=0; i<problem.n_var; i++)
	{
		offsprings[0][i] = x[i] + dInc[i] * problem.inc[i];

		if(offsprings[0][i] > problem.high[i])
			offsprings[0][i] = problem.high[i];
		if(offsprings[0][i] < problem.low[i])
			offsprings[0][i] = problem.low[i];
	}

	// Generate C1 or C3
	if(number >= 2)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if(a <= 0.5)
				offsprings[1][i] = x[i] - dInc[i] * problem.inc[i];
			else
				offsprings[1][i] = y[i] + dInc[i] * problem.inc[i];
			if(offsprings[1][i] > problem.high[i])
				offsprings[1][i] = problem.high[i];
			if(offsprings[1][i] < problem.low[i])
				offsprings[1][i] = problem.low[i];
		}
	}

	// Generate the other one (C1 or C3)
	if(number >= 3)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if(a > 0.5)
				offsprings[2][i] = x[i] - dInc[i] * problem.inc[i];
			else
				offsprings[2][i] = y[i] + dInc[i] * problem.inc[i];
			if(offsprings[2][i] > problem.high[i])
				offsprings[2][i] = problem.high[i];
			if(offsprings[2][i] < problem.low[i])
				offsprings[2][i] = problem.low[i];
		}
	}

	// Generate another C2
	if(number == 4)
	{
	
		for(i=0; i<problem.n_var; i++)
		{
			if (dInc_max[i] < 0)
				dInc[i] = -(rand() % (-dInc_max[i]+1));
			else
				dInc[i] = rand() % (dInc_max[i]+1);

			offsprings[3][i] = x[i] + dInc[i] * problem.inc[i];
			if(offsprings[3][i] > problem.high[i])
				offsprings[3][i] = problem.high[i];
			if(offsprings[3][i] < problem.low[i])
				offsprings[3][i] = problem.low[i];
		}
	}

	delete []dInc;
	delete []dInc_max;
} //JZ_inc
*/

void CBMPOptimizer::Combine_inc(double *x, double *y, double **offsprings, int number)  //JZ_inc
{
	int i;
	double a;

	int* dInc = new int[problem.n_var];
	int* dInc_max = new int[problem.n_var];

	for(i=0; i<problem.n_var; i++)
	{
		dInc_max[i] = int((y[i] - x[i])/(2*problem.inc[i])+(rand()%2*0.5));
		// generate random numbers between 0 to dInc_max
		if (dInc_max[i] < 0)
			dInc[i] = -(rand() % (-dInc_max[i]+1));
		else
			dInc[i] = rand() % (dInc_max[i]+1);
	}

	// Generate C2
	for(i=0; i<problem.n_var; i++)
	{
		offsprings[0][i] = x[i] + dInc[i] * problem.inc[i];

		if(offsprings[0][i] > problem.high[i])
			offsprings[0][i] = problem.high[i];
		if(offsprings[0][i] < problem.low[i])
			offsprings[0][i] = problem.low[i];
	}

	// Generate C1 or C3
	if(number >= 2)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if(a <= 0.5)
				offsprings[1][i] = x[i] - dInc[i] * problem.inc[i];
			else
				offsprings[1][i] = y[i] + dInc[i] * problem.inc[i];
			if(offsprings[1][i] > problem.high[i])
				offsprings[1][i] = problem.high[i];
			if(offsprings[1][i] < problem.low[i])
				offsprings[1][i] = problem.low[i];
		}
	}

	// Generate the other one (C1 or C3)
	if(number >= 3)
	{
		a = rand()%10001/10000.0;
		for(i=0; i<problem.n_var; i++)
		{
			if (dInc_max[i] < 0)
				dInc[i] = -(rand() % (-dInc_max[i]+1));
			else
				dInc[i] = rand() % (dInc_max[i]+1);
			if(a > 0.5)
				offsprings[2][i] = x[i] - dInc[i] * problem.inc[i];
			else
				offsprings[2][i] = y[i] + dInc[i] * problem.inc[i];
			if(offsprings[2][i] > problem.high[i])
				offsprings[2][i] = problem.high[i];
			if(offsprings[2][i] < problem.low[i])
				offsprings[2][i] = problem.low[i];
		}
	}

	// Generate another C2
	if(number == 4)
	{
		for(i=0; i<problem.n_var; i++)
		{
			if (dInc_max[i] < 0)
				dInc[i] = -(rand() % (-dInc_max[i]+1));
			else
				dInc[i] = rand() % (dInc_max[i]+1);

			offsprings[3][i] = x[i] + dInc[i] * problem.inc[i];
			if(offsprings[3][i] > problem.high[i])
				offsprings[3][i] = problem.high[i];
			if(offsprings[3][i] < problem.low[i])
				offsprings[3][i] = problem.low[i];
		}
	}

	delete []dInc;
	delete []dInc_max;
}

void CBMPOptimizer::TryAddRefSet1(double *sol)
{
	int i, j, worst_index;
	double value, worst_value;

	if (!IsNewSolution(problem.refSet1, problem.b1, sol))
		return;

	// turn on counter (06-28-05)
	if (nRunCounter > nMaxRun)
		return;
	
	value = Evaluate(sol);
	
	if (problem.LS)
		ImproveSolution(sol, &value);

	worst_index = problem.order1[problem.b1-1];
	worst_value = problem.value1[worst_index];

	if (value < worst_value)
	{
		i = problem.b1-1;
		while ((i >= 0) && (value < problem.value1[problem.order1[i]]))
			i--;
		i++;

		// Replace solution
		for(j=0; j<problem.n_var; j++)
			problem.refSet1[worst_index][j] = sol[j];

		problem.value1[worst_index] = value;
		problem.iter1[worst_index]  = problem.iter;

		// Update Order
		for(j=problem.b1-1; j>i; j--)
			problem.order1[j] = problem.order1[j-1];
		
		problem.order1[i]    = worst_index;
		problem.new_elements = 1;
	}
}
	
void CBMPOptimizer::TryAddRefSet2(double *sol)
{
	int i, j, worst_index;
	double value, worst_value;

	// It should be noted that solutions are not improved
	// to increase the diversity in RefSet2
	value = DistanceToRefSet(sol);

	worst_index = problem.order2[problem.b2-1];
	worst_value = problem.value2[worst_index];
	
	if(value > worst_value)
	{
		i = problem.b2-1;
		while((i >= 0) && (value > problem.value2[problem.order2[i]]))
			i--;
		i++;

		// Replace solution
		for(j=0; j<problem.n_var; j++)
			problem.refSet2[worst_index][j] = sol[j];

		problem.value2[worst_index] = value;
		problem.iter2[worst_index]  = problem.iter;

		// Update Order
		for(j=problem.b2-1; j>i; j--)
			problem.order2[j] = problem.order2[j-1];
		
		problem.order2[i]    = worst_index;
		problem.new_elements = 1;
	}
}

double CBMPOptimizer::DistanceToRefSet(double *sol)
{
	int i, j;
	double d, min_dist=DBL_MAX;

	for(i=0; i<problem.b1; i++)
	{
		d = 0;
//		for(j=0; j<problem.n_var; j++)
//			d += pow(sol[j] - problem.refSet1[i][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
		for(j=0; j<problem.n_var; j++)
			d += pow((sol[j] - problem.refSet1[i][j])/(problem.high[j] - problem.low[j]), 2);
		if(d == 0.0)
			return 0.0;
		if(min_dist > d)
			min_dist = d;
	}

	for(i=0; i<problem.b2; i++)
	{
		d = 0;
//		for(j=0; j<problem.n_var; j++)
//			d += pow(sol[j] - problem.refSet2[i][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
		for(j=0; j<problem.n_var; j++)
			d += pow((sol[j] - problem.refSet2[i][j])/(problem.high[j] - problem.low[j]), 2);
		if(d == 0.0)
			return 0.0;
		if(min_dist > d)
			min_dist = d;
	}

	return min_dist;
}

void CBMPOptimizer::UpdateRefSet2()
{
	int i, j, k, a, *index2, cont=0;
	double d, dmax;

	problem.iter++;
	problem.digits++;

	double* current    = new double[problem.n_var];
	double* min_dist   = new double[problem.PSize];	
	double* value      = new double[problem.PSize];	
	double** solutions = new double*[problem.PSize];
	for(i=0; i<problem.PSize; i++)
		solutions[i] = new double[problem.n_var];

	for(i=0; i<problem.PSize; i++)
	{
		// Generate new solution
		for(j=0; j<problem.n_var; j++)
			current[j] = GenerateValue(j);

		if(IsNewSolution(solutions, i-1, current))
		{
			// Store solution in matrix "solutions"
			for(j=0; j<problem.n_var; j++)
				solutions[i][j] = current[j];
		}
		else 
		{
			i--;
			cont++;
		}

		if(cont > problem.PSize/2)
		{
			problem.digits++;
			cont = 0;
		}
	}
	
	// Compute minimum distances
	for(i=0; i<problem.PSize; i++)
		min_dist[i] = DistanceToRefSet1(solutions[i]);

	// Add to RefSet
	for(i=0; i<problem.b2; i++)
	{
		// Select the solution with maximum minimum-distance
		dmax = -1;
		for(j=0; j<problem.PSize; j++)
			if(min_dist[j] > dmax)
			{
				dmax = min_dist[j];
				a = j;
			}

		for(j=0; j<problem.n_var; j++)
			problem.refSet2[i][j] = solutions[a][j];
		
		// Update minimum distances
		for(k=0; k<problem.PSize; k++)
		{
			d = 0;
//			for(j=0; j<problem.n_var; j++)
//				d += pow(solutions[k][j] - solutions[a][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
			for(j=0; j<problem.n_var; j++)
				d += pow((solutions[k][j] - solutions[a][j])/(problem.high[j] - problem.low[j]), 2);
			if(d < min_dist[k])
				min_dist[k] = d;
		}
	}

	// Update minimum distances in RefSet2
	for(i=0; i<problem.b2; i++)
	{
	//	for(a=0; a<problem.b2 && a!=i; a++) // correction made on March, 12, 2007
		for(a=i+1; a<problem.b2; a++)
		{
			d = 0;
//			for(j=0; j<problem.n_var; j++)
//				d += pow(problem.refSet2[i][j] - problem.refSet2[a][j], 2);
// Haihong Yang -- March 2nd, 2007
// Commented out above code considering this calculation does not 
// take into consideration of normalization.
// Below is the new version with normalization. Basically, 
// the distance in each dimension will be unitless, and the value range is within 0 and 1
			for(j=0; j<problem.n_var; j++)
				d += pow((problem.refSet2[i][j] - problem.refSet2[a][j])/(problem.high[j] - problem.low[j]), 2);
			if(min_dist[i] > d)
				min_dist[i] = d;
		}
		problem.value2[i] = min_dist[i];
	}

	index2 = new int[problem.b2];
	GetOrderIndices(index2, problem.value2, problem.b2, 1);
	for(i=0; i<problem.b2; i++)
	{
		problem.order2[i] = index2[i];
		problem.iter2[i]  = problem.iter;
	}
	problem.new_elements = 1;

	delete []current;
	delete []min_dist;
	delete []value;
	for(i=0; i<problem.PSize; i++)
		delete []solutions[i];
	delete []solutions;
	delete []index2;
}

void CBMPOptimizer::TryAddEvaluation(double *sol, double current_value)
{
	int i, j;
	int worst_index = problem.PSize;
	bool isSameSolution;

	for(i=0; i<problem.PSize; i++)
	{
		if(problem.evaOrders[i] == -1)
		{
			worst_index = i;
			break;
		}

		// if this is an existing solution, simply return considering
		// this is in the same iteration for the same target
		// (the calculated value will remain the same)
		isSameSolution = true;
		for(j=0; j<problem.n_var; j++)
		{
			if(sol[j] != problem.evaSolutions[i][j])
			{
				isSameSolution = false;
				break;
			}
		}

		if (isSameSolution)
			return;
	}

	int actual_index = 0;
	if (worst_index == problem.PSize)
	{
		// If the current value is better than the worst one in the cached array,
		// replace the worst one.
		if (current_value < problem.evaValues[problem.evaOrders[problem.PSize-1]])
		{
			worst_index = problem.PSize-1;
			actual_index = problem.evaOrders[problem.PSize-1];
		}
		// Otherwise, if the current value is worse than the worst one in the 
		// cached array, simply return.
		else
			return;
	}
	else
	{
		actual_index = worst_index;
	}

	// Populate the appropriate cached solution with the current solution
	// and the evaluated value
	for(j=0; j<problem.n_var; j++)
		problem.evaSolutions[actual_index][j] = sol[j];
	problem.evaValues[actual_index] = current_value;


	int insert_index = 0;
	for(i=0; i<worst_index; i++)
	{
		if(problem.evaValues[problem.evaOrders[i]] > current_value)
		{
			for(j=worst_index-1; j>=i; j--)
				problem.evaOrders[j+1] = problem.evaOrders[j];
			break;
		}
		insert_index++;
	}

	problem.evaOrders[insert_index] = actual_index;
}

void CBMPOptimizer::PerformSearch()
{
	int i, j;
	for(i=0; i<nMaxIter; i++)
	{
		if(problem.new_elements)
		{
			CombineRefSet();
		}
		else
		{	
			UpdateRefSet2();
			CombineRefSet();
		}

		// Rerun the best solution in optimization mode for calculating cost and value
		for(j=0; j<problem.n_var; j++)
			*m_pVariables[j] = problem.refSet1[problem.order1[0]][j];
		m_pBMPRunner->RunModel(RUN_OPTIMIZE);
		
		if(m_pBMPRunner->pWndProgress->Cancelled())
			break;

		double curResult = 0.0;
		double curValue = 0.0;
		double curDelta = 0.0;
		double curValueDelta = 0.0;

		if (m_pBMPRunner->pBMPData->nRunOption == OPTION_MIMIMIZE_COST) // Get total cost of the best solution for Minimize Cost option or TradeOff Curve option
		{
			POSITION pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();

			while (pos != NULL)
			{
				CBMPSite* pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
				curResult += pBMPSite->m_lfCost;
			}
			curValue = problem.value1[problem.order1[0]];
		}
		else if (m_pBMPRunner->pBMPData->nRunOption == OPTION_MAXIMIZE_CONTROL) // Get value of the best solution for Maximize Control option
			curResult = problem.value1[problem.order1[0]];
		else if (m_pBMPRunner->pBMPData->nRunOption == OPTION_TRADE_OFF_CURVE) // Continue if running for tradeoff curve
		{
			POSITION pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();

			while (pos != NULL)
			{
				CBMPSite* pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
				curResult += pBMPSite->m_lfCost;
			}
			curValue = problem.value1[problem.order1[0]];
		}

		if (i == 0) // if first PSize time run, just initialize the Previous Result variable
		{
			m_lfPrevResult = curResult;
			m_lfPrevValue = curValue;
		}
		else // otherwise, compare with the stop delta
		{
			if (m_pBMPRunner->pBMPData->nRunOption == OPTION_MIMIMIZE_COST)
			{
				curDelta = fabs(m_lfPrevResult - curResult);
				curValueDelta = fabs(((m_lfPrevValue-m_lfPrevResult) - (curValue-curResult))/(m_lfPrevValue-m_lfPrevResult));

				if (curValueDelta < 0.001 && curDelta < m_pBMPRunner->pBMPData->lfStopDelta)
				{
					CString strMsg;
					strMsg.Format("Cost of the best solution has been reduced by $%.1lf. The cost reduction is within the stopping delta range. Do you want to continue?", curDelta);
					if (AfxMessageBox(strMsg, MB_YESNO|MB_ICONINFORMATION) != IDYES)
						return;
				}
			}
			else if (m_pBMPRunner->pBMPData->nRunOption == OPTION_MAXIMIZE_CONTROL)
			{
				if (m_lfPrevResult > 0)
				{
//					curDelta = (m_lfPrevResult-curResult)*100/m_lfPrevResult;
					curDelta = fabs((m_lfPrevResult-curResult)*100/m_lfPrevResult);

					if (curDelta < m_pBMPRunner->pBMPData->lfStopDelta)
					{
						CString strMsg;
						strMsg.Format("The control benefit has been improved by %.1lf%. The benefit improvement is within the stopping delta range. Do you want to continue?", curDelta);
						if (AfxMessageBox(strMsg, MB_YESNO|MB_ICONINFORMATION) != IDYES)
							return;
					}
				}
			}
			else if (m_pBMPRunner->pBMPData->nRunOption == OPTION_TRADE_OFF_CURVE)
			{
//				curDelta = m_lfPrevResult - curResult;
//				curValueDelta = ((m_lfPrevValue-m_lfPrevResult) - (curValue-curResult))/(m_lfPrevValue-m_lfPrevResult);
				curDelta = fabs(m_lfPrevResult - curResult);
				curValueDelta = fabs(((m_lfPrevValue-m_lfPrevResult) - (curValue-curResult))/(m_lfPrevValue-m_lfPrevResult));

				if (curValueDelta < 0.001 && curDelta < m_pBMPRunner->pBMPData->lfStopDelta)
				{
					CString strMsg;
					strMsg.Format("Cost of the best solution has been reduced by $%.1lf. The cost reduction is within the stopping delta range. Do you want to continue?", curDelta);
					if (AfxMessageBox(strMsg, MB_YESNO|MB_ICONINFORMATION) != IDYES)
						return;
				}
			}

			m_lfPrevResult = curResult;
			m_lfPrevValue = curValue;
		}
	}
}

void CBMPOptimizer::OutputBestSolutions()
{
	int i, j;
	if (m_pBMPRunner == NULL)
		return;

	if (m_pBMPRunner->pBMPData == NULL)
		return;

	FILE *fp = NULL;
	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	CString strLine, strValue;
	double totalCost;

	double totalSurfaceArea = 0.0;
	double totalExcavatnVol = 0.0;
	double totalSurfStorVol = 0.0;
	double totalSoilStorVol = 0.0;
	double totalUdrnStorVol = 0.0;

	CString	strFilePath = m_pBMPRunner->pBMPData->strOutputDir + "\\BestSolutions.out";
	fp = fopen(LPCSTR(strFilePath), "wt");
	if(fp == NULL)
		return;

	int nEF = 0;
	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
		nEF += pBMPSite->m_factorList.GetCount();
	}

	strValue.Format("%d\t%d\t SS - Best solutions",nEF,m_pBMPRunner->pBMPData->nBMPtype);
	OutputFileHeader(strValue, fp);

	for(i=0; i<m_pBMPRunner->pBMPData->nSolution;i++)
	{
		// mapping solution to ajustable variables
		for(j=0; j<problem.n_var; j++)
			*m_pVariables[j] = problem.refSet1[problem.order1[i]][j];

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
			totalSurfaceArea += pBMPSite->m_lfSurfaceArea;
			totalExcavatnVol += pBMPSite->m_lfExcavatnVol;
			totalSurfStorVol += pBMPSite->m_lfSurfStorVol;
			totalSoilStorVol += pBMPSite->m_lfSoilStorVol;
			totalUdrnStorVol += pBMPSite->m_lfUdrnStorVol;
		}

		//strLine.Format("%d\t%lf\t%lf", i+1, totalCost, problem.value1[problem.order1[i]]);
		strLine.Format("%d\t%lf\t%lf\t%lf\t%lf\t%lf\t%lf", i+1, totalCost, 
					totalSurfaceArea, totalExcavatnVol, totalSurfStorVol, 
					totalSoilStorVol, totalUdrnStorVol);

		pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
			pos1 = pBMPSite->m_factorList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
				strValue.Format("\t%lf", pEF->m_lfCurrent);      //may need to modify JZ 3/21/05
				strLine += strValue;
			}
		}

		//add cost for each unique bmp type
		for(j=0; j<m_pBMPRunner->pBMPData->nBMPtype; j++)
		{
			strValue.Format("\t%lf", m_pBMPRunner->pBMPData->m_pBMPcost[j].m_lfCost);
			strLine += strValue;
		}

		for(j=0; j<problem.n_var; j++)
		{
			strValue.Format("\t%lf", problem.refSet1[problem.order1[i]][j]);
			strLine += strValue;
		}

		strLine += "\n";
		fputs(strLine, fp);
	}
	
	fclose(fp);
}

void CBMPOptimizer::OutputBestSolutionsForTradeOffCurve(int breakNum, FILE* fp)
{
	int i, j;

	if (m_pBMPRunner == NULL)
		return;

	if (m_pBMPRunner->pBMPData == NULL)
		return;

	if (fp == NULL)
		return;

	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	EVALUATION_FACTOR* pEF;
	CString strLine, strValue;
	double totalCost;

	for(i=0; i<m_pBMPRunner->pBMPData->nSolution;i++)
	{

		// mapping solution to ajustable variables
		for(j=0; j<problem.n_var; j++)
			*m_pVariables[j] = problem.refSet1[problem.order1[i]][j];

		strValue.Format("Break%d_Solution%d", breakNum+1, i+1);
		if (!m_pBMPRunner->pBMPData->OpenOutputFiles(strValue))	// time series for the best solution
			return;
//		if (!m_pBMPRunner->OpenOutputFiles(strValue, m_pBMPRunner->pBMPData->nRunOption, RUN_OUTPUT))
//			return;

		m_pBMPRunner->RunModel(RUN_OUTPUT);

		if (!m_pBMPRunner->pBMPData->CloseOutputFiles())
			return;
//		if (!m_pBMPRunner->CloseOutputFiles())
//			return;

		totalCost = 0.0;
		pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
			totalCost += pBMPSite->m_lfCost;
		}

		strLine.Format("%d\t%d", breakNum+1, i+1);
		strValue.Format("\t%lf", totalCost);
		strLine += strValue;

		pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);
			pos1 = pBMPSite->m_factorList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
				//strValue.Format("\t%lf\t%lf", pEF->m_lfTarget, pEF->m_lfCurrent);
				strValue.Format("\t%lf", pEF->m_lfCurrent);
				strLine += strValue;
			}
		}
		
//		strValue.Format("\t%lf", totalCost);
//		strLine += strValue;
		for(j=0; j<problem.n_var; j++)
		{
			strValue.Format("\t%lf", problem.refSet1[problem.order1[i]][j]);
			strLine += strValue;
		}

		//strValue.Format("\t%lf", problem.value1[problem.order1[i]]);
		//strLine += strValue;
		strLine += "\n";
		fputs(strLine, fp);
	}

	fflush(fp);
}

void CBMPOptimizer::OutputDebugFileHeader(FILE* fp)
{
	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	ADJUSTABLE_PARAM* pAP;
	CString strLine, strValue;
	strLine = "Debug Information\n\n";
	fputs(strLine, fp);

	strLine = "Solution #";
	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);

		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			strValue.Format("\tSite%s_%s", pBMPSite->m_strID, pAP->m_strVariable);
			strLine += strValue;
		}
	}

	strLine += "\n\n";
	fputs(strLine, fp);
	fflush(fp);
}

void CBMPOptimizer::OutputDebugInformation(FILE* fp)
{
	int i, j;
	CString strLine, strValue;

	for(i=0; i<problem.b1; i++)
	{
		strLine.Format("RefSet1_%d", i+1);
		for(j=0; j<problem.n_var; j++)
		{
			strValue.Format("\t%lf", problem.refSet1[problem.order1[i]][j]);
	strLine += strValue;
		}
		strLine += "\n";
		fputs(strLine, fp);
	}

	for(i=0; i<problem.b2; i++)
	{
		strLine.Format("RefSet2_%d", i+1);
		for(j=0; j<problem.n_var; j++)
		{
			strValue.Format("\t%lf", problem.refSet2[problem.order2[i]][j]);
			strLine += strValue;
		}
	strLine += "\n";
	fputs(strLine, fp);
	}

	fflush(fp);
}

void CBMPOptimizer::OutputFileHeader(CString header, FILE* fp)
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

void CBMPOptimizer::OutputFileHeaderForTradeOffCurve(FILE* fp)
{
	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	ADJUSTABLE_PARAM* pAP;
	EVALUATION_FACTOR* pEF;
	CString strLine, strValue;
	strLine = "SS - Cost-Effectiveness Curve Solutions\n";
	strLine += "TargetBreak#\tSolution#\tCost($)";
//	strLine += "\tTargetValue";

	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
			strValue.Format("\tSite%s_%s_%d", pBMPSite->m_strID, pEF->m_strFactor, pEF->m_nCalcMode);
			strLine += strValue;
		}
	}

//	strLine += "\tCost($)";

	pos = m_pBMPRunner->pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) m_pBMPRunner->pBMPData->routeList.GetNext(pos);

		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			strValue.Format("\tSite%s_%s", pBMPSite->m_strID, pAP->m_strVariable);
			strLine += strValue;
		}
	}

//	strLine += "\tValue";

	strLine += "\n\n";
	fputs(strLine, fp);
	fflush(fp);
}

