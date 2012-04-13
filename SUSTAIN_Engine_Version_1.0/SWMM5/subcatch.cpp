//-----------------------------------------------------------------------------
//   subcatch.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//             7/5/06   (Build 5.0.008)
//             9/19/06   (Build 5.0.009)
//   Author:   L. Rossman
//
//   Subcatchment runoff & quality functions.
//-----------------------------------------------------------------------------

#include <math.h>
#include <string.h>
#include "headers.h"
#include "odesolve.h"

//-----------------------------------------------------------------------------
// Constants 
//-----------------------------------------------------------------------------
const float MCOEFF    = 1.49;               // constant in Manning Eq.
const float MEXP      = 5./3.;              // exponent in Manning Eq.
const float ODETOL    = 0.0001;             // acceptable error for ODE solver

//-----------------------------------------------------------------------------
// Shared variables   
//-----------------------------------------------------------------------------
static  float     Losses;         // subcatch evap. + infil. loss rate (ft/sec)
static  float     Outflow;        // subcatch outflow rate (ft/sec)
static  float     Vevap;          // subcatch evap. volume over a time step (ft)
static  float     Vinfil;         // subcatch infil. volume over a time step (ft)
static  float     Voutflow;       // subcatch outflow volume over a time step (ft)

//////////////////////////////////////
//  New variable added. (LR - 7/5/06 )
//////////////////////////////////////
static  float     Vponded;        // subcatch ponded volume (ft)

static  TSubarea* theSubarea;     // subarea to which getDdDt() is applied
static  char *RunoffRoutingWords[] = { w_OUTLET,  w_IMPERV, w_PERV, NULL};

//-----------------------------------------------------------------------------
//  Imported variables (declared in RUNOFF.C)
//-----------------------------------------------------------------------------
/////////////////////////////////////////////////////////////
//  The following variables have been renamed. (LR - 7/5/06 )
/////////////////////////////////////////////////////////////
extern  float*    WashoffQual;    // washoff quality for a subcatchment (mass/ft3)
extern  float*    WashoffLoad;    // washoff loads for a impervious landuse (mass/sec)
extern  float*    RemovalLoad;    // removal loads for a pervious landuse (mass/sec)

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)   
//-----------------------------------------------------------------------------
//  subcatch_readParams        (called from parseLine in input.c)
//  subcatch_readSubareaParams (called from parseLine in input.c)
//  subcatch_readLanduseParams (called from parseLine in input.c)
//  subcatch_readInitBuildup   (called from parseLine in input.c)
//  subcatch_validate          (called from project_validate)
//  subcatch_initState         (called from project_init)
//  subcatch_setOldState       (called from runoff_execute)
//  subcatch_getRunon          (called from runoff_execute)
//  subcatch_getRunoff         (called from runoff_execute)
//  subcatch_getWashoff        (called from runoff_execute)
//  subcatch_getBuildup        (called from runoff_execute)
//  subcatch_sweepBuildup      (called from runoff_execute)
//  subcatch_hadRunoff         (called from runoff_execute)
//  subcatch_getWtdOutflow     (called from addWetWeatherInflows in routing.c)
//  subcatch_getWtdWashoff     (called from addWetWeatherInflows in routing.c)
//  subcatch_getResults        (called from output_saveSubcatchResults)

//-----------------------------------------------------------------------------
// Function declarations
//-----------------------------------------------------------------------------
////////////////////////////////////
//  Function removed. (LR - 7/5/06 )
////////////////////////////////////
//static void  getWashoffLoads(int subcatch, float qRunon, float qInflow,
//             float qRunoff, float tStep);

static char  sweptSurfacesDry(int subcatch);
static void  getSubareaRunoff(int subcatch, int subarea, float rainfall,
             float evap, float tStep);
static void  updatePondedDepth(TSubarea* subarea, float* tx);
static void  getDdDt(float t, float* d, float* dddt);

///////////////////////////////////////
//  New functions added. (LR - 7/5/06 )
///////////////////////////////////////
static void  getPondedQual(float wUp[], float qUp, float qPpt, float qEvap,
             float qInfil, float v, float area, float tStep,
             float pondedQual[]);
//static float getCstrQual(float c, float v, float wIn, float qNet, float tStep);
static void  getWashoffQual(int j, float runoff, float tStep,float washoffQual[]);
static void  combineWashoffQual(int j, float pondedQual[], float washoffQual[],
             float tStep);
static float getBmpRemoval(int j, int p);


//=============================================================================

int  subcatch_readParams(int j, char* tok[], int ntoks)
//
//  Input:   j = subcatchment index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads subcatchment parameters from a tokenized  line of input data.
//
//  Data has format:
//    Name  RainGage  Outlet  Area  %Imperv  Width  Slope CurbLength  Snowmelt  
//
{
    int   i, k, m;
    char* id;
    float x[9];

    // --- check for enough tokens
    if ( ntoks < 8 ) return error_setInpError(ERR_ITEMS, "");

    // --- check that named subcatch exists
    id = project_findID(SUBCATCH, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);

    // --- check that rain gage exists
    k = project_findObject(GAGE, tok[1]);
    if ( k < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    x[0] = k;

    // --- check that outlet node or subcatch exists
    m = project_findObject(NODE, tok[2]);
    x[1] = m;
    m = project_findObject(SUBCATCH, tok[2]);
    x[2] = m;
    if ( x[1] < 0.0 && x[2] < 0.0 )
        return error_setInpError(ERR_NAME, tok[2]);

    // --- read area, %imperv, width, slope, & curb length
    for ( i = 3; i < 8; i++)
    {
        if ( ! getFloat(tok[i], &x[i]) || x[i] < 0.0 )
            return error_setInpError(ERR_NUMBER, tok[i]);
    }

    // --- if snowmelt object named, check that it exists
    x[8] = -1;
    if ( ntoks > 8 )
    {
        k = project_findObject(SNOWMELT, tok[8]);
        if ( k < 0 ) return error_setInpError(ERR_NAME, tok[8]);
        x[8] = k;
    }

    // --- assign input values to subcatch's properties
    Subcatch[j].ID = id;
    Subcatch[j].gage       = x[0];
    Subcatch[j].outNode    = x[1];
    Subcatch[j].outSubcatch= x[2];
    Subcatch[j].area       = x[3] / UCF(LANDAREA);
    Subcatch[j].fracImperv = x[4] / 100.0;
    Subcatch[j].width      = x[5] / UCF(LENGTH);
    Subcatch[j].slope      = x[6] / 100.0;
    Subcatch[j].curbLength = x[7];

    // --- create the snow pack object if it hasn't already been created
    if ( x[8] >= 0 )
    {
        if ( !snow_createSnowpack(j, x[8]) )
            return error_setInpError(ERR_MEMORY, "");
    }
    return 0;
}

//=============================================================================

int subcatch_readSubareaParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads subcatchment's subarea parameters from a tokenized 
//           line of input data.
//
//  Data has format:
//    Subcatch  Imperv_N  Perv_N  Imperv_S  Perv_S  PctZero  RouteTo (PctRouted)
//
{
    int   i, j, k, m;
    float x[7];

    // --- check for enough tokens
    if ( ntoks < 7 ) return error_setInpError(ERR_ITEMS, "");

    // --- check that named subcatch exists
    j = project_findObject(SUBCATCH, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

    // --- read in Mannings n, depression storage, & PctZero values
    for (i = 0; i < 5; i++)
    {
        if ( ! getFloat(tok[i+1], &x[i])  || x[i] < 0.0 )
            return error_setInpError(ERR_NAME, tok[i+1]);
    }

    // --- check for valid runoff routing keyword
    m = findmatch(tok[6], RunoffRoutingWords);
    if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[6]);

    // --- get percent routed parameter if present (default is 100)
    x[5] = m;
    x[6] = 1.0;
    if ( ntoks >= 8 )
    {
        if ( ! getFloat(tok[7], &x[6]) || x[6] < 0.0 || x[6] > 100.0 )
            return error_setInpError(ERR_NUMBER, tok[7]);
        x[6] /= 100.0;
    }

    // --- assign input values to each type of subarea
    Subcatch[j].subArea[IMPERV0].N = x[0];
    Subcatch[j].subArea[IMPERV1].N = x[0];
    Subcatch[j].subArea[PERV].N    = x[1];

    Subcatch[j].subArea[IMPERV0].dStore = 0.0;
    Subcatch[j].subArea[IMPERV1].dStore = x[2] / UCF(RAINDEPTH);
    Subcatch[j].subArea[PERV].dStore    = x[3] / UCF(RAINDEPTH);

    Subcatch[j].subArea[IMPERV0].fArea  = Subcatch[j].fracImperv * x[4] / 100.0;
    Subcatch[j].subArea[IMPERV1].fArea  = Subcatch[j].fracImperv * (1.0 - x[4] / 100.0);
    Subcatch[j].subArea[PERV].fArea     = (1.0 - Subcatch[j].fracImperv);

    // --- assume that all runoff from each subarea goes to subcatch outlet
    for (i = IMPERV0; i <= PERV; i++)
    {
        Subcatch[j].subArea[i].routeTo = TO_OUTLET;
        Subcatch[j].subArea[i].fOutlet = 1.0;
    }

    // --- modify routing if pervious runoff routed to impervious area
    //     (fOutlet is the fraction of runoff not routed)
    k = x[5];
    if ( k == TO_IMPERV )
    {
        Subcatch[j].subArea[PERV].routeTo = k;
        Subcatch[j].subArea[PERV].fOutlet = 1.0 - x[6];
    }

    // --- modify routing if impervious runoff routed to pervious area
    if ( k == TO_PERV )
    {
        Subcatch[j].subArea[IMPERV0].routeTo = k;
        Subcatch[j].subArea[IMPERV1].routeTo = k;
        Subcatch[j].subArea[IMPERV0].fOutlet = 1.0 - x[6];
        Subcatch[j].subArea[IMPERV1].fOutlet = 1.0 - x[6];
    }
    return 0;
}

//=============================================================================

int subcatch_readLanduseParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads assignment of landuses to subcatchment from a tokenized 
//           line of input data.
//
//  Data has format:
//    Subcatch  landuse  percent .... landuse  percent
//
{
    int    j, k, m;
    float  f;

    // --- check for enough tokens
    if ( ntoks < 3 ) return error_setInpError(ERR_ITEMS, "");

    // --- check that named subcatch exists
    j = project_findObject(SUBCATCH, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

    // --- process each pair of landuse - percent items
    for ( k = 2; k <= ntoks; k = k+2)
    {
        // --- check that named land use exists and is followed by a percent
        m = project_findObject(LANDUSE, tok[k-1]);
        if ( m < 0 ) return error_setInpError(ERR_NAME, tok[k-1]);
        if ( k+1 > ntoks ) return error_setInpError(ERR_ITEMS, "");
        if ( ! getFloat(tok[k], &f) )
            return error_setInpError(ERR_NUMBER, tok[k]);

        // --- store land use fraction in subcatch's landFactor property
        Subcatch[j].landFactor[m].fraction = f/100.0;
    }
    return 0;
}

//=============================================================================

int subcatch_readInitBuildup(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads initial pollutant buildup on subcatchment from 
//           tokenized line of input data.
//
//  Data has format:
//    Subcatch  pollut  initLoad .... pollut  initLoad
//
{
    int    j, k, m;
    float  x;

    // --- check for enough tokens
    if ( ntoks < 3 ) return error_setInpError(ERR_ITEMS, "");

    // --- check that named subcatch exists
    j = project_findObject(SUBCATCH, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

    // --- process each pair of pollutant - init. load items
    for ( k = 2; k <= ntoks; k = k+2)
    {
        // --- check for valid pollutant name and loading value
		//added
		if(strncmp(tok[k-1],strTSS,MAXFNAME) == 0) 
			tok[k-1] = "SAND";

        m = project_findObject(POLLUT, tok[k-1]);
        if ( m < 0 ) return error_setInpError(ERR_NAME, tok[k-1]);
        if ( k+1 > ntoks ) return error_setInpError(ERR_ITEMS, "");
        if ( ! getFloat(tok[k], &x) )
            return error_setInpError(ERR_NUMBER, tok[k]);

        // --- store loading in subcatch's initBuildup property
        Subcatch[j].initBuildup[m] = x;
    }
    return 0;
}

//=============================================================================

void  subcatch_validate(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: checks for valid subcatchment input parameters.
//
{
    int    i;
    float  area;

    // --- compute alpha (i.e. WCON in old SWMM) for overland flow
    //     NOTE: the area which contributes to alpha for both imperv
    //     subareas w/ and w/o depression storage is the total imperv area.
    for (i = IMPERV0; i <= PERV; i++)
    {
        if ( i == PERV )
        {
            area = (1.0 - Subcatch[j].fracImperv) * Subcatch[j].area;
        }
        else
        {
             area = Subcatch[j].fracImperv * Subcatch[j].area;
        }
        Subcatch[j].subArea[i].alpha = 0.0;
        if ( area > 0.0 && Subcatch[j].subArea[i].N > 0.0 )
        {
            Subcatch[j].subArea[i].alpha = MCOEFF * Subcatch[j].width / area *
                sqrt(Subcatch[j].slope) / Subcatch[j].subArea[i].N;
        }
    }
}

//=============================================================================

void  subcatch_initState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: Initializes the state of a subcatchment.
//
{
    int   i;
    int   p;                           // pollutant index
    float f;                           // fraction of total area
    float area;                        // area (ft2 or acre or ha)
    float curb;                        // curb length (users units)
    float startDrySeconds;             // antecedent dry period (sec)
    float buildup;                     // initial mass buildup (lbs or kg)
	//added
    float iarea;                       // impervious area (ft2 or acre or ha)
    float icurb;                       // curb length on impervious area (users units)
    float parea;                       // pervious area (ft2 or acre or ha)
    float detstorage;                  // initial sediment detached (lbs or kg)
    float rainfall;					   // rainfall (ft/sec)

    // --- initialize rainfall, runoff, & snow depth
    Subcatch[j].rainfall = 0.0;
    Subcatch[j].oldRunoff = 0.0;
    Subcatch[j].newRunoff = 0.0;
    Subcatch[j].oldSnowDepth = 0.0;
    Subcatch[j].newSnowDepth = 0.0;
    Subcatch[j].runon = 0.0;

    // --- set isUsed property of subcatchment's rain gage
    i = Subcatch[j].gage;
    if ( i >= 0 )
    {
        Gage[i].isUsed = TRUE;
        if ( Gage[i].coGage >= 0 ) Gage[Gage[i].coGage].isUsed = TRUE;
    }

    // --- initialize state of infiltration, groundwater, & snow pack objects
    if ( Subcatch[j].infil == j )  infil_initState(j, InfilModel);
    if ( Subcatch[j].groundwater ) gwater_initState(j);
    if ( Subcatch[j].snowpack )    snow_initSnowpack(j);

    // --- initialize state of sub-areas
    for (i = IMPERV0; i <= PERV; i++)
    {
        Subcatch[j].subArea[i].depth  = 0.0;
        Subcatch[j].subArea[i].inflow = 0.0;
        Subcatch[j].subArea[i].runoff = 0.0;
    }

    // --- initialize runoff quality
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        Subcatch[j].oldQual[p] = 0.0;
        Subcatch[j].newQual[p] = 0.0;
/////////////////////////////////////////////////////////////////
//  New property added to support washoff routing. (LR - 7/5/06 )
/////////////////////////////////////////////////////////////////
        Subcatch[j].pondedQual[p] = 0.0;
    }

    // --- initialize pollutant buildup

    // --- first convert antecedent dry days into seconds
    startDrySeconds = StartDryDays*SECperDAY;

    // --- then examine each land use
    for (i = 0; i < Nobjects[LANDUSE]; i++)
    {
        // --- initialize date when last swept
        Subcatch[j].landFactor[i].lastSwept =
            datetime_addSeconds(StartDateTime, -Landuse[i].sweepDays0*SECperDAY);

        // --- determine area and curb length covered by land use
        f = Subcatch[j].landFactor[i].fraction;
        area = f * Subcatch[j].area * UCF(LANDAREA);
        curb = f * Subcatch[j].curbLength;

        // --- determine area and curb length covered by land use
        iarea = area * Landuse[i].pctimp;
        icurb = curb * Landuse[i].pctimp;
        parea = area * (1 - Landuse[i].pctimp);

        // --- examine each pollutant
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            // --- if an initial loading was supplied, then use it to
            //     find the starting buildup over the land use
            buildup = 0.0;
			 
            if ( Subcatch[j].initBuildup[p] > 0.0 )
            {
				if (Pollut[p].sedflag > 0)	// the pollutant is sediment
		            buildup = Subcatch[j].initBuildup[p] * iarea;
				else
		            buildup = Subcatch[j].initBuildup[p] * area;
            }

            // --- otherwise use the land use's buildup function to 
            //     compute a buildup over the antecedent dry period
            else 
			{
				if (Pollut[p].sedflag > 0)	// the pollutant is sediment
					buildup = landuse_getBuildup(i, p, iarea, icurb, buildup,
							   startDrySeconds);
				else
					buildup = landuse_getBuildup(i, p, area, curb, buildup,
							   startDrySeconds);
			}

            // find the starting detached soil over the pervious land use
			detstorage = 0.0;

			if (Pollut[p].sedflag > 0)	// the pollutant is sediment
			{
				// --- use land use's detach function for pervious area
				rainfall = Subcatch[j].rainfall;	//ft/sec

				detstorage = landuse_getDetached(i, area, detstorage, rainfall, 
					startDrySeconds);
			}

            Subcatch[j].landFactor[i].buildup[p] = buildup;
			Subcatch[j].landFactor[i].detstorage[p] = detstorage;
        }
    }
}

//=============================================================================

void subcatch_setOldState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: replaces old state of subcatchment with new state.
//
{
    int i;
    Subcatch[j].oldRunoff = Subcatch[j].newRunoff;
    Subcatch[j].oldSnowDepth = Subcatch[j].newSnowDepth;
    Subcatch[j].runon = 0.0;
    for (i = IMPERV0; i <= PERV; i++)
    {
        Subcatch[j].subArea[i].inflow = 0.0;
    }
    for (i = 0; i < Nobjects[POLLUT]; i++)
    {
        Subcatch[j].oldQual[i] = Subcatch[j].newQual[i];
        Subcatch[j].newQual[i] = 0.0;
    }
}

//=============================================================================

void subcatch_getRunon(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: Routes runoff from a subcatchment to its outlet subcatchment
//           or between its subareas.
//
{
    int   i;                           // subarea index
    int   k;                           // outlet subcatchment index
    float q;                           // runon to outlet subcatchment (ft/sec)
    float q1, q2;                      // runoff from imperv. areas (ft/sec)

    // --- add previous period's runoff from this subcatchment to the
    //     runon of the outflow subcatchment, if it exists
    k = Subcatch[j].outSubcatch;
    if ( k >= 0 && k != j && Subcatch[k].area > 0.0 )
    {
        // --- distribute previous runoff from subcatch j (in cfs)
        //     uniformly over area of subcatch k (ft/sec)
        q = Subcatch[j].oldRunoff / Subcatch[k].area;
        Subcatch[k].runon += q;

        // --- assign this flow to the 3 types of subareas
        for (i = IMPERV0; i <= PERV; i++)
        {
            Subcatch[k].subArea[i].inflow += q;
        }

        // --- add runoff mass load (in mass/sec) to receiving subcatch,
        //     storing it in Subcatch[].newQual for now
        for (i = 0; i < Nobjects[POLLUT]; i++)
        {
            Subcatch[k].newQual[i] += (Subcatch[j].oldRunoff *
                                       Subcatch[j].oldQual[i] * LperFT3);
        }
    }

    // --- add to sub-area inflow any outflow from other subarea in previous period
    //     (NOTE: no transfer of runoff pollutant load, since runoff loads are
    //     based on runoff flow from entire subcatchment.)

    // --- Case 1: imperv --> perv
    if ( Subcatch[j].fracImperv < 1.0 &&
         Subcatch[j].subArea[IMPERV0].routeTo == TO_PERV )
    {
        // --- add area-wtd. outflow from imperv1 subarea to perv area inflow
        q1 = Subcatch[j].subArea[IMPERV0].runoff *
             Subcatch[j].subArea[IMPERV0].fArea;
        q2 = Subcatch[j].subArea[IMPERV1].runoff *
             Subcatch[j].subArea[IMPERV1].fArea;
        Subcatch[j].subArea[PERV].inflow += (q1 + q2) *
             (1.0 - Subcatch[j].subArea[IMPERV0].fOutlet) /
             Subcatch[j].subArea[PERV].fArea;
    }

    // --- Case 2: perv --> imperv
    if ( Subcatch[j].fracImperv > 0.0 &&
         Subcatch[j].subArea[PERV].routeTo == TO_IMPERV &&
         Subcatch[j].subArea[IMPERV1].fArea > 0.0 )
    {
        Subcatch[j].subArea[IMPERV1].inflow +=
            Subcatch[j].subArea[PERV].runoff * 
            (1.0 - Subcatch[j].subArea[PERV].fOutlet) *
            Subcatch[j].subArea[PERV].fArea /
            Subcatch[j].subArea[IMPERV1].fArea;
    }
}

//=============================================================================

float subcatch_getRunoff(int j, float tStep)
//
//  Input:   j = subcatchment index
//           tStep = time step (sec)
//  Output:  returns total runoff produced by subcatchment (ft/sec)
//  Purpose: Computes runoff & new storage depth for subcatchment.
//
{
    int   i;                           // subarea index
    int   k;                           // rain gage index
    float rainfall = 0.0;              // rainfall (ft/sec)
    float snowfall = 0.0;              // snowfall (ft/sec)
    float rainVol;                     // rain volume (ft)
    float evapVol    = 0.0;            // evaporation volume (ft)
    float infilVol   = 0.0;            // infiltration volume (ft)
    float outflowVol = 0.0;            // runoff volume leaving subcatch (ft)
    float outflow;                     // runoff rate leaving subcatch (cfs)
    float runoff;                      // total runoff rate on subcatch (ft/sec)
    float area;                        // total subcatch area (ft2)
    float fArea;                       // subarea fraction of total area
    float netPrecip[3];                // subarea net precipitation (ft/sec)

///////////////////////////////////////////////////////////
////  Added to support washoff calculations. (LR - 7/5/06 )
///////////////////////////////////////////////////////////
    // --- save current depth of ponded water
    Vponded = subcatch_getDepth(j);

    // --- get current rainfall or snowfall from rain gage (in ft/sec)
    k = Subcatch[j].gage;
    if ( k >= 0 )
    {
        gage_getPrecip(k, &rainfall, &snowfall);
    }

    // --- assign total precip. rate to subcatch's rainfall property
    Subcatch[j].rainfall = rainfall + snowfall;

    // --- determine net precipitation input (netPrecip) to each sub-area

    // --- if subcatch has a snowpack, then base netPrecip on possible snow melt
    if ( Subcatch[j].snowpack )
    {
        Subcatch[j].newSnowDepth = 
            snow_getSnowMelt(j, rainfall, snowfall, tStep, netPrecip);
    }

    // --- otherwise netPrecip is just sum of rainfall & snowfall
    else
    {
        for (i=IMPERV0; i<=PERV; i++) netPrecip[i] = rainfall + snowfall;
    }

    // --- initialize loss rate & runoff rates
    Subcatch[j].losses = 0.0;
    outflow = 0.0;
    runoff = 0.0;

    // --- examine each type of sub-area
    for (i = IMPERV0; i <= PERV; i++)
    {
        // --- check that sub-area type exists
        fArea = Subcatch[j].subArea[i].fArea;
        if ( fArea > 0.0 )
        {
            // --- call getSubareaRunoff() which assigns values for
            //     global variables Outflow, Losses, and Voutflow as it
            //     computes runoff from the sub-area
            getSubareaRunoff(j, i, netPrecip[i], Evap.rate, tStep);

            // --- add sub-area results to totals, wtd. by areal coverage
            Subcatch[j].losses += Losses * fArea;
            outflow    += Outflow * fArea;
            evapVol    += Vevap * fArea;
            infilVol   += Vinfil * fArea;
            outflowVol += Voutflow * fArea;
            runoff     += Subcatch[j].subArea[i].runoff * fArea;
        }
    }

    // --- convert outflow from ft/sec to cfs & save as new runoff
    //     NOTE: 'runoff' is total runoff generated from subcatchment,
    //           'outflow' is the portion of the runoff that leaves the
    //           subcatchment (i.e., the portion that is not internally
    //           routed between the pervious and impervious areas).
    area = Subcatch[j].area;
    outflow *= area;
    if ( outflow < MIN_RUNOFF_FLOW ) outflow = 0.0;
    Subcatch[j].newRunoff = outflow;

    // --- compute rainfall+snowfall volume (does not include snowmelt)
    rainVol = Subcatch[j].rainfall * tStep;

    // --- update subcatchment's runoff totals

///////////////////////////////////////////////////////////
////  Added to support washoff calculations. (LR - 7/5/06 )
///////////////////////////////////////////////////////////
    Vevap  = evapVol;
    Vinfil = infilVol;

/////////////////////////////////////////////////////////////////////////////
////  Modified to include outflow in subcatchment stats. (LR - 3/10/06)  ////
/////////////////////////////////////////////////////////////////////////////
    stats_updateSubcatchStats(j, rainVol, Subcatch[j].runon*tStep,
        evapVol, infilVol, outflowVol, outflow);

    // --- update system flow balance
    //     (runoff volume is 0 if outlet is another subcatch)
    if ( Subcatch[j].outNode == -1 &&
         Subcatch[j].outSubcatch != j ) outflowVol = 0.0;
    massbal_updateRunoffTotals(rainVol*area, evapVol*area, infilVol*area,
                               outflowVol*area);
    return runoff;
}

//=============================================================================

float subcatch_getDepth(int j)
//
//  Input:   j = subcatchment index
//  Output:  returns average depth of water (ft)
//  Purpose: finds average depth of water over a subcatchment
//
{
    int   i;
    float fArea;
    float depth = 0.0;

    for (i = IMPERV0; i <= PERV; i++)
    {
        fArea = Subcatch[j].subArea[i].fArea;
        if ( fArea > 0.0 ) depth += Subcatch[j].subArea[i].depth * fArea;
    }
    return depth;
}

//=============================================================================

void subcatch_getBuildup(int j, float tStep)
//
//  Input:   j = subcatchment index
//           tStep = time step (sec)
//  Output:  none
//  Purpose: adds to pollutant buildup on subcatchment.
//
{
    int    i;                          // land use index
    int    p;                          // pollutant index
    float  f;                          // land use fraction
    float  area;                       // land use area (acres or hectares)
    float  curb;                       // land use curb length (user units)
    float  oldBuildup;                 // buildup at start of time step
    float  newBuildup;                 // buildup at end of time step
	//added
    float iarea;                       // impervious area (ft2 or acre or ha)
    float icurb;                       // curb length on impervious area (users units)
    float parea;                       // pervious area (acres or hectares)
    float oldDetStorage;               // detached storage at start of time step (lbs)
    float newDetStorage;               // detached storage at end of time step (lbs)
    float rainfall;					   // rainfall (ft/sec)

    // --- consider each landuse
    for (i = 0; i < Nobjects[LANDUSE]; i++)
    {
        // --- skip landuse if not in subcatch
        f = Subcatch[j].landFactor[i].fraction;
        if ( f == 0.0 ) continue;

        // --- get land area (in acres or hectares) & curb length
        area = f * Subcatch[j].area * UCF(LANDAREA);
        curb = f * Subcatch[j].curbLength;

        // --- determine area and curb length covered by land use
        iarea = area * Landuse[i].pctimp;
        icurb = curb * Landuse[i].pctimp;
        parea = area * (1 - Landuse[i].pctimp);

        // --- examine each pollutant
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            // --- use land use's buildup function to update buildup amount (lbs or kg)
            oldBuildup = Subcatch[j].landFactor[i].buildup[p]; 
			
			if (Pollut[p].sedflag > 0)	// the pollutant is sediment
				newBuildup = landuse_getBuildup(i, p, iarea, icurb, oldBuildup,
							 tStep);
			else
				newBuildup = landuse_getBuildup(i, p, area, curb, oldBuildup,
							 tStep);
			
            Subcatch[j].landFactor[i].buildup[p] = newBuildup;

			if (Pollut[p].sedflag > 0)	// the pollutant is sediment
			{
				// --- use land use's detach function to update detached amount
				oldDetStorage = Subcatch[j].landFactor[i].detstorage[p];
				rainfall = Subcatch[j].rainfall;	// ft/sec

				newDetStorage = landuse_getDetached(i, parea, oldDetStorage, 
					rainfall, tStep);
				
				Subcatch[j].landFactor[i].detstorage[p] = newDetStorage;
			}

            massbal_updateLoadingTotals(BUILDUP_LOAD, p, 
                                       (newBuildup - oldBuildup));
       }
    }
}

//=============================================================================

void subcatch_sweepBuildup(int j, DateTime aDate)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: reduces pollutant buildup over a subcatchment if sweeping occurs.
//
{
    int    i;                          // land use index
    int    p;                          // pollutant index
    float  oldBuildup;                 // buildup before sweeping (lbs or kg)
    float  newBuildup;                 // buildup after sweeping (lbs or kg)

    // --- no sweeping occurs if subcatch's surfaces are not dry
    if ( !sweptSurfacesDry(j) ) return;

    // --- consider each land use
    for (i = 0; i < Nobjects[LANDUSE]; i++)
    {
        // --- skip land use if not in subcatchment 
        if ( Subcatch[j].landFactor[i].fraction == 0.0 ) continue;

        // --- see if land use is subject to sweeping
        if ( Landuse[i].sweepInterval == 0.0 ) continue;

        // --- see if sweep interval has been reached
        if ( aDate - Subcatch[j].landFactor[i].lastSwept >=
            Landuse[i].sweepInterval )
        {
        
            // --- update time when last swept
            Subcatch[j].landFactor[i].lastSwept = aDate;

            // --- examine each pollutant
            for (p = 0; p < Nobjects[POLLUT]; p++)
            {
                // --- reduce buildup by the fraction available
                //     times the sweeping effic.
                oldBuildup = Subcatch[j].landFactor[i].buildup[p];
                newBuildup = oldBuildup * (1.0 - Landuse[i].sweepRemoval *
                             Landuse[i].washoffFunc[p].sweepEffic);
                newBuildup = MIN(oldBuildup, newBuildup);
                newBuildup = MAX(0.0, newBuildup);
                Subcatch[j].landFactor[i].buildup[p] = newBuildup;

                // --- update mass balance totals
                massbal_updateLoadingTotals(SWEEPING_LOAD, p,
                                            oldBuildup - newBuildup);
            }
        }
    }
}

//=============================================================================

void  subcatch_getWashoff(int j, float runoff, float tStep)
////////////////////////////////////////////////////////////
//  This function has been totally rewritten. (LR - 7/5/06 )
////////////////////////////////////////////////////////////
//
//  Input:   j = subcatchment index
//           runoff = total subcatchment runoff (ft/sec)
//           tStep = time step (sec)
//  Output:  none
//  Purpose: computes new runoff quality for subcatchment.
//
//  Considers two separate pollutant generating streams that are combined
//  together:
//  1. complete mix mass balance of pollutants in surface ponding due to
//     runon, deposition, infil., & evap.
//  2. washoff of pollutant buildup as described by the project's land
//     use washoff functions.
//
{
    float v;                           // ponded depth (ft)
    float qUp;                         // runon inflow rate (ft/sec)
    float qPpt;                        // precipitation rate (ft/sec)
    float qInfil;                      // infiltration rate (ft/sec)
    float qEvap;                       // evaporation rate (ft/sec)
    float area;                        // subcatchment area (ft2)
    float massLoad;                    // pollutant load (lbs or kg)
    float *wUp;                        // runon inflow loads (mass/sec)
    float *pondedQual;                 // quality of ponded water (mass/ft3)

    // --- return if there is no area or no pollutants
    area = Subcatch[j].area;
    if ( Nobjects[POLLUT] == 0 || area == 0.0 ) return;

    // --- get flow rates of the various inflows & outflows
    qUp    = Subcatch[j].runon;             // upstream runon
    qPpt   = Subcatch[j].rainfall;          // precipitation (rain + snow)
    qEvap  = Vevap / tStep;                 // evaporation
    qInfil = Vinfil / tStep;                // infiltration
    runoff = Subcatch[j].newRunoff / area;  // runoff that leaves the subcatchment

    // --- assign upstream runon load (computed previously from call to 
    //     subcatch_getRunon) and ponded quality to local variables
    //     merely for notational convenience
    wUp = Subcatch[j].newQual;
    pondedQual = Subcatch[j].pondedQual;

    // --- avgerage the ponded depth volumes over the time step
    v = 0.5 * (Vponded + subcatch_getDepth(j));

    // --- get quality in surface ponding at end of time step
    getPondedQual(wUp, qUp, qPpt, qEvap, qInfil, v, area, tStep, pondedQual);

    // --- get quality in washoff from pollutant buildup
    getWashoffQual(j, runoff, tStep, WashoffQual);

    // --- combine ponded & washoff quality in the subcatchment's outflow
    //     (updates Subcatch[j].newQual[])
    combineWashoffQual(j, pondedQual, WashoffQual, tStep);
}

//=============================================================================

//////////////////////////////////////
//  New function added. (LR - 7/5/06 )
//////////////////////////////////////
void getPondedQual(float wUp[], float qUp, float qPpt, float qEvap,
         float qInfil, float v, float area, float tStep, float pondedQual[])
//
//  Input:   wUp[]  = runon load from upstream subcatchments (mass/sec)
//           qUp    = runon inflow flow rate (ft/sec)
//           qPpt   = precip. rate (ft/sec)
//           qEvap  = evaporation rate (ft/sec)
//           qInfil = infiltration rate (ft/sec)
//           v      = ponded depth (ft)
//           area   = subcatchment area (ft2)
//           tStep  = time step (sec)
//  Output:  pondedQual[] = pollutant concentrations in ponded water (mass/ft3)
//  Purpose: computes new quality of ponded surface water in a subcatchment.
//
{
    int   p;
    float wIn, wPpt;
    float massLoad;
    float qNet = qUp + qPpt - qEvap;

    // --- analyze each individual pollutant
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        // --- compute direct deposition from precip. in mass/sec
        wPpt = Pollut[p].pptConcen * LperFT3 * qPpt * area;

        // --- add direct deposition (in lbs or kg) to mass balance
        massLoad = wPpt * tStep * Pollut[p].mcf;
        massbal_updateLoadingTotals(DEPOSITION_LOAD, p, massLoad);

        // --- add infiltration loss to mass balance
        massLoad = pondedQual[p] * qInfil * area * tStep * Pollut[p].mcf;
        massbal_updateLoadingTotals(INFIL_LOAD, p, massLoad);

        // ---- add direct deposition to runon load from upstream subcatchments
        wIn = (wPpt + wUp[p]) / area;
    
        // --- update ponded concentration (in mass/ft3)
        pondedQual[p] = getCstrQual(pondedQual[p], v, wIn, qNet, tStep);
    }
}

//=============================================================================

//////////////////////////////////////
//  New function added. (LR - 7/5/06 )
//////////////////////////////////////
float getCstrQual(float c, float v, float wIn, float qNet, float tStep)
//
//  Input:   c       = concen. in CSTR at start of time step (mass/ft3)
//           v       = depth of water in CSTR (ft or ft3)
//           wIn     = mass inflow rate (mass/ft2/sec or mass/sec)
//           qNet    = net inflow flow rate (ft/sec or ft3/sec)
//           tStep   = time step (sec)
//  Output:  returns concen. in CSTR at end of time step
//  Purpose: updates the concentration in a continuously stirred tank reactor
//           (CSTR) over a given time step.
//
{
    float cIn, vNet, expp;

    // --- if no volume, then concen. is zero
    if ( v <= 0.0) return 0.0;

    // --- if no flow, no change in concen.
    if ( qNet == 0.0 ) return c;

    // --- net inflow volume w.r.t. reactor volume
    vNet = qNet * tStep / v;

    // --- net outflow can't be > reactor volume
    if (vNet < -1.0) vNet = -1.0;

    // --- inflow concentration
    cIn = wIn / qNet;

    // --- if inflow >> v, then inflow dominates
    if ( vNet > 12.0 ) return MAX(0.0, cIn);

    // --- otherwise combine cIn and c
    expp = exp(-vNet);
    c = cIn * (1.0 - expp) + c * expp;

    // --- negative concen. not allowed
    return MAX(0.0, c);
}

//=============================================================================

//////////////////////////////////////
//  New function added. (LR - 7/5/06 )
//////////////////////////////////////
void  getWashoffQual(int j, float runoff, float tStep, float washoffQual[])
//
//  Input:   j       = subcatchment index
//           runoff  = runoff flow rate over entire subcatchment (ft/sec)
//           tStep   = time step (sec)
//  Output:  washoffQual[] = quality of surface washoff (mass/ft3)
//  Purpose: finds concentrations in washoff from pollutant buildup over a
//           subcatchment within a time step.
//
{
    int   p;                           // pollutant index
    int   i;                           // land use index
    float area;                        // subcatchment area (ft2)

    // --- initialize total washoff quality from subcatchment
    area = Subcatch[j].area;
    for (p = 0; p < Nobjects[POLLUT]; p++) washoffQual[p] = 0.0;

/////////////////////////
// Modified (LR - 9/19/06)
/////////////////////////
    if ( area*runoff <= MIN_RUNOFF_FLOW || area == 0.0 ) return;

    // --- get local washoff mass flow from each landuse and add to total
    for (i = 0; i < Nobjects[LANDUSE]; i++)
    {
        if ( Subcatch[j].landFactor[i].fraction == 0.0 ) continue;

		// from the impervious land
        landuse_getWashoff(i, area, Subcatch[j].landFactor, runoff, tStep,
            WashoffLoad);

		// from the pervious land
        landuse_getRemoval(i, j, area, Subcatch[j].landFactor, runoff, tStep, 
			RemovalLoad);

        for (p = 0; p < Nobjects[POLLUT]; p++) 
		{
	        washoffQual[p] += WashoffLoad[p] + RemovalLoad[p];
        }
    }

    // --- convert from mass/sec to mass/ft3
    runoff *= area;
    for (p = 0; p < Nobjects[POLLUT]; p++) washoffQual[p] /= runoff; 
}

//=============================================================================

//////////////////////////////////////
//  New function added. (LR - 7/5/06 )
//////////////////////////////////////
void combineWashoffQual(int j, float pondedQual[], float washoffQual[],
        float tStep)
//
//  Input:   j             = subcatchment index
//           pondedQual[]  = quality of ponded water (mass/ft3)
//           washoffQual[] = quality of washoff (mass/ft3)
//           tStep         = time step (sec)
//  Output:  updates Subcatch[j].newQual[]
//  Purpose: computes combined concentration of ponded water & washoff streams
{
    int   p;
    float qOut, cOut, bmpRemoval, massLoad;

    qOut = Subcatch[j].newRunoff;
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        // --- add concen. of ponded water to that of washoff
        cOut = pondedQual[p] + WashoffQual[p];
        
        // --- apply any BMP removal
        bmpRemoval = getBmpRemoval(j, p) * cOut;
        massLoad = bmpRemoval * qOut *  tStep * Pollut[p].mcf; 
        massbal_updateLoadingTotals(BMP_REMOVAL_LOAD, p, massLoad);
        cOut -= bmpRemoval;
            
        // --- save new outflow runoff concentration (in mass/L)
        Subcatch[j].newQual[p] = MAX(cOut, 0.0) / LperFT3;

        // --- update total runoff pollutant load from subcatchment
        massLoad = 0.5 * (Subcatch[j].oldQual[p]*Subcatch[j].oldRunoff +
                          Subcatch[j].newQual[p]*Subcatch[j].newRunoff) *
                          LperFT3 * tStep * Pollut[p].mcf;
        Subcatch[j].totalLoad[p] += massLoad;

        // --- update mass balance if runoff goes to an outlet node
        if ( Subcatch[j].outNode >= 0 ) 
        {
            massbal_updateLoadingTotals(RUNOFF_LOAD, p, massLoad);
        }
    }
}

//=============================================================================

//////////////////////////////////////
//  New function added. (LR - 7/5/06 )
//////////////////////////////////////
float getBmpRemoval(int j, int p)
//
//  Input:   j = subcatchment index
//           p = pollutant index
//  Output:  returns a BMP removal fraction for pollutant p
//  Purpose: finds the overall average BMP removal achieved for pollutant p
//           treated in subcatchment j.
{
    int i;
    float r = 0.0;
    for (i = 0; i < Nobjects[LANDUSE]; i++)
    {
        r += Subcatch[j].landFactor[i].fraction *
             Landuse[i].washoffFunc[p].bmpEffic;
    }
    return r;
}

//=============================================================================

float subcatch_getWtdOutflow(int j, float f)
//
//  Input:   j = subcatchment index
//           f = weighting factor.
//  Output:  returns weighted runoff value
//  Purpose: computes wtd. combination of old and new subcatchment runoff.
//
{
    if ( Subcatch[j].area == 0.0 ) return 0.0;
    return (1.0 - f) * Subcatch[j].oldRunoff + f * Subcatch[j].newRunoff;
}

//=============================================================================

float subcatch_getWtdWashoff(int j, int p, float f)
//
//  Input:   j = subcatchment index
//           p = pollutant index
//           f = weighting factor
//  Output:  returns pollutant washoff value
//  Purpose: finds wtd. combination of old and new washoff for a pollutant.
//
{
    return (1.0 - f) * Subcatch[j].oldQual[p] + f * Subcatch[j].newQual[p];
}

//=============================================================================

void  subcatch_getResults(int j, float f, float x[])
//
//  Input:   j = subcatchment index
//           f = weighting factor
//  Output:  x = array of results
//  Purpose: computes wtd. combination of old and new subcatchment results.
//
{
    int    p;                          // pollutant index
    int    k;                          // rain gage index
    float  f1 = 1.0 - f;
    TGroundwater* gw;                  // ptr. to groundwater object

    // --- retrieve rainfall for current report period
    k = Subcatch[j].gage;
    if ( k >= 0 ) x[SUBCATCH_RAINFALL] = Gage[k].reportRainfall;
    else          x[SUBCATCH_RAINFALL] = 0.0;

    // --- retrieve snow depth
    x[SUBCATCH_SNOWDEPTH] = ( f1 * Subcatch[j].oldSnowDepth +
                              f * Subcatch[j].newSnowDepth ) * UCF(RAINDEPTH);

    // --- retrieve runoff and losses
    x[SUBCATCH_LOSSES] = Subcatch[j].losses * UCF(RAINFALL);
    x[SUBCATCH_RUNOFF] = ( f1 * Subcatch[j].oldRunoff +
                           f * Subcatch[j].newRunoff ) * UCF(FLOW);

    // --- retrieve groundwater flow & water table if present
    gw = Subcatch[j].groundwater;
    if ( gw )
    {
        x[SUBCATCH_GW_FLOW] = (f1 * gw->oldFlow + f * gw->newFlow) *
                              Subcatch[j].area * UCF(FLOW);
        x[SUBCATCH_GW_ELEV] = (Aquifer[gw->aquifer].bottomElev +
                              gw->lowerDepth) * UCF(LENGTH);
    }
///////////////////////////////////////////////////////////////
//  Change GW variables to 0 if GW not simulated. (LR - 9/5/05)
///////////////////////////////////////////////////////////////
    else
    {
        x[SUBCATCH_GW_FLOW] = 0.0;
        x[SUBCATCH_GW_ELEV] = 0.0;
    }

    // --- retrieve pollutant washoff
    for (p = 0; p < Nobjects[POLLUT]; p++ )
    {
        x[SUBCATCH_WASHOFF+p] = f1 * Subcatch[j].oldQual[p] +
                                f * Subcatch[j].newQual[p];
    }    
}

//=============================================================================

char  sweptSurfacesDry(int j)
//
//  Input:   j = subcatchment index
//  Output:  returns TRUE if subcatchment surfaces are dry
//  Purpose: checks if surfaces subject to street sweeping are dry.
//
{
    float      depth;                            // depth of standing water (ft)
    TSnowpack* snowpack = Subcatch[j].snowpack;  // snowpack data

    // --- check snow depth on plowable impervious area
    if ( snowpack != NULL )
    {
        if ( snowpack->wsnow[IMPERV0] > MIN_TOTAL_DEPTH ) return FALSE;
    }

    // --- check water depth on impervious surfaces
    if ( Subcatch[j].fracImperv > 0.0 )
    {
       depth = (Subcatch[j].subArea[IMPERV0].depth *
                Subcatch[j].subArea[IMPERV0].fArea) +
               (Subcatch[j].subArea[IMPERV1].depth *
                Subcatch[j].subArea[IMPERV1].fArea);
       depth = depth / Subcatch[j].fracImperv;
       if ( depth > MIN_TOTAL_DEPTH ) return FALSE;
    }
    return TRUE;
}


//=============================================================================
//                              SUB-AREA METHODS
//=============================================================================

void getSubareaRunoff(int j, int i, float precip, float evap, float tStep)
//
//  Input:   j = subcatchment index
//           i = subarea index
//           precip = rainfall + snowmelt over subarea (ft/sec)
//           evap = evaporation (ft/sec)
//           tStep = time step (sec)
//  Output:  none
//  Purpose: computes runoff & losses from a subarea over the current time step.
//
{
    float  tRunoff;                    // time over which runoff occurs (sec)
    float  oldRunoff;                  // runoff from previous time period
    float  availMoisture;              // sum of precipitation & ponded water (ft)
    float  xDepth;                     // ponded depth above dep. storage (ft)
    float  infil;                      // infiltration rate (ft/sec)
    float  surfEvap;                   // evap. used for surface water (ft/sec)
    float  subsurfEvap;                // evap. available for subsurface water
    TSubarea* subarea;                 // pointer to subarea being analyzed

    // --- assign pointer to current subarea
    subarea = &Subcatch[j].subArea[i];

    // --- assume runoff occurs over entire time step
    tRunoff = tStep;

    // --- initialize runoff & losses
    oldRunoff = subarea->runoff;
    subarea->runoff = 0.0;
    Vevap    = 0.0;
    Vinfil   = 0.0;
    Voutflow = 0.0;
    Losses   = 0.0;
    Outflow  = 0.0;

    // --- no runoff if no area
    if ( subarea->fArea == 0.0 ) return;
    subarea->inflow += precip;
    availMoisture = subarea->inflow + subarea->depth / tStep;
    surfEvap = MIN(availMoisture, evap);
    subsurfEvap = evap - surfEvap; 

    // --- compute infiltration loss rate for pervious subarea
    infil = 0.0;
    if ( i == PERV )
    {
        if ( Subcatch[j].infil  == j )
        {

///////////////////////////////////////////////////////////////////////////
//  Need to subtract off evaporation from inflow rate used for infiltration
//  calculations. - (LR via RD - 9/5/05)
///////////////////////////////////////////////////////////////////////////
            infil = infil_getInfil(j, InfilModel, tStep,
                    (subarea->inflow - surfEvap), subarea->depth);

        }
        if ( infil > availMoisture - surfEvap )
        {
            infil = MAX(0.0, availMoisture - surfEvap);
        }

        // --- update groundwater which might limit amount of infiltration
        gwater_getGroundwater(j, subsurfEvap, &infil, tStep);
    }

    // --- compute evaporation & infiltration volumes
    Vevap = surfEvap * tStep;
    Vinfil = infil * tStep;

    // --- if losses exceed available moisture then there's no ponded water
    Losses = surfEvap + infil;
    if ( Losses >= availMoisture )
    {
        Losses = availMoisture;
        subarea->depth = 0.0;
    }

    // --- otherwise update depth of ponded water
    //     and time over which runoff occurs
    else updatePondedDepth(subarea, &tRunoff);

    // --- compute runoff based on updated ponded depth
    xDepth = subarea->depth - subarea->dStore;
    if ( xDepth > MIN_EXCESS_DEPTH )
    {
        // --- case where nonlinear routing is used
        if ( subarea->N > 0.0 )
        {
            subarea->runoff = subarea->alpha * pow(xDepth, MEXP);
        }

        // --- case where no routing is used (Mannings N = 0)
        else
        {
            subarea->runoff = xDepth / tRunoff;
            subarea->depth = subarea->dStore;
        }
    }
    else subarea->runoff = 0.0;

    // --- compute runoff volume leaving subcatchment for mass balance purposes
    //     (fOutlet is the fraction of this subarea's runoff that goes to the
    //     subcatchment outlet as opposed to another subarea of the subcatchment)
    if ( subarea->fOutlet > 0.0 )
    {
        Voutflow = 0.5 * (oldRunoff + subarea->runoff) * tRunoff
                  * subarea->fOutlet;
        Outflow = subarea->fOutlet * subarea->runoff;
    }
}

//=============================================================================

void updatePondedDepth(TSubarea* subarea, float* dt)
//
//  Input:   subarea = ptr. to a subarea,
//           dt = time step (sec)
//  Output:  dt = time ponded depth is above depression storage (sec)
//  Purpose: computes new ponded depth over subarea after current time step.
//
{
    float ix;                          // excess inflow to subarea (ft/sec)
    float dx;                          // depth above depression storage (ft)
    float tx = *dt;                    // time over which dx > 0 (sec)

    // --- excess inflow = total inflow - losses
    ix = subarea->inflow - Losses;

    // --- see if not enough inflow to fill depression storage (dStore)
    if ( subarea->depth + ix*tx <= subarea->dStore )
    {
        subarea->depth += ix * tx;
    }

    // --- otherwise use the ODE solver to integrate flow depth
    else
    {
        // --- if depth < dStore then fill up dStore & reduce time step
        dx = subarea->dStore - subarea->depth;
        if ( dx > 0.0 && ix > 0.0 )
        {
            tx -= dx / ix;
            subarea->depth = subarea->dStore;
        }

        // --- now integrate depth over remaining time step tx
        if ( subarea->alpha > 0.0 && tx > 0.0 )
        {
            theSubarea = subarea;
            odesolve_integrate(&(subarea->depth), 1, 0, tx, ODETOL, tx,
                               getDdDt);
        }
        else
        {
            if ( tx < 0.0 ) tx = 0.0;
            subarea->depth += ix * tx;
        }
    }

    // --- do not allow ponded depth to go negative
    if ( subarea->depth < 0.0 ) subarea->depth = 0.0;

    // --- replace original time step with time ponded depth
    //     is above depression storage
    *dt = tx;
}

//=============================================================================

void  getDdDt(float t, float* d, float* dddt)
//
//  Input:   t = current time (not used)
//           d = stored depth (ft)
//  Output   dddt = derivative of d with respect to time
//  Purpose: evaluates derivative of stored depth w.r.t. time
//           for the subarea whose runoff is being computed.
//
{
    float ix = theSubarea->inflow - Losses;
    float rx = *d - theSubarea->dStore;
    if ( rx < 0.0 )
    {
        rx = 0.0;
    }
    else
    {
        rx = theSubarea->alpha * pow(rx, MEXP);
    }
    *dddt = ix - rx;
}

//=============================================================================
