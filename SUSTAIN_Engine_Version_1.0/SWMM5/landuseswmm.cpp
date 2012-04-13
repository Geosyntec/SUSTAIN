//-----------------------------------------------------------------------------
//   landuse.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman
//
//   Pollutant buildup and washoff functions.
//-----------------------------------------------------------------------------

#include <math.h>
#include <string.h>
#include "headers.h"
//HSPF
#include "../Sediment.h"	

//-----------------------------------------------------------------------------
//  Imported variables (declared in runoff.c)
//-----------------------------------------------------------------------------
///////////////////////////////////////////////////
//  This variable is no longer used. (LR - 7/5/06 )
///////////////////////////////////////////////////
//extern  float*    LocalWashoff;   // washoff load from a land use (mass/sec)

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  landuse_readParams        (called by parseLine in input.c)
//  landuse_readPollutParams  (called by parseLine in input.c)
//  landuse_readBuildupParams (called by parseLine in input.c)
//  landuse_readWashoffParams (called by parseLine in input.c)
//  landuse_getBuildup        (called by subcatch_getBuildup)
//  landuse_getWashoff        (called by getWashoffLoads in subcatch.c)
//  landuse_getRemoval        (called by getRemovalLoads in subcatch.c)

//-----------------------------------------------------------------------------
// Function declarations
//-----------------------------------------------------------------------------
static float  landuse_getBuildupDays(int landuse, int pollut, float buildup);
static float  landuse_getBuildupMass(int landuse, int pollut, float days);
static float  landuse_getRunoffLoad(int landuse, int pollut, float area,
              TLandFactor landFactor[], float runoff, float tStep);
static float  landuse_getWashoffMass(int landuse, int pollut, float buildup,
              float runoff, float area);
static float  landuse_getCoPollutLoad(int pollut, float washoff[], float tStep);

//added
static float  landuse_getRunoffLoad2(int landuse, int subcatch, int pollut, float area,
              TLandFactor landFactor[], float runoff, float tStep);	
static float  landuse_getRemovalMass(int landuse,float area,float tStep,float rainfall,
									 float runoff, float surs, float& dets);


///////////////////////////////////////////////////
//  This function is no longer used. (LR - 7/5/06 )
///////////////////////////////////////////////////
//static float  landuse_getBMPRemoval(int landuse, int pollut, float washoff,
//              float tStep);


//=============================================================================

int  landuse_readParams(int j, char* tok[], int ntoks)
//
//  Input:   j = land use index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads landuse parameters from a tokenized line of input.
//
//  Data format is:
//    landuseID  (sweepInterval sweepRemoval sweepDays0
//HSPF --> pctimp smpf krer jrer affix cover kser jser kger jger frc_sand frc_silt frc_clay)
//
{
    char *id;
    if ( ntoks < 1 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LANDUSE, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    Landuse[j].ID = id;
    if ( ntoks > 1 )
    {
        //if ( ntoks < 4 ) return error_setInpError(ERR_ITEMS, "");
        if ( ntoks < 17 ) return error_setInpError(ERR_ITEMS, "");
        if ( ! getFloat(tok[1], &Landuse[j].sweepInterval) )
            return error_setInpError(ERR_NUMBER, tok[1]);
        if ( ! getFloat(tok[2], &Landuse[j].sweepRemoval) )
            return error_setInpError(ERR_NUMBER, tok[2]);
        if ( ! getFloat(tok[3], &Landuse[j].sweepDays0) )
            return error_setInpError(ERR_NUMBER, tok[3]);
		// HSPF parameters
        if ( ! getFloat(tok[4], &Landuse[j].pctimp) )
            return error_setInpError(ERR_NUMBER, tok[4]);
        if ( ! getFloat(tok[5], &Landuse[j].smpf) )
            return error_setInpError(ERR_NUMBER, tok[5]);
        if ( ! getFloat(tok[6], &Landuse[j].krer) )
            return error_setInpError(ERR_NUMBER, tok[6]);
        if ( ! getFloat(tok[7], &Landuse[j].jrer) )
            return error_setInpError(ERR_NUMBER, tok[7]);
        if ( ! getFloat(tok[8], &Landuse[j].affix) )
            return error_setInpError(ERR_NUMBER, tok[8]);
        if ( ! getFloat(tok[9], &Landuse[j].cover) )
            return error_setInpError(ERR_NUMBER, tok[9]);
        if ( ! getFloat(tok[10], &Landuse[j].kser) )
            return error_setInpError(ERR_NUMBER, tok[10]);
        if ( ! getFloat(tok[11], &Landuse[j].jser) )
            return error_setInpError(ERR_NUMBER, tok[11]);
        if ( ! getFloat(tok[12], &Landuse[j].kger) )
            return error_setInpError(ERR_NUMBER, tok[12]);
        if ( ! getFloat(tok[13], &Landuse[j].jger) )
            return error_setInpError(ERR_NUMBER, tok[13]);
        if ( ! getFloat(tok[14], &Landuse[j].frc_sand) )
            return error_setInpError(ERR_NUMBER, tok[14]);
        if ( ! getFloat(tok[15], &Landuse[j].frc_silt) )
            return error_setInpError(ERR_NUMBER, tok[15]);
        if ( ! getFloat(tok[16], &Landuse[j].frc_clay) )
            return error_setInpError(ERR_NUMBER, tok[16]);
    }
    else
    {
        Landuse[j].sweepInterval = 0.0;
        Landuse[j].sweepRemoval = 0.0;
        Landuse[j].sweepDays0 = 0.0;
        Landuse[j].pctimp = 0.0;
        Landuse[j].smpf = 0.0;
        Landuse[j].krer = 0.0;
        Landuse[j].jrer = 0.0;
        Landuse[j].affix = 0.0;
        Landuse[j].cover = 0.0;
        Landuse[j].kser = 0.0;
        Landuse[j].jser = 0.0;
        Landuse[j].kger = 0.0;
        Landuse[j].jger = 0.0;
        Landuse[j].frc_sand = 0.0;
        Landuse[j].frc_silt = 0.0;
        Landuse[j].frc_clay = 0.0;
     }
	        
    if ( Landuse[j].sweepRemoval < 0.0
        || Landuse[j].sweepRemoval > 1.0 )
        return error_setInpError(ERR_NUMBER, tok[2]);
    return 0;
}

//=============================================================================

int  landuse_readPollutParams(int j, char* tok[], int ntoks)
//
//  Input:   j = pollutant index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads pollutant parameters from a tokenized line of input.
//
//  Data format is:
//  pollutID  cUnits  cRain  cGW  cRDII  kDecay  IsSediment  snowOnly  coPollut  coFrac
//
{
    int   i, k, coPollut, snowFlag, sedFlag;
    float x[4], coFrac;
    char  *id;

    // --- extract pollutant name & units
    if ( ntoks < 6 ) return error_setInpError(ERR_ITEMS, "");

	//added
	sedFlag = 0;

    if ( ntoks >= 7 )
    {
		sedFlag = findmatch(tok[6], SedimentWords);             
			if ( sedFlag < 0 ) return error_setInpError(ERR_KEYWORD, tok[6]);
	}

	if (sedFlag == 4)
	{
		sedFlag = 1;
		tok[0] = "SAND";
	}

    id = project_findID(POLLUT, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    k = findmatch(tok[1], QualUnitsWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[1]);

    // --- extract concen. in rain, gwater, & I&I and decay coeff
    for ( i = 2; i <= 5; i++ )
    {
        if ( ! getFloat(tok[i], &x[i-2]) )
            return error_setInpError(ERR_NUMBER, tok[i]);
    }

    // --- set defaults for snow only flag & co-pollut. parameters
    snowFlag = 0;
    coPollut = -1;
    coFrac = 0.0;

    // --- check for snow only flag
    if ( ntoks >= 8 )
    {
        snowFlag = findmatch(tok[7], NoYesWords);             
        if ( snowFlag < 0 ) return error_setInpError(ERR_KEYWORD, tok[7]);
    }

    // --- check for co-pollutant
    if ( ntoks >= 10 )
    {
        if ( !strcomp(tok[8], "*") )
        {
            coPollut = project_findObject(POLLUT, tok[8]);
            if ( coPollut < 0 ) return error_setInpError(ERR_NAME, tok[8]);
            if ( ! getFloat(tok[9], &coFrac) )
                return error_setInpError(ERR_NUMBER, tok[9]);
        }
    }

    // --- save values for pollutant object   
    Pollut[j].ID = id;
    Pollut[j].units = k;
    if      ( Pollut[j].units == MG ) Pollut[j].mcf = UCF(MASS);
    else if ( Pollut[j].units == UG ) Pollut[j].mcf = UCF(MASS) / 1000.0;
    else                              Pollut[j].mcf = 1.0;
    Pollut[j].pptConcen = x[0];
    Pollut[j].gwConcen  = x[1];
    Pollut[j].rdiiConcen = x[2];
    Pollut[j].kDecay = x[3]/SECperDAY;
    Pollut[j].sedflag = sedFlag;
    Pollut[j].snowOnly = snowFlag;
    Pollut[j].coPollut = coPollut;
    Pollut[j].coFraction = coFrac;
    return 0;
}

//=============================================================================

int  landuse_readBuildupParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads pollutant buildup parameters from a tokenized line of input.
//
//  Data format is:
//    landuseID  pollutID  buildupType  c1  c2  c3  normalizerType
//
{
    int i, j, k, n, p;
    float c[3], tmax;

    if ( ntoks < 3 ) return 0;
    j = project_findObject(LANDUSE, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

	//added
	if(strncmp(tok[1],strTSS,MAXFNAME) == 0) 
		tok[1] = "SAND";
    p = project_findObject(POLLUT, tok[1]);
    if ( p < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    k = findmatch(tok[2], BuildupTypeWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[2]);
    Landuse[j].buildupFunc[p].funcType = k;
    if ( k > NO_BUILDUP )
    {
        if ( ntoks < 7 ) return error_setInpError(ERR_ITEMS, "");
        for (i=0; i<3; i++)
        {
            if ( ! getFloat(tok[i+3], &c[i])  || c[i] < 0.0  )
            return error_setInpError(ERR_NUMBER, tok[i+3]);
        }
        n = findmatch(tok[6], NormalizerWords);
        if (n < 0 ) return error_setInpError(ERR_KEYWORD, tok[6]);
        Landuse[j].buildupFunc[p].normalizer = n;
    }

    // Find time until max. buildup
    switch (Landuse[j].buildupFunc[p].funcType)
    {
      case POWER_BUILDUP:
        // --- check for too small or large an exponent
        if ( c[2] > 0.0 && (c[2] < 0.01 || c[2] > 10.0) )
            return error_setInpError(ERR_KEYWORD, tok[5]);

        // --- find time to reach max. buildup
        // --- use zero if coeffs. are 0        
        if ( c[1]*c[2] == 0.0 ) tmax = 0.0;

        // --- use 10 years if inverse power function tends to blow up
        else if ( log10(c[0]) / c[2] > 3.5 ) tmax = 3650.0;

        // --- otherwise use inverse power function
        else tmax = pow(c[0]/c[1], 1.0/c[2]);
        break;

      case EXPON_BUILDUP:
        if ( c[1] == 0.0 ) tmax = 0.0;
        else tmax = -log(0.001)/c[1];
        break;

      case SATUR_BUILDUP:
        tmax = 1000.0*c[2];
        break;

      default:
        tmax = 0.0;
    }

    // Assign parameters to buildup object
    Landuse[j].buildupFunc[p].coeff[0]   = c[0];
    Landuse[j].buildupFunc[p].coeff[1]   = c[1];
    Landuse[j].buildupFunc[p].coeff[2]   = c[2];
    Landuse[j].buildupFunc[p].maxDays = tmax;
    return 0;
}

//=============================================================================

int  landuse_readWashoffParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads pollutant washoff parameters from a tokenized line of input.
//
//  Data format is:
//    landuseID  pollutID  washoffType  c1  c2  sweepEffic  bmpRemoval
{
    int i, j, p;
    int func;
    float x[4];

    if ( ntoks < 3 ) return 0;
    for (i=0; i<4; i++) x[i] = 0.0;
    func = NO_WASHOFF;
    j = project_findObject(LANDUSE, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

	//added
	if(strncmp(tok[1],strTSS,MAXFNAME) == 0) 
		tok[1] = "SAND";

    p = project_findObject(POLLUT, tok[1]);
    if ( p < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    if ( ntoks > 2 )
    {
        func = findmatch(tok[2], WashoffTypeWords);
        if ( func < 0 ) return error_setInpError(ERR_KEYWORD, tok[2]);
        if ( func != NO_WASHOFF )
        {
            if ( ntoks < 5 ) return error_setInpError(ERR_ITEMS, "");
            if ( ! getFloat(tok[3], &x[0]) )
                    return error_setInpError(ERR_NUMBER, tok[3]);
            if ( ! getFloat(tok[4], &x[1]) )
                    return error_setInpError(ERR_NUMBER, tok[4]);
            if ( ntoks >= 6 )
            {
                if ( ! getFloat(tok[5], &x[2]) )
                        return error_setInpError(ERR_NUMBER, tok[5]);
            }
            if ( ntoks >= 7 )
            {
                if ( ! getFloat(tok[6], &x[3]) )
                        return error_setInpError(ERR_NUMBER, tok[6]);
            }
        }
    }

    // --- check for valid parameter values
    //     x[0] = washoff coeff.
    //     x[1] = washoff expon.
    //     x[2] = sweep effic.
    //     x[3] = BMP effic.
    if ( x[0] < 0.0 ) return error_setInpError(ERR_NUMBER, tok[3]);
    if ( x[1] < -10.0 || x[1] > 10.0 )
        return error_setInpError(ERR_NUMBER, tok[4]);;
    if ( x[2] < 0.0 || x[2] > 100.0 )
        return error_setInpError(ERR_NUMBER, tok[5]);
    if ( x[3] < 0.0 || x[3] > 100.0 )
        return error_setInpError(ERR_NUMBER, tok[6]);

    // --- convert units of washoff coeff.
    if ( func == EXPON_WASHOFF  ) x[0] /= 3600.0;
    if ( func == RATING_WASHOFF ) x[0] *= pow(UCF(FLOW), x[1]);
    if ( func == EMC_WASHOFF    ) x[0] *= LperFT3;

    // --- assign washoff parameters to washoff object
    Landuse[j].washoffFunc[p].funcType = func;
    Landuse[j].washoffFunc[p].coeff = x[0];
    Landuse[j].washoffFunc[p].expon = x[1];
    Landuse[j].washoffFunc[p].sweepEffic = x[2] / 100.0;
    Landuse[j].washoffFunc[p].bmpEffic = x[3] / 100.0;
    return 0;
}

//=============================================================================

float  landuse_getBuildup(int i, int p, float area, float curb, float buildup,
                          float tStep)
//
//  Input:   i = land use index
//           p = pollutant index
//           area = land use area (ac or ha)
//           curb = land use curb length (users units)
//           buildup = current pollutant buildup (lbs or kg)
//           tStep = time increment for buildup (sec)
//  Output:  returns new buildup mass (lbs or kg)
//  Purpose: computes new pollutant buildup on a landuse after a time increment.
//
{
    int    n;                          // normalizer code
    float  days;                       // accumulated days of buildup
    float  perUnit;                    // normalizer value (area or curb length)
	float  buildupmass = 0.0;

    // --- return 0 if no buildup function or time increment
    if ( Landuse[i].buildupFunc[p].funcType == NO_BUILDUP || tStep == 0.0 )
    {
        return 0.0;
    }

    // --- see what buildup is normalized to
    n = Landuse[i].buildupFunc[p].normalizer;
    perUnit = 1.0;
    if ( n == PER_AREA ) perUnit = area;
    if ( n == PER_CURB ) perUnit = curb;
    if ( perUnit == 0.0 ) return 0.0;

    // --- determine equivalent days of current buildup
    days = landuse_getBuildupDays(i, p, buildup/perUnit);

    // --- compute buildup after adding on time increment
    days += tStep / SECperDAY;
    buildupmass = landuse_getBuildupMass(i, p, days) * perUnit;

	return buildupmass;
}

//=============================================================================

float landuse_getBuildupDays(int i, int p, float buildup)
//
//  Input:   i = land use index
//           p = pollutant index
//           buildup = amount of pollutant buildup
//  Output:  returns number of days it takes for buildup to reach a given level
//  Purpose: finds the number of days corresponding to a pollutant buildup.
//
{
    float c0 = Landuse[i].buildupFunc[p].coeff[0];
    float c1 = Landuse[i].buildupFunc[p].coeff[1];
    float c2 = Landuse[i].buildupFunc[p].coeff[2];

    if ( buildup == 0.0 ) return 0.0;
    if ( buildup >= c0 ) return Landuse[i].buildupFunc[p].maxDays;   
    switch (Landuse[i].buildupFunc[p].funcType)
    {
      case POWER_BUILDUP:
        if ( c1*c2 == 0.0 ) return 0.0;
        else return pow( (buildup/c1), (1.0/c2) );

      case EXPON_BUILDUP:
        if ( c0*c1 == 0.0 ) return 0.0;
        else return -log(1. - buildup/c0) / c1;

      case SATUR_BUILDUP:
        if ( c0 == 0.0 ) return 0.0;
        else return buildup*c2 / (c0 - buildup);

      default:
        return 0.0;
    }
}

//=============================================================================

float landuse_getBuildupMass(int i, int p, float days)
//
//  Input:   i = land use index
//           p = pollutant index
//           days = time over which buildup has occurred (days)
//  Output:  returns mass of pollutant buildup (lbs or kg per area or curblength)
//  Purpose: finds amount of buildup of pollutant on a land use.
//
{
    float b;
    float c0 = Landuse[i].buildupFunc[p].coeff[0];
    float c1 = Landuse[i].buildupFunc[p].coeff[1];
    float c2 = Landuse[i].buildupFunc[p].coeff[2];

    if ( days == 0.0 ) return 0.0;
    if ( days >= Landuse[i].buildupFunc[p].maxDays ) return c0;
    switch (Landuse[i].buildupFunc[p].funcType)
    {
      case POWER_BUILDUP:
        b = c1 * pow(days, c2);
        if ( b > c0 ) b = c0;
        break;

      case EXPON_BUILDUP:
        b = c0*(1.0 - exp(-days*c1));
        break;

      case SATUR_BUILDUP:
        b = days*c0/(c2 + days);
        break;

      default: b = 0.0;
    }
    return b;
}

//=============================================================================

////////////////////////////////////////////////////////////
//  This function has been totally rewritten. (LR - 7/5/06 )
////////////////////////////////////////////////////////////
//float landuse_getWashoff(int i, float area, TLandFactor landFactor[],
//                        float fPpt, float runoff, float wUpstrm[],
//                        float wTotal[], float tStep)
void  landuse_getWashoff(int i, float area, TLandFactor landFactor[],
                         float qRunoff, float tStep, float washoffLoad[])
//
//  Input:   i            = land use index
//           area         = subcatchment area (ft2)
//           landFactor[] = array of land use data for subcatchment
//           runoff       = runoff flow rate (ft/sec) over subcatchment
//           tStep        = time step (sec)
//  Output:  washoffLoad[] = pollutant load in surface washoff (mass/sec)
//  Purpose: computes surface washoff load for all pollutants generated by a
//           land use within a subcatchment.
//
{
    int   p;                           //pollutant index
    float fArea;                       //area devoted to land use (ft2)
	//added	
	float iArea;					   //impervious area (ft2) 

    // --- find area devoted to land use
    fArea = landFactor[i].fraction * area;

    // --- find impervious area devoted to land use
	iArea = fArea * Landuse[i].pctimp;

    // --- initialize washoff loads from land use
    for (p = 0; p < Nobjects[POLLUT]; p++) washoffLoad[p] = 0.0;

    // --- compute contribution from direct runoff load
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
		if (Pollut[p].sedflag > 0)	// the pollutant is sediment
			// calculate sediment washoff on impervious area
			washoffLoad[p] +=
				landuse_getRunoffLoad(i, p, iArea, landFactor, qRunoff, tStep);
		else
			// calculate pollutant washoff on total landuse area
			washoffLoad[p] +=
				landuse_getRunoffLoad(i, p, fArea, landFactor, qRunoff, tStep);
    }

    // --- compute contribution from co-pollutant
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        washoffLoad[p] += landuse_getCoPollutLoad(p, washoffLoad, tStep);
    }
}

//=============================================================================

float landuse_getRunoffLoad(int i, int p, float area, TLandFactor landFactor[], 
                            float runoff, float tStep)
//
//  Input:   i = land use index
//           p = pollut. index
//           area = area devoted to land use (ft2)
//           landFactor[] = array of land use data for subcatchment
//           runoff = runoff flow on subcatchment (ft/sec)
//           tStep = time step (sec)
//  Output:  returns runoff load for pollutant (mass/sec)
//  Purpose: computes pollutant load generated by a specific land use.
//
{
    float buildup;
    float washoff;

    // --- compute washoff mass/sec for this pollutant
    buildup = landFactor[i].buildup[p];
    washoff = landuse_getWashoffMass(i, p, buildup, runoff, area);
		
	//added
	if(strncmp(strTSS, "", MAXFNAME) != 0 && Pollut[p].sedflag > 0) 
	{
		// check the sediment class
		if (Pollut[p].sedflag == 1)
		{
			// sand
			buildup *= Landuse[i].frc_sand;
			washoff *= Landuse[i].frc_sand;
		}
		else if (Pollut[p].sedflag == 2)
		{
			// silt
			buildup *= Landuse[i].frc_silt;
			washoff *= Landuse[i].frc_silt;
		}
		else if (Pollut[p].sedflag == 3)
		{
			// clay
			buildup *= Landuse[i].frc_clay;
			washoff *= Landuse[i].frc_clay;
		}
	}
			
    // --- convert washoff to lbs (or kg) over time step so that
    //     buildup and mass balances can be adjusted
    //     (Pollut[].mcf converts from concentration mass units
    //      to either lbs or kg)
    washoff *= tStep * Pollut[p].mcf;

    // --- if buildup modelled, reduce it by amount of washoff
    if ( Landuse[i].buildupFunc[p].funcType != NO_BUILDUP )
    {
        washoff = MIN(washoff, buildup);
        buildup -= washoff;
        landFactor[i].buildup[p] = buildup;
    }

    // --- otherwise add washoff to buildup mass balance totals
    //     so that things will balance
    else massbal_updateLoadingTotals(BUILDUP_LOAD, p, washoff);

    // --- return washoff converted back to mass/sec
    return washoff / tStep / Pollut[p].mcf;
}

//=============================================================================

float landuse_getWashoffMass(int i, int p, float buildup, float runoff,
      float area)
//
//  Input:   i = land use index
//           p = pollutant index
//           buildup = current buildup over land use (lbs or kg)
//           runoff = current runoff on subcatchment (ft/sec)
//           area = area devoted to land use (ft2)
//  Output:  returns pollutant washoff rate (mass/sec)
//  Purpose: finds mass loading of pollutant washed off a land use.
//
//  Notes:   "coeff" for each washoff function was previously adjusted to
//           result in units of mass/sec
//
{
    float washoff;
    float coeff = Landuse[i].washoffFunc[p].coeff;
    float expon = Landuse[i].washoffFunc[p].expon;
    int   func  = Landuse[i].washoffFunc[p].funcType;

    // --- if no washoff function, return 0
    if ( func == NO_WASHOFF ) return 0.0;
    
    // --- if buildup function exists but no current buildup, return 0
    if ( Landuse[i].buildupFunc[p].funcType != NO_BUILDUP && buildup == 0.0 )
        return 0.0;

    if ( func == EXPON_WASHOFF )
    {
        // --- convert runoff to inches/hr (or mm/hr) and 
        //     convert buildup from lbs (or kg) to concen. mass units
        runoff = runoff * UCF(RAINFALL);
        buildup /= Pollut[p].mcf;

        // --- evaluate washoff eqn.
        washoff = coeff * pow(runoff, expon) * buildup;
    }

    else if ( func == RATING_WASHOFF )
    {
        runoff = runoff * area;             // runoff in cfs
        if ( runoff == 0.0 ) washoff = 0.0;
        else washoff = coeff * pow(runoff, expon);
    }

    else if ( func == EMC_WASHOFF )
    {
        runoff = runoff * area;             // runoff in cfs
        washoff = coeff * runoff;           // coeff includes LperFT3 factor
    }

    else washoff = 0.0;

    return washoff;
}

//=============================================================================

float landuse_getCoPollutLoad(int p, float washoff[], float tStep)
//
//  Input:   p = pollutant index
//           washoff = pollut. washoff rate (mass/sec)
//           tStep = time step (sec)
//  Output:  returns washoff mass added by co-pollutant relation (mass/sec)
//  Purpose: finds washoff mass added by a co-pollutant of a given pollutant.
//
{
    int   k;
    float w;
    float load;

    // --- check if pollutant p has a co-pollutant k
    k = Pollut[p].coPollut;
    if ( k >= 0 )
    {
        // --- compute addition to washoff from co-pollutant
        w = Pollut[p].coFraction * washoff[k];

        // --- add to mass balance totals
        load = w * tStep * Pollut[p].mcf;
        massbal_updateLoadingTotals(BUILDUP_LOAD, p, load);
        return w;
    }
    return 0.0;
}

//=============================================================================
///////////////////////////////////////////////////
//  This function is no longer used. (LR - 7/5/06 )
///////////////////////////////////////////////////
//float landuse_getRainLoad(int p, float qPpt, float area, float tStep)
//
//  Input:   p = pollutant index
//           qPpt = rainfall/snowmelt portion of runoff (ft/sec)
//           area = area covered by a land use (ft2)
//           tStep = time step (sec)
//  Output:  returns washoff mass added by direct rainfall deposition (mass/sec)
//  Purpose: finds washoff mass added by direct deposition from rainfall.
//
//{
//    float w;
//    float load;
//
//    w = Pollut[p].pptConcen * qPpt * area * LperFT3;
//    load = w * tStep * Pollut[p].mcf;                      // convert to lbs (kg)
//    massbal_updateLoadingTotals(DEPOSITION_LOAD, p, load);
//    return w;
//}

//=============================================================================
///////////////////////////////////////////////////
//  This function is no longer used. (LR - 7/5/06 )
///////////////////////////////////////////////////
//float landuse_getBMPRemoval(int i, int p, float washoff, float tStep)
//
//  Input:   i = land use index
//           p = pollutant index
//           washoff = pollut. washoff rate (mass/sec)
//           tStep = time step (sec)
//  Output:  returns washoff mass after BMP removal (mass/sec)
//  Purpose: adjusts washoff loading to account for reduction by BMPs.
//
//{
//    float w;
//    float load;
//
//    w = washoff * Landuse[i].washoffFunc[p].bmpEffic;
//    load = w * tStep * Pollut[p].mcf;
//    massbal_updateLoadingTotals(BMP_REMOVAL_LOAD, p, load);
//    return washoff - w;
//}

//=============================================================================

float landuse_getDetached(int i,float area, float dets,float rainfall,float tStep)
//
//  Input:   i = land use index
//			 area = lannduse area (acres)
//           dets = sediment detached storage at the beginning of the time step (lb)
//           tStep = time step (sec)
//  Output:  returns updated sediment detached storage (lb)
//  Purpose: finds mass loading of sediment detached off a pervious land use.
//
{
	int csnofg = 0;
	int crvfg = 0;
	int mon = 1;
	int nxtmon = 1;
	int day = 1;
	int ndays = 31;
	float snocov = 0.0;
	float delt60 = tStep/3600.0;	// hrs/ivl
	float *coverm = NULL;
	float smpf  = Landuse[i].smpf;
	float krer  = Landuse[i].krer;
	float jrer  = Landuse[i].jrer;
	float cover = Landuse[i].cover;

	float det = 0.0;

	// --- compute rainfall+snowfall volume (does not include snowmelt)
	rainfall = rainfall * tStep * 12;	// in/ivl
	dets = dets / 2000.00 / area;		// tons/acre 
					
	detach(crvfg, csnofg, mon, nxtmon, day, ndays, coverm, rainfall, snocov, delt60,
		   smpf, krer, jrer, cover, dets, det);

	// update values
	Landuse[i].cover = cover;

	//convert units back to lbs
	dets = dets * 2000.00 * area;

    return dets;
} 

float landuse_getRemovalMass(int i,float area,float tStep,float rainfall,float runoff, 
							 float surs, float& dets)
{
	int vsivfg = 0;
	int drydfg = 0;
	int mon = 1;
	int nxtmon = 1;
	int day = 1;
	int ndays = 31;
	float nvsi = 0.0;
	float *nvsim = NULL;
	float delt60 = tStep/3600.0;		// number of hours per time step
	float deltd = tStep/(3600.0*24.0);	// number of days per time step
	float affix = Landuse[i].affix;
	float kser  = Landuse[i].kser;
	float jser  = Landuse[i].jser;
	float kger  = Landuse[i].kger;
	float jger  = Landuse[i].jger;

	float sosed = 0.0;	// tons/acre/ivl

	// --- compute rainfall+snowfall volume (does not include snowmelt)
	rainfall = rainfall * tStep * 12;	// in/ivl
	runoff = runoff * tStep * 12;		// in/ivl
	surs *= 12.0;						// inch
	area /= 43559.66;					// acre
	dets = dets / 2000.00 / area;		// tons/acre 
						
	//it is the first interval of the day
	if (vsivfg != 0)
	{
		//net vert. input values are allowed to vary throughout the
		//year
		//interpolate for the daily value
		//units are tons/acre-ivl
		//linearly interpolate nvsi between two values from the
		//monthly array nvsim(12)
		nvsi = dayval(nvsim[mon], nvsim[nxtmon], day, ndays);
	}
	else
	{
		//net vert. input values do not vary throughout the year.
		//nvsi value has been supplied by the run interpreter
	}

	if (vsivfg == 2) 
	{
		if (drydfg == 1)
		{
			//last day was dry, add a whole days load in first interval
			//detailed output will show load added over the whole day.
			float dummy = nvsi * deltd;
			dets = dets + dummy;
		}
		else 
		{
			//dont accumulate until tomorrow, maybe
			nvsi = 0.0;
		}
	}

	//augment the detached sediment storage by external(vertical)
	//inputs of sediment - dets and detsb units are tons/acre

	float slsed = 0.0;		// ??
	float dummy = slsed;
	
	if (vsivfg < 2) 
	  dummy = dummy + nvsi;

	dets += dummy;

	sosed1(runoff, surs, delt60, kser, jser, kger, jger, dets, sosed);

	if(rainfall <0.00001)
		attach(affix, deltd, dets);

	//convert units back to lbs
	dets = dets * 2000.00 * area;

    // --- return removal converted to lb/ivl
	return sosed * 2000.00 * area;
} 

void  landuse_getRemoval(int i, int j, float area, TLandFactor landFactor[],
                         float qRunoff, float tStep,float removalLoad[])
//
//  Input:   i            = land use index
//			 j            = subcatchment index
//           area         = subcatchment area (ft2)
//           landFactor[] = array of land use data for subcatchment
//           runoff       = runoff flow rate (ft/sec) over subcatchment
//           tStep        = time step (sec)
//  Output:  removalLoad[] = pollutant load in surface washoff (mass/sec)
//  Purpose: computes surface washoff load for all pollutants generated by a
//           land use within a subcatchment.
//
{
    int   p;                           //pollutant index
    float fArea;                       //area devoted to land use (ft2)
    float pArea;                       //pervious area devoted to land use (ft2)

    // --- find area devoted to land use
    fArea = landFactor[i].fraction * area;

    // --- find pervious area devoted to land use
    pArea = fArea * (1 - Landuse[i].pctimp);

    // --- initialize washoff loads from land use
    for (p = 0; p < Nobjects[POLLUT]; p++) removalLoad[p] = 0.0;

    // --- compute contribution from direct runoff load
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
		if (Pollut[p].sedflag > 0)	// the pollutant is sediment
			// calculate removal from the pervious area
			removalLoad[p] +=
				landuse_getRunoffLoad2(i, j, p, pArea, landFactor, qRunoff, tStep);
    }

	return;
}

float landuse_getRunoffLoad2(int i, int j, int p, float area, TLandFactor landFactor[], 
                            float runoff, float tStep)
//
//  Input:   i = land use index
//			 j = subcatchment index
//			 p = pollutant index
//           area = area devoted to land use (ft2)
//           landFactor[] = array of land use data for subcatchment
//           runoff = runoff flow on subcatchment (ft/sec)
//           tStep = time step (sec)
//  Output:  returns runoff load for pollutant (mass/sec)
//  Purpose: computes pollutant load generated by a specific land use.
//
{
	float rainfall;		// ft/sec
	float pondeddepth;	// ft
	float detstorage;	// lb	
	float removal;		// lb/ivl

    // --- compute washoff mass/sec for this pollutant
	rainfall = Subcatch[j].rainfall;
	pondeddepth = Subcatch[j].subArea[2].depth;
	detstorage = landFactor[i].detstorage[p];

    removal = landuse_getRemovalMass(i, area, tStep, rainfall, runoff, 
									 pondeddepth, detstorage);

	landFactor[i].detstorage[p] = detstorage;

	// check the sediment class
	if (Pollut[p].sedflag == 1)
		// sand
		removal *= Landuse[i].frc_sand;
	else if (Pollut[p].sedflag == 2)
		// silt
		removal *= Landuse[i].frc_silt;
	else if (Pollut[p].sedflag == 3)
		// clay
		removal *= Landuse[i].frc_clay;
			
    // --- return removal converted from lb/ivl to mass/sec
    return removal / tStep / Pollut[p].mcf;
}

//=============================================================================
