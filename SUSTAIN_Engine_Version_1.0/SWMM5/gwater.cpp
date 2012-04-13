//-----------------------------------------------------------------------------
//   gwater.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//   Author:   L. Rossman
//
//   Groundwater functions.
//-----------------------------------------------------------------------------

#include <stdlib.h>
#include <math.h>
#include "headers.h"
#include "odesolve.h"

//-----------------------------------------------------------------------------
//  Constants
//-----------------------------------------------------------------------------
static const float GWTOL = 0.0001;     // ODE solver tolerance
static const float XTOL  = 0.001;      // tolerance on moisture & depth
enum   GWstates {THETA,                // moisture content of upper GW zone
                 LOWERDEPTH};          // depth of lower sat. GW zone

//-----------------------------------------------------------------------------
//  Shared variables
//-----------------------------------------------------------------------------
//  NOTE: all flux rates are in ft/sec, all depths are in ft.
static float    Infil;            // infiltration rate from surface
static float    MaxEvap;          // max. evaporation rate
static float    AvailEvap;        // available evaporation rate
static float    UpperEvap;        // evaporation rate from upper GW zone
static float    LowerEvap;        // evaporation rate from lower GW zone
static float    UpperPerc;        // percolation rate from upper to lower zone
static float    LowerLoss;        // loss rate from lower GW zone
static float    GWFlow;           // flow rate from lower zone to conveyance node
static float    MaxUpperPerc;     // upper limit on UpperPerc
static float    MaxGWFlowPos;     // upper limit on GWFlow when its positve
static float    MaxGWFlowNeg;     // upper limit on GWFlow when its negative
static float    FracPerv;         // fraction of surface that is pervious
static float    TotalDepth;       // total depth of GW aquifer
static float    NodeInvert;       // elev. of conveyance node invert
static float    NodeDepth;        // current water depth at conveyance node
/////////////////////////////////////////////////////////////////////////////
////Added (KA - 05/30/07)
/////////////////////////////////////////////////////////////////////////////
static float    ThetaMacropore;	  // moisture content in macropores

static TAquifer A;                // aquifer being analyzed
static TGroundwater* GW;          // groundwater object being analyzed

//-----------------------------------------------------------------------------
//  External Functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  gwater_readAquiferParams     (called by input_readLine)
//  gwater_readGroundwaterParams (called by input_readLine)
//  gwater_validateAquifer       (called by swmm_open)
//  gwater_initState             (called by subcatch_initState)
//  gwater_getVolume             (called by massbal_open & massbal_getGwaterError)
//  gwater_getGroundwater        (called by getSubareaRunoff in subcatch.c)

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void  getDxDt(float t, float* x, float* dxdt);
static float getExcessInfil(float* x, float tStep);
static void  getFluxes(float upperVolume, float lowerDepth);
static void  getEvapRates(float theta, float upperDepth);
static float getUpperPerc(float theta, float upperDepth);
static float getGWFlow(float lowerDepth);
static void  updateMassBal(float area,  float tStep);


//=============================================================================

int gwater_readAquiferParams(int j, char* tok[], int ntoks)
//
//  Input:   j = aquifer index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns error message
//  Purpose: reads aquifer parameter values from line of input data
//
//  Data line contains following parameters:
//    ID, porosity, wiltingPoint, fieldCapacity,     conductivity,
//    conductSlope, tensionSlope, upperEvapFraction, lowerEvapDepth,
//    gwRecession,  bottomElev,   waterTableElev,    upperMoisture
//    macroporosity
//
{
    int   i;
    float x[13];	// added one for macroporosity
    char *id;

    // --- check that aquifer exists
    if ( ntoks < 13 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(AQUIFER, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);

    // --- read remaining tokens as floats
    for (i = 0; i < 13; i++) x[i] = 0.0;
    for (i = 1; i < 14; i++)
    {
        if ( ! getFloat(tok[i], &x[i-1]) )
            return error_setInpError(ERR_NUMBER, tok[i]);
    }

    // --- assign parameters to aquifer object
    Aquifer[j].ID = id;
    Aquifer[j].porosity       = x[0];
    Aquifer[j].wiltingPoint   = x[1];
    Aquifer[j].fieldCapacity  = x[2];
    Aquifer[j].conductivity   = x[3] / UCF(RAINFALL);
    Aquifer[j].conductSlope   = x[4];
    Aquifer[j].tensionSlope   = x[5] / UCF(LENGTH);
    Aquifer[j].upperEvapFrac  = x[6];
    Aquifer[j].lowerEvapDepth = x[7] / UCF(LENGTH);
    Aquifer[j].lowerLossCoeff = x[8] / UCF(RAINFALL);
    Aquifer[j].bottomElev     = x[9] / UCF(LENGTH);
    Aquifer[j].waterTableElev = x[10] / UCF(LENGTH);
    Aquifer[j].upperMoisture  = x[11];
    Aquifer[j].macroporosity  = x[12];
    return 0;
}

//=============================================================================

int gwater_readGroundwaterParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns error code
//  Purpose: reads groundwater inflow parameters for a subcatchment from
//           a line of input data.
//
////////////////////////////////////////////////////////////////////
//  Revised input data format (node elev. added as x7) (LR - 9/5/05)
////////////////////////////////////////////////////////////////////
//  Data format is:
//    subcatch  aquifer  node  surfElev  x0 ... x7 (flow parameters)
//
{
    int   i, j, k, n;

///////////////////////////////////////////////////////////////////////////
//  Parameter vector size increased to accommodate node elev. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////////////
    float x[8];

    TGroundwater* gw;

    // --- check that specified subcatchment, aquifer & node exist
    if ( ntoks < 10 ) return error_setInpError(ERR_ITEMS, "");
    j = project_findObject(SUBCATCH, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);
    k = project_findObject(AQUIFER, tok[1]);
    if ( k < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n = project_findObject(NODE, tok[2]);
    if ( n < 0 ) return error_setInpError(ERR_NAME, tok[2]);

    // --- read in the groundwater flow parameters as floats
    for ( i = 0; i < 7; i++ )
    {
        if ( ! getFloat(tok[i+3], &x[i]) ) 
            return error_setInpError(ERR_NUMBER, tok[i+3]);
    }

//////////////////////////////////////////////////////////////////////
//  Code added to read in optional node elevation value. (LR - 9/5/05)
//////////////////////////////////////////////////////////////////////
    // --- read in overridden node invert elev.
    x[7] = MISSING;
    if ( ntoks > 10 )
    {
        if ( ! getFloat(tok[10], &x[7]) ) 
            return error_setInpError(ERR_NUMBER, tok[10]);
        x[7] /= UCF(LENGTH);
    }

    // --- create a groundwater flow object
    if ( !Subcatch[j].groundwater )
    {
        gw = (TGroundwater *) malloc(sizeof(TGroundwater));
        if ( !gw ) return error_setInpError(ERR_MEMORY, "");
        Subcatch[j].groundwater = gw;
    }
    else gw = Subcatch[j].groundwater;

    // --- populate the groundwater flow object with its parameters
    gw->aquifer    = k;
    gw->node       = n;
    gw->surfElev   = x[0] / UCF(LENGTH);
    gw->a1         = x[1];
    gw->b1         = x[2];
    gw->a2         = x[3];
    gw->b2         = x[4];
    gw->a3         = x[5];
    gw->fixedDepth = x[6] / UCF(LENGTH);

///////////////////////////////////////////////////////////////
// Add overridden node invert elev. to gw object. (LR - 9/5/05)
///////////////////////////////////////////////////////////////
    gw->nodeElev   = x[7];

    return 0;
}

//=============================================================================

void  gwater_validateAquifer(int j)
//
//  Input:   j = aquifer index
//  Output:  none
//  Purpose: validates groundwater aquifer properties .
//
{
    if ( Aquifer[j].porosity          <= 0.0 
    ||   Aquifer[j].macroporosity     <  0.0
    ||   Aquifer[j].fieldCapacity     >= Aquifer[j].porosity
    ||   Aquifer[j].wiltingPoint      >= Aquifer[j].fieldCapacity
    ||   Aquifer[j].conductivity      <= 0.0
    ||   Aquifer[j].conductSlope      <  0.0
    ||   Aquifer[j].tensionSlope      <  0.0
    ||   Aquifer[j].upperEvapFrac     <  0.0
    ||   Aquifer[j].lowerEvapDepth    <  0.0
    ||   Aquifer[j].waterTableElev    <  Aquifer[j].bottomElev
    ||   Aquifer[j].upperMoisture     >  Aquifer[j].porosity 
    ||   Aquifer[j].upperMoisture     <  Aquifer[j].wiltingPoint )
        report_writeErrorMsg(ERR_AQUIFER_PARAMS, Aquifer[j].ID);
}

//=============================================================================

void  gwater_initState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: initializes state of subcatchment's groundwater.
//
{
    TAquifer a;
    TGroundwater* gw;
    
    gw = Subcatch[j].groundwater;
    if ( gw )
    {
        a = Aquifer[gw->aquifer];
        gw->theta = a.upperMoisture;
        if ( gw->theta >= a.porosity )
        {
            gw->theta = a.porosity - XTOL;
        }
		gw->theta_mac = gw->theta * a.macroporosity/a.porosity;

        gw->lowerDepth = a.waterTableElev - a.bottomElev;
        if ( gw->lowerDepth >= gw->surfElev - a.bottomElev )
        {
            gw->lowerDepth = gw->surfElev - a.bottomElev - XTOL;
        }
        gw->oldFlow = 0.0;
        gw->newFlow = 0.0;
    }
}

//=============================================================================

float gwater_getVolume(int j)
//
//  Input:   j = subcatchment index
//  Output:  returns total volume of groundwater in ft/ft2
//  Purpose: finds volume of groundwater stored in upper & lower zones
//
{
    TAquifer a;
    TGroundwater* gw;
    float upperDepth;
    gw = Subcatch[j].groundwater;
    if ( gw == NULL ) return 0.0;
    a = Aquifer[gw->aquifer];
    upperDepth = gw->surfElev - a.bottomElev - gw->lowerDepth;

///////////////////////////////////////////
//  Following line corrected. (LR - 9/5/05)
///////////////////////////////////////////
    //return (upperDepth / gw->theta) + (gw->lowerDepth / a.porosity);
    return (upperDepth * gw->theta) + (gw->lowerDepth * a.porosity);
}

//=============================================================================

void gwater_getGroundwater(int j, float evap, float* infil, float tStep)
//
//  Input:   j     = subcatchment index
//           evap  = available evaporation (ft/sec)
//           infil = infiltration rate (ft/sec)
//           tStep = time step (sec)
//  Output:  infil = excess infiltration (ft/sec)
//  Purpose: finds groundwater flow from subcatchment during current time step.
//
{
    int   n;                           // node exchanging groundwater
    float x[2];                        // upper moisture content & lower depth 
    float xInfil;                      // excess infiltration
    float vUpper;                      // upper vol. available for percolation
    float nodeFlow;                    // max. possible GW flow from node

    // --- save subcatchment's groundwater and aquifer objects to 
    //     shared variables
    GW = Subcatch[j].groundwater;
    if ( GW == NULL ) return;
    A = Aquifer[GW->aquifer];

    // --- save fract. pervious, surface evap., & infil. to shared variables
    FracPerv = 1.0 - Subcatch[j].fracImperv;
    if ( FracPerv <= 0.0 ) return;
    MaxEvap = Evap.rate;
    AvailEvap = evap; 
    Infil = (*infil);

    // --- save total depth & outlet node properties to shared variables
    TotalDepth = GW->surfElev - A.bottomElev;
    if ( TotalDepth <= 0.0 ) return;
    n = GW->node;

/////////////////////////////////////////////////////////////////////////////
//  New code added for overriding receiving node's invert elev. (LR - 9/5/05)
/////////////////////////////////////////////////////////////////////////////
    // --- override node's invert if value was provided in the GW object
    if ( GW->nodeElev != MISSING ) NodeInvert = GW->nodeElev;
    else NodeInvert = Node[n].invertElev;
    
    if ( GW->fixedDepth > 0.0 ) NodeDepth = GW->fixedDepth;
    else                        NodeDepth = Node[n].newDepth;

    // --- store state variables in work vector x
    x[THETA] = GW->theta;
    x[LOWERDEPTH] = GW->lowerDepth;
	ThetaMacropore = GW->theta_mac;

    // --- set limits on upper perc
/////////////////////////////////////////////////////////////////////////////
////Added (KA - 05/30/07)
//  code modified for macropores
/////////////////////////////////////////////////////////////////////////////
    //vUpper = (TotalDepth - x[LOWERDEPTH]) * (x[THETA] - A.fieldCapacity);
    vUpper = (TotalDepth - x[LOWERDEPTH]) * (ThetaMacropore - A.fieldCapacity);
    
	vUpper = MAX(0.0, vUpper); 
    MaxUpperPerc = vUpper / tStep;

    // --- set limit on GW flow out of aquifer based on volume of lower zone
    MaxGWFlowPos = x[LOWERDEPTH]*A.porosity / tStep;

    // --- set limit on GW flow into aquifer from drainage system node
    //     based on min. of capacity of upper zone and drainage system
    //     inflow to the node
    MaxGWFlowNeg = (TotalDepth - x[LOWERDEPTH]) * (A.porosity - x[THETA])
                   / tStep;
    nodeFlow = (Node[n].inflow + Node[n].newVolume/tStep) / Subcatch[j].area;
    MaxGWFlowNeg = -MIN(MaxGWFlowNeg, nodeFlow);
    

    // --- limit infiltration to not exceed upper zone capacity
    xInfil = MIN(Infil, getExcessInfil(x, tStep));
    if ( xInfil > 0.0 )
    {
        Infil -= xInfil;
        *infil = xInfil;
    }

    // --- integrate eqns. for d(Theta)/dt and d(LowerDepth)/dt
    //     NOTE: ODE solver must have been initialized previously
    odesolve_integrate(x, 2, 0, tStep, GWTOL, tStep, getDxDt);
    
    // --- keep state variables within allowable bounds
    x[THETA] = MAX(x[THETA], A.wiltingPoint);
    if ( x[THETA] >= A.porosity )
    {
        x[THETA] = A.porosity - XTOL;
    }
    x[LOWERDEPTH] = MAX(x[LOWERDEPTH],  0.0);
    if ( x[LOWERDEPTH] >= TotalDepth )
    {
        x[LOWERDEPTH] = TotalDepth - XTOL;
    }

    // --- save new state values
    GW->theta = x[THETA];
	GW->theta_mac = GW->theta * A.macroporosity/A.porosity;
	ThetaMacropore = GW->theta_mac;

    GW->lowerDepth  = x[LOWERDEPTH];
    getFluxes(GW->theta, GW->lowerDepth);
    GW->oldFlow = GW->newFlow;
    GW->newFlow = GWFlow;

    // --- update mass balance
    updateMassBal(Subcatch[j].area, tStep);
}

//=============================================================================

void updateMassBal(float area, float tStep)
//
//  Input:   area  = subcatchment area (ft2)
//           tStep = time step (sec)
//  Output:  none
//  Purpose: updates GW mass balance with volumes of water fluxes.
//
{
    float vInfil;                      // infiltration volume
    float vUpperEvap;                  // upper zone evap. volume
    float vLowerEvap;                  // lower zone evap. volume
    float vLowerPerc;                  // lower zone deep perc. volume
    float vGwater;                     // volume of exchanged groundwater
    float ft2sec = area * tStep;

    vInfil     = Infil * FracPerv * ft2sec;
    vUpperEvap = UpperEvap * FracPerv * ft2sec;
    vLowerEvap = LowerEvap * FracPerv * ft2sec;
    vLowerPerc = LowerLoss * ft2sec;
    vGwater    = 0.5 * (GW->oldFlow + GW->newFlow) * ft2sec;
    massbal_updateGwaterTotals(vInfil, vUpperEvap, vLowerEvap, vLowerPerc,
                               vGwater);
}

//=============================================================================

float getExcessInfil(float* x, float tStep)
//
//  Input:   x = array with upper vol., lower vol., lower depth
//           tStep = time step (sec)
//  Output:  returns excess infiltration rate (ft/sec)
//  Purpose: finds execess infiltration into upper zone.
//
{
    float upperDepth;
    float availUpperVol;
    float upperPerc;
    float xInfil;

    upperDepth = TotalDepth - x[LOWERDEPTH];
	
/////////////////////////////////////////////////////////////////////////////
////Added (KA - 05/30/07)
//  code modified for macropores.
/////////////////////////////////////////////////////////////////////////////
    //upperPerc = getUpperPerc(x[THETA], upperDepth);
    upperPerc = getUpperPerc(ThetaMacropore, upperDepth);

    upperPerc = MIN(upperPerc, MaxUpperPerc);
    availUpperVol = upperDepth*(A.porosity - x[THETA]) - upperPerc*tStep;
    xInfil = Infil*tStep - availUpperVol/FracPerv;
    return MAX(0.0, xInfil);
}

//=============================================================================

void  getFluxes(float theta, float lowerDepth)
//
//  Input:   upperVolume = vol. depth of upper zone (ft)
//           upperDepth  = depth of upper zone (ft)
//  Output:  none
//  Purpose: computes water fluxes into/out of upper/lower GW zones.
//
{
    float upperDepth;

    // --- find upper zone depth
    lowerDepth = MAX(lowerDepth, 0.0);
    lowerDepth = MIN(lowerDepth, TotalDepth);
    upperDepth = TotalDepth - lowerDepth;

    // --- find evaporation from both zones
    getEvapRates(theta, upperDepth);

    // --- find percolation rate at upper & lower zone boundaries
/////////////////////////////////////////////////////////////////////////////
////Added (KA - 05/30/07)
//  code modified for macropores.
/////////////////////////////////////////////////////////////////////////////
    //UpperPerc = getUpperPerc(theta, upperDepth);
    UpperPerc = getUpperPerc(ThetaMacropore, upperDepth);

    UpperPerc = MIN(UpperPerc, MaxUpperPerc);

    // --- find losses to deep GW
    LowerLoss = A.lowerLossCoeff * lowerDepth / TotalDepth;

    // --- find GW flow from lower zone to conveyance system node
    GWFlow = getGWFlow(lowerDepth);
    if ( GWFlow >= 0.0 ) GWFlow = MIN(GWFlow, MaxGWFlowPos);
    else GWFlow = MAX(GWFlow, MaxGWFlowNeg);
}

//=============================================================================

void  getDxDt(float t, float* x, float* dxdt)
//
//  Input:   t    = current time (not used)
//           x    = array of state variables
//  Output:  dxdt = array of time derivatives of state variables
//  Purpose: computes time derivatives of upper moisture content 
//           and lower depth.
//
{
    float qUpper, qLower;

    getFluxes(x[THETA], x[LOWERDEPTH]);
    qUpper = (Infil - UpperEvap)*FracPerv - UpperPerc;
    qLower = UpperPerc - LowerLoss - (LowerEvap*FracPerv) - GWFlow;

    dxdt[THETA] = qUpper / (TotalDepth - x[LOWERDEPTH]);
    dxdt[LOWERDEPTH] = qLower / (A.porosity - x[THETA]);
}

//=============================================================================

void getEvapRates(float theta, float upperDepth)
//
//  Input:   theta      = moisture content of upper zone
//           upperDepth = depth of upper zone (ft)
//  Output:  none
//  Purpose: computes evapotranspiration out of upper & lower zones.
//
{
    float lowerFrac;
    UpperEvap = A.upperEvapFrac * MaxEvap;
    if ( theta <= A.wiltingPoint || Infil > 0.0 ) UpperEvap = 0.0;
    else UpperEvap = MIN(UpperEvap, AvailEvap);
    if ( A.lowerEvapDepth == 0.0 ) LowerEvap = 0.0;
    else
    {
        lowerFrac = (A.lowerEvapDepth - upperDepth) / A.lowerEvapDepth;
        lowerFrac = MAX(0.0, lowerFrac);
        LowerEvap = (1.0 - A.upperEvapFrac) * MaxEvap * lowerFrac;
        LowerEvap = MIN(LowerEvap, (AvailEvap - UpperEvap));
        LowerEvap = MAX(0.0, LowerEvap);
    }
}

//=============================================================================

float getUpperPerc(float theta, float upperDepth)
//
//  Input:   theta      = moisture content of upper zone
//           upperDepth = depth of upper zone (ft)
//  Output:  returns percolation rate (ft/sec)
//  Purpose: finds percolation rate from upper to lower zone.
//
{
    float delta;                       // unfilled water content of upper zone
    float dhdz;                        // avg. change in head with depth
    float hydcon;                      // unsaturated hydraulic conductivity

    // --- no perc. from upper zone if no depth or moisture content too low    
    if ( upperDepth <= 0.0 || theta <= A.fieldCapacity ) return 0.0;

    // --- compute hyd. conductivity as function of moisture content
    delta = theta - A.porosity;
    hydcon = A.conductivity * exp(delta * A.conductSlope);

    // --- compute integral of dh/dz term
    delta = theta - A.fieldCapacity;
    dhdz = 1.0 + A.tensionSlope * 2.0 * delta / upperDepth;

    // --- compute upper zone percolation rate
    return hydcon * dhdz;
}

//=============================================================================

float getGWFlow(float lowerDepth)
//
//  Input:   lowerDepth = depth of lower zone (ft)
//  Output:  returns groundwater flow rate (ft/sec)
//  Purpose: finds groundwater outflow from lower saturated zone.
//
{
    float t1, t2, t3;
    float gwElev;

    // --- water table must be above node invert for flow to occur
    gwElev = A.bottomElev + lowerDepth;
    if ( gwElev < NodeInvert ) return 0.0;

    // --- compute groundwater component of flow
    if ( GW->b1 == 0.0 ) t1 = GW->a1;
    else t1 = GW->a1 * pow( (gwElev - NodeInvert)*UCF(LENGTH), GW->b1);

    // --- compute surface water component of flow
    if ( GW->b2 == 0.0 ) t2 = GW->a2;
    else t2 = GW->a2 * pow(NodeDepth*UCF(LENGTH), GW->b2);

    // --- compute groundwater/surface water interaction term
    t3 = GW->a3 * lowerDepth * (NodeInvert + NodeDepth - A.bottomElev)
         * UCF(LENGTH) * UCF(LENGTH);
    return (t1 - t2 + t3) / UCF(RAINFALL);
}

//=============================================================================
