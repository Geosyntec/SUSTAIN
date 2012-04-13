//-----------------------------------------------------------------------------
//   infil.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman
//
//   Infiltration functions.
//-----------------------------------------------------------------------------

#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  External Functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  infil_readParams (called by input_readLine)
//  infil_initState  (called by subcatch_initState)
//  infil_getInfil   (called by getSubareaRunoff in subcatch.c)

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static int   horton_setParams(int j, float p[]);
static void  horton_initState(int j);
static float horton_getInfil(int j, float tstep, float irate, float depth);

//static int   grnampt_setParams(int j, float p[]);
static void  grnampt_initState(int j);
static float grnampt_getInfil(int j, float tstep, float irate, float depth);
static float grnampt_getRate(int j, float tstep, float F2, float F);
static float grnampt_getF2(float f1, float c1, float c2, float iv2);
static void  grnampt_setT(int j);

static int   curvenum_setParams(int j, float p[]);
static void  curvenum_initState(int j);
static float curvenum_getInfil(int j, float tstep, float irate, float depth);


//=============================================================================

int infil_readParams(int m, char* tok[], int ntoks)
//
//  Input:   m = infiltration method code
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: sets infiltration parameters from a line of input data.
//
//  Format of data line is:
//     subcatch  p1  p2 ...
{
    int   i, j, n, status;
    float x[5];

    // --- check that subcatchment exists
    j = project_findObject(SUBCATCH, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

    // --- number of input tokens depends on infiltration model m
    if      ( m == HORTON )       n = 5;
    else if ( m == GREEN_AMPT )   n = 4;
    else if ( m == CURVE_NUMBER ) n = 4;
    else return 0;
    if ( ntoks < n ) return error_setInpError(ERR_ITEMS, "");

    // --- parse numerical values from tokens
    for (i = 0; i < 5; i++) x[i] = 0.0;
    for (i = 1; i < n; i++)
    {
        if ( ! getFloat(tok[i], &x[i-1]) )
            return error_setInpError(ERR_NUMBER, tok[i]);
    }

    // --- special case for Horton infil. - last parameter is optional
    if ( m == HORTON && ntoks > n )
    {
        if ( ! getFloat(tok[n], &x[n-1]) )
            return error_setInpError(ERR_NUMBER, tok[n]);
    }

    // --- assign parameter values to infil. object
    Subcatch[j].infil = j;
    switch (m)
    {
      case HORTON:       status = horton_setParams(j, x);   break;
      case GREEN_AMPT:   status = grnampt_setParams(j, x);  break;
      case CURVE_NUMBER: status = curvenum_setParams(j, x); break;
      default:           status = TRUE;
    }
    if ( !status ) return error_setInpError(ERR_NUMBER, "");
    return 0;
}

//=============================================================================

void infil_initState(int j, int m)
//
//  Input:   j = subcatchment index
//           m = infiltration method code
//  Output:  none
//  Purpose: initializes state of infiltration for a subcatchment.
//
{
    switch (m)
    {
      case HORTON:       horton_initState(j);   break;
      case GREEN_AMPT:   grnampt_initState(j);  break;
      case CURVE_NUMBER: curvenum_initState(j); break;
    }
}

//=============================================================================

float infil_getInfil(int j, int m, float tstep, float irate, float depth)
//
//  Input:   j = subcatchment index
//           m = infiltration method code
//           tstep = runoff time step (sec)
//           irate = rainfall rate (ft/sec)
//           depth = depth of surface water on subcatchment (ft)
//  Output:  returns infiltration rate (ft/sec)
//  Purpose: computes infiltration rate depending on infiltration method.
//
{
    switch (m)
    {
      case HORTON:       return horton_getInfil(j, tstep, irate, depth);
      case GREEN_AMPT:   return grnampt_getInfil(j, tstep, irate, depth);
      case CURVE_NUMBER: return curvenum_getInfil(j, tstep, irate, depth);
      default:           return 0.0;
    }
}

//=============================================================================

int horton_setParams(int j, float p[])
//
//  Input:   j = subcatchment index
//           p[] = array of parameter values
//  Output:  returns TRUE if parameters are valid, FALSE otherwise
//  Purpose: assigns Horton infiltration parameters to a subcatchment.
//
{
    int k;
    for (k=0; k<5; k++) if ( p[k] < 0.0 ) return FALSE;

    // --- max. & min. infil rates (ft/sec)
    HortInfil[j].f0      = p[0] / UCF(RAINFALL);
    HortInfil[j].fmin    = p[1] / UCF(RAINFALL);

    // --- convert decay const. to 1/sec
    HortInfil[j].decay = p[2] / 3600.;

///////////////////////////////////////////////////////
////  Correction to conversion constant. (LR - 7/5/06 )
///////////////////////////////////////////////////////
    // --- convert drying time (days) to a regeneration const. (1/sec)
    //     assuming that former is time to reach 98% dry along an
    //     exponential drying curve
    if (p[3] == 0.0 ) p[3] = TINY;
    //HortInfil[j].regen = 0.02 / p[3] / SECperDAY;
    HortInfil[j].regen = -log(1.0-0.98) / p[3] / SECperDAY;

    // --- optional max. infil. capacity (ft) (p[4] = 0 if no value supplied)
    HortInfil[j].Fmax = p[4] / UCF(RAINDEPTH);
    if ( HortInfil[j].f0 < HortInfil[j].fmin ) return FALSE;
    return TRUE;
}

//=============================================================================

void horton_initState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: initializes time on Horton infiltration curve for a subcatchment.
//
{
    HortInfil[j].tp = 0.0;
}

//=============================================================================

float horton_getInfil(int j, float tstep, float irate, float depth)
//
//  Input:   j = subcatchment index
//           tstep =  runoff time step (sec),

//////////////////////////////////////////////
//  Updated definition of irate. (LR - 9/5/05)
//////////////////////////////////////////////
//           irate = net "rainfall" rate (ft/sec),
//                 = rainfall + snowmelt + runon - evaporation

//           depth = depth of ponded water (ft).
//  Output:  returns infiltration rate (ft/sec)
//  Purpose: computes Horton infiltration for a subcatchment.
//
{
    // --- assign local variables
    int   iter;
    float fa, fp = 0.0;
    float Fp, F1, t1, tlim, ex, kt;
    float FF, FF1, r;
    float fmin = HortInfil[j].fmin;
    float Fmax = HortInfil[j].Fmax;
    float tp   = HortInfil[j].tp;
    float df   = HortInfil[j].f0 - fmin;
    float kd   = HortInfil[j].decay;
    float kr   = HortInfil[j].regen;

    // --- special cases of no infil. or constant infil
    if ( df < 0.0 || kd < 0.0 || kr < 0.0 ) return 0.0;
    if ( df == 0.0 || kd == 0.0 )
    {
        fp = HortInfil[j].f0;
        fa = irate + depth / tstep;
        if ( fp > fa ) fp = fa;

//////////////////////////////////////////////
//  Limit fp to be non-negative. (LR - 9/5/05)
//////////////////////////////////////////////
        return MAX(0.0, fp);
    }

////////////////////////////////////////////
//  Modified code starts here. (LR - 9/5/05)

    // --- compute water available for infiltration
    fa = irate + depth / tstep;

    // --- case where there is water to infiltrate
    if ( fa > 0.0 )
    {

//  Modified code ends here. (LR - 9/5/05)
///////////////////////////////////////////

        // --- compute average infil. rate over time step
        t1 = tp + tstep;         // future cumul. time
        tlim = 16.0 / kd;        // for tp >= tlim, f = fmin
        if ( tp >= tlim )
        {
            Fp = fmin * tp + df / kd;
            F1 = Fp + fmin * tstep;
        }
        else
        {
            Fp = fmin * tp + df / kd * (1.0 - exp(-kd * tp));
            F1 = fmin * t1 + df / kd * (1.0 - exp(-kd * t1));
        }
        if (Fmax > 0.0)
        {
             if ( Fmax < Fp ) Fp = Fmax;
             if ( Fmax < F1 ) F1 = Fmax;
        }
        fp = (F1 - Fp) / tstep;

        // --- limit infil rate to available infil
        if ( fp > fa ) fp = fa;

        // --- if fp on flat portion of curve then increase tp by tstep
        if ( t1 > tlim ) tp = t1;

        // --- if infil < available capacity then increase tp by tstep
        else if ( fp < fa ) tp = t1;

        // --- if infil limited by available capcity then
        //     solve F(tp) - F1 = 0 using Newton-Raphson method
        else
        {
            F1 = Fp + fp * tstep;
            tp = tp + tstep / 2.0;
            for ( iter=1; iter<=20; iter++ )
            {
                kt = MIN( 60.0, kd*tp );
                ex = exp(-kt);
                FF = fmin * tp + df / kd * (1.0 - ex) - F1;
                FF1 = fmin + df * ex;
                r = FF / FF1;
                tp = tp - r;
                if ( fabs(r) <= 0.001 * tstep ) break;
            }
        }
    }

    // --- case where infil. capacity is regenerating; update tp.
    else if (kr > 0.0)
    {
        tp = -log(1.0 - exp(-kr * tstep)*(1.0 - exp(-kd * tp))) / kd;
    }
    HortInfil[j].tp = tp;
    return fp;
}

//=============================================================================

int grnampt_setParams(int j, float p[])
//
//  Input:   j = subcatchment index
//           p[] = array of parameter values
//  Output:  returns TRUE if parameters are valid, FALSE otherwise
//  Purpose: assigns Green-Ampt infiltration parameters to a subcatchment.
//
{
    float ksat;                        // sat. hyd. conductivity in in/hr

    if ( p[0] <= 0.0 || p[1] <= 0.0 || p[2] < 0.0 ) return FALSE;
    GAInfil[j].S      = p[0] / UCF(RAINDEPTH);   // Capillary suction head (ft)
    GAInfil[j].Ks     = p[1] / UCF(RAINFALL);    // Sat. hyd. conductivity (ft/sec)
    GAInfil[j].IMDmax = p[2];                    // Max. init. moisture deficit

    // --- find depth of upper soil zone (ft) using Mein's eqn.
    ksat = GAInfil[j].Ks * 12. * 3600.;
    GAInfil[j].L = 4.0 * sqrt(ksat) / 12.;

////////////////////////////////////////////
//  Definition of FUmax added. (LR - 9/5/05)
////////////////////////////////////////////
    // --- set max. water volume of upper layer
    GAInfil[j].FUmax = GAInfil[j].L * GAInfil[j].IMDmax;

    return TRUE;
}

//=============================================================================

void grnampt_initState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: initializes state of Green-Ampt infiltration for a subcatchment.
//
{
    GAInfil[j].IMD = GAInfil[j].IMDmax;
    GAInfil[j].F = 0.0;
    GAInfil[j].FU = GAInfil[j].L * GAInfil[j].IMD;
    GAInfil[j].Sat = FALSE;
    grnampt_setT(j);
}

//=============================================================================

float grnampt_getInfil(int j, float tstep, float irate, float depth)
//
//  Input:   j = subcatchment index
//           tstep =  runoff time step (sec),

/////////////////////////////////////////////
//  Updated definition of irate. (LR -9/5/05)
/////////////////////////////////////////////
//           irate = net "rainfall" rate to upper zone (ft/sec);
//                 = rainfall + snowmelt + runon - evaporation,
//                   does not include ponded water (added on below)

//           depth = depth of ponded water (ft).
//  Output:  returns infiltration rate (ft/sec)
//  Purpose: computes Green-Ampt infiltration for a subcatchment.
//
//  Definition of variables:
//   f      = infiltration rate (ft/sec)
//   IMD    = initial moisture deficit for rain event (ft/ft)
//   IMDmax = max. IMD available (ft/ft)
//   Ks     = saturated hyd. conductivity (ft/sec)
//   S      = capillary suction head (ft)
//   F      = cumulative event infiltration (ft)
//   Fmax   = max. allowable infiltration (ft)
//   T      = cumulative event duration (sec)
//   Tmax   = max. discrete event duration (sec)
//   L      = depth of upper soil zone (ft)
//   FU     = current moisture content of upper zone (ft)
//   FUmax  = saturated moisture content of upper zone (ft)
//   DF     = upper zone moisture depeletion factor (1/sec)
//
{
    // --- initialize infil. rate f to rainfall rate
    float F = GAInfil[j].F;
    float F2;
    float DF;
    float DV;
    float Fs;

/////////////////////////////////////////////////
//  Need to add standing water onto irate first 
//  before initializing f and ivol. (LR - 9/5/05)
/////////////////////////////////////////////////
    float f;      // = irate;
    float ivol;   // = irate * tstep;

    float c1 = GAInfil[j].S * GAInfil[j].IMD;

//////////////////////////////////////////////////////////////////
//  FUmax already defined as part of GAInfil object. (LR - 9/5/05)
//////////////////////////////////////////////////////////////////
    //float FUmax = GAInfil[j].L * GAInfil[j].IMDmax;

    float ts, c2, iv2;

///////////////////////////////////////////////////////////////////////////
//  Add ponded water onto irate and initialize f & ivol here. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////////////
    // --- add ponded water onto potential infiltration
    irate += depth / tstep;
    f = irate;
    ivol = irate * tstep;

    // --- upper soil zone is unsaturated
    if ( !GAInfil[j].Sat )
    {
        // --- update time remaining until upper zone is completely drained
        GAInfil[j].T -= tstep;

        // --- no rainfall; deplete soil moisture
        if ( irate <= 0.0 )
        {

////////////////////////////////////////
//  New code added. (LR via RD - 9/5/05)
////////////////////////////////////////
            // --- return if no upper zone moisture
            if ( GAInfil[j].FU <= 0.0 ) return 0.0;

            DF = GAInfil[j].L / 300. * (12. / 3600.);

///////////////////////////////////////////////////////
// Local FUmax replaced by GAInfil.FUmax. (LR - 9/5/05)
///////////////////////////////////////////////////////
            DV = DF * GAInfil[j].FUmax * tstep;

            GAInfil[j].F -= DV;
            GAInfil[j].FU -= DV;
            if ( GAInfil[j].FU <= 0.0 )
            {
                GAInfil[j].FU = 0.0;
                GAInfil[j].F = 0.0;
                GAInfil[j].IMD = GAInfil[j].IMDmax;

////////////////////////////////
// New code added. (LR - 9/5/05)
////////////////////////////////
                return 0.0;
            }

            // --- if upper zone drained, then redistribute moisture content
            if ( GAInfil[j].T <= 0.0 )
            {
//////////////////////////////////////////////////////////
//  Correction to IMD update formula. (LR via RD - 9/5/05)
//////////////////////////////////////////////////////////
                GAInfil[j].IMD = (GAInfil[j].FUmax - GAInfil[j].FU) /
                                  GAInfil[j].L;

                GAInfil[j].F = 0.0;
            }

////////////////////////////////////////////////////////////////////
//  Return with no infiltration since net inflow <= 0. (LR - 9/5/05)
////////////////////////////////////////////////////////////////////
            //return f;
            return 0.0;
        }

        // --- low rainfall; everything infiltrates
        if ( irate <= GAInfil[j].Ks )
        {
            F2 = F + ivol;
            f = grnampt_getRate(j, tstep, F2, F);

            // --- if sufficient time to drain upper zone, then redistribute
            if ( GAInfil[j].T <= 0.0 )
            {
//////////////////////////////////////////////////////////
//  Correction to IMD update formula. (LR via RD - 9/5/05)
//////////////////////////////////////////////////////////
                GAInfil[j].IMD = (GAInfil[j].FUmax - GAInfil[j].FU) /
                                  GAInfil[j].L;

                GAInfil[j].F = 0.0;
            }
            return f;
        }

        // --- rainfall > hyd. conductivity; renew time to drain upper zone
        grnampt_setT(j);

        // --- check if surface already saturated
        Fs = c1 * GAInfil[j].Ks / (irate - GAInfil[j].Ks);
        if ( F - Fs >= 0.0 )
        {
            GAInfil[j].Sat = TRUE;
        }

        // --- check if all water infiltrates
        else if ( Fs - F >= ivol )
        {
            F2 = F + ivol;
            f = grnampt_getRate(j, tstep, F2, F);
            return f;
        }

        // --- otherwise surface saturates during time interval
        else
        {
            ts  = (Fs - F) / irate;
            c2  = c1 * log(Fs + c1) - GAInfil[j].Ks * (tstep - ts);
            iv2 = (tstep - ts) * irate / 2.;
            F2  = grnampt_getF2(Fs, c1, c2, iv2);
            f   = grnampt_getRate(j, tstep, F2, Fs);
            GAInfil[j].Sat = TRUE;
            return f;
        }
    }

    // --- upper soil zone saturated:

    // --- renew time to drain upper zone
    grnampt_setT(j);

    // --- compute volume of potential infiltration
    if ( c1 <= 0.0 ) F2 = GAInfil[j].Ks * tstep + F;
    else
    {
        c2 = c1 * log(F + c1) - GAInfil[j].Ks * tstep;
        iv2 = tstep * irate / 2.;
        F2 = grnampt_getF2(F, c1, c2, iv2);
    }

    // --- excess water will remain on surface

//////////////////////////////////////////
//  ivol now includes depth. (LR - 9/5/05)
//////////////////////////////////////////
    //if ( F2 - F <= ivol + depth )
    if ( F2 - F <= ivol )

    {
        f = grnampt_getRate(j, tstep, F2, F);
        return f;
    }

    // --- all rain + ponded water infiltrates

//////////////////////////////////////////
//  ivol now includes depth. (LR - 9/5/05)
//////////////////////////////////////////
    //F2 = F + ivol + depth;
    F2 = F + ivol;

    f = grnampt_getRate(j, tstep, F2, F);
    GAInfil[j].Sat = FALSE;
    return f;
}

//=============================================================================

float grnampt_getRate(int j, float tstep, float F2, float F)
//
//  Input:   j = subcatchment index
//           tstep =  runoff time step (sec),
//           F2 = new cumulative event infiltration volume (ft)
//           F = old cumulative event infiltration volume (ft)
//  Output:  returns infiltration rate (ft/sec)
//  Purpose: computes infiltration rate from change in infiltration volume.
//
{
    float f = (F2 - GAInfil[j].F) / tstep;
    GAInfil[j].FU += F2 - F;
    if ( GAInfil[j].FU > GAInfil[j].FUmax ) GAInfil[j].FU = GAInfil[j].FUmax;
    GAInfil[j].F = F2;
    return f;
}

//=============================================================================

float grnampt_getF2(float f1, float c1, float c2, float iv2)
//
//  Input:   f1 = old infiltration volume (ft)
//           c1, c2 =  equation terms
//           iv2 = half of rainfall volume over time step (ft)
//  Output:  returns infiltration volume at end of time step (ft)
//  Purpose: computes new infiltration volume over a time step
//           using Green-Ampt formula for saturated upper soil zone
//
{
    int   i;
    float f2 = f1;
    float df2;

    // --- use Newton-Raphson method to solve governing nonlinear equation
    for ( i = 1; i <= 20; i++ )
    {
        df2 = (f2 - f1 - c1 * log(f2 + c1) + c2) / (1.0 - c1 / (f2 + c1) );
        if ( fabs(df2) < 0.0001 ) return f2;
        f2 -= df2;
    }
    return f1 + iv2;
}

//=============================================================================

void grnampt_setT(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: resets maximum time to drain upper soil zone for Green-Ampt
//           infiltration.
//
{
    float DF = GAInfil[j].L / 300.0 * (12. / 3600.);
    GAInfil[j].T = 6.0 / (100.0 * DF);
}

//=============================================================================

int curvenum_setParams(int j, float p[])
//
//  Input:   j = subcatchment index
//           p[] = array of parameter values
//  Output:  returns TRUE if parameters are valid, FALSE otherwise
//  Purpose: assigns Curve Number infiltration parameters to a subcatchment.
//
{
    float ksat;

    // --- convert Curve Number to max. infil. capacity
    if ( p[0] < 10.0 ) p[0] = 10.0;
    if ( p[0] > 99.0 ) p[0] = 99.0;
    CNInfil[j].Smax    = (1000.0 / p[0] - 10.0) / 12.0;
    if ( CNInfil[j].Smax < 0.0 ) return FALSE;

    // --- compute inter-event time (sec) from hyd. conductivity
    if (p[1] > 0.0)
    {
        ksat = p[1] * UCF(RAINFALL) * 12.0 * 3600.0;
        CNInfil[j].Tmax = 4.5 / sqrt(ksat) * 3600.0;
    }
    else CNInfil[j].Tmax = 6.0 * 3600.0;

    // --- convert drying time (days) to a regeneration const. (1/sec)
    //     assuming that former is time to reach 98% dry along an
    //     exponential drying curve
    if ( p[2] > 0.0 )
    {
        CNInfil[j].regen = -log(1.0-0.98) / p[2] / SECperDAY;
    }
    else CNInfil[j].regen = 0.0;

    return TRUE;
}

//=============================================================================

void curvenum_initState(int j)
//
//  Input:   j = subcatchment index
//  Output:  none
//  Purpose: initializes state of Curve Number infiltration for a subcatchment.
//
{
    CNInfil[j].S  = CNInfil[j].Smax;
    CNInfil[j].P  = 0.0;
    CNInfil[j].F  = 0.0;
    CNInfil[j].T  = 0.0;
    CNInfil[j].Se = CNInfil[j].Smax;
    CNInfil[j].f  = 0.0;
}

//=============================================================================

float curvenum_getInfil(int j, float tstep, float irate, float depth)
//
//  Input:   j = subcatchment index
//           tstep = runoff time step (sec),

/////////////////////////////////////////////
//  Updated definition of irate. (LR -9/5/05)
/////////////////////////////////////////////
//           irate = net "rainfall" rate (ft/sec);
//                 = rainfall + snowmelt + runon - evaporation,

//           depth = depth of ponded water (ft)
//  Output:  returns infiltration rate (ft/sec)
//  Purpose: computes infiltration rate using the Curve Number method.
//
{
    float F1;                          // new cumulative infiltration (ft)
    float f1 = 0.0;                    // new infiltration rate (ft/sec)
    float fa = irate + depth/tstep;    // max. available infil. rate (ft/sec)

    // --- case where there is rainfall
    if ( irate > 0.0 )
    {
        // --- check if new rain event
        if ( CNInfil[j].T >= CNInfil[j].Tmax )
        {
            CNInfil[j].P = 0.0;
            CNInfil[j].F = 0.0;
            CNInfil[j].f = 0.0;
            CNInfil[j].Se = CNInfil[j].S;
        }
        CNInfil[j].T = 0.0;

        // --- update cumul. precip. & cumul. infil.
        CNInfil[j].P += irate * tstep;
        F1 = CNInfil[j].P * (1.0 - CNInfil[j].P /
             (CNInfil[j].P + CNInfil[j].Se));

        // --- compute infil. rate
        f1 = (F1 - CNInfil[j].F) / tstep;
        if ( f1 < 0.0 ) f1 = 0.0;
        CNInfil[j].F += f1 * tstep;
    }

    // --- case of no rainfall
    else
    {
        // --- update inter-event time
        CNInfil[j].T += tstep;

        // --- if there is ponded water then use previous infil. rate
        if ( depth > MIN_TOTAL_DEPTH ) f1 = CNInfil[j].f;
    }

    // --- if there is some infiltration
    if ( f1 > 0.0 )
    {
        // --- limit infil rate to max. available rate
        f1 = MIN(f1, fa);

//////////////////////////////////////////
//  Make f1 be non-negative. (LR - 9/5/05)
//////////////////////////////////////////
        f1 = MAX(f1, 0.0);

        // --- reduce infil. capacity if a regen. constant was supplied
        if ( CNInfil[j].regen > 0.0 )
        {
            CNInfil[j].S -= f1 * tstep;
            if ( CNInfil[j].S < 0.0 ) CNInfil[j].S = 0.0;
        }
    }

    // --- otherwise regenerate infil. capacity
    else
    {
        CNInfil[j].S += CNInfil[j].regen * (CNInfil[j].Smax - CNInfil[j].S)
                        * tstep;
    }
    CNInfil[j].f = f1;
    return f1;
}

//=============================================================================
