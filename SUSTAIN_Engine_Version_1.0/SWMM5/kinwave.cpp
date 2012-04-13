//-----------------------------------------------------------------------------
//   kinwave.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             3/10/06  (Build 5.0.007)
//   Author:   L. Rossman
//
//   Kinematic wave flow routing functions.
//-----------------------------------------------------------------------------

#include <math.h>
#include "headers.h"
#include "findroot.h"

//-----------------------------------------------------------------------------
//  Constants 
//-----------------------------------------------------------------------------
static const float WX      = 0.6;      // distance weighting
static const float WT      = 0.6;      // time weighting
static const float EPSIL   = 0.001;    // convergence criterion

//-----------------------------------------------------------------------------
//  Shared variables
//-----------------------------------------------------------------------------
static float   Beta1;
static float   C1;
static float   C2;
static float   Afull;
static float   Qfull;
static TXsect* pXsect;

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  kinwave_execute  (called by flowrout_execute)

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static int   solveContinuity(float qin, float ain, float* aout);
static void  evalContinuity(float a, float* f, float* df);

//=============================================================================

int kinwave_execute(int j, float* qinflow, float* qoutflow, float tStep)
//
//  Input:   j = link index
//           qinflow = inflow at current time (cfs)
//           tStep = time step (sec)
//  Output:  qoutflow = outflow at current time (cfs),
//           returns number of iterations used
//  Purpose: finds outflow over time step tStep given flow entering a
//           conduit using Kinematic Wave flow routing.
//
//  t
//  |          qin, ain |-------------------| qout, aout
//  |                   |  Flow --->        |
//  |----> x     q1, a1 |-------------------| q2, a2
//
//
{
    int   k;
    int   result = 1;
    float dxdt, dq;
    float ain, aout;
    float qin, qout;
    float a1, a2, q1, q2;

    // --- no routing for non-conduit link
    (*qoutflow) = (*qinflow); 
    if ( Link[j].type != CONDUIT ) return result;

    // --- no routing for dummy xsection
    if ( Link[j].xsect.type == DUMMY ) return result;

    // --- assign module-level variables
    pXsect = &Link[j].xsect;
    Qfull = Link[j].qFull;
    Afull = Link[j].xsect.aFull;
    k = Link[j].subIndex;
    Beta1 = Conduit[k].beta / Qfull;
 
    // --- normalize flows and areas
    q1 = Conduit[k].q1 / Qfull;
    q2 = Conduit[k].q2 / Qfull;
    a1 = Conduit[k].a1 / Afull;
    a2 = Conduit[k].a2 / Afull;
    qin = (*qinflow) / Conduit[k].barrels / Qfull;

    // --- inflow not allowed to be > full flow
////////////////////////////////////////////////////////////
//  Check for Max. flow instead of full flow. (LR - 3/10/06)
////////////////////////////////////////////////////////////
//  if ( qin > 0.999 )
//  {
//      qin = 1.0;
//      ain = 1.0;
//  }
    if ( qin > Conduit[k].qMax / Qfull )
    {
        qin = Conduit[k].qMax / Qfull;
        ain = xsect_getAmax(pXsect);
    }

    // --- get normalized inlet area corresponding to inlet flow
    else ain = xsect_getAofS(pXsect, qin/Beta1) / Afull;

    // --- check for no flow
    if ( qin <= TINY && q2 <= TINY )
    {
        qout = 0.0;
        aout = 0.0;
    }


    // --- check for full flow
///////////////////////////////////////////////
//  This check has been removed. (LR - 3/10/06)
///////////////////////////////////////////////
//  else if ( qin >= 0.999 && q2 >= 0.999 )
//  {
//      qout = 1.0;
//      aout = 1.0;
//  }

    // --- otherwise solve finite difference form of continuity eqn.
    else
    {
        // --- compute constant factors
        dxdt = Conduit[k].length / tStep * Afull / Qfull;
        dq   = q2 - q1;
        C1   = dxdt * WT / WX;
        C2   = (1.0 - WT) * (ain - a1);
        C2   = C2 - WT * a2;
        C2   = C2 * dxdt / WX;
        C2   = C2 + (1.0 - WX) / WX * dq - qin;
    
        // --- starting guess for aout is value from previous time step
        aout = a2;

        // --- solve continuity equation for aout
        result = solveContinuity(qin, ain, &aout);

        // --- report error if continuity eqn. not solved
        if ( result == -1 )
        {
            report_writeErrorMsg(ERR_KINWAVE, Link[j].ID);
            return 1;
        }
        if ( result <= 0 ) result = 1;

        // --- compute normalized outlet flow from outlet area
        qout = Beta1 * xsect_getSofA(pXsect, aout*Afull);
    }

    // --- save new flows and areas
    Conduit[k].q1 = qin * Qfull;
    Conduit[k].a1 = ain * Afull;
    Conduit[k].q2 = qout * Qfull;
    Conduit[k].a2 = aout * Afull;
    (*qinflow)  = Conduit[k].q1 * Conduit[k].barrels;
    (*qoutflow) = Conduit[k].q2 * Conduit[k].barrels;
    return result;
}

//=============================================================================

int solveContinuity(float qin, float ain, float* aout)
//
//  Input:   qin = upstream normalized flow
//           ain = upstream normalized area
//           aout = downstream normalized area
//  Output:  new value for aout; returns an error code
//  Purpose: solves continuity equation  f(a) = Beta1*S(a) + C1*a + C2 = 0
//           for 'a' using the Newton-Raphson root finder function.
//           Return code has the following meanings:
//           >= 0 number of function evaluations used
//           -1   Newton function failed
//           -2   flow always above max. flow
//           -3   flow always below zero
//
//     Note: pXsect (pointer to conduit's cross-section), and constants Beta1,
//           C1, and C2 are module-level shared variables assigned values in
//           kinwave_execute().
//
{
    int   n;                           // # evaluations or error code
    float aLo, aHi, aTmp;              // lower/upper bounds on a
    float fLo, fHi;                    // lower/upper bounds on f
    float tol = EPSIL;                 // absolute convergence tol.

    // --- first determine bounds on 'a' so that f(a) passes through 0.

    // --- set upper bound to area at full flow
    aHi = 1.0;
    fHi = 1.0 + C1 + C2;

    // --- try setting lower bound to area where section factor is maximum
    aLo = xsect_getAmax(pXsect) / Afull;
    if ( aLo < aHi )
    {
        fLo = ( Beta1 * pXsect->sMax ) + (C1 * aLo) + C2;
    }
    else fLo = fHi;

    // --- if fLo and fHi have same sign then set lower bound to 0
    if ( fHi*fLo > 0.0 )
    {
        aHi = aLo;
        fHi = fLo;
        aLo = 0.0;
        fLo = C2;
    }

    // --- proceed with search for root if fLo and fHi have different signs
    if ( fHi*fLo <= 0.0 )
    {
        // --- start search at midpoint of lower/upper bounds
        //     if initial value outside of these bounds
		if ( *aout < aLo || *aout > aHi ) *aout = 0.5*(aLo + aHi);

        // --- if fLo > fHi then switch aLo and aHi
        if ( fLo > fHi )
        {
            aTmp = aLo;
            aLo  = aHi;
            aHi  = aTmp;
        }

        // --- call the Newton root finder method passing it the 
        //     evalContinuity function to evaluate the function
        //     and its derivatives
        n = findroot_Newton(aLo, aHi, aout, tol, evalContinuity);

        // --- check if root finder succeeded
        if ( n <= 0 ) n = -1;
    }

    // --- if lower/upper bound functions both negative then use full flow
    else if ( fLo < 0.0 )
    {
        if ( qin > 1.0 ) *aout = ain;
        else *aout = 1.0;
        n = -2;
    }

    // --- if lower/upper bound functions both positive then use no flow
    else if ( fLo > 0 )
    {
        *aout = 0.0;
        n = -3;
    }
    else n = -1;
    return n;
}

//=============================================================================

void evalContinuity(float a, float* f, float* df)
//
//  Input:   a = outlet normalized area
//  Output:  f = value of continuity eqn.
//           df = derivative of continuity eqn.
//  Purpose: computes value of continuity equation (f) and its derivative (df)
//           w.r.t. normalized area for link with normalized outlet area 'a'.
//
{
    *f  = (Beta1 * xsect_getSofA(pXsect, a*Afull)) + (C1 * a) + C2;
    *df = (Beta1 * Afull * xsect_getdSdA(pXsect, a*Afull)) + C1;
}

//=============================================================================
