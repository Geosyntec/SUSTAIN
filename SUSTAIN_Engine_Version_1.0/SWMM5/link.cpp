//-----------------------------------------------------------------------------
//   link.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             6/14/05  (Build 5.0.005b)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//             7/5/06   (Build 5.0.008)
//             9/19/06  (Build 5.0.009)
//   Author:   L. Rossman
//
//   Conveyance system link functions
//-----------------------------------------------------------------------------

#include <string.h>
#include <stdlib.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  Constants
//-----------------------------------------------------------------------------
// moved to funcs.h
//static const float MIN_DELTA_Z = 0.001; // minimum elevation change for conduit
                                        // slopes (ft)

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  link_readParams      (called by parseLine in input.c)
//  link_readXsectParams (called by parseLine in input.c)
//  link_readLossParams  (called by parseLine in input.c)
//  link_validate        (called by swmm_open in swmm5.c)
//  link_initState       (called by initObjects in swmm5.c)
//  link_setOldHydState  (called by routing_execute in routing.c)
//  link_setOldQualState (called by routing_execute in routing.c)
//  link_getResults      (called by output_saveLinkResults)
//  link_getFroude
//  link_getInflow
//  link_setOutfallDepth
//  link_getYcrit
//  link_getYnorm
//  link_getVelocity

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void  link_setParams(int j, int type, int n1, int n2, int k, float x[]);

static int   conduit_readParams(int j, int k, char* tok[], int ntoks);
static void  conduit_validate(int j, int k);
static void  conduit_initState(int j, int k);
static void  conduit_reverse(int j, int k);
static float conduit_getLengthFactor(int j, int k);
static float conduit_getInflow(int j);
static void  conduit_updateStats(int j, float dt, DateTime aDate);

static int   pump_readParams(int j, int k, char* tok[], int ntoks);
static void  pump_validate(int j, int k);
static void  pump_initState(int j, int k);
static float pump_getInflow(int j);

static int   orifice_readParams(int j, int k, char* tok[], int ntoks);
static void  orifice_validate(int j, int k);
static float orifice_getInflow(int j);
static float orifice_getFlow(int j, int k, float head, float f);

static int   weir_readParams(int j, int k, char* tok[], int ntoks);
static void  weir_validate(int j, int k);
static float weir_getInflow(int j);
static float weir_getOpenArea(int j, float y);
static void  weir_getFlow(int j, int k, float head, float dir, int hasFlapGate,
                          float* q1, float* q2);
static float weir_getdqdh(int k, float dir, float h, float q1, float q2);

static int   outlet_readParams(int j, int k, char* tok[], int ntoks);
static float outlet_getFlow(int k, float head);
static float outlet_getInflow(int j);


//=============================================================================

int link_readParams(int j, int type, int k, char* tok[], int ntoks)
//
//  Input:   j     = link index
//           type  = link type code
//           k     = link type index
//           tok[] = array of string tokens
//           ntoks = number of tokens   
//  Output:  returns an error code
//  Purpose: reads parameters for a specific type of link from a 
//           tokenized line of input data.
//
{
    switch ( type )
    {
      case CONDUIT: return conduit_readParams(j, k, tok, ntoks);
      case PUMP:    return pump_readParams(j, k, tok, ntoks);
      case ORIFICE: return orifice_readParams(j, k, tok, ntoks);
      case WEIR:    return weir_readParams(j, k, tok, ntoks);
      case OUTLET:  return outlet_readParams(j, k, tok, ntoks);
      default: return 0;
    }
}

//=============================================================================

int link_readXsectParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens   
//  Output:  returns an error code
//  Purpose: reads a link's cross section parameters from a tokenized
//           line of input data.
//
{
    int   i, j, k;
    float x[4];

    // --- get index of link
    if ( ntoks < 6 ) return error_setInpError(ERR_ITEMS, "");
    j = project_findObject(LINK, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);

    // --- get code of xsection shape
    k = findmatch(tok[1], XsectTypeWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[1]);

    // --- assign default number of barrels to conduit
    if ( Link[j].type == CONDUIT ) Conduit[Link[j].subIndex].barrels = 1;

    // --- for irregular shape, find index of transect object
    if ( k == IRREGULAR )
    {
        i = project_findObject(TRANSECT, tok[2]);
        if ( i < 0.0 ) return error_setInpError(ERR_NAME, tok[2]);
        Link[j].xsect.type = k;
        Link[j].xsect.transect = i;
    }
    else
    {
        // --- parse and save geometric parameters
        for (i = 2; i <= 5; i++)
        {
            if ( !getFloat(tok[i], &x[i-2]) )
                return error_setInpError(ERR_NUMBER, tok[i]);
        }
        if ( !xsect_setParams(&Link[j].xsect, k, x, UCF(LENGTH)) )
        {
            return error_setInpError(ERR_NUMBER, "");
        }

        // --- parse number of barrels if present
        if ( Link[j].type == CONDUIT && ntoks >= 7 )
        {
            i = atof(tok[6]) + 0.01;
            if ( i <= 0 ) return error_setInpError(ERR_NUMBER, tok[6]);
            else Conduit[Link[j].subIndex].barrels = (char)i;
        }
    }
    return 0;
}

//=============================================================================

int link_readLossParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens   
//  Output:  returns an error code
//  Purpose: reads local loss parameters for a link from a tokenized
//           line of input data.
//
{
    int   i, j, k;
    float x[3];

    if ( ntoks < 4 ) return error_setInpError(ERR_ITEMS, "");
    j = project_findObject(LINK, tok[0]);
    if ( j < 0 ) return error_setInpError(ERR_NAME, tok[0]);
    for (i=1; i<=3; i++)
    {
        if ( ! getFloat(tok[i], &x[i-1]) )
        return error_setInpError(ERR_NUMBER, tok[i]);
    }
    k = 0;
    if ( ntoks >= 5 )
    {
        k = findmatch(tok[4], NoYesWords);             
        if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[4]);
    }
    Link[j].cLossInlet   = x[0];
    Link[j].cLossOutlet  = x[1];
    Link[j].cLossAvg     = x[2];
    Link[j].hasFlapGate  = k;
    return 0;
}

//=============================================================================

void  link_setParams(int j, int type, int n1, int n2, int k, float x[])
//
//  Input:   j   = link index
//           type = link type code
//           n1   = index of upstream node
//           n2   = index of downstream node
//           k    = index of link's sub-type
//           x    = array of parameter values
//  Output:  none
//  Purpose: sets parameters for a link.
//
{
    Link[j].node1       = n1;
    Link[j].node2       = n2;
    Link[j].type        = type;
    Link[j].subIndex    = k;
    Link[j].z1          = 0.0;
    Link[j].z2          = 0.0;
    Link[j].q0          = 0.0;
    Link[j].qFull       = 0.0;
    Link[j].setting     = 1.0;
    Link[j].hasFlapGate = 0;
    Link[j].qLimit      = 0.0;         // 0 means that no limit is defined
    Link[j].direction   = 1;

    switch (type)
    {
      case CONDUIT:
        Conduit[k].length    = x[0] / UCF(LENGTH);
        Conduit[k].modLength = Conduit[k].length;
        Conduit[k].roughness = x[1];
        Link[j].z1           = x[2] / UCF(LENGTH);
        Link[j].z2           = x[3] / UCF(LENGTH);
        Link[j].q0           = x[4] / UCF(FLOW);
        Link[j].qLimit      = x[5] / UCF(FLOW);
        break;

      case PUMP:
        Pump[k].pumpCurve    = x[0];
        Link[j].hasFlapGate  = FALSE;
        Link[j].setting      = x[1];
        break;

      case ORIFICE:
        Orifice[k].type      = x[0];
        Link[j].z1           = x[1] / UCF(LENGTH);
        Link[j].z2           = Link[j].z1;
        Orifice[k].cDisch    = x[2];
        Link[j].hasFlapGate  = (x[3] > 0.0) ? 1 : 0; 
        break;

      case WEIR:
        Weir[k].type         = x[0];
        Link[j].z1           = x[1] / UCF(LENGTH);
        Link[j].z2           = Link[j].z1;
        Weir[k].crestHt      = Link[j].z1;
        Weir[k].cDisch1      = x[2];
        Link[j].hasFlapGate  = (x[3] > 0.0) ? 1 : 0;
        Weir[k].endCon       = x[4];
        Weir[k].cDisch2      = x[5];
        break;

      case OUTLET:
        Link[j].z1           = x[0] / UCF(LENGTH);
        Link[j].z2           = Link[j].z1;
        Outlet[k].crestHt    = Link[j].z1;
        Outlet[k].qCoeff     = x[1];
        Outlet[k].qExpon     = x[2];
        Outlet[k].qCurve     = x[3];
        Link[j].hasFlapGate  = (x[4] > 0.0) ? 1 : 0;

        xsect_setParams(&Link[j].xsect, DUMMY, NULL, 0.0);
        break;

    }
}

//=============================================================================

void  link_validate(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: validates a link's properties.
//
{
    int   n;
    switch ( Link[j].type )
    {
      case CONDUIT: conduit_validate(j, Link[j].subIndex); break;
      case PUMP:    pump_validate(j, Link[j].subIndex);    break;
      case ORIFICE: orifice_validate(j, Link[j].subIndex); break;
      case WEIR:    weir_validate(j, Link[j].subIndex);    break;
    }

    // --- force max. depth of end nodes to be >= link crown height
    //     at non-storage nodes
    n = Link[j].node1;
    if ( Node[n].type != STORAGE )
    {
        Node[n].fullDepth = MAX(Node[n].fullDepth,
                            Link[j].z1 + Link[j].xsect.yFull);
    }
    n = Link[j].node2;
    if ( Node[n].type != STORAGE )
    {
        Node[n].fullDepth = MAX(Node[n].fullDepth,
                            Link[j].z2 + Link[j].xsect.yFull);
    }
}

//=============================================================================

void link_initState(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: initializes a link's state variables at start of simulation.
//
{
    int   p;

    // --- initialize hydraulic state
    Link[j].oldFlow   = Link[j].q0;
    Link[j].newFlow   = Link[j].q0;
    Link[j].oldDepth  = 0.0;
    Link[j].newDepth  = 0.0;
    Link[j].oldVolume = 0.0;
    Link[j].newVolume = 0.0;
    Link[j].isClosed  = FALSE;
    if ( Link[j].type == CONDUIT ) conduit_initState(j, Link[j].subIndex);
    if ( Link[j].type == PUMP    ) pump_initState(j, Link[j].subIndex);
    
    // --- initialize water quality state
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        Link[j].oldQual[p] = 0.0;
        Link[j].newQual[p] = 0.0;
    }
}

//=============================================================================

float  link_getInflow(int j)
//
//  Input:   j = link index
//  Output:  returns link flow rate (cfs)
//  Purpose: finds total flow entering a link during current time step.
//
{
    if ( Link[j].setting == 0.0 ||
         Link[j].isClosed ) return 0.0;
    switch ( Link[j].type )
    {
      case CONDUIT: return conduit_getInflow(j);
      case PUMP:    return pump_getInflow(j);
      case ORIFICE: return orifice_getInflow(j);
      case WEIR:    return weir_getInflow(j);
      case OUTLET:  return outlet_getInflow(j);
      default:      return node_getOutflow(Link[j].node1, j);
    }
}

//=============================================================================

void link_setOldHydState(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: replaces link's old hydraulic state values with current ones.
//
{
    int k;
    Link[j].oldDepth  = Link[j].newDepth;
    Link[j].oldFlow   = Link[j].newFlow;
    Link[j].oldVolume = Link[j].newVolume;
    if ( Link[j].type == CONDUIT )
    {
        k = Link[j].subIndex;
        Conduit[k].q1Old = Conduit[k].q1;
        Conduit[k].q2Old = Conduit[k].q2;
    }
}

//=============================================================================

void link_setOldQualState(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: replaces link's old water quality state values with current ones.
//
{
    int p;
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        Link[j].oldQual[p] = Link[j].newQual[p];
        Link[j].newQual[p] = 0.0;
    }
}

//=============================================================================

void link_setFlapGate(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: sets isClosed property of links with flap gates or connected to
//           outfalls with tide gates, depending on current flow direction.
//
{
    int   n1, n2;
    float h1, h2;

    Link[j].isClosed = FALSE;
    if ( Link[j].type == CONDUIT || Link[j].type == PUMP ) return;
    n1 = Link[j].node1;
    n2 = Link[j].node2;
    h1 = Node[n1].invertElev + Node[n1].newDepth;
    h2 = Node[n2].invertElev + Node[n2].newDepth;
    if ( Link[j].hasFlapGate &&
         Link[j].direction * (h2 - h1) > 0.0 ) Link[j].isClosed = TRUE;
}

//=============================================================================

void link_getResults(int j, float f, float x[])
//
//  Input:   j = link index
//           f = time weighting factor
//  Output:  x = array of weighted results
//  Purpose: retrieves time-weighted average of old and new results for a link.
//
{
    int   p;                      // pollutant index
    float y,                      // depth
          q,                      // flow
          v,                      // velocity
          fr,                     // Froude no.
          c;                      // capacity
    float f1 = 1.0 - f;

    y = f1*Link[j].oldDepth + f*Link[j].newDepth;
    q = f1*Link[j].oldFlow + f*Link[j].newFlow;
    v = link_getVelocity(j, q, y);
    fr = link_getFroude(j, v, y);
    c = 0.0;
    if ( Link[j].type != PUMP && Link[j].xsect.type != DUMMY )
        c = y / Link[j].xsect.yFull;

    x[LINK_DEPTH]    = y * UCF(LENGTH);
    x[LINK_FLOW]     = q * UCF(FLOW) * (float)Link[j].direction;
    x[LINK_VELOCITY] = v * UCF(LENGTH) * (float)Link[j].direction;
    x[LINK_FROUDE]   = fr;
    x[LINK_CAPACITY] = c;
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        x[LINK_QUAL+p] = f1*Link[j].oldQual[p] + f*Link[j].newQual[p];
    }
}

//=============================================================================

void link_setOutfallDepth(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: sets depth at outfall node connected to link j.
//
{
    int    k;                          // conduit index
    int    n;                          // outfall node index
    float  z;                          // invert offset height (ft)
    float  q;                          // flow rate (cfs)
    float  yCrit = 0.0;                // critical flow depth (ft)
    float  yNorm = 0.0;                // normal flow depth (ft)

    // --- find which end node of link is an outfall
    if ( Node[Link[j].node2].type == OUTFALL )
    {
        n = Link[j].node2;
        z = Link[j].z2;
    }
    else if ( Node[Link[j].node1].type == OUTFALL )
    {
        n = Link[j].node1;
        z = Link[j].z1;
    }
    else return;
    
    // --- find both normal & critical depth for current flow
    if ( Link[j].type == CONDUIT )
    {
        k = Link[j].subIndex;
        q = fabs(Link[j].newFlow / Conduit[k].barrels);
        yNorm = link_getYnorm(j, q);
        yCrit = link_getYcrit(j, q);
    }

    // --- set new depth at node
    node_setOutletDepth(n, yNorm, yCrit, z);
}

//=============================================================================

float link_getYcrit(int j, float q)
//
//  Input:   j = link index
//           q = link flow rate (cfs)
//  Output:  returns critical depth (ft)
//  Purpose: computes critical depth for given flow rate.
//
{
    return (float)xsect_getYcrit(&Link[j].xsect, q);
}

//=============================================================================

float  link_getYnorm(int j, float q)
//
//  Input:   j = link index
//           q = link flow rate (cfs)
//  Output:  returns normal depth (ft)
//  Purpose: computes normal depth for given flow rate.
//
{
    int   k;
    double s, a, y;

    if ( Link[j].type != CONDUIT ) return 0.0;
    if ( Link[j].xsect.type == DUMMY ) return 0.0;
    q = fabs(q);
    if ( q <= 0.0 ) return 0.0;

//////////////////////////////////////////////////////////////////////////////
//  Comparison should be made against max. flow, not full flow. (LR - 3/10/06)
//////////////////////////////////////////////////////////////////////////////
//  if ( q >= Link[j].qFull ) return Link[j].xsect.yFull;
    k = Link[j].subIndex;
    if ( q > Conduit[k].qMax ) return Link[j].xsect.yFull;

    s = q / Conduit[k].beta;
    a = xsect_getAofS(&Link[j].xsect, s);
    y = xsect_getYofA(&Link[j].xsect, a);
    return (float)y;
}

//=============================================================================

float link_getVelocity(int j, float flow, float depth)
//
//  Input:   j     = link index
//           flow  = link flow rate (cfs)
//           depth = link flow depth (ft)
//  Output:  returns flow velocity (fps)
//  Purpose: finds flow velocity given flow and depth.
//
{
    float area;
    float veloc = 0.0;
    int   k;

    if ( depth <= 0.01 ) return 0.0;
    if ( Link[j].type == CONDUIT )
    {
        k = Link[j].subIndex;
        flow /= Conduit[k].barrels;
        area = xsect_getAofY(&Link[j].xsect, depth);
        if (area > FUDGE ) veloc = flow / area;
    }
    return veloc;
}

//=============================================================================

float link_getFroude(int j, float v, float y)
//
//  Input:   j = link index
//           v = flow velocity (fps)
//           y = flow depth (ft)
//  Output:  returns Froude Number
//  Purpose: computes Froude Number for given velocity and flow depth
//
{
    float   yMin;
    float   yMax;
    TXsect* xsect = &Link[j].xsect;

    // --- return 0 if link is not a conduit
    if ( Link[j].type != CONDUIT ) return 0.0;
    if ( y <= FUDGE ) return 0.0;

    // --- find effective flow depth y
    yMin = 0.04 * xsect->yFull;
    yMax = xsect->yFull;
    if      ( y < yMin ) y = yMin;          // don't let y be < 4% yFull
    else if ( y >= yMax ) y = yMax;         // don't let y be > yFull
    else if ( xsect_isOpen(xsect->type) )   // use hyd. depth for open channel
    {
        y = xsect_getAofY(xsect, y) / xsect_getWofY(xsect, y);
    }

    // --- compute Froude No. from effective depth & velocity
    return fabs(v) / sqrt(GRAVITY * y);
}


//=============================================================================
//                    C O N D U I T   M E T H O D S
//=============================================================================

int  conduit_readParams(int j, int k, char* tok[], int ntoks)
//
//  Input:   j = link index
//           k = conduit index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads conduit parameters from a tokenzed line of input.
//
{
    int   i;
    int   n1, n2;
    float x[6];
    char* id;

    // --- check for valid ID and end node IDs
    if ( ntoks < 7 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LINK, tok[0]);                // link ID
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    n1 = project_findObject(NODE, tok[1]);            // upstrm. node
    if ( n1 < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n2 = project_findObject(NODE, tok[2]);            // dwnstrm. node
    if ( n2 < 0 ) return error_setInpError(ERR_NAME, tok[2]);

    // --- parse required conduit parameters
    for (i=3; i<=6; i++)
    {
        if ( !getFloat(tok[i], &x[i-3]) )
            return error_setInpError(ERR_NUMBER, tok[i]);
    }

    // --- parse optional parameters
    x[4] = 0.0;                                       // init. flow
    if ( ntoks >= 8 )
    {
        if ( !getFloat(tok[7], &x[4]) )
        return error_setInpError(ERR_NUMBER, tok[7]);
    }
    x[5] = 0.0;
    if ( ntoks >= 9 )
    {
        if ( !getFloat(tok[8], &x[5]) )
        return error_setInpError(ERR_NUMBER, tok[8]);
    }

    // --- add parameters to data base
    Link[j].ID = id;
    link_setParams(j, CONDUIT, n1, n2, k, x);
    return 0;
}

//=============================================================================

void  conduit_validate(int j, int k)
//
//  Input:   j = link index
//           k = conduit index
//  Output:  none
//  Purpose: validates a conduit's properties.
//
{
    float aa;
    float elev1, elev2;
    float lengthFactor, roughness;

    // --- if irreg. xsection, assign transect roughness to conduit
    if ( Link[j].xsect.type == IRREGULAR )
    {
        xsect_setIrregXsectParams(&Link[j].xsect);
        Conduit[k].roughness = Transect[Link[j].xsect.transect].roughness;
    }

    // --- check for valid length & roughness
    if ( Conduit[k].length <= 0.0 )
        report_writeErrorMsg(ERR_LENGTH, Link[j].ID);
    if ( Conduit[k].roughness <= 0.0 )
        report_writeErrorMsg(ERR_ROUGHNESS, Link[j].ID);
    if ( Conduit[k].barrels <= 0 )
        report_writeErrorMsg(ERR_BARRELS, Link[j].ID);

    // --- check for valid xsection
    if ( Link[j].xsect.type != DUMMY )
    {
        if ( Link[j].xsect.type < 0 )
            report_writeErrorMsg(ERR_NO_XSECT, Link[j].ID);
        else if ( Link[j].xsect.aFull <= 0.0 )
            report_writeErrorMsg(ERR_XSECT, Link[j].ID);
    }

//////////////////////////////////////////////
//  Added per suggestion by TS. (LR - 7/5/06 )
//////////////////////////////////////////////
    // --- check for non-negative offsets
    if ( Link[j].z1 < 0.0 || Link[j].z2 < 0.0)
        report_writeErrorMsg(ERR_OFFSET, Link[j].ID);

    if ( ErrorCode ) return;

//////////////////////
//  LR - added 6/14/05
//////////////////////
    // --- adjust conduit offsets for partly filled circular xsection
    if ( Link[j].xsect.type == FILLED_CIRCULAR )
    {
        Link[j].z1 += Link[j].xsect.yBot;
        Link[j].z2 += Link[j].xsect.yBot;
    }

    // --- compute conduit slope 
    elev1 = Link[j].z1 + Node[Link[j].node1].invertElev;
    elev2 = Link[j].z2 + Node[Link[j].node2].invertElev;
    if ( fabs(elev1 - elev2) < MIN_DELTA_Z )
    {
        Conduit[k].slope = MIN_DELTA_Z / Conduit[k].length;
    }
    else Conduit[k].slope = (elev1 - elev2) / Conduit[k].length;

    // --- reverse orientation of conduit if using dynamic wave routing 
    //     and slope is negative
      if ( RouteModel == DW &&
           Conduit[k].slope < 0.0 &&
           Link[j].xsect.type != DUMMY )
      {
          conduit_reverse(j, k);
      }

    // --- lengthen conduit if lengthening option is in effect
    if ( RouteModel == DW &&
         LengtheningStep > 0.0 &&
         Link[j].xsect.type != DUMMY )
    {
        lengthFactor = conduit_getLengthFactor(j,k);     
        Conduit[k].modLength = lengthFactor * Conduit[k].length;
    }
    else lengthFactor = 1.0;

    // --- compute modified slope, roughness & roughness factor
    Conduit[k].slope /= lengthFactor;
    roughness = Conduit[k].roughness / sqrt(lengthFactor);
    Conduit[k].roughFactor = GRAVITY * SQR(roughness/PHI);

    // --- compute full flow through cross section
    if ( Link[j].xsect.type == DUMMY ) Conduit[k].beta = 0.0;
    else Conduit[k].beta = PHI * sqrt(fabs(Conduit[k].slope)) / roughness;
    Link[j].qFull = Link[j].xsect.sFull * Conduit[k].beta;
    Conduit[k].qMax = Link[j].xsect.sMax * Conduit[k].beta;

    // --- see if flow is supercritical most of time
    //     by comparing normal & critical velocities.
    //     (factor of 0.3 is for circular pipe 95% full)
    // NOTE: this factor is used for modified Kinematic Wave routing.
    aa = Conduit[k].beta / sqrt(32.2) *
         pow(Link[j].xsect.yFull, 0.1666667) * 0.3;
    if ( aa >= 1.0 ) Conduit[k].superCritical = TRUE;
    else             Conduit[k].superCritical = FALSE;

    // --- set value of hasLosses flag
    if ( Link[j].cLossInlet  == 0.0 &&
         Link[j].cLossOutlet == 0.0 &&
         Link[j].cLossAvg    == 0.0
       ) Conduit[k].hasLosses = FALSE;
    else Conduit[k].hasLosses = TRUE;
}

//=============================================================================

void conduit_reverse(int j, int k)
//
//  Input:   j = link index
//           k = conduit index
//  Output:  none
//  Purpose: reverses direction of a conduit
//
{
    int i;
    float z;
    float cLoss;
    i = Link[j].node1;
    Link[j].node1 = Link[j].node2;
    Link[j].node2 = i;
    z = Link[j].z1;
    Link[j].z1 = Link[j].z2;
    Link[j].z2 = z;
    cLoss = Link[j].cLossInlet;
    Link[j].cLossInlet = Link[j].cLossOutlet;
    Link[j].cLossOutlet = cLoss;
    Conduit[k].slope = -Conduit[k].slope;
    Link[j].direction *= (signed char)-1;
    Link[j].q0 = -Link[j].q0;
}


//=============================================================================

float conduit_getLengthFactor(int j, int k)
//
//  Input:   j = link index
//           k = conduit index
//  Output:  returns factor by which a conduit should be lengthened
//  Purpose: computes amount of conduit lengthing to improve numerical stability.
//
//  The following form of the Courant criterion is used:
//      L = t * v * (1 + Fr) / Fr
//  where L = conduit length, t = time step, v = velocity, & Fr = Froude No.
//  After substituting Fr = v / sqrt(gy), where y = flow depth, we get:
//    L = t * ( sqrt(gy) + v )
//
{
    float ratio;
    float yFull;
    float vFull;
    float tStep;

    // --- evaluate flow depth and velocity at full normal flow condition
    yFull = Link[j].xsect.yFull;
    if ( xsect_isOpen(Link[j].xsect.type) )
    {
        yFull = Link[j].xsect.aFull / xsect_getWofY(&Link[j].xsect, yFull);
    }
    vFull = Link[j].xsect.sFull * PHI * sqrt(fabs(Conduit[k].slope))
            / Conduit[k].roughness / Link[j].xsect.aFull;

    // --- determine ratio of Courant length to actual length
    if ( LengtheningStep == 0.0 ) tStep = RouteStep;
    else                          tStep = MIN(RouteStep, LengtheningStep);
    ratio = (sqrt(GRAVITY*yFull) + vFull) * tStep / Conduit[k].length;

    // --- return max. of 1.0 and ratio
    if ( ratio > 1.0 ) return ratio;
    else return 1.0;
}

//=============================================================================

void  conduit_initState(int j, int k)
//
//  Input:   j = link index
//           k = conduit index
//  Output:  none
//  Purpose: sets initial conduit depth to normal depth of initial flow
//
{
    Link[j].newDepth = link_getYnorm(j, Link[j].q0 / Conduit[k].barrels);
    Link[j].oldDepth = Link[j].newDepth;
}

//=============================================================================

float conduit_getInflow(int j)
//
//  Input:   j = link index
//  Output:  returns flow in link (cfs)
//  Purpose: finds inflow to conduit from upstream node.
//
{
    float qIn = node_getOutflow(Link[j].node1, j);
    return qIn;
}


//=============================================================================
//                        P U M P   M E T H O D S
//=============================================================================

int  pump_readParams(int j, int k, char* tok[], int ntoks)
//
//  Input:   j = link index
//           k = pump index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads pump parameters from a tokenized line of input.
//
{
    int   m;
    int   n1, n2;
    float x[2];
    char* id;

    // --- check for valid ID and end node IDs
    if ( ntoks < 5 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LINK, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    n1 = project_findObject(NODE, tok[1]);
    if ( n1 < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n2 = project_findObject(NODE, tok[2]);
    if ( n2 < 0 ) return error_setInpError(ERR_NAME, tok[2]);

    // --- parse curve name
    m = project_findObject(CURVE, tok[3]);
    if ( m < 0 ) return error_setInpError(ERR_NAME, tok[3]);
    x[0] = m;

    // --- parse init. status if present
    x[1] = 1.0;
    if ( ntoks >= 5 )
    {
        m = findmatch(tok[4], OffOnWords);
        if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[4]);
        x[1] = m;
    }

    // --- add parameters to data base
    Link[j].ID = id;
    link_setParams(j, PUMP, n1, n2, k, x);
    return 0;
}

//=============================================================================

void  pump_validate(int j, int k)
//
//  Input:   j = link index
//           k = pump index
//  Output:  none
//  Purpose: validates a pump's properties
//
{
    int    m;
    double x, y;

    Link[j].xsect.yFull = 0.0;

    // --- check for valid curve type
    m = Pump[k].pumpCurve;
    if ( m < 0 ) report_writeErrorMsg(ERR_NO_CURVE, Link[j].ID);
    else if ( Curve[m].curveType < PUMP1_CURVE ||
              Curve[m].curveType > PUMP4_CURVE )
              report_writeErrorMsg(ERR_NO_CURVE, Link[j].ID);

    // --- store pump curve type with pump's parameters
    else 
    {
        Pump[k].type = Curve[m].curveType - PUMP1_CURVE;
        //table_getLastEntry(&Curve[m], &x, &y);           ////Removed (LR - 3/10/06)

/////////////////////////////////////////////////////////////////
//  New code added to determine highest pump flow. (LR - 3/10/06)
/////////////////////////////////////////////////////////////////
        if ( table_getFirstEntry(&Curve[m], &x, &y) )
        {
            Link[j].qFull = y;       
            while ( table_getNextEntry(&Curve[m], &x, &y) )
            {
                Link[j].qFull = MAX(y, Link[j].qFull);
            }
        }
        Link[j].qFull /= UCF(FLOW);
   }
}

//=============================================================================

void  pump_initState(int j, int k)
//
//  Input:   j = link index
//           k = pump index
//  Output:  none
//  Purpose: assigns wet well volume to inlet node of Type 1 pump.
//
{
    int    m;                          // pump curve index
    int    n1;                         // upstream node index
    float  vft3;                       // volume (ft3)
    double v;                          // volume (ft3 or m3)
    double q;                          // pump flow (not used)
    if ( Pump[k].type == TYPE1_PUMP )
    {
        n1 = Link[j].node1;
        if ( Node[n1].type != STORAGE )
        {
            m = Pump[k].pumpCurve;
            table_getLastEntry(&Curve[m], &v, &q);
            vft3 = v / UCF(VOLUME);
            Node[n1].fullVolume = MAX(Node[n1].fullVolume, vft3);
        }
    }
}

//=============================================================================

float pump_getInflow(int j)
//
//  Input:   j = link index
//  Output:  returns pump flow (cfs)
//  Purpose: finds flow produced by a pump.
//
{
    int     k, m;
    int     n1, n2;
    float   vol, depth, head;
    float   qIn;

///////////////////////////////////////
//  New variables added. (LR - 9/19/06)
///////////////////////////////////////
    float   qIn1, dh = 0.001;

    k = Link[j].subIndex;
    m = Pump[k].pumpCurve;
    n1 = Link[j].node1;
    n2 = Link[j].node2;

    // --- no flow if no pump curve or setting is closed
    if ( m < 0 || Link[j].setting == 0.0 ) return 0.0;

    // --- pumping rate depends on pump curve type
    switch(Curve[m].curveType)
    {
      case PUMP1_CURVE:
        vol = Node[n1].newVolume * UCF(VOLUME);
        qIn = table_intervalLookup(&Curve[m], vol) / UCF(FLOW);
        break;

      case PUMP2_CURVE:
        depth = Node[n1].newDepth * UCF(LENGTH);
        qIn = table_intervalLookup(&Curve[m], depth) / UCF(FLOW);
        break;

      case PUMP3_CURVE:
        head = ( (Node[n2].newDepth + Node[n2].invertElev) -
                 (Node[n1].newDepth + Node[n1].invertElev) ) * UCF(LENGTH);
        qIn = table_lookup(&Curve[m], head) / UCF(FLOW);

//////////////////////////////////////////////////
//  New code added to compute dqdh. (LR - 9/19/06)
//////////////////////////////////////////////////
        // --- compute dQ/dh (slope of pump curve) and
        //     reverse sign since flow decreases with increasing head
        qIn1 = table_lookup(&Curve[m], (head+dh)*UCF(LENGTH)) / UCF(FLOW);
        Link[j].dqdh = -(qIn1 - qIn) / dh;

        break;

      case PUMP4_CURVE:
        depth = Node[n1].newDepth * UCF(LENGTH);
        qIn = table_lookup(&Curve[m], depth) / UCF(FLOW);

//////////////////////////////////////////////////
//  New code added to compute dqdh. (LR - 9/19/06)
//////////////////////////////////////////////////
        // --- compute dQ/dh (slope of pump curve)
        qIn1 = table_lookup(&Curve[m], (depth+dh)*UCF(LENGTH)) / UCF(FLOW);
        Link[j].dqdh = (qIn1 - qIn) / dh;
        break;

      default: qIn = 0.0;
    }

    // --- do not allow reverse flow through pump
    if ( qIn < 0.0 )  qIn = 0.0;

//////////////////////////////////////////////////////
//  Adjust qIn by any controller setting. (LR - 9/5/05
//////////////////////////////////////////////////////
    return qIn * Link[j].setting; 
}


//=============================================================================
//                    O R I F I C E   M E T H O D S
//=============================================================================

int  orifice_readParams(int j, int k, char* tok[], int ntoks)
//
//  Input:   j = link index
//           k = orifice index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads orifice parameters from a tokenized line of input.
//
{
    int   m;
    int   n1, n2;
    float x[4];
    char* id;

    // --- check for valid ID and end node IDs
    if ( ntoks < 6 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LINK, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    n1 = project_findObject(NODE, tok[1]);
    if ( n1 < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n2 = project_findObject(NODE, tok[2]);
    if ( n2 < 0 ) return error_setInpError(ERR_NAME, tok[2]);

    // --- parse orifice parameters
    m = findmatch(tok[3], OrificeTypeWords);
    if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[3]);
    x[0] = m;                                            // type

//////////////////////////////////////////////////////////////////
//  Modified to check for negative height & cDisch. (LR - 7/5/06 )
//////////////////////////////////////////////////////////////////
    if ( ! getFloat(tok[4], &x[1]) || x[1] < 0.0 )       // height
        return error_setInpError(ERR_NUMBER, tok[4]);
    if ( ! getFloat(tok[5], &x[2]) || x[2] < 0.0 )       // cDisch
        return error_setInpError(ERR_NUMBER, tok[5]);
    x[3] = 0.0;
    if ( ntoks >= 7 )
    {
        m = findmatch(tok[6], NoYesWords);               
        if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[6]);
        x[3] = m;                                          // flap gate
    }

    // --- add parameters to data base
    Link[j].ID = id;
    link_setParams(j, ORIFICE, n1, n2, k, x);
    return 0;
}

//=============================================================================

void  orifice_validate(int j, int k)
//
//  Input:   j = link index
//           k = orifice index
//  Output:  none
//  Purpose: validates an orifice's properties
//
{
    int err = 0;
    float qFull;

    // --- check for valid xsection
    if ( Link[j].xsect.type != RECT_CLOSED
    &&   Link[j].xsect.type != CIRCULAR ) err = ERR_REGULATOR_SHAPE;
    if ( err > 0 )
    {
        report_writeErrorMsg(err, Link[j].ID);
        return;
    }

    // --- compute partial flow adjustment
    qFull = orifice_getFlow(j, k, Link[j].setting*Link[j].xsect.yFull, 1.0);
    Orifice[k].cFull = qFull / Link[j].xsect.aFull;

    // --- compute an equivalent length
    Orifice[k].length = 2.0 * RouteStep * sqrt(GRAVITY * Link[j].xsect.yFull);

///////////////////////////////////////////////////////////////////////////
////  Length should be max. of equiv. length & 200 ft. (RD - 7/5/06 )  ////
///////////////////////////////////////////////////////////////////////////
    Orifice[k].length = MAX(200.0, Orifice[k].length);

    Orifice[k].surfArea = 0.0;
}

//=============================================================================

float orifice_getInflow(int j)
//
//  Input:   j = link index
//  Output:  orifice flow rate (cfs)
//  Purpose: finds the flow through an orifice.
//
{
    int k, n1, n2;
    float head, h1, h2, y1, dir;
    float f, hcrest, hcrown;

    // --- get indexes of end nodes and link's orifice
    n1 = Link[j].node1;
    n2 = Link[j].node2;
    k  = Link[j].subIndex;

    // --- find heads at upstream & downstream nodes
    if ( RouteModel == DW )
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n2].newDepth + Node[n2].invertElev;
    }
    else
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n1].invertElev;
    }
    dir = (h1 >= h2) ? +1.0 : -1.0; 
           
    // --- exchange h1 and h2 for reverse flow
    y1 = Node[n1].newDepth;
    if ( dir < 0.0 )
    {
        head = h1;
        h1 = h2;
        h2 = head;
        y1 = Node[n2].newDepth;
    }

    // --- compute elevations of orifice crest and crown
    if ( Orifice[k].type == SIDE_ORIFICE )
    {
        hcrest = Node[n1].invertElev + Link[j].z1;
        hcrown = hcrest + Link[j].xsect.yFull;
    }

////////////////////////
//  LR - revised 6/14/05
////////////////////////
    if ( Orifice[k].type == BOTTOM_ORIFICE )
    {
        hcrest = Node[n1].invertElev + Link[j].z1;
        hcrown = hcrest;
    }
    
    // --- compute head on orifice & fraction full
    head = h1 - MAX(h2, hcrest);
    if ( h1 < hcrown ) f = (h1 - hcrest) / Link[j].xsect.yFull;
    else f = 1.0;

    // --- return if head is negligible or flap gate closed
    if ( head <= FUDGE || y1 <= FUDGE ||
         (Link[j].hasFlapGate && dir < 0.0) )
    {
        Link[j].newDepth = 0.0;
        Link[j].flowClass = DRY;
        Orifice[k].surfArea = FUDGE * Orifice[k].length;
        return 0.0;
    }

    // --- determine flow depth & flow class
    Link[j].newDepth = f * Link[j].xsect.yFull;
    Link[j].flowClass = SUBCRITICAL;
    if ( hcrest > h2 )
    {
        if ( dir == 1.0 ) Link[j].flowClass = DN_CRITICAL;
        else              Link[j].flowClass = UP_CRITICAL;
    }

    // --- update surface area & compute flow
    Orifice[k].surfArea = 
        xsect_getWofY(&Link[j].xsect, f*Link[j].xsect.yFull) *
        Orifice[k].length;
    return dir * orifice_getFlow(j, k, head, f);
}

//=============================================================================

float orifice_getFlow(int j, int k,  float head, float f)
//
//  Input:   j    = link index
//           k    = orifice index
//           head = head across orifice
//           f    = fraction of orifice area open
//  Output:  returns flow through an orifice
//  Purpose: computes flow through an orifice given head.
//
{
    float area, q;

    // --- find area (possibly reduced by control setting)
    area = Link[j].setting * Link[j].xsect.aFull;

    // --- case where orifice is only partly full
    if ( f < 1.0 )
    {
         q = area * Orifice[k].cFull * pow(f, 1.5);
         Link[j].dqdh = 1.5 * q / f / Link[j].xsect.yFull;
    }

    // --- case where orifice is submerged
    else
    {
        q = Orifice[k].cDisch * area * sqrt(2.0 * GRAVITY * head);
        Link[j].dqdh = q / (2.0 * head);
    }
    return q;
}


//=============================================================================
//                           W E I R   M E T H O D S
//=============================================================================

int   weir_readParams(int j, int k, char* tok[], int ntoks)
//
//  Input:   j = link index
//           k = weir index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads weir parameters from a tokenized line of input.
//
{
    int   m;
    int   n1, n2;
    float x[6];
    char* id;

    // --- check for valid ID and end node IDs
    if ( ntoks < 6 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LINK, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    n1 = project_findObject(NODE, tok[1]);
    if ( n1 < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n2 = project_findObject(NODE, tok[2]);
    if ( n2 < 0 ) return error_setInpError(ERR_NAME, tok[2]);

    // --- parse weir parameters
    m = findmatch(tok[3], WeirTypeWords);
    if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[3]);
    x[0] = m;                                              // type

//////////////////////////////////////////////////////////////////
//  Modified to check for negative height & cDisch. (LR - 7/5/06 )
//////////////////////////////////////////////////////////////////
    if ( ! getFloat(tok[4], &x[1]) || x[1] < 0.0 )         // height
        return error_setInpError(ERR_NUMBER, tok[4]);
    if ( ! getFloat(tok[5], &x[2]) || x[2] < 0.0 )         // cDisch1
        return error_setInpError(ERR_NUMBER, tok[5]);
    x[3] = 0.0;
    x[4] = 0.0;
    x[5] = 0.0;
    if ( ntoks >= 7 )
    {
        m = findmatch(tok[6], NoYesWords);             
        if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[6]);
        x[3] = m;                                          // flap gate
    }


/////////////////////////////////////////////////////////
//  Modified to check for negative values. (LR - 7/5/06 )
/////////////////////////////////////////////////////////

    if ( ntoks >= 8 )
    {
        if ( ! getFloat(tok[7], &x[4]) || x[4] < 0.0 )     // endCon
            return error_setInpError(ERR_NUMBER, tok[7]);
    }
    if ( ntoks >= 9 )
    {
        if ( ! getFloat(tok[8], &x[5]) || x[5] < 0.0 )     // cDisch2
            return error_setInpError(ERR_NUMBER, tok[8]);
    }

    // --- add parameters to data base
    Link[j].ID = id;
    link_setParams(j, WEIR, n1, n2, k, x);
    return 0;
}

//=============================================================================

void  weir_validate(int j, int k)
//
//  Input:   j = link index
//           k = weir index
//  Output:  none
//  Purpose: validates a weir's properties
//
{
    int err = 0;
    float q, q1, q2, head;

    // --- check for valid cross section
    switch ( Weir[k].type)
    {
      case TRANSVERSE_WEIR:
      case SIDEFLOW_WEIR:
        if ( Link[j].xsect.type != RECT_OPEN ) err = ERR_REGULATOR_SHAPE;
        Weir[k].slope = 0.0;
        break;
        
      case VNOTCH_WEIR:
        if ( Link[j].xsect.type != TRIANGULAR ) err = ERR_REGULATOR_SHAPE;
        else
        {
            Weir[k].slope = Link[j].xsect.aFull /
                            Link[j].xsect.yFull / Link[j].xsect.yFull; 
        }
        break;

      case TRAPEZOIDAL_WEIR:
        if ( Link[j].xsect.type != TRAPEZOIDAL ) err = ERR_REGULATOR_SHAPE;
        else
        {
            Weir[k].slope = (Link[j].xsect.aFull - 
                             Link[j].xsect.wMax * Link[j].xsect.yFull) /
                            (Link[j].xsect.yFull * Link[j].xsect.yFull);
        }
        break;
    }
    if ( err > 0 )
    {
        report_writeErrorMsg(err, Link[j].ID);
        return;
    }

    // --- compute an equivalent length
    Weir[k].length = 2.0 * RouteStep * sqrt(GRAVITY * Link[j].xsect.yFull);

///////////////////////////////////////////////////////////////////////////
////  Length should be max. of equiv. length & 200 ft. (RD - 7/5/06 )  ////
///////////////////////////////////////////////////////////////////////////
    Weir[k].length = MAX(200.0, Weir[k].length);
    Weir[k].surfArea = 0.0;

    // --- find flow through weir when water level equals weir height
    head = Link[j].xsect.yFull;
    weir_getFlow(j, k, head, 1.0, Link[j].hasFlapGate, &q1, &q2);
    q = q1 + q2;

    // --- compute orifice coeff. (for CFS flow units)
    Weir[k].cSurcharge = q / (Link[j].xsect.aFull * sqrt(2.0 * GRAVITY * head));
}

//=============================================================================

float weir_getInflow(int j)
//
//  Input:   j = link index
//  Output:  returns weir flow rate (cfs)
//  Purpose: finds the flow over a weir.
//
{
    int   n1;           // index of upstream node
    int   n2;           // index of downstream node
    int   k;            // index of weir
    float q1;           // flow through central part of weir (cfs)
    float q2;           // flow through end sections of weir (cfs)
    float head;         // head on weir (ft)
    float h1;           // upstrm nodal head (ft)
    float h2;           // downstrm nodal head (ft)
    float hcrest;       // head at weir crest (ft)
    float hcrown;       // head at weir crown (ft)
    float y;            // water depth in weir (ft)
    float dir;          // direction multiplier
    float ratio;
    float weirPower[] = {1.5,       // transverse weir
                         5./3.,     // side flow weir
                         2.5,       // v-notch weir
                         1.5};      // trapezoidal weir

    n1 = Link[j].node1;
    n2 = Link[j].node2;
    k  = Link[j].subIndex;
    if ( RouteModel == DW )
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n2].newDepth + Node[n2].invertElev;
    }
    else
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n1].invertElev;
    }
    dir = (h1 > h2) ? +1.0 : -1.0;            

////////////////////////
//  LR - revised 6/14/05
////////////////////////
    // --- exchange h1 and h2 for reverse flow
    y = Node[n1].newDepth;
    if ( dir < 0.0 )
    {
        head = h1;
        h1 = h2;
        h2 = head;
        y = Node[n2].newDepth;
    }

    // --- find head of weir's crest and crown
    hcrest = Node[n1].invertElev + Weir[k].crestHt;
    hcrown = hcrest + Link[j].xsect.yFull;

    // --- adjust crest ht. for partially open weir
    hcrest += (1.0 - Link[j].setting) * Link[j].xsect.yFull;

    // --- compute head relative to weir crest
    head = h1 - hcrest;

////////////////////////
//  LR - revised 6/14/05
////////////////////////
    // --- return if head is negligible or flap gate closed
    Link[j].dqdh = 0.0;
    if ( head <= FUDGE || hcrest >= hcrown ||
         //y <= FUDGE || 
         (Link[j].hasFlapGate && dir < 0.0) )
    {
        Link[j].newDepth = 0.0;
        Link[j].flowClass = DRY;
        return 0.0;
    }

    // --- determine flow class
    Link[j].flowClass = SUBCRITICAL;
    if ( hcrest > h2 )
    {
        if ( dir == 1.0 ) Link[j].flowClass = DN_CRITICAL;
        else              Link[j].flowClass = UP_CRITICAL;
    }

    // --- compute new equivalent surface area
    y = Link[j].xsect.yFull - (hcrown - MIN(h1, hcrown));
    Weir[k].surfArea = xsect_getWofY(&Link[j].xsect, y) * Weir[k].length;

    // --- if under surcharge condition then use equiv. orifice eqn.
    if ( h1 >= hcrown )
    {
        head = h1 - MAX(h2, hcrest);
        q1 = dir * Weir[k].cSurcharge *
             weir_getOpenArea(j, hcrown - hcrest) *
             sqrt(2.0 * GRAVITY * head);
        Link[j].dqdh = q1 / (2.0 * head);
        Link[j].newDepth = Link[j].xsect.yFull;
        return q1;
    }

    // --- otherwise use weir eqn. to find flows through central (q1)
    //     and end sections (q2) of weir
    weir_getFlow(j, k, head, dir, Link[j].hasFlapGate, &q1, &q2);
    Link[j].dqdh = weir_getdqdh(k, dir, head, q1, q2);

    // --- apply Villemonte eqn. to correct for submergence
    if ( h2 > hcrest )
    {
        ratio = (h2 - hcrest) / (h1 - hcrest);
        q1 *= pow( (1.0 - pow(ratio, weirPower[Weir[k].type])), 0.385);
        if ( q2 > 0.0 )
            q2 *= pow( (1.0 - pow(ratio, weirPower[VNOTCH_WEIR])), 0.385);
    }

   // --- return total flow through weir
   Link[j].newDepth = h1 - (Node[n1].invertElev + Weir[k].crestHt);
   return dir * (q1 + q2);
}

//=============================================================================

void weir_getFlow(int j, int k,  float head, float dir, int hasFlapGate,
                 float* q1, float* q2)
//
//  Input:   j    = link index
//           k    = weir index
//           head = head across weir (ft)
//           dir  = flow direction indicator
//           hasFlapGate = flap gate indicator
//  Output:  q1 = flow through central portion of weir (cfs)
//           q2 = flow through end sections of weir (cfs)
//  Purpose: computes flow over weir given head.
//
{
    float length;
    float h;
    float y;
    float hLoss;
    float area;
    float veloc;

    // --- convert weir length & head to original units
    length = Link[j].xsect.wMax * UCF(LENGTH);
    h = head * UCF(LENGTH);

    // --- reduce length when end contractions present
    length -= 0.1 * Weir[k].endCon * h;
    length = MAX(length, 0.0);

    // --- q1 = flow through central portion of weir,
    //     q2 = flow through end sections of trapezoidal weir
    *q1 = 0.0;
    *q2 = 0.0;

    // --- use appropriate formula for weir flow
    switch (Weir[k].type)
    {
      case TRANSVERSE_WEIR:
        *q1 = Weir[k].cDisch1 * length * pow(h, 1.5);
        break;

      case SIDEFLOW_WEIR:
        // --- weir behaves as a transverse weir under reverse flow
        if ( dir < 0.0 )
            *q1 = Weir[k].cDisch1 * length * pow(h, 1.5);
        else
            *q1 = Weir[k].cDisch1 * length * pow(h, 5./3.);
        break;

      case VNOTCH_WEIR:
        *q1 = Weir[k].cDisch1 * Weir[k].slope * pow(h, 2.5);
        break;

      case TRAPEZOIDAL_WEIR:
        y = (1.0 - Link[j].setting) * Link[j].xsect.yFull;
        length = xsect_getWofY(&Link[j].xsect, y) * UCF(LENGTH);
        *q1 = Weir[k].cDisch1 * length * pow(h, 1.5);
        *q2 = Weir[k].cDisch2 * Weir[k].slope * pow(h, 2.5);
    }

    // --- convert CMS flows to CFS
    if ( UnitSystem == SI )
    {
        *q1 /= M3perFT3;
        *q2 /= M3perFT3;
    }

    // --- apply ARMCO adjustment for headloss from flap gate
    if ( hasFlapGate )
    {
        // --- compute flow area & velocity for current weir flow
        area = xsect_getAofY(&Link[j].xsect, head);
        veloc = (*q1 + *q2) / area;

        // --- compute headloss and subtract from original head
        hLoss = (4.0 / GRAVITY) * veloc * veloc *
                 exp(-1.15 * veloc / sqrt(head) );
        head = head - hLoss;
        if ( head < 0.0 ) head = 0.0;

        // --- make recursive call to this function, with hasFlapGate
        //     set to false, to find flow values at adjusted head value
        weir_getFlow(j, k, head, dir, FALSE, q1, q2);
    }
}

//=============================================================================

float weir_getOpenArea(int j, float y)
//
//  Input:   j = link index
//           y = height from weir crest to weir crown (ft)
//  Output:  returns area between weir crest and crown (ft2)
//  Purpose: finds flow area of partially open weir
//
{
    y = Link[j].xsect.yFull - y;
    return Link[j].xsect.aFull - xsect_getAofY(&Link[j].xsect, y);
}

//=============================================================================

float  weir_getdqdh(int k, float dir, float h, float q1, float q2)
{
    float q1h;
    float q2h;

    if ( fabs(h) < FUDGE ) return 0.0;
    q1h = fabs(q1/h);
    q2h = fabs(q2/h);
    switch (Weir[k].type)
    {
      case TRANSVERSE_WEIR: return 1.5 * q1h;

      case SIDEFLOW_WEIR:
        // --- weir behaves as a transverse weir under reverse flow
        if ( dir < 0.0 ) return 1.5 * q1h;
        else return 5./3. * q1h;

      case VNOTCH_WEIR: return 2.5 * q1h;

      case TRAPEZOIDAL_WEIR: return 1.5 * q1h + 2.5 * q2h;
    }
    return 0.0;
}
 

//=============================================================================
//               O U T L E T    D E V I C E    M E T H O D S
//=============================================================================

int outlet_readParams(int j, int k, char* tok[], int ntoks)
//
//  Input:   j = link index
//           k = outlet index
//           tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads outlet parameters from a tokenized  line of input.
//
{
    int   i, m, n;
    int   n1, n2;
    float x[5];
    char* id;

    // --- check for valid ID and end node IDs
    if ( ntoks < 6 ) return error_setInpError(ERR_ITEMS, "");
    id = project_findID(LINK, tok[0]);
    if ( id == NULL ) return error_setInpError(ERR_NAME, tok[0]);
    n1 = project_findObject(NODE, tok[1]);
    if ( n1 < 0 ) return error_setInpError(ERR_NAME, tok[1]);
    n2 = project_findObject(NODE, tok[2]);
    if ( n2 < 0 ) return error_setInpError(ERR_NAME, tok[2]);

/////////////////////////////////////////////////////////
//  Modified to check for negative height. (LR - 7/5/06 )
/////////////////////////////////////////////////////////
    // --- get height above invert
    if ( ! getFloat(tok[3], &x[0]) || x[0] < 0.0 )
        return error_setInpError(ERR_NUMBER, tok[3]);

    // --- see if outlet flow relation is tabular or functional
    m = findmatch(tok[4], RelationWords);
    if ( m < 0 ) return error_setInpError(ERR_KEYWORD, tok[4]);
    x[1] = 0.0;
    x[2] = 0.0;
    x[3] = -1.0;
    x[4] = 0.0;

    // --- get params. for functional outlet device
    if ( m == FUNCTIONAL )
    {
        if ( ntoks < 7 ) return error_setInpError(ERR_ITEMS, "");
        if ( ! getFloat(tok[5], &x[1]) )
            return error_setInpError(ERR_NUMBER, tok[5]);
        if ( ! getFloat(tok[6], &x[2]) )
            return error_setInpError(ERR_NUMBER, tok[6]);
        n = 7;
    }

    // --- get name of outlet rating curve
    else
    {
        i = project_findObject(CURVE, tok[5]);
        if ( i < 0 ) return error_setInpError(ERR_NAME, tok[5]);
        x[3] = i;
        n = 6;
    }

    // --- check if flap gate specified
    if ( ntoks > n)
    {
        i = findmatch(tok[n], NoYesWords);               
        if ( i < 0 ) return error_setInpError(ERR_KEYWORD, tok[n]);
        x[4] = i;
    }

    // --- add parameters to data base
    Link[j].ID = id;
    link_setParams(j, OUTLET, n1, n2, k, x);
    return 0;
}

//=============================================================================

float outlet_getInflow(int j)
//
//  Input:   j = link index
//  Output:  outlet flow rate (cfs)
//  Purpose: finds the flow through an outlet.
//
{
    int k, n1, n2;
    float head, hcrest, h1, h2, y1, dir;

    // --- get indexes of end nodes
    n1 = Link[j].node1;
    n2 = Link[j].node2;
    k  = Link[j].subIndex;

    // --- find heads at upstream & downstream nodes
    if ( RouteModel == DW )
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n2].newDepth + Node[n2].invertElev;
    }
    else
    {
        h1 = Node[n1].newDepth + Node[n1].invertElev;
        h2 = Node[n1].invertElev;
    }
    dir = (h1 >= h2) ? +1.0 : -1.0; 
           
    // --- exchange h1 and h2 for reverse flow
    y1 = Node[n1].newDepth;
    if ( dir < 0.0 )
    {
        h1 = h2;
        y1 = Node[n2].newDepth;
    }
    hcrest = Node[n1].invertElev + Link[j].z1;
    head = h1 - hcrest;
    if ( head <= FUDGE || y1 <= FUDGE ||
         (Link[j].hasFlapGate && dir < 0.0) )
    {
        Link[j].newDepth = 0.0;
        Link[j].flowClass = DRY;
        return 0.0;
    }
    Link[j].newDepth = head;
    Link[j].flowClass = SUBCRITICAL;
    return dir * Link[j].setting * outlet_getFlow(k, head);
}

//=============================================================================

float outlet_getFlow(int k, float head)
//
//  Input:   k    = outlet index
//           head = head across outlet (ft)
//  Output:  returns outlet flow rate (cfs)
//  Purpose: computes flow rate through an outlet given head.
//
{
    int m;
    float h;

    // --- convert head to original units
    h = head * UCF(LENGTH);

    // --- look-up flow in rating curve table if provided
    m = Outlet[k].qCurve;
    if ( m >= 0 ) return table_lookup(&Curve[m], h) / UCF(FLOW);
    
    // --- otherwise use function to find flow
    else return Outlet[k].qCoeff * pow(h, Outlet[k].qExpon) / UCF(FLOW);
}

//=============================================================================

