//-----------------------------------------------------------------------------
//   dynwave.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             7/5/06   (Build 5.0.008)
//             9/19/06  (Build 5.0.009)
//   Author:   L. Rossman
//             R. Dickinson
//
//   Dynamic wave flow routing functions.
//
//   This module solves the dynamic wave flow routing equations using
//   Picard Iterations (i.e., a method of successive approximations)
//   to solve the explicit form of the continuity and momentum equations
//   for conduits.
//-----------------------------------------------------------------------------

#include <malloc.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//     Constants 
//-----------------------------------------------------------------------------
static const float MINSURFAREA =  12.566;   // min. nodal surface area (~4 ft diam.)
static const float MAXVELOCITY =  50.;      // max. allowable velocity (ft/sec)
static const float MINTIMESTEP =  0.5;      // min. time step (sec)
static const float OMEGA       =  0.5;      // under-relaxation parameter
static const float STOP_TOL    =  0.005;    // Picard iteration stop criterion
static const int   MAXSTEPS    =  4;        // max. number of Picard iterations

//-----------------------------------------------------------------------------
//  Data Structures
//-----------------------------------------------------------------------------
typedef struct 
{
    char    converged;                 // TRUE if iterations for a node done
    float   newSurfArea;               // current surface area (ft2)
    float   oldSurfArea;               // previous surface area (ft2)
    float   sumdqdh;                   // sum of dqdh from adjoining links
    float   dYdT;                      // change in depth w.r.t. time (ft/sec)
} TXnode;

typedef struct
{
    char    bypassed;                  // TRUE if can bypass calcs. for a link
    float   surfArea1;                 // surf. area at upstrm end of link (ft2)
    float   surfArea2;                 // surf. area at dnstrm end of link (ft2)
} TXlink;

//-----------------------------------------------------------------------------
//  Shared Variables
//-----------------------------------------------------------------------------
static float   MinSurfAreaFt2;         // actual min. nodal surface area (ft2)
static float   VariableStep;           // size of variable time step (sec)
static float   Omega;                  // actual under-relaxation parameter
static float   CriticalDepth;          // critical flow depth (ft)
static float   NormalDepth;            // normal flow depth (ft)
static float   Fasnh;                  // fraction between norm. & crit. depth
static int     Converged;              // TRUE if Picard iterations converged
static int     Steps;                  // number of Picard iterations
static TXnode* Xnode;
static TXlink* Xlink;

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  dynwave_init            (called by flowrout_init)
//  dynwave_getRoutingStep  (called by flowrout_getRoutingStep)
//  dynwave_execute         (called by flowrout_execute)

//-----------------------------------------------------------------------------
//  Function declarations
//-----------------------------------------------------------------------------
static void   execRoutingStep(int links[], float dt);
static void   initNodeState(int i);
static void   findConduitFlow(int i, float dt);
static void   findNonConduitFlow(int i, float dt);
static float  getModPumpFlow(int i, float q, float dt);
static void   updateNodeFlows(int i, float q);

static float  getConduitFlow(int link, float qin, float dt);
static int    getFlowClass(int link, float q, float h1, float h2,
              float y1, float y2);
static void   findSurfArea(int link, float length, float* h1, float* h2,
              float* y1, float* y2);
static float  findLocalLosses(int link, float a1, float a2, float aMid,
              float q);
static void   findNonConduitSurfArea(int link);

static float  getWidth(TXsect* xsect, float y);
static float  getArea(TXsect* xsect, float y);
static float  getHydRad(TXsect* xsect, float y);
static float  checkFlapGate(int j, int n1, int n2, float q);
static float  checkNormalFlow(int j, float q, float h1, float h2, float a1,
              float a2, float r1);

static void   setNodeDepth(int node, float dt);
static void   setWetWellVolume(int i, float dV, float dt);
static float  getFloodedDepth(int i, int canPond, float dV, float yMax,
              float dt);

static float  getVariableStep(float maxStep);
static float  getLinkStep(float tMin, int *minLink);
static float  getNodeStep(float tMin, int *minNode);

//=============================================================================

void dynwave_init()
//
//  Input:   none
//  Output:  none
//  Purpose: initializes dynamic wave routing method.
//
{
    int i;

    VariableStep = 0.0;
    if ( MinSurfArea == 0.0 ) MinSurfAreaFt2 = MINSURFAREA;
    else MinSurfAreaFt2 = MinSurfArea / UCF(LENGTH) / UCF(LENGTH);
    Xnode = (TXnode *) calloc(Nobjects[NODE], sizeof(TXnode));
    Xlink = (TXlink *) calloc(Nobjects[LINK], sizeof(TXlink));

    // --- initialize node surface areas
    for (i = 0; i < Nobjects[NODE]; i++ )
    {
        Xnode[i].newSurfArea = 0.0;
        Xnode[i].oldSurfArea = 0.0;
    }
    for (i = 0; i < Nobjects[LINK]; i++)
    {
        Link[i].flowClass = DRY;
        Link[i].dqdh = 0.0;
    }
}

//=============================================================================

void  dynwave_close()
//
//  Input:   none
//  Output:  none
//  Purpose: frees memory allocated for dynamic wave routing method.
//
{
    FREE(Xnode);
    FREE(Xlink);
}

//=============================================================================

float dynwave_getRoutingStep(float fixedStep)
//
//  Input:   fixedStep = user-supplied fixed time step (sec)
//  Output:  returns routing time step (sec)
//  Purpose: computes variable routing time step if applicable.
//
{
    // --- use user-supplied fixed step if variable step option turned off
    //     or if its smaller than the min. allowable variable time step
    if ( CourantFactor == 0.0 ) return fixedStep;
    if ( fixedStep < MINTIMESTEP ) return fixedStep;

    // --- at start of simulation (when current variable step is zero)
    //     use the minimum allowable time step
    if ( VariableStep == 0.0 )
    {
        VariableStep = MINTIMESTEP;
    }

    // --- otherwise compute variable step based on current flow solution
    else VariableStep = getVariableStep(fixedStep);

    // --- adjust step to be a multiple of a millisecond
    VariableStep = floor(1000.0 * VariableStep) / 1000.0;
    return VariableStep;
}

//=============================================================================

int dynwave_execute(int links[], float tStep)
//
//  Input:   links = array of topo sorted links indexes
//           tStep = time step (sec)
//  Output:  returns number of iterations used
//  Purpose: routes flows through drainage network over current time step.
//
{
    int i;

    // --- initialize
    if ( ErrorCode ) return 0;
    Steps = 0;
    Converged = FALSE;
    Omega = OMEGA;
    for (i=0; i<Nobjects[NODE]; i++)
    {
        Xnode[i].converged = FALSE;
        Xnode[i].dYdT = 0.0;
    }
    for (i=0; i<Nobjects[LINK]; i++)
    {
        Xlink[i].bypassed = FALSE;
        Xlink[i].surfArea1 = 0.0;
        Xlink[i].surfArea2 = 0.0;
    }

    // --- a2 preserves conduit area from solution at last time step
    for ( i=0; i<Nlinks[CONDUIT]; i++) Conduit[i].a2 = Conduit[i].a1;

    // --- keep iterating until convergence 
    while ( Steps < MAXSTEPS )
    {
        // --- execute a routing step & check for nodal convergence
        execRoutingStep(links, tStep);
        Steps++;
        if ( Steps > 1 )
        {
            if ( Converged ) break;

            // --- check if link calculations can be skipped in next step
            for (i=0; i<Nobjects[LINK]; i++)
            {
                if ( Xnode[Link[i].node1].converged &&
                     Xnode[Link[i].node2].converged )
                     Xlink[i].bypassed = TRUE;
                else Xlink[i].bypassed = FALSE;
            }
        }
    }
    return Steps;
}

//=============================================================================

void execRoutingStep(int links[], float dt)
//
//  Input:   links = array of link indexes
//           dt    = time step (sec)
//  Output:  none
//  Purpose: solves momentum eq. in links and continuity eq. at nodes
//           over specified time step.
//
{
    int    i;                          // node or link index
    float  yOld;                       // old node depth (ft)

    // --- re-initialize state of each node
    for ( i = 0; i < Nobjects[NODE]; i++ ) initNodeState(i);
    Converged = TRUE;

    // --- find new flows in conduit links and non-conduit links
    for ( i=0; i<Nobjects[LINK]; i++) findConduitFlow(links[i], dt);
    for ( i=0; i<Nobjects[LINK]; i++) findNonConduitFlow(links[i], dt);

    // --- compute outfall depths based on flow in connecting link
    for ( i = 0; i < Nobjects[LINK]; i++ ) link_setOutfallDepth(i);

    // --- compute new depth for all non-outfall nodes and determine if
    //     depth change from previous iteration is below tolerance
    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        if ( Node[i].type == OUTFALL ) continue;
        yOld = Node[i].newDepth;
        setNodeDepth(i, dt);
        Xnode[i].converged = TRUE;
        if ( fabs(yOld - Node[i].newDepth) > STOP_TOL )
        {
            Converged = FALSE;
            Xnode[i].converged = FALSE;
        }
    }
}

//=============================================================================

void initNodeState(int i)
//
//  Input:   i = node index
//  Output:  none
//  Purpose: initializes node's surface area, inflow & outflow
//
{
    // --- initialize nodal surface area
    if ( AllowPonding )
    {
        Xnode[i].newSurfArea = node_getPondedArea(i, Node[i].newDepth);
    }
    else
    {
        Xnode[i].newSurfArea = node_getSurfArea(i, Node[i].newDepth);
    }

    if ( Xnode[i].newSurfArea < MinSurfAreaFt2 )
    {
        Xnode[i].newSurfArea = MinSurfAreaFt2;
    }

    // --- initialize nodal inflow & outflow
    Node[i].inflow = Node[i].newLatFlow;
    Node[i].outflow = 0.0;
    Xnode[i].sumdqdh = 0.0;
}

//=============================================================================

void findConduitFlow(int i, float dt)
//
//  Input:   i = link index
//           dt = time step (sec)
//  Output:  none
//  Purpose: finds new flow in a conduit-type link
//
{
    float  qOld;                       // old link flow (cfs)
    float  barrels;                    // number of barrels in conduit

    // --- do nothing if link not a conduit
    if ( Link[i].type != CONDUIT || Link[i].xsect.type == DUMMY) return;

    // --- get link flow from last "full" time step
    qOld = Link[i].oldFlow;

    // --- solve momentum eqn. to update conduit flow
    if ( !Xlink[i].bypassed )
    {
////////////////////////////////////////
//  Initialize dqdh here. (LR - 9/19/06)
////////////////////////////////////////
        Link[i].dqdh = 0.0;

        Link[i].newFlow = getConduitFlow(i, qOld, dt);
    }
    // NOTE: if link was bypassed, then its flow and surface area values
    //       from the previous iteration will still be valid.

    // --- add surf. area contributions to upstream/downstream nodes
    barrels = Conduit[Link[i].subIndex].barrels;
    Xnode[Link[i].node1].newSurfArea += Xlink[i].surfArea1 * barrels;
    Xnode[Link[i].node2].newSurfArea += Xlink[i].surfArea2 * barrels;

//////////////////////////////////////////////////////////
//  Updating of node sumdqdh moved to here. (LR - 9/19/06)
//////////////////////////////////////////////////////////
    // --- update summed value of dqdh at each end node
    Xnode[Link[i].node1].sumdqdh += Link[i].dqdh;
    Xnode[Link[i].node2].sumdqdh += Link[i].dqdh;

    // --- update outflow/inflow at upstream/downstream nodes
    updateNodeFlows(i, Link[i].newFlow);
}

//=============================================================================

void findNonConduitFlow(int i, float dt)
//
//  Input:   i = link index
//           dt = time step (sec)
//  Output:  none
//  Purpose: finds new flow in a non-conduit-type link
//
{
    float  qLast;                      // previous link flow (cfs)
    float  qNew;                       // new link flow (cfs)
    int    k, m;

    // --- ignore non-dummy conduit links
    if ( Link[i].type == CONDUIT && Link[i].xsect.type != DUMMY ) return;

    // --- update flow in link if not bypassed
    if ( !Xlink[i].bypassed )
    {
        // --- get link flow from last iteration
        qLast = Link[i].newFlow;

////////////////////////////////////////
//  Initialize dqdh here. (LR - 9/19/06)
////////////////////////////////////////
        Link[i].dqdh = 0.0;

        // --- get new inflow to link from its upstream node
        //     (link_getInflow returns 0 if flap gate closed or pump is offline)
        qNew = link_getInflow(i);
        if ( Link[i].type == PUMP ) qNew = getModPumpFlow(i, qNew, dt);

        // --- find surface area at each end of link
        findNonConduitSurfArea(i);

        // --- apply under-relaxation with flow from previous iteration;
        // --- do not allow flow to change direction without first being 0
        if ( Steps > 0 )
        {
            qNew = (1.0 - Omega) * qLast + Omega * qNew;
            if ( qNew * qLast < 0.0 ) qNew = 0.001 * SGN(qNew);
        }
        Link[i].newFlow = qNew;
    }

    // --- add surf. area contributions to upstream/downstream nodes
    Xnode[Link[i].node1].newSurfArea += Xlink[i].surfArea1;
    Xnode[Link[i].node2].newSurfArea += Xlink[i].surfArea2;

//////////////////////////////////////////////////////////
//  Updating of node sumdqdh moved to here. (LR - 9/19/06)
//////////////////////////////////////////////////////////
    // --- update summed value of dqdh at each end node
    //     (but not for discharge node of Type 4 pumps)
    Xnode[Link[i].node1].sumdqdh += Link[i].dqdh;
    if ( Link[i].type == PUMP )
    {
        k = Link[i].subIndex;
        m = Pump[k].pumpCurve;
        if ( Curve[m].curveType != PUMP4_CURVE )
            Xnode[Link[i].node2].sumdqdh += Link[i].dqdh;
    }
    else Xnode[Link[i].node2].sumdqdh += Link[i].dqdh;

    // --- update outflow/inflow at upstream/downstream nodes
    updateNodeFlows(i, Link[i].newFlow);
}

//=============================================================================

float getModPumpFlow(int i, float q, float dt)
//
//  Input:   i = link index
//           q = pump flow from pump curve (cfs)
//           dt = time step (sec)
//  Output:  returns modified pump flow rate (cfs)
//  Purpose: modifies pump curve pumping rate depending on amount of water
//           available at pump's inlet node.
//
{
    int    j = Link[i].node1;          // pump's inlet node index
    int    k = Link[i].subIndex;       // pump's index
    float  newNetInflow;               // inflow - outflow rate (cfs)
    float  netFlowVolume;              // inflow - outflow volume (ft3)
    float  y;                          // node depth (ft)

    if ( q == 0.0 ) return q;

    // --- case where inlet node is a storage node: 
    //     prevent node volume from going negative
    if ( Node[j].type == STORAGE ) return node_getMaxOutflow(j, q, dt); 

    // --- case where inlet is a non-storage node
    switch ( Pump[k].type )
    {
      // --- for Type1 pump, a volume is computed for inlet node,
      //     so make sure it doesn't go negative
      case TYPE1_PUMP:
        return node_getMaxOutflow(j, q, dt);

      // --- for other types of pumps, if pumping rate would make depth
      //     at upstream node negative, then set pumping rate = inflow
      case TYPE2_PUMP:
      case TYPE4_PUMP:
      case TYPE3_PUMP:
         newNetInflow = Node[j].inflow - Node[j].outflow - q;
         netFlowVolume = 0.5 * (Node[j].oldNetInflow + newNetInflow ) * dt;
         y = Node[j].oldDepth + netFlowVolume / Xnode[j].newSurfArea;
         if ( y <= 0.0 ) return Node[j].inflow;
    }
    return q;
}

//=============================================================================

void  findNonConduitSurfArea(int i)
//
//  Input:   i = link index
//  Output:  none
//  Purpose: finds the surface area contributed by a non-conduit
//           link to its upstream and downstream nodes.
//
{
    if ( Link[i].type == ORIFICE )
    {
        Xlink[i].surfArea1 = Orifice[Link[i].subIndex].surfArea / 2.;
    }

///////////////////////////////////////////////////////////////////
// Ignore weir surface area for SWMM 4 compatibility. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////
    //else if ( Link[i].type == WEIR )
    //{
    //    Xlink[i].surfArea1 = Weir[Link[i].subIndex].surfArea / 2.;
    //}
    else Xlink[i].surfArea1 = 0.0;
    Xlink[i].surfArea2 = Xlink[i].surfArea1;
    if ( Link[i].flowClass == UP_CRITICAL ) Xlink[i].surfArea1 = 0.0;
    if ( Link[i].flowClass == DN_CRITICAL ) Xlink[i].surfArea2 = 0.0;
}

//=============================================================================

void updateNodeFlows(int i, float q)
//
//  Input:   i = link index
//           q = link flow rate (cfs)
//  Output:  none
//  Purpose: updates cumulative inflow & outflow at link's end nodes.
//
{
    if ( q >= 0.0 )
    {
        Node[Link[i].node1].outflow += q;
        Node[Link[i].node2].inflow  += q;
    }
    else
    {
        Node[Link[i].node1].inflow   -= q;
        Node[Link[i].node2].outflow  -= q;
    }

//////////////////////////////////////////////////
//  This code was moved to the findConduitFlow and
//  findNonConduitFlow functions. - (LR - 9/19/06)
//////////////////////////////////////////////////
//    Xnode[Link[i].node1].sumdqdh += Link[i].dqdh;
//    Xnode[Link[i].node2].sumdqdh += Link[i].dqdh;
}

//=============================================================================

////////////////////////////////////////////////////////////////////////
////  Some comments were changed to clarify that the heads used in the
////  conduit flow computations are the heads of the flow depths in the
////  conduit, which are not necessarily the same as the heads at the
////  upstream/downstream nodes. (LR - 7/5/06 )
////////////////////////////////////////////////////////////////////////

float  getConduitFlow(int j, float qOld, float dt)
//
//  Input:   j        = link index
//           qOld     = flow from previous iteration (cfs)
//           dt       = time step (sec)
//  Output:  returns new flow value (cfs)
//  Purpose: updates flow in conduit link by solving finite difference
//           form of continuity and momentum equations.
//
{
    int   k;                           // index of conduit
    int   n1, n2;                      // indexes of end nodes
    float h1, h2;                      // upstream/dounstream flow heads (ft)
    float y1, y2;                      // upstream/downstream flow depths (ft)
    float a1, a2;                      // upstream/downstream flow areas (ft2)
    float r1;                          // upstream hyd. radius (ft)
    float yMid, rMid, aMid;            // mid-stream or avg. values of y, r, & a
    float qLast;                       // flow from previous iteration (cfs)
    float aLast;                       // area from previous time step (ft2)
    float v;                           // velocity (ft/sec)
    float sigma;                       // inertial damping factor
    float length;                      // effective conduit length (ft)
    float dq1, dq2, dq3, dq4, dq5;     // terms in momentum eqn.
    float denom;                       // denominator of flow update formula
    float q;                           // new flow value (cfs)
    float barrels;                     // number of barrels in conduit
    TXsect* xsect = &Link[j].xsect;    // ptr. to conduit's cross section data

    // --- get most current heads at upstream and downstream ends of conduit
    k =  Link[j].subIndex;
    n1 = Link[j].node1;
    n2 = Link[j].node2;
    h1 = Node[n1].newDepth + Node[n1].invertElev;
    h2 = Node[n2].newDepth + Node[n2].invertElev;
    h1 = MAX(h1, Node[n1].invertElev + Link[j].z1);
    h2 = MAX(h2, Node[n2].invertElev + Link[j].z2);

    // --- get unadjusted upstream and downstream flow depths in conduit
    //    (flow depth = head in conduit - elev. of conduit invert)
    y1 = h1 - (Node[n1].invertElev + Link[j].z1);
    y2 = h2 - (Node[n2].invertElev + Link[j].z2);
    y1 = MAX(y1, FUDGE);
    y2 = MAX(y2, FUDGE);

    // --- flow depths can't exceed full depth of conduit
    y1 = MIN(y1, xsect->yFull);
    y2 = MIN(y2, xsect->yFull);

    // --- get flow from last time step & previous iteration 
    barrels = Conduit[k].barrels;
    qOld /= barrels;
    qLast = Conduit[k].q1;

    // -- get area from solution at previous time step
    //    which was saved in Conduit[k].a2
    aLast = Conduit[k].a2;
    aLast = MAX(aLast, FUDGE);

    // --- use Courant-modified length instead of conduit's actual length
    length = Conduit[k].modLength;

    // --- find flow classification & corresponding surface area
    //     contributions to upstream and downstream nodes
    Link[j].flowClass = getFlowClass(j, qLast, h1, h2, y1, y2);
    findSurfArea(j, length, &h1, &h2, &y1, &y2);

    // --- compute area at each end of conduit & hyd. radius at upstream end
    a1 = getArea(xsect, y1);
    a2 = getArea(xsect, y2);
    r1 = getHydRad(xsect, y1);

    // --- compute area & hyd. radius at midpoint
    yMid = 0.5 * (y1 + y2);
    aMid = getArea(xsect, yMid);
    rMid = getHydRad(xsect, yMid);

    // --- alternate approach not currently used, but might produce better
    //     Bernoulli energy balance for steady flows
    //aMid = (a1+a2)/2.0;
    //rMid = (r1+getHydRad(xsect,y2))/2.0;

    // --- set new flow to zero if conduit is dry or if flap gate is closed
    if ( Link[j].flowClass == DRY ||
         Link[j].flowClass == UP_DRY ||
         Link[j].flowClass == DN_DRY ||
         Link[j].isClosed ||
         aMid <= FUDGE )
    {
        Conduit[k].a1 = 0.5 * (a1 + a2);
        Conduit[k].q1 = 0.0;;
        Conduit[k].q2 = 0.0;
        Link[j].dqdh  = GRAVITY * dt * aMid / length * barrels;
        Link[j].froude = 0.0;
        Link[j].newDepth = MIN(yMid, Link[j].xsect.yFull);
        Link[j].newVolume = Conduit[k].a1 * Conduit[k].length * barrels;
        return 0.0;
    }

    // --- compute velocity from last flow estimate
    v = qLast / aMid;
    if ( fabs(v) > MAXVELOCITY )  v = MAXVELOCITY * SGN(qLast);

    // --- compute Froude No.
    Link[j].froude = link_getFroude(j, v, yMid);
    if ( Link[j].flowClass == SUBCRITICAL &&
         Link[j].froude > 1.0 ) Link[j].flowClass = SUPCRITICAL;

    // --- find inertial damping factor (sigma)
    if ( InertDamping == NONE ) sigma = 1.0;
    else if ( InertDamping == SOME )
    {
        if      ( Link[j].froude <= 0.5 ) sigma = 1.0;
        else if ( Link[j].froude >= 1.0 ) sigma = 0.0;
        else    sigma = 2.0 * (1.0 - Link[j].froude);
    }
    else sigma = 0.0;

    // --- use full inertial damping if conduit is surcharged
    if ( h1 >= Node[n1].crownElev && h2 >= Node[n2].crownElev ) sigma = 0.0;

    // --- compute terms of momentum eqn.:
    // --- 1. friction slope term
    dq1 = dt * Conduit[k].roughFactor / pow(rMid, 1.33333) * fabs(v);

    // --- 2. energy slope term
    dq2 = dt * GRAVITY * aMid * (h2 - h1) / length;

    // --- 3 & 4. inertial terms
    dq3 = 0.0;
    dq4 = 0.0;
    if ( sigma > 0.0 )
    {
        dq3 = 2.0 * v * (aMid - aLast) * sigma;
        dq4 = dt * v * v * (a2 - a1) / length * sigma;
    }

    // --- 5. local losses term
    dq5 = 0.0;
    if ( Conduit[k].hasLosses )
    {
        dq5 = findLocalLosses(j, a1, a2, aMid, qLast) / 2.0 / length * dt;
    }

    // --- combine terms to find new conduit flow
    denom = 1.0 + dq1 + dq5;
    q = (qOld - dq2 + dq3 + dq4) / denom;

    // --- compute derivative of flow w.r.t. head
    Link[j].dqdh = 1.0 / denom  * GRAVITY * dt * aMid / length * barrels;

    // --- check if normal flow limitation applies
    q = checkNormalFlow(j, q, h1, h2, a1, a2, r1);

    // --- apply under-relaxation weighting between new & old flows;
    // --- do not allow change in flow direction without first being zero 
    if ( Steps > 0 )
    {
        q = (1.0 - Omega) * qLast + Omega * q;
        if ( q * qLast < 0.0 ) q = 0.001 * SGN(q);
    }

    // --- check if user-supplied flow limit applies
    if ( Link[j].qLimit > 0.0 )
    {
         if ( fabs(q) > Link[j].qLimit ) q = SGN(q) * Link[j].qLimit;
    }

    // --- check for reverse flow with closed flap gate
    q = checkFlapGate(j, n1, n2, q);

    // --- save new values of area, flow, depth, & volume
    Conduit[k].a1 = aMid;
    Conduit[k].q1 = q;
    Conduit[k].q2 = q;
    Link[j].newDepth  = MIN(yMid, xsect->yFull);
    aMid = (a1 + a2) / 2.0;
    aMid = MIN(aMid, xsect->aFull);
    Link[j].newVolume = aMid * Conduit[k].length * barrels;
    return q * barrels;
}

//=============================================================================

///////////////////////////////////////////////////////////////////////
////  This function was modified to account for the case where the
////  "offset" height for an outfall conduit might depend on the stage
////  elevation of the outfall's boundary condition. (LR - 7/5/06 )
///////////////////////////////////////////////////////////////////////

int getFlowClass(int j, float q, float h1, float h2, float y1, float y2)
//
//  Input:   j  = conduit link index
//           q  = current conduit flow (cfs)
//           h1 = head at upstream end of conduit (ft)
//           h2 = head at downstream end of conduit (ft)
//           y1 = upstream flow depth in conduit (ft)
//           y2 = downstream flow depth in conduit (ft)
//  Output:  returns flow classification code
//  Purpose: determines flow class for a conduit based on depths at each end.
//
{
    int    n1, n2;                     // indexes of upstrm/downstrm nodes
    int    flowClass;                  // flow classification code
    float  ycMin, ycMax;               // min/max critical depths (ft)
    float  z1, z2;                     // offsets of conduit inverts (ft)

    // --- get upstream & downstream node indexes
    n1 = Link[j].node1;
    n2 = Link[j].node2;

    // --- get upstream & downstream conduit invert offsets
    z1 = Link[j].z1;
    z2 = Link[j].z2;

    // --- base offset of an outfall conduit on outfall's depth
    if ( Node[n1].type == OUTFALL ) z1 = MAX(0.0, (z1 - Node[n1].newDepth));
    if ( Node[n2].type == OUTFALL ) z2 = MAX(0.0, (z2 - Node[n2].newDepth));

    // --- default class is SUBCRITICAL
    flowClass = SUBCRITICAL;
    Fasnh = 1.0;

    // --- case where both ends of conduit are wet
    if ( y1 > FUDGE && y2 > FUDGE )
    {
        if ( q < 0.0 )
        {
            // --- upstream end at critical depth if flow depth is
            //     below conduit's critical depth and an upstream 
            //     conduit offset exists
            if ( z1 > 0.0 )
            {
                NormalDepth   = link_getYnorm(j, fabs(q));
                CriticalDepth = link_getYcrit(j, fabs(q));
                ycMin = MIN(NormalDepth, CriticalDepth);
                if ( y1 < ycMin ) flowClass = UP_CRITICAL;
            }
        }

        // --- case of normal direction flow
        else
        {
            // --- downstream end at smaller of critical and normal depth
            //     if downstream flow depth below this and a downstream
            //     conduit offset exists
            if ( z2 > 0.0 )
            {
                NormalDepth = link_getYnorm(j, fabs(q));
                CriticalDepth = link_getYcrit(j, fabs(q));
                ycMin = MIN(NormalDepth, CriticalDepth);
                ycMax = MAX(NormalDepth, CriticalDepth);
                if ( y2 < ycMin ) flowClass = DN_CRITICAL;
                else if ( y2 < ycMax )
                {
                    if ( ycMax - ycMin < FUDGE ) Fasnh = 0.0;
                    else Fasnh = (ycMax - y2) / (ycMax - ycMin);
                }
            }
        }
    }

    // --- case where no flow at either end of conduit
    else if ( y1 <= FUDGE && y2 <= FUDGE ) flowClass = DRY;

    // --- case where downstream end of pipe is wet, upstream dry
    else if ( y2 > FUDGE )
    {
        // --- flow classification is UP_DRY if downstream head <
        //     invert of upstream end of conduit
        if ( h2 < Node[n1].invertElev + Link[j].z1 ) flowClass = UP_DRY;

        // --- otherwise, the downstream head will be >= upstream
        //     conduit invert creating a flow reversal and upstream end
        //     should be at critical depth, providing that an upstream
        //     offset exists (otherwise subcritical condition is maintained)
        else if ( z1 > 0.0 )
        {
            NormalDepth   = link_getYnorm(j, fabs(q));
            CriticalDepth = link_getYcrit(j, fabs(q));
            flowClass = UP_CRITICAL;
        }
    }

    // --- case where upstream end of pipe is wet, downstream dry
    else
    {
        // --- flow classification is DN_DRY if upstream head <
        //     invert of downstream end of conduit
        if ( h1 < Node[n2].invertElev + Link[j].z2 ) flowClass = DN_DRY;

        // --- otherwise flow at downstream end should be at critical depth
        //     providing that a downstream offset exists (otherwise
        //     subcritical condition is maintained)
        else if ( z2 > 0.0 )
        {
            NormalDepth = link_getYnorm(j, fabs(q));
            CriticalDepth = link_getYcrit(j, fabs(q));
            flowClass = DN_CRITICAL;
        }
    }
    return flowClass;
}

//=============================================================================

void findSurfArea(int j, float length, float* h1, float* h2,
                  float* y1, float* y2)
//
//  Input:   j  = conduit link index
//           q  = current conduit flow (cfs)
//           length = conduit length (ft)
//           h1 = head at upstream end of conduit (ft)
//           h2 = head at downstream end of conduit (ft)
//           y1 = upstream flow depth (ft)
//           y2 = downstream flow depth (ft)
//  Output:  updated values of h1, h2, y1, & y2;
//  Purpose: assigns surface area of conduit to its up and downstream nodes.
//
{
    int     n1, n2;                    // indexes of upstrm/downstrm nodes
    float   flowDepth1;                // flow depth at upstrm end (ft)
    float   flowDepth2;                // flow depth at downstrm end (ft)
    float   flowDepthMid;              // flow depth at midpt. (ft)
    float   width1;                    // top width at upstrm end (ft)
    float   width2;                    // top width at downstrm end (ft)
    float   widthMid;                  // top width at midpt. (ft)
    float   surfArea1 = 0.0;           // surface area at upstream node (ft2)
    float   surfArea2 = 0.0;           // surface area st downstrm node (ft2)
    TXsect* xsect = &Link[j].xsect;

    // --- get node indexes & current flow depths
    n1 = Link[j].node1;
    n2 = Link[j].node2;
    flowDepth1 = *y1;
    flowDepth2 = *y2;

    // --- add conduit's surface area to its end nodes depending on flow class
    switch ( Link[j].flowClass )
    {
      case SUBCRITICAL:
        flowDepthMid = 0.5 * (flowDepth1 + flowDepth2);
        if ( flowDepthMid < FUDGE ) flowDepthMid = FUDGE;
        width1 =   getWidth(xsect, flowDepth1);
        width2 =   getWidth(xsect, flowDepth2);
        widthMid = getWidth(xsect, flowDepthMid);
        surfArea1 = (width1 + widthMid) * length / 4.;
        surfArea2 = (widthMid + width2) * length / 4. * Fasnh;
        break;

      case UP_CRITICAL:
        flowDepth1 = CriticalDepth;
        if ( NormalDepth < CriticalDepth ) flowDepth1 = NormalDepth;
        *h1 = Node[n1].invertElev + Link[j].z1 + flowDepth1;
        flowDepthMid = 0.5 * (flowDepth1 + flowDepth2);
        if ( flowDepthMid < FUDGE ) flowDepthMid = FUDGE;
        width2   = getWidth(xsect, flowDepth2);
        widthMid = getWidth(xsect, flowDepthMid);
        surfArea2 = (widthMid + width2) * length * 0.5;
        break;

      case DN_CRITICAL:
        flowDepth2 = CriticalDepth;
        if ( NormalDepth < CriticalDepth ) flowDepth2 = NormalDepth;
        *h2 = Node[n2].invertElev + Link[j].z2 + flowDepth2;
        width1 = getWidth(xsect, flowDepth1);
        flowDepthMid = 0.5 * (flowDepth1 + flowDepth2);
        if ( flowDepthMid < FUDGE ) flowDepthMid = FUDGE;
        widthMid = getWidth(xsect, flowDepthMid);
        surfArea1 = (width1 + widthMid) * length * 0.5;
        break;

      case UP_DRY:
        flowDepth1 = FUDGE;
        flowDepthMid = 0.5 * (flowDepth1 + flowDepth2);
        if ( flowDepthMid < FUDGE ) flowDepthMid = FUDGE;
        width1 = getWidth(xsect, flowDepth1);
        width2 = getWidth(xsect, flowDepth2);
        widthMid = getWidth(xsect, flowDepthMid);

        // --- assign avg. surface area of downstream half of conduit
        //     to the downstream node
        surfArea2 = (widthMid + width2) * length / 4.;

        // --- if there is no free-fall at upstream end, assign the
        //     upstream node the avg. surface area of the upstream half
        if ( Link[j].z1 <= 0.0 )
        {
            surfArea1 = (width1 + widthMid) * length / 4.;
        }
        break;

      case DN_DRY:
        flowDepth2 = FUDGE;
        flowDepthMid = 0.5 * (flowDepth1 + flowDepth2);
        if ( flowDepthMid < FUDGE ) flowDepthMid = FUDGE;
        width1 = getWidth(xsect, flowDepth1);
        width2 = getWidth(xsect, flowDepth2);
        widthMid = getWidth(xsect, flowDepthMid);

        // --- assign avg. surface area of upstream half of conduit
        //     to the upstream node
        surfArea1 = (widthMid + width1) * length / 4.;

        // --- if there is no free-fall at downstream end, assign the
        //     downstream node the avg. surface area of the downstream half
        if ( Link[j].z2 <= 0.0 )
        {
            surfArea2 = (width2 + widthMid) * length / 4.;
        }
        break;

      case DRY:
        surfArea1 = FUDGE * length / 2.0;
        surfArea2 = surfArea1;
        break;
    }
    Xlink[j].surfArea1 = surfArea1;
    Xlink[j].surfArea2 = surfArea2;
    *y1 = flowDepth1;
    *y2 = flowDepth2;
}

//=============================================================================

float findLocalLosses(int j, float a1, float a2, float aMid, float q)
//
//  Input:   j    = link index
//           a1   = upstream area (ft2)
//           a2   = downstream area (ft2)
//           aMid = midpoint area (ft2)
//           q    = flow rate (cfs)
//  Output:  returns local losses (ft/sec)
//  Purpose: computes local losses term of momentum equation.
//
{
    float  losses = 0.0;
    q = fabs(q);

////////////////////////////////////////////////////
//// Area adjustment term added. (LR - 7/5/06 ) ////
////////////////////////////////////////////////////
    if ( a1 > FUDGE ) losses += Link[j].cLossInlet  * (q/a1) * aMid/a1;
    if ( a2 > FUDGE ) losses += Link[j].cLossOutlet * (q/a2) * aMid/a2;
    if ( aMid  > FUDGE ) losses += Link[j].cLossAvg * (q/aMid);

    return losses;
}

//=============================================================================

float getWidth(TXsect* xsect, float y)
//
//  Input:   xsect = ptr. to conduit cross section
//           y     = flow depth (ft)
//  Output:  returns top width (ft)
//  Purpose: computes top width of flow surface in conduit.
//
{
    float yNorm = y/xsect->yFull;
    if ( yNorm < 0.04 ) y = 0.04*xsect->yFull;
    if ( yNorm > 0.96 &&
         !xsect_isOpen(xsect->type) ) y = 0.96*xsect->yFull;
    return xsect_getWofY(xsect, y);
}

//=============================================================================

float getArea(TXsect* xsect, float y)
//
//  Input:   xsect = ptr. to conduit cross section
//           y     = flow depth (ft)
//  Output:  returns flow area (ft2)
//  Purpose: computes area of flow cross-section in a conduit.
//
{
    float area;                        // flow area (ft2)
    y = MIN(y, xsect->yFull);
    area = xsect_getAofY(xsect, y);
    area = MAX(area, FUDGE);
    return area;
}

//=============================================================================

float getHydRad(TXsect* xsect, float y)
//
//  Input:   xsect = ptr. to conduit cross section
//           y     = flow depth (ft)
//  Output:  returns hydraulic radius (ft)
//  Purpose: computes hydraulic radius of flow cross-section in a conduit.
//
{
    float hRadius;                     // hyd. radius (ft)
    y = MIN(y, xsect->yFull);
    hRadius = xsect_getRofY(xsect, y);
    hRadius = MAX(hRadius, FUDGE);
    return hRadius;
}

//=============================================================================

float checkFlapGate(int j, int n1, int n2, float q)
//
//  Input:   j = link index
//           n1 = index of upstream node
//           n2 = index of downstream node
//  Output:  returns flow in link (cfs)
//  Purpose: checks if flow in link should be zero due to closed flap gate.
//
{
    int n = -1;

    // --- return 0 if have reverse flow & link has a flap gate
    if ( Link[j].hasFlapGate )
    {
        if ( q * (float)Link[j].direction < 0.0 ) return 0.0;
    }
    
    // --- check for Outfall node on end of link where flow enters link
    if ( q < 0.0 ) n = n2;
    if ( q > 0.0 ) n = n1;
    if ( n >= 0 &&
         Node[n].type == OUTFALL &&
         Outfall[Node[n].subIndex].hasFlapGate ) return 0.0;
    return q;
}

//=============================================================================

float checkNormalFlow(int j, float q, float h1, float h2, float a1,
                      float a2, float r1)
{
    int   check = FALSE;
    int   k = Link[j].subIndex;
    int   n1 = Link[j].node1;
    int   n2 = Link[j].node2;
    float z1, z2, y1, y2, qNorm;
    float f1;
    float f2;

    // --- check for positive & non-surcharged flow
    if ( q <= 0.0 || h1 > Node[n1].crownElev ) return q;

//////////////////////////////////////////////////////////////////////////
//// Compute invert elevs. before checking conditions. (LR - 7/5/06 ) ////
//////////////////////////////////////////////////////////////////////////
    z1 = Node[n1].invertElev + Link[j].z1;
    z2 = Node[n2].invertElev + Link[j].z2;
    
    // --- check if water surface slope < conduit slope
    if ( NormalFlowLtd == FALSE || Node[n1].type == OUTFALL ||
         Node[n2].type == OUTFALL )
    {
        if ( h1 - h2 < z1 - z2 ) check = TRUE;
    }

    // --- check if Fr >= 1.0
    if ( NormalFlowLtd == TRUE && Node[n1].type != OUTFALL &&
         Node[n2].type != OUTFALL )
    {
        y1 = h1 - z1;
        y2 = h2 - z2;
        if ( y1 > FUDGE && y2 > FUDGE )
        {
            f1 = q / a1 / sqrt(GRAVITY * y1);
            f2 = q / a2 / sqrt(GRAVITY * y2);
            if ( f1 >= 1.0 ) check = TRUE;
            if ( f2 >= 1.0 ) check = TRUE;
        }
    }

    // --- check if normal flow < dynamic flow
    if ( check )
    {
        qNorm = Conduit[k].beta * a1 * pow(r1, 2./3.);
        return MIN(q, qNorm);
    }
    else return q;
}

//=============================================================================

void setNodeDepth(int i, float dt)
//
//  Input:   i  = node index
//           dt = time step (sec)
//  Output:  none
//  Purpose: sets depth at non-outfall node after current time step.
//
{
    int    canPond;                    // TRUE if node can pond overflows
    float  dQ;                         // inflow minus outflow at node (cfs)
    float  dV;                         // change in node volume (ft3)
    float  dy;                         // change in node depth (ft)
    float  yMax;                       // max. depth at node (ft)
    float  yOld;                       // node depth at previous time step (ft)
    float  yLast;                      // previous node depth (ft)
    float  yNew;                       // new node depth (ft)
    float  yCrown;                     // depth to node crown (ft)
    float  surfArea;                   // node surface area (ft2)
    float  denom;                      // denominator term
    float  corr;                       // correction factor
    float  f;                          // relative surcharge depth

    // --- see if node can pond water above it
    canPond = (AllowPonding && Node[i].pondedArea > 0.0);

    // --- initialize values
    yCrown = Node[i].crownElev - Node[i].invertElev;
    yOld = Node[i].oldDepth;
    yLast = Node[i].newDepth;
    Node[i].overflow = 0.0;
    surfArea = Xnode[i].newSurfArea;
    
    // --- determine average net flow volume into node over the time step
    dQ = Node[i].inflow - Node[i].outflow;
    dV = 0.5 * (Node[i].oldNetInflow + dQ) * dt;

    // --- find new volume at nodes with Type 1 pumps
    if ( Node[i].type != STORAGE && Node[i].fullVolume > 0.0 )
    {
        setWetWellVolume(i, dV, dt);
        return;
    }

    // --- if node not surcharged, base depth change on surface area        
    if ( yLast <= yCrown || Node[i].type == STORAGE || canPond )
    {
        dy = dV / surfArea;
        Xnode[i].oldSurfArea = Xnode[i].newSurfArea;
        yNew = yOld + dy;

        // --- apply under-relaxation to new depth estimate
        if ( Steps > 0 )
        {
            yNew = (1.0 - Omega) * yLast + Omega * yNew;
        }
    }

    // --- if node surcharged, base depth change on dqdh
    //     NOTE: depth change is w.r.t depth from previous
    //     iteration; also, do not apply under-relaxation.
    else
    {
        // --- apply correction factor for upstream terminal nodes
        corr = 1.0;
        if ( Node[i].degree < 0 ) corr = 0.6;

        // --- allow surface area from last non-surcharged condition
        //     to influence dqdh if depth close to crown depth
        denom = Xnode[i].sumdqdh;
        if ( yLast < 1.25 * yCrown )
        {
            f = (yLast - yCrown) / yCrown;
            denom += (Xnode[i].oldSurfArea/dt -
                      Xnode[i].sumdqdh) * exp(-15.0 * f);
        }

        // --- compute new estimate of node depth
        dy = corr * dQ / denom;
        yNew = yLast + dy;
        if ( yNew < yCrown ) yNew = yCrown - FUDGE;
    }

    // --- depth cannot be negative
    if ( yNew < 0 ) yNew = 0.0;

    // --- compute change in depth w.r.t. time
    Xnode[i].dYdT = fabs(yNew - yOld) / dt;

    // --- limit depth based on overflow or ponding
    yMax = Node[i].fullDepth;
    if ( canPond == FALSE ) yMax += Node[i].surDepth;
    if ( yNew > yMax )
    {
        yNew = MIN(yNew, getFloodedDepth(i, canPond, dV, yMax, dt));
    }

    // --- for depth less than max. depth compute new volume
    else Node[i].newVolume = node_getVolume(i, yNew);

    // --- save new depth for node
    Node[i].newDepth = yNew;
}

//=============================================================================

void setWetWellVolume(int i, float dV, float dt)
//
//  Input:   i  = node index
//           dV = change in volume over time step (ft3)
//           dt = time step (sec)
//  Output:  none
//  Purpose: computes new volume and depth at wet well node of a Type 1 pump.
//
{
    float  vNew;
    vNew = Node[i].oldVolume + dV;
    vNew = MAX(0.0, vNew);
    Node[i].overflow = (vNew - Node[i].fullVolume) / dt;
    Node[i].overflow = MAX(0.0, Node[i].overflow);
    vNew = MIN(vNew, Node[i].fullVolume);
    Node[i].newVolume = vNew;
    Node[i].newDepth = vNew / Node[i].fullVolume * Node[i].fullDepth;
    Xnode[i].dYdT = 0.0;
}

//=============================================================================

float getFloodedDepth(int i, int canPond, float dV, float yMax, float dt)
//
//  Input:   i  = node index
//           canPond = TRUE if water can pond over node
//           dV = change in volume over time step (ft3)
//           yMax = max. depth at node before ponding (ft)
//           dt = time step (sec)
//  Output:  returns depth at node when flooded (ft)
//  Purpose: computes new volume and depth at wet well node of a Type 1 pump.
//
{
    float  yPonded;                    // depth of ponded water (ft)

    // --- determine overflow lost from system
    if ( canPond == FALSE )
    {
        Node[i].overflow = (Node[i].oldVolume + dV - Node[i].fullVolume ) / dt;
        Node[i].overflow = MAX(0.0, Node[i].overflow);
        Node[i].newVolume = Node[i].fullVolume;
        return yMax;
    }

    // --- determine volume & depth of ponded water
    else
    {
        Node[i].newVolume = Node[i].oldVolume + dV;
        if ( Node[i].newVolume <= Node[i].fullVolume )
        {
            return yMax;
        }
        else
        {
            yPonded = Node[i].fullDepth +
                      (Node[i].newVolume - Node[i].fullVolume) /
                       Node[i].pondedArea;
            return yPonded;
        }
    }
}

//=============================================================================

float getVariableStep(float maxStep)
//
//  Input:   maxStep = user-supplied max. time step (sec)
//  Output:  returns time step (sec)
//  Purpose: finds time step that satisfies stability criterion but
//           is no greater than the user-supplied max. time step.
//
{
    int   minLink = -1;                // index of link w/ min. time step
    int   minNode = -1;                // index of node w/ min. time step
    float tMin;                        // allowable time step (sec)
    float tMinLink;                    // allowable time step for links (sec)
    float tMinNode;                    // allowable time step for nodes (sec)

    // --- find stable time step for links & then nodes
    tMin = maxStep;
    tMinLink = getLinkStep(tMin, &minLink);
    tMinNode = getNodeStep(tMinLink, &minNode);

    // --- use smaller of the link and node time step
    tMin = tMinLink;
    if ( tMinNode < tMin )
    {
        tMin = tMinNode ;
        minLink = -1;
    }

    // --- update count of times the minimum node or link was critical
    stats_updateCriticalTimeCount(minNode, minLink);

    // --- don't let time step go below an absolute minimum
    if ( tMin < MINTIMESTEP ) tMin = MINTIMESTEP;
    return tMin;
}

//=============================================================================

float getLinkStep(float tMin, int *minLink)
//
//  Input:   tMin = critical time step found so far (sec)
//  Output:  minLink = index of link with critical time step;
//           returns critical time step (sec)
//  Purpose: finds critical time step for conduits based on Courant criterion.
//
{
    int   i;                           // link index
    int   k;                           // conduit index
    float q;                           // conduit flow (cfs)
    float t;                           // time step (sec)
    float tLink = tMin;                // critical link time step (sec)

    // --- examine each conduit link
    for ( i = 0; i < Nobjects[LINK]; i++ )
    {
        if ( Link[i].type == CONDUIT )
        {
           // --- skip conduits with negligible flow, area or Fr
            k = Link[i].subIndex;
            q = fabs(Link[i].newFlow) / Conduit[k].barrels;
            if ( q <= 0.05 * Link[i].qFull
            ||   Conduit[k].a1 <= FUDGE
            ||   Link[i].froude <= 0.01 
               ) continue;

            // --- compute time step to satisfy Courant condition
            t = Link[i].newVolume / Conduit[k].barrels / q;
            t = t * Conduit[k].modLength / Conduit[k].length;
            t = t * Link[i].froude / (1.0 + Link[i].froude) * CourantFactor;

            // --- update critical link time step
            if ( t < tLink )
            {
                tLink = t;
                *minLink = i;
            }
        }
    }
    return tLink;
}

//=============================================================================

float getNodeStep(float tMin, int *minNode)
//
//  Input:   tMin = critical time step found so far (sec)
//  Output:  minNode = index of node with critical time step;
//           returns critical time step (sec)
//  Purpose: finds critical time step for nodes based on max. allowable
//           projected change in depth.
//
{
    int   i;                           // node index
    float maxDepth;                    // max. depth allowed at node (ft)
    float dYdT;                        // change in depth per unit time (ft/sec)
    float t1;                          // time needed to reach depth limit (sec)
    float tNode = tMin;                // critical node time step (sec)

    // --- find smallest time so that estimated change in nodal depth
    //     does not exceed safety factor * maxdepth
    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        // --- see if node can be skipped
        if ( Node[i].type == OUTFALL ) continue;
        if ( Node[i].newDepth <= FUDGE) continue;
        if ( Node[i].newDepth >
             Node[i].crownElev - Node[i].invertElev ) continue;

        // --- define max. allowable depth change using crown elevation
        maxDepth = (Node[i].crownElev - Node[i].invertElev) * 0.25;
        if ( maxDepth < FUDGE ) continue;
        dYdT = Xnode[i].dYdT;
        if (dYdT < FUDGE ) continue;

        // --- compute time to reach max. depth & compare with critical time
        t1 = maxDepth / dYdT;
        if ( t1 < tNode )
        {
            tNode = t1;
            *minNode = i;
        }
    }
    return tNode;
}

//=============================================================================
