//-----------------------------------------------------------------------------
//   flowrout.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             5/25/05  (Build 5.0.005a)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//   Author:   L. Rossman
//
//   Flow routing functions.
//-----------------------------------------------------------------------------

#include <stdlib.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  Constants
//-----------------------------------------------------------------------------
static const float  OMEGA   = 0.55;    // under-relaxation parameter
static const int    MAXITER = 10;      // max. iterations for storage updating
static const float  STOPTOL = 0.005;   // storage updating stopping tolerance

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  flowrout_init            (called by routing_open)
//  flowrout_close           (called by routing_close)
//  flowrout_getRoutingStep  (called routing_getRoutingStep)
//  flowrout_execute         (called routing_execute)

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void  initLinkDepths(void);
static void  initNodeDepths(void);
static void  initNodes(void);
static void  initLinks(void);
static void  validateTreeLayout(void);      
static void  validateGeneralLayout(void);
static void  updateStorageState(int i, int j, int links[], float dt);
static float getStorageOutflow(int node, int j, int links[], float dt);
static float getLinkInflow(int link, float dt);
static void  setNewNodeState(int node, float dt);
static void  setNewLinkState(int link);
static void  updateNodeDepth(int node, float y);
static int   steadyflow_execute(int link, float* qin, float* qout);


//=============================================================================

void flowrout_init(int routingModel)
//
//  Input:   routingModel = routing model code
//  Output:  none
//  Purpose: initializes flow routing system.
//
{
    // --- initialize for dynamic wave routing 
    if ( routingModel == DW )
    {
        // --- check for valid conveyance network layout
        validateGeneralLayout();
        dynwave_init();

        // --- initialize node & link depths if not using a hotstart file
        if ( Fhotstart1.mode == NO_FILE )
        {
            initNodeDepths();
            initLinkDepths();
        }
    }

    // --- validate network layout for kinematic wave routing
    else validateTreeLayout();

    // --- initialize node & link volumes
    initNodes();
    initLinks();
}

//=============================================================================

void  flowrout_close(int routingModel)
//
//  Input:   routingModel = routing method code
//  Output:  none
//  Purpose: closes down routing method used.
//
{
    if ( routingModel == DW ) dynwave_close();
}

//=============================================================================

float flowrout_getRoutingStep(int routingModel, float fixedStep)
//
//  Input:   routingModel = type of routing method used
//           fixedStep = user-assigned max. routing step (sec)
//  Output:  returns adjusted value of routing time step (sec)
//  Purpose: finds variable time step for dynamic wave routing.
//
{
    if ( routingModel == DW )
    {
        return dynwave_getRoutingStep(fixedStep);
    }
    return fixedStep;
}

//=============================================================================

int flowrout_execute(int links[], int routingModel, float tStep)
//
//  Input:   links = array of link indexes in topo-sorted order
//           routingModel = type of routing method used
//           tStep = routing time step (sec)
//  Output:  returns number of computational steps taken
//  Purpose: routes flow through conveyance network over current time step.
//
{
    int   i, j;
    int   n1;                          // upstream node of link
    float qin;                         // link inflow (cfs)
    float qout;                        // link outflow (cfs)
    float steps;                       // computational step count

    // --- set updated state of all nodes to False
    if ( ErrorCode ) return 0;
    for (j = 0; j < Nobjects[NODE]; j++) Node[j].updated = FALSE;

    // --- execute dynamic wave routing if called for
    if ( routingModel == DW )
    {
        steps = dynwave_execute(links, tStep);
        return steps;
    }

    // --- otherwise examine each link, moving from upstream to downstream
    steps = 0.0;
    for (i = 0; i < Nobjects[LINK]; i++)
    {
        // --- see if upstream node is a storage unit whose state needs updating
        j = links[i];
        n1 = Link[j].node1;
        if ( Node[n1].type == STORAGE ) updateStorageState(n1, i, links, tStep);

        // --- retrieve inflow at upstream end of link
        qin  = getLinkInflow(j, tStep);

        // route flow through link
        if ( routingModel == SF ) steps += steadyflow_execute(j, &qin, &qout);
        else steps += kinwave_execute(j, &qin, &qout, tStep);
        Link[j].newFlow = qout;

        // adjust outflow at upstream node and inflow at downstream node
        Node[ Link[j].node1 ].outflow += qin;
        Node[ Link[j].node2 ].inflow += qout;
    }
    if ( Nobjects[LINK] > 0 ) steps /= Nobjects[LINK];

    // --- update state of each non-updated node and link
    for ( j=0; j<Nobjects[NODE]; j++) setNewNodeState(j, tStep);
    for ( j=0; j<Nobjects[LINK]; j++) setNewLinkState(j);
    return (int)(steps+0.5);
}

//=============================================================================

void validateTreeLayout()
//
//  Input:   none
//  Output:  none
//  Purpose: validates tree-like conveyance system layout used for Steady
//           and Kinematic Wave flow routing
//
{
    int   j, node1, node2;
    float elev1, elev2;

    // --- check nodes
    for ( j = 0; j < Nobjects[NODE]; j++ )
    {
        switch ( Node[j].type )
        {
          // --- dividers must have only 2 outlet links
          case DIVIDER:
            if ( Node[j].degree > 2 )
            {
                report_writeErrorMsg(ERR_DIVIDER, Node[j].ID);
            }
            break;

          // --- outfalls cannot have any outlet links
          case OUTFALL:
            if ( Node[j].degree > 0 )
            {
                report_writeErrorMsg(ERR_OUTFALL, Node[j].ID);
            }
            break;

          // --- storage nodes can have multiple outlets
          case STORAGE: break;

          // --- all other nodes allowed only one outlet link
          default:
            if ( Node[j].degree > 1 )
            {
                report_writeErrorMsg(ERR_MULTI_OUTLET, Node[j].ID);
            }
        }
    }

    // ---  check links 
    for (j=0; j<Nobjects[LINK]; j++)
    {
        node1 = Link[j].node1;
        switch ( Link[j].type )
        {
          // --- conduits cannot have adverse slope
          case CONDUIT:
            node2 = Link[j].node2;
            elev1 = Link[j].z1 + Node[node1].invertElev;
            elev2 = Link[j].z2 + Node[node2].invertElev;
            if ( elev1 < elev2 )
            {
                report_writeErrorMsg(ERR_SLOPE, Link[j].ID);
            }
            break;

          // --- regulator links must be outlets of storage nodes
          case ORIFICE:
          case WEIR:
          case OUTLET:
            if ( Node[node1].type != STORAGE )
            {
                report_writeErrorMsg(ERR_REGULATOR, Link[j].ID);
            }
        }
    }
}

//=============================================================================

void validateGeneralLayout()
//
//  Input:   none
//  Output:  nonw
//  Purpose: validates general conveyance system layout.
//
{
    int i, j;
    int outletCount = 0;

    // --- use node inflow attribute to count inflow connections
    for ( i=0; i<Nobjects[NODE]; i++ ) Node[i].inflow = 0.0;

    // --- examine each link
    for ( j=0; j<Nobjects[LINK]; j++ )
    {
        // --- update inflow link count of downstream node
        i = Link[j].node1;
        if ( Node[i].type != OUTFALL ) i = Link[j].node2;
        Node[i].inflow += 1.0;

        // --- if link is dummy link then it must
        //     be the only link exiting the upstream node 

////////////////////////
//  LR - revised 5/24/05
////////////////////////
        if ( Link[j].type == CONDUIT && Link[j].xsect.type == DUMMY )
        {
            i = Link[j].node1;
            if ( Node[i].degree > 1 )
            {
                report_writeErrorMsg(ERR_MULTI_DUMMY_OUTLET, Node[i].ID);
            }
        }
    }

    // --- check each node to see if it qualifies as an outlet node
    //     (meaning that degree = 0)
    for ( i=0; i<Nobjects[NODE]; i++ )
    {
        // --- if node is of type Outfall, check that it has only 1
        //     connecting link (which can either be an outflow or inflow link)
        if ( Node[i].type == OUTFALL )
        {
            if ( Node[i].degree + (int)Node[i].inflow > 1 )
            {
                report_writeErrorMsg(ERR_OUTFALL, Node[i].ID);
            }
            else outletCount++;
        }

//////////////////////////////////////
// This section removed. (LR - 9/5/05)
//////////////////////////////////////
/*
        //  Check that interior node not mistaken for WQ outfall
        else if ( Nobjects[POLLUT] > 0 &&        // analyzing WQ
                  Node[i].degree == 0 &&         // has no outflow links
                 (int)Node[i].inflow > 1 )       // but has multiple inflow links
        {
            report_writeErrorMsg(ERR_OUTFALL, Node[i].ID);
        }
*/
    }
    if ( outletCount == 0 ) report_writeErrorMsg(ERR_NO_OUTLETS, "");

    // --- reset node inflows back to zero
    for ( i=0; i<Nobjects[NODE]; i++ )
    {
        if ( Node[i].inflow == 0.0 ) Node[i].degree = -Node[i].degree;
        Node[i].inflow = 0.0;
    }
}

//=============================================================================

void initNodeDepths(void)
//
//  Input:   none
//  Output:  none
//  Purpose: sets initial depth at nodes for Dyn. Wave flow routing.
//
{
    int   i;                           // link or node index
    int   n;                           // node index
    float y;                           // node water depth (ft)

    // --- use Node[].inflow as a temporary accumulator for depth in 
    //     connecting links and Node[].outflow as a temporary counter
    //     for the number of connecting links
    for (i=0; i<Nobjects[NODE]; i++)
    {
        Node[i].inflow  = 0.0;
        Node[i].outflow = 0.0;
    }

    // --- total up flow depths in all connecting links into nodes
    for (i=0; i<Nobjects[LINK]; i++)
    {
        if ( Link[i].newDepth > FUDGE ) y = Link[i].newDepth + Link[i].z1;
        else y = 0.0;
        n = Link[i].node1;
        Node[n].inflow += y;
        Node[n].outflow += 1.0;
        n = Link[i].node2;
        Node[n].inflow += y;
        Node[n].outflow += 1.0;
    }

    // --- if no user-supplied depth then set initial depth at non-storage/
    //     non-outfall nodes to average of depths in connecting links
    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        if ( Node[i].type == OUTFALL ) continue;
        if ( Node[i].type == STORAGE ) continue;
        if ( Node[i].initDepth > 0.0 ) continue;
        if ( Node[i].outflow > 0.0 )
        {
            Node[i].newDepth = Node[i].inflow / Node[i].outflow;
        }
    }

    // --- compute initial depths at all outfall nodes
    for ( i = 0; i < Nobjects[LINK]; i++ ) link_setOutfallDepth(i);
}

//=============================================================================
         
void initLinkDepths()
//
//  Input:   none
//  Output:  none
//  Purpose: sets initial flow depths in conduits under Dyn. Wave routing.
//
{
    int    i;                          // link index
    float  y, y1, y2;                  // depths (ft)

    // --- examine each link
    for (i=0; i<Nobjects[LINK]; i++)
    {
        // --- examine each conduit
        if ( Link[i].type == CONDUIT )
        {
            // --- skip conduits with user-assigned initial flows
            //     (their depths have already been set to normal depth)
            if ( Link[i].q0 != 0.0 ) continue;

            // --- set depth to average of depths at end nodes
            y1 = Node[Link[i].node1].newDepth - Link[i].z1;
            y1 = MAX(y1, 0.0);
            y1 = MIN(y1, Link[i].xsect.yFull);
            y2 = Node[Link[i].node2].newDepth - Link[i].z2;
            y2 = MAX(y2, 0.0);
            y2 = MIN(y2, Link[i].xsect.yFull);
            y = 0.5 * (y1 + y2);
            y = MAX(y, FUDGE);
            Link[i].newDepth = y;
        }
    }
}

//=============================================================================

void initNodes()
//
//  Input:   none
//  Output:  none
//  Purpose: sets initial inflow/outflow and volume for each node
//
{
    int i;

    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        // --- set default crown elevations here
        Node[i].crownElev = Node[i].invertElev;
        if ( Node[i].type == STORAGE )
        {
            Node[i].crownElev += Node[i].fullDepth;
        }

        // --- initialize node inflow, outflow, & volume
        Node[i].inflow = Node[i].newLatFlow;
        Node[i].outflow = 0.0;
        Node[i].newVolume = node_getVolume(i, Node[i].newDepth);
    }

    // --- update nodal inflow/outflow at ends of each link
    //     (needed for Steady Flow & Kin. Wave routing)
    for ( i = 0; i < Nobjects[LINK]; i++ )
    {
        if ( Link[i].newFlow >= 0.0 )
        {
            Node[Link[i].node1].outflow += Link[i].newFlow;
            Node[Link[i].node2].inflow  += Link[i].newFlow;
        }
        else
        {
            Node[Link[i].node1].inflow   -= Link[i].newFlow;
            Node[Link[i].node2].outflow  -= Link[i].newFlow;
        }
    }
}

//=============================================================================

void initLinks()
//
//  Input:   none
//  Output:  none
//  Purpose: sets initial upstream/downstream conditions in links.
//
//  Note: initNodes() must have been called first to properly
//        initialize each node's crown elevation.
//
{
    int    i;                          // link index
    int    j;                          // node index
    int    k;                          // conduit or pump index
    float  z;                          // crown elev. (ft)

    // --- examine each link
    for ( i = 0; i < Nobjects[LINK]; i++ )
    {
        // --- examine each conduit
        if ( Link[i].type == CONDUIT )
        {
            // --- assign initial flow to both ends of conduit
            k = Link[i].subIndex;
            Conduit[k].q1 = Link[i].newFlow / Conduit[k].barrels;
            Conduit[k].q2 = Conduit[k].q1;

            Conduit[k].q1Old = Conduit[k].q1;
            Conduit[k].q2Old = Conduit[k].q2;

            // --- find areas based on initial flow depth
            Conduit[k].a1 = xsect_getAofY(&Link[i].xsect, Link[i].newDepth);
            Conduit[k].a2 = Conduit[k].a1;

            // --- compute initial volume from area
            Link[i].newVolume = Conduit[k].a1 * Conduit[k].length *
                                Conduit[k].barrels;
            Link[i].oldVolume = Link[i].newVolume;

            // --- update crown elev. of nodes at either end
            j = Link[i].node1;
            z = Node[j].invertElev + Link[i].z1 + Link[i].xsect.yFull;
            Node[j].crownElev = MAX(Node[j].crownElev, z);
            j = Link[i].node2;
            z = Node[j].invertElev + Link[i].z2 + Link[i].xsect.yFull;
            Node[j].crownElev = MAX(Node[j].crownElev, z);
        }
    }
}

//=============================================================================

float getLinkInflow(int j, float dt)
//
//  Input:   j  = link index
//           dt = routing time step (sec)
//  Output:  returns link inflow (cfs)
//  Purpose: finds flow into upstream end of link at current time step under
//           Steady or Kin. Wave routing.
//
{
    int   n1 = Link[j].node1;
    float q;
    if ( Link[j].type == CONDUIT ||
         Link[j].type == PUMP ||
         Node[n1].type == STORAGE ) q = link_getInflow(j);
    else q = 0.0;
    return node_getMaxOutflow(n1, q, dt);
}

//=============================================================================

void updateStorageState(int i, int j, int links[], float dt)
//
//  Input:   i = index of storage node
//           j = current position in links array
//           links = array of topo-sorted link indexes
//           dt = routing time step (sec)
//  Output:  none
//  Purpose: updates depth and volume of a storage node using successive
//           approximation with under-relaxation for Steady or Kin. Wave
//           routing.
//
{
    int    iter;                       // iteration counter
    int    stopped;                    // TRUE when iterations stop
    float  vFixed;                     // fixed terms of flow balance eqn.
    float  v2;                         // new volume estimate (ft3)
    float  d1;                         // initial value of storage depth (ft)
    float  d2;                         // updated value of storage depth (ft)
    float  outflow;                    // outflow rate from storage (cfs)

    // --- see if storage node needs updating
    if ( Node[i].type != STORAGE ) return;
    if ( Node[i].updated ) return;

    // --- compute terms of flow balance eqn.
    //       v2 = v1 + (inflow - outflow)*dt
    //     that do not depend on storage depth at end of time step
    vFixed = Node[i].oldVolume + 
             0.5 * (Node[i].oldNetInflow + Node[i].inflow) * dt;
    d1 = Node[i].newDepth;

    // --- iterate finding outflow (which depends on depth) and subsequent
    //     new volume and depth until negligible depth change occurs
    iter = 1;
    stopped = FALSE;
    while ( iter < MAXITER && !stopped )
    {
        // --- find total flow in all outflow links
        outflow = getStorageOutflow(i, j, links, dt);

        // --- find new volume from flow balance eqn.
        v2 = vFixed - 0.5 * outflow * dt;

        // --- constrain volume to be between 0 and full value,
        //     while computing any overflow that might exist
        v2 = MAX(0.0, v2);
        Node[i].overflow = 0.0;
        if ( v2 > Node[i].fullVolume )
        {
            Node[i].overflow = (v2 - Node[i].fullVolume) / dt;
            v2 = Node[i].fullVolume;
        }

        // --- update node's volume and depth 
        Node[i].newVolume = v2;
        d2 = node_getDepth(i, v2);
        Node[i].newDepth = d2;

        // --- use under-relaxation to estimate new depth value
        //     and stop if close enough to previous value
        d2 = (1.0 - OMEGA)*d1 + OMEGA*d2;
        if ( fabs(d2 - d1) <= STOPTOL ) stopped = TRUE;

        // --- update old depth with new value and continue to iterate
        Node[i].newDepth = d2;
        d1 = d2;
        iter++;
    }

    // --- mark node as being updated
    Node[i].updated = TRUE;
}

//=============================================================================

float getStorageOutflow(int i, int j, int links[], float dt)
//
//  Input:   i = index of storage node
//           j = current position in links array
//           links = array of topo-sorted link indexes
//           dt = routing time step (sec)
//  Output:  returns total outflow from storage node (cfs)
//  Purpose: computes total flow released from a storage node.
//
{
    int   k, m;
    float outflow = 0.0;

    for (k = j; k < Nobjects[LINK]; k++)
    {
        m = links[k];
        if ( Link[m].node1 != i ) break;
        outflow += getLinkInflow(m, dt);
    }
    return outflow;        
}

//=============================================================================

void setNewNodeState(int j, float dt)
//
//  Input:   j  = node index
//           dt = time step (sec)
//  Output:  none
//  Purpose: updates state of node after current time step
//           for Steady Flow or Kinematic Wave flow routing.
//
{
    int   canPond;                     // TRUE if ponding can occur at node  
    float newNetInflow;                // inflow - outflow at node (cfs)

    // --- update stored volume using mid-point integration
    newNetInflow = Node[j].inflow - Node[j].outflow;
    Node[j].newVolume = Node[j].oldVolume +
                        0.5 * (Node[j].oldNetInflow + newNetInflow) * dt;
    if ( Node[j].newVolume < 0.0 ) Node[j].newVolume = 0.0;

    // --- determine any overflow lost from system
    Node[j].overflow = 0.0;
    canPond = (AllowPonding && Node[j].pondedArea > 0.0);
    if ( Node[j].newVolume > Node[j].fullVolume )
    {
        if ( !canPond )
        {
            Node[j].overflow = (Node[j].newVolume - Node[j].fullVolume) / dt;
            Node[j].newVolume = Node[j].fullVolume;

            // --- ignore any negligible overflow
            if ( Node[j].overflow <= FUDGE ) Node[j].overflow = 0.0;
        }
    }

    // --- compute a depth from volume
    //     (depths at upstream nodes are subsequently adjusted in
    //     setNewLinkState to reflect depths in connected conduit)
    if ( canPond )
    {
        Node[j].newDepth = node_getPondedDepth(j, Node[j].newVolume);
    }
    else
    {
        Node[j].newDepth = node_getDepth(j, Node[j].newVolume);
    }
}

//=============================================================================

void setNewLinkState(int j)
//
//  Input:   j = link index
//  Output:  none
//  Purpose: updates state of link after current time step under
//           Steady Flow or Kinematic Wave flow routing
//
{
    int   k;
    float a, y1, y2;

    Link[j].newDepth = 0.0;
    Link[j].newVolume = 0.0;

    if ( Link[j].type == CONDUIT )
    {
        // --- find avg. depth from entry/exit conditions
        k = Link[j].subIndex;
        a = 0.5 * (Conduit[k].a1 + Conduit[k].a2);   // avg. area
        Link[j].newVolume = a * Conduit[k].length * Conduit[k].barrels;
        y1 = xsect_getYofA(&Link[j].xsect, Conduit[k].a1);
        y2 = xsect_getYofA(&Link[j].xsect, Conduit[k].a2);
        Link[j].newDepth = 0.5 * (y1 + y2);

        // --- update depths at end nodes
        updateNodeDepth(Link[j].node1, y1 + Link[j].z1);
        updateNodeDepth(Link[j].node2, y2 + Link[j].z2);
    }
}

//=============================================================================

void updateNodeDepth(int i, float y)
//
//  Input:   i = node index
//           y = flow depth (ft)
//  Output:  none
//  Purpose: updates water depth at a node with a possibly higher value.
//
{
    // --- storage nodes were updated elsewhere
    if ( Node[i].type == STORAGE ) return;

    // --- if node is flooded, then use full depth
    if ( Node[i].overflow > 0.0 ) y = Node[i].fullDepth;

    // --- if current new depth below y
    if ( Node[i].newDepth < y )
    {
        // --- update new depth
        Node[i].newDepth = y;

        // --- depth cannot exceed full depth (if value exists)
        if ( Node[i].fullDepth > 0.0 && y > Node[i].fullDepth )
        {
            Node[i].newDepth = Node[i].fullDepth;
        }
    }
}

//=============================================================================

int steadyflow_execute(int j, float* qin, float* qout)
//
//  Input:   j = link index
//           qin = inflow to link (cfs)
//  Output:  qin = adjusted inflow to link (limited by flow capacity) (cfs)
//           qout = link's outflow (cfs)
//           returns 1 if successful
//  Purpose: performs steady flow routing through a single link.
//
{
    int   k;
    float s;
    float q;

    // --- use Manning eqn. to compute flow area for conduits
    if ( Link[j].type == CONDUIT )
    {
        k = Link[j].subIndex;
        q = (*qin) / Conduit[k].barrels;
        if ( Link[j].xsect.type == DUMMY ) Conduit[k].a1 = 0.0;
        else 
        {

//////////////////////////////////////////////////////////////////////////////
//  Comparison should be made against max. flow, not full flow. (LR - 3/10/06)
//////////////////////////////////////////////////////////////////////////////
//          if ( q > Link[j].qFull )
//          {
//              q = Link[j].qFull;
//              Conduit[k].a1 = Link[j].xsect.aFull;
//
            if ( q > Conduit[k].qMax )
            {
                q = Conduit[k].qMax;
                Conduit[k].a1 = xsect_getAmax(&Link[j].xsect);
                (*qin) = q * Conduit[k].barrels;
            }
            else
            {
                s = q / Conduit[k].beta;
                Conduit[k].a1 = xsect_getAofS(&Link[j].xsect, s);
            }
        }
        Conduit[k].a2 = Conduit[k].a1;
        Conduit[k].q1 = q;
        Conduit[k].q2 = q;
        (*qout) = q * Conduit[k].barrels;
    }
    else (*qout) = (*qin);
    return 1;
}

//=============================================================================
