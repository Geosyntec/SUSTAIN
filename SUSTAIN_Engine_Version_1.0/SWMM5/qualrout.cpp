//-----------------------------------------------------------------------------
//   qualrout.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//   Author:   L. Rossman
//
//   Water quality routing functions.
//-----------------------------------------------------------------------------

#include <stdio.h>
#include <stdlib.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  qualrout_execute  (called by routing_execute)

//-----------------------------------------------------------------------------
//  Function declarations
//-----------------------------------------------------------------------------
/* moved to func.h
static void  findLinkMassFlow(int i);
static void  findNodeQual(int j);
static void  findLinkQual(int i, float tStep);
static void  findStorageQual(int j, float tStep);
static void  updateHRT(int j, float v, float q, float tStep);
*/
//=============================================================================

void qualrout_execute(float tStep)
//
//  Input:   tStep = routing time step (sec)
//  Output:  none
//  Purpose: routes water quality constituents through the drainage
//           network over the current time step.
//
{
    int   i, j;

    // --- find mass flow from each link to its downstream node
    for ( i=0; i<Nobjects[LINK]; i++ ) findLinkMassFlow(i);

    // --- find new water quality concentration at each node  
    for (j = 0; j < Nobjects[NODE]; j++)
    {

//////////////////////////////////////////////////////
// New code added to save inflow concen. (LR - 9/5/05)
//////////////////////////////////////////////////////
        // --- save inflow concentrations if treatment applied
        if ( Node[j].treatment )
        {
            treatmnt_setInflow(j, Node[j].inflow, Node[j].newQual);
        }
       
        // --- use a mass balance eqn. for storage nodes
        //     or nodes with ponded volume
        if ( Node[j].type == STORAGE ||
             Node[j].newVolume > TINY ) findStorageQual(j, tStep);

        // --- otherwise base node quality on mass inflow divided by
        //     flow inflow
        else findNodeQual(j);

        // --- apply treatment to new quality values
        if ( Node[j].treatment )
            treatmnt_treat(j, Node[j].inflow, Node[j].newVolume, tStep);
    }

    // --- find new water quality in each link
    for ( i=0; i<Nobjects[LINK]; i++ ) findLinkQual(i, tStep);
}

//=============================================================================

void findLinkMassFlow(int i)
//
//  Input:   i = link index
//  Output:  none
//  Purpose: adds constituent mass flow out of link to the total
//           accumulation at the link's downstream node.
//
{
    int   j, k, p;
    float qLink;

    // --- identify index of downstream node
    j = Link[i].node2;
    if ( Link[i].newFlow < 0.0 ) j = Link[i].node1;

    // --- find inflow to downstream node
    //     (for conduits, use Conduit[k].q2 since for
    //     KW routing, inlet flow can differ from outlet flow)
    qLink = Link[i].newFlow;

///////////////////////////////////////////////////////////////////////////////
//  Test not needed since Link.newFlow now is same as Conduit.q2. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////////////////
//    if ( Link[i].type == CONDUIT )
//    {
//        k = Link[i].subIndex;
//        qLink = Conduit[k].q2 * (float)Conduit[k].barrels;
//    }

    // --- add mass inflow from link to total at downstream node
    //     (contributions from lateral inflows already computed)
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        Node[j].newQual[p] += fabs(qLink) * Link[i].oldQual[p];
    }
}

//=============================================================================

void findNodeQual(int j)
//
//  Input:   j = node index
//  Output:  none
//  Purpose: finds new quality in a node with no storage volume.
//
{
    int   p;
    float qNode;

    // --- if there is flow at node then concen. is mass inflow
    //     divided by node flow
    qNode = Node[j].inflow;
    if ( qNode > TINY )
    {
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            Node[j].newQual[p] /= qNode;
        }
    }

    // --- if node has no water, then concen. is 0
    else if (Node[j].newDepth < TINY )
    {
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            Node[j].newQual[p] = 0.0;
        }
    }

    // --- otherwise concen. is same as old value
    else
    { 
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            Node[j].newQual[p] = Node[j].oldQual[p];
        }
    }
}

//=============================================================================

void findLinkQual(int i, float tStep)
//
//  Input:   i = link index
//           tStep = routing time step (sec)
//  Output:  none
//  Purpose: finds new quality in a link after the current time step.
//
{
    int   j, k, p;
    float kDecay, massReacted;
    float q, vOld, vPlus;
    float c, cMax, wAdded;

    // --- find flow at inlet of link
    //     (for conduits, use Conduit[k].q1 since for
    //     KW routing, inlet flow can differ from outlet flow)
    q = Link[i].newFlow;
    if ( Link[i].type == CONDUIT )
    {
        k = Link[i].subIndex;
        q = Conduit[k].q1 * (float)Conduit[k].barrels;
    }

    // --- find old volume (vOld) & volume after inflow is added (vPlus)
    vOld = Link[i].oldVolume;
    vPlus = vOld + (fabs(q) * tStep);

    // --- identify index of upstream node
    j = Link[i].node1;
    if ( q < 0.0 ) j = Link[i].node2;

    // --- for non-conduit links (i.e., those with no length or volume),
    //     concentrations equal those of upstream node

/////////////////////////////////////////////////////////
//  Test updated to include Dummy conduits. (LR - 9/5/05)
/////////////////////////////////////////////////////////
    if ( Link[i].type != CONDUIT || Link[i].xsect.type == DUMMY)
    {
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            Link[i].newQual[p] = Node[j].newQual[p];
        }
        return;
    }

    // --- for each pollutant
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        // --- find exponential 1st order decay over time step
        c = Link[i].oldQual[p];
        kDecay = Pollut[p].kDecay;
        if ( kDecay != 0.0 )
        {
            c = c * exp(-kDecay * tStep);
            massReacted = (Link[i].oldQual[p] - c) * vOld / tStep;
            massbal_addReactedMass(p, massReacted);
        }

        // --- compute upper bound on mixture concentration
        cMax = MAX(c, Node[j].newQual[p]);

        // --- compute mass added over time step
        wAdded = Node[j].newQual[p] * fabs(q) * tStep;

        // --- combine inflow with old volume to compute new concen.
        if ( vPlus <= TINY ) c = 0.0;
        else c = (c*vOld + wAdded) / vPlus;
        c = MIN(c, cMax);
        c = MAX(c, 0.0);
        Link[i].newQual[p] = c;
    }
}

//=============================================================================

void  findStorageQual(int j, float tStep)
//
//  Input:   j = node index
//           tStep = routing time step (sec)
//  Output:  none
//  Purpose: finds new quality in a node with storage volume.
//
{
    int   p;
    float kDecay, massReacted;
    float qIn, vOld, vPlus;
    float wAdded, c, cMax;

    // --- find old volume (vOld) & volume after inflow is added (vPlus)
    vOld = Node[j].oldVolume;
    qIn = Node[j].inflow;
    vPlus = vOld + qIn * tStep;

    // --- update hyd. residence time for storage nodes
    //     (HRT can be used in treatment functions)
    if ( Node[j].type == STORAGE ) updateHRT(j, vOld, Node[j].inflow, tStep);

    // --- for each pollutant
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        // --- find 1st order exponential decay over time step
        c = Node[j].oldQual[p];

//////////////////////////////////////////////////////////////////////////
//  Only apply decay if no separate treatment function used. (LR - 9/5/05)
//////////////////////////////////////////////////////////////////////////
        if ( Node[j].treatment == NULL ||
             Node[j].treatment[p].equation == NULL )
        {
            kDecay = Pollut[p].kDecay;
            if ( kDecay != 0.0 )
            {
                c = c * exp(-kDecay * tStep);
                massReacted = (Node[j].oldQual[p] - c) * vOld / tStep;
                massbal_addReactedMass(p, massReacted);
            }
        }

        // --- compute mass added over time step
        //     (total mass inflow rate was previously accumulated
        //      in Node[j].newQual[p])
        wAdded = Node[j].newQual[p] * tStep;

        // --- compute upper bound on mixture concentration
        cMax = c;
        if ( qIn > TINY )
        {
            cMax = MAX(cMax, (Node[j].newQual[p] / qIn) );
        }

        // --- if new volume negligible, then concen. = 0
        if ( Node[j].newVolume <= TINY ) c = 0.0;

        // --- otherwise combine inflow with old volume to compute new concen.
        //     (note that if vPlus <= TINY then there's no change in c)
        else if ( vPlus > TINY ) c = (c*vOld + wAdded) / vPlus;
        c = MIN(c, cMax);
        c = MAX(c, 0.0);
        Node[j].newQual[p] = c;
    }
}

//=============================================================================

void updateHRT(int j, float v, float q, float tStep)
//
//  Input:   j = node index
//           v = storage volume (ft3)
//           q = inflow rate (cfs)
//           tStep = time step (sec)
//  Output:  none
//  Purpose: updates hydraulic residence time (i.e., water age) at a 
//           storage node.
//
{
    int   k = Node[j].subIndex;
    float hrt = Storage[k].hrt;
    if ( v == 0.0 ) Storage[k].hrt = 0.0;
    else
    {
        hrt = hrt * (1.0 - q * tStep / v) + tStep;
        Storage[k].hrt = MAX(hrt, 0.0);
    }
}
