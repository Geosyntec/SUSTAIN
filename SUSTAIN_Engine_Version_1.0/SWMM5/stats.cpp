//-----------------------------------------------------------------------------
//   stats.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman (EPA)
//             R. Dickinson (CDM)
//
//   Simulation statistics functions.
//-----------------------------------------------------------------------------

#include <stdlib.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  Shared variables
//-----------------------------------------------------------------------------
#define MAX_STATS 5
static TSysStats       SysStats;
static TSubcatchStats* SubcatchStats;
static TNodeStats*     NodeStats;
static TLinkStats*     LinkStats;
static TMaxStats       MaxMassBalErrs[MAX_STATS];
static TMaxStats       MaxCourantCrit[MAX_STATS];

/////////////////////////////////////////////////////////
//  New array added for storage statistics. (LR - 9/5/05)
/////////////////////////////////////////////////////////
static TStorageStats*  StorageStats;

//////////////////////////////////////
//  New variables added. (LR - 7/5/06)
//////////////////////////////////////
static TOutfallStats*  OutfallStats;
static float           SysOutfallFlow;
static float           MaxOutfallFlow;
static float           MaxRunoffFlow;


//-----------------------------------------------------------------------------
//  Imported variables
//-----------------------------------------------------------------------------
extern double*        NodeInflow;      // defined in massbal.c
extern double*        NodeOutflow;     // defined in massbal.c

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  stats_open                    (called from swmm_start in swmm5.c)
//  stats_close                   (called from swmm_end in swmm5.c)
//  stats_report                  (called from swmm_end in swmm5.c)
//  stats_updateSubcatchStats     (called from subcatch_getRunoff)
//  stats_updateFlowStats         (called from routing_execute)
//  stats_updateCriticalTimeCount (called from getVariableStep in dynwave.c)

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void stats_updateNodeStats(int node, float tStep, DateTime aDate);
static void stats_updateLinkStats(int link, float tStep, DateTime aDate);
static void stats_findMaxStats(void);
static void stats_updateMaxStats(TMaxStats maxStats[], int i, int j, float x);

//=============================================================================

int  stats_open()
//
//  Input:   none
//  Output:  returns an error code
//  Purpose: opens the simulation statistics system.
//
{
    int j, k;

    // --- set all pointers to NULL
    NodeStats = NULL;
    LinkStats = NULL;

/////////////////////////////////////////////////////////
//  New array for storage statistics added. (LR - 9/5/05)
/////////////////////////////////////////////////////////
    StorageStats = NULL;
/////////////////////////////////////////////////////////
//  New array for outfall statistics added. (LR - 7/5/06)
/////////////////////////////////////////////////////////
    OutfallStats = NULL;

    // --- allocate memory for & initialize subcatchment statistics
    SubcatchStats = NULL;
    if ( Nobjects[SUBCATCH] > 0 )
    {
        SubcatchStats = (TSubcatchStats *) calloc(Nobjects[SUBCATCH],
                                               sizeof(TSubcatchStats));
        if ( !SubcatchStats )
        {
            report_writeErrorMsg(ERR_MEMORY, "");
            return ErrorCode;
        }
        for (j=0; j<Nobjects[SUBCATCH]; j++)
        {
            SubcatchStats[j].precip = 0.0;
            SubcatchStats[j].runon  = 0.0;
            SubcatchStats[j].evap   = 0.0;
            SubcatchStats[j].infil  = 0.0;
            SubcatchStats[j].runoff = 0.0;

////////////////////////////////////////////////////////
////  Max. runoff flow initialized. (LR - 3/10/06)  ////
////////////////////////////////////////////////////////
            SubcatchStats[j].maxFlow = 0.0;
        }
    }

    // --- allocate memory for node & link stats
    if ( Nobjects[LINK] > 0 )
    {
        NodeStats = (TNodeStats *) calloc(Nobjects[NODE], sizeof(TNodeStats));
        LinkStats = (TLinkStats *) calloc(Nobjects[LINK], sizeof(TLinkStats));
        if ( !NodeStats || !LinkStats )
        {
            report_writeErrorMsg(ERR_MEMORY, "");
            return ErrorCode;
        }
    }

    // --- initialize node stats
    if ( NodeStats ) for ( j = 0; j < Nobjects[NODE]; j++ )
    {
        NodeStats[j].avgDepth = 0.0;
        NodeStats[j].maxDepth = 0.0;
        NodeStats[j].maxDepthDate = StartDateTime;
        NodeStats[j].avgDepthChange = 0.0;
        NodeStats[j].volFlooded = 0.0;
        NodeStats[j].timeFlooded = 0.0;
        NodeStats[j].timeCourantCritical = 0.0;
        NodeStats[j].maxLatFlow = 0.0;                     //Added (LR - 7/5/06)
        NodeStats[j].maxInflow = 0.0;                      //Added (LR - 7/5/06)
        NodeStats[j].maxOverflow = 0.0;                    //Added (LR - 7/5/06)
        NodeStats[j].maxInflowDate = StartDateTime;        //Added (LR - 7/5/06)
        NodeStats[j].maxOverflowDate = StartDateTime;      //Added (LR - 7/5/06)
    }

    // --- initialize link stats
    if ( LinkStats ) for ( j = 0; j < Nobjects[LINK]; j++ )
    {
        LinkStats[j].maxFlow = 0.0;
        LinkStats[j].maxVeloc = 0.0;
        LinkStats[j].maxDepth = 0.0;                       //Added (LR - 7/5/06)
        LinkStats[j].avgFlowChange = 0.0;
        LinkStats[j].avgFroude = 0.0;
        LinkStats[j].timeSurcharged = 0.0;
        LinkStats[j].timeCourantCritical = 0.0;
        for (k=0; k<MAX_FLOW_CLASSES; k++)
            LinkStats[j].timeInFlowClass[k] = 0.0;
    }

/////////////////////////////////////////////////////////
//  New array for storage statistics added. (LR - 9/5/05)
//  Count of storage units corrected. (LR - 7/5/06)
/////////////////////////////////////////////////////////
    // --- allocate memory for & initialize storage unit statistics
    if ( Nnodes[STORAGE] > 0 )
    {
        StorageStats = (TStorageStats *) calloc(Nnodes[STORAGE],
                           sizeof(TStorageStats));
        if ( !StorageStats )
        {
            report_writeErrorMsg(ERR_MEMORY, "");
            return ErrorCode;
        }
        else for ( j = 0; j < Nnodes[STORAGE]; j++ )
        {
            StorageStats[j].avgVol = 0.0;
            StorageStats[j].maxVol = 0.0;
            StorageStats[j].maxFlow = 0.0;
            StorageStats[j].maxVolDate = StartDateTime;
        }
    }

/////////////////////////////////////////////////////////
//  New array for outfall statistics added. (LR - 7/5/06)
/////////////////////////////////////////////////////////
    // --- allocate memory for & initialize outfall statistics
    if ( Nnodes[OUTFALL] > 0 )
    {
        OutfallStats = (TOutfallStats *) calloc(Nnodes[OUTFALL],
                           sizeof(TOutfallStats));
        if ( !OutfallStats )
        {
            report_writeErrorMsg(ERR_MEMORY, "");
            return ErrorCode;
        }
        else for ( j = 0; j < Nnodes[OUTFALL]; j++ )
        {
            OutfallStats[j].avgFlow = 0.0;
            OutfallStats[j].maxFlow = 0.0;
            OutfallStats[j].totalPeriods = 0;
            if ( Nobjects[POLLUT] > 0 )
            {
                OutfallStats[j].totalLoad =
                    (float *) calloc(Nobjects[POLLUT], sizeof(float));
                if ( !OutfallStats[j].totalLoad )
                {
                    report_writeErrorMsg(ERR_MEMORY, "");
                    return ErrorCode;
                }
                for (k=0; k<Nobjects[POLLUT]; k++)
                    OutfallStats[j].totalLoad[k] = 0.0;
            }
            else OutfallStats[j].totalLoad = NULL;
        }
    }

    // --- initialize system stats
    MaxRunoffFlow = 0.0;                         //Added (LR - 7/5/06)
    MaxOutfallFlow = 0.0;                        //Added (LR - 7/5/06)
    SysStats.maxTimeStep = 0.0;
    SysStats.minTimeStep = (float)RouteStep;
    SysStats.avgTimeStep = 0.0;
    SysStats.avgStepCount = 0.0;
    SysStats.steadyStateCount = 0.0;
    return 0;
}

//=============================================================================

void  stats_close()
//
//  Input:   none
//  Output:  
//  Purpose: closes the simulation statistics system.
//
{
    int j;                                       //Added (LR - 7/5/06)

    FREE(SubcatchStats);
    FREE(NodeStats);
    FREE(LinkStats);

/////////////////////////////////////////////////////////
//  New array for storage statistics added. (LR - 9/5/05)
/////////////////////////////////////////////////////////
    FREE(StorageStats); 

/////////////////////////////////////////////////////////
//  New array for outfall statistics added. (LR - 9/5/05)
/////////////////////////////////////////////////////////
    if ( OutfallStats )
    {
        for ( j=0; j<Nnodes[OUTFALL]; j++ )
            FREE(OutfallStats[j].totalLoad);
        FREE(OutfallStats);
    }
}

//=============================================================================

void  stats_report()
//
//  Input:   none
//  Output:  none
//  Purpose: reports simulation statistics.
//
{
///////////////////////////////////////////////////////////////////////
//  Modified to report max. runoff flow & washoff loads. (LR - 7/5/06 )
///////////////////////////////////////////////////////////////////////
    if ( Nobjects[SUBCATCH] > 0 )
    {
        report_writeSubcatchStats(SubcatchStats, MaxRunoffFlow);
        if ( Nobjects[POLLUT] > 0 ) report_writeSubcatchLoads();
   }

    if ( Nobjects[LINK] > 0 )
    {
        report_writeNodeStats(NodeStats);

///////////////////////////////////////////////////////
// Reporting of storage statistics added. (LR - 9/5/05)
///////////////////////////////////////////////////////
        report_writeStorageStats(StorageStats);

///////////////////////////////////////////////////////
// Reporting of outfall statistics added. (LR - 7/5/06)
///////////////////////////////////////////////////////
       report_writeOutfallStats(OutfallStats, MaxOutfallFlow);

        report_writeLinkStats(LinkStats);
        stats_findMaxStats();
        report_writeMaxStats(MaxMassBalErrs, MaxCourantCrit, MAX_STATS);
        report_writeSysStats(&SysStats); 
    }
}

//=============================================================================

//////////////////////////////////////////////////////////
////  New argument added to function. (LR - 3/10/06)  ////
//////////////////////////////////////////////////////////
void   stats_updateSubcatchStats(int j, float rainVol, float runonVol,
           float evapVol, float infilVol, float runoffVol, float runoff)
//
//  Input:   j = subcatchment index
//           rainVol   = rainfall + snowfall volume (ft)
//           runonVol  = runon volume from other subcatchments (ft)
//           evapVol   = evaporation volume (ft)
//           infilVol  = infiltration volume (ft)
//           runoffVol = runoff volume (ft)
//           runoff    = runoff rate (cfs)
//  Output:  none
//  Purpose: updates totals of runoff components for a specific subcatchment.
//
{
    SubcatchStats[j].precip += rainVol;
    SubcatchStats[j].runon  += runonVol;
    SubcatchStats[j].evap   += evapVol;
    SubcatchStats[j].infil  += infilVol;
    SubcatchStats[j].runoff += runoffVol;

////////////////////////////////////////////////////////
////  Updating of max runoff added. (LR - 3/10/06)  ////
////////////////////////////////////////////////////////
    SubcatchStats[j].maxFlow = MAX(SubcatchStats[j].maxFlow, runoff);
}

//=============================================================================

/////////////////////////////////////
//  New function added. (LR - 7/5/06)
/////////////////////////////////////
void  stats_updateMaxRunoff()
//
//   Input:   none
//   Output:  updates global variable MaxRunoffFlow
//   Purpose: updates value of maximum system runoff rate.
//
{
    int j;
    float sysRunoff = 0.0;
    
    for (j=0; j<Nobjects[SUBCATCH]; j++) sysRunoff += Subcatch[j].newRunoff;
    MaxRunoffFlow = MAX(MaxRunoffFlow, sysRunoff);
}    

//=============================================================================

void   stats_updateFlowStats(float tStep, DateTime aDate, int stepCount,
       int steadyState)
//
//  Input:   tStep = routing time step (sec)
//           aDate = current date/time
//           stepCount = # steps required to solve routing at current time period
//           steadyState = TRUE if steady flow conditions exist
//  Output:  none
//  Purpose: updates various flow routing statistics at current time period.
//
{
    int   j;

/////////////////////////////////
//  New line added. (LR - 9/5/05)
//  New line added. (LR - 7/5/06)
/////////////////////////////////
    // --- update stats only after reporting period begins
    if ( aDate < ReportStart ) return;
    SysOutfallFlow = 0.0;

    // --- update node & link stats
    for ( j=0; j<Nobjects[NODE]; j++ )
        stats_updateNodeStats(j, tStep, aDate);
    for ( j=0; j<Nobjects[LINK]; j++ )
        stats_updateLinkStats(j, tStep, aDate);

    // --- update time step stats
    //     (skip initial time step for min. value)
    if ( StepCount > 1 )
    {
        SysStats.minTimeStep = MIN(SysStats.minTimeStep, tStep);
    }
    SysStats.avgTimeStep += tStep;
    SysStats.maxTimeStep = MAX(SysStats.maxTimeStep, tStep);

    // --- update iteration step count stats
    SysStats.avgStepCount += stepCount;

    // --- update count of times in steady state
    SysStats.steadyStateCount += steadyState;

/////////////////////////////////
//  New line added. (LR - 7/5/06)
/////////////////////////////////
    // --- update max. system outfall flow
    MaxOutfallFlow = MAX(MaxOutfallFlow, SysOutfallFlow);
}

//=============================================================================
   
void stats_updateCriticalTimeCount(int node, int link)
//
//  Input:   node = node index
//           link = link index
//  Output:  none
//  Purpose: updates count of times a node or link was time step-critical.
//
{
    if      ( node >= 0 ) NodeStats[node].timeCourantCritical += 1.0;
    else if ( link >= 0 ) LinkStats[link].timeCourantCritical += 1.0;
}

//=============================================================================

void stats_updateNodeStats(int j, float tStep, DateTime aDate)
//
//  Input:   j = node index
//           tStep = routing time step (sec)
//           aDate = current date/time
//  Output:  none
//  Purpose: updates flow statistics for a node.
//
{
/////////////////////////////////////
// New variables added. (LR - 9/5/05)
//                      (LR - 7/5/06)
/////////////////////////////////////
    int   k, p;
    float newVolume = Node[j].newVolume;
    float newDepth = Node[j].newDepth;

    NodeStats[j].avgDepth += newDepth;
    if ( newDepth > NodeStats[j].maxDepth )
    {
        NodeStats[j].maxDepth = newDepth;
        NodeStats[j].maxDepthDate = aDate;
    }
    NodeStats[j].avgDepthChange += fabs(newDepth - Node[j].oldDepth);
    if ( Node[j].type != OUTFALL
    &&   newDepth >= Node[j].fullDepth + Node[j].surDepth )
    {
        NodeStats[j].timeFlooded += tStep;
        NodeStats[j].volFlooded  += tStep * Node[j].overflow;
    }

////////////////////////////////////////////////////////////////
//  New code for updating storage node statistics. (LR - 9/5/05)
////////////////////////////////////////////////////////////////
    // --- update storage statistics
    if ( Node[j].type == STORAGE )
    {
        k = Node[j].subIndex;
        StorageStats[k].avgVol += newVolume;
        if ( newVolume > StorageStats[k].maxVol )
        {
            StorageStats[k].maxVol = newVolume;
            StorageStats[k].maxVolDate = aDate;
        }
        StorageStats[k].maxFlow = MAX(StorageStats[k].maxFlow, Node[j].outflow);
    }

////////////////////////////////////////////////////////////////
//  New code for updating outfall node statistics. (LR - 7/5/06)
////////////////////////////////////////////////////////////////
    // --- update outfall statistics
    if ( Node[j].type == OUTFALL && Node[j].inflow >= MIN_RUNOFF_FLOW )
    {
        k = Node[j].subIndex;
        OutfallStats[k].avgFlow += Node[j].inflow;
        OutfallStats[k].maxFlow = MAX(OutfallStats[k].maxFlow, Node[j].inflow);
        OutfallStats[k].totalPeriods++;
        for (p=0; p<Nobjects[POLLUT]; p++)
        {
            OutfallStats[k].totalLoad[p] += Node[j].inflow *
                Node[j].newQual[p] * LperFT3 * tStep * Pollut[p].mcf;
        }
        SysOutfallFlow += Node[j].inflow;
    }

/////////////////////////////////////////////////////////////
//  New code for updating node flow statistics. (LR - 7/5/06)
/////////////////////////////////////////////////////////////
    NodeStats[j].maxLatFlow = MAX(Node[j].newLatFlow, NodeStats[j].maxLatFlow);
    if ( Node[j].inflow > NodeStats[j].maxInflow )
    {
        NodeStats[j].maxInflow = Node[j].inflow;
        NodeStats[j].maxInflowDate = aDate;
    }
    if ( Node[j].overflow > NodeStats[j].maxOverflow )
    {
        NodeStats[j].maxOverflow = Node[j].overflow;
        NodeStats[j].maxOverflowDate = aDate;
    }
}

//=============================================================================

void  stats_updateLinkStats(int j, float tStep, DateTime aDate)
//
//  Input:   j = link index
//           tStep = routing time step (sec)
//           aDate = current date/time
//  Output:  none
//  Purpose: updates flow statistics for a link.
//
{
    int   k;
    float q, v;

    //if ( Link[j].type == CONDUIT )                       //Removed (LR - 3/10/06)
    //{                                                    //Removed (LR - 3/10/06)
        // --- update max. flow
        q = fabs(Link[j].newFlow);
        if ( q > LinkStats[j].maxFlow )
        {
            LinkStats[j].maxFlow = q;
            LinkStats[j].maxFlowDate = aDate;
        }

        // --- update max. velocity
        v = link_getVelocity(j, q, Link[j].newDepth);
        if ( v > LinkStats[j].maxVeloc )
        {
            LinkStats[j].maxVeloc = v;
            LinkStats[j].maxVelocDate = aDate;
        }

        // --- update max. depth                           //Added (LR - 7/5/06)
        if ( Link[j].newDepth > LinkStats[j].maxDepth )    //Added (LR - 7/5/06)
        {                                                  //Added (LR - 7/5/06)
            LinkStats[j].maxDepth = Link[j].newDepth;      //Added (LR - 7/5/06)
        }                                                  //Added (LR - 7/5/06)

    if ( Link[j].type == PUMP )                            //Added (LR - 3/10/06)
    {                                                      //Added (LR - 3/10/06)
        if ( q >= Link[j].qFull )                          //Added (LR - 3/10/06)
            LinkStats[j].timeSurcharged += tStep;          //Added (LR - 3/10/06)
    }                                                      //Added (LR - 3/10/06)
    else if ( Link[j].type == CONDUIT )                    //Added (LR - 3/10/06)
    {                                                      //Added (LR - 3/10/06)
        // --- update sums used to compute avg. Fr and flow change
        LinkStats[j].avgFroude += Link[j].froude; 
        LinkStats[j].avgFlowChange += fabs(Link[j].newFlow - Link[j].oldFlow);
    
        // --- update flow classification distribution
        k = Link[j].flowClass;
        if ( k >= 0 && k < MAX_FLOW_CLASSES )
        {
            ++LinkStats[j].timeInFlowClass[k];
        }

        // --- update time conduit is surcharged
        k = Link[j].subIndex;
        if ( q >= Link[j].qFull * (float)Conduit[k].barrels ||
             Link[j].newDepth >= Link[j].xsect.yFull)
        {
            LinkStats[j].timeSurcharged += tStep;
        }
    }
}

//=============================================================================

void  stats_findMaxStats()
//
//  Input:   none
//  Output:  none
//  Purpose: finds nodes & links with highest mass balance errors
//           & highest times Courant time-step critical.
//
{
    int   j;
    float x;

    // --- initialize max. stats arrays
    for (j=0; j<MAX_STATS; j++)
    {
        MaxMassBalErrs[j].objType = NODE;
        MaxMassBalErrs[j].index   = -1;
        MaxMassBalErrs[j].value   = 0.0;
        MaxCourantCrit[j].index   = -1;
        MaxCourantCrit[j].value   = 0.0;
    }

    // --- find nodes with largest mass balance errors
    for (j=0; j<Nobjects[NODE]; j++)
    {
        // --- skip terminal nodes and nodes with negligible inflow
        if ( Node[j].degree <= 0  ) continue;
        if ( NodeInflow[j] <= 0.1 ) continue;

        // --- evaluate mass balance error
        if      ( NodeInflow[j]  > 0.0 ) x = 1.0 - NodeOutflow[j] / NodeInflow[j];
        else if ( NodeOutflow[j] > 0.0 ) x = -1.0;
        else                             x = 0.0;
        stats_updateMaxStats(MaxMassBalErrs, NODE, j, 100.0*x);
    }

    // --- stop if not using a variable time step
    if ( RouteModel != DW || CourantFactor == 0.0 ) return;

    // --- find nodes most frequently Courant critical
    for (j=0; j<Nobjects[NODE]; j++)
    {
        x = NodeStats[j].timeCourantCritical / StepCount;
        stats_updateMaxStats(MaxCourantCrit, NODE, j, 100.0*x);
    }

    // --- find links most frequently Courant critical
    for (j=0; j<Nobjects[LINK]; j++)
    {
        x = LinkStats[j].timeCourantCritical / StepCount;
        stats_updateMaxStats(MaxCourantCrit, LINK, j, 100.0*x);
    }
}

//=============================================================================

void  stats_updateMaxStats(TMaxStats maxStats[], int i, int j, float x)
//
//  Input:   maxStats[] = array of critical statistics values
//           i = object category (NODE or LINK)
//           j = object index
//           x = value of statistic for the object
//  Output:  none
//  Purpose: updates the collection of most critical statistics
//
{
    int   k;
    TMaxStats maxStats1, maxStats2;
    maxStats1.objType = i;
    maxStats1.index   = j;
    maxStats1.value   = x;
    for (k=0; k<MAX_STATS; k++)
    {
        if ( fabs(maxStats1.value) > fabs(maxStats[k].value) )
        {
            maxStats2 = maxStats[k];
            maxStats[k] = maxStats1;
            maxStats1 = maxStats2;
        }
    }
}

//=============================================================================
