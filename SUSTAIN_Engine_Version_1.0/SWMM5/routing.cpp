//-----------------------------------------------------------------------------
//   routing.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             10/19/05 (Build 5.0.006a)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman
//
//   Conveyance system routing functions.
//-----------------------------------------------------------------------------

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
// Shared variables
//-----------------------------------------------------------------------------
static int* SortedLinks;
static int  InSteadyState;

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
// routing_open            (called by swmm_start in swmm5.c)
// routing_getRoutingStep  (called by swmm_step in swmm5.c)
// routing_execute         (called by swmm_step in swmm5.c)
// routing_close           (called by swmm_end in swmm5.c)

//-----------------------------------------------------------------------------
// Function declarations
//-----------------------------------------------------------------------------
///////////////////////////////////////////////////////////////////////////
//  openHotstartFiles has been replaced by two new functions. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////////////
//static int  openHotstartFiles(void);

static int  openHotstartFile1(void);
static int  openHotstartFile2(void);

static void saveHotstartFile(void);
static void readHotstartFile(void);
static void addExternalInflows(DateTime currentDate);
static void addDryWeatherInflows(DateTime currentDate);
static void addWetWeatherInflows(double routingTime);
static void addGroundwaterInflows(double routingTime);
static void addRdiiInflows(DateTime currentDate);
static void addIfaceInflows(DateTime currentDate);
static void findEvap(float routingStep);
static void removeOutflows(void);
static void routeFlows(float routingStep);
static void routeQuality(float routingStep);
static int  systemHasChanged(int routingModel);

//=============================================================================

//////////////////////////////////////////////////////////////
//  This function was substantially rewritten. (LR - 10/19/05)
//////////////////////////////////////////////////////////////
int routing_open(int routingModel)
//
//  Input:   routingModel = routing method code
//  Output:  returns an error code
//  Purpose: initializes the routing analyzer.
//
{
    //  --- initialize steady state indicator
    InSteadyState = FALSE;

    // --- open treatment system
    if ( !treatmnt_open() ) return ErrorCode;

    // --- topologically sort the links
    SortedLinks = NULL;
    if ( Nobjects[LINK] > 0 )
    {
        SortedLinks = (int *) calloc(Nobjects[LINK], sizeof(int));
        if ( !SortedLinks )
        {
            report_writeErrorMsg(ERR_MEMORY, "");
            return ErrorCode;
        }
        toposort_sortLinks(SortedLinks);
        if ( ErrorCode ) return ErrorCode;
    }

    // --- open any routing interface files
    iface_openRoutingFiles();
    if ( ErrorCode ) return ErrorCode;

    // --- open hot start files
    if ( !openHotstartFile1() ) return ErrorCode;
    if ( !openHotstartFile2() ) return ErrorCode;

    // --- initialize the flow routing model
    flowrout_init(routingModel);
    return ErrorCode;
}

//=============================================================================

void routing_close(int routingModel)
//
//  Input:   routingModel = routing method code
//  Output:  none
//  Purpose: closes down the routing analyzer.
//
{
    // --- close hotstart file if in use
//////////////////////////////////////////////////
//  Following line no longer needed. (LR - 9/5/05)
//////////////////////////////////////////////////
    //if ( Fhotstart1.file ) fclose(Fhotstart1.file);

    if ( Fhotstart2.file )
    {
        // --- save latest results if called for
        if ( Fhotstart2.mode == SAVE_FILE ) saveHotstartFile();
        fclose(Fhotstart2.file);
    }

    // --- close any routing interface files
    iface_closeRoutingFiles();

    // --- free allocated memory
    flowrout_close(routingModel);
    treatmnt_close();
    FREE(SortedLinks);
}

//=============================================================================

float routing_getRoutingStep(int routingModel, float fixedStep)
//
//  Input:   routingModel = routing method code
//           fixedStep = user-supplied time step (sec)
//  Output:  returns a routing time step (sec)
//  Purpose: determines time step used for flow routing at current time period.
//
{
    if ( Nobjects[LINK] == 0 ) return fixedStep;
    else return flowrout_getRoutingStep(routingModel, fixedStep);
}

//=============================================================================

void routing_execute(int routingModel, float routingStep)
//
//  Input:   routingModel = routing method code
//           routingStep = routing time step (sec)
//  Output:  none
//  Purpose: executes the routing process at the current time period.
//
{
    int      j;
    int      stepCount = 1;
    int      actionCount;
    DateTime currentDate;

    // --- update continuity with current state
    //     applied over 1/2 of time step
    if ( ErrorCode ) return;
    massbal_updateRoutingTotals(routingStep/2.);

    // --- evaluate control rules at current date and elapsed time
    currentDate = getDateTime(NewRoutingTime);
    actionCount = controls_evaluate(currentDate, currentDate - StartDateTime,
                      (double)routingStep/SECperDAY);

    // --- update value of elapsed routing time (in milliseconds)
    OldRoutingTime = NewRoutingTime;
    NewRoutingTime = NewRoutingTime + 1000.0 * routingStep;
    currentDate = getDateTime(NewRoutingTime);

    // --- initialize mass balance totals for time step
    massbal_initTimeStepTotals();

    // --- replace old water quality state with new state
    if ( Nobjects[POLLUT] > 0 )
    {
        for (j=0; j<Nobjects[NODE]; j++) node_setOldQualState(j);
        for (j=0; j<Nobjects[LINK]; j++) link_setOldQualState(j);
    }

    // --- add lateral inflows to nodes
    for (j = 0; j < Nobjects[NODE]; j++)
    {
        Node[j].oldLatFlow  = Node[j].newLatFlow;
        Node[j].newLatFlow  = 0.0;
    }
    addExternalInflows(currentDate);
    addDryWeatherInflows(currentDate);
    addWetWeatherInflows(NewRoutingTime);
    addGroundwaterInflows(NewRoutingTime);
    addRdiiInflows(currentDate);
    addIfaceInflows(currentDate);

    // --- check if can skip steady state periods
    if ( SkipSteadyState )
    {
        if ( OldRoutingTime == 0.0
        ||   actionCount > 0
        ||   systemHasChanged(routingModel) ) InSteadyState = FALSE;
        else InSteadyState = TRUE;
    }

    // --- find new hydraulic state if system has changed
    if ( InSteadyState == FALSE )
    {
        // --- replace old hydraulic state values with current ones
        for (j = 0; j < Nobjects[LINK]; j++) link_setOldHydState(j);
        for (j = 0; j < Nobjects[NODE]; j++)
        {
            node_setOldHydState(j);
            node_initInflow(j, routingStep);
        }


        // --- route flow through the drainage network
        if ( Nobjects[LINK] > 0 )
        {
            stepCount = flowrout_execute(SortedLinks, routingModel, routingStep);
        }
    }

    // --- route quality through the drainage network
    if ( Nobjects[POLLUT] > 0 )
    {
        qualrout_execute(routingStep);
    }

    // --- remove evaporation & system outflows from nodes
    findEvap(routingStep);
    removeOutflows();

    // --- update continuity with new totals
    //     applied over 1/2 of routing step
    massbal_updateRoutingTotals(routingStep/2.);

    // --- update summary statistics
    if ( RptFlags.flowStats && Nobjects[LINK] > 0 )
    {
        stats_updateFlowStats(routingStep, currentDate, stepCount, InSteadyState);
    }
}

//=============================================================================

void addExternalInflows(DateTime currentDate)
//
//  Input:   currentDate = current date/time
//  Output:  none
//  Purpose: adds direct external inflows to nodes at current date.
//
{
    int    j, p;
    float  q, w;
    TExtInflow* inflow;

    // --- for each node with a defined external inflow
    for (j = 0; j < Nobjects[NODE]; j++)
    {
        inflow = Node[j].extInflow;
        if ( !inflow ) continue;

        // --- get flow inflow
        q = 0.0;
        while ( inflow )
        {
            if ( inflow->type == FLOW_INFLOW )
            {
                q = inflow_getExtInflow(inflow, currentDate);
                break;
            }
            else inflow = inflow->next;
        }
        if ( fabs(q) < FLOW_TOL ) q = 0.0;

        // --- add flow inflow to node's lateral inflow
        Node[j].newLatFlow += q;
        massbal_addInflowFlow(EXTERNAL_INFLOW, q);

        // --- get pollutant mass inflows
        inflow = Node[j].extInflow;
        while ( inflow )
        {
            if ( inflow->type != FLOW_INFLOW )
            {
                p = inflow->param;
                w = inflow_getExtInflow(inflow, currentDate);
                if ( inflow->type == CONCEN_INFLOW ) w *= q;
                Node[j].newQual[p] += w;
                massbal_addInflowQual(EXTERNAL_INFLOW, p, w);
            }
            inflow = inflow->next;
        }
    }
}

//=============================================================================

void addDryWeatherInflows(DateTime currentDate)
//
//  Input:   currentDate = current date/time
//  Output:  none
//  Purpose: adds dry weather inflows to nodes at current date.
//
{
    int      j, p;
    int      month, day, hour;
    float    q, w;
    TDwfInflow* inflow;

    // --- get month (zero-based), day-of-week (zero-based),
    //     & hour-of-day for routing date/time
    month = datetime_monthOfYear(currentDate) - 1;
    day   = datetime_dayOfWeek(currentDate) - 1;
    hour  = datetime_hourOfDay(currentDate);

    // --- for each node with a defined dry weather inflow
    for (j = 0; j < Nobjects[NODE]; j++)
    {
        inflow = Node[j].dwfInflow;
        if ( !inflow ) continue;

        // --- get flow inflow (i.e., the inflow whose param code is -1)
        q = 0.0;
        while ( inflow )
        {
            if ( inflow->param < 0 )
            {
                q = inflow_getDwfInflow(inflow, month, day, hour);
                break;
            }
            inflow = inflow->next;
        }
        if ( fabs(q) < FLOW_TOL ) q = 0.0;

        // --- add flow inflow to node's lateral inflow
        Node[j].newLatFlow += q;
        massbal_addInflowFlow(DRY_WEATHER_INFLOW, q);

        // --- get pollutant mass inflows
        inflow = Node[j].dwfInflow;
        while ( inflow )
        {
            if ( inflow->param >= 0 )
            {
                p = inflow->param;
                w = q * inflow_getDwfInflow(inflow, month, day, hour);
                Node[j].newQual[p] += w;
                massbal_addInflowQual(DRY_WEATHER_INFLOW, p, w);
            }
            inflow = inflow->next;
        }
    }
}

//=============================================================================

void addWetWeatherInflows(double routingTime)
//
//  Input:   routingTime = elasped time (millisec)
//  Output:  none
//  Purpose: adds runoff inflows to nodes at current elapsed time.
//
{
    int    i, j, p;
    float  q, w;
    double f;

    // --- find where current routing time lies between latest runoff times
    if ( Nobjects[SUBCATCH] == 0 ) return;
    f = (routingTime - OldRunoffTime) / (NewRunoffTime - OldRunoffTime);
    if ( f < 0.0 ) f = 0.0;
    if ( f > 1.0 ) f = 1.0;

    // for each subcatchment outlet node,
    // add interpolated runoff flow & pollutant load to node's inflow
    for (i = 0; i < Nobjects[SUBCATCH]; i++)
    {
        j = Subcatch[i].outNode;
        if ( j >= 0)
        {
            // add runoff flow to lateral inflow
            q = subcatch_getWtdOutflow(i, f);     // current runoff flow
            if ( fabs(q) < FLOW_TOL ) q = 0.0;
            Node[j].newLatFlow += q;
            massbal_addInflowFlow(WET_WEATHER_INFLOW, q);

            // add pollutant load
            for (p = 0; p < Nobjects[POLLUT]; p++)
            {
                w = q * subcatch_getWtdWashoff(i, p, f);
                Node[j].newQual[p] += w;
                massbal_addInflowQual(WET_WEATHER_INFLOW, p, w);
            }
        }
    }
}

//=============================================================================

void addGroundwaterInflows(double routingTime)
//
//  Input:   routingTime = elasped time (millisec)
//  Output:  none
//  Purpose: adds groundwater inflows to nodes at current elapsed time.
//
{
    int    i, j, p;
    float  q, w;
    double f;
    TGroundwater* gw;

    // --- find where current routing time lies between latest runoff times
    if ( Nobjects[SUBCATCH] == 0 ) return;
    f = (routingTime - OldRunoffTime) / (NewRunoffTime - OldRunoffTime);
    if ( f < 0.0 ) f = 0.0;
    if ( f > 1.0 ) f = 1.0;

    // --- for each subcatchment
    for (i = 0; i < Nobjects[SUBCATCH]; i++)
    {
        // --- see if subcatch contains groundwater
        gw = Subcatch[i].groundwater;
        if ( gw )
        {
            // --- identify node receiving groundwater flow
            j = gw->node;
            if ( j >= 0 )
            {
                // add groundwater flow to lateral inflow
                q = ( (1.0 - f)*(gw->oldFlow) + f*(gw->newFlow) )
                    * Subcatch[i].area;
                if ( fabs(q) < FLOW_TOL ) continue;
                Node[j].newLatFlow += q;
                massbal_addInflowFlow(GROUNDWATER_INFLOW, q);

                // add pollutant load (for positive inflow)
                if ( q > 0.0 )
                {
                    for (p = 0; p < Nobjects[POLLUT]; p++)
                    {
                        w = q * Pollut[p].gwConcen;
                        Node[j].newQual[p] += w;
                        massbal_addInflowQual(GROUNDWATER_INFLOW, p, w);
                    }
                }
            }
        }
    }
}

//=============================================================================

void addRdiiInflows(DateTime currentDate)
//
//  Input:   currentDate = current date/time
//  Output:  none
//  Purpose: adds RDII inflows to nodes at current date.
//
{
    int    i, j, p;
    float  q, w;
    int    numRdiiNodes;

    // --- see if any nodes have RDII at current date
    numRdiiNodes = rdii_getNumRdiiFlows(currentDate);

    // --- add RDII flow to each node's lateral inflow
    for (i=0; i<numRdiiNodes; i++)
    {
        rdii_getRdiiFlow(i, &j, &q);
        if ( j < 0 ) continue;
        if ( fabs(q) < FLOW_TOL ) continue;
        Node[j].newLatFlow += q;
        massbal_addInflowFlow(RDII_INFLOW, q);

        // add pollutant load (for positive inflow)
        if ( q > 0.0 )
        {
            for (p = 0; p < Nobjects[POLLUT]; p++)
            {
                w = q * Pollut[p].pptConcen;
                Node[j].newQual[p] += w;
                massbal_addInflowQual(RDII_INFLOW, p, w);
            }
        }
    }
}

//=============================================================================

void addIfaceInflows(DateTime currentDate)
//
//  Input:   currentDate = current date/time
//  Output:  none
//  Purpose: adds inflows from routing interface file to nodes at current date.
//
{
    int    i, j, p;
    float  q, w;
    int    numIfaceNodes;

    // --- see if any nodes have interface inflows at current date
    if ( Finflows.mode != USE_FILE ) return;
    numIfaceNodes = iface_getNumIfaceNodes(currentDate);

    // --- add interface flow to each node's lateral inflow
    for (i=0; i<numIfaceNodes; i++)
    {
        j = iface_getIfaceNode(i);
        if ( j < 0 ) continue;
        q = iface_getIfaceFlow(i);
        if ( fabs(q) < FLOW_TOL ) continue;
        Node[j].newLatFlow += q;
        massbal_addInflowFlow(EXTERNAL_INFLOW, q);

        // add pollutant load (for positive inflow)
        if ( q > 0.0 )
        {
            for (p = 0; p < Nobjects[POLLUT]; p++)
            {
                w = q * iface_getIfaceQual(i, p);
                Node[j].newQual[p] += w;
                massbal_addInflowQual(EXTERNAL_INFLOW, p, w);
            }
        }
    }
}

//=============================================================================

int  systemHasChanged(int routingModel)
//
//  Input:   none
//  Output:  returns TRUE if external inflows or hydraulics have changed
//           from the previous time step
//  Purpose: checks if the hydraulic state of the system has changed from
//           the previous time step.
//
{
    int   j, k;
    float diff;

    // --- check if external inflows have changed
    for (j=0; j<Nobjects[NODE]; j++)
    {
        diff = Node[j].oldLatFlow - Node[j].newLatFlow;
        if ( fabs(diff) > FLOW_TOL ) return TRUE;
    }

    // --- if system was already in steady state & there are no changes
    //     in inflows, then system must remain in steady state
    if ( InSteadyState ) return FALSE;

    // --- check for changes in node volume
    for (j=0; j<Nobjects[NODE]; j++)
    {
        diff = Node[j].newVolume - Node[j].oldVolume;
        if ( fabs(diff) > VOLUME_TOL ) return TRUE;
    }

    // --- check for other routing changes
    switch (routingModel)
    {
    // --- for dynamic wave routing, check if node depths have changed
    case DW:
        for (j=0; j<Nobjects[NODE]; j++)
        {
            diff = Node[j].oldDepth - Node[j].newDepth;
            if ( fabs(diff) > DEPTH_TOL ) return TRUE;
        }
        break;

    // --- for other routing methods, check if flows have changed
    case SF:
    case KW:
        for (j=0; j<Nobjects[LINK]; j++)
        {
            if ( Link[j].type == CONDUIT )
            {
                k = Link[j].subIndex;
                diff = Conduit[k].q1Old - Conduit[k].q1;
                if ( fabs(diff) > FLOW_TOL ) return TRUE;
                diff = Conduit[k].q2Old - Conduit[k].q2;
                if ( fabs(diff) > FLOW_TOL ) return TRUE;
            }
            else
            {
                diff = Link[j].oldFlow - Link[j].newFlow;
                if ( fabs(diff) > FLOW_TOL ) return TRUE;
            }
        }
        break;
    default: return TRUE;
    }
    return FALSE;
}

//=============================================================================

void findEvap(float routingStep)
//
//  Input:   routingStep = routing time step (sec)
//  Output:  none
//  Purpose: computes evaporation volume lost from nodes over current time step.
//
{
    int i;
    double evapLoss = 0.0;

    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        evapLoss += node_getEvapLoss(i, Evap.rate, routingStep);
    }
    massbal_addNodeEvap(evapLoss);
}

//=============================================================================

void removeOutflows()
//
//  Input:   none
//  Output:  none
//  Purpose: finds flows that leave the system and add these to mass
//           balance totals
//
{
    int    i, p;
    int    isFlooded;
    float  q, w;

    for ( i = 0; i < Nobjects[NODE]; i++ )
    {
        // --- determine flows leaving the system
        q = node_getSystemOutflow(i, &isFlooded);
        if ( q != 0.0 )
        {
            massbal_addOutflowFlow(q, isFlooded);
            for ( p = 0; p < Nobjects[POLLUT]; p++ )
            {
                w = q * Node[i].newQual[p];
                massbal_addOutflowQual(p, w, isFlooded);
            }
        }
    }
}

//=============================================================================

//////////////////////////////////////////////////////////////////
//  This function has been split into two functions. (LR - 9/5/05)
//////////////////////////////////////////////////////////////////
//int openHotstartFiles()

int openHotstartFile1()
//
//  Input:   none
//  Output:  none
//  Purpose: opens a previously saved hotstart file.
//
{
    INT4  nNodes;
    INT4  nLinks;
    INT4  nPollut;
    INT4  flowUnits;
    char  fileStamp[] = "SWMM5-HOTSTART";
    char  fStamp[] = "SWMM5-HOTSTART";

    // --- try to open the file
    if ( Fhotstart1.mode != USE_FILE ) return TRUE;
    if ( (Fhotstart1.file = fopen(Fhotstart1.name, "r+b")) == NULL)
    {
        report_writeErrorMsg(ERR_HOTSTART_FILE_OPEN, Fhotstart1.name);
        return FALSE;
    }

    // --- check that file contains proper header records
    fread(fStamp, sizeof(char), strlen(fileStamp), Fhotstart1.file);
    if ( strcmp(fStamp, fileStamp) != 0 )
    {
        report_writeErrorMsg(ERR_HOTSTART_FILE_FORMAT, "");
        return FALSE;
    }
    nNodes = -1;
    nLinks = -1;
    nPollut = -1;
    flowUnits = -1;
    fread(&nNodes, sizeof(INT4), 1, Fhotstart1.file);
    fread(&nLinks, sizeof(INT4), 1, Fhotstart1.file);
    fread(&nPollut, sizeof(INT4), 1, Fhotstart1.file);
    fread(&flowUnits, sizeof(INT4), 1, Fhotstart1.file);
    if ( nNodes != Nobjects[NODE]
    ||   nLinks != Nobjects[LINK]
    ||   nPollut   != Nobjects[POLLUT]
    ||   flowUnits != FlowUnits )
    {
         report_writeErrorMsg(ERR_HOTSTART_FILE_FORMAT, "");
         return FALSE;
    }

    // --- read contents of the file and close it
    readHotstartFile();
    fclose(Fhotstart1.file);
    if ( ErrorCode ) return FALSE;
    else return TRUE;
}

//=============================================================================

int openHotstartFile2()
//
//  Input:   none
//  Output:  none
//  Purpose: opens a new hotstart file to save results to.
//
{
    INT4  nNodes;
    INT4  nLinks;
    INT4  nPollut;
    INT4  flowUnits;
    char  fileStamp[] = "SWMM5-HOTSTART";

    // --- try to open file
    if ( Fhotstart2.mode != SAVE_FILE ) return TRUE;
    if ( (Fhotstart2.file = fopen(Fhotstart2.name, "w+b")) == NULL)
    {
        report_writeErrorMsg(ERR_HOTSTART_FILE_OPEN, Fhotstart2.name);
        return FALSE;
    }

    // --- write file stamp, # nodes, # links, & # pollutants to file
    nNodes = Nobjects[NODE];
    nLinks = Nobjects[LINK];
    nPollut = Nobjects[POLLUT];
    flowUnits = FlowUnits;
    fwrite(fileStamp, sizeof(char), strlen(fileStamp), Fhotstart2.file);
    fwrite(&nNodes, sizeof(INT4), 1, Fhotstart2.file);
    fwrite(&nLinks, sizeof(INT4), 1, Fhotstart2.file);
    fwrite(&nPollut, sizeof(INT4), 1, Fhotstart2.file);
    fwrite(&flowUnits, sizeof(INT4), 1, Fhotstart2.file);
    return TRUE;
}

//=============================================================================

void  saveHotstartFile(void)
//
//  Input:   none
//  Output:  none
//  Purpose: saves current state of all nodes and links to hotstart file.
//
{
    int   i, j;
    float zero = 0.0f;

    for (i = 0; i < Nobjects[NODE]; i++)
    {
        fwrite(&Node[i].newDepth,    sizeof(float), 1, Fhotstart2.file);
        fwrite(&Node[i].newLatFlow,  sizeof(float), 1, Fhotstart2.file);
        for (j = 0; j < Nobjects[POLLUT]; j++)
            fwrite(&Node[i].newQual[j], sizeof(float), 1, Fhotstart2.file);

        // --- write out 0 here for back compatibility
        for (j = 0; j < Nobjects[POLLUT]; j++ )
            fwrite(&zero, sizeof(float), 1, Fhotstart2.file);
    }
    for (i = 0; i < Nobjects[LINK]; i++)
    {
        fwrite(&Link[i].newFlow,   sizeof(float), 1, Fhotstart2.file);
        fwrite(&Link[i].newDepth,  sizeof(float), 1, Fhotstart2.file);
        fwrite(&Link[i].setting,   sizeof(float), 1, Fhotstart2.file);
        for (j = 0; j < Nobjects[POLLUT]; j++)
            fwrite(&Link[i].newQual[j], sizeof(float), 1, Fhotstart2.file);
    }
}

//=============================================================================

void readHotstartFile(void)
//
//  Input:   none
//  Output:  none
//  Purpose: reads initial state of all nodes and links from hotstart file.
//
{
    int   i, j;
    long  kount = 0;
    float zero;

    for (i = 0; i < Nobjects[NODE]; i++)
    {
        kount += fread(&Node[i].newDepth, sizeof(float), 1, Fhotstart1.file);
        kount += fread(&Node[i].newLatFlow, sizeof(float), 1, Fhotstart1.file);
        for (j = 0; j < Nobjects[POLLUT]; j++)
            kount += fread(&Node[i].newQual[j], sizeof(float), 1,
                           Fhotstart1.file);

        // --- read in zero here for back compatibility
        for (j = 0; j < Nobjects[POLLUT]; j++)
            kount += fread(&zero, sizeof(float), 1, Fhotstart1.file);
    }
    for (i = 0; i < Nobjects[LINK]; i++)
    {
        kount += fread(&Link[i].newFlow,   sizeof(float), 1, Fhotstart1.file);
        kount += fread(&Link[i].newDepth,  sizeof(float), 1, Fhotstart1.file);
        kount += fread(&Link[i].setting,   sizeof(float), 1, Fhotstart1.file);
        for (j = 0; j < Nobjects[POLLUT]; j++)
            kount += fread(&Link[i].newQual[j], sizeof(float), 1,
                           Fhotstart1.file);
    }
    if ( kount < Nobjects[NODE] * (2 + 2*Nobjects[POLLUT]) +
                 Nobjects[LINK] * (3 + Nobjects[POLLUT]) )
    {
         report_writeErrorMsg(ERR_HOTSTART_FILE_READ, "");
    }
}

//=============================================================================
