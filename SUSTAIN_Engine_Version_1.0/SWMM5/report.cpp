//-----------------------------------------------------------------------------
//   report.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/05   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman
//
//   Report writing functions.
//-----------------------------------------------------------------------------

#include <malloc.h>
#include <string.h>
#include <math.h>
#include <time.h>
#include "headers.h"

#define WRITE(x) (report_writeLine((x)))

#define LINE_10 "----------"

//////////////////////////////////////////////////////
////  Actually contains 51 '-'s, not 41. (LR - 7/5/06)
//////////////////////////////////////////////////////
#define LINE_41 \
"---------------------------------------------------"

#define LINE_61 \
"-------------------------------------------------------------"

//-----------------------------------------------------------------------------
//  Shared variables   
//-----------------------------------------------------------------------------
static char* LoadUnitsWords[] = { w_LBS, w_KG, w_LOGN };
static char* NodeTypeWords[]  = { w_JUNCTION, w_OUTFALL, w_STORAGE, w_DIVIDER };
static char* LinkTypeWords[]  = { w_CONDUIT, w_PUMP, w_ORIFICE, w_WEIR, w_OUTLET };
static time_t SysTime;

//-----------------------------------------------------------------------------
//  Imported variables
//-----------------------------------------------------------------------------
extern float* SubcatchResults;         // Results vectors defined in OUTPUT.C
extern float* NodeResults;             //  "
extern float* LinkResults;             //  "
extern char   ErrString[81];           // defined in ERROR.C

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void report_Options(void);
static void report_Subcatchments(void);
static void report_SubcatchHeader(char *id);
static void report_Nodes(void);
static void report_NodeHeader(char *id);
static void report_Links(void);
static void report_LinkHeader(char *id);
static void report_LoadingErrors(int p1, int p2, TLoadingTotals* totals);
static void report_QualErrors(int p1, int p2, TRoutingTotals* totals);
/////////////////////////////////////
//  New function added. (LR - 7/5/06)
/////////////////////////////////////
static void writeNodeFlowStats(TNodeStats nodeStats[]);


//=============================================================================

int report_readOptions(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns an error code
//  Purpose: reads reporting options from a line of input
//
{
    char  k;
    int   j, m, t;
    if ( ntoks < 2 ) return error_setInpError(ERR_ITEMS, "");
    k = (char)findmatch(tok[0], ReportWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[0]);
    switch ( k )
    {
      case 0: // Input
        m = findmatch(tok[1], NoYesWords);
        if      ( m == YES ) RptFlags.input = TRUE;
        else if ( m == NO )  RptFlags.input = FALSE;
        else                 return error_setInpError(ERR_KEYWORD, tok[1]);
        return 0;

      case 1: // Continuity
        m = findmatch(tok[1], NoYesWords);
        if      ( m == YES ) RptFlags.continuity = TRUE;
        else if ( m == NO )  RptFlags.continuity = FALSE;
        else                 return error_setInpError(ERR_KEYWORD, tok[1]);
        return 0;

      case 2: // Flow Statistics
        m = findmatch(tok[1], NoYesWords);
        if      ( m == YES ) RptFlags.flowStats = TRUE;
        else if ( m == NO )  RptFlags.flowStats = FALSE;
        else                 return error_setInpError(ERR_KEYWORD, tok[1]);
        return 0;

      case 3: // Controls
        m = findmatch(tok[1], NoYesWords);
        if      ( m == YES ) RptFlags.controls = TRUE;
        else if ( m == NO )  RptFlags.controls = FALSE;
        else                 return error_setInpError(ERR_KEYWORD, tok[1]);
        return 0;

      case 4:  m = SUBCATCH;  break;  // Subcatchments
      case 5:  m = NODE;      break;  // Nodes
      case 6:  m = LINK;      break;  // Links

      case 7: // Node Statistics
        m = findmatch(tok[1], NoYesWords);
        if      ( m == YES ) RptFlags.nodeStats = TRUE;
        else if ( m == NO )  RptFlags.nodeStats = FALSE;
        else                 return error_setInpError(ERR_KEYWORD, tok[1]);
        return 0;

      default: return error_setInpError(ERR_KEYWORD, tok[1]);
    }
    k = (char)findmatch(tok[1], NoneAllWords);
    if ( k < 0 )
    {
        k = SOME;
        for (t = 1; t < ntoks; t++)
        {
            j = project_findObject(m, tok[t]);
            if ( j < 0 ) return error_setInpError(ERR_NAME, tok[t]);
            switch ( m )
            {
              case SUBCATCH:  Subcatch[j].rptFlag = TRUE;  break;
              case NODE:      Node[j].rptFlag = TRUE;  break;
              case LINK:      Link[j].rptFlag = TRUE;  break;
            }
        }
    }
    switch ( m )
    {
      case SUBCATCH: RptFlags.subcatchments = k;  break;
      case NODE:     RptFlags.nodes = k;  break;
      case LINK:     RptFlags.links = k;  break;
    }
    return 0;
}

//=============================================================================

void report_writeLine(char *line)
//
//  Input:   line = line of text
//  Output:  none
//  Purpose: writes line of text to report file.
//
{
    if ( Frpt.file ) fprintf(Frpt.file, "\n  %s", line);
}

//=============================================================================

void report_writeSysTime(void)
//
//  Input:   none
//  Output:  none
//  Purpose: writes starting/ending processing times to report file.
//
{
    char    theTime[9];
    double  elapsedTime;
    if ( Frpt.file )
    {
        fprintf(Frpt.file, FMT20, ctime(&SysTime));
        elapsedTime = difftime(time(0), SysTime);
        fprintf(Frpt.file, FMT21);
        if ( elapsedTime < 1.0 ) fprintf(Frpt.file, "< 1 sec");
        else
        {
            datetime_timeToStr(elapsedTime/SECperDAY, theTime);
            fprintf(Frpt.file, "%s", theTime);
        }
    }
}

//=============================================================================

void report_writeLogo()
//
//  Input:   none
//  Output:  none
//  Purpose: writes report header lines to report file.
//
{
    fprintf(Frpt.file, FMT08);
    fprintf(Frpt.file, FMT09);
    fprintf(Frpt.file, FMT10);
    time(&SysTime);                    // Save starting wall clock time
}

//=============================================================================

void report_writeTitle()
//
//  Input:   none
//  Output:  none
//  Purpose: writes project title to report file.
//
{
    int i;
    if ( ErrorCode ) return;
    for (i=0; i<MAXTITLE; i++) if ( strlen(Title[i]) > 0 )
    {
        WRITE(Title[i]);
    }
    report_Options();
}

//=============================================================================

void report_Options()
//
//  Input:   none
//  Output:  none
//  Purpose: writes analysis options in use to report file.
//
{
    char str[80];
    WRITE("");
    WRITE("****************");
    WRITE("Analysis Options");
    WRITE("****************");
    fprintf(Frpt.file, "\n  Flow Units ............... %s",
        FlowUnitWords[FlowUnits]);
    if ( Nobjects[SUBCATCH] > 0 )
    fprintf(Frpt.file, "\n  Infiltration Method ...... %s", 
        InfilModelWords[InfilModel]);
    if ( Nobjects[LINK] > 0 )
    fprintf(Frpt.file, "\n  Flow Routing Method ...... %s",
        RouteModelWords[RouteModel]);
    datetime_dateToStr(StartDate, str);
    fprintf(Frpt.file, "\n  Starting Date ............ %s", str);
    datetime_timeToStr(StartTime, str);
    fprintf(Frpt.file, " %s", str);
    datetime_dateToStr(EndDate, str);
    fprintf(Frpt.file, "\n  Ending Date .............. %s", str);
    datetime_timeToStr(EndTime, str);
    fprintf(Frpt.file, " %s", str);

////////////////////////////////////////////
//  Antecedent Dry Days added. (LR - 9/5/05)
////////////////////////////////////////////
    fprintf(Frpt.file, "\n  Antecedent Dry Days ...... %.1f", StartDryDays);

////////////////////////////////////
//  New option added. (LR - 3/10/06)
////////////////////////////////////
    if ( IgnoreRainfall) fprintf(Frpt.file,
                       "\n  Rainfall Ignored ......... YES");

    datetime_timeToStr(datetime_encodeTime(0, 0, ReportStep), str);
    fprintf(Frpt.file, "\n  Report Time Step ......... %s", str);
    if ( Nobjects[SUBCATCH] > 0 )
    {
        datetime_timeToStr(datetime_encodeTime(0, 0, WetStep), str);
        fprintf(Frpt.file, "\n  Wet Time Step ............ %s", str);
        datetime_timeToStr(datetime_encodeTime(0, 0, DryStep), str);
        fprintf(Frpt.file, "\n  Dry Time Step ............ %s", str);
    }
    if ( Nobjects[LINK] > 0 )
    {
        fprintf(Frpt.file, "\n  Routing Time Step ........ %.2f sec", RouteStep);
    }
    WRITE("");
}


//=============================================================================
//      INPUT SUMMARY REPORT
//=============================================================================

void report_writeInput()
//
//  Input:   none
//  Output:  none
//  Purpose: writes summary of input data to report file.
//
{
    int m;
    int i, k;
    if ( ErrorCode ) return;

    WRITE("");
    WRITE("*************");
    WRITE("Element Count");
    WRITE("*************");
    fprintf(Frpt.file, "\n  Number of rain gages ...... %d", Nobjects[GAGE]);
    fprintf(Frpt.file, "\n  Number of subcatchments ... %d", Nobjects[SUBCATCH]);
    fprintf(Frpt.file, "\n  Number of nodes ........... %d", Nobjects[NODE]);
    fprintf(Frpt.file, "\n  Number of links ........... %d", Nobjects[LINK]);
    fprintf(Frpt.file, "\n  Number of pollutants ...... %d", Nobjects[POLLUT]);
    fprintf(Frpt.file, "\n  Number of land uses ....... %d", Nobjects[LANDUSE]);

    if ( Nobjects[POLLUT] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("*****************");
        WRITE("Pollutant Summary");
        WRITE("*****************");
        fprintf(Frpt.file,
    "\n                              Ppt.      GW         Kdecay");
        fprintf(Frpt.file,
    "\n  Name                Units   Concen.   Concen.    1/days    CoPollutant");
        fprintf(Frpt.file,
    "\n  ----------------------------------------------------------------------");
        for (i = 0; i < Nobjects[POLLUT]; i++)
        {
            fprintf(Frpt.file, "\n  %-20s%5s%10.2f%10.2f%10.2f", Pollut[i].ID,
                QualUnitsWords[Pollut[i].units], Pollut[i].pptConcen,
                Pollut[i].gwConcen, Pollut[i].kDecay*SECperDAY);
            if ( Pollut[i].coPollut >= 0 )
                fprintf(Frpt.file, "    %-s  (%.2f)",
                    Pollut[Pollut[i].coPollut].ID, Pollut[i].coFraction);
        }
    }

    if ( Nobjects[LANDUSE] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("***************");
        WRITE("Landuse Summary");
        WRITE("***************");
        fprintf(Frpt.file,
    "\n                        Sweeping   Maximum      Last");
        fprintf(Frpt.file,
    "\n  Name                  Interval   Removal     Swept");
        fprintf(Frpt.file,
    "\n  --------------------------------------------------");
        for (i=0; i<Nobjects[LANDUSE]; i++)
        {
             fprintf(Frpt.file, "\n  %-20s%10.2f%10.2f%10.2f", Landuse[i].ID,
                 Landuse[i].sweepInterval, Landuse[i].sweepRemoval,
                 Landuse[i].sweepDays0);
        }
    }

    if ( Nobjects[GAGE] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("****************");
        WRITE("Raingage Summary");
        WRITE("****************");
    fprintf(Frpt.file,
"\n                                          Data        Interval");
    fprintf(Frpt.file,
"\n  Name                Data Source         Type           hours");
    fprintf(Frpt.file,
"\n  ------------------------------------------------------------");
        for (i = 0; i < Nobjects[GAGE]; i++)
        {
            if ( Gage[i].tSeries >= 0 )
            {
                fprintf(Frpt.file, "\n  %-20s%-20s%-10s%10.2f",
                    Gage[i].ID, Tseries[Gage[i].tSeries].ID,
                    RainTypeWords[Gage[i].rainType],
//////////////////////////////////////////////////////////////////////////////
//  Conversion from integer seconds to decimal hours corrected. (LR - 3/10/06)
//////////////////////////////////////////////////////////////////////////////
                    (float)(Gage[i].rainInterval)/3600.0);
            }
            else fprintf(Frpt.file, "\n  %-20s%-20s", Gage[i].ID, Gage[i].fname);
        }
    }

    if ( Nobjects[SUBCATCH] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("********************");
        WRITE("Subcatchment Summary");
        WRITE("********************");

////////////////////////////////////////////////////////////////////
////  Outlet added to subcatchment properties listed. (LR - 7/5/06 )
////////////////////////////////////////////////////////////////////
        fprintf(Frpt.file,
"\n  Name                      Area     Width   %%Imperv    %%Slope    Rain Gage            Outlet          ");
        fprintf(Frpt.file,
"\n  -------------------------------------------------------------------------------------------------------");
        for (i = 0; i < Nobjects[SUBCATCH]; i++)
        {
            fprintf(Frpt.file,"\n  %-20s%10.2f%10.2f%10.2f%10.4f    %-20s",
                Subcatch[i].ID, Subcatch[i].area*UCF(LANDAREA),
                Subcatch[i].width*UCF(LENGTH),  Subcatch[i].fracImperv*100.0,
                Subcatch[i].slope*100.0, Gage[Subcatch[i].gage].ID);
            if ( Subcatch[i].outNode >= 0 )
            {
                fprintf(Frpt.file, " %-20s", Node[Subcatch[i].outNode].ID);
            }
            else if ( Subcatch[i].outSubcatch >= 0 )
            {
                fprintf(Frpt.file, " %-20s", Subcatch[Subcatch[i].outSubcatch].ID);
            }
        }
    }

    if ( Nobjects[NODE] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("************");
        WRITE("Node Summary");
        WRITE("************");

/////////////////////////////////////////////////////////
//  Ponded area added to node summary. (LR - 3/10/06)
//  External Inflow added to node summary. (LR - 7/5/06 )
/////////////////////////////////////////////////////////
        fprintf(Frpt.file,
"\n                                          Invert      Max.    Ponded    External");
        fprintf(Frpt.file,
"\n  Name                Type                 Elev.     Depth      Area    Inflow  ");
        fprintf(Frpt.file,
"\n  ------------------------------------------------------------------------------");
        for (i = 0; i < Nobjects[NODE]; i++)
        {
            fprintf(Frpt.file, "\n  %-20s%-16s%10.2f%10.2f%10.0f", Node[i].ID,
                NodeTypeWords[Node[i].type-JUNCTION],
                Node[i].invertElev*UCF(LENGTH),
                Node[i].fullDepth*UCF(LENGTH),
                Node[i].pondedArea*UCF(LENGTH)*UCF(LENGTH));
            if ( Node[i].extInflow || Node[i].dwfInflow || Node[i].rdiiInflow )
            {
                fprintf(Frpt.file, "    Yes");
            }
        }
    }

    if ( Nobjects[LINK] > 0 )
    {
        WRITE("");
        WRITE("");
        WRITE("************");
        WRITE("Link Summary");
        WRITE("************");
        fprintf(Frpt.file,
"\n  Name            From Node       To Node         Type            Length    %%Slope         N");
        fprintf(Frpt.file,
"\n  ------------------------------------------------------------------------------------------");
        for (i = 0; i < Nobjects[LINK]; i++)
        {
            fprintf(Frpt.file, "\n  %-16s%-16s%-16s%-12s",
                Link[i].ID, Node[Link[i].node1].ID, Node[Link[i].node2].ID,
                LinkTypeWords[Link[i].type-CONDUIT]);
            if (Link[i].type == CONDUIT)
            {
                k = Link[i].subIndex;
                fprintf(Frpt.file, "%10.0f%10.4f%10.4f",
                    Conduit[k].length*UCF(LENGTH), Conduit[k].slope*100.0,
                    Conduit[k].roughness);
            }
        }

///////////////////////////////////////////////////////////////////
//  Number of barrels added to Cross Section Summary. (LR - 7/5/06)
///////////////////////////////////////////////////////////////////
        WRITE("");
        WRITE("");
        WRITE("*********************");
        WRITE("Cross Section Summary");
        WRITE("*********************");
        fprintf(Frpt.file,
"\n                                        Full     Full     Hyd.     Max.   No. of     Full");
        fprintf(Frpt.file,    
"\n  Conduit          Shape               Depth     Area     Rad.    Width  Barrels     Flow");
        fprintf(Frpt.file,
"\n  ---------------------------------------------------------------------------------------");
        for (i = 0; i < Nobjects[LINK]; i++)
        {
            if (Link[i].type == CONDUIT)
            {
                k = Link[i].subIndex;
                fprintf(Frpt.file, "\n  %-16s %-16s %8.2f %8.2f %8.2f %8.2f   %3d    %8.2f",
                    Link[i].ID,
                    XsectTypeWords[Link[i].xsect.type],
                    Link[i].xsect.yFull*UCF(LENGTH),
                    Link[i].xsect.aFull*UCF(LENGTH)*UCF(LENGTH),
                    Link[i].xsect.rFull*UCF(LENGTH),
                    Link[i].xsect.wMax*UCF(LENGTH),
                    Conduit[k].barrels,
                    Link[i].qFull*UCF(FLOW));
            }
        }
    }

    if (Nobjects[TRANSECT] > 0)
    {
        WRITE("");
        WRITE("");
        WRITE("****************");
        WRITE("Transect Summary");
        WRITE("****************");
        for (i = 0; i < Nobjects[TRANSECT]; i++)
        {
            fprintf(Frpt.file, "\n\n  Transect %s", Transect[i].ID);
            fprintf(Frpt.file, "\n  Area:  ");
            for ( m = 1; m <= 25; m++)
            {
                 if ( m % 5 == 1 ) fprintf(Frpt.file,"\n          ");
                 fprintf(Frpt.file, "%10.4f ", Transect[i].areaTbl[m]);
            }
            fprintf(Frpt.file, "\n  Hrad:  ");
            for ( m = 1; m <= 25; m++)
            {
                 if ( m % 5 == 1 ) fprintf(Frpt.file,"\n          ");
                 fprintf(Frpt.file, "%10.4f ", Transect[i].hradTbl[m]);
            }
            fprintf(Frpt.file, "\n  Width: ");
            for ( m = 1; m <= 25; m++)
            {
                 if ( m % 5 == 1 ) fprintf(Frpt.file,"\n          ");
                 fprintf(Frpt.file, "%10.4f ", Transect[i].widthTbl[m]);
            }
        }
    }
    WRITE("");
}


//=============================================================================
//      SIMULATION RESULTS REPORT
//=============================================================================

void report_writeReport()
//
//  Input:   none
//  Output:  none
//  Purpose: writes simulation results to report file.
//
{
    if ( ErrorCode ) return;
    if ( Nperiods == 0 ) return;
    if ( RptFlags.subcatchments != NONE ) report_Subcatchments();
    if ( RptFlags.nodes != NONE )         report_Nodes();
    if ( RptFlags.links != NONE )         report_Links();
}

//=============================================================================

void report_Subcatchments()
//
//  Input:   none
//  Output:  none
//  Purpose: writes results for selected subcatchments to report file.
//
{
    int      j, p;
    long     period;
    DateTime days;
    char     theDate[12];
    char     theTime[9];

    if ( Nobjects[SUBCATCH] == 0 ) return;
    WRITE("");
    WRITE("********************");
    WRITE("Subcatchment Results");
    WRITE("********************");
    for (j = 0; j < Nobjects[SUBCATCH]; j++)
    {
        if ( RptFlags.subcatchments == ALL || Subcatch[j].rptFlag == TRUE)
        {
            report_SubcatchHeader(Subcatch[j].ID);
            for ( period = 1; period <= Nperiods; period++ )
            {
                output_readDateTime(period, &days);
                datetime_dateToStr(days, theDate);
                datetime_timeToStr(days, theTime);
                output_readSubcatchResults(period, j);

//////////////////////////////////////////////////////////
//  Modified to add Losses to report table. (LR - 3/10/06)
//////////////////////////////////////////////////////////
                fprintf(Frpt.file, "\n  %11s %8s %10.3f%10.3f%10.4f",
                    theDate, theTime, SubcatchResults[SUBCATCH_RAINFALL],
                    SubcatchResults[SUBCATCH_LOSSES],
                    SubcatchResults[SUBCATCH_RUNOFF]);

                for (p = 0; p < Nobjects[POLLUT]; p++)
                    fprintf(Frpt.file, "%10.3f",
                        SubcatchResults[SUBCATCH_WASHOFF+p]);
            }
            WRITE("");
        }
    }
}

//=============================================================================

void  report_SubcatchHeader(char *id)
//
//  Input:   id = subcatchment ID name
//  Output:  none
//  Purpose: writes table headings for subcatchment results to report file.
//
{
    int i;
    WRITE("");
    fprintf(Frpt.file,"\n  <<< Subcatchment %s >>>", id);
    WRITE(LINE_41);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);

///////////////////////////////////////////////////////////////////////
//  Modified to add a Losses column to the report table. (LR - 3/10/06)
///////////////////////////////////////////////////////////////////////

    fprintf(Frpt.file,
    "\n  Date        Time       Rainfall    Losses    Runoff");

    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", Pollut[i].ID);

    if ( UnitSystem == US ) fprintf(Frpt.file, 
    "\n                            in/hr     in/hr %9s", FlowUnitWords[FlowUnits]);
    else fprintf(Frpt.file, 
    "\n                            mm/hr     mm/hr %9s", FlowUnitWords[FlowUnits]);

    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", QualUnitsWords[Pollut[i].units]);

    WRITE(LINE_41);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);
}

//=============================================================================

void report_Nodes()
//
//  Input:   none
//  Output:  none
//  Purpose: writes results for selected nodes to report file.
//
{
    int      j, p;
    long     period;
    DateTime days;
    char     theDate[20];
    char     theTime[20];

    if ( Nobjects[NODE] == 0 ) return;
    WRITE("");
    WRITE("************");
    WRITE("Node Results");
    WRITE("************");
    for (j = 0; j < Nobjects[NODE]; j++)
    {
        if ( RptFlags.nodes == ALL || Node[j].rptFlag == TRUE)
        {
            report_NodeHeader(Node[j].ID);
            for ( period = 1; period <= Nperiods; period++ )
            {
                output_readDateTime(period, &days);
                datetime_dateToStr(days, theDate);
                datetime_timeToStr(days, theTime);
                output_readNodeResults(period, j);
                fprintf(Frpt.file, "\n  %11s %8s  %9.3f %9.3f %9.3f %9.3f",
                    theDate, theTime, NodeResults[NODE_INFLOW],
                    NodeResults[NODE_OVERFLOW], NodeResults[NODE_DEPTH],
                    NodeResults[NODE_HEAD]);
                for (p = 0; p < Nobjects[POLLUT]; p++)
                    fprintf(Frpt.file, " %9.3f", NodeResults[NODE_QUAL + p]);
            }
            WRITE("");
        }
    }
}

//=============================================================================

void  report_NodeHeader(char *id)
//
//  Input:   id = node ID name
//  Output:  none
//  Purpose: writes table headings for node results to report file.
//
{
    int i;
    char lengthUnits[9];
    WRITE("");
    fprintf(Frpt.file,"\n  <<< Node %s >>>", id);
    WRITE(LINE_61);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);

    fprintf(Frpt.file,
    "\n                           Inflow  Flooding     Depth      Head");
    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", Pollut[i].ID);
    if ( UnitSystem == US) strcpy(lengthUnits, "feet");
    else strcpy(lengthUnits, "meters");
    fprintf(Frpt.file,
    "\n  Date        Time      %9s %9s %9s %9s",
        FlowUnitWords[FlowUnits], FlowUnitWords[FlowUnits],
        lengthUnits, lengthUnits);
    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", QualUnitsWords[Pollut[i].units]);

    WRITE(LINE_61);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);
}

//=============================================================================

void report_Links()
//
//  Input:   none
//  Output:  none
//  Purpose: writes results for selected links to report file.
//
{
    int      j, p;
    long     period;
    DateTime days;
    char     theDate[12];
    char     theTime[9];

    if ( Nobjects[LINK] == 0 ) return;
    WRITE("");
    WRITE("************");
    WRITE("Link Results");
    WRITE("************");
    for (j = 0; j < Nobjects[LINK]; j++)
    {
        if ( RptFlags.links == ALL || Link[j].rptFlag == TRUE)
        {
            report_LinkHeader(Link[j].ID);
            for ( period = 1; period <= Nperiods; period++ )
            {
                output_readDateTime(period, &days);
                datetime_dateToStr(days, theDate);
                datetime_timeToStr(days, theTime);
                output_readLinkResults(period, j);
                fprintf(Frpt.file, "\n  %11s %8s  %9.3f %9.1f %9.3f %9.1f",
                    theDate, theTime, LinkResults[LINK_FLOW], 
                    LinkResults[LINK_VELOCITY], LinkResults[LINK_DEPTH],
                    LinkResults[LINK_CAPACITY]*100.0);
                for (p = 0; p < Nobjects[POLLUT]; p++)
                    fprintf(Frpt.file, " %9.3f", LinkResults[LINK_QUAL + p]);
            }
            WRITE("");
        }
    }
}

//=============================================================================

void  report_LinkHeader(char *id)
//
//  Input:   id = link ID name
//  Output:  none
//  Purpose: writes table headings for link results to report file.
//
{
    int i;
    WRITE("");
    fprintf(Frpt.file,"\n  <<< Link %s >>>", id);
    WRITE(LINE_61);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);

    fprintf(Frpt.file, 
    "\n                             Flow  Velocity     Depth   Percent");
    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", Pollut[i].ID);

    if ( UnitSystem == US )
        fprintf(Frpt.file,
        "\n  Date        Time     %10s    ft/sec      feet      Full",
        FlowUnitWords[FlowUnits]);
    else
        fprintf(Frpt.file,
        "\n  Date        Time     %10s     m/sec    meters      Full",
        FlowUnitWords[FlowUnits]);
    for (i = 0; i < Nobjects[POLLUT]; i++)
        fprintf(Frpt.file, "%10s", QualUnitsWords[Pollut[i].units]);

    WRITE(LINE_61);
    for (i = 0; i < Nobjects[POLLUT]; i++) fprintf(Frpt.file, LINE_10);
}


//=============================================================================
//      CONTINUITY ERROR REPORT
//=============================================================================

void report_writeRunoffError(TRunoffTotals* totals, double totalArea)
//
//  Input:  totals = accumulated runoff totals
//          totalArea = total area of all subcatchments
//  Output:  none
//  Purpose: writes runoff continuity error to report file. 
//
{
    WRITE("");

    if ( Frunoff.mode == USE_FILE )
    {
        fprintf(Frpt.file,
        "\n  **************************"
        "\n  Runoff Quantity Continuity"
        "\n  **************************"
        "\n  Runoff supplied by interface file %s", Frunoff.name);
        WRITE("");
        return;
    }

    fprintf(Frpt.file,
    "\n  **************************        Volume         Depth");
    if ( UnitSystem == US) fprintf(Frpt.file,
    "\n  Runoff Quantity Continuity     acre-feet        inches");
    else fprintf(Frpt.file,
    "\n  Runoff Quantity Continuity     hectare-m            mm");
    fprintf(Frpt.file,
    "\n  **************************     ---------       -------");

    if ( Nobjects[SNOWMELT] > 0 )
    {
        fprintf(Frpt.file, "\n  Initial Snow Cover .......%14.3f%14.3f",
            totals->initSnowCover * UCF(LENGTH) * UCF(LANDAREA),
            totals->initSnowCover / totalArea * UCF(RAINDEPTH));
    }

    fprintf(Frpt.file, "\n  Total Precipitation ......%14.3f%14.3f",
            totals->rainfall * UCF(LENGTH) * UCF(LANDAREA),
            totals->rainfall / totalArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Evaporation Loss .........%14.3f%14.3f",
            totals->evap * UCF(LENGTH) * UCF(LANDAREA),
            totals->evap / totalArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Infiltration Loss ........%14.3f%14.3f",
            totals->infil * UCF(LENGTH) * UCF(LANDAREA),
            totals->infil / totalArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Surface Runoff ...........%14.3f%14.3f",
            totals->runoff * UCF(LENGTH) * UCF(LANDAREA),
            totals->runoff / totalArea * UCF(RAINDEPTH));

    if ( Nobjects[SNOWMELT] > 0 )
    {
        fprintf(Frpt.file, "\n  Snow Removed .............%14.3f%14.3f",
            totals->snowRemoved * UCF(LENGTH) * UCF(LANDAREA),
            totals->snowRemoved / totalArea * UCF(RAINDEPTH));
        fprintf(Frpt.file, "\n  Final Snow Cover .........%14.3f%14.3f",
            totals->finalSnowCover * UCF(LENGTH) * UCF(LANDAREA),
            totals->finalSnowCover / totalArea * UCF(RAINDEPTH));
    }

    fprintf(Frpt.file, "\n  Final Surface Storage ....%14.3f%14.3f",
            totals->finalStorage * UCF(LENGTH) * UCF(LANDAREA),
            totals->finalStorage / totalArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Continuity Error (%%) .....%14.3f",
            totals->pctError);
    WRITE("");
}

//=============================================================================

void report_writeLoadingError(TLoadingTotals* totals)
//
//  Input:   totals = accumulated pollutant loading totals
//           area = total area of all subcatchments
//  Output:  none
//  Purpose: writes runoff loading continuity error to report file. 
//
{
    int p1, p2;
    p1 = 1;
    p2 = MIN(5, Nobjects[POLLUT]);
    while ( p1 <= Nobjects[POLLUT] )
    {
        report_LoadingErrors(p1-1, p2-1, totals);
        p1 = p2 + 1;
        p2 = p1 + 4;
        p2 = MIN(p2, Nobjects[POLLUT]);
    }
}

//=============================================================================

void report_LoadingErrors(int p1, int p2, TLoadingTotals* totals)
//
//  Input:   p1 = index of first pollutant to report
//           p2 = index of last pollutant to report
//           totals = accumulated pollutant loading totals
//           area = total area of all subcatchments
//  Output:  none
//  Purpose: writes runoff loading continuity error to report file for
//           up to 5 pollutants at a time. 
//
{
    int    i;
    int    p;
    double cf = 1.0;
    char   units[15];

    WRITE("");
    fprintf(Frpt.file, "\n  **************************");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14s", Pollut[p].ID);
    }
    fprintf(Frpt.file, "\n  Runoff Quality Continuity ");
    for (p = p1; p <= p2; p++)
    {
        i = UnitSystem;
        if ( Pollut[p].units == COUNT ) i = 2;
        strcpy(units, LoadUnitsWords[i]);
        fprintf(Frpt.file, "%14s", units);
    }
    fprintf(Frpt.file, "\n  **************************");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "    ----------");
    }

    fprintf(Frpt.file, "\n  Initial Buildup ..........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].initLoad*cf);
    }
    fprintf(Frpt.file, "\n  Surface Buildup ..........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].buildup*cf);
    }
    fprintf(Frpt.file, "\n  Wet Deposition ...........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].deposition*cf);
    }
    fprintf(Frpt.file, "\n  Sweeping Removal .........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].sweeping*cf);
    }


/////////////////////////////////////
// New load type added (LR - 7/5/06 )
/////////////////////////////////////
    fprintf(Frpt.file, "\n  Infiltration Loss ........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].infil*cf);
    }


    fprintf(Frpt.file, "\n  BMP Removal ..............");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].bmpRemoval*cf);
    }
    fprintf(Frpt.file, "\n  Surface Runoff ...........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].runoff*cf);
    }
    fprintf(Frpt.file, "\n  Remaining Buildup ........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].finalLoad*cf);
    }
    fprintf(Frpt.file, "\n  Continuity Error (%%) .....");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", totals[p].pctError);
    }
    WRITE("");
}

//=============================================================================

void report_writeGwaterError(TGwaterTotals* totals, double gwArea)
//
//  Input:   totals = accumulated groundwater totals
//           gwArea = total area of all subcatchments with groundwater
//  Output:  none
//  Purpose: writes groundwater continuity error to report file. 
//
{
    WRITE("");
    fprintf(Frpt.file,
    "\n  **************************        Volume         Depth");
    if ( UnitSystem == US) fprintf(Frpt.file,
    "\n  Groundwater Continuity         acre-feet        inches");
    else fprintf(Frpt.file,
    "\n  Groundwater Continuity         hectare-m            mm");
    fprintf(Frpt.file,
    "\n  **************************     ---------       -------");
    fprintf(Frpt.file, "\n  Initial Storage ..........%14.3f%14.3f",
            totals->initStorage * UCF(LENGTH) * UCF(LANDAREA),
            totals->initStorage / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Infiltration .............%14.3f%14.3f", 
            totals->infil * UCF(LENGTH) * UCF(LANDAREA),
            totals->infil / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Upper Zone ET ............%14.3f%14.3f", 
            totals->upperEvap * UCF(LENGTH) * UCF(LANDAREA),
            totals->upperEvap / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Lower Zone ET ............%14.3f%14.3f", 
            totals->lowerEvap * UCF(LENGTH) * UCF(LANDAREA),
            totals->lowerEvap / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Deep Percolation .........%14.3f%14.3f", 
            totals->lowerPerc * UCF(LENGTH) * UCF(LANDAREA),
            totals->lowerPerc / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Groundwater Flow .........%14.3f%14.3f",
            totals->gwater * UCF(LENGTH) * UCF(LANDAREA),
            totals->gwater / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Final Storage ............%14.3f%14.3f",
            totals->finalStorage * UCF(LENGTH) * UCF(LANDAREA),
            totals->finalStorage / gwArea * UCF(RAINDEPTH));

    fprintf(Frpt.file, "\n  Continuity Error (%%) .....%14.3f",
            totals->pctError);
    WRITE("");
}

//=============================================================================

void report_writeFlowError(TRoutingTotals *totals)
//
//  Input:  totals = accumulated flow routing totals
//  Output:  none
//  Purpose: writes flow routing continuity error to report file. 
//
{
    float ucf1, ucf2;

    ucf1 = UCF(LENGTH) * UCF(LANDAREA);
    if ( UnitSystem == US) ucf2 = MGDperCFS / SECperDAY;
    else                   ucf2 = MLDperCFS / SECperDAY;

    WRITE("");
    fprintf(Frpt.file,
    "\n  **************************        Volume        Volume");
    if ( UnitSystem == US) fprintf(Frpt.file,
    "\n  Flow Routing Continuity        acre-feet      Mgallons");
    else fprintf(Frpt.file,
    "\n  Flow Routing Continuity        hectare-m       Mliters");
    fprintf(Frpt.file,
    "\n  **************************     ---------     ---------");

    fprintf(Frpt.file, "\n  Dry Weather Inflow .......%14.3f%14.3f",
            totals->dwInflow * ucf1, totals->dwInflow * ucf2);

    fprintf(Frpt.file, "\n  Wet Weather Inflow .......%14.3f%14.3f",
            totals->wwInflow * ucf1, totals->wwInflow * ucf2);

    fprintf(Frpt.file, "\n  Groundwater Inflow .......%14.3f%14.3f",
            totals->gwInflow * ucf1, totals->gwInflow * ucf2);

    fprintf(Frpt.file, "\n  RDII Inflow ..............%14.3f%14.3f",
            totals->iiInflow * ucf1, totals->iiInflow * ucf2);

    fprintf(Frpt.file, "\n  External Inflow ..........%14.3f%14.3f",
            totals->exInflow * ucf1, totals->exInflow * ucf2);

    fprintf(Frpt.file, "\n  External Outflow .........%14.3f%14.3f",
            totals->outflow * ucf1, totals->outflow * ucf2);

    fprintf(Frpt.file, "\n  Surface Flooding .........%14.3f%14.3f",
            totals->flooding * ucf1, totals->flooding * ucf2);

    fprintf(Frpt.file, "\n  Evaporation Loss .........%14.3f%14.3f",
            totals->reacted * ucf1, totals->reacted * ucf2);

    fprintf(Frpt.file, "\n  Initial Stored Volume ....%14.3f%14.3f",
            totals->initStorage * ucf1, totals->initStorage * ucf2);

    fprintf(Frpt.file, "\n  Final Stored Volume ......%14.3f%14.3f",
            totals->finalStorage * ucf1, totals->finalStorage * ucf2);

    fprintf(Frpt.file, "\n  Continuity Error (%%) .....%14.3f",
            totals->pctError);
    WRITE("");
}

//=============================================================================

void report_writeQualError(TRoutingTotals QualTotals[])
//
//  Input:   totals = accumulated quality routing totals for each pollutant
//  Output:  none
//  Purpose: writes quality routing continuity error to report file. 
//
{
    int p1, p2;
    p1 = 1;
    p2 = MIN(5, Nobjects[POLLUT]);
    while ( p1 <= Nobjects[POLLUT] )
    {
        report_QualErrors(p1-1, p2-1, QualTotals);
        p1 = p2 + 1;
        p2 = p1 + 4;
        p2 = MIN(p2, Nobjects[POLLUT]);
    }
}

//=============================================================================

void report_QualErrors(int p1, int p2, TRoutingTotals QualTotals[])
{
    int   i;
    int   p;
    char  units[15];

    WRITE("");
    fprintf(Frpt.file, "\n  **************************");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14s", Pollut[p].ID);
    }
    fprintf(Frpt.file, "\n  Quality Routing Continuity");
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        i = UnitSystem;
        if ( Pollut[p].units == COUNT ) i = 2;
        strcpy(units, LoadUnitsWords[i]);
        fprintf(Frpt.file, "%14s", units);
    }
    fprintf(Frpt.file, "\n  **************************");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "    ----------");
    }

    fprintf(Frpt.file, "\n  Dry Weather Inflow .......");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].dwInflow);
    }

    fprintf(Frpt.file, "\n  Wet Weather Inflow .......");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].wwInflow);
    }

    fprintf(Frpt.file, "\n  Groundwater Inflow .......");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].gwInflow);
    }

    fprintf(Frpt.file, "\n  RDII Inflow ..............");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].iiInflow);
    }

    fprintf(Frpt.file, "\n  External Inflow ..........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].exInflow);
    }

    fprintf(Frpt.file, "\n  Internal Flooding ........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].flooding);
    }

    fprintf(Frpt.file, "\n  External Outflow .........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].outflow);
    }

    fprintf(Frpt.file, "\n  Mass Reacted .............");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].reacted);
    }

    fprintf(Frpt.file, "\n  Initial Stored Mass ......");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].initStorage);
    }

    fprintf(Frpt.file, "\n  Final Stored Mass ........");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].finalStorage);
    }

    fprintf(Frpt.file, "\n  Continuity Error (%%) .....");
    for (p = p1; p <= p2; p++)
    {
        fprintf(Frpt.file, "%14.3f", QualTotals[p].pctError);
    }
    WRITE("");
}


//=============================================================================
//      SIMULATION STATISTICS REPORT
//=============================================================================

/////////////////////////////////////
//  New function added. (LR - 9/5/05)
/////////////////////////////////////
void   report_writeControlActionsHeading()
{
    WRITE("");
    WRITE("*********************");
    WRITE("Control Actions Taken");
    WRITE("*********************");
    fprintf(Frpt.file, "\n");
}

//=============================================================================

void   report_writeControlAction(DateTime aDate, char* linkID, float value,
                                 char* ruleID)
//
//  Input:   aDate  = date/time of rule action
//           linkID = ID of link being controlled
//           value  = new status value of link
//           ruleID = ID of rule implementing the action
//  Output:  none
//  Purpose: reports action taken by a control rule.
//
{
    char     theDate[12];
    char     theTime[9];
    datetime_dateToStr(aDate, theDate);
    datetime_timeToStr(aDate, theTime);

////////////////////////////////
//  Line modified. (LR - 9/5/05)
////////////////////////////////
    fprintf(Frpt.file,
            "  %11s: %8s Link %s setting changed to %6.2f by Control %s\n",
            theDate, theTime, linkID, value, ruleID);
}

//=============================================================================

void report_writeSubcatchStats(TSubcatchStats subcatchStats[], float maxRunoff)

//////////////////////////////////////////////////////////
////  Modified to report peak runoff. (LR - 3/10/06)  ////
////  and to report peak system runoff. (LR - 7/5/06) ////
//////////////////////////////////////////////////////////
{
    int    j;
    float  a, aSum, x, r;
    TSubcatchStats totals;

    if ( Nobjects[SUBCATCH] == 0 ) return;
    aSum = 0.0;
    totals.precip = 0.0;
    totals.runon  = 0.0;
    totals.evap   = 0.0;
    totals.infil  = 0.0;
    totals.runoff = 0.0;
    totals.maxFlow = maxRunoff * UCF(FLOW);                //Modified (LR - 7/5/06)
    WRITE("");
    WRITE("***************************");
    WRITE("Subcatchment Runoff Summary");
    WRITE("***************************");
    WRITE("");
    fprintf(Frpt.file,
"\n  --------------------------------------------------------------------------------------"
"\n                       Total     Total     Total     Total     Total      Peak    Runoff"
"\n                      Precip     Runon      Evap     Infil    Runoff    Runoff     Coeff");
    if ( UnitSystem == US ) fprintf(Frpt.file,
"\n  Subcatchment            in        in        in        in        in %9s",
        FlowUnitWords[FlowUnits]);
    else fprintf(Frpt.file,
"\n  Subcatchment            mm        mm        mm        mm        mm %9s",
        FlowUnitWords[FlowUnits]);
    fprintf(Frpt.file,
"\n  --------------------------------------------------------------------------------------");
    for ( j = 0; j < Nobjects[SUBCATCH]; j++ )
    {
        a = Subcatch[j].area;
        aSum += a;
        fprintf(Frpt.file, "\n  %-16s", Subcatch[j].ID);
        x = subcatchStats[j].precip * UCF(RAINDEPTH);
        fprintf(Frpt.file, "%10.3f", x);
        totals.precip += x * a;
        x = subcatchStats[j].runon * UCF(RAINDEPTH); 
        fprintf(Frpt.file, "%10.3f", x);
        totals.runon += x * a;
        x = subcatchStats[j].evap * UCF(RAINDEPTH);
        fprintf(Frpt.file, "%10.3f", x);
        totals.evap += x * a;
        x = subcatchStats[j].infil * UCF(RAINDEPTH); 
        fprintf(Frpt.file, "%10.3f", x);
        totals.infil += x * a;
        x = subcatchStats[j].runoff * UCF(RAINDEPTH);
        fprintf(Frpt.file, "%10.3f", x);
        totals.runoff += x * a;

        x = subcatchStats[j].maxFlow * UCF(FLOW);
        fprintf(Frpt.file, "%10.2f", x);
        //totals.maxFlow = MAX(totals.maxFlow, x);         //Deleted (LR - 7/5/06)

        r = subcatchStats[j].precip + subcatchStats[j].runon;
        if ( r > 0.0 ) r = subcatchStats[j].runoff / r;
        fprintf(Frpt.file, "%10.3f", r);
    }
    if ( aSum > 0.0 )
    {
        fprintf(Frpt.file,
"\n  --------------------------------------------------------------------------------------");
        fprintf(Frpt.file, "\n  System          ");
        fprintf(Frpt.file, "%10.3f", totals.precip / aSum);
        fprintf(Frpt.file, "%10.3f", totals.runon / aSum);
        fprintf(Frpt.file, "%10.3f", totals.evap / aSum);
        fprintf(Frpt.file, "%10.3f", totals.infil / aSum);
        fprintf(Frpt.file, "%10.3f", totals.runoff / aSum);

        fprintf(Frpt.file, "%10.2f", totals.maxFlow);

        r = totals.precip + totals.runon;
        if ( r > 0.0 ) r = totals.runoff / r;
        fprintf(Frpt.file, "%10.3f", r);
    }
    WRITE("");
}

//=============================================================================

//////////////////////////////////
////  New function. (LR - 7/5/06 )
//////////////////////////////////
void report_writeSubcatchLoads()
{
    int i, j, p;
    double x;
    double* totals; 
    char  units[15];
    char  subcatchLine[] = "----------------";
    char  pollutLine[]   = "--------------";

    // --- create an array to hold total loads for each pollutant
    totals = (double *) calloc(Nobjects[POLLUT], sizeof(double));
    if ( totals ) for (p = 0; p < Nobjects[POLLUT]; p++) totals[p] = 0.0;

    // --- print the table headings 
    WRITE("");
    WRITE("****************************");
    WRITE("Subcatchment Washoff Summary");
    WRITE("****************************");
    WRITE("");
    fprintf(Frpt.file, "\n  %s", subcatchLine);
    for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "%s", pollutLine);
    fprintf(Frpt.file, "\n                  ");
    for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "%14s", Pollut[p].ID);
    fprintf(Frpt.file, "\n  Subcatchment    ");
    for (p = 0; p < Nobjects[POLLUT]; p++)
    {
        i = UnitSystem;
        if ( Pollut[p].units == COUNT ) i = 2;
        strcpy(units, LoadUnitsWords[i]);
        fprintf(Frpt.file, "%14s", units);
    }
    fprintf(Frpt.file, "\n  %s", subcatchLine);
    for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "%s", pollutLine);

    // --- print the pollutant loadings from each subcatchment
    for ( j = 0; j < Nobjects[SUBCATCH]; j++ )
    {
        fprintf(Frpt.file, "\n  %-16s", Subcatch[j].ID);
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            x = Subcatch[j].totalLoad[p];
            if ( totals ) totals[p] += x;
            if ( Pollut[p].units == COUNT ) x = LOG10(x);
            fprintf(Frpt.file, "%14.3f", x); 
        }
    }

    // --- print the total loading of each pollutant
    if ( totals )
    {
        fprintf(Frpt.file, "\n  %s", subcatchLine);
        for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "%s", pollutLine);
        fprintf(Frpt.file, "\n  System          ");
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            x = totals[p];
            if ( Pollut[p].units == COUNT ) x = LOG10(x);
            fprintf(Frpt.file, "%14.3f", x); 
        }
    }
    FREE(totals);
    WRITE("");
}

//=============================================================================

void report_writeNodeStats(TNodeStats nodeStats[])
//
//  Input:   nodeStats = array of simulation statistics for nodes
//  Output:  none
//  Purpose: writes simulation statistics for nodes to report file.
//
{
    int j, days, hrs, mins;

    if ( Nobjects[LINK] == 0 ) return;
    WRITE("");
    WRITE("******************");
    WRITE("Node Depth Summary");
    WRITE("******************");
    WRITE("");

/////////////////////////////////////////
//  Table header modified. (LR - 3/10/06)
/////////////////////////////////////////
    fprintf(Frpt.file,
"\n  ----------------------------------------------------------------------------------------"
"\n                                 Average  Maximum  Maximum  Time of Max     Total    Total"
"\n                                   Depth    Depth      HGL   Occurrence  Flooding  Minutes");
    if ( UnitSystem == US ) fprintf(Frpt.file,
"\n  Node                 Type         Feet     Feet     Feet  days hr:min   acre-in  Flooded");
    else fprintf(Frpt.file,
"\n  Node                 Type       Meters   Meters   Meters  days hr:min     ha-mm  Flooded");
    fprintf(Frpt.file,
"\n  ----------------------------------------------------------------------------------------");

    for ( j = 0; j < Nobjects[NODE]; j++ )
    {
        fprintf(Frpt.file, "\n  %-20s", Node[j].ID);       ////Modified (LR - 3/10/06)
        fprintf(Frpt.file, " %-9s ",                       ////Added (LR - 3/10/06)
            NodeTypeWords[Node[j].type]);                  ////Added (LR - 3/10/06)

        getElapsedTime(nodeStats[j].maxDepthDate, &days, &hrs, &mins);
        fprintf(Frpt.file, "%7.2f  %7.2f  %7.2f  %4d  %02d:%02d",
            nodeStats[j].avgDepth / StepCount * UCF(LENGTH),
            nodeStats[j].maxDepth * UCF(LENGTH),
            (nodeStats[j].maxDepth + Node[j].invertElev) * UCF(LENGTH),
            days, hrs, mins);
        if ( nodeStats[j].volFlooded == 0.0 )
        {
            fprintf(Frpt.file, "         0");
        }
        else
        {
            fprintf(Frpt.file, "%10.2f",
                nodeStats[j].volFlooded * UCF(LANDAREA) * UCF(RAINDEPTH));
        }
        fprintf(Frpt.file, "  %7.0f", nodeStats[j].timeFlooded / 60.0);
    }
    WRITE("");
    writeNodeFlowStats(nodeStats);                         //Added (LR - 7/5/06)
}

//=============================================================================

//////////////////////////////////
////  New function. (LR - 7/5/06 )
//////////////////////////////////
void writeNodeFlowStats(TNodeStats nodeStats[])
//
//  Input:   nodeStats = array of simulation statistics for nodes
//  Output:  none
//  Purpose: writes flow statistics for nodes to report file.
//
{
    int j;
    int days1, hrs1, mins1;
    int days2, hrs2, mins2;

    WRITE("");
    WRITE("*****************");
    WRITE("Node Flow Summary");
    WRITE("*****************");
    WRITE("");

    fprintf(Frpt.file,
"\n  ------------------------------------------------------------------------------------"
"\n                                  Maximum  Maximum                Maximum             "
"\n                                  Lateral    Total  Time of Max  Flooding  Time of Max"
"\n                                   Inflow   Inflow   Occurrence  Overflow   Occurrence");
    fprintf(Frpt.file,
"\n  Node                 Type           %3s      %3s  days hr:min       %3s  days hr:min",
        FlowUnitWords[FlowUnits], FlowUnitWords[FlowUnits], FlowUnitWords[FlowUnits]);
    fprintf(Frpt.file,
"\n  ------------------------------------------------------------------------------------");

    for ( j = 0; j < Nobjects[NODE]; j++ )
    {
        fprintf(Frpt.file, "\n  %-20s", Node[j].ID);
        fprintf(Frpt.file, " %-9s", NodeTypeWords[Node[j].type]);

        getElapsedTime(nodeStats[j].maxInflowDate, &days1, &hrs1, &mins1);
        getElapsedTime(nodeStats[j].maxOverflowDate, &days2, &hrs2, &mins2);
        fprintf(Frpt.file, "  %7.2f  %7.2f  %4d  %02d:%02d  %7.2f",
            nodeStats[j].maxLatFlow * UCF(FLOW),
            nodeStats[j].maxInflow * UCF(FLOW),
            days1, hrs1, mins1, nodeStats[j].maxOverflow * UCF(FLOW));
        if ( nodeStats[j].maxOverflow > 0.0 )
            fprintf(Frpt.file, "  %4d  %02d:%02d", days2, hrs2, mins2);
    }
    WRITE("");
}

//=============================================================================

/////////////////////////////////////////////////////////
//  New storage volume summary table added. (LR - 9/5/05)
/////////////////////////////////////////////////////////
void report_writeStorageStats(TStorageStats storageStats[])
//
//  Input:   storageStats = array of simulation statistics for storage units
//  Output:  none
//  Purpose: writes simulation statistics for storage units to report file.
//
{
    int   j, k, days, hrs, mins;
    float avgVol, maxVol, pctAvgVol, pctMaxVol;

    if ( Nnodes[STORAGE] > 0 )
    {
        WRITE("");
        WRITE("**********************");
        WRITE("Storage Volume Summary");
        WRITE("**********************");
        WRITE("");

/////////////////////////////////////////
//  Table header modified. (LR - 3/10/06)
/////////////////////////////////////////
        fprintf(Frpt.file,
"\n  --------------------------------------------------------------------------------------"
"\n                         Average     Avg       Maximum     Max    Time of Max    Maximum"
"\n                          Volume    Pcnt        Volume    Pcnt     Occurrence    Outflow");
        if ( UnitSystem == US ) fprintf(Frpt.file,
"\n  Storage Unit          1000 ft3    Full      1000 ft3    Full    days hr:min        ");
        else fprintf(Frpt.file,
"\n  Storage Unit           1000 m3    Full       1000 m3    Full    days hr:min        ");
        fprintf(Frpt.file, "%3s", FlowUnitWords[FlowUnits]);
        fprintf(Frpt.file,
"\n  --------------------------------------------------------------------------------------");
        for ( j = 0; j < Nobjects[NODE]; j++ )
        {
            if ( Node[j].type != STORAGE ) continue;
            k = Node[j].subIndex;

            fprintf(Frpt.file, "\n  %-20s", Node[j].ID);   ////Modified (LR - 3/10/06)

            avgVol = storageStats[k].avgVol / StepCount;
            maxVol = storageStats[k].maxVol;
            pctMaxVol = 0.0;
            pctAvgVol = 0.0;
            if ( Node[j].fullVolume > 0.0 )
            {
                pctAvgVol = avgVol / Node[j].fullVolume * 100.0;
                pctMaxVol = maxVol / Node[j].fullVolume * 100.0;
            }
            fprintf(Frpt.file, "%10.3f    %4.0f    %10.3f    %4.0f",
                avgVol*UCF(VOLUME)/1000.0, pctAvgVol, maxVol*UCF(VOLUME)/1000.0, pctMaxVol);
            getElapsedTime(storageStats[k].maxVolDate, &days, &hrs, &mins);
            fprintf(Frpt.file, "    %4d  %02d:%02d  %9.2f",
                days, hrs, mins, storageStats[k].maxFlow*UCF(FLOW));
        }
        WRITE("");
    }
}

//=============================================================================

//////////////////////////////////////////////////////////
//  New outfall loading summary table added. (LR - 7/5/06)
//////////////////////////////////////////////////////////
void report_writeOutfallStats(TOutfallStats outfallStats[], float maxFlow)
//
//  Input:   outfallStats = array of simulation statistics for outfall nodes
//  Output:  none
//  Purpose: writes simulation statistics for outfall nodess to report file.
//
{
    char   units[15];
    int    i, j, k, p;
    float  x;
    float  outfallCount, flowCount;
    float  flowSum, freqSum;
    float* totals = NULL;

    if ( Nnodes[OUTFALL] > 0 )
    {
        // --- initial totals
        totals = (float *) calloc(Nobjects[POLLUT], sizeof(float));
        flowSum = 0.0;
        freqSum = 0.0;

        // --- print table title
        WRITE("");
        WRITE("***********************");
        WRITE("Outfall Loading Summary");
        WRITE("***********************");
        WRITE("");

        // --- print table column headers
        fprintf(Frpt.file, "\n  -----------------------------------------------");
        for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "--------------");
        fprintf(Frpt.file, "\n                        Flow       Avg.      Max.");
        for (p=0; p<Nobjects[POLLUT]; p++) fprintf(Frpt.file,"         Total");
        fprintf(Frpt.file, "\n                        Freq.      Flow      Flow");  
        for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "%14s", Pollut[p].ID);
        fprintf(Frpt.file, "\n  Outfall Node          Pcnt.      %3s       %3s ",
            FlowUnitWords[FlowUnits], FlowUnitWords[FlowUnits]);
        for (p = 0; p < Nobjects[POLLUT]; p++)
        {
            i = UnitSystem;
            if ( Pollut[p].units == COUNT ) i = 2;
            strcpy(units, LoadUnitsWords[i]);
            fprintf(Frpt.file, "%14s", units);
        }
        fprintf(Frpt.file, "\n  -----------------------------------------------");
        for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "--------------");

        // --- identify each outfall node
        for (j=0; j<Nobjects[NODE]; j++)
        {
            if ( Node[j].type != OUTFALL ) continue;
            k = Node[j].subIndex;
            flowCount = outfallStats[k].totalPeriods;

            // --- print node ID, flow freq., avg. flow, & max. flow
            fprintf(Frpt.file, "\n  %-20s", Node[j].ID);
            x = 100.*flowCount/(float)StepCount;
            fprintf(Frpt.file, "%7.2f", x);
            freqSum += x;
            if ( flowCount > 0 )
                x = outfallStats[k].avgFlow*UCF(FLOW)/flowCount;
            else
                x = 0.0;
            flowSum += x;
            fprintf(Frpt.file, "  %8.2f", x);
            fprintf(Frpt.file, "  %8.2f", outfallStats[k].maxFlow*UCF(FLOW));

            // --- print load of each pollutant for outfall
            for (p=0; p<Nobjects[POLLUT]; p++)
            {
                x = outfallStats[k].totalLoad[p];
                if ( totals ) totals[p] += x;
                if ( Pollut[p].units == COUNT ) x = LOG10(x);
                fprintf(Frpt.file, "%14.3f", x); 
            }
        }

        // --- print total outfall loads
        outfallCount = Nnodes[OUTFALL];
        fprintf(Frpt.file, "\n  -----------------------------------------------");
        for (p = 0; p < Nobjects[POLLUT]; p++) fprintf(Frpt.file, "--------------");
        fprintf(Frpt.file, "\n  System              %7.2f  %8.2f  %8.2f",
            freqSum/outfallCount, flowSum, maxFlow*UCF(FLOW));
        if (totals)
        {
            for (p = 0; p < Nobjects[POLLUT]; p++)
            {
                x = totals[p];
                if ( Pollut[p].units == COUNT ) x = LOG10(x);
                fprintf(Frpt.file, "%14.3f", x); 
            }
        }
        WRITE("");
        FREE(totals);
    } 
}

//=============================================================================

void report_writeLinkStats(TLinkStats linkStats[])
//
//  Input:   linkStats = array of simulation statistics for links
//  Output:  none
//  Purpose: writes simulation statistics for links to report file.
//
{

//////////////////////////////////////////////////////////////
//  This routine has undergone major revisions.  (LR - 7/5/06)
//////////////////////////////////////////////////////////////

    int i, j, k, days, hrs, mins;
    float v, fullDepth;

    if ( Nobjects[LINK] == 0 ) return;
    WRITE("");
    WRITE("********************");
    WRITE("Link Flow Summary");
    WRITE("********************");
    WRITE("");
    fprintf(Frpt.file,
"\n  -----------------------------------------------------------------------------------------"
"\n                                 Maximum  Time of Max   Maximum    Max/    Max/       Total"
"\n                                    Flow   Occurrence  Velocity    Full    Full     Minutes");
    if ( UnitSystem == US ) fprintf(Frpt.file,
"\n  Link                 Type          %3s  days hr:min    ft/sec    Flow   Depth  Surcharged",
        FlowUnitWords[FlowUnits]);
    else fprintf(Frpt.file, 
"\n  Link                 Type          %3s  days hr:min     m/sec    Flow   Depth  Surcharged",
        FlowUnitWords[FlowUnits]);
    fprintf(Frpt.file,
"\n  ------------------------------------------------------------------------------------------");

    for ( j = 0; j < Nobjects[LINK]; j++ )
    {
        // --- print link ID
        k = Link[j].subIndex;
        fprintf(Frpt.file, "\n  %-20s", Link[j].ID);

        // --- print link type
        if ( Link[j].xsect.type == DUMMY ) fprintf(Frpt.file, " DUMMY   ");
        else if ( Link[j].xsect.type == IRREGULAR ) fprintf(Frpt.file, " CHANNEL ");
        else fprintf(Frpt.file, " %-7s ", LinkTypeWords[Link[j].type]);

        // --- print max. flow & time of occurrence
        getElapsedTime(linkStats[j].maxFlowDate, &days, &hrs, &mins);
        fprintf(Frpt.file, "%9.2f  %4d  %02d:%02d",
            linkStats[j].maxFlow*UCF(FLOW), days, hrs, mins);

        // --- print max flow / flow capacity & minutes surcharged for pumps
        if ( Link[j].type == PUMP && Link[j].qFull > 0.0)
        {
            fprintf(Frpt.file, "          ");
            fprintf(Frpt.file, "  %6.2f",
                linkStats[j].maxFlow / Link[j].qFull);
            fprintf(Frpt.file, "          %10.0f",
                linkStats[j].timeSurcharged/60.0);
            continue;
        }

        // --- stop printing for dummy conduits
        if ( Link[j].xsect.type == DUMMY ) continue;
        //    || Link[j].type != CONDUIT) continue;

        // --- stop printing for outlet links (since they don't have xsections)
        if ( Link[j].type == OUTLET ) continue;

        // --- print max velocity & max/full flow for conduits
        if ( Link[j].type == CONDUIT )
        {
            v = linkStats[j].maxVeloc*UCF(LENGTH);
            if ( v > 50.0 ) fprintf(Frpt.file, "    >50.00");
            else fprintf(Frpt.file, "   %7.2f", v);
            fprintf(Frpt.file, "  %6.2f", linkStats[j].maxFlow / Link[j].qFull /
                                          (float)Conduit[k].barrels);
        }
        else fprintf(Frpt.file, "                  ");

        // --- print max/full depth
        fullDepth = Link[j].xsect.yFull;
        if ( Link[j].type == ORIFICE &&
             Orifice[k].type == BOTTOM_ORIFICE ) fullDepth = 0.0;
        if ( fullDepth > 0.0 )
        {
            fprintf(Frpt.file, "  %6.2f", linkStats[j].maxDepth / fullDepth); 
        }
        else fprintf(Frpt.file, "        ");

        // --- print minutes surcharged
        fprintf(Frpt.file, "  %10.0f", linkStats[j].timeSurcharged/60.0);
    }
    WRITE("");
    if ( RouteModel != DW ) return;

    WRITE("");
    WRITE("***************************");
    WRITE("Flow Classification Summary");
    WRITE("***************************");
    WRITE("");
    fprintf(Frpt.file,
"\n  -----------------------------------------------------------------------------------------"
"\n                      Adjusted    --- Fraction of Time in Flow Class ----   Avg.     Avg.  "
"\n                       /Actual         Up    Down  Sub   Sup   Up    Down   Froude   Flow  "
"\n  Conduit               Length    Dry  Dry   Dry   Crit  Crit  Crit  Crit   Number   Change"
"\n  -----------------------------------------------------------------------------------------");
    for ( j = 0; j < Nobjects[LINK]; j++ )
    {
        if ( Link[j].type != CONDUIT ) continue;
        if ( Link[j].xsect.type == DUMMY ) continue;
        k = Link[j].subIndex;
        fprintf(Frpt.file, "\n  %-20s", Link[j].ID);
        fprintf(Frpt.file, "  %6.2f ", Conduit[k].modLength / Conduit[k].length);
        for ( i=0; i<MAX_FLOW_CLASSES; i++ )
        {
            fprintf(Frpt.file, "  %4.2f",
                linkStats[j].timeInFlowClass[i] /= StepCount);
        }
        fprintf(Frpt.file, "   %6.2f", linkStats[j].avgFroude / StepCount);
        fprintf(Frpt.file, "   %6.4f", linkStats[j].avgFlowChange /
                                       Link[j].qFull / StepCount);
    }
    WRITE("");
}

//=============================================================================

void report_writeMaxStats(TMaxStats maxMassBalErrs[], TMaxStats maxCourantCrit[],
                          int nMaxStats)
//
//  Input:   maxMassBal[] = nodes with highest mass balance errors
//           maxCourantCrit[] = nodes most often Courant time step critical
//           maxLinkTimes[] = links most often Courant time step critical
//           nMaxStats = number of most critical nodes/links saved
//  Output:  none
//  Purpose: lists nodes & links with highest mass balance errors and 
//           time Courant time step critical
//
{
    int i, j, k;

    if ( RouteModel != DW || Nobjects[LINK] == 0 ) return;

///////////////////////////////////////////////////////////////////
//  Check for at least one element with a max. error. (LR - 9/5/05)
///////////////////////////////////////////////////////////////////
    if ( nMaxStats <= 0 ) return;
    if ( maxMassBalErrs[0].index >= 0 )
    {
        WRITE("");
        WRITE("*************************");
        WRITE("Highest Continuity Errors");    
        WRITE("*************************");
        for (i=0; i<nMaxStats; i++)
        {
            j = maxMassBalErrs[i].index;
            if ( j < 0 ) continue;
            fprintf(Frpt.file, "\n  Node %s (%.2f%%)",
                Node[j].ID, maxMassBalErrs[i].value);
        }
        WRITE("");
    }

    if ( CourantFactor == 0.0 ) return;
    WRITE("");
    WRITE("***************************");
    WRITE("Time-Step Critical Elements");    
    WRITE("***************************");
    k = 0;
    for (i=0; i<nMaxStats; i++)
    {
        j = maxCourantCrit[i].index;
        if ( j < 0 ) continue;
        k++;
        if ( maxCourantCrit[i].objType == NODE )
             fprintf(Frpt.file, "\n  Node %s", Node[j].ID);
        else fprintf(Frpt.file, "\n  Link %s", Link[j].ID);
        fprintf(Frpt.file, " (%.2f%%)", maxCourantCrit[i].value);
    }
    if ( k == 0 ) fprintf(Frpt.file, "\n  None");
    WRITE("");
}

//=============================================================================

void report_writeSysStats(TSysStats* sysStats)
//
//  Input:   sysStats = simulation statistics for overall system
//  Output:  none
//  Purpose: writes simulation statistics for overall system to report file.
//
{
    float x;

    if ( Nobjects[LINK] == 0 || StepCount == 0 ) return;
    WRITE("");
    WRITE("*************************");
    WRITE("Routing Time Step Summary");
    WRITE("*************************");
    fprintf(Frpt.file,
        "\n  Minimum Time Step           :  %7.2f sec",
        sysStats->minTimeStep);
    fprintf(Frpt.file,
        "\n  Average Time Step           :  %7.2f sec",
        sysStats->avgTimeStep / StepCount);
    fprintf(Frpt.file,
        "\n  Maximum Time Step           :  %7.2f sec",
        sysStats->maxTimeStep);
    x = sysStats->steadyStateCount / StepCount * 100.0;
    fprintf(Frpt.file,
        "\n  Percent in Steady State     :  %7.2f", MIN(x, 100.0));
    fprintf(Frpt.file,
        "\n  Average Iterations per Step :  %7.2f",
        sysStats->avgStepCount / StepCount);
    WRITE("");
}


//=============================================================================
//      RAINFALL DATA REPORTING
//=============================================================================

void report_writeRainStats(int i, TRainStats* r)
//
//  Input:   i = rain gage index
//           r = rain file summary statistics
//  Output:  none
//  Purpose: writes summary of rain data read from file to report file.
//
{
    char date1[] = "***********";
    char date2[] = "***********";
    if ( i < 0 )
    {
        WRITE("");
        WRITE("*********************");
        WRITE("Rainfall File Summary");
        WRITE("*********************");
        fprintf(Frpt.file,
"\n  Station    First        Last         Recording   Periods    Periods    Periods");
        fprintf(Frpt.file,
"\n  ID         Date         Date         Frequency    w/Rain    Missing    Malfunc.");
        fprintf(Frpt.file,
"\n  -------------------------------------------------------------------------------\n");
    }
    else
    {
        if ( r->startDate != NO_DATE ) datetime_dateToStr(r->startDate, date1);
        if ( r->endDate   != NO_DATE ) datetime_dateToStr(r->endDate, date2);
        fprintf(Frpt.file, "  %-10s %-11s  %-11s  %5d min    %6d     %6d     %6d\n",
            Gage[i].staID, date1, date2, Gage[i].rainInterval/60,
            r->periodsRain, r->periodsMissing, r->periodsMalfunc);
    }        
}


//=============================================================================
//      RDII REPORTING
//=============================================================================

void report_writeRdiiStats(float rainVol, float rdiiVol)
//
//  Input:   rainVol = total rainfall volume over sewershed
//           rdiiVol = total RDII volume produced
//  Output:  none
//  Purpose: writes summary of RDII inflow to report file.
//
{
    float ratio;
    float ucf1, ucf2;

    ucf1 = UCF(LENGTH) * UCF(LANDAREA);
    if ( UnitSystem == US) ucf2 = MGDperCFS / SECperDAY;
    else                   ucf2 = MLDperCFS / SECperDAY;

    WRITE("");
    fprintf(Frpt.file,
    "\n  **********************           Volume        Volume");
    if ( UnitSystem == US) fprintf(Frpt.file,
    "\n  Rainfall Dependent I/I        acre-feet      Mgallons");
    else fprintf(Frpt.file,
    "\n  Rainfall Dependent I/I        hectare-m       Mliters");
    fprintf(Frpt.file,
    "\n  **********************        ---------     ---------");

    fprintf(Frpt.file, "\n  Sewershed Rainfall ......%14.3f%14.3f",
            rainVol * ucf1, rainVol * ucf2);

    fprintf(Frpt.file, "\n  RDII Produced ...........%14.3f%14.3f",
            rdiiVol * ucf1, rdiiVol * ucf2);

    if ( rainVol == 0.0 ) ratio = 0.0;
    else ratio = rdiiVol / rainVol;
    fprintf(Frpt.file, "\n  RDII Ratio ..............%14.3f", ratio);
    WRITE("");
}


//=============================================================================
//      ERROR REPORTING
//=============================================================================

void report_writeErrorMsg(int code, char* s)
//
//  Input:   code = error code
//           s = error message text
//  Output:  none
//  Purpose: writes error message to report file.
//
{
    if ( Frpt.file )
    {
        WRITE("");
        fprintf(Frpt.file, error_getMsg(code), s);
    }
    ErrorCode = code;
}

//=============================================================================

void report_writeErrorCode()
//
//  Input:   none
//  Output:  none
//  Purpose: writes error message to report file.
//
{
    if ( Frpt.file )
    {
        if ( (ErrorCode >= ERR_MEMORY && ErrorCode <= ERR_TIMESTEP)
        ||   (ErrorCode >= ERR_FILE_NAME && ErrorCode <= ERR_OUT_FILE)
        ||   (ErrorCode == ERR_SYSTEM) )
            fprintf(Frpt.file, error_getMsg(ErrorCode));
    }
}

//=============================================================================

void report_writeInputErrorMsg(int k, int sect, char* line, long lineCount)
//
//  Input:   k = error code
//           sect = number of input data section where error occurred
//           line = line of data containing the error
//           lineCount = line number of data file containing the error
//  Output:  none
//  Purpose: writes input error message to report file.
//
{
    if ( Frpt.file )
    {
        report_writeErrorMsg(k, ErrString);
        if ( sect < 0 ) fprintf(Frpt.file, FMT17, lineCount);
        else            fprintf(Frpt.file, FMT18, lineCount, SectWords[sect]);
        fprintf(Frpt.file, "\n  %s", line);
    }
}

//=============================================================================

