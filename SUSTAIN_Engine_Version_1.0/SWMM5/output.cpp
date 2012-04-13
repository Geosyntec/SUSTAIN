//-----------------------------------------------------------------------------
//   output.c
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/06   (Build 5.0.005)
//   Author:   L. Rossman
//
//   Binary output file access functions.
//-----------------------------------------------------------------------------

#include <stdlib.h>
#include <string.h>
#include "headers.h"

enum InputDataType {INPUT_TYPE_CODE, INPUT_AREA, INPUT_INVERT, INPUT_MAX_DEPTH,
                    INPUT_OFFSET, INPUT_LENGTH};

//-----------------------------------------------------------------------------
//  Shared variables    
//-----------------------------------------------------------------------------
static long      IDStartPos;           // starting file position of ID names
static long      InputStartPos;        // starting file position of input data
static long      OutputStartPos;       // starting file position of output data
static long      BytesPerPeriod;       // bytes saved per simulation time period
static int       NsubcatchResults;     // number of subcacthment output variables
static int       NnodeResults;         // number of node output variables
static int       NlinkResults;         // number of link output variables
static float     SysResults[MAX_SYS_RESULTS];    // values of system output vars.

//-----------------------------------------------------------------------------
//  Exportable variables (shared with report.c)
//-----------------------------------------------------------------------------
float*           SubcatchResults;
float*           NodeResults;
float*           LinkResults;

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static void output_openOutFile(void);
static void output_saveID(char* id, FILE* file);
static void output_saveNodeResults(double reportTime, FILE* file);
static void output_saveLinkResults(double reportTime, FILE* file);

//-----------------------------------------------------------------------------
//  External functions (declared in funcs.h)
//-----------------------------------------------------------------------------
//  output_open                   (called by swmm_start in swmm5.c)
//  output_end                    (called by swmm_end in swmm5.c)
//  output_close                  (called by swmm_close in swmm5.c)
//  output_saveResults            (called by swmm_step in swmm5.c)
//  output_saveSubcatchResults    (called by runoff_saveToFile)
//  output_readDateTime           (called by routines in report.c)
//  output_readSubcatchResults    (called by report_Subcatchments)
//  output_readNodeResults        (called by report_Nodes)
//  output_readLinkResults        (called by report_Links)


//=============================================================================

int output_open()
//
//  Input:   none
//  Output:  returns an error code
//  Purpose: writes basic project data to binary output file.
//
{
    int   nPolluts = Nobjects[POLLUT];
    int   j;
    int   m;
    INT4  k;
    float z;

    // --- open binary output file
    output_openOutFile();
    if ( ErrorCode ) return ErrorCode;

    // --- subcatchment results consist of Rainfall, Snowdepth, Losses, Runoff, 
    //     GW Flow, GW Elev, and Washoff
    NsubcatchResults = MAX_SUBCATCH_RESULTS - 1 + nPolluts;

    // --- node results consist of Depth, Head, Volume, Lateral Inflow,
    //     Total Inflow, Overflow and Quality
    NnodeResults = MAX_NODE_RESULTS - 1 + nPolluts;

    // --- link results consist of Depth, Flow, Velocity, Froude No.,
    //     Capacity and Quality
    NlinkResults = MAX_LINK_RESULTS - 1 + nPolluts;

    BytesPerPeriod = sizeof(DateTime)
        + Nobjects[SUBCATCH] * NsubcatchResults * sizeof(float)
        + Nobjects[NODE] * NnodeResults * sizeof(float)
        + Nobjects[LINK] * NlinkResults * sizeof(float)
        + MAX_SYS_RESULTS * sizeof(float);
    Nperiods = 0;

    SubcatchResults = NULL;
    NodeResults = NULL;
    LinkResults = NULL;
    SubcatchResults = (float *) calloc(NsubcatchResults, sizeof(float));
    NodeResults = (float *) calloc(NnodeResults, sizeof(float));
    LinkResults = (float *) calloc(NlinkResults, sizeof(float));
    if ( !SubcatchResults || !NodeResults || !LinkResults )
    {
        report_writeErrorMsg(ERR_MEMORY, "");
        return ErrorCode;
    }

    fseek(Fout.file, 0, SEEK_SET);
    k = MAGICNUMBER;
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // Magic number
    k = VERSION;
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // Version number
    k = FlowUnits;
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // Flow units
    k = Nobjects[SUBCATCH];
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // # subcatchments
    k = Nobjects[NODE];
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // # nodes
    k = Nobjects[LINK];
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // # links
    k = Nobjects[POLLUT];
    fwrite(&k, sizeof(INT4), 1, Fout.file);   // # pollutants

    // --- save ID names of subcatchments, nodes, links, & pollutants
    IDStartPos = ftell(Fout.file);
    for (j=0; j<Nobjects[SUBCATCH]; j++) output_saveID(Subcatch[j].ID, Fout.file);
    for (j=0; j<Nobjects[NODE];     j++) output_saveID(Node[j].ID, Fout.file);
    for (j=0; j<Nobjects[LINK];     j++) output_saveID(Link[j].ID, Fout.file);
    for (j=0; j<Nobjects[POLLUT];   j++) output_saveID(Pollut[j].ID, Fout.file);

    // --- save codes of pollutant concentration units
    for (j=0; j<Nobjects[POLLUT]; j++)
    {
        k = Pollut[j].units;
        fwrite(&k, sizeof(INT4), 1, Fout.file);
    }

    InputStartPos = ftell(Fout.file);

    // --- save subcatchment area
    k = 1;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_AREA;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (j=0; j<Nobjects[SUBCATCH]; j++)
    {
         SubcatchResults[0] = Subcatch[j].area * UCF(LANDAREA);
         fwrite(&SubcatchResults[0], sizeof(float), 1, Fout.file);
    }

    // --- save node type, invert, & max. depth
    k = 3;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_TYPE_CODE;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_INVERT;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_MAX_DEPTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (j=0; j<Nobjects[NODE]; j++)
    {
        k = Node[j].type;
        NodeResults[0] = Node[j].invertElev * UCF(LENGTH);
        NodeResults[1] = Node[j].fullDepth * UCF(LENGTH);
        fwrite(&k, sizeof(INT4), 1, Fout.file);
        fwrite(NodeResults, sizeof(float), 2, Fout.file);
    }

    // --- save link type, offsets, max. depth, & length
    k = 5;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_TYPE_CODE;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_OFFSET;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_OFFSET;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_MAX_DEPTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = INPUT_LENGTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);

    for (j=0; j<Nobjects[LINK]; j++)
    {
        k = Link[j].type;
        if ( k == PUMP )
        {
            for (m=0; m<4; m++) LinkResults[m] = 0.0;
        }
        else
        {
            LinkResults[0] = Link[j].z1 * UCF(LENGTH);
            LinkResults[1] = Link[j].z2 * UCF(LENGTH);
            if ( Link[j].direction < 0 )
            {
                z = LinkResults[0];
                LinkResults[0] = LinkResults[1];
                LinkResults[1] = z;
            }
            if ( k == OUTLET ) LinkResults[2] = 0.0;
            else LinkResults[2] = Link[j].xsect.yFull * UCF(LENGTH);
            if ( k == CONDUIT )
            {
                m = Link[j].subIndex;
                LinkResults[3] = Conduit[m].length * UCF(LENGTH);
            }
            else LinkResults[3] = 0.0;
        }
        fwrite(&k, sizeof(INT4), 1, Fout.file);
        fwrite(LinkResults, sizeof(float), 4, Fout.file);
    }

    // --- save number & codes of subcatchment result variables
    k = NsubcatchResults;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_RAINFALL;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_SNOWDEPTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_LOSSES;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_RUNOFF;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_GW_FLOW;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = SUBCATCH_GW_ELEV;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (j=0; j<nPolluts; j++)
    {
        k = SUBCATCH_WASHOFF + j;
        fwrite(&k, sizeof(INT4), 1, Fout.file);
    }

    // --- save number & codes of node result variables
    k = NnodeResults;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_DEPTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_HEAD;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_VOLUME;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_LATFLOW;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_INFLOW;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = NODE_OVERFLOW;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (j=0; j<nPolluts; j++)
    {
        k = NODE_QUAL + j;
        fwrite(&k, sizeof(INT4), 1, Fout.file);
    }

    // --- save number & codes of link result variables
    k = NlinkResults;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = LINK_FLOW;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = LINK_DEPTH;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = LINK_VELOCITY;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = LINK_FROUDE;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = LINK_CAPACITY;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (j=0; j<nPolluts; j++)
    {
        k = LINK_QUAL + j;
        fwrite(&k, sizeof(INT4), 1, Fout.file);
    }

    // --- save number & codes of system result variables
    k = MAX_SYS_RESULTS;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    for (k=0; k<MAX_SYS_RESULTS; k++) fwrite(&k, sizeof(INT4), 1, Fout.file);

    // --- save starting report date & report step 
    fwrite(&ReportStart, sizeof(DateTime), 1, Fout.file);
    k = ReportStep;
    if ( fwrite(&k, sizeof(INT4), 1, Fout.file) < 1)
    {
        report_writeErrorMsg(ERR_OUT_WRITE, "");
        return ErrorCode;
    }
    OutputStartPos = ftell(Fout.file);

    // --- check if file will become too large to write a report
    if ( RptFlags.subcatchments != NONE ||
         RptFlags.nodes != NONE ||
         RptFlags.links != NONE )
    {
        if ( (float)OutputStartPos + (float)BytesPerPeriod * TotalDuration
             / 1000.0 / (float)ReportStep >= (float)MAXFILESIZE )
        {
            report_writeErrorMsg(ERR_FILE_SIZE, "");
        }
    }
    return ErrorCode;
}

//=============================================================================

void output_openOutFile()
//
//  Input:   none
//  Output:  none
//  Purpose: opens a project's binary output file.
//
{
    // --- close output file if already opened
    if (Fout.file != NULL) fclose(Fout.file); 

    // --- else if file name supplied then set file mode to SAVE
    else if (strlen(Fout.name) != 0) Fout.mode = SAVE_FILE;

    // --- otherwise set file mode to SCRATCH & generate a name
    else
    {
        Fout.mode = SCRATCH_FILE;
        getTmpName(Fout.name);
    }

    // --- try to open the file
    if ( (Fout.file = fopen(Fout.name, "w+b")) == NULL)
    {
        writecon(FMT14);
        ErrorCode = ERR_OUT_FILE;
    }
}

//=============================================================================

void output_saveResults(double reportTime)
//
//  Input:   reportTime = elapsed simulation time (millisec)
//  Output:  none
//  Purpose: writes computed results for current report time to binary file.
//
{
    int i;
    DateTime reportDate = getDateTime(reportTime);

    if ( reportDate < ReportStart ) return;
    for (i=0; i<MAX_SYS_RESULTS; i++) SysResults[i] = 0.0;
    fwrite(&reportDate, sizeof(DateTime), 1, Fout.file);
    if (Nobjects[SUBCATCH] > 0)
        output_saveSubcatchResults(reportTime, Fout.file);
    if (Nobjects[NODE] > 0)
        output_saveNodeResults(reportTime, Fout.file);
    if (Nobjects[LINK] > 0)
        output_saveLinkResults(reportTime, Fout.file);
    fwrite(SysResults, sizeof(float), MAX_SYS_RESULTS, Fout.file);
    if ( Foutflows.mode == SAVE_FILE )
        iface_saveOutletResults(reportDate, Foutflows.file);
    Nperiods++;
}

//=============================================================================

void output_end()
//
//  Input:   none
//  Output:  none
//  Purpose: writes closing records to binary file.
//
{
    INT4 k;
    fwrite(&IDStartPos, sizeof(INT4), 1, Fout.file);
    fwrite(&InputStartPos, sizeof(INT4), 1, Fout.file);
    fwrite(&OutputStartPos, sizeof(INT4), 1, Fout.file);
    k = Nperiods;
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = error_getCode(ErrorCode);
    fwrite(&k, sizeof(INT4), 1, Fout.file);
    k = MAGICNUMBER;
    if (fwrite(&k, sizeof(INT4), 1, Fout.file) < 1)
    {
        report_writeErrorMsg(ERR_OUT_WRITE, "");
    }
}

//=============================================================================

void output_close()
//
//  Input:   none
//  Output:  none
//  Purpose: frees memory used for accessing the binary file.
//
{
    FREE(SubcatchResults);
    FREE(NodeResults);
    FREE(LinkResults);
}

//=============================================================================

void output_saveID(char* id, FILE* file)
//
//  Input:   id = name of an object
//           file = ptr. to binary output file
//  Output:  none
//  Purpose: writes an object's name to the binary output file.
//
{
    INT4 n = strlen(id);
    fwrite(&n, sizeof(INT4), 1, file);
    fwrite(id, sizeof(char), n, file);
}

//=============================================================================

void output_saveSubcatchResults(double reportTime, FILE* file)
//
//  Input:   reportTime = elapsed simulation time (millisec)
//           file = ptr. to binary output file
//  Output:  none
//  Purpose: writes computed subcatchment results to binary file.
//
{
    int      j;
    double   f;
    float    area;
    float    totalArea = 0.0; 
    DateTime reportDate = getDateTime(reportTime);

    // --- update reported rainfall at each rain gage
    for ( j=0; j<Nobjects[GAGE]; j++ )
    {
        gage_setReportRainfall(j, reportDate);
    }

    // --- find where current reporting time lies between latest runoff times
    f = (reportTime - OldRunoffTime) / (NewRunoffTime - OldRunoffTime);

    // --- write subcatchment results to file
    for ( j=0; j<Nobjects[SUBCATCH]; j++)
    {
        // --- retrieve interpolated results for reporting time & write to file
        subcatch_getResults(j, f, SubcatchResults);
        fwrite(SubcatchResults, sizeof(float), NsubcatchResults, file);

        // --- update system-wide results
        area = Subcatch[j].area * UCF(LANDAREA);
        totalArea += area;
        SysResults[SYS_RAINFALL] += SubcatchResults[SUBCATCH_RAINFALL] * area;
        SysResults[SYS_SNOWDEPTH] += SubcatchResults[SUBCATCH_SNOWDEPTH] * area;
        SysResults[SYS_LOSSES] += SubcatchResults[SUBCATCH_LOSSES] * area;
        SysResults[SYS_RUNOFF] += SubcatchResults[SUBCATCH_RUNOFF];
    }

    // --- normalize system-wide results to catchment area
    if ( UnitSystem == SI )
       SysResults[SYS_TEMPERATURE]  =  (5./9.) * (Temp.ta - 32.0);
    else SysResults[SYS_TEMPERATURE] = Temp.ta;
    SysResults[SYS_RAINFALL]  /= totalArea;
    SysResults[SYS_SNOWDEPTH] /= totalArea;
    SysResults[SYS_LOSSES]    /= totalArea;
}

//=============================================================================

void output_saveNodeResults(double reportTime, FILE* file)
//
//  Input:   reportTime = elapsed simulation time (millisec)
//           file = ptr. to binary output file
//  Output:  none
//  Purpose: writes computed node results to binary file.
//
{
    extern TRoutingTotals StepFlowTotals;  // defined in massbal.c
    int j;

    // --- find where current reporting time lies between latest routing times
    double f = (reportTime - OldRoutingTime) /
               (NewRoutingTime - OldRoutingTime);

    // --- write node results to file
    for (j=0; j<Nobjects[NODE]; j++)
    {
        // --- retrieve interpolated results for reporting time & write to file
        node_getResults(j, f, NodeResults);
        fwrite(NodeResults, sizeof(float), NnodeResults, file);

        // --- update system-wide results
        SysResults[SYS_FLOODING] += NodeResults[NODE_OVERFLOW];
        SysResults[SYS_STORAGE] += NodeResults[NODE_VOLUME];
        if ( Node[j].degree == 0 ) 
        {
            SysResults[SYS_OUTFLOW] += NodeResults[NODE_INFLOW];
        }
    }
    SysResults[SYS_DWFLOW] = StepFlowTotals.dwInflow * UCF(FLOW);
    SysResults[SYS_GWFLOW] = StepFlowTotals.gwInflow * UCF(FLOW);
    SysResults[SYS_IIFLOW] = StepFlowTotals.iiInflow * UCF(FLOW);
    SysResults[SYS_EXFLOW] = StepFlowTotals.exInflow * UCF(FLOW);
    SysResults[SYS_INFLOW] = SysResults[SYS_RUNOFF] +
                             SysResults[SYS_DWFLOW] +
                             SysResults[SYS_GWFLOW] +
                             SysResults[SYS_IIFLOW] +
                             SysResults[SYS_EXFLOW];
}

//=============================================================================

void output_saveLinkResults(double reportTime, FILE* file)
//
//  Input:   reportTime = elapsed simulation time (millisec)
//           file = ptr. to binary output file
//  Output:  none
//  Purpose: writes computed link results to binary file.
//
{
    int j;

    // --- find where current reporting time lies between latest routing times
    double f = (reportTime - OldRoutingTime) /
               (NewRoutingTime - OldRoutingTime);

    // --- write link results to file
    for (j=0; j<Nobjects[LINK]; j++)
    {
        // --- retrieve interpolated results for reporting time & write to file
        link_getResults(j, f, LinkResults);
        fwrite(LinkResults, sizeof(float), NlinkResults, file);

        // --- update system-wide results
        SysResults[SYS_STORAGE] += ((1.0-f)*Link[j].oldVolume +
                                   f*Link[j].newVolume) * UCF(VOLUME);
    }
}

//=============================================================================

void output_readDateTime(long period, DateTime* days)
//
//  Input:   period = index of reporting time period
//  Output:  days = date/time value
//  Purpose: retrieves the date/time for a specific reporting period
//           from the binary output file.
//
{
    long bytePos = OutputStartPos + (period-1)*BytesPerPeriod;
    fseek(Fout.file, bytePos, SEEK_SET);
    *days = NO_DATE;
    fread(days, sizeof(DateTime), 1, Fout.file);
}

//=============================================================================

void output_readSubcatchResults(long period, int index)
//
//  Input:   period = index of reporting time period
//           index = subcatchment index
//  Output:  none
//  Purpose: reads computed results for a subcatchment at a specific time
//           period.
//
{
    long bytePos = OutputStartPos + (period-1)*BytesPerPeriod;
    bytePos += sizeof(DateTime) + index*NsubcatchResults*sizeof(float);
    fseek(Fout.file, bytePos, SEEK_SET);
    fread(SubcatchResults, sizeof(float), NsubcatchResults, Fout.file);
}

//=============================================================================

void output_readNodeResults(long period, int index)
//
//  Input:   period = index of reporting time period
//           index = node index
//  Output:  none
//  Purpose: reads computed results for a node at a specific time period.
//
{
    long bytePos = OutputStartPos + (period-1)*BytesPerPeriod;
    bytePos += sizeof(DateTime) + Nobjects[SUBCATCH]*
               NsubcatchResults*sizeof(float);
    bytePos += index*NnodeResults*sizeof(float);
    fseek(Fout.file, bytePos, SEEK_SET);
    fread(NodeResults, sizeof(float), NnodeResults, Fout.file);
}

//=============================================================================

void output_readLinkResults(long period, int index)
//
//  Input:   period = index of reporting time period
//           index = link index
//  Output:  none
//  Purpose: reads computed results for a link at a specific time period.
//
{
    long bytePos = OutputStartPos + (period-1)*BytesPerPeriod;
    bytePos += sizeof(DateTime) + Nobjects[SUBCATCH]*
               NsubcatchResults*sizeof(float);
    bytePos += Nobjects[NODE]*NnodeResults*sizeof(float);
    bytePos += index*NlinkResults*sizeof(float);
    fseek(Fout.file, bytePos, SEEK_SET);
    fread(LinkResults, sizeof(float), NlinkResults, Fout.file);
    fread(SysResults, sizeof(float), MAX_SYS_RESULTS, Fout.file);
}

//=============================================================================
