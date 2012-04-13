//-----------------------------------------------------------------------------
//   funcs.h
//
//   Project:  EPA SWMM5
//   Version:  5.0
//   Date:     5/6/06   (Build 5.0.005)
//             9/5/05   (Build 5.0.006)
//             3/10/06  (Build 5.0.007)
//             7/5/06   (Build 5.0.008)
//   Author:   L. Rossman
//
//   Global interfacing functions.
//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------
//   moved from SWMM5.cpp
//-----------------------------------------------------------------------------

static float Ucf[9][2] =
      {//  US      SI
      {43200.0,   1097280.0 },         // RAINFALL (in/hr, mm/hr --> ft/sec)
      {12.0,      304.8     },         // RAINDEPTH (in, mm --> ft)
      {1036800.0, 26334720.0},         // EVAPRATE (in/day, mm/day --> ft/sec)
      {1.0,       0.3048    },         // LENGTH (ft, m --> ft)
      {2.2956e-5, 0.92903e-5},         // LANDAREA (ac, ha --> ft2)
      {1.0,       0.02832   },         // VOLUME (ft3, m3 --> ft3)
      {1.0,       1.608     },         // WINDSPEED (mph, km/hr --> mph)
      {1.0,       1.8       },         // TEMPERATURE (deg F, deg C --> deg F)
      {2.203e-6,  1.0e-6    }          // MASS (lb, kg --> mg)
      };
static float Qcf[6] =                  // Flow Conversion Factors:
		{ 1.0,     448.831, 0.64632,   // cfs, gpm, mgd --> cfs
        0.02832, 28.317,  2.4466 };    // cms, lps, mld --> cfs

//moved from link.cpp
static const float MIN_DELTA_Z = 0.001; // minimum elevation change for conduit
                                        // slopes (ft)

//-----------------------------------------------------------------------------
//   Project Manager Methods
//-----------------------------------------------------------------------------
void    project_open(char *f1, char *f2, char *f3);
void    project_close(void);
void    project_readInput(void);
int     project_readOption(char* s1, char* s2);
void    project_validate(void);
int     project_init(void);
int     project_addObject(int type, char* id, int n);
int     project_findObject(int type, char* id);
char*   project_findID(int type, char* id);

// moved from project.cpp
void initPointers(void);
void setDefaults(void);
void openFiles(char *f1, char *f2, char *f3);
void createObjects(void);
void deleteObjects(void);
void createHashTables(void);
void deleteHashTables(void);

///////////////////////////////////////////
//  Argument types modified. (LR - 7/5/06 )
///////////////////////////////////////////
float** project_createMatrix(int nrows, int ncols);

void    project_freeMatrix(float** m);

//-----------------------------------------------------------------------------
//   Input Reader Methods
//-----------------------------------------------------------------------------
int    input_countObjects(void);
int    input_readData(void);

//-----------------------------------------------------------------------------
//   Report Writer Methods
//-----------------------------------------------------------------------------
int    report_readOptions(char* tok[], int ntoks);
void   report_writeLine(char* line);
void   report_writeLogo(void);
void   report_writeTitle(void);
void   report_writeInput(void);
void   report_writeRainStats(int gage, TRainStats* rainStats);
void   report_writeRdiiStats(float totalRain, float totalRdii);
void   report_writeNodeStats(TNodeStats* nodeStats);
void   report_writeLinkStats(TLinkStats* linkStats);
void   report_writeMaxStats(TMaxStats massBalErrs[], TMaxStats CourantCrit[],
       int nMaxStats);
void   report_writeSysStats(TSysStats* sysStats);
void   report_writeReport(void);
void   report_writeControlAction(DateTime aDate, char* linkID, float value,
       char* ruleID);
void   report_writeRunoffError(TRunoffTotals* totals, double area);
void   report_writeLoadingError(TLoadingTotals* totals);
void   report_writeGwaterError(TGwaterTotals* totals, double area);
void   report_writeFlowError(TRoutingTotals* totals);
void   report_writeQualError(TRoutingTotals* totals);
void   report_writeErrorMsg(int code, char* msg);
void   report_writeErrorCode(void);
void   report_writeInputErrorMsg(int k, int sect, char* line, long lineCount);
void   report_writeSysTime(void);

/////////////////////////////////////
//  New argument added. (LR - 7/5/06)
/////////////////////////////////////
void   report_writeSubcatchStats(TSubcatchStats* subcatchStats, float maxRunoff);

/////////////////////////////////////
// New functions added. (LR - 9/5/05)
/////////////////////////////////////
void   report_writeControlActionsHeading(void);
void   report_writeStorageStats(TStorageStats* storageStats);

//////////////////////////////////////
// New functions added, (LR - 7/5/06 )
//////////////////////////////////////
void   report_writeSubcatchLoads(void);
void   report_writeOutfallStats(TOutfallStats* outfallStats, float maxFlow);

//-----------------------------------------------------------------------------
//   Temperature/Evaportation Methods
//-----------------------------------------------------------------------------
int      climate_readParams(char* tok[], int ntoks);
int      climate_readEvapParams(char* tok[], int ntoks);
void     climate_validate(void);
void     climate_openFile(void);
void     climate_initState(void);
void     climate_setState(DateTime aDate);
DateTime climate_getNextEvap(DateTime aDate); 

//-----------------------------------------------------------------------------
//   Rainfall Processing Methods
//-----------------------------------------------------------------------------
void   rain_open(void);
void   rain_close(void);

//-----------------------------------------------------------------------------
//   Snowmelt Processing Methods
//-----------------------------------------------------------------------------
int    snow_readMeltParams(char* tok[], int ntoks);
int    snow_createSnowpack(int subcacth, int snowIndex);
void   snow_initSnowpack(int subcatch);
void   snow_initSnowmelt(int snowIndex);
void   snow_setMeltCoeffs(int snowIndex, float season);
void   snow_plowSnow(int subcatch, float tStep);
float  snow_getSnowMelt(int subcatch, float rainfall, float snowfall,
                        float tStep, float netPrecip[]);
float  snow_getSnowCover(int subcatch);

//-----------------------------------------------------------------------------
//   Runoff Analyzer Methods
//-----------------------------------------------------------------------------
int    runoff_open(void);
void   runoff_execute(void);
void   runoff_close(void);

//-----------------------------------------------------------------------------
//   Conveyance System Routing Methods
//-----------------------------------------------------------------------------
int    routing_open(int routingModel);
float  routing_getRoutingStep(int routingModel, float fixedStep);
void   routing_execute(int routingModel, float routingStep);
void   routing_close(int routingModel);

//-----------------------------------------------------------------------------
//   Output Filer Methods
//-----------------------------------------------------------------------------
int    output_open(void);
void   output_end(void);
void   output_close(void);
void   output_saveResults(double reportTime);
void   output_saveSubcatchResults(double reportTime, FILE* file);
void   output_readDateTime(long period, DateTime *aDate);
void   output_readSubcatchResults(long period, int area);
void   output_readNodeResults(long period, int node);
void   output_readLinkResults(long period, int link);

//-----------------------------------------------------------------------------
//   Infiltration Methods
//-----------------------------------------------------------------------------
int    infil_readParams(int model, char* tok[], int ntoks);
void   infil_initState(int area, int model);
float  infil_getInfil(int area, int model, float tstep, float rainfall,
                      float depth);

int    grnampt_setParams(int j, float p[]);	

//-----------------------------------------------------------------------------
//   Groundwater Methods
//-----------------------------------------------------------------------------
int    gwater_readAquiferParams(int aquifer, char* tok[], int ntoks);
int    gwater_readGroundwaterParams(char* tok[], int ntoks);
void   gwater_validateAquifer(int aquifer);
void   gwater_initState(int subcatch);
void   gwater_getGroundwater(int subcatch, float evap, float* infil,
                             float tStep);
float  gwater_getVolume(int subcatch);

//-----------------------------------------------------------------------------
//   RDII Methods
//-----------------------------------------------------------------------------
int    rdii_readRdiiInflow(char* tok[], int ntoks);
void   rdii_deleteRdiiInflow(int node);
void   rdii_initUnitHyd(int unitHyd);
int    rdii_readUnitHydParams(char* tok[], int ntoks);
void   rdii_openRdii(void);
void   rdii_closeRdii(void);
int    rdii_getNumRdiiFlows(DateTime aDate);
void   rdii_getRdiiFlow(int index, int* node, float* q);


//-----------------------------------------------------------------------------
//   Landuse Methods
//-----------------------------------------------------------------------------
int    landuse_readParams(int landuse, char* tok[], int ntoks);
int    landuse_readPollutParams(int pollut, char* tok[], int ntoks);
int    landuse_readBuildupParams(char* tok[], int ntoks);
int    landuse_readWashoffParams(char* tok[], int ntoks);
float  landuse_getBuildup(int landuse, int pollut, float area, float curb,
       float buildup, float tStep);
float  landuse_getDetached(int landuse,float area, float detstorage,
						   float rainfall, float tStep);

/////////////////////////////////////////
//  Argument list changed. (LR - 7/5/06 )
/////////////////////////////////////////
void  landuse_getWashoff(int landuse, float area, TLandFactor landFactor[],
      float runoff, float tStep, float washoffLoad[]);

void  landuse_getRemoval(int landuse, int subcatch, float area, TLandFactor landFactor[],
      float runoff, float tStep, float removalLoad[]);

////////////////////////////////////
//  Function deleted. (LR - 7/5/06 )
////////////////////////////////////
//float  landuse_getRainLoad(int pollut, float precip, float area, float tStep);

//-----------------------------------------------------------------------------
//   Flow/Quality Routing Methods
//-----------------------------------------------------------------------------
void   flowrout_init(int routingModel);
void   flowrout_close(int routingModel);
float  flowrout_getRoutingStep(int routingModel, float fixedStep);
int    flowrout_execute(int links[], int routingModel, float tStep);

void   qualrout_execute(float tStep);

void   toposort_sortLinks(int links[]);

int    kinwave_execute(int link, float* qin, float* qout, float tStep);

void   dynwave_init(void);
void   dynwave_close(void);
float  dynwave_getRoutingStep(float fixedStep);
int    dynwave_execute(int links[], float tStep);

//-----------------------------------------------------------------------------
// moved from qualrout.cpp
void  findLinkMassFlow(int i);
void  findNodeQual(int j);
void  findLinkQual(int i, float tStep);
void  findStorageQual(int j, float tStep);
void  updateHRT(int j, float v, float q, float tStep);


//-----------------------------------------------------------------------------
//   Treatment Methods
//-----------------------------------------------------------------------------
int    treatmnt_open(void);
void   treatmnt_close(void);
int    treatmnt_readExpression(char* tok[], int ntoks);
void   treatmnt_delete(int node);
void   treatmnt_treat(int node, float q, float v, float tStep);

////////////////////////////////////
// New function added. (LR - 9/5/05)
////////////////////////////////////
void  treatmnt_setInflow(int node, float qIn, float wIn[]);


//-----------------------------------------------------------------------------
//   Mass Balance Methods
//-----------------------------------------------------------------------------
int    massbal_open(void);
void   massbal_close(void);
void   massbal_report(void);
void   massbal_updateRunoffTotals(float vRainfall, float vEvap, float vInfil,
       float vRunoff);
void   massbal_updateLoadingTotals(int type, int pollut, float w);
void   massbal_updateGwaterTotals(float vInfil, float vUpperEvap,
       float vLowerEvap, float vLowerPerc, float vGwater);
void   massbal_updateRoutingTotals(float tStep);
void   massbal_initTimeStepTotals(void);
void   massbal_addInflowFlow(int type, double q);
void   massbal_addInflowQual(int type, int pollut, double w);
void   massbal_addOutflowFlow(double q, int isFlooded);
void   massbal_addOutflowQual(int pollut, double mass, int isFlooded);
void   massbal_addNodeEvap(double evapLoss);
void   massbal_addReactedMass(int pollut, double mass);

//-----------------------------------------------------------------------------
//   Simulation Statistics Methods
//-----------------------------------------------------------------------------
int    stats_open(void);
void   stats_close(void);
void   stats_report(void);
void   stats_updateCriticalTimeCount(int node, int link);
void   stats_updateFlowStats(float tStep, DateTime aDate, int stepCount,
       int steadyState);

/////////////////////////////////////////////////////////
////  New argument added to function (LR - 3/10/06)  ////
/////////////////////////////////////////////////////////
void   stats_updateSubcatchStats(int subcatch, float rainVol, float runonVol,
           float evapVol, float infilVol, float runoffVol, float runoff);

/////////////////////////////////////
//  New function added. (LR - 7/5/06)
/////////////////////////////////////
void  stats_updateMaxRunoff(void);

//-----------------------------------------------------------------------------
//   Raingage Methods
//-----------------------------------------------------------------------------
int      gage_readParams(int gage, char* tok[], int ntoks);
void     gage_initState(int gage);
void     gage_setState(int gage, DateTime aDate);
float    gage_getPrecip(int gage, float *rainfall, float *snowfall);
void     gage_setReportRainfall(int gage, DateTime aDate);
DateTime gage_getNextRainDate(int gage, DateTime aDate);

//-----------------------------------------------------------------------------
//   Subcatchment Methods
//-----------------------------------------------------------------------------
int    subcatch_readParams(int subcatch, char* tok[], int ntoks);
int    subcatch_readSubareaParams(char* tok[], int ntoks);
int    subcatch_readLanduseParams(char* tok[], int ntoks);
int    subcatch_readInitBuildup(char* tok[], int ntoks);
void   subcatch_validate(int subcatch);
void   subcatch_initState(int subcatch);
void   subcatch_setOldState(int subcatch);
void   subcatch_getRunon(int subcatch);
float  subcatch_getRunoff(int subcatch, float tStep);
float  subcatch_getDepth(int subcatch);
void   subcatch_getWashoff(int subcatch, float runoff, float tStep);
void   subcatch_getBuildup(int subcatch, float tStep);
void   subcatch_sweepBuildup(int subcatch, DateTime aDate);
float  subcatch_getWtdOutflow(int subcatch, float wt);
float  subcatch_getWtdWashoff(int subcatch, int pollut, float wt);
void   subcatch_getResults(int subcatch, float wt, float x[]);

float getCstrQual(float c, float v, float wIn, float qNet, float tStep);

//-----------------------------------------------------------------------------
//   Conveyance System Node Methods
//-----------------------------------------------------------------------------
int    node_readParams(int node, int type, int subIndex, char* tok[], int ntoks);
void   node_validate(int node);
void   node_initState(int node);
void   node_setOldHydState(int node);
void   node_setOldQualState(int node);
void   node_initInflow(int node, float tStep);
void   node_setOutletDepth(int node, float yNorm, float yCrit, float z);
void   node_setDividerCutoff(int node, int link);
double node_getSurfArea(int node, float depth);
float  node_getDepth(int node, float volume);
float  node_getVolume(int node, float depth);
float  node_getPondedDepth(int node, float volume);
float  node_getPondedArea(int node, float depth);
float  node_getOutflow(int node, int link);
double node_getEvapLoss(int node, float evapRate, float tStep);
float  node_getMaxOutflow(int node, float q, float tStep);
float  node_getSystemOutflow(int node, int *isFlooded);
void   node_getResults(int node, float wt, float x[]);

//-----------------------------------------------------------------------------
//   Conveyance System Inflow Methods
//-----------------------------------------------------------------------------
int    inflow_readExtInflow(char* tok[], int ntoks);
int    inflow_readDwfInflow(char* tok[], int ntoks);
int    inflow_readDwfPattern(char* tok[], int ntoks);
void   inflow_initDwfPattern(int pattern);
float  inflow_getExtInflow(TExtInflow* inflow, DateTime aDate);
float  inflow_getDwfInflow(TDwfInflow* inflow, int m, int d, int h);
void   inflow_deleteExtInflows(int node);
void   inflow_deleteDwfInflows(int node);

//-----------------------------------------------------------------------------
//   Routing Interface File Methods
//-----------------------------------------------------------------------------
int    iface_readFileParams(char* tok[], int ntoks);
void   iface_openRoutingFiles(void);
void   iface_closeRoutingFiles(void);
int    iface_getNumIfaceNodes(DateTime aDate);
int    iface_getIfaceNode(int index);
float  iface_getIfaceFlow(int index);
float  iface_getIfaceQual(int index, int pollut);
void   iface_saveOutletResults(DateTime reportDate, FILE* file);

//-----------------------------------------------------------------------------
//   Conveyance System Link Methods
//-----------------------------------------------------------------------------
int    link_readParams(int link, int type, int subIndex, char* tok[], int ntoks);
int    link_readXsectParams(char* tok[], int ntoks);
int    link_readLossParams(char* tok[], int ntoks);
void   link_validate(int link);
void   link_initState(int link);
void   link_setOldHydState(int link);
void   link_setOldQualState(int link);
void   link_setFlapGate(int link);
float  link_getInflow(int link);
void   link_setOutfallDepth(int link);
float  link_getYcrit(int link, float q);
float  link_getYnorm(int link, float q);
float  link_getVelocity(int link, float q, float y);
float  link_getFroude(int link, float v, float y);
void   link_getResults(int link, float wt, float x[]);

//-----------------------------------------------------------------------------
//   Link Cross-Section Methods
//-----------------------------------------------------------------------------
int    xsect_isOpen(int type);
int    xsect_setParams(TXsect *xsect, int type, float p[], float ucf);
void   xsect_setIrregXsectParams(TXsect *xsect);
double xsect_getAmax(TXsect* xsect);
double xsect_getSofA(TXsect* xsect, double area);
double xsect_getYofA(TXsect* xsect, double area);
double xsect_getRofA(TXsect* xsect, double area);
double xsect_getAofS(TXsect* xsect, double sFactor);
double xsect_getdSdA(TXsect* xsect, double area);
double xsect_getAofY(TXsect* xsect, double y);
double xsect_getRofY(TXsect* xsect, double y);
double xsect_getWofY(TXsect* xsect, double y);
double xsect_getYcrit(TXsect* xsect, double q);

//-----------------------------------------------------------------------------
//   Cross-Section Transect Methods
//-----------------------------------------------------------------------------
int    transect_create(int n);
void   transect_delete(void);
int    transect_readParams(int* count, char* tok[], int ntoks);
void   transect_validate(int j);

//-----------------------------------------------------------------------------
// moved from transect.cpp
int    setParams(int transect, char* id, float x[]);	
int    setManning(float n[]);							
int    addStation(float x, float y);					
float  getFlow(int k, float a, float wp, int findFlow);	
void   getGeometry(int i, int j, float y);				
void   getSliceGeom(int k, float y, float yu, float yd, float *w,
					float *a, float *wp);

//-----------------------------------------------------------------------------
//   Control Rule Methods
//-----------------------------------------------------------------------------
int    controls_create(int n);
void   controls_delete(void);
int    controls_addRuleClause(int rule, int keyword, char* Tok[], int nTokens);
int    controls_evaluate(DateTime currentTime, DateTime elapsedTime, 
                         double tStep);

//-----------------------------------------------------------------------------
//   Table & Time Series Methods
//-----------------------------------------------------------------------------
int    table_readCurve(char* tok[], int ntoks);
int    table_readTimeseries(char* tok[], int ntoks);
int    table_addEntry(TTable* table, double x, double y);
void   table_deleteEntries(TTable* table);
void   table_init(TTable* table);
int    table_validate(TTable* table);
double table_interpolate(double x, double x1, double y1, double x2, double y2);
double table_lookup(TTable* table, double x);
double table_intervalLookup(TTable* table, double x);
double table_inverseLookup(TTable* table, double y);
int    table_getFirstEntry(TTable* table, double* x, double* y);
int    table_getLastEntry(TTable *table, double *x, double *y);
int    table_getNextEntry(TTable* table, double* x, double* y);
void   table_tseriesInit(TTable *table);
double table_tseriesLookup(TTable* table, double t, char extend);
double table_getArea(TTable* table, double x);
double table_getInverseArea(TTable* table, double a);

/////////////////////////////////////
//  New function added. (LR - 9/5/05)
/////////////////////////////////////
double table_lookupEx(TTable* table, double x);

//-----------------------------------------------------------------------------
//   Utility Methods
//-----------------------------------------------------------------------------
float    UCF(int quantity);                   // units conversion factor
int      getInt(char *s, int *y);             // get integer from string
int      getDouble(char *s, double *y);       // get double from string
int      getFloat(char *s, float *y);         // get float from string
char*    getTmpName(char *s);
int      findmatch(char *s, char *keyword[]); // search for matching keyword
int      match(char *str, char *substr);      // true if substr matches part of str
int      strcomp(char *s1, char *s2);         // case insensitive string compare
char*    sstrncpy(char *dest, const char *src,
         size_t maxlen);                      // safe string copy
void     writecon(char *s);                   // writes string to console
DateTime getDateTime(double elapsedMsec);     // convert elapsed time to date
void     getElapsedTime(DateTime aDate, int* days, int* hrs, int* mins);


