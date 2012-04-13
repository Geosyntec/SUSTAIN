//-----------------------------------------------------------------------------
//   globals.h
//
//   Project: EPA SWMM5
//   Version: 5.0
//   Date:    5/6/05   (Build 5.0.005)
//            3/10/06  (Build 5.0.007)
//   Author:  L. Rossman
//
//   Global Variables
//-----------------------------------------------------------------------------

EXTERN TFile
                  Finp,                     // Input file
                  Fout,                     // Output file
                  Frpt,                     // Report file
                  Fclimate,                 // Climate file
                  Frain,                    // Rainfall file
                  Frunoff,                  // Runoff file
                  Frdii,                    // RDII inflow file
                  Fhotstart1,               // Hotstart input file
                  Fhotstart2,               // Hotstart output file
                  Finflows,                 // Inflows routing file
                  Foutflows;                // Outflows routing file
EXTERN long
                  Nperiods,                 // Number of reporting periods
                  StepCount;                // Number of routing steps used

EXTERN char
                  Msg[MAXMSG+1],            // Text of output message
                  Title[MAXTITLE][MAXMSG+1],// Project title
                  TmpDir[MAXFNAME+1],       // Temporary file directory

				  strTSS[MAXFNAME+1];		

EXTERN TRptFlags
                  RptFlags;                 // Reporting options

EXTERN int
                  Nobjects[MAX_OBJ_TYPES],  // Number of each object type
                  Nnodes[MAX_NODE_TYPES],   // Number of each node sub-type
                  Nlinks[MAX_LINK_TYPES],   // Number of each link sub-type
                  UnitSystem,               // Unit system
                  FlowUnits,                // Flow units
                  InfilModel,               // Infiltration method
                  RouteModel,               // Flow routing method
                  AllowPonding,             // Allow water to pond at nodes
                  InertDamping,             // Degree of inertial damping
                  NormalFlowLtd,            // Use normal flow limiting
                  SlopeWeighting,           // Use slope weighting
                  Compatibility,            // SWMM 5/3/4 compatibility
                  SkipSteadyState,          // Skip over steady state periods
////////////////////////////////////
//  New option added. (LR - 3/10/06)
////////////////////////////////////
                  IgnoreRainfall,           // Ignore rainfall/runoff

                  ErrorCode,                // Error code number
                  WarningCode,              // Warning code number
                  WetStep,                  // Runoff wet time step (sec)
                  DryStep,                  // Runoff dry time step (sec)
                  ReportStep,               // Reporting time step (sec)
                  SweepStart,               // Day of year when sweeping starts
                  SweepEnd;                 // Day of year when sweeping ends

EXTERN float
                  RouteStep,                // Routing time step (sec)
                  LengtheningStep,          // Time step for lengthening (sec)
                  StartDryDays,             // Antecedent dry days
                  CourantFactor,            // Courant time step factor
                  MinSurfArea;              // Minimum nodal surface area

EXTERN double
                  RunoffError,              // Runoff continuity error
                  GwaterError,              // Groundwater continuity error
                  FlowError,                // Flow routing error
                  QualError;                // Quality routing error

EXTERN DateTime
                  StartDate,                // Starting date
                  StartTime,                // Starting time
                  StartDateTime,            // Starting Date+Time
                  EndDate,                  // Ending date
                  EndTime,                  // Ending time
                  EndDateTime,              // Ending Date+Time
                  ReportStartDate,          // Report start date
                  ReportStartTime,          // Report start time
                  ReportStart;              // Report start Date+Time

EXTERN double
                  ReportTime,               // Current reporting time (msec)
                  OldRunoffTime,            // Previous runoff time (msec)
                  NewRunoffTime,            // Current runoff time (msec)
                  OldRoutingTime,           // Previous routing time (msec)
                  NewRoutingTime,           // Current routing time (msec)
                  TotalDuration;            // Simulation duration (msec)

EXTERN TTemp      Temp;                     // Temperature data
EXTERN TEvap      Evap;                     // Evaporation data
EXTERN TWind      Wind;                     // Wind speed data
EXTERN TSnow      Snow;                     // Snow melt data

EXTERN TSnowmelt* Snowmelt;                 // Array of snow melt objects
EXTERN TGage*     Gage;                     // Array of rain gages
EXTERN TSubcatch* Subcatch;                 // Array of subcatchments
EXTERN TAquifer*  Aquifer;                  // Array of groundwater aquifers
EXTERN TUnitHyd*  UnitHyd;                  // Array of unit hydrographs
EXTERN TNode*     Node;                     // Array of nodes
EXTERN TOutfall*  Outfall;                  // Array of outfall nodes
EXTERN TDivider*  Divider;                  // Array of divider nodes
EXTERN TStorage*  Storage;                  // Array of storage nodes
EXTERN TLink*     Link;                     // Array of links
EXTERN TConduit*  Conduit;                  // Array of conduit links
EXTERN TPump*     Pump;                     // Array of pump links
EXTERN TOrifice*  Orifice;                  // Array of orifice links
EXTERN TWeir*     Weir;                     // Array of weir links
EXTERN TOutlet*   Outlet;                   // Array of outlet device links
EXTERN TPollut*   Pollut;                   // Array of pollutants
EXTERN TLanduse*  Landuse;                  // Array of landuses
EXTERN TPattern*  Pattern;                  // Array of time patterns
EXTERN TTable*    Curve;                    // Array of curve tables
EXTERN TTable*    Tseries;                  // Array of time series tables
EXTERN TTransect* Transect;                 // Array of transect data
EXTERN THorton*   HortInfil;                // Horton infiltration data
EXTERN TGrnAmpt*  GAInfil;                  // Green-Ampt infiltration data
EXTERN TCurveNum* CNInfil;                  // Curve No. infiltration data

