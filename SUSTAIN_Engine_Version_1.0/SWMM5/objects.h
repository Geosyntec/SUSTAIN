//-----------------------------------------------------------------------------
//   objects.h
//
//   Project: EPA SWMM5
//   Version: 5.0
//   Date:    5/6/05   (Build 5.0.005)
//            9/5/05   (Build 5.0.006)
//            3/10/06  (Build 5.0.007)
//            7/5/06   (Build 5.0.008)
//   Author:  L. Rossman (EPA)
//            R. Dickinson (CDM)
//
//   Definitions of data structures.
//
//   Most SWMM 5 "objects" are represented as C data structures.
//
//   NOTE: the units shown next to each structure's properties are SWMM's
//         internal units and may be different than the units required
//         for the property as it appears in the input file.
//
//   NOTE: in many structure definitions, a blank line separates the set of
//         input properties from the set of computed output properties.
//-----------------------------------------------------------------------------

#include "mathexpr.h"

//----------------------------------
// DEFINITION OF REAL AND INT4 TYPES
//----------------------------------
typedef  double        REAL;
typedef  long          INT4;


//-----------------------------------------
// LINKED LIST ENTRY FOR TABLES/TIME SERIES
//-----------------------------------------
struct  TableEntry
{
   REAL    x;
   REAL    y;
   struct  TableEntry* next;
};
typedef struct TableEntry TTableEntry;


//-------------------------
// CURVE/TIME SERIES OBJECT
//-------------------------
typedef struct
{
   char*         ID;              // Table/time series ID
   int           curveType;       // type of curve tabulated
   REAL          lastDate;        // last input date for time series
   REAL          x1, x2;          // current bracket on x-values
   REAL          y1, y2;          // current bracket on y-values
   TTableEntry*  firstEntry;      // first data point
   TTableEntry*  lastEntry;       // last data point
   TTableEntry*  thisEntry;       // current data point
}  TTable;


//-----------------
// RAIN GAGE OBJECT
//-----------------
typedef struct
{
   char*         ID;              // raingage name
   int           dataSource;      // data from time series or file 
   int           tSeries;         // rainfall data time series index
   char          fname[MAXFNAME+1]; // name of rainfall data file
   char          staID[MAXMSG+1]; // station number
   DateTime      startFileDate;   // starting date of data read from file
   DateTime      endFileDate;     // ending date of data read from file
   int           rainType;        // intensity, volume, cumulative
   int           rainInterval;    // recording time interval (seconds)
   int           rainUnits;       // rain depth units (US or SI)
   float         snowFactor;      // snow catch deficiency correction

   long          startFilePos;    // starting byte position in Rain file
   long          endFilePos;      // ending byte position in Rain file
   long          currentFilePos;  // current byte position in Rain file
   float         rainAccum;       // cumulative rainfall
   float         unitsFactor;     // units conversion factor (to inches or mm)
   DateTime      startDate;       // start date of current rainfall
   DateTime      endDate;         // end date of current rainfall
   DateTime      nextDate;        // next date with recorded rainfall
   float         rainfall;        // current rainfall (in/hr or mm/hr)
   float         nextRainfall;    // next rainfall (in/hr or mm/hr)
   float         reportRainfall;  // rainfall value used for reported results
   int           coGage;          // index of gage with same rain timeseries
   int           isUsed;          // TRUE if gage used by any subcatchment
}  TGage;


//-------------------
// TEMPERATURE OBJECT
//-------------------
typedef struct
{
   int           dataSource;      // data from time series or file 
   int           tSeries;         // temperature data time series index
   DateTime      fileStartDate;   // starting date of data read from file
   float         elev;            // elev. of study area (ft)
   float         anglat;          // latitude (degrees)
   float         dtlong;          // longitude correction (hours)

   float         ta;              // air temperature (deg F)
   float         tmax;            // previous day's max. temp. (deg F)
   float         ea;              // saturation vapor pressure (in Hg)
   float         gamma;           // psychrometric constant
   float         tanAnglat;       // tangent of latitude angle
}  TTemp;


//-----------------
// WINDSPEED OBJECT
//-----------------
typedef struct
{
   int          type;             // monthly or file data
   float        aws[12];          // monthly avg. wind speed (mph)

   float         ws;              // wind speed (mph)
}  TWind;


//------------
// SNOW OBJECT
//------------
typedef struct
{
   float         snotmp;          // temp. dividing rain from snow (deg F)
   float         tipm;            // antecedent temp. index parameter
   float         rnm;             // ratio of neg. melt to melt coeff.
   float         adc[2][10];      // areal depletion curves

   float         season;          // snowmelt season
   double        removed;         // total snow plowed out of system (ft3)
}  TSnow;


//-------------------
// EVAPORATION OBJECT
//-------------------
typedef struct
{
    int          type;            // type of evaporation data
    int          tSeries;         // time series index
    float        monthlyEvap[12]; // monthly evaporation values
    float        panCoeff[12];    // monthly pan coeff. values

    float        rate;            // current evaporation rate (ft/sec)
}   TEvap;


//---------------------
// HORTON INFILTRATION
//---------------------
typedef struct
{
   float         fmin;            // minimum infil. rate (ft/sec)
   float         Fmax;            // maximum total infiltration (ft);
   float         decay;           // decay coeff. of infil. rate (1/sec)
   float         regen;           // regeneration coeff. of infil. rate (1/sec)

   float         tp;              // present time on infiltration curve (sec)
   float         f0;              // initial infil. rate (ft/sec)
}  THorton;


//-------------------------
// GREEN-AMPT INFILTRATION
//-------------------------
typedef struct
{
   float         S;               // avg. capillary suction (ft)
   float         Ks;              // saturated conductivity (ft/sec)
   float         IMDmax;          // max. soil moisture deficit (ft/ft)

   float         IMD;             // current soil moisture deficit
   float         F;               // current cumulative infiltration (ft)
   float         T;               // time needed to drain upper zone (sec)
   float         L;               // depth of upper soil zone (ft)
   float         FU;              // current moisture content of upper zone (ft)
   float         FUmax;           // saturated moisture content of upper zone (ft)
   char          Sat;             // saturation flag
}  TGrnAmpt;


//------------------------------
// SCS CURVE NUMBER INFILTRATION
//------------------------------
typedef struct
{
   float         Smax;            // max. infiltration capacity (ft)
   float         regen;           // infil. capacity regeneration constant (1/sec)
   float         Tmax;            // maximum inter-event time (sec)

   float         S;               // current infiltration capacity (ft)
   float         F;               // current cumulative infiltration (ft)
   float         P;               // current cumulative precipitation (ft)
   float         T;               // current inter-event time (sec)
   float         Se;              // current event infiltration capacity (ft)
   float         f;               // previous infiltration rate (ft/sec)

}  TCurveNum;


//-------------------
// AQUIFER OBJECT
//-------------------
typedef struct
{
    char*       ID;               // aquifer name
    float       porosity;         // soil porosity
    float       wiltingPoint;     // soil wilting point
    float       fieldCapacity;    // soil field capacity
    float       conductivity;     // soil hyd. conductivity (ft/sec)
    float       conductSlope;     // slope of conductivity v. moisture curve
    float       tensionSlope;     // slope of tension v. moisture curve
    float       upperEvapFrac;    // evaporation available in upper zone
    float       lowerEvapDepth;   // evap depth existing in lower zone (ft)
    float       lowerLossCoeff;   // coeff. for losses to deep GW (ft/sec)
    float       bottomElev;       // elevation of bottom of aquifer (ft)
    float       waterTableElev;   // initial water table elevation (ft)
    float       upperMoisture;    // initial moisture content of unsat. zone
/////////////////////////////////////////////////////////////////////////////
//  new parameter for macropores.
/////////////////////////////////////////////////////////////////////////////
    float       macroporosity;    // soil porosity in macropores (large pores for gravitational water)
}   TAquifer;


//------------------------
// GROUNDWATER OBJECT
//------------------------
typedef struct
{
    int           aquifer;        // index of associated gw aquifer 
    int           node;           // index of node receiving gw flow
    float         surfElev;       // elevation of ground surface (ft)
    float         a1, b1;         // ground water outflow coeff. & exponent
    float         a2, b2;         // surface water outflow coeff. & exponent
    float         a3;             // surf./ground water interaction coeff.
    float         fixedDepth;     // fixed surface water water depth (ft)

//////////////////////////////////////
//  New attribute added. (LR - 9/5/05)
//////////////////////////////////////
    float         nodeElev;       // elevation of receiving node invert (ft)

//////////////////////////////////////
//added
//////////////////////////////////////
    float         theta_mac;      // upper zone moisture content in macropores

    float         theta;          // upper zone moisture content
    float         lowerDepth;     // depth of saturated zone (ft)
    float         oldFlow;        // gw outflow from previous time period (cfs)
    float         newFlow;        // gw outflow from current time period (cfs)
} TGroundwater;


//----------------
// SNOWMELT OBJECT
//----------------
// Snowmelt objects contain parameters that describe the melting
// process of snow packs on 3 different types of surfaces:
//   1 - plowable impervious area
//   2 - non-plowable impervious area
//   3 - pervious area
typedef struct
{
   char*         ID;              // snowmelt parameter set name
   float         snn;             // fraction of impervious area plowable
   float         si[3];           // snow depth for 100% cover
   float         dhmin[3];        // min. melt coeff. for each surface (ft/sec-F)
   float         dhmax[3];        // max. melt coeff. for each surface (ft/sec-F)
   float         tbase[3];        // base temp. for melting (F)
   float         fwfrac[3];       // free water capacity / snow depth
   float         wsnow[3];        // initial snow depth on each surface (ft)
   float         fwnow[3];        // initial free water in snow pack (ft)
   float         weplow;          // depth at which plowing begins (ft)
   float         sfrac[5];        // fractions moved to other areas by plowing
   int           toSubcatch;      // index of subcatch receiving plowed snow

   float         dhm[3];          // melt coeff. for each surface (ft/sec-F)
}  TSnowmelt;


//----------------
// SNOWPACK OBJECT
//----------------
// Snowpack objects describe the state of the snow melt process on each
// of 3 types of snow surfaces.
typedef struct
{
   int           snowmeltIndex;   // index of snow melt parameter set
   float         fArea[3];        // fraction of total area of each surface
   float         wsnow[3];        // depth of snow pack (ft)
   float         fw[3];           // depth of free water in snow pack (ft)
   float         coldc[3];        // cold content of snow pack
   float         ati[3];          // antecedent temperature index (deg F)
   float         sba[3];          // initial ASC of linear ADC
   float         awe[3];          // initial AWESI of linear ADC
   float         sbws[3];         // final AWESI of linear ADC
   float         imelt[3];        // immediate melt (ft)
}  TSnowpack;


//---------------
// SUBAREA OBJECT
//---------------
// An array of 3 subarea objects is associated with each subcatchment object.
// They describe the runoff process on 3 types of surfaces:
//   1 - impervious with no depression storage
//   2 - impervious with depression storage
//   3 - pervious
typedef struct
{
   int           routeTo;         // code indicating where outflow is sent
   float         fOutlet;         // fraction of outflow to outlet
   float         N;               // Manning's n
   float         fArea;           // fraction of total area
   float         dStore;          // depression storage (ft)

   float         alpha;           // overland flow factor
   float         inflow;          // inflow rate (ft/sec)
   float         runoff;          // runoff rate (ft/sec)
   float         depth;           // depth of surface runoff (ft)
}  TSubarea;


//-------------------------
// LAND AREA LANDUSE FACTOR
//-------------------------
typedef struct
{
   float         fraction;        // fraction of land area with land use
   float*        buildup;         // array of buildups for each pollutant

   //added
   float*		 detstorage;	  // array of detached storage for sediment
   
   DateTime      lastSwept;       // date/time of last street sweeping
}  TLandFactor;


//--------------------
// SUBCATCHMENT OBJECT
//--------------------
typedef struct
{
   char*         ID;              // subcatchment name
   char          rptFlag;         // reporting flag
   int           gage;            // raingage index
   int           outNode;         // outlet node index
   int           outSubcatch;     // outlet subcatchment index
   int           infil;           // infiltration object index
   TSubarea      subArea[3];      // sub-area data
   float         width;           // overland flow width (ft)
   float         area;            // area (ft2)
   float         fracImperv;      // fraction impervious
   float         slope;           // slope (ft/ft)
   float         curbLength;      // total curb length (ft)
   float*        initBuildup;     // initial pollutant buildup (mass/ft2)
   TLandFactor*  landFactor;      // array of land use factors
   TGroundwater* groundwater;     // associated groundwater data
   TSnowpack*    snowpack;        // associated snow pack data

   float         rainfall;        // current rainfall (ft/sec)
   float         losses;          // current infil + evap losses (ft/sec)

////////////////////////////////////////////////////////
//  This property is never used anywhere. (LR - 7/5/06 )
////////////////////////////////////////////////////////
//   float         depth;           // depth of surface water (ft)

   float         runon;           // runon from other subcatchments (cfs)
   float         oldRunoff;       // previous runoff (cfs)
   float         newRunoff;       // current runoff (cfs)
   float         oldSnowDepth;    // previous snow depth (ft)
   float         newSnowDepth;    // current snow depth (ft)
   float*        oldQual;         // previous runoff quality (mass/L)
   float*        newQual;         // current runoff quality (mass/L)

////////////////////////////////////////////////////////////
//  New properties added for washoff routing. (LR - 7/5/06 )
////////////////////////////////////////////////////////////
   float*        pondedQual;      // ponded surface water quality (mass/ft3)
   float*        totalLoad;       // total washoff load (lbs or kg)
}  TSubcatch;


//-----------------------
// TIME PATTERN DATA
//-----------------------
typedef struct
{
   char*        ID;               // time pattern name
   int          type;             // time pattern type code
   int          count;            // number of factors
   float        factor[24];       // time pattern factors
}  TPattern;


//------------------------------
// DIRECT EXTERNAL INFLOW OBJECT
//------------------------------
struct ExtInflow
{
   int            param;         // pollutant index (flow = -1)
   int            type;          // CONCEN or MASS
   int            tSeries;       // index of inflow time series
   float          cFactor;       // units conversion factor for mass inflow

//////////////////////////////////////////
////  New attributes added. (LR - 7/5/06 )
//////////////////////////////////////////
   float          baseline;      // constant baseline value
   float          sFactor;       // time series scaling factor

   struct ExtInflow* next;       // pointer to next inflow data object
};
typedef struct ExtInflow TExtInflow;


//-------------------------------
// DRY WEATHER FLOW INFLOW OBJECT
//-------------------------------
struct DwfInflow
{
   int            param;          // pollutant index (flow = -1)
   float          avgValue;       // average value (cfs or concen.)
   int            patterns[4];    // monthly, daily, hourly, weekend time patterns
   struct DwfInflow* next;        // pointer to next inflow data object
};
typedef struct DwfInflow TDwfInflow;


//-------------------
// RDII INFLOW OBJECT
//-------------------
typedef struct
{
   int           unitHyd;         // index of unit hydrograph
   float         area;            // area of sewershed (ft2)
}  TRdiiInflow;


//-----------------------------
// UNIT HYDROGRAPH GROUP OBJECT
//-----------------------------
typedef struct
{
   char*         ID;              // name of the unit hydrograph object
   int           rainGage;        // index of rain gage
   float         r[12][3];        // fraction of rainfall becoming I&I
   long          tBase[12][3];    // time base of each UH in each month (sec)
   long          tPeak[12][3];    // time to peak of each UH in each month (sec).
}  TUnitHyd;


//-----------------
// TREATMENT OBJECT
//-----------------
typedef struct
{
    int          treatType;       // treatment equation type: REMOVAL/CONCEN
    ExprTree*    equation;        // removal eqn. stored as expression tree
} TTreatment;


//------------
// NODE OBJECT
//------------
typedef struct
{
   char*         ID;              // node ID
   int           type;            // node type code
   int           subIndex;        // index of node's sub-category
   char          rptFlag;         // reporting flag
   float         invertElev;      // invert elevation (ft)
   float         initDepth;       // initial storage level (ft)
   float         fullDepth;       // dist. from invert to surface (ft)
   float         surDepth;        // added depth under surcharge (ft)
   float         pondedArea;      // area filled by ponded water (ft2)
   TExtInflow*   extInflow;       // pointer to external inflow data
   TDwfInflow*   dwfInflow;       // pointer to dry weather flow inflow data
   TRdiiInflow*  rdiiInflow;      // pointer to RDII inflow data
   TTreatment*   treatment;       // array of treatment data

   int           degree;          // number of outflow links
   char          updated;         // true if state has been updated
   float         crownElev;       // top of highest connecting conduit (ft)
   float         inflow;          // total inflow (cfs)
   float         outflow;         // total outflow (cfs)
   float         oldVolume;       // previous volume (ft3)
   float         newVolume;       // current volume (ft3)
   float         fullVolume;      // max. storage available (ft3)
   float         overflow;        // overflow rate (cfs)
   float         oldDepth;        // previous water depth (ft)
   float         newDepth;        // current water depth (ft)
   float         oldLatFlow;      // previous lateral inflow (cfs)
   float         newLatFlow;      // current lateral inflow (cfs)
   float*        oldQual;         // previous quality state
   float*        newQual;         // current quality state
   float         oldFlowInflow;   // previous flow inflow
   float         oldNetInflow;    // previous net inflow
}  TNode;


//---------------
// OUTFALL OBJECT
//---------------
typedef struct
{
   int        type;               // outfall type code
   char       hasFlapGate;        // true if contains flap gate
   float      fixedStage;         // fixed outfall stage (ft)
   int        tideCurve;          // index of tidal stage curve
   int        stageSeries;        // index of outfall stage time series
}  TOutfall;


//--------------------
// STORAGE UNIT OBJECT
//--------------------
typedef struct
{
   float       fEvap;             // fraction of evaporation realized
   float       aConst;            // surface area at zero height (ft2)
   float       aCoeff;            // coeff. of area v. height curve
   float       aExpon;            // exponent of area v. height curve
   int         aCurve;            // index of tabulated area v. height curve

   float       hrt;               // hydraulic residence time (sec)
}  TStorage;


//--------------------
// FLOW DIVIDER OBJECT
//--------------------
typedef struct
{
   int         link;              // index of link with diverted flow
   int         type;              // divider type code
   float       qMin;              // minimum inflow for diversion (cfs)
   float       qMax;              // flow when weir is full (cfs)
   float       dhMax;             // height of weir (ft)
   float       cWeir;             // weir discharge coeff.
   int         flowCurve;         // index of inflow v. diverted flow curve
}  TDivider;


//-----------------------------
// CROSS SECTION DATA STRUCTURE
//-----------------------------
typedef struct
{
   int           type;            // type code of cross section shape
   int           transect;        // index of transect (if applicable)
   float         yFull;           // depth when full (ft)
   float         wMax;            // width at widest point (ft)
   float         aFull;           // area when full (ft2)
   float         rFull;           // hyd. radius when full (ft)
   float         sFull;           // section factor when full (ft^4/3)
   float         sMax;            // section factor at max. flow (ft^4/3)

   // These variables have different meanings depending on section shape
   float         yBot;            // depth of bottom section
   float         aBot;            // area of bottom section
   float         sBot;            // slope of bottom section
   float         rBot;            // radius of bottom section
}  TXsect;


//--------------------------------------
// CROSS SECTION TRANSECT DATA STRUCTURE
//--------------------------------------
#define  N_TRANSECT_TBL  26       // size of transect geometry tables
typedef struct
{
    char*        ID;                        // section ID
    float        yFull;                     // depth when full (ft)
    float        aFull;                     // area when full (ft2)
    float        rFull;                     // hyd. radius when full (ft)
    float        wMax;                      // width at widest point (ft)

////////////////////////////////////////
//  New attributes added. (LR - 3/10/06)
////////////////////////////////////////
    float        sMax;                      // section factor at max. flow (ft^4/3)
    float        aMax;                      // area at max. flow (ft2)

    float        roughness;                 // Manning's n
    double       areaTbl[N_TRANSECT_TBL];   // table of area v. depth
    double       hradTbl[N_TRANSECT_TBL];   // table of hyd. radius v. depth
    double       widthTbl[N_TRANSECT_TBL];  // table of top width v. depth
    int          nTbl;                      // size of geometry tables
}   TTransect;


//------------
// LINK OBJECT
//------------
typedef struct
{
   char*         ID;              // link ID
   int           type;            // link type code
   int           subIndex;        // index of link's sub-category
   char          rptFlag;         // reporting flag
   int           node1;           // start node index
   int           node2;           // end node index
   float         z1;              // upstrm invert ht. above node invert (ft)
   float         z2;              // downstrm invert ht. above node invert (ft)
   TXsect        xsect;           // cross section data
   float         q0;              // initial flow (cfs)
   float         qLimit;          // constraint on max. flow (cfs)
   float         cLossInlet;      // inlet loss coeff.
   float         cLossOutlet;     // outlet loss coeff.
   float         cLossAvg;        // avg. loss coeff.
   int           hasFlapGate;     // true if flap gate present

   float         oldFlow;         // previous flow rate (cfs)
   float         newFlow;         // current flow rate (cfs)
   float         oldDepth;        // previous flow depth (ft)
   float         newDepth;        // current flow depth (ft)
   float         oldVolume;       // previous flow volume (ft3)
   float         newVolume;       // current flow volume (ft3)
   float         qFull;           // flow when full (cfs)
   float         setting;         // control setting
   float         froude;          // Froude number
   float*        oldQual;         // previous quality state
   float*        newQual;         // current quality state
   int           flowClass;       // flow classification
   float         dqdh;            // change in flow w.r.t. head (ft2/sec)
   signed char   direction;       // flow direction flag
   char          isClosed;        // flap gate closed flag
}  TLink;


//---------------
// CONDUIT OBJECT
//---------------
typedef struct
{
   float         length;          // conduit length (ft)
   float         roughness;       // Manning's n
   char          barrels;         // number of barrels

   float         modLength;       // modified conduit length (ft)
   float         roughFactor;     // roughness factor for DW routing
   float         slope;           // slope
   float         beta;            // discharge factor
   float         qMax;            // max. flow (cfs)
   float         a1, a2;          // upstream & downstream areas (ft2)
   float         q1, q2;          // upstream & downstream flows per barrel (cfs)
   float         q1Old, q2Old;    // previous values of q1 & q2 (cfs)
   char          superCritical;   // super-critical flow flag
   char          hasLosses;       // local losses flag
}  TConduit;


//------------
// PUMP OBJECT
//------------
typedef struct
{
   int           type;            // pump type
   int           pumpCurve;       // pump curve table index
}  TPump;


//---------------
// ORIFICE OBJECT
//---------------
typedef struct
{
   int           type;            // orifice type code
   int           shape;           // orifice shape code
   float         cDisch;          // discharge coeff.

   float         cFull;           // discharge / area when full
   float         length;          // equivalent length (ft)
   float         surfArea;        // equivalent surface area (ft2) 
}  TOrifice;


//------------
// WEIR OBJECT
//------------
typedef struct
{
   int           type;            // weir type code
   int           shape;           // weir shape code
   float         crestHt;         // crest height above node invert (ft)
   float         cDisch1;         // discharge coeff.
   float         cDisch2;         // discharge coeff. for ends
   float         endCon;          // end contractions

   float         cSurcharge;      // cDisch for equiv. orifice under surcharge
   float         length;          // equivalent length (ft)
   float         slope;           // slope for Vnotch & Trapezoidal weirs
   float         surfArea;        // equivalent surface area (ft2)
}  TWeir;


//---------------------
// OUTLET DEVICE OBJECT
//---------------------
typedef struct
{
    float        crestHt;         // crest ht. above node invert (ft)
    float        qCoeff;          // discharge coeff.
    float        qExpon;          // discharge exponent
    int          qCurve;          // index of discharge rating curve
}   TOutlet;


//-----------------
// POLLUTANT OBJECT
//-----------------
typedef struct
{
   char*         ID;              // Pollutant ID
   int           units;           // units
   float         mcf;             // mass conversion factor
   float         pptConcen;       // precip. concen.
   float         gwConcen;        // groundwater concen.
   float         rdiiConcen;      // RDII concen.
   float         kDecay;          // decay constant (1/sec)
   int           coPollut;        // co-pollutant index
   float         coFraction;      // co-pollutant fraction
   int           snowOnly;        // TRUE if buildup occurs only under snow
   int           sedflag;         // pollutant tye (sand, silt, clay, total)
}  TPollut;


//------------------------
// BUILDUP FUNCTION OBJECT
//------------------------
typedef struct
{
   int           normalizer;      // normalizer code (area or curb length)
   int           funcType;        // buildup function type code
   float         coeff[3];        // coeffs. of buildup function
   float         maxDays;         // time to reach max. buildup (days)
}  TBuildup;


//------------------------
// WASHOFF FUNCTION OBJECT
//------------------------
typedef struct
{
   int           funcType;        // washoff function type code
   float         coeff;           // function coeff.
   float         expon;           // function exponent
   float         sweepEffic;      // street sweeping fractional removal
   float         bmpEffic;        // best mgt. practice fractional removal
}  TWashoff;


//---------------
// LANDUSE OBJECT
//---------------
typedef struct
{
	char*         ID;				// landuse name
	float         sweepInterval;	// street sweeping interval (days)
	float         sweepRemoval;		// fraction of buildup available for sweeping
	float         sweepDays0;		// days since last sweeping at start
	
	// HSPF parameters
	float         pctimp;			// percent imperviousness of the landuse
	float         smpf;				// supporting management practice factor
	float         krer;				// coefficient in the soil detachment equation
	float         jrer;				// exponent in the soil detachment equation
	float         affix;			// fraction by which detached sediment storage 
									// decreases each day as a result of soil compaction
	float         cover;			// fraction of land surface which is shielded from 
									// rainfall erosion (not considering snow cover)
	float         kser;				// coefficient in the detached sediment washoff equation
	float         jser;				// exponent in the detached sediment washoff equation
	float         kger;				// coefficient in the matrix soil scour equation, 
									// which simulates gully erosion
	float         jger;				// exponent in the matrix soil scour equation, 
									// which simulates gully erosion
	float         frc_sand;			// total sediment fraction for sand
	float         frc_silt;			// total sediment fraction for silt
	float         frc_clay;			// total sediment fraction for clay

	TBuildup*     buildupFunc;		// array of buildup functions for pollutants
	TWashoff*     washoffFunc;		// array of washoff functions for pollutants
}  TLanduse;


//--------------------------
// REPORTING FLAGS STRUCTURE
//--------------------------
typedef struct
{
   char          report;          // TRUE if results report generated
   char          input;           // TRUE if input summary included
   char          subcatchments;   // TRUE if subcatchment results reported
   char          nodes;           // TRUE if node results reported
   char          links;           // TRUE if link results reported
   char          continuity;      // TRUE if continuity errors reported
   char          flowStats;       // TRUE if routing link flow stats. reported
   char          nodeStats;       // TRUE if routing node depth stats. reported
   char          controls;        // TRUE if control actions reported
   int           linesPerPage;    // number of lines printed per page
}  TRptFlags;


//-------------------------------
// CUMULATIVE RUNOFF TOTALS
//-------------------------------
typedef struct
{                                 // All volume totals are in ft.
   double        rainfall;        // rainfall volume 
   double        evap;            // evaporation loss
   double        infil;           // infiltration loss
   double        runoff;          // runoff volume
   double        initStorage;     // inital surface storage
   double        finalStorage;    // final surface storage
   double        initSnowCover;   // initial snow cover
   double        finalSnowCover;  // final snow cover
   double        snowRemoved;     // snow removal
   double        pctError;        // continuity error (%)
}  TRunoffTotals;


//--------------------------
// CUMULATIVE LOADING TOTALS
//--------------------------
typedef struct
{                                 // All loading totals are in lbs.
   double        initLoad;        // initial loading
   double        buildup;         // loading added from buildup
   double        deposition;      // loading added from wet deposition
   double        sweeping;        // loading removed by street sweeping
   double        bmpRemoval;      // loading removed by BMPs

///////////////////////////////////////////////
//  New element add to structure. (LR - 7/5/06 )
////////////////////////////////////////////////
   double        infil;           // loading removed by infiltration

   double        runoff;          // loading removed by runoff
   double        finalLoad;       // final loading
   double        pctError;        // continuity error (%)
}  TLoadingTotals;


//------------------------------
// CUMULATIVE GROUNDWATER TOTALS
//------------------------------
typedef struct
{                                 // All GW flux totals are in feet.
   double        infil;           // surface infiltration
   double        upperEvap;       // upper zone evaporation loss
   double        lowerEvap;       // lower zone evaporation loss
   double        lowerPerc;       // percolation out of lower zone
   double        gwater;          // groundwater flow
   double        initStorage;     // initial groundwater storage
   double        finalStorage;    // final groundwater storage
   double        pctError;        // continuity error (%)
}  TGwaterTotals;


//----------------------------
// CUMULATIVE ROUTING TOTALS
//----------------------------
typedef struct
{                                  // All routing totals are in ft3.
   double        dwInflow;         // dry weather inflow
   double        wwInflow;         // wet weather inflow
   double        gwInflow;         // groundwater inflow
   double        iiInflow;         // RDII inflow
   double        exInflow;         // direct inflow
   double        flooding;         // internal flooding
   double        outflow;          // external outflow
   double        reacted;          // reaction losses
   double        initStorage;      // initial storage volume
   double        finalStorage;     // final storage volume
   double        pctError;         // continuity error
}  TRoutingTotals;


//-----------------------
// SYSTEM-WIDE STATISTICS
//-----------------------
typedef struct
{
   float         minTimeStep;
   float         maxTimeStep;
   float         avgTimeStep;
   float         avgStepCount;
   float         steadyStateCount;
}  TSysStats;


//--------------------
// RAINFALL STATISTICS
//--------------------
typedef struct
{
   DateTime    startDate;
   DateTime    endDate;
   long        periodsRain;
   long        periodsMissing;
   long        periodsMalfunc;
}  TRainStats;


//------------------------
// SUBCATCHMENT STATISTICS
//------------------------
typedef struct
{
    float        precip;
    float        runon;
    float        evap;
    float        infil;
    float        runoff;
//////////////////////////////////////////////
////  New attribute added (LR - 3/10/06)  ////
//////////////////////////////////////////////
    float        maxFlow;         
}  TSubcatchStats;


//----------------
// NODE STATISTICS
//----------------
typedef struct
{
   float         avgDepth;
   float         maxDepth;
   DateTime      maxDepthDate;
   float         avgDepthChange;
   float         volFlooded;
   float         timeFlooded;
   float         timeCourantCritical;
//////////////////////////////////////////////
////  New attributes added (LR - 7/5/06)  ////
//////////////////////////////////////////////
   float         maxLatFlow;
   float         maxInflow;
   float         maxOverflow;
   DateTime      maxInflowDate;
   DateTime      maxOverflowDate;
}  TNodeStats;


/////////////////////////////////////////////////////////////
//  New structure added for storage statistics. (LR - 9/5/05)
/////////////////////////////////////////////////////////////
//-------------------
// STORAGE STATISTICS
//-------------------
typedef struct
{
   float         avgVol;
   float         maxVol;
   float         maxFlow;
   DateTime      maxVolDate;
}  TStorageStats;


/////////////////////////////////////////////////////////////
//  New structure added for outfall statistics. (LR - 7/5/06)
/////////////////////////////////////////////////////////////
//-------------------
// OUTFALL STATISTICS
//-------------------
typedef struct
{
   float        avgFlow;
   float        maxFlow;
   float*       totalLoad;
   int          totalPeriods;
}  TOutfallStats;


//----------------
// LINK STATISTICS
//----------------
typedef struct
{
   float         maxFlow;
   DateTime      maxFlowDate;
   float         maxVeloc;
   DateTime      maxVelocDate;

/////////////////////////////////////////////
////  New attribute added (LR - 7/5/06)  ////
/////////////////////////////////////////////
   float         maxDepth;

   float         avgFlowChange;
   float         avgFroude;
   float         timeSurcharged;
   float         timeInFlowClass[MAX_FLOW_CLASSES];
   float         timeCourantCritical;
}  TLinkStats;


//-------------------------
// MAXIMUM VALUE STATISTICS
//-------------------------
typedef struct
{
   int           objType;         // either NODE or LINK
   int           index;           // node or link index
   float         value;           // value of node or link statistic
}  TMaxStats; 


//------------------
// REPORT FIELD INFO
//------------------
typedef struct 
{
   char          Name[80];        // name of reported variable 
   char          Units[80];       // units of reported variable
   char          Enabled;         // TRUE if appears in report table
   int           Precision;       // number of decimal places when reported
}  TRptField;


//-----------------
// FILE INFORMATION
//-----------------
typedef struct
{
   char          name[MAXFNAME+1];     // file name
   char          mode;                 // NONE, SCRATCH, USE, or SAVE
   char          state;                // current state (OPENED, CLOSED)
   FILE*         file;                 // FILE structure pointer
}  TFile;
