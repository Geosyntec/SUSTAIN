//-----------------------------------------------------------------------------
//   climate.c
//
//   Project: EPA SWMM5
//   Version: 5.0
//   Date:    5/6/05   (Build 5.0.005)
//            9/5/05   (Build 5.0.006)
//            9/19/06  (Build 5.0.009)
//   Author:  L. Rossman
//
//   Climate related functions.
//-----------------------------------------------------------------------------

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <math.h>
#include "headers.h"

//-----------------------------------------------------------------------------
//  Constants
//-----------------------------------------------------------------------------
enum ClimateFileFormats {UNKNOWN_FORMAT,
                         USER_PREPARED,     // SWMM 5's own user format
                         TD3200,            // NCDC TD3200 format
                         DLY0204};          // Canadian DLY02 or DLY04 format
static const int   MAXCLIMATEVARS  = 4;
static const int   MAXDAYSPERMONTH = 32;

// These variables are used when processing climate files.
enum   ClimateVarType {TMIN, TMAX, EVAP, WIND};
static char* ClimateVarWords[] = {"TMIN", "TMAX", "EVAP", "WDMV", NULL};

//-----------------------------------------------------------------------------
//  Shared variables
//-----------------------------------------------------------------------------
// Temperature variables
static float    Tmin;                  // min. daily temperature (deg F)
static float    Tmax;                  // max. daily temperature (deg F)
static float    Trng;                  // 1/2 range of daily temperatures
static float    Trng1;                 // prev. max - current min. temp.
static float    Tave;                  // average daily temperature (deg F)
static float    Hrsr;                  // time of min. temp. (hrs)
static float    Hrss;                  // time of max. temp (hrs) 
static float    Hrday;                 // avg. of min/max temp times
static float    Dhrdy;                 // hrs. between min. & max. temp. times
static float    Dydif;                 // hrs. between max. & min. temp. times
static DateTime LastDay;               // date of last day with temp. data

// Climate file variables
static int      FileFormat;            // file format (see ClimateFileFormats)
static int      FileYear;              // current year of file data
static int      FileMonth;             // current month of year of file data
static int      FileDay;               // current day of month of file data
static int      FileLastDay;           // last day of current month of file data
static int      FileElapsedDays;       // number of days read from file
static float    FileValue[4];          // current day's values of climate data
static float    FileData[4][32];       // month's worth of daily climate data
static char     FileLine[MAXLINE+1];   // line from climate data file

//-----------------------------------------------------------------------------
//  External functions (defined in funcs.h)
//-----------------------------------------------------------------------------
//  climate_readParams                 // called by input_parseLine
//  climate_readEvapParams             // called by input_parseLine
//  climate_validate                   // called by project_validate
//  climate_openFile                   // called by runoff_open
//  climate_initState                  // called by project_init
//  climate_setState                   // called by runoff_execute
//  climate_getNextEvap                // called by runoff_getTimeStep

//-----------------------------------------------------------------------------
//  Local functions
//-----------------------------------------------------------------------------
static int  getFileFormat(void);
static void readFileLine(int *year, int *month);
static void readUserFileLine(int *year, int *month);
static void readTD3200FileLine(int *year, int *month);
static void readDLY0204FileLine(int *year, int *month);
static void readFileValues(void);
static void setFileValues(int param);
static void setEvap(DateTime theDate);
static void setTemp(DateTime theDate);
static void setWind(DateTime theDate);
static void updateTempTimes(int day);
static void updateFileValues(DateTime theDate);
static void parseUserFileLine(void);
static void parseTD3200FileLine(void);
static void parseDLY0204FileLine(void);
static void setTD3200FileValues(int param);



//=============================================================================

int  climate_readParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns error code
//  Purpose: reads climate/temperature parameters from input line of data
//
//  Format of data can be
//    TIMESERIES  name
//    FILE        name
//    WINDSPEED   MONTHLY  v1  v2  ...  v12
//    WINDSPEED   FILE 
//    SNOWMELT    v1  v2  ...  v6
//    ADC         IMPERV/PERV  v1  v2  ...  v10
//
{
    int      i, j, k;
    float    x[6], y;
    DateTime aDate;

    // --- identify keyword
    k = findmatch(tok[0], TempKeyWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[0]);
    switch (k)
    {
      case 0: // Time series name
        // --- check that time series name exists
        if ( ntoks < 2 ) return error_setInpError(ERR_ITEMS, "");
        i = project_findObject(TSERIES, tok[1]);
        if ( i < 0 ) return error_setInpError(ERR_NAME, tok[1]);

        // --- record the time series as being the data source for temperature
        Temp.dataSource = TSERIES_TEMP;
        Temp.tSeries = i;
        break;

      case 1: // Climate file
        // --- record file as being source of temperature data
        if ( ntoks < 2 ) return error_setInpError(ERR_ITEMS, "");
        Temp.dataSource = FILE_TEMP;

        // --- save name and usage mode of external climate file
        Fclimate.mode = USE_FILE;
        sstrncpy(Fclimate.name, tok[1], MAXFNAME);

        // --- save starting date to read from file if one is provided
        Temp.fileStartDate = NO_DATE;
        if ( ntoks > 2 )
        {
            if ( *tok[2] != '*')
            {
                if ( !datetime_strToDate(tok[2], &aDate) )
                    return error_setInpError(ERR_DATETIME, tok[2]);
                Temp.fileStartDate = aDate;
            }
        }
        break;

      case 2: // Wind speeds
        // --- check if wind speeds will be supplied from climate file
        if ( strcomp(tok[1], w_FILE) )
        {
            Wind.type = FILE_WIND;
        }

        // --- otherwise read 12 monthly avg. wind speed values
        else
        {
            if ( ntoks < 14 ) return error_setInpError(ERR_ITEMS, "");
            Wind.type = MONTHLY_WIND;
            for (i=0; i<12; i++)
            {
                if ( !getFloat(tok[i+2], &y) )
                    return error_setInpError(ERR_NUMBER, tok[i+2]);
                Wind.aws[i] = y;
            }
        }
        break;

      case 3: // Snowmelt params
        if ( ntoks < 7 ) return error_setInpError(ERR_ITEMS, "");
        for (i=1; i<7; i++)
        {
            if ( !getFloat(tok[i], &x[i-1]) )
                return error_setInpError(ERR_NUMBER, tok[i]);
        }
        // --- convert deg. C to deg. F for snowfall temperature
        if ( UnitSystem == SI ) x[0] = 9./5.*x[0] + 32.0;
        Snow.snotmp = x[0];
        Snow.tipm   = x[1];
        Snow.rnm    = x[2];
        Temp.elev   = x[3] / UCF(LENGTH);
        Temp.anglat = x[4];
        Temp.dtlong = x[5] / 60.0;
        break;

      case 4:  // Areal Depletion Curve data
        // --- check if data is for impervious or pervious areas
        if ( ntoks < 12 ) return error_setInpError(ERR_ITEMS, "");
        if      ( match(tok[1], w_IMPERV) ) i = 0;
        else if ( match(tok[1], w_PERV)   ) i = 1;
        else return error_setInpError(ERR_KEYWORD, tok[1]);

        // --- read 10 fractional values
        for (j=0; j<10; j++)
        {
            if ( !getFloat(tok[j+2], &y) || y < 0.0 || y > 1.0 ) 
                return error_setInpError(ERR_NUMBER, tok[j+2]);
            Snow.adc[i][j] = y;
        }
        break;
    }
    return 0;
}

//=============================================================================

int climate_readEvapParams(char* tok[], int ntoks)
//
//  Input:   tok[] = array of string tokens
//           ntoks = number of tokens
//  Output:  returns error code
//  Purpose: reads evaporation parameters from input line of data.
//
//  Data formats are:
//    CONSTANT  value
//    MONTHLY   v1 ... v12
//    TIMESERIES name
//    FILE      (v1 ... v12)
//
{
    int i, k;
    float  x;

    // --- find keyword indicating what form the evaporation data is in
    if ( ntoks < 2 ) return error_setInpError(ERR_ITEMS, "");
    k = findmatch(tok[0], EvapTypeWords);
    if ( k < 0 ) return error_setInpError(ERR_KEYWORD, tok[0]);

    // --- process data depending on its form
    Evap.type = k;
    switch ( k )
    {
      case CONSTANT_EVAP:
        // --- for constant evap., fill monthly avg. values with same number
        if ( !getFloat(tok[1], &x) )
            return error_setInpError(ERR_NUMBER, tok[1]);
        for (i=0; i<12; i++) Evap.monthlyEvap[i] = x;
        break;

      case MONTHLY_EVAP:
        // --- for monthly evap., read a value for each month of year
        if ( ntoks < 13 ) return error_setInpError(ERR_ITEMS, "");
        for ( i=0; i<12; i++)
            if ( !getFloat(tok[i+1], &Evap.monthlyEvap[i]) )
                return error_setInpError(ERR_NUMBER, tok[i+1]);
        break;           

      case TIMESERIES_EVAP:
        // --- for time series evap., read name of time series
        i = project_findObject(TSERIES, tok[1]);
        if ( i < 0 ) return error_setInpError(ERR_NAME, tok[1]);
        Evap.tSeries = i;
        break;

      case FILE_EVAP:
        // --- for evap. from climate file, read monthly pan coeffs.
        //     if they are provided (default values are 1.0)
        if ( ntoks > 1 )
        {
            if ( ntoks < 13 ) return error_setInpError(ERR_ITEMS, "");
            for (i=0; i<12; i++)
            {
                if ( !getFloat(tok[i+1], &Evap.panCoeff[i]) )
                    return error_setInpError(ERR_NUMBER, tok[i+1]);
            }
        }
        break;
    }
    return 0;
}

//=============================================================================

void climate_validate()
//
//  Input:   none
//  Output:  none
//  Purpose: validates climatological variables
//
{
    float    a, z, pa;

    // --- check if climate file was specified when used 
    //     to supply wind speed or evap rates
    if ( Wind.type == FILE_WIND || Evap.type == FILE_EVAP )
    {
        if ( Fclimate.mode == NO_FILE )
        {
            report_writeErrorMsg(ERR_NO_CLIMATE_FILE, "");
        }
    }

    // --- snow melt parameters tipm & rnm must be fractions
    if ( Snow.tipm < 0.0 ) Snow.tipm = 0.0;
    if ( Snow.tipm > 1.0 ) Snow.tipm = 1.0;
    if ( Snow.rnm < 0.0 ) Snow.rnm = 0.0;
    if ( Snow.rnm > 1.0 ) Snow.rnm = 1.0;

    // --- latitude should be between -60 & 60 degrees
    a = Temp.anglat;
    if ( a < -60.0 ) a = -60.0;
    if ( a > 60.0  ) a = 60.0;
    Temp.tanAnglat = tan(a * PI / 180.0);

    // --- compute psychrometric constant
    z = Temp.elev / 1000.0;
    if ( z <= 0.0 ) pa = 29.9;
    else  pa = 29.9 - 1.02*z + 0.0032*pow(z, 2.4); // atmos. pressure
    Temp.gamma = 0.000359 * pa;
}

//=============================================================================

void climate_openFile()
//
//  Input:   none
//  Output:  none
//  Purpose: opens a climate file and reads in first set of values.
//
{
    int i, m, y;

    // --- open the file
    if ( (Fclimate.file = fopen(Fclimate.name, "rt")) == NULL )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_OPEN, Fclimate.name);
        return;
    }

    // --- initialize values of file's climate variables
    //     (Temp.ta was previously initialized in project.c)
    FileValue[TMIN] = Temp.ta;
    FileValue[TMAX] = Temp.ta;
    FileValue[EVAP] = 0.0;
    FileValue[WIND] = 0.0;

    // --- find climate file's format
    FileFormat = getFileFormat();
    if ( FileFormat == UNKNOWN_FORMAT )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_READ, Fclimate.name);
        return;
    }


fprintf(Frpt.file, "\n  File Format = %d", FileFormat);

    // --- position file at start of range of dates to be used
    rewind(Fclimate.file);
    strcpy(FileLine, "");

////////////////////////////////////////////////////////////////////////
//  Code modified to begin reading climate file at either user-specified
//  month/year or at start of simulation period. (LR - 9/5/05)
////////////////////////////////////////////////////////////////////////
    if ( Temp.fileStartDate == NO_DATE )
        datetime_decodeDate(StartDate, &FileYear, &FileMonth, &FileDay); 
    else
        datetime_decodeDate(Temp.fileStartDate, &FileYear, &FileMonth, &FileDay);

//////////////////////////////////////////////////////////////////
//  Modified to keep reading until EOF encountered. (LR - 9/19/06)
//////////////////////////////////////////////////////////////////
    while ( !feof(Fclimate.file) )
    {
        strcpy(FileLine, "");
        readFileLine(&y, &m);
        if ( y == FileYear && m == FileMonth ) break;
    }
    if ( feof(Fclimate.file) )
    {
        report_writeErrorMsg(ERR_CLIMATE_END_OF_FILE, Fclimate.name);
        return;
    }
    

    // --- initialize file dates and current climate variable values 
    if ( !ErrorCode )
    {
        FileElapsedDays = 0;
        FileLastDay = datetime_daysPerMonth(FileYear, FileMonth);
        readFileValues();
        for (i=TMIN; i<=WIND; i++)
        {
            if ( FileData[i][FileDay] == MISSING ) continue;
            FileValue[i] = FileData[i][FileDay];
        }
    }
}

//=============================================================================

void climate_initState()
//
//  Input:   none
//  Output:  none
//  Purpose: initializes climate state variables.
//
{
    LastDay = NO_DATE;
    Temp.tmax = MISSING;
    Snow.removed = 0.0;
}

//=============================================================================

void climate_setState(DateTime theDate)
//
//  Input:   theDate = simulation date
//  Output:  none
//  Purpose: sets climate variables for current date.
//
{
    if ( Fclimate.mode == USE_FILE ) updateFileValues(theDate);
    if ( Temp.dataSource != NO_TEMP ) setTemp(theDate);
    setEvap(theDate);
    setWind(theDate);
}

//=============================================================================

DateTime climate_getNextEvap(DateTime days)
//
//  Input:   days = current simulation date (in whole days)
//  Output:  returns date (in whole days) when evaporation rate next changes
//  Purpose: finds date for next change in evaporation.
//
{
    int yr, mon, day, k;
    double d, e;

    switch ( Evap.type )
    {
      case CONSTANT_EVAP:
        return days + 365.;

      case MONTHLY_EVAP:
        datetime_decodeDate(days, &yr, &mon, &day);
        if ( mon == 12 )
        {
            mon = 1;
            yr++;
        }
        else mon++;
        return datetime_encodeDate(yr, mon, 1);

      case TIMESERIES_EVAP:
        k = Evap.tSeries;
        while ( table_getNextEntry(&Tseries[k], &d, &e) )
        {
            if ( d > days ) return d;
        }
        return days + 365.;

      case FILE_EVAP:
        return days + 1.0;

      default: return days + 365.;
    }
}

//=============================================================================

void updateFileValues(DateTime theDate)
//
//  Input:   theDate = current simulation date
//  Output:  none
//  Purpose: updates daily climate variables for new day or reads in
//           another month worth of values if a new month begins.
//
//  NOTE:    counters FileElapsedDays, FileDay, FileMonth, FileYear and
//           FileLastDay were initialized in climate_openFile().
//
{
    int i;
    int deltaDays;

    // --- see if a new day has begun
    deltaDays = (int)(floor(theDate) - floor(StartDateTime));
    if ( deltaDays > FileElapsedDays )
    {
        // --- advance day counters
        FileElapsedDays++;
        FileDay++;

        // --- see if new month of data needs to be read from file
        if ( FileDay > FileLastDay )
        {
            FileMonth++;
            if ( FileMonth > 12 )
            {
                FileMonth = 1;
                FileYear++;
            }
            readFileValues();
            FileDay = 1;
            FileLastDay = datetime_daysPerMonth(FileYear, FileMonth);
        }

        // --- set climate variables for new day
        for (i=TMIN; i<=WIND; i++)
        {
            // --- no change in current value if its missing
            if ( FileData[i][FileDay] == MISSING ) continue;
            FileValue[i] = FileData[i][FileDay];
        }
    }
}

//=============================================================================

void setTemp(DateTime theDate)
//
//  Input:   theDate = simulation date
//  Output:  none
//  Purpose: updates temperatures for new simulation date.
//
{
    int      j;                        // snow data object index
    int      k;                        // time series index
    int      day;                      // day of year
    DateTime theDay;                   // calendar day
    float    hour;                     // hour of day

    // --- see if a new day has started
    theDay = floor(theDate);
    if ( theDay > LastDay )
    {
        // --- update min. & max. temps & their time of day
        day = datetime_dayOfYear(theDate);
        if ( Temp.dataSource == FILE_TEMP )
        {
            Tmin = FileValue[TMIN];
            Tmax = FileValue[TMAX];
            updateTempTimes(day);
        }

        // --- compute snow melt coefficients based on day of year
        Snow.season = sin(0.0172615*(day-81.0));
        for (j=0; j<Nobjects[SNOWMELT]; j++)
        {
            snow_setMeltCoeffs(j, Snow.season);
        }

        // --- update date of last day analyzed
        LastDay = theDate;
    }

    // --- for min/max daily temps. from climate file,
    //     compute hourly temp. by sinusoidal interp.
    if ( Temp.dataSource == FILE_TEMP )
    {
        hour = (theDate - theDay) * 24.0;
        if ( hour < Hrsr )
            Temp.ta = Tmin + Trng1/2.0 * sin(PI/Dydif * (Hrsr - hour));
        else if ( hour >= Hrsr && hour <= Hrss )
            Temp.ta = Tave + Trng * sin(PI/Dhrdy * (Hrday - hour));
        else
            Temp.ta = Tmax - Trng * sin(PI/Dydif * (hour - Hrss));
    }

    // --- for user-supplied temperature time series,
    //     get temperature value from time series
    if ( Temp.dataSource == TSERIES_TEMP )
    {
        k = Temp.tSeries;
        if ( k >= 0)
        {
            Temp.ta = table_tseriesLookup(&Tseries[k], theDate, TRUE);

            // --- convert from deg. C to deg. F if need be
            if ( UnitSystem == SI )
            {
                Temp.ta = (9./5.) * Temp.ta + 32.0;
            }
        }
    }

    // --- compute saturation vapor pressure
    Temp.ea = 8.1175e6 * exp(-7701.544 / (Temp.ta + 405.0265) );
}

//=============================================================================

void setEvap(DateTime theDate)
//
//  Input:   theDate = simulation date
//  Output:  none
//  Purpose: sets evaporation rate (ft/sec) for a specified date.
//
{
    int yr, mon, day, k;

    switch ( Evap.type )
    {
      case CONSTANT_EVAP:
        Evap.rate = Evap.monthlyEvap[0];
        break;

      case MONTHLY_EVAP:
        datetime_decodeDate(theDate, &yr, &mon, &day);
        Evap.rate = Evap.monthlyEvap[mon-1];
        break;

      case TIMESERIES_EVAP:
        k = Evap.tSeries;
        Evap.rate = table_tseriesLookup(&Tseries[k], theDate, TRUE);
        break;

      case FILE_EVAP:
        Evap.rate = FileValue[EVAP];
        datetime_decodeDate(theDate, &yr, &mon, &day);
        Evap.rate *= Evap.panCoeff[mon-1];
        break;

      default: Evap.rate = 0.0;
    }

    // --- convert rate from in/day (mm/day) to ft/sec
    Evap.rate /= UCF(EVAPRATE);
}

//=============================================================================

void setWind(DateTime theDate)
//
//  Input:   theDate = simulation date
//  Output:  none
//  Purpose: sets wind speed (mph) for a specified date.
//
{
    int yr, mon, day;

    switch ( Wind.type )
    {
      case MONTHLY_WIND:
        datetime_decodeDate(theDate, &yr, &mon, &day);
        Wind.ws = Wind.aws[mon-1] / UCF(WINDSPEED);
        break;

      case FILE_WIND:
        Wind.ws = FileValue[WIND];
        break;

      default: Wind.ws = 0.0;
    }
}

//=============================================================================

void updateTempTimes(int day)
//
//  Input:   day = day of year
//  Output:  none
//  Purpose: computes time of day when min/max temperatures occur.
//           (min. temp occurs at sunrise, max. temp. at 3 hrs. < sunset)
//
{
    float decl;                        // earth's declination
    float hrang;                       // hour angle of sunrise/sunset
    float arg;

    decl  = 0.40928*cos(0.017202*(172.0-day));
    arg = -tan(decl)*Temp.tanAnglat;
    if      ( arg <= -1.0 ) arg = PI;
    else if ( arg >= 1.0 )  arg = 0.0;
    else                    arg = acos(arg);
    hrang = 3.8197 * arg;
    Hrsr  = 12.0 - hrang + Temp.dtlong;
    Hrss  = 12.0 + hrang + Temp.dtlong - 3.0;
    Dhrdy = Hrsr - Hrss;
    Dydif = 24.0 + Hrsr - Hrss;
    Hrday = (Hrsr + Hrss) / 2.0;
    Tave  = (Tmin + Tmax) / 2.0;
    Trng  = (Tmax - Tmin) / 2.0;
    if ( Temp.tmax == MISSING ) Trng1 = Tmax - Tmin;
    else                        Trng1 = Temp.tmax - Tmin;
    Temp.tmax = Tmax;
}

//=============================================================================

int  getFileFormat()
//
//  Input:   none
//  Output:  returns code number of climate file's format
//  Purpose: determines what format the climate file is in.
//
///////////////////////////////////////////////////////////////////
//  Modified the tests for TD3200 & Canadian formats (LR - 9/19/06)
///////////////////////////////////////////////////////////////////
{
    char recdType[4] = "";
    char elemType[4] = "";
    char filler[5] = "";                                   // New line (LR - 9/19/06)
    char staID[80];
    char s[80];
    char line[MAXLINE];
    int  y, m, d, n;

    // --- read first line of file
    if ( fgets(line, MAXLINE, Fclimate.file) == NULL ) return UNKNOWN_FORMAT;

    // --- check for TD3200 format
    sstrncpy(recdType, line, 3);
    sstrncpy(filler, &line[23], 4);
    if ( strcmp(recdType, "DLY") == 0 &&
         strcmp(filler, "9999")  == 0 ) return TD3200;     // New line (LR - 9/19/06)

    // --- check for DLY0204 format
    if ( strlen(FileLine) >= 233 )                         // New test (LR - 9/19/06)
    { 
        sstrncpy(elemType, &line[13], 3);
        n = atoi(elemType);
        if ( n == 1 || n == 2 ) return DLY0204;
    }
    
    // --- check for USER_PREPARED format
    n = sscanf(line, "%s %d %d %d %s", staID, &y, &m, &d, s);
    if ( n == 5 ) return USER_PREPARED;
    return UNKNOWN_FORMAT;
}

//=============================================================================

void readFileLine(int *y, int *m)
//
//  Input:   none
//  Output:  y = year
//           m = month
//  Purpose: reads year & month from next line of climate file.
//
{
    // --- read next line from climate data file
    if ( strlen(FileLine) == 0 )
    {
///////////////////////////////////////////////////////////////
//  Modified to report end of file error message. (LR - 9/5/05)
///////////////////////////////////////////////////////////////
//  Previous modification removed. (LR - 9/19/06)
/////////////////////////////////////////////////
        if ( fgets(FileLine, MAXLINE, Fclimate.file) == NULL ) return;
    }

    // --- parse year & month from line
    switch (FileFormat)
    {
    case  USER_PREPARED: readUserFileLine(y, m);   break;
    case  TD3200:        readTD3200FileLine(y,m);  break;
    case  DLY0204:       readDLY0204FileLine(y,m); break;
    }
}

//=============================================================================

void readUserFileLine(int* y, int* m)
//
//  Input:   none
//  Output:  y = year
//           m = month
//  Purpose: reads year & month from line of User-Prepared climate file.
//
{
    int n;
    char staID[80];
    n = sscanf(FileLine, "%s %d %d", staID, y, m);
    if ( n < 3 )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_READ, Fclimate.name);
    }
}

//=============================================================================

void readTD3200FileLine(int* y, int* m)
//
//  Input:   none
//  Output:  y = year
//           m = month
//  Purpose: reads year & month from line of TD-3200 climate file.
//
{
    char recdType[4] = "";
    char year[5] = "";
    char month[3] = "";
    int  len;

    // --- check for minimum number of characters
    len = strlen(FileLine);
    if ( len < 30 )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_READ, Fclimate.name);
        return;
    }

    // --- check for proper type of record
    sstrncpy(recdType, FileLine, 3);
    if ( strcmp(recdType, "DLY") != 0 )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_READ, Fclimate.name);
        return;
    }

    // --- get record's date 
    sstrncpy(year,  &FileLine[17], 4);
    sstrncpy(month, &FileLine[21], 2);
    *y = atoi(year);
    *m = atoi(month);
}

//=============================================================================

void readDLY0204FileLine(int* y, int* m)
//
//  Input:   none
//  Output:  y = year
//           m = month
//  Purpose: reads year & month from line of DLY02 or DLY04 climate file.
//
{
    char year[5] = "";
    char month[3] = "";
    int  len;

    // --- check for minimum number of characters
    len = strlen(FileLine);
    if ( len < 16 )
    {
        report_writeErrorMsg(ERR_CLIMATE_FILE_READ, Fclimate.name);
        return;
    }

    // --- get record's date 
    sstrncpy(year,  &FileLine[7], 4);
    sstrncpy(month, &FileLine[11], 2);
    *y = atoi(year);
    *m = atoi(month);
}

//=============================================================================

void readFileValues()
//
//  Input:   none
//  Output:  none
//  Purpose: reads next month's worth of data from climate file.
//
{
    int  i, j;
    int  y, m;

    // --- initialize FileData array to missing values
    for ( i=0; i<MAXCLIMATEVARS; i++)
    {
        for (j=0; j<MAXDAYSPERMONTH; j++) FileData[i][j] = MISSING;
    }

    while ( !ErrorCode )
    {
        // --- return when date on line is after current file date
        if ( feof(Fclimate.file) ) return;
        readFileLine(&y, &m);
        if ( y > FileYear || m > FileMonth ) return;

        // --- parse climate values from file line
        switch (FileFormat)
        {
        case  USER_PREPARED: parseUserFileLine();   break;
        case  TD3200:        parseTD3200FileLine();  break;
        case  DLY0204:       parseDLY0204FileLine(); break;
        }
        strcpy(FileLine, "");
    }
}

//=============================================================================

void parseUserFileLine()
//
//  Input:   none
//  Output:  none
//  Purpose: parses climate variable values from a line of a user-prepared
//           climate file.
//
{
    int   n;
    int   y, m, d;
    char  staID[80];
    char  s0[80];
    char  s1[80];
    char  s2[80];
    char  s3[80];
    float x;

    // --- read day, Tmax, Tmin, Evap, & Wind from file line
    n = sscanf(FileLine, "%s %d %d %d %s %s %s %s",
        staID, &y, &m, &d, s0, s1, s2, s3);
    if ( n < 4 ) return;
    if ( d < 1 || d > 31 ) return;

    // --- process TMAX
    if ( strlen(s0) > 0 && *s0 != '*' )
    {
        x = atof(s0);
        if ( UnitSystem == SI ) x = 9./5.*x + 32.0;
        FileData[TMAX][d] =  x;
    }
 
    // --- process TMIN
    if ( strlen(s1) > 0 && *s1 != '*' )
    {
        x = atof(s1);
        if ( UnitSystem == SI ) x = 9./5.*x + 32.0;
        FileData[TMIN][d] =  x;
    }

    // --- process EVAP
    if ( strlen(s2) > 0 && *s2 != '*' ) FileData[EVAP][d] = atof(s2);

    // --- process WIND
    if ( strlen(s3) > 0 && *s3 != '*' ) FileData[WIND][d] = atof(s3);
}

//=============================================================================

void parseTD3200FileLine()
//
//  Input:   none
//  Output:  none
//  Purpose: parses climate variable values from a line of a TD3200 file.
//
{
    int  i;
    char param[5] = "";

    // --- parse parameter name
    sstrncpy(param, &FileLine[11], 4);

    // --- see if parameter is temperature, evaporation or wind speed
    for (i=0; i<MAXCLIMATEVARS; i++)
    {
        if (strcmp(param, ClimateVarWords[i]) == 0 ) setTD3200FileValues(i);
    }
}

//=============================================================================

void setTD3200FileValues(int i)
//
//  Input:   i = climate variable code
//  Output:  none
//  Purpose: reads month worth of values for climate variable from TD-3200 file.
//
{
    char valCount[4] = "";
    char day[3] = "";
    char sign[2] = "";
    char value[6] = "";
    char flag2[2] = "";
    float x;
    int  nValues;
    int  j, k, d;
    int  lineLength;

    // --- parse number of days with data from cols. 27-29 of file line
    sstrncpy(valCount, &FileLine[27], 3);
    nValues = atoi(valCount);
    lineLength = strlen(FileLine);

    // --- check for enough characters on line
    if ( lineLength >= 12*nValues + 30 )
    {
        // --- for each day's value
        for (j=0; j<nValues; j++)
        {
            // --- parse day, value & flag from file line
            k = 30 + j*12;
            sstrncpy(day,   &FileLine[k], 2);
            sstrncpy(sign,  &FileLine[k+4], 1);
            sstrncpy(value, &FileLine[k+5], 5);
            sstrncpy(flag2, &FileLine[k+11], 1);

            // --- if value is valid then store it in FileData array
            d = atoi(day);
            if ( strcmp(value, "99999") != 0 
                 && ( flag2[0] == '0' || flag2[0] == '1')
                 &&   d > 0
                 &&   d <= 31 )
            {
                // --- convert from string value to numerical value
                x = atof(value);
                if ( sign[0] == '-' ) x = -x;

                // --- convert evaporation from hundreths of inches
                if ( i == EVAP ) x /= 100.0;
                
                // --- convert wind speed from miles/day to miles/hour
                if ( i == WIND ) x /= 24.0;
 
                // --- store value
                FileData[i][d] = x;
            }
        }
    }
}

//=============================================================================

void parseDLY0204FileLine()
//
//  Input:   none
//  Output:  none
//  Purpose: parses a month's worth of climate variable values from a line of
//           a DLY02 or DLY04 climate file.
//
{
    int  j, k, p;
    char param[4] = "";
    char sign[2]  = "";
    char value[6] = "";
    char code[2]  = "";
    float x;

    // --- parse parameter name
    sstrncpy(param, &FileLine[14], 3);

    // --- see if parameter is min or max temperature
    p = atoi(param);
    if ( p == 1 ) p = TMAX;
    else if ( p == 2 ) p = TMIN;
    else return;

    // --- check for 233 characters on line
    if ( strlen(FileLine) < 233 ) return;

    // --- for each of 31 days
    k = 17;
    for (j=1; j<=31; j++)
    {
        // --- parse value & flag from file line
        sstrncpy(sign,  &FileLine[k], 1);
        sstrncpy(value, &FileLine[k+1], 5);
        sstrncpy(code,  &FileLine[k+6], 1);
        k += 7;

        // --- if value is valid then store it in FileData array
        if ( strcmp(value, "99999") != 0 && strcmp(value, "     ") != 0 )
        {
            // --- convert from integer tenths of a degree C to degrees F
            x = atof(value) / 10.0;
            if ( sign[0] == '-' ) x = -x;
            x = 9./5.*x + 32.0;
            FileData[p][j] = x;
        }
    }
}

//=============================================================================
