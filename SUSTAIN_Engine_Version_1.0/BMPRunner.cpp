// BMPRunner.cpp: implementation of the CBMPRunner class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "LandUse.h"
#include "BMPSite.h"
#include "SiteLandUse.h"
#include "SitePointSource.h"
#include "Sediment.h"
#include "BMPData.h"
#include "BMPRunner.h"
#include <math.h>
#include <afxtempl.h>
#include "ProgressWnd.h"	 
#include "BMPOptimizer.h"	 

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#define SMALLNUM	1.0E-15

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBMPRunner::CBMPRunner()
{
	optcounter = 0;
	outcounter = 0;
	lInitRunTime = 0.0;
	nMaxRun = 0;	
	pBMPData = NULL;
	pWndProgress = NULL;	
	time_i = COleDateTime(1890,1,1,0,0,0);
	fp = NULL;
}

CBMPRunner::CBMPRunner(CBMPData* bmpData)
{
	optcounter = 0;
	outcounter = 0;
	lInitRunTime = 0.0;
	nMaxRun = 0;	
	pBMPData = bmpData;
	pWndProgress = NULL;	
	time_i = COleDateTime(1890,1,1,0,0,0);
	fp = NULL;
}

CBMPRunner::~CBMPRunner()
{

}

void CBMPRunner::bmp_a(int nInfiltMethod,int nGAindex, bool underdrain_on,int timestep,
	 int npeople,int ddays,int releasetype,int weirtype,int& counter,double oinflow,
	 double BMParea,double orifice_area,double orificeheight,double orificecoef,
	 double weirwidth,double weirheight,double weirangle,double cisternoutflow,
	 double soildepth,double soilporosity,double finalf,double vegparma,double holtpar,
	 double udfinalf,double udsoildepth,double udsoilporosity,double FC,double WP,
	 double ETrate,double& AET,double& perc,double& ovolume,double& ostage,
	 double& infilt,double& orifice,double& weir,double& osa,double& ostorage,
	 double& udout,double& seepage)
{
	if (BMParea <= 0)
		return;

	int nivl = 60 / timestep;	// number of intervals per hour 
	double g = 32.2;			// ft/sec^2

	double r__1 = 0., r__2 = 0., r__3 = 0.,h = 0., f = 0.;
	double nweir = 0.,norifice = 0.,nudout = 0.,ninfilt = 0.,nperc = 0.0;
	double nseepage = 0.,nAET = 0., nAETsurf = 0.;
	double tempweir = 0.,temporifice = 0.,tempudout = 0.,tempinfilt = 0.,tempperc = 0.;
	double tempseepage = 0.,tempAET = 0.;

	// at the begining of the timestep
	double AWC      = soildepth*12.0*(FC - WP);	//available water capacity for plants only (in)
	double nvolume  = ovolume;				//available water in the water column (ft3)
	double nsa      = osa;					//available pores of the soil column (in)
	double nstorage = ostorage;				//available pores of the underdrain column (in)
	double inflow   = oinflow / nivl;		//ft3/hr to ft3/ivl
	double nstage   = ostage;				//ft

	if (inflow > SMALLNUM)
		counter = 0;
	else
		counter++;

	//calculate the max available space in soil column (in)
	double nsamax = soildepth*12.0*(soilporosity - WP); 

	if (nsamax < SMALLNUM)
		nsamax = 0.0;

	//calculate the max available space in underdrain column (in)
	double nstoragemax = udsoildepth*12.0*udsoilporosity;

	if (nstoragemax < SMALLNUM)
		nstoragemax = 0.0;

	for (int i=0; i<nivl; i++) 
	{
		ninfilt  = 0.0;		// surface infiltration (in/ivl)
		norifice = 0.0;		// orifice outflow (ft3/ivl)
		nweir    = 0.0;		// weir outflow (ft3/ivl)
		nudout   = 0.0;		// underdrain outflow (ft3/ivl)
		nperc    = 0.0;		// percolation to underdrain storage (in/ivl)
		nAETsurf = 0.0;		// actual evapotranspiration from the surface (in/ivl) 
		nAET     = 0.0;		// actual evapotranspiration from the subsurface (in/ivl)
		nseepage = 0.0;		// deep percolation to groundwater storage (in/ivl)

		float tStep = timestep * 60.0;				//sec/ivl
		float depth = nstage;						//ft
		float finflow = inflow / (BMParea * tStep);	//ft/sec

		//update the volume
		nvolume += inflow;	// ft3

		//update the stage
		nstage  = nvolume / BMParea;  //ft

		//WATER COLUMN
		if (nstage > SMALLNUM)			
		{
			// calculate the evaporation from the surface (in/ivl)
			nAETsurf = min(nstage*12.0,ETrate / nivl);
				
			//update the stage (ft)
			nstage -= nAETsurf/12.0;

			if (nstage > SMALLNUM)	
			{
				if (nInfiltMethod == 0)
				{
					// calculate infiltration potential per hour using Holtan Method (in/hr)
					f = holtpar * vegparma * pow(nsa, 1.4) + finalf;

					// f is converted from in/hr to in/ivl
					f /= nivl;
				}
				else
				{
					float surfEvap = nAETsurf / (12.0 * tStep);	//ft/sec
					// calculate infiltration rate per hour using Green-Ampt Method (ft/sec)
					f = infil_getInfil(nGAindex,GREEN_AMPT,tStep,(finflow - surfEvap),depth);

					//convert infiltration rate from ft/sec to in/ivl
					f = f / 12.0 * tStep;
				}
				
				//infiltration is the min of potential infiltration or available water in basin
				ninfilt = min(f, nstage*12.0);	// in/ivl

				//update the stage (ft)
				nstage -= ninfilt/12.0;

				if (nstage > SMALLNUM)	
				{
					//update the surface storage
					nvolume = nstage * BMParea; //ft3

					//calculate the orifice and weir outflow
				}
				else
				{
					ninfilt += (nstage * 12);	//add to infiltrated water(in/ivl)
					nstage = 0.0;	//dry surface
					nvolume = 0.0;	//no surface water
					norifice = 0.0;	//no orifice outflow
					nweir    = 0.0;	//no weir outflow
				}
			}
			else
			{
				nAETsurf += nstage * 12;//add small water to evaporated water (in/ivl)
				nstage = 0.0;	//dry surface
				nvolume = 0.0;	//no surface water storage
				ninfilt = 0.0;	//no water available for infiltration
				norifice = 0.0;	//no orifice outflow
				nweir    = 0.0;	//no weir outflow
			}
		} 
		else  // if there is no water currently in the basin... 
		{
			nAETsurf = nstage * 12;	//all water evaporated (in/ivl)
			nstage = 0.0;	//dry surface
			nvolume = 0.0;	//no surface storage
			ninfilt = 0.0;	//no infiltration
			norifice = 0.0;	//no orifice outflow
			nweir    = 0.0;	//no weir outflow
		}

		// SOIL LAYER
		if (nsamax > 0)
		{
			// extra water after saturation (in)
			double excesswater = ninfilt - nsa;
			
			//sub-surface evapotranspiration (in/ivl)
			nAET = ETrate / nivl - nAETsurf;

			if (nAET > (nsamax + excesswater))
				nAET = nsamax + excesswater;
				
			if (excesswater >= 0) // saturation 
			{
				nperc = min(finalf/nivl, excesswater + nsamax - nAET - AWC);
				nperc = max(0.0, nperc);

				double excessinfilt = max(0.0, excesswater - nAET - nperc);
				ninfilt -= excessinfilt;//reduce the amount of infiltartion	
				
				// send back to surface storage
				nstage += excessinfilt/12; //ft
				nvolume = nstage * BMParea; //ft3

				//update the soil available storage
				nsa = nsa - ninfilt + nperc + nAET;
			}
			else   // no saturation
			{
				if (-excesswater < nsamax)	// there is water to percolate
				{
					r__1 = (nsamax + excesswater) / nsamax;
					r__1 = (0.5 + 0.5 * r__1);
					r__1 = max(0.0, r__1);

					nperc = min(r__1 * finalf / nivl, excesswater + nsamax - nAET - AWC);	
					nperc = max(0.0, nperc);
					nsa   = nsa - ninfilt + nperc + nAET;
				}
				else
				{
					// no percolation to underdrain storage
					nperc = 0.0; 
					nsa   = nsamax;
				}
			}
		}
		else
		{
			// there is no storage capacity
			nAET  = 0.0;
			nsa   = 0.0;
			nperc = ninfilt; 
		}

		nsa = max(nsa, 0.0);	
		nsa = min(nsa, nsamax);	
		
		// UNDERDRAIN SOIL LAYER
		if(nstoragemax > 0)
		{
			// extra water after saturation (in)
			double excesswater = nperc - nstorage;

			if (excesswater >= 0) // saturation 
			{
				nseepage = min(udfinalf/nivl, excesswater + nstoragemax);
				nudout   = 0.0;

				if(underdrain_on)
				{
					nudout = max(0.0, excesswater - nseepage);
				}
				else
				{
					double excessperc = max(0.0, excesswater - nseepage);
					nperc   -= excessperc;

					//check if soil column has space to hold this excess water
					if(nsa > 0)
					{
						//check the available space
						if(nsa < excessperc)
						{
							ninfilt -= (excessperc - nsa);
							nsa = 0.0;	// saturated (no more space)

							// send back to the surface storage
							nstage += (excessperc - nsa)/12; //ft
							nvolume = nstage * BMParea; //ft3
						}
						else
						{
							nsa -= excessperc;
						}
					}
					else
					{
						//return this excess water to water column
						ninfilt -= excessperc;

						// send back to the surface storage
						nstage += excessperc / 12; //ft
						nvolume = nstage * BMParea; //ft3
					}
				}

				nstorage = nstorage - nperc + nseepage + nudout;
			}
			else	// no saturation (no underdrain outflow)
			{
				nudout = 0.0;

				if (-excesswater < nstoragemax)	// there is water to percolate
				{
					nseepage = min(udfinalf/nivl, excesswater + nstoragemax);
					nstorage = nstorage - nperc + nseepage;
				}
				else
				{
					nseepage = 0.0; 
					nstorage = nstoragemax;
				}
			}
		}
		else
		{
			// no underdrain outflow
			nudout   = 0.0;
			nstorage = 0.0;
			nseepage = nperc;	// whatever percolate will enter to groundwater storage
		}

		nstorage = max(nstorage, 0.0);	
		nstorage = min(nstorage, nstoragemax);
		
		//WATER COLUMN
		if (nstage > SMALLNUM)			
		{
			// CALCULATE ORIFICE FLOW
			if (nstage > orificeheight)
			{
				h = nstage - orificeheight;			// height of water above the orifice
				double norifice_max = h * BMParea;	// ft3

				if (releasetype == 1) 
				{
					// cistern outflow
					norifice = cisternoutflow / nivl * npeople;	//ft3/ivl
					if (norifice > norifice_max)
						norifice = norifice_max;
				}
				else if (releasetype == 2)
				{
					if (counter > ddays * 24)
					{
						norifice = sqrt(2 * g * h) * orifice_area  * orificecoef;//cfs
						norifice = norifice * 3600 / nivl;	//ft3/ivl
						if (norifice > norifice_max)
							norifice = norifice_max;
					}
					else
					{
						norifice = 0.0;
					}
				}
				else
				{
					norifice = sqrt(2 * g * h) * orifice_area  * orificecoef;  //cfs 	
					norifice = norifice * 3600 / nivl;	//ft3/ivl
					if (norifice > norifice_max)
						norifice = norifice_max;
				}
			}
			else
			{
				// no orifice outflow
				norifice = 0.0;
			}

			if (norifice > 0)
			{
				// RECALCULATE VOLUME AND STAGE
				r__1    = nvolume - norifice;	//ft3
				nvolume = max(r__1, 0.0);
				nstage  = nvolume / BMParea;  //ft
			}

			if (nstage > SMALLNUM)			
			{
				if (nstage > weirheight && weirheight > 0)
				{
					// CALCULATE WEIR OVERFLOW  
					h = nstage - weirheight;	// ft 

					double nweir_max = h * BMParea;	// ft3												// h is the lookup index1 for weir coefficient

					if (weirtype == 1)
					{
						//Chow Equation, Replaces Linsey lookup table for weir coefficient
						double c_weir = 0.0;

						if (h / weirheight < 10)
							c_weir = 3.27 + 0.4 * h / weirheight;
						else
							c_weir = 5.68 * pow(1 + weirheight / h, 1.5);

						nweir = c_weir * weirwidth * sqrt(h * h * h);  //cfs 
						nweir = nweir * 3600 / nivl;	//ft3/ivl
						if (nweir > nweir_max)
							nweir = nweir_max;
					}
					else // weirtype == 2
					{
						weirangle = min(179.9, max(1.0, weirangle));
						r__1      = 3.141592653 * weirangle / 180.0;
						r__2      = tan(r__1 / 2);
						r__3      = pow(h, 2.5);
						nweir     = 2.4824 * r__2 * r__3;	//cfs 
						nweir     = nweir * 3600 / nivl;	//ft3/ivl
						if (nweir > nweir_max)
							nweir = nweir_max;
					}
				}
				else
				{
					// no weir outflow
					nweir = 0.0;
				}

				if (nweir > 0)
				{
					// RECALCULATE VOLUME
					r__1    = nvolume - nweir;	//ft3
					nvolume = max(r__1, 0.0);

					//update the stage
					nstage  = nvolume / BMParea;  //ft
					if (nstage < SMALLNUM)
					{
						nAETsurf += nstage * 12; //ft3
						nstage = 0.0;
						nvolume = 0.0;
					}
				}
			}
			else
			{
				nAETsurf += nstage * 12; //ft3
				nweir = 0.0;
				nstage = 0.0;
				nvolume = 0.0;
			}
		} 
		else  // if there is no water currently in the basin... 
		{
			nAETsurf += nstage * 12;//all water evaporated (in/ivl)
			norifice = 0.0;	//no orifice outflow
			nweir    = 0.0;	//no weir outflow
			nstage   = 0.0;	//dry surface
			nvolume  = 0.0;	//no surface water
		}

		// accumulate the values
		temporifice += norifice;							// ft3/hr
		tempweir    += nweir;								// ft3/hr
		tempudout   += (nudout   / 12 * BMParea);			// ft3/hr
		tempinfilt  += (ninfilt  / 12 * BMParea);			// ft3/hr
		tempperc    += (nperc    / 12 * BMParea);			// ft3/hr
		tempseepage += (nseepage / 12 * BMParea);			// ft3/hr
		tempAET     += ((nAETsurf+nAET) / 12 * BMParea);	// ft3/hr
    }

	// OUTPUT PARAMETERS  
	ostage    = nstage;					//ft
    osa       = nsa;					//in
	ostorage  = nstorage;				//in
    ovolume   = nvolume;				//ft3	
	orifice   = temporifice / 3600.0;	//cfs	
	weir      = tempweir    / 3600.0;	//cfs	
	udout     = tempudout   / 3600.0;	//cfs	
	infilt    = tempinfilt  / 3600.0;	//cfs	
	perc      = tempperc    / 3600.0;	//cfs	
	AET       = tempAET     / 3600.0;	//cfs	
	seepage   = tempseepage / 3600.0;	//cfs	

	return;
}

void CBMPRunner::bmp_b(int nInfiltMethod,int nGAindex, bool underdrain_on,int timestep,
	 double oinflow,double BMPdepth,double BMPwidth,double BMPlength,double slope1,
	 double slope2,double slope3,double man_n,double soildepth,double soilporosity,
	 double finalf,double vegparma,double holtpar,double udfinalf,double udsoildepth,
	 double udsoilporosity,double FC,double WP,double ETrate,double& AET,double& perc,
	 double& ovolume,double& ostage,double& infilt,double& channel,double& weir,
	 double& osa,double& ostorage,double& udout,double& seepage)
{
	double bot_area = BMPlength*BMPwidth;	// bottom area (ft2)

	if ((bot_area*BMPdepth) <= 0 || (slope1*slope2) <= 0 || man_n <= 0)
		return;

	int nivl = 60 / timestep;	// number of intervals per hour 
	double g = 32.2;			// ft/sec2
	double r__1 = 0., r__2 = 0., h = 0., f = 0.;
	double nweir = 0.,nchannel = 0.,nudout = 0.,ninfilt = 0.,nperc = 0.;
	double nseepage = 0.,nAET = 0.,nAETsurf = 0.,overflow_max = 0.;
	double tempweir = 0.,tempchannel = 0.,tempudout = 0.,tempinfilt = 0.,tempperc = 0.;
	double tempseepage = 0.,tempAET = 0.;
	double x_area = 0., a = 0., b = 0., c = 0., d = 0.;

	//calculate the top width (ft)
	double top_width = (BMPdepth/slope1 + BMPdepth/slope2 + BMPwidth); 
	
	//calculate the maximum surface area (ft2)
	double s_area_max = top_width * BMPlength;	 

	// at the begining of the timestep
	double AWC      = soildepth*s_area_max*(FC - WP);//available water capacity for plants only (ft3)
	double nvolume  = ovolume;				//available water in the water column (ft3)
	double nsa      = osa/12.0*s_area_max;	//available pores of the soil column (ft3)
	double nstorage = ostorage/12.0*bot_area;//available pores of the underdrain column (ft3)
	double inflow   = oinflow / nivl;		//ft3/hr to ft3/ivl
	double nstage   = ostage;				//ft

	// calculate surface area (ft2)
	double sur_area = (nstage/slope1 + nstage/slope2 + BMPwidth) * BMPlength;	

	if (sur_area > s_area_max)
		sur_area = s_area_max;

	//calculate the max capacity of the water column (ft3)
	double vol_max = BMPdepth * (top_width + BMPwidth) / 2.0 * BMPlength; 

	//calculate the max available space in soil column (ft3)
	double nsamax = soildepth * s_area_max * (soilporosity - WP); 

	if (nsamax < SMALLNUM)
		nsamax = 0.0;

	//calculate the max available space in under-drain column (ft3)
	double nstoragemax = udsoildepth*bot_area*udsoilporosity;

	if (nstoragemax < SMALLNUM)
		nstoragemax = 0.0;

	for (int i=0; i<nivl; i++) 
	{
		ninfilt  = 0.0;		// surface infiltration (ft3/ivl)
		nchannel = 0.0;		// channel outflow (ft3/ivl)
		nweir    = 0.0;		// weir outflow (ft3/ivl)
		nudout   = 0.0;		// underdrain outflow (ft3/ivl)
		nperc    = 0.0;		// percolation to underdrain storage (ft3/ivl)
		nAETsurf = 0.0;		// actual evapotranspiration from the surface (ft3/ivl) 
		nAET     = 0.0;		// actual evapotranspiration (ft3/ivl)
		nseepage = 0.0;		// deep percolation to groundwater storage (ft3/ivl)

		float tStep = timestep * 60.0;//sec/ivl
		float avdepth = nvolume/s_area_max;//ft
		float finflow = inflow/s_area_max*tStep;//ft/sec

		//update the volume
		nvolume += inflow;	// ft3

		//update cross-section area (ft2), stage (ft), and surface area (ft2)
		UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
			slope1, slope2, x_area, nstage, sur_area);

		//WATER COLUMN
		if (nstage > SMALLNUM)			
		{
			// calculate the evaporation from the surface (ft3/ivl)
			nAETsurf = min(nvolume, ETrate/12.0/nivl*sur_area);
				
			//update the volume
			nvolume -= nAETsurf;// ft3

			//update cross-section area, stage, and surface area
			UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
				slope1, slope2, x_area, nstage, sur_area);

			if (nstage > SMALLNUM)	
			{
				if (nInfiltMethod == 0)
				{
					// calculate infiltration potential per hour using Holtan Method (in/hr)
					f = holtpar * vegparma * pow(nsa, 1.4) + finalf;

					// f is converted from in/hr to in/ivl
					f /= nivl;
				}
				else
				{
					float surfEvap = nAETsurf / (s_area_max * tStep);//ft/sec
					
					// calculate infiltration rate per hour using Green-Ampt Method (ft/sec)
					f = infil_getInfil(nGAindex,GREEN_AMPT,tStep,(finflow - surfEvap),avdepth);

					//convert infiltration rate from ft/sec to in/ivl
					f = f / 12.0 * tStep;
				}

				//infiltration is the min of potential infiltration or available water in basin
				ninfilt = min(f/12.0*s_area_max, nvolume);	//ft3/ivl

				//update the volume
				nvolume -= ninfilt;	// ft3

				//update cross-section area, stage, and surface area
				UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
					slope1, slope2, x_area, nstage, sur_area);

				if (nstage <= SMALLNUM)	
				{
					nAETsurf += nvolume;//add small portion to evaporation
					nstage = 0.0;	//dry surface
					nvolume = 0.0;	//no surface water
					nweir = 0.0;	//no weir outflow
					nchannel = 0.0;	//no orifice outflow
				}
			}
			else
			{
				nAETsurf += nvolume;	//ft3
				nstage = 0.0;			//dry surface
				nvolume = 0.0;			//no surface water
				ninfilt = 0.0;			//no water available
				nweir = 0.0;			//no weir outflow
				nchannel = 0.0;			//no orifice outflow
			}
		} 
		else  // if there is no water currently in the basin... 
		{
			nAETsurf = nvolume;
			nstage = 0.0;	//dry surface
			nvolume = 0.0;	//no surface water
			ninfilt = 0.0;	//no water available
			nweir = 0.0;	//no weir outflow
			nchannel = 0.0;	//no orifice outflow
		}

		//The soil column width is equal to the channel top width and infiltration 
		//occurs only through the current water surface area. Therefore, it is important 
		//to adjust the infiltrated water depth (ninfilt).
		//double adjustinfilt = sur_area/s_area_max;  
		
		// SOIL LAYER
		if (nsamax > 0)
		{
			// extra water after saturation (ft3)
			double excesswater = ninfilt - nsa;
			
			//sub-surface evapotranspiration (ft3/ivl)
			nAET = ETrate/12.0/nivl*s_area_max - nAETsurf;

			if (nAET > (nsamax + excesswater))
				nAET = nsamax + excesswater;
				
			if (excesswater >= 0) // saturation 
			{
				nperc = min(finalf/12.0/nivl*s_area_max, excesswater + nsamax - nAET - AWC);
				nperc = max(0.0, nperc);//ft3

				double excessinfilt = max(0.0, excesswater - nAET - nperc);
				ninfilt -= excessinfilt;//ft3

				//update the volume
				nvolume += excessinfilt;// ft3

				//update cross-section area, stage, and surface area
				UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
					slope1, slope2, x_area, nstage, sur_area);

				nsa = nsa - ninfilt + nperc + nAET;
			}
			else   // no saturation
			{
				if (-excesswater < nsamax)	// there is water to percolate
				{
					r__1 = (nsamax + excesswater) / nsamax;
					r__1 = (0.5 + 0.5 * r__1);
					r__1 = max(0.0, r__1);

					nperc = min(r__1 * finalf/12.0/nivl*s_area_max, excesswater + nsamax - nAET - AWC);	
					nperc = max(0.0, nperc);
					nsa   = nsa - ninfilt + nperc + nAET;
				}
				else
				{
					// no percolation to underdrain storage
					nperc = 0.0; 
					nsa   = nsamax;
				}
			}
		}
		else
		{
			// there is no storage capacity
			nAET  = 0.0;
			nsa   = 0.0;
			nperc = ninfilt; //ft3
		}

		nsa = max(nsa, 0.0);	
		nsa = min(nsa, nsamax);	

		//since under drain column is just below the channel bed and it is assumed 
		//that some sort of carrier (mesh or pipe etc) carry percolated water from 
		//the channel sides to the under drain column. Therefore, it is important 
		//to adjust the percolated water depth (nperc).
		//double adjustperc = s_area_max/bot_area;  
		
		// UNDERDRAIN SOIL LAYER
		if(nstoragemax > 0)
		{
			// extra water after saturation (ft3)
			double excesswater = nperc - nstorage;

			if (excesswater >= 0) // saturation 
			{
				nseepage = min(udfinalf/12.0/nivl*bot_area, excesswater + nstoragemax);
				nudout   = 0.0;

				if(underdrain_on)
				{
					nudout = max(0.0, excesswater - nseepage);
				}
				else
				{
					double excessperc = max(0.0, excesswater - nseepage);
					nperc -= excessperc;

					//check if soil column has space to hold this excess water
					if(nsa > 0)
					{
						//check the available space
						if(nsa < excessperc)
						{
							ninfilt -= (excessperc - nsa);
							nsa = 0.0;	// saturated (no more space)
							
							//update the volume
							nvolume += (excessperc - nsa);// ft3

							//update cross-section area, stage, and surface area
							UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
								slope1, slope2, x_area, nstage, sur_area);
						}
						else
						{
							nsa -= excessperc;
						}
					}
					else
					{
						//return this excess water to water column
						ninfilt -= excessperc;

						//update the volume
						nvolume += excessperc;// ft3

						//update cross-section area, stage, and surface area
						UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
							slope1, slope2, x_area, nstage, sur_area);
					}
				}

				nstorage = nstorage - nperc + nseepage + nudout;
			}
			else	// no saturation (no underdrain outflow)
			{
				nudout = 0.0;

				if (-excesswater < nstoragemax)	// there is water to percolate
				{
					nseepage = min(udfinalf/12.0/nivl*bot_area, excesswater + nstoragemax);
					nstorage = nstorage - nperc + nseepage;
				}
				else
				{
					nseepage = 0.0; 
					nstorage = nstoragemax;
				}
			}
		}
		else
		{
			// no underdrain outflow
			nudout   = 0.0;
			nstorage = 0.0;
			nseepage = nperc;// percolate will go to groundwater storage
		}

		nstorage = max(nstorage, 0.0);	
		nstorage = min(nstorage, nstoragemax);
		
		//WATER COLUMN
		if (nstage > SMALLNUM)			
		{
			if (nstage > BMPdepth)
			{
				// CALCULATE SWALE CAPACITY OVERFLOW USING WEIR OVERFLOW FROM SIDE OF CHANNEL  
				h = nstage - BMPdepth;	// ft 

				double nweir_max = h * s_area_max;	// ft3												// h is the lookup index1 for weir coefficient

				//Chow Equation, Replaces Linsey lookup table for weir coefficient
				double c_weir = 0.0;

				if (h / BMPdepth < 10)
					c_weir = 3.27 + 0.4 * h / BMPdepth;
				else
					c_weir = 5.68 * pow(1 + BMPdepth / h, 1.5);
				
				nweir = c_weir * BMPlength * sqrt(h * h * h);  //cfs 
				nweir = nweir * 3600 / nivl;	//ft3/ivl
				if (nweir > nweir_max)
					nweir = nweir_max;

				//update the volume
				nvolume -= nweir;// ft3

				//update cross-section area, stage, and surface area
				UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
					slope1, slope2, x_area, nstage, sur_area);

				if (nstage <= SMALLNUM)	
				{
					nAETsurf += nvolume;//add small portion to evaporation
					nstage = 0.0;	//dry surface
					nvolume = 0.0;	//no surface water
					nchannel = 0.0;	//no orifice outflow
				}
			}
			else
			{
				// no weir outflow
				nweir = 0.0;
			}

			// COMPUTE CHANNEL OUTFLOW  
			// solve for velocity using manning's equation
			
			h = nstage;
			r__1 = h / slope1;
			r__2 = h / slope2;
			
			double wet_p = BMPwidth + sqrt(h*h + r__1*r__1) + sqrt(h*h + r__2*r__2); //ft
			
			double HR = x_area / wet_p; //ft
			
			double velocity = 1.49 / man_n * pow(HR, 2/3) * pow(slope3, 0.5); //ft/sec

			// solve for channel flow using continuity and manning's velocity
			nchannel = velocity * x_area;	//cfs 
			nchannel = nchannel * tStep;	//ft3/ivl
			if (nchannel > nvolume)
				nchannel = nvolume;
					
			//update the volume
			nvolume -= nchannel;// ft3

			//update cross-section area, stage, and surface area
			UpdateXareaStageSarea(nvolume, vol_max, s_area_max, BMPdepth, BMPwidth, BMPlength,
				slope1, slope2, x_area, nstage, sur_area);
		} 
		else  // if there is no water currently in the basin... 
		{
			nAETsurf += nvolume;//ft3
			nstage = 0.0;		//dry surface
			nvolume = 0.0;		//no surface water
			nweir    = 0.0;		//no weir outflow
			nchannel = 0.0;		//no orifice outflow
		}

		// accumulate the values
		tempchannel += nchannel;		// ft3/hr
		tempweir    += nweir;			// ft3/hr
		tempudout   += nudout;			// ft3/hr
		tempinfilt  += ninfilt;			// ft3/hr
		tempperc    += nperc;			// ft3/hr
		tempseepage += nseepage;		// ft3/hr
		tempAET     += (nAET+nAETsurf);	// ft3/hr
    }

	// OUTPUT PARAMETERS  
	ostage    = nstage;					//ft
    osa       = nsa/s_area_max*12.0;	//in
	ostorage  = nstorage/bot_area*12.0;	//in
    ovolume   = nvolume;				//ft3	
	channel   = tempchannel / 3600.0;	//cfs	
	weir      = tempweir    / 3600.0;	//cfs	
	udout     = tempudout   / 3600.0;	//cfs	
	infilt    = tempinfilt  / 3600.0;	//cfs	
	perc      = tempperc    / 3600.0;	//cfs	
	AET       = tempAET     / 3600.0;	//cfs	
	seepage   = tempseepage / 3600.0;	//cfs	

	return;
}

void CBMPRunner::UpdateXareaStageSarea(double nvolume,double vol_max,double s_area_max,
	 double BMPdepth,double BMPwidth,double BMPlength,double slope1,double slope2,
	 double& x_area,double& nstage,double& sur_area)
{
		//update cross-section area and stage
		if (nvolume > vol_max)
		{
			x_area = nvolume / BMPlength;// ft2
			
			double overflow_max = nvolume - vol_max;// ft3
			nstage = overflow_max / s_area_max + BMPdepth;// ft
		}
		else
		{
			x_area = nvolume / BMPlength; // ft2
			
			// quadratic formula coefficients for calculating stage
			double a = slope1 + slope2;
			double b = BMPwidth * 2 * slope1 * slope2;
			double c = x_area * -2 * slope1 * slope2;
			double d = b * b - 4 * a * c;// (b^2 - 4ac)
			nstage = (-b + pow(d, 0.5)) / (2 * a);// (-b + (b^2 - 4ac)^0.5 ) / 2a
		}
		
		// update surface area (ft2)
		sur_area = (nstage/slope1 + nstage/slope2 + BMPwidth) * BMPlength;	

		if (sur_area > s_area_max)
			sur_area = s_area_max;

	return;
}

void CBMPRunner::advect(double imat,double svol,double sro,double evol,double ero,
	 double delts,double crrat,double& conc,double& romat)
{
	double js = 0.0;
    double sconc = conc;// save starting concentration 

	if (fabs(sro) > fThreshold)
	{
		double rat = svol / (sro * delts);
		if (rat < crrat)
			js = rat / crrat;
		else
			js = 1.0;
	}
    
	double cojs = 1.0 - js;
	double srovol = js * sro * delts;
	double erovol = cojs * ero * delts;

    if (fabs(evol) > fThreshold)
	{
		// reach/res contains water; perform advection normally  
		// calculate new concentration of material in reach/res based  
		// on quantity of material entering during interval (imat),  
		// weighted volume of outflow based on conditions at start of  
		// ivl (srovol), and weighted volume of outflow based on  
		// conditions at end of ivl (erovol)  
		conc = (imat + sconc * (svol - srovol)) / (evol + erovol);				
		if (conc < 0)	
			conc = 0;
		// calculate total amount of material leaving reach/res during  
		// interval  

		if (ero > 0.001)//cfs
		{
			romat = srovol * sconc + erovol * conc;									
			if (romat < 0)	
				romat = 0;
		}
		else
		{
			romat = 0;	
		}
    }
	else
	{
		// reach/res has gone dry during the interval; set conc equal to  
		// an undefined value  
		// conc = -1e10;
		conc = 0;
		// calculate total amount of material leaving during interval;  
		// this is equal to material inflow + material initially present  

		if (ero > 0.001)	//cfs
		{
			romat = imat + sconc * svol;	
			if (romat < 0)	
				romat = 0;
		}
		else
		{
			romat = 0;
		}
    }
	return;
}

void CBMPRunner::RunModel(int nRunMode)
{
	//define local variables
	int  i=0, j=0, k=0;	//loop counters
	int  NPOL  = 0;		//number of pollutants (before spliting sediment);
	int  NWQ   = 0;		//number of pollutants (after spliting sediment);
	int  NLAND = 0;		//number of land use types
	int  NBMP  = 0;		//number of BMPs
	int  NBMPtype = 0;  //number of unique BMP types
	long t=0, N=0;		//simulation counters (N = total)
	double lfNumOfYears = 0.0;		//number of years
	double lfWetDaysPerYear = 0.0;	//number of wet days per year
	double lfAll = 0.0;				//total number of seconds for the simulation period

	SYSTEMTIME tm;					//system time
	COleDateTime tStart, tEnd;		//date time for the start and end of simulation
	COleDateTimeSpan tSpan, tSpan0;	//time spans
	POSITION pos, pos1;				//position pointers 
	CSiteLandUse *pSiteLU;			//land use pointer
	CSitePointSource *pSitePS;		//point source pointer
	CBMPSite *pBMPSite, *pBMPSiteUp;//BMP site pointers
	US_BMPSITE *pUS;				//upstream BMP pointer

	//check the run counter to update the progress bar
	if (nRunMode == RUN_OPTIMIZE)
		optcounter++;
	else if (nRunMode == RUN_OUTPUT)
		outcounter++;

	//get the system time at the beginning of the simulation
	if (nRunMode == RUN_INIT)
	{
		GetLocalTime(&tm);
		time_i = COleDateTime(tm);
	}

	NWQ = pBMPData->nNWQ;
	NPOL = pBMPData->nPollutant;
	NLAND = pBMPData->siteluList.GetCount();
	NBMP = pBMPData->routeList.GetCount();
	NBMPtype = pBMPData->nBMPtype;
	//initialize the BMP cost
	for (i=0; i<NBMPtype; i++)
		pBMPData->m_pBMPcost[i].m_lfCost = 0.0;
	tStart = pBMPData->startDate;
	tEnd = pBMPData->endDate;
	tSpan0 = tEnd - tStart;
	N = (long)tSpan0.GetTotalHours() + 24;
	lfAll = tSpan0.GetTotalSeconds();
	//initialize the progress bar
	pWndProgress->SetRange(0, 100);			 
	pWndProgress->SetText("");

	//define local dynamic arrays
	int    *counter_p      = new int[NBMP];				
    double *BmpFlowInput   = new double[NBMP];
	double *bmpoflow       = new double[NBMP];
	double *vol_p          = new double[NBMP];
	double *bmpvol_p       = new double[NBMP];
	double *osa_p          = new double[NBMP];
	double *ostorage_p     = new double[NBMP];
	double *weir_p         = new double[NBMP];
	double *orifice_p      = new double[NBMP];
	double *infilt_p       = new double[NBMP];
	double *undrain_p      = new double[NBMP];
	double *seepage_p      = new double[NBMP];
	//output variables
	double *bmpvol_s       = new double[NBMP];
	double *bmpstage_s     = new double[NBMP];
    double *BmpFlowInput_s = new double[NBMP];
	double *weir_s         = new double[NBMP];
	double *orifice_s      = new double[NBMP];
	double *bmpudout_s     = new double[NBMP];
	double *bmpbypass_s    = new double[NBMP];
	double *bmpoutflow_s   = new double[NBMP];
	double *infilt_s       = new double[NBMP];
	double *perc_s         = new double[NBMP];
	double *AET_s          = new double[NBMP];
	double *seepage_s      = new double[NBMP];
//	double *usstorage_s    = new double[NBMP];
//	double *udstorage_s    = new double[NBMP];
	double *bmpoflow_w     = new double[NBMP];			
	double *bmpoflow_o     = new double[NBMP];			
	double *bmpoflow_ud    = new double[NBMP];			
	double *bmpoflow_ut    = new double[NBMP];	//untreated for the exceeding design drainage area
	double *rbsedtot	   = new double[NBMP];

	double *BmpWqInput     = new double[NBMP*NWQ];
	double *bmpc           = new double[NBMP*NWQ];
	double *bmpc2          = new double[NBMP*NWQ];
	double *bmpqconc_sand  = new double[NBMP*NWQ];	//lb of qual per lb of sand
	double *bmpqconc_silt  = new double[NBMP*NWQ];	//lb of qual per lb of silt
	double *bmpqconc_clay  = new double[NBMP*NWQ];	//lb of qual per lb of clay
	double *bmpmassout     = new double[NBMP*NWQ];
	double *bmpmassout_w   = new double[NBMP*NWQ];
	double *bmpmassout_o   = new double[NBMP*NWQ];
	double *bmpmassout_ud  = new double[NBMP*NWQ];
	double *bmpmassout_ut  = new double[NBMP*NWQ];//untreated for the exceeding design drainage area
	double *bmudconc       = new double[NBMP*NWQ];
	double *romat_o2       = new double[NBMP*NWQ];
	double *romat_ut2      = new double[NBMP*NWQ];
	double *rbsed		   = new double[NBMP*NWQ];

	//output variables
	double *BmpWqInput_s   = new double[NBMP*NPOL];
	double *bmpmass_w_s    = new double[NBMP*NPOL];
	double *bmpmass_o_s    = new double[NBMP*NPOL];
	double *bmpmass_ud_s   = new double[NBMP*NPOL];
	double *bmpmass_ut_s   = new double[NBMP*NPOL];
	double *bmpmassout_s   = new double[NBMP*NPOL];
	double *bmpconcout_s   = new double[NBMP*NPOL];

	// evaluation factor calculation
	int	   *nExceedFlag    = new int[NBMP];			// number of times flow exceeds threshold flow 
	double *bmpTotalFlow   = new double[NBMP];		// cumulative hourly flows (ft3/simulation)
	double *bmpAAFlowVol   = new double[NBMP];		// annual average flow (ft3/yr)
	double *bmpPkDisFlow   = new double[NBMP];		// peak flow per simulation (cfs) 
	double *bmpFlowExcFreq = new double[NBMP];		// exceeding frequency of flow compared to threshold flow (/yr) 
	double *bmpTotalLoad   = new double[NBMP*NPOL];	// cumulative hourly load (lb/simulation)
	double *bmpAALoad      = new double[NBMP*NPOL];	// annual average load (lb/yr)
	double *bmpAAConc      = new double[NBMP*NPOL];	// annual average concentration (mg/l)
	double *bmpMAConc      = new double[NBMP*NPOL];	// maximum running average concentration (mg/l)
	//optional
	int	   *nExceedFlow    = new int[NBMP];			// number of times flow exceeds threshold flow 
	int	   *nExceedConc    = new int[NBMP*NPOL];	// number of times conc exceeds threshold conc during a wet day 
	double *lfSumFlow	   = new double[NBMP];		// ft3/day
	double *lfSumMass	   = new double[NBMP*NPOL];	// lb/day
	double *bmpConcExcDays = new double[NBMP*NPOL];	// number of days conc exceeds threshold conc per year 
	//---------------------------------------------------------------------

	for(i=0; i<NBMP; i++)
	{
		// get model parameters
		pos = pBMPData->routeList.FindIndex(i);
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetAt(pos);

		double WP			  = pBMPSite->m_lfWPoint;	
		double soildepth      = pBMPSite->m_lfSoilDepth;//ft
		double soilporosity   = pBMPSite->m_lfPorosity;
		double udsoildepth    = pBMPSite->m_lfUndDepth;	//ft
		double udsoilporosity = pBMPSite->m_lfUndVoid;	

		bmpvol_s[i]       = 0.0;
		vol_p[i]          = 0.0;
		bmpvol_p[i]       = 0.0;
		infilt_s[i]       = 0.0;
		perc_s[i]         = 0.0;
		AET_s[i]          = 0.0;
		seepage_s[i]      = 0.0;
//		usstorage_s[i]    = 0.0;
//		udstorage_s[i]    = 0.0;
		osa_p[i]          = soildepth * (soilporosity - WP) * 12; //max water holding capacity (in)
		//assume underdrain storage media can release all water under gravity
		ostorage_p[i]     = udsoildepth * udsoilporosity * 12;	 //max water holding capacity (in)
		weir_p[i]         = 0.0;
		orifice_p[i]      = 0.0;
		infilt_p[i]       = 0.0;
		undrain_p[i]      = 0.0;
		seepage_p[i]      = 0.0;
		weir_s[i]         = 0.0;
		orifice_s[i]      = 0.0;
		bmpoutflow_s[i]   = 0.0;
		bmpoflow[i]       = 0.0;
		bmpoflow_w[i]     = 0.0;		
		bmpoflow_o[i]     = 0.0;		
		bmpoflow_ud[i]    = 0.0;		
		bmpoflow_ut[i]    = 0.0;
		BmpFlowInput_s[i] = 0.0;
		bmpstage_s[i]     = 0.0;
		bmpudout_s[i]     = 0.0;
		bmpbypass_s[i]    = 0.0;
		counter_p[i]      = 0;					
		bmpTotalFlow[i]	  =	0.0;
		bmpAAFlowVol[i]	  =	0.0;	
		bmpPkDisFlow[i]	  = 0.0;	
		bmpFlowExcFreq[i] =	0.0;	
		nExceedFlow[i]    =	0;	
		nExceedFlag[i]    =	0;	
		rbsedtot[i]       = 0.0;		
		lfSumFlow[i]	  = 0.0;
	}

	for(i=0; i<NBMP*NWQ; i++)
	{
		bmpc[i]         = 0.0;
		bmpc2[i]        = 0.0;
		bmpqconc_sand[i]= 0.0;
		bmpqconc_silt[i]= 0.0;
		bmpqconc_clay[i]= 0.0;
		bmpmassout[i]   = 0.0;
		bmpmassout_w[i] = 0.0;
		bmpmassout_o[i] = 0.0;
		bmpmassout_ud[i]= 0.0;
		bmpmassout_ut[i]= 0.0;
		bmudconc[i]     = 0.0;
		romat_o2[i]     = 0.0;		
		romat_ut2[i]    = 0.0;		
		rbsed[i]        = 0.0;	
	}

	for(i=0; i<NBMP*NPOL; i++)
	{
		BmpWqInput_s[i] = 0.0;
		bmpmass_w_s[i] = 0.0;
		bmpmass_o_s[i] = 0.0;
		bmpmass_ud_s[i] = 0.0;
		bmpmass_ut_s[i] = 0.0;
		bmpmassout_s[i] = 0.0;
		bmpconcout_s[i] = 0.0;
		bmpTotalLoad[i] = 0.0;
		bmpAALoad[i]    = 0.0;		
		bmpAAConc[i]    = 0.0;		
		bmpMAConc[i]    = 0.0;		
		nExceedConc[i]  = 0;
		lfSumMass[i]	= 0.0;
		bmpConcExcDays[i] = 0.0;
	}

	// calculate BMP cost for each BMP site
	// Cost ($) = (LinearCost * Length + AreaCost * Area + TotalVolumeCost * TotalVolume 
	// + MediaVolumeCost * SoilMediaVolume + UnderDrainVolumeCost * UnderDrainVolume
	// + Unitcost * Count + ConstantCost) * (1 + PercentCost / 100)

	int ii = 0;
	double totalCost = 0.0;
	pos = pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

		pBMPSite->m_lfCost = 0.0;

		pBMPSite->m_lfSurfaceArea = 0.0;
		pBMPSite->m_lfExcavatnVol = 0.0;
		pBMPSite->m_lfSurfStorVol = 0.0;
		pBMPSite->m_lfSoilStorVol = 0.0;
		pBMPSite->m_lfUdrnStorVol = 0.0;

		double lfLinearCost = pBMPSite->m_costParam.m_lfLinearCost;			
		double lfAreaCost = pBMPSite->m_costParam.m_lfAreaCost;            				
		double lfTotalVolumeCost = pBMPSite->m_costParam.m_lfTotalVolumeCost;     				
		double lfMediaVolumeCost = pBMPSite->m_costParam.m_lfMediaVolumeCost;     				
		double lfUnderDrainVolumeCost = pBMPSite->m_costParam.m_lfUnderDrainVolumeCost;				
		double lfConstantCost = pBMPSite->m_costParam.m_lfConstantCost;        				
		double lfPercentCost = pBMPSite->m_costParam.m_lfPercentCost;  
		
		// initialize BMP Cost parameters			 
		double lfLengthExp = pBMPSite->m_costParam.m_lfLengthExp;
		double lfAreaExp = pBMPSite->m_costParam.m_lfAreaExp;
		double lfTotalVolExp = pBMPSite->m_costParam.m_lfTotalVolExp;
		double lfMediaVolExp = pBMPSite->m_costParam.m_lfMediaVolExp;
		double lfUDVolExp = pBMPSite->m_costParam.m_lfUDVolExp;

		double soilporosity   = pBMPSite->m_lfPorosity;
		double udsoilporosity = pBMPSite->m_lfUndVoid;	
		double sqft2acre = 2.295675e-005;
		double cuft2acft = 2.295675e-005;

		//assign array size
		if (pBMPSite->m_pConc != NULL)
			delete []pBMPSite->m_pConc;

		if (pBMPSite->m_nPolRotMethod*NWQ > 0)
			pBMPSite->m_pConc = new double[pBMPSite->m_nPolRotMethod*NWQ];
		
		//initialize the values
		for (i=0; i<pBMPSite->m_nPolRotMethod*NWQ; i++)
			pBMPSite->m_pConc[i] = 0.0;

		// initialize running average parameters
//		if (pBMPSite->m_factorList.GetCount() > 0)
		{
			int qsize = 24;	// hourly time step 
			double value1 = 0.0;
			while (!pBMPSite->qFlow.empty())
				pBMPSite->qFlow.pop();
			while (pBMPSite->qFlow.size() != qsize)
				pBMPSite->qFlow.push(value1);

			pBMPSite->m_lfThreshFlow = 0;
			if (pBMPSite->m_RAConc != NULL)
				delete []pBMPSite->m_RAConc;
			//pBMPSite->m_RAConc = new POLLUT_RAConc[NWQ];
			pBMPSite->m_RAConc = new POLLUT_RAConc[NPOL];	//don't split TSS

			//for (i=0; i<NWQ; ++i)
			for (i=0; i<NPOL; ++i)
			{
				pBMPSite->m_RAConc[i].m_nRDays = 0;
				pBMPSite->m_RAConc[i].m_lfRFlow = NULL;
				pBMPSite->m_RAConc[i].m_lfRLoad = NULL;
				pBMPSite->m_RAConc[i].m_lfThreshConc = 0.0;	

				//initialze
				while (!pBMPSite->m_RAConc[i].qMass.empty())
					pBMPSite->m_RAConc[i].qMass.pop();
				while (pBMPSite->m_RAConc[i].qMass.size() != qsize)
					pBMPSite->m_RAConc[i].qMass.push(value1);

				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1 && ef->m_nFactorType == FEF)
						pBMPSite->m_lfThreshFlow = ef->m_lfThreshold;
					if ((ef->m_nFactorGroup-1) == i && ef->m_nFactorType == CEF)	
						pBMPSite->m_RAConc[i].m_lfThreshConc = ef->m_lfConcThreshold;
					if ((ef->m_nFactorGroup-1) == i && ef->m_nFactorType == MAC)
						pBMPSite->m_RAConc[i].m_nRDays = ef->m_nCalcDays;		 
				}
				if (pBMPSite->m_RAConc[i].m_nRDays > 0)
				{
					pBMPSite->m_RAConc[i].m_lfRFlow = new double[pBMPSite->m_RAConc[i].m_nRDays*24];
					pBMPSite->m_RAConc[i].m_lfRLoad = new double[pBMPSite->m_RAConc[i].m_nRDays*24];
					for (j=0; j<pBMPSite->m_RAConc[i].m_nRDays*24; j++)
					{
						pBMPSite->m_RAConc[i].m_lfRFlow[j] = 0.0;
						pBMPSite->m_RAConc[i].m_lfRLoad[j] = 0.0;
					}
				}
			}
		}

		double volsediment = 0.0;

		if (pBMPSite->m_nBMPClass == CLASS_A)
		{
			BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;

			// copy data
			int nIndex = pBMPSite->m_nGAInfil_Index;
			CopyGAInfil(&pBMP->m_pGAInfil, &GAInfil[nIndex]);

			int    releasetype    = pBMP->m_nORelease;		    
			double basinlength    = pBMP->m_lfBasinLength;		//ft
			double basinwidth     = pBMP->m_lfBasinWidth;		//ft
			double weirheight     = pBMP->m_lfWeirHeight;		//ft
			double soildepth      = pBMPSite->m_lfSoilDepth;	//ft	
			double udsoildepth    = pBMPSite->m_lfUndDepth;		//ft

			double BMParea = basinlength * basinwidth;			//ft^2	 
			
			// check if this BMP is cistern or rainbarrel, if so then 
			// basinlength = diameter (ft) and basinwidth = number of devices
			if (releasetype == 1 || releasetype == 2)
				//BMParea = 3.142857/4.0*pow(basinlength,2)*basinwidth;	// ft2 
				BMParea = 3.142857/4.0*pow(basinlength,2);	// ft2 
			
			double BMPdepth2 = weirheight + soildepth + udsoildepth;	// ft	 
			
			//initilaize the cost
			pBMPSite->m_lfCost = 0.0;
			
			if (BMParea > 0 && BMPdepth2 > 0)		 
				
				//Cost ($) = ((LinearCost)Length^(LengthExp) 
				//+ (AreaCost)Area^(AreaExp)  
				//+ (TotalVolumeCost)TotalVolume^(TotalVolExp) 
				//+ (MediaVolumeCost)SoilMediaVolume^(MediaVolExp) 
				//+ (UnderDrainVolumeCost)UnderDrainVolume^(UDVolExp) 
				//+ (Unitcost) +  (ConstantCost)) * (1+PercentCost/100)

				pBMPSite->m_lfCost = (lfLinearCost * pow(basinlength, lfLengthExp) 
				+ lfAreaCost * pow(BMParea, lfAreaExp) 
				+ lfTotalVolumeCost	* pow((BMParea * BMPdepth2), lfTotalVolExp) 
				+ lfMediaVolumeCost	* pow((BMParea * soildepth), lfMediaVolExp) 
				+ lfUnderDrainVolumeCost * pow((BMParea * udsoildepth), lfUDVolExp) 
				+ lfConstantCost) * (1 + lfPercentCost / 100);
			
			//get the total cost for all units
			pBMPSite->m_lfCost *= pBMPSite->m_lfBMPUnit;

			//assign the cost to its unique BMPtype
			for (i=0; i<NBMPtype; i++)
			{
				if (pBMPData->m_pBMPcost[i].m_strBMPType.CompareNoCase(pBMPSite->m_strType) == 0)
				{
					pBMPData->m_pBMPcost[i].m_lfCost += pBMPSite->m_lfCost;
					break;
				}
			}

			pBMPSite->m_lfSurfaceArea = BMParea * pBMPSite->m_lfBMPUnit * sqft2acre;//acre
			pBMPSite->m_lfExcavatnVol = BMParea * BMPdepth2 * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfSurfStorVol = BMParea * weirheight * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfSoilStorVol = BMParea * soildepth * soilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfUdrnStorVol = BMParea * udsoildepth * udsoilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			
			// calculate initial volume of sediment in the bed (ft3)
			volsediment = basinlength * pBMPSite->m_sediment.m_lfBEDWID
						* pBMPSite->m_sediment.m_lfBEDDEP
						* (1.0 - pBMPSite->m_sediment.m_lfBEDPOR);
		}
		else if (pBMPSite->m_nBMPClass == CLASS_B)
		{
			BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;

			// copy data
			int nIndex = pBMPSite->m_nGAInfil_Index;
			CopyGAInfil(&pBMP->m_pGAInfil, &GAInfil[nIndex]);

			double BMPlength   = pBMP->m_lfBasinLength;
			double BMPwidth    = pBMP->m_lfBasinWidth;
			double BMPdepth    = pBMP->m_lfMaximumDepth;
			double soildepth   = pBMPSite->m_lfSoilDepth;	
			double udsoildepth = pBMPSite->m_lfUndDepth;	
			double slope1      = pBMP->m_lfSideSlope1;
			double slope2      = pBMP->m_lfSideSlope2;

			//calculate the top width (ft)
			double top_width = (BMPdepth/slope1 + BMPdepth/slope2 + BMPwidth); 

			double BMParea     = BMPlength * top_width;						// ft^2		 
			double BMPdepth2   = BMPdepth + soildepth + udsoildepth;		// ft	 

			//initilaize the cost
			pBMPSite->m_lfCost = 0.0;

			if (BMParea > 0 && BMPdepth2 > 0)		 

				//Cost ($) = ((LinearCost)Length^(LengthExp) 
				//+ (AreaCost)Area^(AreaExp)  
				//+ (TotalVolumeCost)TotalVolume^(TotalVolExp) 
				//+ (MediaVolumeCost)SoilMediaVolume^(MediaVolExp) 
				//+ (UnderDrainVolumeCost)UnderDrainVolume^(UDVolExp) 
				//+ (Unitcost) +  (ConstantCost)) * (1+PercentCost/100)

				pBMPSite->m_lfCost = (lfLinearCost * pow(BMPlength, lfLengthExp) 
				+ lfAreaCost * pow(BMParea, lfAreaExp) 
				+ lfTotalVolumeCost	* pow((BMParea * BMPdepth2), lfTotalVolExp) 
				+ lfMediaVolumeCost	* pow((BMParea * soildepth), lfMediaVolExp) 
				+ lfUnderDrainVolumeCost * pow((BMParea * udsoildepth), lfUDVolExp) 
				+ lfConstantCost) * (1 + lfPercentCost / 100);
			
			//get the total cost for all units
			pBMPSite->m_lfCost *= pBMPSite->m_lfBMPUnit;

			//assign the cost to its unique BMPtype
			for (i=0; i<NBMPtype; i++)
			{
				if (pBMPData->m_pBMPcost[i].m_strBMPType.CompareNoCase(pBMPSite->m_strType) == 0)
				{
					pBMPData->m_pBMPcost[i].m_lfCost += pBMPSite->m_lfCost;
					break;
				}
			}

			pBMPSite->m_lfSurfaceArea = BMParea * pBMPSite->m_lfBMPUnit * sqft2acre;//acre
			pBMPSite->m_lfExcavatnVol = BMParea * BMPdepth2 * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfSurfStorVol = BMParea * BMPdepth * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfSoilStorVol = BMParea * soildepth * soilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			pBMPSite->m_lfUdrnStorVol = BMParea * udsoildepth * udsoilporosity * pBMPSite->m_lfBMPUnit * cuft2acft;//ac-ft
			
			// calculate initial volume of sediment in the bed (ft3)
			volsediment = BMPlength * pBMPSite->m_sediment.m_lfBEDWID
						* pBMPSite->m_sediment.m_lfBEDDEP
						* (1.0 - pBMPSite->m_sediment.m_lfBEDPOR);
		}
		else if (pBMPSite->m_nBMPClass == CLASS_C)
		{
			//  Purpose: initializes a link's state variables at start of simulation.
			BMP_C* pBMP = (BMP_C*) pBMPSite->m_pSiteProp;

			int nIndex = pBMP->m_nIndex;

			// copy data
			CopyLink(&pBMP->m_pTLink, &Link[nIndex], NWQ);
			CopyConduit(&pBMP->m_pTConduit, &Conduit[nIndex]);
			CopyTransect(&pBMP->m_pTTransect, &Transect[nIndex]);

			// check if conduit type is DUMMY
			if(Link[nIndex].xsect.type != DUMMY)
			{
				// --- assign initial flow to both ends of conduit
				k = Link[nIndex].subIndex;
				if (Conduit[k].barrels > 0)
					Conduit[k].q1 = Link[nIndex].newFlow / Conduit[k].barrels;
				else
					Conduit[k].q1 = 0.0;;

				Conduit[k].q2 = Conduit[k].q1;

				Conduit[k].q1Old = Conduit[k].q1;
				Conduit[k].q2Old = Conduit[k].q2;

				// --- find areas based on initial flow depth
				Conduit[k].a1 = xsect_getAofY(&Link[nIndex].xsect, Link[nIndex].newDepth);
				Conduit[k].a2 = Conduit[k].a1;

				// --- compute initial volume from area
				Link[nIndex].newVolume = max(0.0, Conduit[k].a1 * Conduit[k].length *
									Conduit[k].barrels);
				Link[nIndex].oldVolume = Link[nIndex].newVolume;

				// --- initialize flow
				Link[nIndex].oldFlow = Link[nIndex].q0;
				Link[nIndex].newFlow = Link[nIndex].q0;

				// --- initialize water depth
				if (Conduit[k].barrels > 0)
					Link[nIndex].newDepth = link_getYnorm(nIndex, Link[nIndex].q0 / Conduit[k].barrels);
				else
					Link[nIndex].newDepth	= 0;
				Link[nIndex].oldDepth = Link[nIndex].newDepth;
			}

			// calculate initial volume of sediment in the bed (ft3)
			volsediment = Conduit[k].length * Conduit[k].barrels 
								 * pBMPSite->m_sediment.m_lfBEDWID
								 * pBMPSite->m_sediment.m_lfBEDDEP
								 * (1.0 - pBMPSite->m_sediment.m_lfBEDPOR);
		}
		else if (pBMPSite->m_nBMPClass == CLASS_D)
		{
			// disabled for the first release (under testing)
			// Read output files
//			if (!pBMPData->ReadVFSMODFiles(nRunMode,pBMPSite->m_strID))
//			{
//				CString strErr;
//				strErr.Format("Unable to read VFSMOD output files for BMPSite ID: %s", pBMPSite->m_strID);
//				AfxMessageBox(strErr);
//				return;
//			}  
		}
		
		for (j=0; j<NWQ; j++)
		{
			if(pBMPData->nSedflag[j] == 2)		// silt 
				rbsed[ii*NWQ+j] = volsediment * pBMPSite->m_sediment.m_lfSILT_FRAC 
								* pBMPSite->m_silt.m_lfRHO;
			else if(pBMPData->nSedflag[j] == 3)	// clay 
				rbsed[ii*NWQ+j] = volsediment * pBMPSite->m_sediment.m_lfCLAY_FRAC 
								* pBMPSite->m_clay.m_lfRHO;
		}

		//check if the bmpsite has tradeoff curve
		if (pBMPSite->m_nBreakPoints > 0)
		{
			//break point id: 0 for initial, -1 for PreDev, and -2 for PostDev condition
			//get the break point index
			int nBrPtIndex = int(pBMPSite->m_lfBreakPtID + 2.0);
			
			if (nRunMode == RUN_POSTDEV)
				nBrPtIndex = 0;				//index = -2+2=0
			else if (nRunMode == RUN_PREDEV)
				nBrPtIndex = 1;				//index (-1+2=1)
			else if(nRunMode == RUN_INIT)
				nBrPtIndex = 2;				//index = 0+2=2

			pBMPSite->m_lfCost = pBMPSite->m_TradeOff[nBrPtIndex].m_lfCost*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult;

			if (pBMPSite->m_nBreakPoints > 1)
			{
				// load time series data for tradeoff curve
				if (!pBMPSite->LoadTradeOffCurveData(nBrPtIndex, pBMPData->startDate, pBMPData->endDate))
				{
					CString strErr;
					strErr.Format("Check Cost-Effectiveness Curve Data for BMPSite ID: %s", pBMPSite->m_strID);
					AfxMessageBox(strErr);
					return;
				}  
			}
		}

		ii++;
		totalCost += pBMPSite->m_lfCost;
	}
	
	// start simulation
	int dayflag = 0;
	for(t=0; t<N; t++)
	{
		// initialization
		for(i=0; i<NBMP; i++)
			BmpFlowInput[i] = 0.0;
		for(i=0; i<NBMP*NWQ; i++)
			BmpWqInput[i] = 0.0;
		
		// sum landuse flow and wq
		if (nRunMode == RUN_PREDEV)
		{
			i = 0;	// BMP array
			pos = pBMPData->routeList.GetHeadPosition();
			while (pos != NULL)
			{
				pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

				// for the pre-developed LU flowing into this bmp site, summarize the flow and wq
				if (pBMPData->nLandSimulation == 0)
				{
					//external land simulation option
					BmpFlowInput[i] += pBMPSite->m_preLU->m_pData[t*pBMPSite->m_preLU->m_nQualNum]*pBMPSite->m_lfSiteDArea;	// in-acres/hr
					
					int nIndex = 0;
					for (j=0; j<NPOL; j++)
					{
						if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
						{
							// need to split the TSS into sand, silt, and clay
							double SED_FR[3];
							SED_FR[0] = pBMPSite->m_preLU->m_lfsand_fr;
							SED_FR[1] = pBMPSite->m_preLU->m_lfsilt_fr;
							SED_FR[2] = pBMPSite->m_preLU->m_lfclay_fr;
							for (k=0; k<3; k++)
							{
								BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_preLU->m_pData[t*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea*SED_FR[k];	// lbs/hr
								nIndex++;
							}
						}
						else
						{
							BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_preLU->m_pData[t*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea;	// lbs/hr
							nIndex++;
						}
					}
				}
				else if (pBMPSite->m_pDataPreLU != NULL)
				{
					//internal land simulation option
					BmpFlowInput[i] += pBMPSite->m_pDataPreLU[t*pBMPSite->m_nQualNum] / 3630.00;	// in-acres/hr
					
					//pollutant index
					for (j=0; j<NWQ; j++)
						BmpWqInput[i*NWQ+j] += pBMPSite->m_pDataPreLU[t*pBMPSite->m_nQualNum+1+j];			// lbs/hr
				}

				//add point source data
				pos1 = pBMPSite->m_sitepsList.GetHeadPosition();
				while (pos1 != NULL)
				{
					pSitePS = (CSitePointSource*) pBMPSite->m_sitepsList.GetNext(pos1);
						
//					if (pBMPData->nLandSimulation == 0)
//					{
						BmpFlowInput[i] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum]*pSitePS->m_lfMult;	// in-acres/hr
						
						int nIndex = 0;
						for (j=0; j<NPOL; j++)
						{
							if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
							{
								// need to split the TSS into sand, silt, and clay
								double SED_FR[3];
								SED_FR[0] = pSitePS->m_lfSand;
								SED_FR[1] = pSitePS->m_lfSilt;
								SED_FR[2] = pSitePS->m_lfClay;
								for (k=0; k<3; k++)
								{
									BmpWqInput[i*NWQ+nIndex] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult*SED_FR[k];	// lbs/hr
									nIndex++;
								}
							}
							else
							{
								BmpWqInput[i*NWQ+nIndex] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult;	// lbs/hr
								nIndex++;
							}
						}
//					}
//					else
//					{
//						BmpFlowInput[i] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum]*pSitePS->m_lfMult;	// in-acres/hr

						//pollutant index
//						int nPollIndex = 0;
//						for (j=0; j<NWQ; j++)
//						{
//							if (pBMPData->nSedflag[j] == SILT || pBMPData->nSedflag[j] == CLAY)
//								nPollIndex--;
//							BmpWqInput[i*NWQ+j] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult;	// lbs/hr
//							BmpWqInput_s[i*NPOL+nPollIndex] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult;
//							nPollIndex++;
//						}
//					}
				}

				//add trade off curve data if needed for predeveloped condition
				int nBrPtIndex = 1;	//index (-1+2=1)

				if (pBMPSite->m_nBreakPoints > 0)
				{
					if (pBMPSite->m_TradeOff != NULL)
					{
						BmpFlowInput[i] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult;	// in-acres/hr
						
						int nIndex = 0;
						for (j=0; j<NPOL; j++)
						{
							if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
							{
								// need to split the TSS into sand, silt, and clay
								double SED_FR[3];
								SED_FR[0] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfSand;
								SED_FR[1] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfSilt;
								SED_FR[2] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfClay;
								for (k=0; k<3; k++)
								{
									BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum+1+j]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult*SED_FR[k];	// lbs/hr
									nIndex++;
								}
							}
							else
							{
								BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum+1+j]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult;	// lbs/hr
								nIndex++;
							}
						}
					}
				}

				i++;
			}
		}
		else
		{
			// sum landuse flow and wq
			i = 0;
			pos = pBMPData->routeList.GetHeadPosition();
			while (pos != NULL)
			{
				pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

				// for all the LUs flowing into this bmp site, summarize the flow and wq
				if (pBMPData->nLandSimulation == 0)
				{
					//external land simulation option
					pos1 = pBMPSite->m_siteluList.GetHeadPosition();
					while (pos1 != NULL)
					{
						pSiteLU = (CSiteLandUse*) pBMPSite->m_siteluList.GetNext(pos1);
						BmpFlowInput[i] += pSiteLU->m_pLU->m_pData[t*pSiteLU->m_pLU->m_nQualNum]*pSiteLU->m_lfArea;	// in-acres/hr

						int nIndex = 0;
						for (j=0; j<NPOL; j++)
						{
							if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
							{
								// need to split the TSS into sand, silt, and clay
								double SED_FR[3];
								SED_FR[0] = pSiteLU->m_pLU->m_lfsand_fr;
								SED_FR[1] = pSiteLU->m_pLU->m_lfsilt_fr;
								SED_FR[2] = pSiteLU->m_pLU->m_lfclay_fr;
								for (k=0; k<3; k++)
								{
									BmpWqInput[i*NWQ+nIndex] += pSiteLU->m_pLU->m_pData[t*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea*SED_FR[k];	// lbs/hr
									nIndex++;
								}
							}
							else
							{
								BmpWqInput[i*NWQ+nIndex] += pSiteLU->m_pLU->m_pData[t*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea;	// lbs/hr
								nIndex++;
							}
						}
					}
				}
				else if (pBMPSite->m_pDataMixLU != NULL)
				{
					//internal land simulation option
					BmpFlowInput[i] += pBMPSite->m_pDataMixLU[t*pBMPSite->m_nQualNum] / 3630.00;	// in-acres/hr

					//pollutant index
					for (j=0; j<NWQ; j++)
						BmpWqInput[i*NWQ+j] += pBMPSite->m_pDataMixLU[t*pBMPSite->m_nQualNum+1+j];	// lbs/hr
				}

				//add point source data
				pos1 = pBMPSite->m_sitepsList.GetHeadPosition();
				while (pos1 != NULL)
				{
					pSitePS = (CSitePointSource*) pBMPSite->m_sitepsList.GetNext(pos1);
					BmpFlowInput[i] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum]*pSitePS->m_lfMult;	// in-acres/hr
					int nIndex = 0;
					for (j=0; j<NPOL; j++)
					{
						if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
						{
							// need to split the TSS into sand, silt, and clay
							double SED_FR[3];
							SED_FR[0] = pSitePS->m_lfSand;
							SED_FR[1] = pSitePS->m_lfSilt;
							SED_FR[2] = pSitePS->m_lfClay;
							for (k=0; k<3; k++)
							{
								BmpWqInput[i*NWQ+nIndex] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult*SED_FR[k];	// lbs/hr
								nIndex++;
							}
						}
						else
						{
							BmpWqInput[i*NWQ+nIndex] += pSitePS->m_pDataPS[t*pSitePS->m_nQualNum+1+j]*pSitePS->m_lfMult;	// lbs/hr
							nIndex++;
						}
					}
				}

				//add trade off curve data if needed

				//get the break point index
				int nBrPtIndex = int(pBMPSite->m_lfBreakPtID + 2.0);

				if (nRunMode == RUN_POSTDEV)
					nBrPtIndex = 0;				//index = -2+2=0
				else if(nRunMode == RUN_INIT)
					nBrPtIndex = 2;				//index = 0+2=2

				if (pBMPSite->m_nBreakPoints > 0 && nBrPtIndex >= 0)
				{
					if (pBMPSite->m_TradeOff != NULL)
					{
						BmpFlowInput[i] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult;	// in-acres/hr
						
						int nIndex = 0;
						for (j=0; j<NPOL; j++)
						{
							if (pBMPData->m_pPollutant[j].m_nSedfg == TSS)
							{
								// need to split the TSS into sand, silt, and clay
								double SED_FR[3];
								SED_FR[0] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfSand;
								SED_FR[1] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfSilt;
								SED_FR[2] = pBMPSite->m_TradeOff[nBrPtIndex].m_lfClay;
								for (k=0; k<3; k++)
								{
									BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum+1+j]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult*SED_FR[k];	// lbs/hr
									nIndex++;
								}
							}
							else
							{
								BmpWqInput[i*NWQ+nIndex] += pBMPSite->m_TradeOff[nBrPtIndex].m_pDataBrPt[t*pBMPSite->m_TradeOff[nBrPtIndex].m_nQualNum+1+j]*pBMPSite->m_TradeOff[nBrPtIndex].m_lfMult;	// lbs/hr
								nIndex++;
							}
						}
					}
				}

				i++;
			}
		}
		
		// start loop through each BMP site
		for (i=0; i<NBMP; i++)
		{			          
			// get model parameters
			pos = pBMPData->routeList.FindIndex(i);
			pBMPSite = (CBMPSite*) pBMPData->routeList.GetAt(pos);

			pos1 = pBMPSite->m_usbmpsiteList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pUS = (US_BMPSITE*) pBMPSite->m_usbmpsiteList.GetNext(pos1);
				pBMPSiteUp = pUS->m_pUSBMPSite;

				int i0 = ::FindObIndexFromList(pBMPData->routeList, pBMPSiteUp);
				
				// getting upstream flow and wq		
				if (pUS->m_nOutletType == TOTAL)
				{
					BmpFlowInput[i] += bmpoflow[i0]*3600/3630;//in-acre/hr
					
					for (j=0; j<NWQ; j++)
						BmpWqInput[i*NWQ+j] += bmpmassout[i0*NWQ+j];//lb/hr
				}
				else if (pUS->m_nOutletType == WEIR_)
				{
					BmpFlowInput[i] += bmpoflow_w[i0]*3600/3630;//in-acre/hr
					
					for (j=0; j<NWQ; j++)
						BmpWqInput[i*NWQ+j] += bmpmassout_w[i0*NWQ+j];//lb/hr
				}
				else if (pUS->m_nOutletType == ORIFICE_CHANNEL)
				{
					BmpFlowInput[i] += bmpoflow_o[i0]*3600/3630;//in-acre/hr
					//untreated bypass flow
					BmpFlowInput[i] += bmpoflow_ut[i0]*3600/3630;//in-acre/hr

					for (j=0; j<NWQ; j++)
					{
						BmpWqInput[i*NWQ+j] += bmpmassout_o[i0*NWQ+j];//lb/hr
						//untreated bypass pollutant
						BmpWqInput[i*NWQ+j] += bmpmassout_ut[i0*NWQ+j];// lb/hr
					}
				}
				else if (pUS->m_nOutletType == UNDERDRAIN)
				{
					BmpFlowInput[i] += bmpoflow_ud[i0]*3600/3630;//in-acre/hr
					
					for (j=0; j<NWQ; j++)
						BmpWqInput[i*NWQ+j] += bmpmassout_ud[i0*NWQ+j];//lb/hr
				}
			}

			// initialize Holtan equation parameters
			double soildepth      = pBMPSite->m_lfSoilDepth;
			double soilporosity   = pBMPSite->m_lfPorosity;
			double FC			  = pBMPSite->m_lfFCapacity;
			double WP			  = pBMPSite->m_lfWPoint;
			double vegparma       = pBMPSite->m_holtanParam.m_lfVegA;
			double finalf         = pBMPSite->m_holtanParam.m_lfFInfilt;
			double udsoildepth    = pBMPSite->m_lfUndDepth;
			double udsoilporosity = pBMPSite->m_lfUndVoid;		
			double udfinalf       = pBMPSite->m_lfUndInfilt;		
			bool   underdrain_on  = pBMPSite->m_bUndSwitch;		
			double *holtm         = pBMPSite->m_holtanParam.m_lfGrowth;
			int    nInfiltMethod  = pBMPSite->m_nInfiltMethod;
			int    nGAindex       = pBMPSite->m_nGAInfil_Index;

			// if underdrain option is not checked, then zero all associated parameters
			if(!underdrain_on)
			{
				udsoildepth    = 0.0;
				udsoilporosity = 0.0;
				udfinalf       = 0.0;
			}

			//calculate the max available storage in soil column (in)
			double nsamax = soildepth*12.0*(soilporosity - WP); 

			if (nsamax < SMALLNUM)
				nsamax = 0.0;

			//calculate the max available space in under-drain column (in)
			double nstoragemax = udsoildepth*12.0*udsoilporosity;

			if (nstoragemax < SMALLNUM)
				nstoragemax = 0.0;

			int	   timestep    = pBMPData->nBMPTimeStep;	// minutes  
			double ovolume     = bmpvol_p[i];//ft3
			double osa         = osa_p[i];
			double ostorage    = ostorage_p[i];
			double oinflow     = BmpFlowInput[i] * 3630;	// acre-in/hr to ft3/hr
			double oinflow2	   = oinflow;//ft3/hr
			double BMParea     = 0.0;	// ft^2	
			double BMParea_max = 0.0;	// ft^2	
			double ostage      = 0.0;
			double weir        = 0.0;
			double orifice     = 0.0;
			double channel     = 0.0;
			double qout1       = 0.0;
			double qout2       = 0.0;
			double udout       = 0.0;
			double utout       = 0.0;
			double bmpout      = 0.0;
			double infilt      = 0.0;
			double perc        = 0.0;	// place holder for output
			double AET         = 0.0;	// place holder for output
			double seepage     = 0.0;	// place holder for output
			double crrat       = 1.5;	// need to be user input
			double vol1		   = 0.0;		
			double vol2		   = 0.0;		

			// calculate mon, nxtmon, day, ndays first 
			COleDateTimeSpan tspan(0, t, 0, 0);
			COleDateTime tCurrent = pBMPData->startDate + tspan;
			int nYear  = tCurrent.GetYear();
			int mon    = tCurrent.GetMonth() - 1;
			int day    = tCurrent.GetDay() - 1;
			int hour   = tCurrent.GetHour();
			int nxtmon = (mon + 1) % 12;

			int ndays  = 31; // for Jan, Mar, May, Jul, Aug, Oct, Dec
			if (mon == 3 || mon == 5 || mon == 8 || mon == 10) // for Apr, Jun, Sep, Nov
				ndays = 30;
			if(mon == 1) // for Feb
			{
				if(nYear%400 == 0)
					ndays = 29;
				else if (nYear%4 == 0 && nYear%100 != 0)
					ndays = 29;
				else
					ndays = 28;
			}

			double holtpar = 0.0;
			if (ndays > 0)
				holtpar = (holtm[nxtmon] - holtm[mon]) * day / ndays + holtm[mon];

			// get ET constant rate (in/day)
			int nETflag = pBMPData->nETflag;
			double ETrate = pBMPData->lfmonET[mon];	// (in/day)
			COleDateTimeSpan tsSpan = COleDateTimeSpan(0,t,0,0);
			long nTSIndex = (long)tsSpan.GetTotalDays();
			
			if (nETflag == 1)
			{
				//get ET rate from the climate file (in/day)
				double evap = pBMPData->m_pDataClimate[nTSIndex*pBMPData->m_nNum+2];
				ETrate *= evap;	
			}
			else if (nETflag == 2)
			{
				//calculate PET rate from the temperature data (in/day)
				double cts  = ETrate;			
				double lat  = pBMPData->lfLatitude;
				double Tmax = pBMPData->m_pDataClimate[nTSIndex*pBMPData->m_nNum];
				double Tmin = pBMPData->m_pDataClimate[nTSIndex*pBMPData->m_nNum+1];
				double Tavf = (Tmax + Tmin) / 2.0;
				double tavc = (Tavf - 32.0) * 5.0 / 9.0;
				double lfday = tCurrent.GetDayOfYear();
				
				ETrate = pet_Hamon(lat, cts, tavc, lfday);
			}

			//convert ETrate from in/day to in/hr
			ETrate /= 24.0;

			int counter = counter_p[i];
			double ndevice = pBMPSite->m_lfBMPUnit;
			double lfDDarea = pBMPSite->m_lfDDarea;		//acre
			double lfAccDArea = pBMPSite->m_lfAccDArea;	//acre

			if (pBMPSite->m_nBMPClass == CLASS_A && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
			{
				// initialize class A BMP related parameters
				BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;
				
				double basinlength    = pBMP->m_lfBasinLength;		// ft
				double basinwidth     = pBMP->m_lfBasinWidth;		// ft
				int	   npeople        = pBMP->m_nPeople;			
				int    ddays          = pBMP->m_nDays;				
				double orificeheight  = pBMP->m_lfOrificeHeight;
				double orificediam    = pBMP->m_lfOrificeDiameter;	// in
				double weirheight     = pBMP->m_lfWeirHeight;		// ft
				double weirwidth      = pBMP->m_lfWeirWidth;
				int    releasetype    = pBMP->m_nORelease;		
				int    weirtype       = pBMP->m_nWeirType;  
				double weirangle      = pBMP->m_lfWeirAngle;
				double cisternoutflow = pBMP->m_lfRelease[hour];			
				int    exittype       = pBMP->m_nExitType;
				double lfPI           = 3.14159;
				double orifice_area   = (lfPI/4.)*(orificediam*orificediam)/144.;  //ft2 
				double orificecoef    = 0.61;

				BMParea = basinlength * basinwidth;	// ft^2	
	
				// check if this BMP is cistern or rainbarrel, if so then 
				// basinlength = diameter (ft) and basinwidth = number of devices
				if (releasetype == 1 || releasetype == 2)
					BMParea   = 3.142857/4.0*pow(basinlength,2); // ft2 

				BMParea_max = BMParea;	

				switch (exittype)
				{
					case 1:
						orificecoef = 1.0;
						break;
					case 2:
						orificecoef = 0.61;
						break;
					case 3:
						orificecoef = 0.61;
						break;
					case 4:
						orificecoef = 0.5;
						break;
					default:
						orificecoef = 1.0;
						break;
				}

				if (BMParea > 0 && ndevice > 0)		
				{
					//check the design drainage area
					if (lfDDarea > 0 && lfAccDArea > (lfDDarea*ndevice))
					{
						utout = oinflow*(lfAccDArea-(lfDDarea*ndevice))/lfAccDArea;//ft3/hr
						oinflow2 = oinflow - utout;//ft3/hr
						utout /= 3600.0;//cfs 
					}
					else
					{
						utout = 0.0;//cfs
						oinflow2 = oinflow;//ft3/hr
					}

					ovolume /= ndevice;//ft3
					oinflow2 /= ndevice;//ft3/hr

					bmp_a(nInfiltMethod,nGAindex,underdrain_on,timestep,npeople,ddays,
						  releasetype,weirtype,counter,oinflow2,BMParea,orifice_area,
						  orificeheight,orificecoef,weirwidth,weirheight,weirangle,
						  cisternoutflow,soildepth,soilporosity,finalf,vegparma,
						  holtpar,udfinalf,udsoildepth,udsoilporosity,FC,WP,ETrate,
						  AET,perc,ovolume,ostage,infilt,orifice,weir,osa,ostorage,
						  udout,seepage);

					ovolume  *= ndevice;	//ft3
					oinflow2 *= ndevice;	//ft3
					weir     *= ndevice;	//cfs
					orifice  *= ndevice;	//cfs
					udout    *= ndevice;	//cfs
					AET      *= ndevice;	//cfs
					infilt   *= ndevice;	//cfs
					perc     *= ndevice;	//cfs
					seepage  *= ndevice;	//cfs
				}
				else // if area is 0 then outflow = inflow 
				{
					utout = oinflow/3600.0;//total inflow is untreated (cfs)
					ovolume = 0.0;	// no stored volume
					weir = 0;
					orifice = 0.0;
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
					ostage = 0.0;
				}
			}
			else if (pBMPSite->m_nBMPClass == CLASS_B && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
			{
				// initialize class B BMP related parameters
				BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;
				double BMPlength = pBMP->m_lfBasinLength;	// ft
				double BMPwidth  = pBMP->m_lfBasinWidth;	// ft
				double BMPdepth  = pBMP->m_lfMaximumDepth;	// ft
				double slope1    = pBMP->m_lfSideSlope1;
				double slope2    = pBMP->m_lfSideSlope2;
				double slope3    = pBMP->m_lfSideSlope3;
				double man_n     = pBMP->m_lfManning;

				BMParea = BMPlength * BMPwidth;	// ft^2	

				//calculate the top width (ft)
				double top_width = (BMPdepth/slope1 + BMPdepth/slope2 + BMPwidth); 

				//calculate the maximum surface area (ft2)
				BMParea_max = top_width * BMPlength;	 

				if (BMParea > 0 && BMPdepth > 0 && ndevice > 0)		
				{
					//check the design drainage area
					if (lfDDarea > 0 && lfAccDArea > (lfDDarea*ndevice))
					{
						utout = oinflow*(lfAccDArea-(lfDDarea*ndevice))/lfAccDArea;//ft3/hr
						oinflow2 = oinflow - utout;//ft3/hr
						utout /= 3600.0;//cfs 
					}
					else
					{
						utout = 0.0;//cfs
						oinflow2 = oinflow;//ft3/hr
					}

					ovolume /= ndevice;//ft3
					oinflow2 /= ndevice;//ft3/hr

					bmp_b(nInfiltMethod,nGAindex,underdrain_on,timestep,oinflow2,
						  BMPdepth,BMPwidth,BMPlength,slope1,slope2,slope3,man_n,
						  soildepth,soilporosity,finalf,vegparma,holtpar,udfinalf,
						  udsoildepth,udsoilporosity,FC,WP,ETrate,AET,perc,ovolume,
						  ostage,infilt,channel,weir,osa,ostorage,udout,seepage);

					ovolume  *= ndevice;//ft3
					oinflow2 *= ndevice;//ft3
					weir     *= ndevice;//cfs
					channel  *= ndevice;//cfs
					udout    *= ndevice;//cfs
					AET      *= ndevice;//cfs
					infilt   *= ndevice;//cfs
					perc     *= ndevice;//cfs
					seepage  *= ndevice;//cfs
					orifice = channel;
				}
				else // if area is 0 then outflow = inflow    
				{
					utout = oinflow/3600.0;//total inflow is untreated (cfs)
					ovolume = 0.0;	// no stored volume
					weir = 0;
					channel = 0.0;
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
					orifice = channel;
					ostage = 0.0;
				}
			}
			else if (pBMPSite->m_nBMPClass == CLASS_C && nRunMode != RUN_PREDEV) 
			{
				// initialize class C Conduit related parameters
				BMP_C* pBMP = (BMP_C*) pBMPSite->m_pSiteProp;

				int nIndex = pBMP->m_nIndex;
			    k = Link[nIndex].subIndex;

				double BMPvolume = Conduit[k].length * xsect_getAmax(&Link[nIndex].xsect);	// ft3			
				
				if (BMPvolume > 0 && Link[nIndex].xsect.type != DUMMY)
				{
					int ii, nivl;
					float  delts = timestep * 60.0;			// sec/ivl 
					float  deltd = delts / (24.0*3600.0);	// day/ivl 
					float  qoutflow = 0.0;
					float  overflow = 0.0;	
					double qout  = 0.0;
					double qover = 0.0;

					nivl = 60 / timestep;	// number of intervals per hour (time step is in minutes)	 

					for (ii=0; ii<nivl; ii++)	
					{
						float qinflow = oinflow / 3600.0;	// from ft^3/hr to cfs	
						qoutflow = 0.0;

						// --- replace old hydraulic state values with current ones
						link_setOldHydState(nIndex);

						// routing using KW model
						kinwave_execute(nIndex, &qinflow, &qoutflow,delts);

						Link[nIndex].newFlow = qoutflow;

						// calculate overflow (cfs)
						overflow = max(0.0, (oinflow/3600.0)-qinflow);

						qout  += qoutflow;	// cfs
						qover += overflow;	// cfs

						//updates state of link after current time step 
						// --- find avg. depth from entry/exit conditions
						float a = 0.5 * (Conduit[k].a1 + Conduit[k].a2);   // avg. area
						Link[nIndex].newVolume = max(0.0, a * Conduit[k].length * Conduit[k].barrels);
						float y1 = xsect_getYofA(&Link[nIndex].xsect, Conduit[k].a1);
						float y2 = xsect_getYofA(&Link[nIndex].xsect, Conduit[k].a2);
						Link[nIndex].newDepth = 0.5 * (y1 + y2);

						// --- get velocity (ft/sec)
						double avvele = link_getVelocity(nIndex, Link[nIndex].newFlow, Link[nIndex].newDepth);//ft/s
						double avdepe = Link[nIndex].newDepth;	// ft
						double avdepm = avdepe * FOOT2METER;	// m
						double slope = Conduit[k].slope;
						double hrade = getHydRad(&Link[nIndex].xsect, avdepe);	//ft
						double hradm = hrade * FOOT2METER;	// m
						double tw = 20.0;//degreeC

						double ros  = Conduit[k].q2Old * Conduit[k].barrels;	//cfs
						double ro   = Conduit[k].q2 * Conduit[k].barrels;		//cfs
						double rosm = ros * CFS2CMS;//m3/s
						double rom  = ro * CFS2CMS;	//m3/s
						double vols = Link[nIndex].oldVolume;	//ft3
						double vol  = Link[nIndex].newVolume;	//ft3
						double volsm = vols * CF2CM;//m3
						double volm  = vol * CF2CM;	//m3
						
						double js   = 0.0;

						if (fabs(ros) > fThreshold)
						{
							double rat = vols / (ros * delts);
							if (rat < crrat)
								js = rat / crrat;
							else
								js = 1.0;
						}

						double cojs    = 1.0 - js;
						double srovol  = js * ros * delts;	//ft3
						double erovol  = cojs * ro * delts;	//ft3
						double srovolm = srovol * CF2CM;	//m3
						double erovolm = erovol * CF2CM;	//m3

						// calculate the water quality
						for (j=0; j<NWQ; j++)
						{
							//replaces old water quality state values with current ones
							Link[nIndex].oldQual[j] = Link[nIndex].newQual[j];
							//Link[nIndex].newQual[j] = 0.0;// lb/ft3
							double conc   = Link[nIndex].oldQual[j];// lb/ft3
							double concs  = Link[nIndex].oldQual[j];
							double massin = BmpWqInput[i*NWQ+j]/nivl;// lb/ivl
							double massin2 = massin;//lb/timestep
							double rsed1  = rbsed[i*NWQ+j];			// lb
							double rsed1tot = rbsedtot[i];			// lb
							double romat  = 0.0;					// lb

							//check if any bypass
							if (overflow > 0 && oinflow > 0)
							{
								massin2 = massin * (oinflow - overflow*3600.0)/oinflow;// lb/timestep
								romat_ut2[i*NWQ+j] += max(0.0, massin - massin2);//lb
							}

							if (pBMPData->nSedflag[j] == SAND)
							{
								//simulate sand
								int    sandfg = 3;	// user-specified power function method
								double ksand  = pBMPSite->m_sand.m_lfKSAND;
								double expsnd = pBMPSite->m_sand.m_lfEXPSND;
								double db50e  = pBMPSite->m_sand.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double w = pBMPSite->m_sand.m_lfW*0.0254*delts;	// fall velocity (m/ivl)
								double wsande = w * 3.28 / delts; // from m/ivl to ft/sec
								double depscr = 0., rsed = 0.;
								
								//dummy variables (not required for sandfg = 3)
								double twide = 0.,fsl = 0.;

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2 *= POUND2GRAM;	//gram
									rsed1   *= POUND2GRAM;	//gram
									conc    *= LBpCFT2MGpL;	//mg/l

									sandld(massin2,volsm,srovolm,volm,erovolm,ksand,avvele, 
										   expsnd,rom,sandfg,db50e,hrade,slope,tw,wsande,
										   twide,db50m,fsl,avdepe,&conc,&rsed,&rsed1,&depscr,
										   &romat);

									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	(g)
											rsed1  += rsed;	// mg/l*m3	(g)
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	//lb/ft3
									romat /= POUND2GRAM;	//lb
									rsed1 /= POUND2GRAM;	//lb
								}
								else
								{
									// conduit has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving conduit 
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								Link[nIndex].newQual[j] = conc;	// lb/ft3
								romat_o2[i*NWQ+j] += romat;		// lb
								rbsed[i*NWQ+j] = rsed1;			// lb
							}
							else if (pBMPData->nSedflag[j] == SILT)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_silt.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_silt.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_silt.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_silt.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
									if (avdepm > 0.0)
										//use formula appropriate to a river or stream
										tau = gamma * hradm * slope;	// (kg/m^2)

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// conduit has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving conduit 
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								Link[nIndex].newQual[j] = conc;	// lb/ft3
								romat_o2[i*NWQ+j] += romat;		// lb
								rbsed[i*NWQ+j] = rsed1;			// lb
								rbsedtot[i] += rsed1;
							}
							else if (pBMPData->nSedflag[j] == CLAY)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_clay.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_clay.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_clay.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_clay.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
									if (avdepm > 0.0)
										//use formula appropriate to a river or stream
										tau = gamma * hradm * slope;	// (kg/m^2)

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// conduit has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving conduit 
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								Link[nIndex].newQual[j] = conc;	// lb/ft3
								romat_o2[i*NWQ+j] += romat;		// lb
								rbsed[i*NWQ+j] = rsed1;			// lb
								rbsedtot[i] += rsed1;
							}
							else	// not sediment
							{
								double decay  = pBMPSite->m_pDecay[j]/3600.0;	// per sec

								if (fabs(volm) > fThreshold)
								{
									//findLinkQual(nIndex, delts);
									findLinkQual2(nIndex,delts,massin2,decay,conc);

									// calculate total amount of material leaving conduit 
									// during the interval  
									if (ro > 0)
									{
										romat = srovol * concs + erovol * conc; //	lb/timestep
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}
								else
								{
									// conduit has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving conduit 
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update conc
								Link[nIndex].newQual[j] = conc;	// lb/ft3
								romat_o2[i*NWQ+j] += romat;
							}
						}
					}

					if (nivl > 0)
					{
						qoutflow = qout / nivl;
						overflow = qover / nivl;
					}

					utout = overflow;// cfs
					ovolume = Link[nIndex].newVolume;	// ft3
					ostage = Link[nIndex].newDepth;		// ft
					channel = qoutflow;					// cfs
					weir = 0.0;//cfs
					orifice = channel;//cfs
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
				}
				else // if area is 0 then outflow = inflow    
				{
					utout = oinflow/3600.0;//total inflow is untreated (cfs)
					ovolume = 0.0;	// no stored volume
					weir = 0;
					orifice = 0.0;
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
					ostage = 0.0;
				}
			}
			else if (pBMPSite->m_nBMPClass == CLASS_D && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
			{
				// initialize class D related parameters
				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				BMParea = pBMP->m_lfLength * pBMP->m_lfWidth;

				if (BMParea > 0)
				{
					//get the VFSMOD results (implemented later on)
					utout = oinflow/3600.0;//total inflow is untreated (cfs)
					ovolume = 0.0;	// no stored volume
					weir = 0;
					orifice = 0.0;
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
					ostage = 0.0;
				}
				else // if area is 0 then outflow = inflow    
				{
					utout = oinflow/3600.0;//total inflow is untreated (cfs)
					ovolume = 0.0;	// no stored volume
					weir = 0;
					orifice = 0.0;
					udout = 0;
					AET = 0;
					infilt = 0;
					perc = 0;
					seepage = 0;
					ostage = 0.0;
				}
			}
			else   //Clase X		// Dummy BMP
			{
				utout = oinflow/3600.0;//total inflow is untreated (cfs)
				ovolume = 0.0;	// no stored volume
				weir = 0;
				orifice = 0.0;
				udout = 0;
				AET = 0;
				infilt = 0;
				perc = 0;
				seepage = 0;
				ostage = 0.0;
			}

			qout1= weir_p[i] + orifice_p[i] + undrain_p[i] + seepage_p[i]; //cfs
			qout2= weir + orifice + udout + seepage;//cfs
			bmpout = weir + orifice + udout + utout;//cfs
			
			vol1 = vol_p[i];//ft3		
			vol2 = max(0.0, vol1 + oinflow2 - (qout2 + AET) * 3600.00);	//ft3		

			//double value1 = bmpoflow[i]*3600.00; //ft3/s to ft3/hr
			double value1 = bmpout*3600.0; //ft3/s to ft3/hr
			double value0 = pBMPSite->qFlow.front();
			lfSumFlow[i] -= value0;
			lfSumFlow[i] += value1;
			//round to 3rd decimal place
			//lfSumFlow[i] = floor(lfSumFlow[i]*1.0E3+0.5)/1.0E3;
			double lfFlowVolume = floor(lfSumFlow[i]+0.5);
			pBMPSite->qFlow.pop();
			pBMPSite->qFlow.push(value1);
			
			// evaluation factor calculation
			if (pBMPSite->m_factorList.GetCount() > 0)
			{
				bmpTotalFlow[i] += (bmpout*3600.00);//cumulative hourly flow for BMP (ft3/hr)
				if (bmpPkDisFlow[i] < bmpout)
					bmpPkDisFlow[i] = bmpout;// keep track of peak discharge for the entire simulation period (cfs)
				if (bmpout > pBMPSite->m_lfThreshFlow)
				{
					if (nExceedFlag[i] == 0)
						nExceedFlow[i]++;		// count the exceeding flow over threshold value
					nExceedFlag[i] = 1;
				}
				else
				{
					nExceedFlag[i] = 0;
				}
			}
				
			// Pollutants 
			double delts  = 3600;//sec/ivl (ivl=hr)
			double deltd = delts / (24.0*3600.0);// day/ivl 
						
			int	nRoutingMethod = pBMPSite->m_nPolRotMethod;	
			int	nRemovalMethod = pBMPSite->m_nPolRemMethod;	

			for (int jj=0; jj<pBMPData->m_nSedQualFlag+1; jj++)
			{
				//pollutant index
				int nPollIndex = 0;

				for (j=0; j<NWQ; j++)
				{
					double massin = BmpWqInput[i*NWQ+j];//lb/hr
					double massin2 = massin;//lb/hr
					double conc      = 0.0;
					double romat     = 0.0;
					double romat_w   = 0.0;//weir
					double romat_o   = 0.0;//orifice
					double romat_i   = 0.0;//infiltration
					double romat_ud  = 0.0;//underdrain
					double romat_ut  = 0.0;//bypass
					double decay     = 0.0;
					double lfK       = 0.0;
					double lfCstar   = 0.0;
					double BMPAREA   = 0.0;

					//check if any sediment-associated qual
					if (pBMPData->m_nSedQualFlag == 1 && jj == 0)
					{
						//need to simulate sediment before other quals
						if (pBMPData->nSedflag[j] == 0)	//not sediment
						{
							nPollIndex++;
							continue;
						}
					}
					else if (pBMPData->m_nSedQualFlag == 1 && jj == 1)
					{
						//sediment is already simulated
						if (pBMPData->nSedflag[j] != 0) //sediment
						{
							if (NWQ > NPOL)	// there is split of total sediment
								if (pBMPData->nSedflag[j] == SILT || pBMPData->nSedflag[j] == CLAY)
									nPollIndex--;
							nPollIndex++;
							continue;
						}

						//sediment-associated pollutant
						if (pBMPData->m_pPollutant[nPollIndex].m_nSedQual == 1)	
						{
							//find the index for sediment classes
							int nSandIndex = -1;
							int nSiltIndex = -1;
							int nClayIndex = -1;
							for (int kk=0; kk<NWQ; kk++)
							{
								if (pBMPData->nSedflag[kk] == SAND)
									nSandIndex = kk;
								else if (pBMPData->nSedflag[kk] == SILT)
									nSiltIndex = kk;
								else if (pBMPData->nSedflag[kk] == CLAY)
									nClayIndex = kk;
							}

							//find the BMP performance for sediment
							double lfEfficiency_Sand = 0;
							double lfEfficiency_Silt = 0;
							double lfEfficiency_Clay = 0;

							double lfBmpWqInput_Sand = 0;
							double lfBmpWqInput_Silt = 0;
							double lfBmpWqInput_Clay = 0;

							double lfbmpmassout_Sand = 0;
							double lfbmpmassout_Silt = 0;
							double lfbmpmassout_Clay = 0;

							double lfbmpmassout_ut_Sand = 0;
							double lfbmpmassout_ut_Silt = 0;
							double lfbmpmassout_ut_Clay = 0;

							double lfbmpmassout_w_Sand = 0;
							double lfbmpmassout_w_Silt = 0;
							double lfbmpmassout_w_Clay = 0;

							double lfbmpmassout_o_Sand = 0;
							double lfbmpmassout_o_Silt = 0;
							double lfbmpmassout_o_Clay = 0;

							double lfbmpmassout_ud_Sand = 0;
							double lfbmpmassout_ud_Silt = 0;
							double lfbmpmassout_ud_Clay = 0;

							double lfsand_qfr = pBMPData->m_pPollutant[nPollIndex].m_lfsand_qfr;
							double lfsilt_qfr = pBMPData->m_pPollutant[nPollIndex].m_lfsilt_qfr;
							double lfclay_qfr = pBMPData->m_pPollutant[nPollIndex].m_lfclay_qfr;
							
							if (nSandIndex != -1)
							{
								lfBmpWqInput_Sand = BmpWqInput[i*NWQ+nSandIndex];
								lfbmpmassout_Sand = bmpmassout[i*NWQ+nSandIndex];
								lfbmpmassout_ut_Sand = bmpmassout_ut[i*NWQ+nSandIndex];
								lfbmpmassout_w_Sand = bmpmassout_w[i*NWQ+nSandIndex];
								lfbmpmassout_o_Sand = bmpmassout_o[i*NWQ+nSandIndex];
								lfbmpmassout_ud_Sand = bmpmassout_ud[i*NWQ+nSandIndex];
							}
							if (nSiltIndex != -1)
							{
								lfBmpWqInput_Silt = BmpWqInput[i*NWQ+nSiltIndex];
								lfbmpmassout_Silt = bmpmassout[i*NWQ+nSiltIndex];
								lfbmpmassout_ut_Silt = bmpmassout_ut[i*NWQ+nSiltIndex];
								lfbmpmassout_w_Silt = bmpmassout_w[i*NWQ+nSiltIndex];
								lfbmpmassout_o_Silt = bmpmassout_o[i*NWQ+nSiltIndex];
								lfbmpmassout_ud_Silt = bmpmassout_ud[i*NWQ+nSiltIndex];
							}
							if (nClayIndex != -1)
							{
								lfBmpWqInput_Clay = BmpWqInput[i*NWQ+nClayIndex];
								lfbmpmassout_Clay = bmpmassout[i*NWQ+nClayIndex];
								lfbmpmassout_ut_Clay = bmpmassout_ut[i*NWQ+nClayIndex];
								lfbmpmassout_w_Clay = bmpmassout_w[i*NWQ+nClayIndex];
								lfbmpmassout_o_Clay = bmpmassout_o[i*NWQ+nClayIndex];
								lfbmpmassout_ud_Clay = bmpmassout_ud[i*NWQ+nClayIndex];
							}

							if (lfBmpWqInput_Sand > 0)
							{
								lfEfficiency_Sand = 1 - (lfBmpWqInput_Sand-lfbmpmassout_Sand)/lfBmpWqInput_Sand;
							}
							if (lfBmpWqInput_Silt > 0)
							{
								lfEfficiency_Silt = 1 - (lfBmpWqInput_Silt-lfbmpmassout_Silt)/lfBmpWqInput_Silt;
							}
							if (lfBmpWqInput_Clay > 0)
							{
								lfEfficiency_Clay = 1 - (lfBmpWqInput_Clay-lfbmpmassout_Clay)/lfBmpWqInput_Clay;
							}

							//calculate the sed-associated qual load
							double qconc_sand = bmpqconc_sand[i*NWQ+j];// lb/lb
							double qconc_silt = bmpqconc_silt[i*NWQ+j];// lb/lb
							double qconc_clay = bmpqconc_clay[i*NWQ+j];// lb/lb

							//check if there is no inflow but there is outflow
							if (massin == 0 && lfbmpmassout_Sand + lfbmpmassout_Silt + lfbmpmassout_Clay > 0)
							{
								romat = qconc_sand * lfbmpmassout_Sand + qconc_silt * lfbmpmassout_Silt + qconc_clay * lfbmpmassout_Clay;
							}
							else
							{
								romat = massin * (lfsand_qfr*lfEfficiency_Sand 
									+ lfsilt_qfr*lfEfficiency_Silt 
									+ lfclay_qfr*lfEfficiency_Clay);
							}

							double lfsandout = 0.0;
							double lfsiltout = 0.0;
							double lfclayout = 0.0;

							if (lfbmpmassout_Sand > 0)
								lfsandout = lfbmpmassout_ut_Sand/lfbmpmassout_Sand;
							if (lfbmpmassout_Silt > 0)
								lfsiltout = lfbmpmassout_ut_Silt/lfbmpmassout_Silt;
							if (lfbmpmassout_Clay > 0)
								lfclayout = lfbmpmassout_ut_Clay/lfbmpmassout_Clay;

							romat_ut = romat * (lfsand_qfr*lfsandout 
								+ lfsilt_qfr*lfsiltout + lfclay_qfr*lfclayout);

							lfsandout = 0.0;
							lfsiltout = 0.0;
							lfclayout = 0.0;

							if (lfbmpmassout_Sand > 0)
								lfsandout = lfbmpmassout_w_Sand/lfbmpmassout_Sand;
							if (lfbmpmassout_Silt > 0)
								lfsiltout = lfbmpmassout_w_Silt/lfbmpmassout_Silt;
							if (lfbmpmassout_Clay > 0)
								lfclayout = lfbmpmassout_w_Clay/lfbmpmassout_Clay;

							romat_w = romat * (lfsand_qfr*lfsandout 
								+ lfsilt_qfr*lfsiltout + lfclay_qfr*lfclayout);

							lfsandout = 0.0;
							lfsiltout = 0.0;
							lfclayout = 0.0;

							if (lfbmpmassout_Sand > 0)
								lfsandout = lfbmpmassout_o_Sand/lfbmpmassout_Sand;
							if (lfbmpmassout_Silt > 0)
								lfsiltout = lfbmpmassout_o_Silt/lfbmpmassout_Silt;
							if (lfbmpmassout_Clay > 0)
								lfclayout = lfbmpmassout_o_Clay/lfbmpmassout_Clay;

							romat_o = romat * (lfsand_qfr*lfsandout 
								+ lfsilt_qfr*lfsiltout + lfclay_qfr*lfclayout);

							lfsandout = 0.0;
							lfsiltout = 0.0;
							lfclayout = 0.0;

							if (lfbmpmassout_Sand > 0)
								lfsandout = lfbmpmassout_ud_Sand/lfbmpmassout_Sand;
							if (lfbmpmassout_Silt > 0)
								lfsiltout = lfbmpmassout_ud_Silt/lfbmpmassout_Silt;
							if (lfbmpmassout_Clay > 0)
								lfclayout = lfbmpmassout_ud_Clay/lfbmpmassout_Clay;

							romat_ud = romat * (lfsand_qfr*lfsandout 
								+ lfsilt_qfr*lfsiltout + lfclay_qfr*lfclayout);

							//calculate the sed-associated qual conc (lb of qual / lb of sed)
//							conc = bmpc[i*NWQ+j];// lb/lb
//							if (lfbmpmassout_Sand + lfbmpmassout_Silt + lfbmpmassout_Clay > 0)
//								conc = romat / (lfbmpmassout_Sand + lfbmpmassout_Silt + lfbmpmassout_Clay);

							//save values
							if (lfbmpmassout_Sand > 0)
								bmpqconc_sand[i*NWQ+j] = romat * lfsand_qfr / lfbmpmassout_Sand;//lb/lb
							if (lfbmpmassout_Silt > 0)
								bmpqconc_silt[i*NWQ+j] = romat * lfsilt_qfr / lfbmpmassout_Silt;//lb/lb
							if (lfbmpmassout_Clay > 0)
								bmpqconc_clay[i*NWQ+j] = romat * lfclay_qfr / lfbmpmassout_Clay;//lb/lb

							goto MM;
						}
					}

					if (NWQ > NPOL)	// there is split of total sediment
						if (pBMPData->nSedflag[j] == SILT || pBMPData->nSedflag[j] == CLAY)
							nPollIndex--;

					//output variable
					BmpWqInput_s[i*NPOL+nPollIndex] += massin;//lb/hr

					//check for untreated over flow
					if (utout > 0 && oinflow > 0)
					{
						massin2 = massin * (oinflow - utout*3600.0)/oinflow;// lb/hr
						romat_ut = max(0.0, massin-massin2);//lb/hr
					}

					if (pBMPSite->m_nBMPClass == CLASS_A && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
					{
						BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;
						BMPAREA = pBMP->m_lfBasinLength * pBMP->m_lfBasinWidth;

						if (BMPAREA > SMALLNUM && ndevice > 0) 
						{
							double avdepe = ovolume/BMPAREA;// ft
							double avdepm = avdepe * FOOT2METER;// m
							double ros  = weir_p[i] + orifice_p[i] + infilt_p[i];//cfs
							double ro   = weir + orifice + infilt;//cfs
							double rosm = ros * CFS2CMS;//m3/s
							double rom  = ro * CFS2CMS;	//m3/s
							double vols = bmpvol_p[i];//ft3
							double vol  = ovolume;//ft3
							double volsm = vols * CF2CM;//m3
							double volm  = vol * CF2CM;	//m3
							double xarea = vol / pBMP->m_lfBasinLength;//ft2
							double avvele = ro/xarea;//ft/s
							double avvelm = avvele * FpS2MpS;//m/s
							
							double slope = 0.001;
							double hrade = xarea/(2*avdepe+pBMP->m_lfBasinWidth);//ft
							double hradm = hrade * FOOT2METER;	// m
							double tw = 20.0;//degreeC

							
							double js   = 0.0;

							if (fabs(ros) > fThreshold)
							{
								double rat = vols / (ros * delts);
								if (rat < crrat)
									js = rat / crrat;
								else
									js = 1.0;
							}

							double cojs    = 1.0 - js;
							double srovol  = js * ros * delts;	//ft3
							double erovol  = cojs * ro * delts;	//ft3
							double srovolm = srovol * CF2CM;	//m3
							double erovolm = erovol * CF2CM;	//m3

							conc = bmpc[i*NWQ+j];				// lb/ft3
							double concs  = bmpc[i*NWQ+j];		// lb/ft3
							double rsed1  = rbsed[i*NWQ+j];		// lb
							double rsed1tot = rbsedtot[i];		// lb

							//simulate sediment
							if (pBMPData->nSedflag[j] == SAND)
							{
								//simulate sand
								int    sandfg = 3;	// user-specified power function method
								double ksand  = pBMPSite->m_sand.m_lfKSAND;
								double expsnd = pBMPSite->m_sand.m_lfEXPSND;
								double db50e  = pBMPSite->m_sand.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double w = pBMPSite->m_sand.m_lfW*0.0254*delts;	// fall velocity (m/ivl)
								double wsande = w * 3.28 / delts; // from m/ivl to ft/sec
								double depscr = 0., rsed = 0.;
									
								//dummy variables (not required for sandfg = 3)
								double twide = 0.,fsl = 0.;

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2 *= POUND2GRAM;	//gram
									rsed1   *= POUND2GRAM;	//gram
									conc    *= LBpCFT2MGpL;	//mg/l

									sandld(massin2,volsm,srovolm,volm,erovolm,ksand,avvele, 
										   expsnd,rom,sandfg,db50e,hrade,slope,tw,wsande,
										   twide,db50m,fsl,avdepe,&conc,&rsed,&rsed1,&depscr,
										   &romat);

									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	(g)
											rsed1  += rsed;	// mg/l*m3	(g)
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	//lb/ft3
									romat /= POUND2GRAM;	//lb
									rsed1 /= POUND2GRAM;	//lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3
								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb

								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else if (pBMPData->nSedflag[j] == SILT)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_silt.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_silt.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_silt.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_silt.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)

								double db50e  = pBMPSite->m_silt.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double akappa = 0.4;	// Karman constant 
								double grav = 9.81;		// m/sec^2
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
								  
									if (avdepm > 0.0)
									{
										//use formula appropriate to a lake- from "hydraulics
										//of sediment transport", by w.h. graf- eq.8.49
										double ustar = avvelm/(17.66+(log10(avdepm/(96.5*db50m)))*2.3/akappa);
										
										tau = gamma*pow(ustar,2)/grav;
									}

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3

								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb
								rbsedtot[i] += rsed1;

								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else if (pBMPData->nSedflag[j] == CLAY)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_clay.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_clay.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_clay.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_clay.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)

								double db50e  = pBMPSite->m_clay.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double akappa = 0.4;	// Karman constant 
								double grav = 9.81;		// m/sec^2
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
								  
									if (avdepm > 0.0)
									{
										//use formula appropriate to a lake- from "hydraulics
										//of sediment transport", by w.h. graf- eq.8.49
										double ustar = avvelm/(17.66+(log10(avdepm/(96.5*db50m)))*2.3/akappa);
										
										tau = gamma*pow(ustar,2)/grav;
									}

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3

								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb
								rbsedtot[i] += rsed1;

								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else	// not sediment
							{
								if (nRoutingMethod <= 1)
								{
									conc = bmpc[i*NWQ+j];			//lb/ft3		
									advect(massin2,vol1,qout1,vol2,qout2,delts,crrat,conc,romat);

									if (nRemovalMethod == 0)
									{
										//1st order decay
										decay = pBMPSite->m_pDecay[j];//per hour
										conc = conc * exp(-decay);
									}
									else if (nRemovalMethod == 1)
									{
										//kadlec and knight method
										//Cout = Cstar + (Cin - Cstar) * exp (-k/q)
										lfK = pBMPSite->m_pK[j];		//ft/hr
										lfCstar = pBMPSite->m_pCstar[j];//lb/ft3
										
										//maintain minimum conc of lfCstar
										conc = max(conc, lfCstar);

										double lfq = bmpout*delts/BMPAREA;
										if (lfq > 0 && conc > lfCstar)
											conc = lfCstar + (conc - lfCstar) * exp(-lfK/lfq);
									}
								}
								else if (nRoutingMethod > 1)
								{
									//CSTRs in series routing
									int nCSTRs = nRoutingMethod;
									float v = vol1/nCSTRs;	//ft3
									float lfmassin = massin2/delts;	//lb/sec
									float lfinflow = oinflow2/delts;//cfs
									float lfoutflow = qout2;		//cfs
									float tStep = delts;			//sec/hr

									for(k=0; k<nCSTRs; k++)
									{
										//calculate inputs for each CSTR
										float c = pBMPSite->m_pConc[j*nCSTRs+k];	//lb/ft3
										float wIn = lfmassin;						//lb/sec
										float qNet = lfinflow+(lfoutflow-lfinflow)*k/nCSTRs;		//cfs
										float qNetout = lfinflow+(lfoutflow-lfinflow)*(k+1)/nCSTRs;	//cfs

										//get the new conc (lb/ft3)
										conc = getCstrQual(c,v,wIn,qNet,tStep);

										//calculate massout (lb/hr)
										romat = conc * qNetout * delts;

										if (nRemovalMethod == 0)
										{
											//1st order decay
											decay = pBMPSite->m_pDecay[j];//per hour
											conc = conc * exp(-decay);	
										}
										else if (nRemovalMethod == 1)
										{
											//kadlec and knight method
											//Cout = Cstar + (Cin - Cstar) * exp (-k/q)
											lfK = pBMPSite->m_pK[j];		//ft/hr
											lfCstar = pBMPSite->m_pCstar[j];//lb/ft3

											//maintain minimum conc of lfCstar
											conc = max(conc, lfCstar);

											double lfq = lfoutflow*delts/BMPAREA;
											if (lfq > 0 && conc > lfCstar)
												conc = lfCstar + (conc - lfCstar) * exp(-lfK/lfq);
										}

										//update conc
										pBMPSite->m_pConc[j*nCSTRs+k] = conc;

										//calculate massin to the next reactor(lb/sec)
										lfmassin = romat / delts;
									}
								}

								if (qout2 > 0)
								{
									romat_w = weir/qout2*romat;
									romat_o = orifice/qout2*romat;
									romat_ud = udout/qout2*romat;
									romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
								}
							}
						}
						else
						{
							//no volume stored
							conc = 0.0;			
							romat_w = 0.0;
							romat_o = 0.0;
							romat_ud = 0.0;
							romat_ut = massin;	//lb/hr
							romat = 0.0;
						}
					}
					else if (pBMPSite->m_nBMPClass == CLASS_B && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
					{
						BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;
						BMPAREA = pBMP->m_lfBasinLength * pBMP->m_lfBasinWidth;

						if (BMPAREA > SMALLNUM && ndevice > 0) 
						{
							double avdepe = ovolume/BMPAREA;// ft
							double avdepm = avdepe * FOOT2METER;	// m
							double ros  = weir_p[i] + orifice_p[i] + infilt_p[i];//cfs
							double ro   = weir + orifice + infilt;//cfs
							double rosm = ros * CFS2CMS;//m3/s
							double rom  = ro * CFS2CMS;	//m3/s
							double vols = bmpvol_p[i];//ft3
							double vol  = ovolume;//ft3
							double volsm = vols * CF2CM;//m3
							double volm  = vol * CF2CM;	//m3
							double xarea = vol / pBMP->m_lfBasinLength;//ft2
							double slope = pBMP->m_lfSideSlope3;
							double slope1 = pBMP->m_lfSideSlope1;
							double slope2 = pBMP->m_lfSideSlope2;
							double r__1 = avdepe / slope1;
							double r__2 = avdepe / slope2;
							double wet_p = pBMP->m_lfBasinWidth + 
										   sqrt(avdepe*avdepe + r__1*r__1) + 
										   sqrt(avdepe*avdepe + r__2*r__2); //ft
							double hrade = xarea / wet_p; //ft
							double hradm = hrade * FOOT2METER;	// m
							double tw = 20.0;//degreeC
							double avvele = 1.49/pBMP->m_lfManning*pow(hrade,2/3)*
											pow(slope,0.5); //ft/sec
							double avvelm = avvele * FpS2MpS;//m/s
							
							double js   = 0.0;

							if (fabs(ros) > fThreshold)
							{
								double rat = vols / (ros * delts);
								if (rat < crrat)
									js = rat / crrat;
								else
									js = 1.0;
							}

							double cojs    = 1.0 - js;
							double srovol  = js * ros * delts;	//ft3
							double erovol  = cojs * ro * delts;	//ft3
							double srovolm = srovol * CF2CM;	//m3
							double erovolm = erovol * CF2CM;	//m3

							conc   = bmpc[i*NWQ+j];		// lb/ft3
							double concs  = bmpc[i*NWQ+j];
							double rsed1  = rbsed[i*NWQ+j];		// lb
							double rsed1tot = rbsedtot[i];		// lb

							//simulate sediment
							if (pBMPData->nSedflag[j] == SAND)
							{
								//simulate sand
								int    sandfg = 3;	// user-specified power function method
								double ksand  = pBMPSite->m_sand.m_lfKSAND;
								double expsnd = pBMPSite->m_sand.m_lfEXPSND;
								double db50e  = pBMPSite->m_sand.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double w = pBMPSite->m_sand.m_lfW*0.0254*delts;	// fall velocity (m/ivl)
								double wsande = w * 3.28 / delts; // from m/ivl to ft/sec
								double depscr = 0., rsed = 0.;
									
								//dummy variables (not required for sandfg = 3)
								double twide = 0.,fsl = 0.;

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2 *= POUND2GRAM;	//gram
									rsed1   *= POUND2GRAM;	//gram
									conc    *= LBpCFT2MGpL;	//mg/l

									sandld(massin2,volsm,srovolm,volm,erovolm,ksand,avvele, 
										   expsnd,rom,sandfg,db50e,hrade,slope,tw,wsande,
										   twide,db50m,fsl,avdepe,&conc,&rsed,&rsed1,&depscr,
										   &romat);

									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	(g)
											rsed1  += rsed;	// mg/l*m3	(g)
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	//lb/ft3
									romat /= POUND2GRAM;	//lb
									rsed1 /= POUND2GRAM;	//lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin2 + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3
								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb

								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else if (pBMPData->nSedflag[j] == SILT)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_silt.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_silt.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_silt.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_silt.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)

								double db50e  = pBMPSite->m_silt.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
								  
									if (avdepm > 0.0)
									{
										//use formula appropriate to a river or stream
										tau = gamma * hradm * slope;// (kg/m^2)
									}

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3

								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb
								rbsedtot[i] += rsed1;
								
								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else if (pBMPData->nSedflag[j] == CLAY)
							{
								//convert settling velocity from in/sec to m/ivl
								double w = pBMPSite->m_clay.m_lfW*0.0254*delts;	// fall velocity (m/timestep)
								double taucd = pBMPSite->m_clay.m_lfTAUCD*0.4535924/0.09290304;	//critical bed shear stress for deposition (kg/m^2)
								double taucs = pBMPSite->m_clay.m_lfTAUCS*0.4535924/0.09290304;	//critical bed shear stress for scour (kg/m^2)
								//convert erodibility coeff from lb/ft^2/day to kg/m^2/ivl
								double m = pBMPSite->m_clay.m_lfM*deltd*0.4535924/0.09290304;	// erodibility coefficient of the sediment (kg/m^2/timestep)

								double db50e  = pBMPSite->m_clay.m_lfD/12.0;//ft
								double db50m  = db50e * 304.8; // mm
								double gamma = 1000;	// kg/m^3	 
								double tau = 0;			// (kg/m^2)		  

								if (fabs(volm) > fThreshold)
								{
									//convert to metric units
									massin2  *= POUND2GRAM;	//gram
									rsed1    *= POUND2GRAM;	//gram
									rsed1tot *= POUND2GRAM;	//gram
									conc     *= LBpCFT2MGpL;//mg/l
									romat     = 0.0;

									advect(massin2,volsm,rosm,volm,rom,delts,crrat,conc,romat);
									
									double depscr = 0;		// mg/l*m3 = g
									double frcsed1 = 0.0;

									if (rsed1tot > 0.0)
									{
										frcsed1 = rsed1 / rsed1tot;
									}
									else
									{
										//no bed at start of interval, assume equal fractions 
										frcsed1 = 0.5;
									}

									//calculate exchange between bed and suspended sediment
									double rsed  = conc*volm;	// storage in suspension (g)
								  
									if (avdepm > 0.0)
									{
										//use formula appropriate to a river or stream
										tau = gamma * hradm * slope;// (kg/m^2)
									}

									if(avdepe > 0.17)
										BDEXCH(avdepm,w,tau,taucd,taucs,m,volm,frcsed1,&rsed,&rsed1,&depscr);	

									//update concentration 
									if (volm > 0)
										conc = rsed/volm;				// mg/l
									
									//set small concentrations to zero
									if (fabs(conc) < SMALLNUM) 
									{
										//small conc., set to zero
										if (depscr > 0.0) 
										{
											//deposition has occurred, add small storage to deposition
											depscr += rsed;	// mg/l*m3	
											rsed1  += rsed;	// mg/l*m3
										}
										else
										{
											//add small storage to outflow
											romat += rsed;
											depscr = 0.0;
										}
										rsed = 0.0;
										conc = 0.0;
									}

									//convert back to english units
									conc  /= LBpCFT2MGpL;	// lb/ft3
									romat /= POUND2GRAM;	// lb
									rsed1 /= POUND2GRAM;	// lb
								}
								else
								{
									// bmp has gone dry during the interval; 
									// set conc equal to zero
									conc = 0;

									// calculate total amount of material leaving  
									// during the interval;  
									// this is equal to material inflow + material initially 
									// present  
									if (ro > 0)
									{
										romat = massin + concs * vols;	
										if (romat < 0)	romat = 0;
									}
									else
									{
										romat = 0;	
									}
								}

								// update values
								bmpc[i*NWQ+j] = conc;// lb/ft3

								romat_w = 0.0;
								romat_o = 0.0;
								romat_i = 0.0;

								if (ro > 0)
								{
									romat_w = weir/ro*romat;
									romat_o = orifice/ro*romat;
									romat_i = infilt/ro*romat;
								}

								rbsed[i*NWQ+j] = rsed1;// lb
								rbsedtot[i] += rsed1;
								
								//soil and under-drain column concentration
								double conc2 = bmpc2[i*NWQ+j];//lb/ft3
								double romat2 = 0.0;//lb		
								double qout21 = undrain_p[i] + seepage_p[i];//cfs		
								double qout22 = udout + seepage;//cfs		
								double vol21 = max(0.0, vol_p[i]-bmpvol_p[i]);//ft3		
								double vol22 = max(0.0, vol2 - ovolume);//ft3	
								
								advect(romat_i,vol21,qout21,vol22,qout22,delts,crrat,conc2,romat2);
								
								bmpc2[i*NWQ+j] = conc2;// lb/ft3

								romat_ud = 0.0;
								if (qout22 > 0)
									romat_ud = udout/qout22*romat2;;
								//apply reduction
								romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
							}
							else	// not sediment
							{
								if (nRoutingMethod <= 1)
								{
									conc = bmpc[i*NWQ+j];			//lb/ft3		
									advect(massin2, vol1, qout1, vol2, qout2, delts, crrat, 
										conc, romat);

									if (nRemovalMethod == 0)
									{
										//1st order decay
										decay = pBMPSite->m_pDecay[j];
										conc = conc * exp(-decay);	// decay is per hour
									}
									else if (nRemovalMethod == 1)
									{
										//kadlec and knight method
										//Cout = Cstar + (Cin - Cstar) * exp (-k/q)
										lfK = pBMPSite->m_pK[j];
										lfCstar = pBMPSite->m_pCstar[j];

										//maintain minimum conc of lfCstar
										conc = max(conc, lfCstar);

										double lfq = bmpout*delts/BMPAREA;
										if (lfq > 0 && conc > lfCstar)
											conc = lfCstar + (conc - lfCstar) * exp(-lfK/lfq);
									}
								}
								else if (nRoutingMethod > 1)
								{
									//CSTRs in series routing
									int nCSTRs = nRoutingMethod;
									float v = vol1/nCSTRs;			//ft3
									float lfmassin = massin2/delts;	//lb/sec
									float lfinflow = oinflow2/delts;//cfs
									float lfoutflow = qout2;		//cfs
									float tStep = delts;			//sec/hr

									for(k=0; k<nCSTRs; k++)
									{
										//calculate inputs for each CSTR
										float c = pBMPSite->m_pConc[j*nCSTRs+k];	//lb/ft3
										float wIn = lfmassin;						//lb/sec
										float qNet = lfinflow+(lfoutflow-lfinflow)*k/nCSTRs;		//ft3/sec
										float qNetout = lfinflow+(lfoutflow-lfinflow)*(k+1)/nCSTRs;	//ft3/sec

										//get the new conc (lb/ft3)
										conc = getCstrQual(c,v,wIn,qNet,tStep);

										//calculate massout (lb/ivl)
										romat = conc * qNetout * delts;

										if (nRemovalMethod == 0)
										{
											//1st order decay
											decay = pBMPSite->m_pDecay[j];
											conc = conc * exp(-decay);	// decay is per hour
										}
										else if (nRemovalMethod == 1)
										{
											//kadlec and knight method
											//Cout = Cstar + (Cin - Cstar) * exp (-k/q)
											lfK = pBMPSite->m_pK[j];
											lfCstar = pBMPSite->m_pCstar[j];
										
											//maintain minimum conc of lfCstar
											conc = max(conc, lfCstar);

											double lfq = lfoutflow*delts/BMPAREA;
											if (lfq > 0 && conc > lfCstar)
												conc = lfCstar + (conc - lfCstar) * exp(-lfK/lfq);
										}

										//update conc
										pBMPSite->m_pConc[j*nCSTRs+k] = conc;

										//calculate massin to the next reactor(lb/sec)
										lfmassin = romat / delts;
									}
								}

								if (qout2 > 0)
								{
									romat_w = weir/qout2*romat;
									romat_o = orifice/qout2*romat;
									romat_ud = udout/qout2*romat;
									romat_ud *= (1.0 - pBMPSite->m_pUndRemoval[j]);
								}
							}
						}
						else
						{
							//no volume stored
							//conc = massin / oinflow;//lb/ft3	
							conc = 0.0;//unknown
							romat_w = 0.0;
							romat_o = 0.0;
							romat_ud = 0.0;
							romat_ut = massin;//lb
							romat = 0.0;
						}
					}
					else if (pBMPSite->m_nBMPClass == CLASS_C && nRunMode != RUN_PREDEV)			
					{
						BMP_C* pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
						int nIndex = pBMP->m_nIndex;
						double BMPvolume = Conduit[nIndex].length * xsect_getAmax(&Link[nIndex].xsect);			
					
						if (BMPvolume > SMALLNUM && Link[nIndex].xsect.type != DUMMY)
						{
							conc = Link[nIndex].newQual[j];//lb/ft3
							romat_w = 0.0;
							romat_o = romat_o2[i*NWQ+j];//lb
							romat_ud = 0.0;
							romat_ut = romat_ut2[i*NWQ+j];//lb
							romat = romat_o;
							romat_o2[i*NWQ+j] = 0.0;
							romat_ut2[i*NWQ+j] = 0.0;
						}
						else
						{
							//no volume stored
							//conc = massin / oinflow;//lb/ft3	
							conc = 0.0;//unknown
							romat_w = 0.0;
							romat_o = 0.0;
							romat_ud = 0.0;
							romat_ut = massin;//lb
							romat = 0.0;
						}
					}
					else if (pBMPSite->m_nBMPClass == CLASS_D && nRunMode != RUN_PREDEV && nRunMode != RUN_POSTDEV)
					{
						BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
						BMPAREA = pBMP->m_lfLength * pBMP->m_lfWidth;

						if (BMPAREA > SMALLNUM) 
						{
							//get VFSMOD results
							conc = 0.0;//unknown
							romat_w = 0.0;
							romat_o = 0.0;//unknown
							romat_ud = 0.0;
							romat_ut = massin;//lb
							romat = 0.0;
						}
						else
						{
							//no volume stored
							//conc = massin / oinflow;//lb/ft3	
							conc = 0.0;//unknown
							romat_w = 0.0;
							romat_o = 0.0;
							romat_ud = 0.0;
							romat_ut = massin;//lb
							romat = 0.0;
						}
					}
					else//dummy BMP or assessment point
					{
						//no volume stored
						//conc = massin / oinflow;//lb/ft3	
						conc = 0.0;//unknown
						romat_w = 0.0;//lb
						romat_o = 0.0;//lb
						romat_ud = 0.0;//lb
						romat_ut = massin;//lb
						romat = 0.0;
					}
MM:
					//save values
					bmpc[i*NWQ+j] = conc;//lb/ft3
//					bmpconcout_s[i*NPOL+nPollIndex] += conc;						
					bmpmassout_w[i*NWQ+j]  = __max(0.0, romat_w);//lb
					bmpmassout_o[i*NWQ+j]  = __max(0.0, romat_o);//lb		
					bmpmassout_ud[i*NWQ+j] = __max(0.0, romat_ud);//lb
					bmpmassout_ut[i*NWQ+j] = __max(0.0, romat_ut);//lb
					bmpmassout[i*NWQ+j]    = __max(0.0, romat_w + romat_o + romat_ud + romat_ut);
					bmpmass_w_s[i*NPOL+nPollIndex] += romat_w;		
					bmpmass_o_s[i*NPOL+nPollIndex] += romat_o;		
					bmpmass_ud_s[i*NPOL+nPollIndex] += romat_ud;		
					bmpmass_ut_s[i*NPOL+nPollIndex] += romat_ut;		
					bmpmassout_s[i*NPOL+nPollIndex] += bmpmassout[i*NWQ+j];		

					double value0 = pBMPSite->m_RAConc[nPollIndex].qMass.front();
					//double value1 = romat; //lb
					double value1 = bmpmassout[i*NWQ+j]; //lb
					lfSumMass[i*NPOL+nPollIndex] -= value0;
					lfSumMass[i*NPOL+nPollIndex] += value1;

					//round to 3rd decimal place
					//lfSumMass[i*NWQ+j] = floor(lfSumMass[i*NWQ+j]*1.0E3+0.5)/1.0E3;
					pBMPSite->m_RAConc[nPollIndex].qMass.pop();
					pBMPSite->m_RAConc[nPollIndex].qMass.push(value1);

					// evaluation factor calculation
					if (pBMPSite->m_factorList.GetCount() > 0)
					{
						//Priya compute the cumulative load
						//bmpTotalLoad[i*NWQ+j] += bmpmassout[i*NWQ+j];	// cumulative load (lb/hr)
						bmpTotalLoad[i*NPOL+nPollIndex] += bmpmassout[i*NWQ+j];// cumulative load (lb/hr)

						double bmpoutconc = 0.0;
						if (lfFlowVolume > 20)	//ft3/day
							bmpoutconc = lfSumMass[i*NPOL+nPollIndex]/(lfSumFlow[i]*28.31685); //lb/liter
						
						if (pBMPData->nWeatherFile == 1)
						{
							if (pBMPData->pWEATHERDATA[pBMPData->lStartIndex+t].bWetInt)
							{
								if (bmpoutconc > pBMPSite->m_RAConc[nPollIndex].m_lfThreshConc)
								{
									nExceedConc[i*NPOL+nPollIndex]++;
								}
							}
						}

						if (pBMPSite->m_RAConc[nPollIndex].m_nRDays > 0)
						{
							int nRHours = pBMPSite->m_RAConc[nPollIndex].m_nRDays*24;
							pBMPSite->m_RAConc[nPollIndex].m_lfRFlow[nRHours-1] = bmpout;					// cfs
							pBMPSite->m_RAConc[nPollIndex].m_lfRLoad[nRHours-1] += bmpc[i*NWQ+j]*bmpout;	// lb/ft3*cfs
						}
					}
					nPollIndex++;
				}
			}

			// evaluation factor calculation
			for (j=0; j<NPOL; j++)
			{
				if (pBMPSite->m_factorList.GetCount() > 0)
				{
					double RTLoad = 0;
					double RTFlow = 0;
					double RAConc = 0;
					if (pBMPSite->m_RAConc[j].m_nRDays > 0)
					{
						int nRHours = pBMPSite->m_RAConc[j].m_nRDays*24;

						for (k=0; k<nRHours; ++k)
						{
							RTFlow += pBMPSite->m_RAConc[j].m_lfRFlow[k];
							RTLoad += pBMPSite->m_RAConc[j].m_lfRLoad[k];
						}

						if (RTFlow > 0)
						RAConc = RTLoad/RTFlow*LBpCFT2MGpL;//mg/l
						
						if (bmpMAConc[i*NPOL+j] < RAConc)
							bmpMAConc[i*NPOL+j] = RAConc;

						for (k=0; k<(nRHours-1); ++k)
						{
							pBMPSite->m_RAConc[j].m_lfRFlow[k] = pBMPSite->m_RAConc[j].m_lfRFlow[k+1];
							pBMPSite->m_RAConc[j].m_lfRLoad[k] = pBMPSite->m_RAConc[j].m_lfRLoad[k+1];
						}

						pBMPSite->m_RAConc[j].m_lfRFlow[nRHours-1] = 0.0;			
						pBMPSite->m_RAConc[j].m_lfRLoad[nRHours-1] = 0.0;
					}
				}
			}

			//save for the next timestep
			bmpoflow_w[i]		= weir;//cfs							
			bmpoflow_o[i]		= orifice;//cfs						
			bmpoflow_ud[i]		= udout;//cfs							
			bmpoflow_ut[i]		= utout;//cfs
			bmpoflow[i]			= bmpout;//cfs

			vol_p[i]			= vol2;		//ft3
			bmpvol_p[i]			= ovolume;	//ft3
			osa_p[i]			= osa;		//in
			ostorage_p[i]		= ostorage;	//in
			weir_p[i]			= weir;		//cfs
			orifice_p[i]		= orifice;	//cfs
			infilt_p[i]			= infilt;	//cfs
			undrain_p[i]		= udout;	//cfs
			seepage_p[i]		= seepage;	//cfs
			counter_p[i]		= counter;

			//output variables
			BmpFlowInput_s[i]	+= oinflow/3600.0;// ft^3/hr to cfs 
			bmpvol_s[i]			+= ovolume;	//ft3		
			bmpstage_s[i]		+= ostage;	//ft		
			infilt_s[i]			+= infilt;	//cfs		
			perc_s[i]			+= perc;	//cfs		
			AET_s[i]			+= AET;		//cfs		
			seepage_s[i]		+= seepage;	//cfs		
//			usstorage_s[i]		+= (nsamax-osa)/12*BMParea_max*ndevice;		// ft3		
//			udstorage_s[i]		+= (nstoragemax-ostorage)/12*BMParea*ndevice;	// ft3			
			weir_s[i]			+= weir;		//cfs			
			orifice_s[i]		+= orifice;	//cfs		
			bmpudout_s[i]		+= udout;		//cfs			
			bmpbypass_s[i]		+= utout;		//cfs			
			bmpoutflow_s[i]		+= bmpoflow[i];//cfs	

			// output time series results 
 			if(nRunMode != RUN_OPTIMIZE)
			{
				// save output to files
				if (pBMPSite->m_factorList.GetCount() > 0)
				{
					int ptoption;
					if (pBMPData->nOutputTimeStep == 0)	// daily
						ptoption = 24;
					else 
						ptoption = 1;					// hourly

					if(ptoption == 1 || (ptoption == 24 && t%24 == 0 && t != 0))
					{
						double cv = 1.0;
						if(ptoption == 24)
							cv = 24.0;

						int nIndex = ::FindObIndexFromList(pBMPData->bmpsiteList, pBMPSite) + 1;
						CString strContent;
						strContent.Format("%s  %d  %d  %d  %d  %d\t", pBMPSite->m_strID, nYear, mon+1, day+1, hour+1, 0);

						CString strAdd;
						strAdd.Format("%e\t%e\t%e\t%e\t%e\t%e\t%e\t%e\t%e\t%e\t%e\t%e\t",
							bmpvol_s[i]/cv,						// ft^3
							bmpstage_s[i]/cv,					// ft
							BmpFlowInput_s[i]/cv,				// cfs
							weir_s[i]/cv,						// cfs
							orifice_s[i]/cv,					// cfs
							bmpudout_s[i]/cv,					// cfs
							bmpbypass_s[i]/cv,					// cfs
							bmpoutflow_s[i]/cv,					// cfs
							infilt_s[i]/cv, 					// cfs
							perc_s[i]/cv,						// cfs
							AET_s[i]/cv,						// cfs
							seepage_s[i]/cv);					// cfs
//							usstorage_s[i]/cv,					// cfs
//							udstorage_s[i]/cv);					// cfs
						strContent += strAdd;

						for(j=0; j<NPOL; j++)
						{
							strAdd.Format("%e\t%e\t%e\t%e\t%e\t%e\t", BmpWqInput_s[i*NPOL+j], bmpmass_w_s[i*NPOL+j], bmpmass_o_s[i*NPOL+j], bmpmass_ud_s[i*NPOL+j], bmpmass_ut_s[i*NPOL+j], bmpmassout_s[i*NPOL+j]);	// lb
							strContent += strAdd;

							bmpconcout_s[i*NPOL+j] = 0.0;
							if (bmpoutflow_s[i] > 0)
								bmpconcout_s[i*NPOL+j] = bmpmassout_s[i*NPOL+j]/(bmpoutflow_s[i]*3600.00);//lb/ft3

							strAdd.Format("%e", bmpconcout_s[i*NPOL+j]*LBpCFT2MGpL);  //mg/l

							if(j != NWQ-1)
								strAdd += "\t";
							strContent += strAdd;

							//re-initialize
							BmpWqInput_s[i*NPOL+j] = 0.0;	// Mass entering the BMP (lbs)
							bmpmass_w_s[i*NPOL+j] = 0.0;	// Mass leaving the BMP (lbs)
							bmpmass_o_s[i*NPOL+j] = 0.0;	// Mass leaving the BMP (lbs)
							bmpmass_ud_s[i*NPOL+j] = 0.0;	// Mass leaving the BMP (lbs)
							bmpmass_ut_s[i*NPOL+j] = 0.0;	// Mass bypassing the BMP (lbs)
							bmpmassout_s[i*NPOL+j] = 0.0;	// Mass leaving the BMP (lbs)
							bmpconcout_s[i*NPOL+j] = 0.0;	// Outflow concentration (mg/l)
						}

						strContent += "\n";
						fputs((LPCSTR)strContent, pBMPSite->m_fileOut);
						fflush(pBMPSite->m_fileOut);

						// re-initialize
						bmpvol_s[i]       = 0.0;	// BMP volume (ft3)
						infilt_s[i]       = 0.0;	// Infiltration (cfs)
						perc_s[i]         = 0.0;	// percolation (cfs)
						AET_s[i]          = 0.0;	// actual evapotranspiration (cfs)
						seepage_s[i]      = 0.0;	// seepage (cfs)
//						usstorage_s[i]    = 0.0;	// BMP volume (ft3)
//						udstorage_s[i]    = 0.0;	// BMP volume (ft3)
						weir_s[i]         = 0.0;	// Weir outflow (cfs)
						orifice_s[i]      = 0.0;	// Orifice or channel outflow (cfs)
						BmpFlowInput_s[i] = 0.0;	// Total inflow (cfs)
						bmpstage_s[i]     = 0.0;	// Water depth (ft)
						bmpudout_s[i]     = 0.0;	// Underdrain outflow (cfs)
						bmpbypass_s[i]    = 0.0;	// untreated outflow (cfs)
						bmpoutflow_s[i]	  = 0.0;	// Total outflow (cfs)
					}
				}
			}
		}

		if(dayflag == t)
		{
			dayflag += 24; 

			CString strMsg, strForDdg, strE;
			COleDateTimeSpan span = COleDateTimeSpan(0,t,0,0);
			COleDateTime tCurrent = tStart + span;

			int nSMonth = tCurrent.GetMonth();
			int nSDay = tCurrent.GetDay();
			int nSYear = tCurrent.GetYear();

			GetLocalTime(&tm);
			tSpan = COleDateTime(tm) - time_i;
			int dd_elap = int(tSpan.GetDays());
			int hh_elap = int(tSpan.GetHours());
			int mm_elap = int(tSpan.GetMinutes());
			int ss_elap = int(tSpan.GetSeconds());

			if (nRunMode == RUN_INIT)
			{
				strMsg.Format("BMP Simulation:\t Initial BMPs Scenario\n");
				strForDdg = strMsg;
			}
			else if (nRunMode == RUN_PREDEV)
			{
				strMsg.Format("BMP Simulation:\t Pre-Development Scenario\n");
				strForDdg = strMsg;
			}
			else if (nRunMode == RUN_POSTDEV)
			{
				strMsg.Format("BMP Simulation:\t Post-Development Scenario\n");
				strForDdg = strMsg;
			}
			else if (nRunMode == RUN_OPTIMIZE)
			{
				strMsg.Format("BMP Simulation:\t Optimization Scenario %d of %d\n", optcounter, nMaxRun);
				strForDdg = strMsg;
			}
			else if (nRunMode == RUN_OUTPUT)
			{
				strMsg.Format("BMP Simulation:\t Best Solution Scenario %d of %d\n", outcounter, pBMPData->nSolution);
				strForDdg = strMsg;
			}

			strMsg.Format("Calculating:\t %02d-%02d-%04d\n", nSMonth, nSDay, nSYear);
			strForDdg += strMsg;

			if (nRunMode == RUN_OPTIMIZE)
				strE.Format("\nTime Elapsed:\t %02d:%02d:%02d:%02d\tMax Run Time: %2.2lf hrs", dd_elap, hh_elap, mm_elap, ss_elap, pBMPData->lfMaxRunTime);
			else
				strE.Format("\nTime Elapsed:\t %02d:%02d:%02d:%02d\n", dd_elap, hh_elap, mm_elap, ss_elap);
			
			strForDdg += strE;

			double lfPart = span.GetTotalSeconds();
//			double lfAll  = tSpan0.GetTotalSeconds();
			double lfPerc = lfPart/lfAll;

			if(pWndProgress->GetSafeHwnd() != NULL)
			{
				pWndProgress->SetText(strForDdg);
				pWndProgress->SetPos((int)(lfPerc*100));
				pWndProgress->PeekAndPump();
			}

			if(pWndProgress->Cancelled()) 
				goto L001;
		}
	}
L001:
	// find the run time if (nRunMode == RUN_INIT)
	if (nRunMode == RUN_INIT)
	{
		GetLocalTime(&tm);
		tSpan = COleDateTime(tm) - time_i;
		lInitRunTime = double(tSpan.GetTotalSeconds()*1000.00);	//milliseconds
	}
	
	// evaluation factor calculation
	lfNumOfYears = N/(24.0 * 365.0);//number of year 

	for (i=0; i<NBMP; i++)
	{			          
		// compute evaluation factors
		pos = pBMPData->routeList.FindIndex(i);
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetAt(pos);
		if (pBMPSite->m_factorList.GetCount() > 0)
		{
			if (lfNumOfYears > 0)
			{
				bmpAAFlowVol[i] = bmpTotalFlow[i] / lfNumOfYears;	// ft3/yr
				bmpFlowExcFreq[i] = nExceedFlow[i] / lfNumOfYears;	// /yr
				
				for (j=0; j<NPOL; j++)
				{
					bmpAALoad[i*NPOL+j] = bmpTotalLoad[i*NPOL+j] / lfNumOfYears;// lb/yr
					if (bmpAAFlowVol[i] > 0)
						bmpAAConc[i*NPOL+j] = bmpAALoad[i*NPOL+j] / bmpAAFlowVol[i] * LBpCFT2MGpL;	//mg/l

					bmpConcExcDays[i*NPOL+j] = nExceedConc[i*NPOL+j] / (24.0 * lfNumOfYears);
					//lfWetDaysPerYear = pBMPData->lfWetDays/lfNumOfYears;
					lfWetDaysPerYear = pBMPData->nWetInt/(24.0 * lfNumOfYears);
				}
			}
		}
	}

	// start loop through each BMP site for assigning and outputting evaluation factors
	for (i=0; i<NBMP; i++)
	{
		pos = pBMPData->routeList.FindIndex(i);
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetAt(pos);
		int nIndex = ::FindObIndexFromList(pBMPData->bmpsiteList, pBMPSite) + 1;
		CString strValue = "";
		if (pBMPSite->m_factorList.GetCount() > 0)
		{
			if(nRunMode == RUN_INIT)			// Existing Scenario (Before Optimization Scenario)
			{
				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1)
					{
						if (ef->m_nFactorType == AAFV)
							ef->m_lfInit = bmpAAFlowVol[i];
						else if (ef->m_nFactorType == PDF)
							ef->m_lfInit = bmpPkDisFlow[i];
						else if (ef->m_nFactorType == FEF)
							ef->m_lfInit = bmpFlowExcFreq[i];
					}
					else if (ef->m_nFactorGroup > 0)
					{
						if (ef->m_nFactorType == AAL)
							ef->m_lfInit = bmpAALoad[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == AAC)
							ef->m_lfInit = bmpAAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == MAC)
							ef->m_lfInit = bmpMAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == CEF)	
							ef->m_lfInit = bmpConcExcDays[i*NPOL + ef->m_nFactorGroup - 1];
					}
					else
						AfxMessageBox("Check card 815!, Factor Group can not have zero value");
					
					CString strFactor = ef->m_strFactor;
					if (strFactor.Right(2) == "_%")
						strFactor.TrimRight("_%");
					else if (strFactor.Right(2) == "_S")
						strFactor.TrimRight("_S");
					else if (strFactor.Right(2) == "_s")
						strFactor.TrimRight("_s");

					if (ef->m_nFactorType == CEF)	
						strValue.Format("%s\t%s\t%.2lf\t%.2lf\n", pBMPSite->m_strID, strFactor, ef->m_lfInit, lfWetDaysPerYear);
					else
						strValue.Format("%s\t%s\t%.5lf\t%.5lf\n", pBMPSite->m_strID, strFactor, ef->m_lfInit, totalCost);
					fputs(LPCSTR(strValue), fp);
				}
			}
			else if(nRunMode == RUN_PREDEV)			// PRE-DEVELOPED Scenario (Before Optimization Scenario)
			{
				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1)
					{
						if (ef->m_nFactorType == AAFV)
							ef->m_lfPreDev = bmpAAFlowVol[i];
						else if (ef->m_nFactorType == PDF)
							ef->m_lfPreDev = bmpPkDisFlow[i];
						else if (ef->m_nFactorType == FEF)
							ef->m_lfPreDev = bmpFlowExcFreq[i];
					}
					else if (ef->m_nFactorGroup > 0)
					{
						if (ef->m_nFactorType == AAL)
							ef->m_lfPreDev = bmpAALoad[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == AAC)
							ef->m_lfPreDev = bmpAAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == MAC)
							ef->m_lfPreDev = bmpMAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == CEF)	
							ef->m_lfPreDev = bmpConcExcDays[i*NPOL + ef->m_nFactorGroup - 1];
					}
					else
						AfxMessageBox("Check card 815!, Factor Group can not have zero value");

					CString strFactor = ef->m_strFactor;
					if (strFactor.Right(2) == "_%")
						strFactor.TrimRight("_%");
					else if (strFactor.Right(2) == "_S")
						strFactor.TrimRight("_S");
					else if (strFactor.Right(2) == "_s")
						strFactor.TrimRight("_s");

					if (ef->m_nFactorType == CEF)	
						strValue.Format("%s\t%s\t%.2lf\t%.2lf\n", pBMPSite->m_strID, strFactor, ef->m_lfPreDev, lfWetDaysPerYear);
					else
						strValue.Format("%s\t%s\t%.5lf\t%.5lf\n", pBMPSite->m_strID, strFactor, ef->m_lfPreDev, 0.0);
					fputs(LPCSTR(strValue), fp);
				}
			}
			else if(nRunMode == RUN_POSTDEV)// POST-DEVELOPED Scenario (Before Optimization and Without BMPs)
			{
				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1)
					{
						if (ef->m_nFactorType == AAFV)
							ef->m_lfPostDev = bmpAAFlowVol[i];
						else if (ef->m_nFactorType == PDF)
							ef->m_lfPostDev = bmpPkDisFlow[i];
						else if (ef->m_nFactorType == FEF)
							ef->m_lfPostDev = bmpFlowExcFreq[i];
					}
					else if (ef->m_nFactorGroup > 0)
					{
						if (ef->m_nFactorType == AAL)
							ef->m_lfPostDev = bmpAALoad[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == AAC)
							ef->m_lfPostDev = bmpAAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == MAC)
							ef->m_lfPostDev = bmpMAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == CEF)	
							ef->m_lfPostDev = bmpConcExcDays[i*NPOL + ef->m_nFactorGroup - 1];
					}
					else
						AfxMessageBox("Check card 815!, Factor Group can not have zero value");

					CString strFactor = ef->m_strFactor;
					if (strFactor.Right(2) == "_%")
						strFactor.TrimRight("_%");
					else if (strFactor.Right(2) == "_S")
						strFactor.TrimRight("_S");
					else if (strFactor.Right(2) == "_s")
						strFactor.TrimRight("_s");

					if (ef->m_nFactorType == CEF)	
						strValue.Format("%s\t%s\t%.2lf\t%.2lf\n", pBMPSite->m_strID, strFactor, ef->m_lfPostDev, lfWetDaysPerYear);
					else
						strValue.Format("%s\t%s\t%.5lf\t%.5lf\n", pBMPSite->m_strID, strFactor, ef->m_lfPostDev, 0.0);
					fputs(LPCSTR(strValue), fp);
				}
			}
			else if(nRunMode == RUN_OPTIMIZE)	//optimizing
			{
				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1)
					{
						if (ef->m_nFactorType == AAFV)
							ef->m_lfCurrent = bmpAAFlowVol[i];
						else if (ef->m_nFactorType == PDF)
							ef->m_lfCurrent = bmpPkDisFlow[i];
						else if (ef->m_nFactorType == FEF)
							ef->m_lfCurrent = bmpFlowExcFreq[i];
					}
					else if (ef->m_nFactorGroup > 0)
					{
						if (ef->m_nFactorType == AAL)
							ef->m_lfCurrent = bmpAALoad[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == AAC)
							ef->m_lfCurrent = bmpAAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == MAC)
							ef->m_lfCurrent = bmpMAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == CEF)	
							ef->m_lfCurrent = bmpConcExcDays[i*NPOL + ef->m_nFactorGroup - 1];
					}
					else
						AfxMessageBox("Check card 815!, Factor Group can not have zero value");
				}
			}
			else if(nRunMode == RUN_OUTPUT)	//outputting
			{
				pos1 = pBMPSite->m_factorList.GetHeadPosition();
				while (pos1 != NULL)
				{
					EVALUATION_FACTOR* ef = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
					if (ef->m_nFactorGroup == -1)
					{
						if (ef->m_nFactorType == AAFV)
							ef->m_lfCurrent = bmpAAFlowVol[i];
						else if (ef->m_nFactorType == PDF)
							ef->m_lfCurrent = bmpPkDisFlow[i];
						else if (ef->m_nFactorType == FEF)
							ef->m_lfCurrent = bmpFlowExcFreq[i];
					}
					else if (ef->m_nFactorGroup > 0)
					{
						if (ef->m_nFactorType == AAL)
							ef->m_lfCurrent = bmpAALoad[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == AAC)
							ef->m_lfCurrent = bmpAAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == MAC)
							ef->m_lfCurrent = bmpMAConc[i*NPOL + ef->m_nFactorGroup - 1];
						else if (ef->m_nFactorType == CEF)	
							ef->m_lfCurrent = bmpConcExcDays[i*NPOL + ef->m_nFactorGroup - 1];
					}
					else
						AfxMessageBox("Check card 815!, Factor Group can not have zero value");

					if (pBMPData->nRunOption == OPTION_MIMIMIZE_COST)
					{
						double output1 = 0;
						if (ef->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
						{
							if (ef->m_lfInit > 0)	
								output1 = ef->m_lfCurrent/ef->m_lfInit*100;
//							else
//								output1 = ef->m_lfCurrent;
						}
						else if (ef->m_nCalcMode == CALC_VALUE) // if the calculation mode is value
						{
							output1 = ef->m_lfCurrent;
						}
						else // if the calculation mode is scale
						{
							if (ef->m_lfInit - ef->m_lfPreDev > 0)
								output1 = (ef->m_lfCurrent - ef->m_lfPreDev) / (ef->m_lfInit-ef->m_lfPreDev);
//							else
//								output1 = ef->m_lfCurrent;
						}
						strValue.Format("%s\t%s\t%.5lf\t%.5lf\n", pBMPSite->m_strID, ef->m_strFactor, output1, ef->m_lfTarget);
						fputs(LPCSTR(strValue), fp);
					}
					else if (pBMPData->nRunOption == OPTION_MAXIMIZE_CONTROL)
					{
//						if (ef->m_nFactorType == CEF)
//							strValue.Format("%s\t%s\t%.2lf\t%.2lf\n", pBMPSite->m_strID, ef->m_strFactor, ef->m_lfCurrent, lfWetDaysPerYear);
//						else
							strValue.Format("%s\t%s\t%.5lf\n", pBMPSite->m_strID, ef->m_strFactor, ef->m_lfCurrent);
						fputs(LPCSTR(strValue), fp);
					}
				}
			}
		}
	}

	fflush(fp);

	//release memory
	if (counter_p != NULL)			delete []counter_p;
	if (BmpFlowInput != NULL) 		delete []BmpFlowInput;
	if (bmpoflow != NULL)			delete []bmpoflow;
	if (vol_p != NULL)				delete []vol_p;
	if (bmpvol_p != NULL)			delete []bmpvol_p;
	if (osa_p != NULL)				delete []osa_p;
	if (ostorage_p != NULL)			delete []ostorage_p;
	if (weir_p != NULL)				delete []weir_p;
	if (orifice_p != NULL)			delete []orifice_p;
	if (infilt_p != NULL)			delete []infilt_p;
	if (undrain_p != NULL)			delete []undrain_p;
	if (seepage_p != NULL)			delete []seepage_p;

	//output variables
	if (bmpvol_s != NULL)			delete []bmpvol_s;
	if (bmpstage_s != NULL)			delete []bmpstage_s;
	if (BmpFlowInput_s != NULL)		delete []BmpFlowInput_s;
	if (weir_s != NULL)				delete []weir_s;
	if (orifice_s != NULL)			delete []orifice_s;
	if (bmpudout_s != NULL)			delete []bmpudout_s;
	if (bmpbypass_s != NULL)		delete []bmpbypass_s;
	if (bmpoutflow_s != NULL)		delete []bmpoutflow_s;
	if (infilt_s != NULL)			delete []infilt_s;
	if (perc_s != NULL)				delete []perc_s;
	if (AET_s != NULL)				delete []AET_s;
	if (seepage_s != NULL)			delete []seepage_s;
//	if (usstorage_s != NULL)		delete []usstorage_s;
//	if (udstorage_s != NULL)		delete []udstorage_s;
	if (bmpoflow_w != NULL)			delete []bmpoflow_w;			
	if (bmpoflow_o != NULL)			delete []bmpoflow_o;			
	if (bmpoflow_ud != NULL)		delete []bmpoflow_ud;			
	if (bmpoflow_ut != NULL)		delete []bmpoflow_ut;
	if (rbsedtot != NULL)			delete []rbsedtot; 

	if (BmpWqInput != NULL)			delete []BmpWqInput;
	if (bmpc != NULL)				delete []bmpc;
	if (bmpc2 != NULL)				delete []bmpc2;
	if (bmpqconc_sand != NULL)		delete []bmpqconc_sand;
	if (bmpqconc_silt != NULL)		delete []bmpqconc_silt;
	if (bmpqconc_clay != NULL)		delete []bmpqconc_clay;
	if (bmpmassout != NULL)			delete []bmpmassout;
	if (bmpmassout_w != NULL)		delete []bmpmassout_w;
	if (bmpmassout_o != NULL)		delete []bmpmassout_o;
	if (bmpmassout_ud != NULL)		delete []bmpmassout_ud;
	if (bmpmassout_ut != NULL)		delete []bmpmassout_ut;
	if (bmudconc != NULL)			delete []bmudconc;
	if (romat_o2 != NULL)			delete []romat_o2; 
	if (romat_ut2 != NULL)			delete []romat_ut2; 
	if (rbsed != NULL)				delete []rbsed; 

	//output variables
	if (BmpWqInput_s != NULL)		delete []BmpWqInput_s;
	if (bmpmass_w_s != NULL)		delete []bmpmass_w_s;
	if (bmpmass_o_s != NULL)		delete []bmpmass_o_s;
	if (bmpmass_ud_s != NULL)		delete []bmpmass_ud_s;
	if (bmpmass_ut_s != NULL)		delete []bmpmass_ut_s;
	if (bmpmassout_s != NULL)		delete []bmpmassout_s;
	if (bmpconcout_s != NULL)		delete []bmpconcout_s;

	// evaluation factor calculation
	if (nExceedFlow != NULL)		delete []nExceedFlow;
	if (nExceedFlag != NULL)		delete []nExceedFlag;
	if (bmpTotalFlow != NULL)		delete []bmpTotalFlow;
	if (bmpAAFlowVol != NULL)		delete []bmpAAFlowVol;
	if (bmpPkDisFlow != NULL)		delete []bmpPkDisFlow;
	if (bmpFlowExcFreq != NULL)		delete []bmpFlowExcFreq;
	if (bmpTotalLoad != NULL)		delete []bmpTotalLoad; 
	if (bmpAALoad != NULL)			delete []bmpAALoad; 
	if (bmpAAConc != NULL)			delete []bmpAAConc; 
	if (bmpMAConc != NULL)			delete []bmpMAConc; 
	//optional	
	if (nExceedConc != NULL)		delete []nExceedConc;
	if (lfSumFlow != NULL)			delete []lfSumFlow; 
	if (lfSumMass != NULL)			delete []lfSumMass; 
	if (bmpConcExcDays != NULL)		delete []bmpConcExcDays;

	//release memory for the tradeoff curve
	pos = pBMPData->routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) pBMPData->routeList.GetNext(pos);

		if (pBMPSite->m_nBreakPoints > 1)
		{
			int nBrPtIndex = int(pBMPSite->m_lfBreakPtID + 2.0);
			
			if (nRunMode == RUN_POSTDEV)
				nBrPtIndex = 0;				//index = -2+2=0
			else if (nRunMode == RUN_PREDEV)
				nBrPtIndex = 1;				//index (-1+2=1)
			else if(nRunMode == RUN_INIT)
				nBrPtIndex = 2;				//index = 0+2=2

			pBMPSite->UnLoadTradeOffCurveData(nBrPtIndex);
		}
	}
}

bool CBMPRunner::OpenOutputFiles(const CString& runID, int nRunOption, int nRunMode)
{
	CString strFilePath;
	strFilePath.Format("%s_Eval.out", runID);
	strFilePath = pBMPData->strOutputDir + strFilePath;
	fp = fopen(LPCSTR(strFilePath), "wt");
	if(fp == NULL)
	{
		AfxMessageBox("Cannot open file " + strFilePath + " for writing.");
		return false;
	}

	WriteFileHeader(nRunOption, nRunMode);

	return true;
}

bool CBMPRunner::CloseOutputFiles()
{
	if (fp != NULL)
	{
		fclose(fp);
		fp = NULL;
	}

	return true;
}

void CBMPRunner::WriteFileHeader(int nRunOption, int nRunMode)
{
	fputs("TT-----------------------------------------------------------------------------------------\n",fp);
	fputs("TT\n",fp);
	fputs("TT SUSTAIN: System for Urban Stormwater Treatment and Analysis INtegration\n",fp);
	fputs("TT\n",fp);
	fputs("TT-----------------------------------------------------------------------------------------\n",fp);
	fputs("TT BMP Site Assessment Results\n",fp);
	fprintf(fp, "TT %s\n",SUSTAIN_VERSION);
	fputs("TT\n",fp);
	fputs("TT Designed and maintained by:\n",fp);
	fputs("TT     Tetra Tech, Inc.\n",fp);
	fputs("TT     10306 Eaton Place, Suite 340\n",fp);
	fputs("TT     Fairfax, VA 22030\n",fp);
	fputs("TT     (703) 385-6000\n",fp);
	fputs("TT-----------------------------------------------------------------------------------------\n",fp);

	fputs("TT	\n", fp);
	SYSTEMTIME tm;
	GetLocalTime(&tm);
	CString str;
	str.Format("TT This output file was created at %02d:%02d:%02d%s on %02d/%02d/%04d\n",(tm.wHour>12)?tm.wHour-12:tm.wHour,tm.wMinute,tm.wSecond,(tm.wHour>=12)?"pm":"am",tm.wMonth,tm.wDay,tm.wYear);
	fputs(LPCSTR(str),fp);

	fputs("TT    \n", fp);
	fputs("TT-----------------------------------------------------------------------------------------\n",fp);
	fputs("TT    \n", fp);

	if (nRunMode == RUN_INIT || nRunMode == RUN_PREDEV || nRunMode == RUN_POSTDEV)
		fprintf(fp, "Assessment Point (ID)     Factor Name     Factor Value     Total Cost\n");
	else if (nRunOption == OPTION_MIMIMIZE_COST)
		fprintf(fp, "Assessment Point (ID)     Factor Name     Factor Value		Target Value\n");
	else if (nRunOption == OPTION_MAXIMIZE_CONTROL)
		fprintf(fp, "Assessment Point (ID)     Factor Name     Factor Value\n");
	fflush(fp);
}				

//SWMM5
void findLinkQual2(int i,float tStep,double wAdded,double kDecay,double& c)
//
//  Input:   i = link index
//			 j = pollutant index
//           tStep = routing time step (sec)
//  Output:  none
//  Purpose: finds new quality in a link after the current time step.
//
{
    int   k;
    double q, vOld, vPlus;
//    float c;

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

    // --- find exponential 1st order decay over time step
//    c = Link[i].oldQual[j];
    if ( kDecay != 0.0 )
        c = c * exp(-kDecay * tStep);

    // --- combine inflow with old volume to compute new concen.
    if ( vPlus <= TINY ) c = 0.0;
    else c = (c*vOld + wAdded) / vPlus;
    c = MAX(c, 0.0);
//    Link[i].newQual[j] = c;

	return;
}

float getHydRad(TXsect* xsect, float y)
//
//  Input:   xsect = ptr. to conduit cross section
//           y     = flow depth (ft)
//  Output:  returns hydraulic radius (ft)
//  Purpose: computes hydraulic radius of flow cross-section in a conduit.
//
{
    float hRadius;                     // hyd. radius (ft)
    y = MIN(y, xsect->yFull);
    hRadius = xsect_getRofY(xsect, y);
    hRadius = MAX(hRadius, FUDGE);
    return hRadius;
}
