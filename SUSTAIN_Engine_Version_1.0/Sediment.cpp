//-----------------------------------------------------------------------------
//   Sediment.cpp
//
//   Project:  EPA SUSTAIN
//   Version:  1.0
//   Date:     3/20/07   
//
//-----------------------------------------------------------------------------

#include "stdafx.h"
#include <math.h>
#include <string.h>
#include "Sediment.h"


float dayval(float mval1,float mval2,int day,int ndays)
{
	//Linearly interpolate a value for this day (DAYVAL), given  
	//values for the start of this month and next month (MVAL1 and  
	//MVAL2).  ndays is the number of days in this month.  

    float rday, rndays;
	
    rday = day;
    rndays = ndays;

    return mval1 + (mval2 - mval1) * (rday - 1) / rndays;
}


int detach(int crvfg,int csnofg,int mon,int nxtmon,int day,int ndays,
		   float *coverm,float rain,float snocov,float delt60,float smpf, 
		   float krer,float jrer,float& cover,float& dets,float& det)
{
	//Detach soil by rainfall  
	float d1, d2, cr;
 
 	//it is the first interval of the day 
	if (crvfg == 1) 
	{
		// erosion related cover is allowed to vary throughout the 
		// year 
		// interpolate for the daily value 
		// linearly interpolate cover between two values from the 
		// monthly array coverm(12) 
		cover = dayval(coverm[mon], coverm[nxtmon], day, ndays);
	}
	else
	{
		// erosion related cover does not vary throughout the year. 
		// cover value has been supplied by the run interpreter 
	}

	if (rain > 0.0) 
	{
		// simulate detachment because it is raining 
		// find the proportion of area shielded from raindrop impact by 
		// snowpack and other cover 
		if (csnofg == 1) 	
		{
			// snow is being considered 
			if (snocov > 0.0) 
			{
				// there is a snowpack 
				cr = cover + (1.0 - cover) * snocov;
			} 
			else 
			{
				cr = cover;
			}
		} 
		else
		{
			cr = cover;
		}

		// calculate the rate of soil detachment, delt60= delt/60 -  
		// units are tons/acre-ivl  
		d1 =  (rain / delt60);	// rain is in in/ivl
		d2 =  (jrer);
		det = delt60 * (1.0 - cr) * smpf * krer * pow(d1, d2);
		
		// augment detached sediment storage - units are tons/acre  
		dets += det;
	}
	else
	{	 	
		// no rain - either it is snowing or it is "dry"  
		det = 0.0;
	}

    return 0;
}  


int attach(float affix,float deltd,float& dets)
{
	// Simulate attachment or compaction of detached sediment on the  
	// surface. The calculation is done at the start of each day, if  
	// the previous day was dry  
    
	dets *= (1.0 - affix * deltd);

    return 0;
} 

int sosed1(float runoff,float surs,float delt60,float kser,float jser,
		   float kger,float jger,float& dets,float& sosed)
{
    float d1 = 0., d2 = 0., arg = 0., stcap = 0., wssd = 0., scrsd = 0.;
	
	// Remove both detached surface sediment and soil matrix by surface  
	// Flow using method 1  
    if (runoff <= 0.0) goto L30;

		// surface runoff occurs, so sediment and soil matrix  
		// particles may be removed, delt60= delt/60  
		// get argument used in transport equations  
		arg = surs + runoff;	// inches
		
		// calculate capacity for removing detached sediment - units  
		// are tons/acre-ivl  
		d1 =  (arg / delt60);
		d2 =  (jser);
		stcap = delt60 * kser * pow(d1, d2);	// tons/ac/ivl
		
		if (stcap <= dets) goto L10;

			// there is insufficient detached storage, base sediment  
			// removal on that available, wssd is in tons/acre-ivl  
			wssd = dets * runoff / arg;
			goto L20;
L10:
	// there is sufficient detached storage, base sediment  
	// removal on the calculated capacity  
	wssd = stcap * runoff / arg;
L20:
	dets -= wssd;
	
	// calculate scour of matrix soil by surface runoff -  
	// units are tons/acre-ivl  
	d1 =  (arg / delt60);
	d2 =  (jger);
	scrsd = delt60 * kger * pow(d1, d2);
	scrsd = scrsd * runoff / arg;
	
	// total removal by runoff  
	sosed = wssd + scrsd;
	goto L40;
L30:
	// no runoff occurs, so no removal by runoff  
	wssd = 0.0;
    scrsd = 0.0;
    sosed = 0.0;
L40:
    return 0;
}

int BDEXCH(double AVDEPM,double W,double TAU,double TAUCD,double TAUCS,
		   double M,double VOL,double FRCSED,double *SUSP,double *BED,
		   double *DEPSCR)
{
	// Simulate deposition and scour of a cohesive sediment 
	// fraction-silt or clay
    double DEPMAS,EXPNT,SCR,SCRMAS;
 
	if (W>0.0 && TAU<TAUCD && *SUSP>1.0E-30) 
	{
		//deposition will occur
		EXPNT = -W/AVDEPM*(1.0 - TAU/TAUCD);
		DEPMAS= *SUSP*(1.0 - exp(EXPNT));		// mg/l*m3
		*SUSP  = *SUSP - DEPMAS;				// mg/l*m3
		*BED   = *BED + DEPMAS;					// mg/l*m3
	}
	else
	{
		//no deposition- concentrations unchanged
		DEPMAS= 0.0;							// mg/l*m3
	}
 
	if (TAU>TAUCS && M>0.0) 
	{
		if (TAUCS == 0)	TAUCS = 1.0E-15;

		//scour can occur- units are:
		//m- kg/m2.ivl  avdepm- m  scr- mg/l
		SCR= FRCSED*M/AVDEPM*1000.*(TAU/TAUCS - 1.0);	// mg/l

		//check availability of material
		SCRMAS= SCR*VOL;						// mg/l*m3

		if (SCRMAS > *BED) 
		SCRMAS= *BED;							// mg/l*m3

		//update storages
		*SUSP= *SUSP + SCRMAS;					// mg/l*m3
		*BED = *BED - SCRMAS;					// mg/l*m3
	}
	else
	{
		//no scour
		SCRMAS= 0.0;							// mg/l*m3
	}
	
	//calculate net deposition or scour
	*DEPSCR= DEPMAS - SCRMAS;					// mg/l*m3

    return 0;
}
 
int sandld(double isand,double vols,double srovol,double vol,double erovol,
		   double ksand,double avvele,double expsnd,double rom,int sandfg,
		   double db50e,double hrade,double slope,double tw, double wsande,
		   double twide,double db50m,double fsl,double avdepe,double *sand,
		   double *rsand,double *bdsand,double *depscr,double *rosand)
{
	// Simulate behavior of sand/gravel  
	// variables are r4 unless otherwise stated  
    
	int d50err, hrerr, ferror, velerr;
	double gsi;
	
    double sands = *sand;	// mg/l
	double psand = 0.0;
	double scour = 0.0;
	double prosnd = 0.0;
	double pscour = 0.0;
	
    if (vol > 0.0)	// m3
	{
		// rchres contains water 
		if (rom > 0.0 && avdepe > 0.17)
		{
			// there is outflow from the rchres- perform advection 
			// calculate potential value of sand 
			switch (sandfg)
			{
			case 1:  
				// toffaleti equation 
				toffal(&avvele, &db50e, &hrade, &slope, &tw, &wsande, &gsi);
				
				// convert potential sand transport rate to a concentration 
				// in mg/l 
				psand = gsi * twide * 10.5 / rom;
				break;
			case 2:  
				// colby equation 
 				colby(&avvele, &db50m, &hrade, &fsl, &tw, &gsi, &ferror, 
 					  &d50err, &hrerr, &velerr);
				
				if (ferror == 1)
				{
					CString WarnMessage;
					WarnMessage.Format("fatal error ocurred in colby method\n - one or more variables went outside valid range\n - switch to toffaleti method\n");
					TRACE(WarnMessage);
					
					// switch to toffaleti method 
					toffal(&avvele, &db50e, &hrade, &slope, &tw, &wsande, &gsi);
				}
				
				// convert potential sand transport rate to conc in mg/l 
				psand = gsi * twide * 10.5 / rom;
				break;
			case 3:  
				// input power function 
				psand = ksand * pow(avvele, expsnd); // mg/l
				break;
			default:
				break;
			}
			
			// calculate potential outflow of sand during ivl (mg/l * m3 = g)
			prosnd = sands * srovol + psand * erovol;
			
			// calculate potential scour from, or to deposition, bed storage
			// scour is expressed as qty.vol/l.ivl (g)
			pscour = vol * psand - vols * sands + prosnd - isand;
			
			if (pscour < *bdsand)
			{
				// potential scour is satisfied by bed storage; new conc. 
				// of sandload is potential conc. 
				scour = pscour;
				*sand = psand;
				*rsand = *sand * vol;
				*bdsand -= scour;
			}
			else
			{
				// potential scour cannot be satisfied; all of the 
				// available bed storage is scoured 
				scour = *bdsand;
				*bdsand = 0.0;
				
				// calculate new conc. of suspended sandload 
				*sand = (isand + scour + sands * (vols - srovol)) / (vol + erovol);
				
				// calculate new storage of suspended sandload 
				*rsand = *sand * vol;
			}
			
			// calculate total amount of sand leaving rchres during ivl 
			*rosand = srovol * sands + erovol * *sand;
		} 
		else 
		{
			// no outflow (still water) or water depth less than two inches 
			*sand = 0.0;
			*rsand = 0.0;
			scour = -(isand) - sands * vols;
			*bdsand -= scour;
			*rosand = 0.0;
		}
	} 
	else 
	{
		// rchres is dry; set sand equal to an undefined number 
		*sand = 0.0;
		*rsand = 0.0;
		
		// calculate total amount of sand settling out during interval; 
		// this is equal to sand inflow + sand initially present 
		scour = -(isand) - sands * vols;
		
		// update bed storage 
		*bdsand -= scour;
		
		// assume zero outflow of sand 
		*rosand = 0.0;
    }
	
	// calculate depth of bed scour or deposition; positive for 
	// deposition 
    *depscr = -scour;
	
    return 0;
}  


int toffal(double *v,double *fdiam,double *fhrad,double *slope,
		   double *tempr,double *vset,double *gsi)
{
	// This subroutine uses toffaleti's method to calculate the capacity  
	// of the flow to transport sand.  
	//  called by: sandld  

	// V     - average velocity of flow (ft/s)  
	// FDIAM - median bed sediment diameter (ft)  
	// FHRAD - hydraulic radius (ft) 
	// SLOPE - energy or river bed slope  
	// TEMPR - water temperature (deg c)  
	// VSET  - settling velocity (ft/s)  
	// GSI   - total capacity of the rchres (tons/day.ft)  
    
	double d1, d2, d3, d4, d5, d6, d7, d8;
	
    double oczl, oczm, oczu, tmpr, zinv, afunc, ustar, k4, p1, k4func,
		ac, d65, cz, zi, zm, tt, zn, zo, zp, zq, c2d, rprime, zo2, fd11,
		fd25, cli, cmi, gsb, gsl, cnv, gsm, gsu, vis, ack4;
	
	
	
	// Convert water temp from degrees c to degrees f  
    tmpr = *tempr * 1.8 + 32.0;
	
	
	// For water temperatures greater than 32f and less than 100f  
	// The kinematic viscosity can be written as the following:  
    d1 =  tmpr;
    vis = pow(d1, -0.864) * 4.106e-4;
	
	// Assuming the d50 grain size is approximately equal to the  
	// Geometric mean grain size and sigma-g = 1.5, the d65 grain  
	// Size can be determined as 1.17*d50.  
	
    d65 = *fdiam * 1.17;
    cnv = tmpr * 4.8e-4 + 0.1198;
    cz = 260.67 - tmpr * 0.667;
    tt = (tmpr * 9e-5 + 0.051) * 1.1;
    zi = *vset * *v / (cz * *fhrad * *slope);
    if (zi < cnv) 
		zi = cnv * 1.5;
	
	// The manning-strickler equation is used here to  
	// Determine the hydraulic radius component due to
	// Grain roughness (r').  taken from the 1975 asce
	// "sedimentation engineering",pg. 128.  
	
    d1 =  (*v);
    d2 =  d65;
    d3 =  (*slope);
    rprime = pow(d1, 1.5) * pow(d2, 0.25) / pow(d3, 0.75) * 0.00349;
    d1 =  (rprime * *slope * 32.2);
    ustar = pow(d1, 0.5);
    d1 =  (vis * 1e5);
    afunc = pow(d1, 0.333) / (ustar * 10.0);
    if (afunc <= 0.5) 
	{
		d1 =  (afunc / 4.89);
		ac = pow(d1, -1.45);
    } 
	else if (afunc <= 0.66) 
	{
		d1 =  (afunc / 0.0036);
		ac = pow(d1, 0.67);
    } 
	else if (afunc <= 0.72) 
	{
		d1 =  (afunc / 0.29);
		ac = pow(d1, 4.17);
    } 
	else if (afunc <= 1.25) 
	{
		ac = 48.0;
    } 
	else if (afunc > 1.25) 
	{
		d1 =  (afunc / 0.304);
		ac = pow(d1, 2.74);
    }
	
    k4func = afunc * *slope * d65 * 1e5;
    if (k4func <= 0.24) 
	{
		k4 = 1.0;
    } 
	else if (k4func <= 0.35) 
	{
		d1 =  k4func;
		k4 = pow(d1, 1.1) * 4.81;
    } 
	else if (k4func > 0.35) 
	{
		d1 =  k4func;
		k4 = pow(d1, -1.05) * 0.49;
    }
	
    ack4 = ac * k4;
    if (ack4 - 16.0 < 0.0) 
	{
		ack4 = 16.0;
		k4 = 16.0 / ac;
    }
    oczu = cnv + 1.0 - zi * 1.5;
    oczm = cnv + 1.0 - zi;
    oczl = cnv + 1.0 - zi * 0.756;
    zinv = cnv - zi * 0.758;
    zm = -zinv;
    zn = zinv + 1.0;
    zo = zi * -0.736;
    zp = zi * 0.244;
    zq = zi * 0.5;
	
	// Cli has been multiplied by 1.0e30 to keep it from  
	// Exceeding the computer overflow limit  
	
    d1 =  (*v);
    d2 =  (*fhrad);
    d3 =  zm;
    d4 =  (tt * ac * k4 * *fdiam);
    d5 =  (*fhrad / 11.24);
    d6 =  zn;
    d7 =  (*fdiam * 2.0);
    d8 =  oczl;
    cli = oczl * 5.6e22 * pow(d1, 2.333) / pow(d2, d3)
		/ pow(d4, 1.667) / (cnv + 1.0) / (pow(d5,d6) - pow(d7, d8));
	
    zo2 = zo / 2.0;
    d1 =  (*fdiam * 2.0 / *fhrad);
    d2 =  zo2;
    p1 = pow(d1, d2);
    c2d = cli * p1;
    c2d = c2d * p1 / 1e30;
	
	// Check to see if the calculated value is reasonable  
	// (< 100.0), and adjust it if it is not.  
	
    if (c2d > 100.0) 
		cli = cli * 100.0 / c2d;
	
	
	// Cmi has been multiplied by 1.0e30 to keep it from  
	// Exceeding the computer overflow limit  
	
    d1 =  (*fhrad);
    d2 =  zm;
    cmi = cli * 43.2 * (cnv + 1.0) * *v * pow(d1, d2);
	
	// Calculate transport capacity of the upper layer  
	
    fd11 = *fhrad / 11.24;
    fd25 = *fhrad / 2.5;
    d1 =  fd11;
    d2 =  zp;
    d3 =  fd25;
    d4 =  zq;
    d5 =  (*fhrad);
    d6 =  oczu;
    d7 =  fd25;
    d8 =  oczu;
    gsu = cmi * pow(d1, d2) * pow(d3, d4) * (pow(d5,d6) - pow(d7, d8)) / (oczu * 1e30);
	
	// Calculate the capacity of the middle layer  
	
    d1 =  fd11;
    d2 =  zp;
    d3 =  fd25;
    d4 =  oczm;
    d5 =  fd11;
    d6 =  oczm;
    gsm = cmi * pow(d1, d2) * (pow(d3, d4) - pow(d5, d6)) / (oczm * 1e30);
	
	// Calculate the capacity of the lower layer  
	
    d1 =  fd11;
    d2 =  zn;
    d3 =  (*fdiam * 2.0);
    d4 =  oczl;
    gsl = cmi * (pow(d1, d2) - pow(d3, d4)) / (oczl * 1e30);
	
	// Calculate the capacity of the bed layer  
	
    d1 =  (*fdiam * 2.0);
    d2 =  zn;
    gsb = cmi * pow(d1, d2) / 1e30;
	
	// Total capacity of the rchres (gsi has units of tons/day/ft)  
	
    *gsi = gsu + gsm + gsl + gsb;
    if (*gsi <= 0.0)
		*gsi = 0.0;
	
    return 0;
}  


//     This subroutine uses colby's method to calculate the capacity of  
//     the flow to transport sand.  
//      called by: sandld  
//     the colby method has the following units and applicable ranges of  
//     variables.  
//        average velocity.............v.......fps.........1-10 fps  
//        hydraulic radius.............fhrad...ft..........1-100 ft  
//        median bed material size.....db50....mm..........0.1-0.8 mm  
//        temperature..................tmpr....deg f.......32-100 deg.  
//        fine sediment concentration..fsl.....mg/liter....0-200000 ppm  
//        total sediment load..........gsi.....ton/day.ft..  

//	   V      - average velocity (ft/s)  
//	   DB50M  - median bed sediment diameter (mm)  
//	   FHRAD  - hydraulic radius   (ft)  
//     FSL    - total concentration of suspended silt and clay  
//              (fine sediment) (mg/l)  
//     MESSU  - ftn unit no. to be used for printout of messages  
//     TEMPR  - water temperature (degrees c)  
//     GSI    - total sand transport (tons/day.ft width)  
//     FERROR - fatal error flag (if on, switch to toffaleti method)  
//     D50ERR - ???  
//     HRERR  - ???  
//     VELERR - ???  

int colby(double *v,double *db50m,double *fhrad,double *fsl,double *tempr, 
		  double *gsi,int *ferror,int *d50err,int *hrerr,int *velerr)
{
    // Initialized data 
	
	double g[192] ={1.0, 0.3, 0.06, 0.0, 3.0, 3.3, 2.5, 2.0, 5.4, 9.0,
					10.0, 20.0, 11.0, 26.0,	50.0, 150.0, 17.0, 49.0, 130.0, 500.0,
					29.0, 101.0, 400.0, 1350.0, 44.0, 160.0, 700.0, 2500.0, 60.0, 220.0,
					1e3, 4400.0, 0.38, 0.06, 0.0, 0.0, 1.6, 1.2, 0.65, 0.1,
					3.7, 5.0, 4.0, 3.0, 10.0, 18.0, 30.0, 52.0, 17.0, 40.0,
					80.0, 160.0, 36.0, 95.0, 230.0, 650.0, 60.0, 150.0, 415.0, 1200.0,
					81.0, 215.0, 620.0, 1500.0, 0.14, 0.0, 0.0, 0.0, 1.0, 0.6,
					0.15, 0.0, 3.3, 3.0, 1.7, 0.5, 11.0, 15.0, 17.0, 14.0,
					20.0, 35.0, 49.0, 70.0, 44.0, 85.0, 150.0, 250.0, 71.0, 145.0,
					290.0, 500.0, 100.0, 202.0, 400.0, 700.0, 0.0, 0.0, 0.0, 0.0,
					0.7, 0.3, 0.06, 0.0, 2.9, 2.3, 1.0, 0.06, 11.5, 13.0,
					12.0, 7.0, 22.0, 31.0, 40.0, 50.0, 47.0, 84.0, 135.0, 210.0,
					75.0, 140.0, 240.0, 410.0, 106.0, 190.0, 350.0, 630.0, 0.0, 0.0,
					0.0, 0.0, 0.44, 0.06, 0.0, 0.0, 2.8, 1.8, 0.6, 0.0,
					12.0, 12.5, 10.0, 4.5, 24.0, 30.0, 35.0, 37.0, 52.0, 78.0,
					120.0, 190.0, 83.0, 180.0, 215.0, 380.0, 120.0, 190.0, 305.0, 550.0,
					0.0, 0.0, 0.0, 0.0, 0.3, 0.0, 0.0, 0.0, 2.9, 1.4,
					0.3, 0.0, 14.0, 11.0, 7.7, 3.0, 27.0, 29.0, 30.0, 30.0,
					57.0, 75.0, 110.0, 170.0, 90.0, 140.0, 200.0, 330.0, 135.0, 190.0,
					290.0, 520.0};
    double d50g[6]={0.1, 0.2, 0.3, 0.4, 0.6, 0.8};
    double temp[7]={32.0, 40.0, 50.0, 70.0, 80.0, 90.0, 100.0};
    double f[50] = {1.0, 1.1, 1.6, 2.6, 4.2, 1.0, 1.1, 1.65, 2.75, 4.9,
					1.0, 1.1, 1.7, 3.0, 5.5, 1.0, 1.12, 1.9, 3.6, 7.0,
					1.0, 1.17, 2.05, 4.3, 8.7, 1.0, 1.2, 2.3, 5.5, 11.2,
					1.0, 1.22, 2.75, 8.0, 22.0, 1.0, 1.25, 3.0, 9.6, 29.0,
					1.0, 1.3, 3.5, 12.0, 43.0, 1.0, 1.4, 4.9, 22.0, 120.0};
    double t[28] = {1.2, 1.15, 1.1, 0.96, 0.9, 0.85, 0.82, 1.35, 1.25, 1.12,
					0.92, 0.86, 0.8, 0.75, 1.6, 1.4, 1.2, 0.89, 0.8, 0.72,
					0.66, 2.0, 1.65, 1.3, 0.85, 0.72, 0.63, 0.55};
    double df[10] ={0.1, 0.2, 0.3, 0.6, 1.0, 2.0, 6.0, 10.0, 20.0, 100.0};
    double cf[5] = {0.0, 1e4, 5e4, 1e5, 1.5e5};
    double p[11] = {0.6, 0.9, 1.0, 1.0, 0.83, 0.6, 0.4, 0.25, 0.15, 0.09, 0.05};
    double dp[11] ={0.1, 0.15, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0};
    double dg[4] = {0.1, 1.0, 10.0, 100.0};
    double vg[8] = {1.0, 1.5, 2.0, 3.0, 4.0, 6.0, 8.0, 10.0};
    double r1, r2, r3;
    double d1;
	
    double gtuc, tmpr;
    int i, j, k;
    double x[4];	// was [2][2] ;
    int i1, j1, j3, k1;
    double p1, p2;
    int ii[2], jj[2], kk[2];
    double xa[2], xd, xf[4]	// was [2][2] 
		, xg[2], xr, xt[4] // was [2][2] 
		, xv, xx[2], yy[2], zz[2];
		int id1, id2, if1, if2, ip1, ip2, it1, it2, iv1, iv2;
    double xn1, xn2, xn3, xn4, db50, cfd, cff, fff, cft, tcf, xct[2],
		xdx, xdy, xdz, xnt;
    int id501, id502;
	
	
    db50 = *db50m;
    tmpr = *tempr * 1.8 + 32.0;
	
	//      fsl....fine sediment (i.e. cohesive sediment or wash) load 
	//      in mg/liter 
	
    *ferror = 0;
    *d50err = 0;
    *hrerr = 0;
    *velerr = 0;
	
    if (db50 >= d50g[0] && db50 <= d50g[5])
	{
		goto L10;
    }
    *ferror = 1;
    *d50err = 1;
L10:
    if (*fhrad >= dg[0] && *fhrad <= dg[3])
	{
		goto L20;
    }
    *ferror = 1;
    *hrerr = 1;
L20:
    if (*v >= vg[0] && *v <= vg[7])
	{
		goto L30;
    }
    *ferror = 1;
    *velerr = 1;
L30:
    if (*ferror != 0)
		goto L400;

    if (tmpr >= 32.0)
		goto L40;

    tmpr = 32.0;
L40:
    if (tmpr <= 100.0)
		goto L45;

    tmpr = 100.0;
L45:
    id1 = 0;
    id2 = 0;
    for (i = 1; i <= 3; ++i)
	{
		if (*fhrad < dg[i - 1] || *fhrad > dg[i])
		{
			goto L50;
		}
		id1 = i;
		id2 = i + 1;
		goto L70;
L50:
		// L60:
		;
    }
L70:
    iv1 = 0;
    iv2 = 0;
    for (i = 1; i <= 7; ++i)
	{
		if (*v < vg[i - 1] || *v > vg[i])
		{
			goto L80;
		}
		iv1 = i;
		iv2 = i + 1;
		goto L100;
L80:
		// L90: 
		;
    }
L100:
    id501 = 0;
    id502 = 0;
    for (i = 1; i <= 5; ++i)
	{
		if (db50 < d50g[i - 1] || db50 > d50g[i])
		{
			goto L110;
		}
		id501 = i;
		id502 = i + 1;
		goto L130;
L110:
		// L120: 
		;
    }
L130:
    ii[0] = id1;
    ii[1] = id2;
    jj[0] = iv1;
    jj[1] = iv2;
    kk[0] = id501;
    kk[1] = id502;
    for (i = 1; i <= 2; ++i)
	{
		i1 = ii[i - 1];
		xx[i - 1] = log10(dg[i1 - 1]);
		for (j = 1; j <= 2; ++j) {
			j1 = jj[j - 1];
			yy[j - 1] = log10(vg[j1 - 1]);
			for (k = 1; k <= 2; ++k) {
				k1 = kk[k - 1];
				zz[k - 1] = log10(d50g[k1 - 1]);
				if (g[i1+(j1+((k1<<3)<<2))-37] > 0.0) 
				{
					goto L160;
				}
				for (j3 = j1; j3 <= 7; ++j3) {
					if (g[i1 + (j3 + ((k1 << 3) << 2)) - 37] > 0.0) {
						goto L150;
					}
					// L140: 
				}
L150:
				r1 = vg[j1 - 1] / vg[j3 - 1];
				r2 = g[i1 + (j3 + 1 + ((k1 << 3) << 2)) - 37] / g[i1 + (j3 + ((
					k1 << 3) << 2)) - 37];
				r3 = vg[j3] / vg[j3 - 1];
				x[j + (k << 1) - 3] = log10(g[i1 + (j3 + ((k1 << 3) << 2)) -
					37]) + log10(r1) * log10(r2) / log10(r3);
				goto L170;
L160:
				x[j + (k << 1) - 3] = log10(g[i1 + (j1 + ((k1 << 3) << 2)) -
					37]);
L170:
				//L180: 
				;
			}
			// L190: 
		}
		xd = log10(db50) - zz[0];
		xn1 = x[2] - x[0];
		xn2 = x[3] - x[1];
		xdz = zz[1] - zz[0];
		xa[0] = x[0] + xn1 * xd / xdz;
		xa[1] = x[1] + xn2 * xd / xdz;
		xv = log10(*v) - yy[0];
		xn3 = xa[1] - xa[0];
		xdy = yy[1] - yy[0];
		xg[i - 1] = xa[0] + xn3 * xv / xdy;
		// L200: 
    }
    xn4 = xg[1] - xg[0];
    xr = log10(*fhrad) - xx[0];
    xdx = xx[1] - xx[0];
    gtuc = xg[0] + xn4 * xr / xdx;
    d1 = gtuc;
    gtuc = pow(10.0, d1);
	
	//       gtuc is uncorrected gt in lb/sec/ft 
	
	//       next apply fine sediment load and temperature /
	//                                             corrections 
	
	//       if (tmpr .ne. 60.) go to 210 
	r1 = tmpr - 60.0;
    if (fabs(r1) > 1e-5)
		goto L210;

    cft = 1.0;
    goto L250;
L210:
    it1 = 0;
    it2 = 0;
    for (i = 1; i <= 6; ++i) {
		if (tmpr < temp[i - 1] || tmpr > temp[i]) {
			goto L220;
		}
		it1 = i;
		it2 = i + 1;
		goto L240;
L220:
		// L230: 
		;
    }
L240:
    xt[0] = log10(t[it1 + id1 * 7 - 8]);
    xt[1] = log10(t[it2 + id1 * 7 - 8]);
    xt[2] = log10(t[it1 + id2 * 7 - 8]);
    xt[3] = log10(t[it2 + id2 * 7 - 8]);
    r1 = tmpr / temp[it1 - 1];
    r2 = temp[it2 - 1] / temp[it1 - 1];
    xnt = log10(r1) / log10(r2);
    xct[0] = xt[0] + xnt * (xt[1] - xt[0]);
    xct[1] = xt[2] + xnt * (xt[3] - xt[2]);
    cft = xct[0] + (xct[1] - xct[0]) * xr / xdx;
    d1 =  cft;
    cft = pow(10.0, d1);
L250:
	
	//        fine sediment load correction 
	
    if (*fsl > 10.0)
		goto L260;

    cff = 1.0;
    goto L350;
L260:
    id1 = 0;
    id2 = 0;
    for (i = 1; i <= 9; ++i) {
		if (*fhrad < df[i - 1] || *fhrad > df[i]) {
			goto L270;
		}
		id1 = i;
		id2 = i + 1;
		goto L290;
L270:
		// L280: 
		;
    }
L290:
    if (*fsl <= 1e5)
		goto L300;

    if1 = 4;
    if2 = 5;
    goto L340;
L300:
    if1 = 0;
    if2 = 0;
    for (i = 1; i <= 4; ++i) {
		if (*fsl < cf[i - 1] || *fsl > cf[i]) {
			goto L310;
		}
		if1 = i;
		if2 = i + 1;
		goto L330;
L310:
		//L320: 
		;
    }
L330:
L340:
    xf[0] = log10(f[if1 + id1 * 5 - 6]);
    xf[3] = log10(f[if2 + id2 * 5 - 6]);
    xf[2] = log10(f[if1 + id2 * 5 - 6]);
    xf[1] = log10(f[if2 + id1 * 5 - 6]);
    xnt = (*fsl - cf[if1 - 1]) / (cf[if2 - 1] - cf[if1 - 1]);
    xct[0] = xf[0] + xnt * (xf[1] - xf[0]);
    xct[1] = xf[2] + xnt * (xf[3] - xf[2]);
    r1 = *fhrad / df[id1 - 1];
    r2 = df[id2 - 1] / df[id1 - 1];
    xnt = log10(r1) / log10(r2);
    cff = xct[0] + xnt * (xct[1] - xct[0]);
    d1 =  cff;
    cff = pow(10.0, d1);
L350:
    tcf = cft * cff - 1.0;
    cfd = 1.0;
    if (db50 >= 0.2 && db50 <= 0.3) {
		goto L390;
    }
    ip1 = 0;
    ip2 = 0;
    for (i = 1; i <= 10; ++i) {
		if (db50 < dp[i - 1] || db50 > dp[i]) {
			goto L360;
		}
		ip1 = i;
		ip2 = i + 1;
		goto L380;
L360:
		// L370: 
		;
    }
L380:
    p2 = log10(p[ip2 - 1]);
    p1 = log10(p[ip1 - 1]);
    r1 = db50 / dp[ip1 - 1];
    r2 = dp[ip2 - 1] / dp[ip1 - 1];
    xnt = log10(r1) / log10(r2);
    cfd = p1 + xnt * (p2 - p1);
    d1 =  cfd;
    cfd = pow(10.0, d1);
L390:
    fff = cfd * tcf;
    fff += 1.0;
    *gsi = fff * gtuc;
	
L400:
    return 0;
} 
