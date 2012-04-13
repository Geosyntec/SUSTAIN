// BMPData.cpp : implementation of the CBMPData class
//
#define	MAXLINE 1024

#include "stdafx.h"
#include <direct.h>
#include <stdlib.h>
#include <stdio.h>
#include <math.h>
#include "LandUse.h"
#include "BMPSite.h"
#include "SiteLandUse.h"
#include "SitePointSource.h"
#include "BMPData.h"
#include "BMPRunner.h"
#include "StringToken.h"
#include "Global.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBMPData construction/destruction

CBMPData::CBMPData()
{
	nStrategy = STRATEGY_SCATTER_SEARCH;
	strInputDir = "";
	strOutputDir = "";
	strMixLUFileName = "";
	strPreLUFileName = "";
	strError = "";			// for reporting error occurring while loading data or saving result
	nLandSimulation = 0;	// Land Simulation Option (0-External,1-Internal)
	nWeatherFile = 0;		// Weather File Option (0-No File,1-Precip File)
	m_nNum = 0;
	nETflag = 0;
	m_nSedQualFlag = 0;
	lfLatitude = 50.0;
	for (int i=0; i<12; i++)
		lfmonET[i] = 0.0;
	m_pDataClimate = NULL;
	strClimateFileName = "";
	nBMPTimeStep = 60;		// BMP Simulation Time Step (1 - 60 minutes)
	nOutputTimeStep = 1;	// 0 for daily, 1 for hourly
	nBIORETENTION = 0;		
	nWETPOND = 0;			
	nCISTERN = 0;			
	nDRYPOND = 0;			
	nINFILTRATIONTRENCH = 0;
	nGREENROOF = 0;			
	nPOROUSPAVEMENT = 0;	
	nRAINBARREL = 0;
	nREGULATOR = 0;
	nSWALE = 0;	
	nBMPtype = 0;
	nBMPA = 0;
	nBMPB = 0;
	nBMPC = 0;				// conduit (01-2005)
	nBMPD = 0;				// bufferstrip (06-2007)
	nN = 0;
	nMAXITER = 0;
	nNPOL = 0;
	nIELOUT = 0;
	nKPG = 0;
	lfTHETAW = 0.0;
	lfCR = 0.0; 
	nAdjVariable = 0;
	nEvalFactor = 0;
	nRunOption = OPTION_NO_OPTIMIZATION; // 0 -- No Optimization, 1 -- Minimize Cost, 3 -- Maximize Control, 2 -- Generate Trade-off Curve
	nSolution = 1;
	nTargetBreak = 1;
	lfCostLimit = 0;
	lfStopDelta = 0;
	lfMaxRunTime = 0;
	startDate = COleDateTime(1890,1,1,0,0,0);
	endDate = COleDateTime(1890,1,1,0,0,0);
	polmultiplier = NULL;
	nSedflag = NULL;
	m_pPollutant = NULL;
	m_pBMPcost = NULL;
	nPollutant = 0;
	nNWQ = 0;

	//optional weather data
	lRecords = 0;
	lStartIndex = 0;
	lEndIndex = 0;
	nWetPeriod = 0;
	nWetInt = 0;
	lfWetDays = 0;
	pWEATHERDATA = NULL;
	pWetPeriod = NULL;
}

CBMPData::~CBMPData()
{
	routeList.RemoveAll();

	while(!luList.IsEmpty())
		delete (CLandUse *) luList.RemoveTail();
	while(!siteluList.IsEmpty())
		delete (CSiteLandUse *) siteluList.RemoveTail();
	while(!bmpsiteList.IsEmpty())
		delete (CBMPSite *) bmpsiteList.RemoveTail();
	while (!bmpGroupList.IsEmpty())
		delete (BMP_GROUP*) bmpGroupList.RemoveTail();
	while(!sitepsList.IsEmpty())
		delete (CSitePointSource *) sitepsList.RemoveTail();

	if (nSedflag != NULL)
		delete []nSedflag;
	if (polmultiplier != NULL)
		delete []polmultiplier;
	if (m_pPollutant != NULL)
		delete []m_pPollutant;
	if (m_pBMPcost != NULL)
		delete []m_pBMPcost;
	if (m_pDataClimate != NULL)
		delete []m_pDataClimate;
	
	//optional weather data
	if(pWEATHERDATA!= NULL)	
		delete[]pWEATHERDATA;
	if(pWetPeriod != NULL)
		delete[]pWetPeriod;
}

/////////////////////////////////////////////////////////////////////////////
// CBMPData member functions

CBMPSite* CBMPData::FindBMPSite(const CString& strID)
{
	POSITION pos;
	CBMPSite* pBMPSite;

	//get the head position
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (strID.CompareNoCase(pBMPSite->m_strID) == 0)
			return pBMPSite;
	}
	return NULL;
}

CLandUse* CBMPData::FindLandUse(int nLuID)
{
	POSITION pos = luList.FindIndex(nLuID-1);
	if (pos != NULL)
		return (CLandUse*) luList.GetAt(pos);
	else
		return NULL;
}

bool CBMPData::ReadInputFile(CString strFileName)
{
	FILE *fpin = NULL;
	char strLine[MAXLINE];
	bool retVal = true;
	int  nSection;
	
	// open the file for reading
	fpin = fopen (strFileName, "rt");
	if(fpin == NULL)
	{
		strError = "Cannot open file " + strFileName + " for reading";
		return false;
	}
	
	// skip the line start with a 'C'
	while(!feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		CString str(strLine);
		str.TrimLeft();

		if(str.GetLength() == 0)
			continue;

		if(str[0] == 'C' || str[0] == 'c')
		{
			// if 'C' is followed by a number, read this section
			if(str.GetLength() >= 2)
				if(str[1] >= '0' && str[1] <= '9')
				{
					// get section number
					CStringToken strToken(str);
					CString strNum = strToken.NextToken();
					strNum = strNum.Mid(1);
					nSection = atoi((LPCSTR)strNum);

					if(!ReadFileSection(fpin, nSection))
					{
						retVal = false;
						break;
					}
				}
		}
	}

	fclose(fpin);
	return retVal;
}

bool CBMPData::ReadFileSection(FILE *fpin, int nSection)
{
	int i, j, nPollutants, nIndex;
	CString str;
	int ncount = 0; // count number of transects
	CString id;		// transect ID name

	SkipCommentLine(fpin);

	switch(nSection)
	{
		case 700:
			{
				//read input directory
				ReadDataLine(fpin, str);
				CStringToken strToken(str);
				
				//read the land simulation control
				nLandSimulation = atoi((LPCSTR)strToken.NextToken());
				if (nLandSimulation != 1) nLandSimulation = 0;
			
				//read the land output directory path
				strInputDir = strToken.NextToken();
				if(strInputDir.GetAt(strInputDir.GetLength()-1)	!= '\\')
					strInputDir += "\\";

				//read time series data file path for mixed landuse output
				if (nLandSimulation == 1)
				{
					strMixLUFileName = strToken.NextToken();
					// make sure full path for time series data file
					if(strMixLUFileName.Find(":", 0) == -1)
						strMixLUFileName = strInputDir + strMixLUFileName;
					
					//read time series data file path for predeveloped landuse output
					strPreLUFileName = strToken.NextToken();
					// make sure full path for time series data file
					if(strPreLUFileName.Find(":", 0) == -1)
						strPreLUFileName = strInputDir + strPreLUFileName;
				}

				// read start date
				int nYear, nMonth, nDay;
				ReadDataLine(fpin, str);
				sscanf(str, "%d %d %d", &nYear, &nMonth, &nDay);
				startDate.SetTime(0, 0, 0);
				startDate.SetDate(nYear, nMonth, nDay);

				// read end date
				ReadDataLine(fpin, str);
				sscanf(str, "%d %d %d", &nYear, &nMonth, &nDay);
				endDate.SetTime(0, 0, 0);
				endDate.SetDate(nYear, nMonth, nDay);

				//read bmp simulation time step, output time step mode, 
				ReadDataLine(fpin, str);
				CStringToken strToken1(str);
				nBMPTimeStep = atoi((LPCSTR)strToken1.NextToken());
				nOutputTimeStep = atoi((LPCSTR)strToken1.NextToken());

				if (nBMPTimeStep < 1) nBMPTimeStep = 1;
				if (nBMPTimeStep > 60) nBMPTimeStep = 60;
				if (nOutputTimeStep != 1) nOutputTimeStep = 0;

				//read the output directory path
				strOutputDir = strToken1.NextToken();
				if(strOutputDir.GetAt(strOutputDir.GetLength()-1)	!= '\\')
					strOutputDir += "\\";

				// CREATE  the output directory if it is necessary
				TRY
				{
					_mkdir(LPCSTR(strOutputDir));
				}
				CATCH_ALL(e)
				{
				}
				END_CATCH_ALL

				//read the ET controls
				ReadDataLine(fpin, str);
				CStringToken strToken2(str);
				nETflag = atoi((LPCSTR)strToken2.NextToken());

				if (nETflag > 0)
				{
					//read the climate file path
					strClimateFileName = strToken2.NextToken();
					// make sure full path for time series data file
					if(strClimateFileName.Find(":", 0) == -1)
						strClimateFileName = strInputDir + strClimateFileName;

					// number of parameters reading from the climate file
					m_nNum = 2;	//TMAX and TMIN

					if (nETflag == 1)
						m_nNum += 1;	// also read EVAP
					else if (nETflag == 2)	
						lfLatitude = atof((LPCSTR)strToken2.NextToken());
				}
				
				//read the monthly ET values
				ReadDataLine(fpin, str);
				CStringToken strToken3(str);
				for (i=0; i<12; i++)
					lfmonET[i] = atof((LPCSTR)strToken3.NextToken());

				//read the weather file control for calculating the wet period (this option is not provided on the interface)
				if (strToken3.HasMoreTokens())
				{
					nWeatherFile = atoi((LPCSTR)strToken3.NextToken());
					if (nWeatherFile != 1) nWeatherFile = 0;
				}
			}
			break;
		case 705:
			{
				if (m_pPollutant != NULL)
				{
					delete []m_pPollutant;
					m_pPollutant = NULL; 
				}

				if (polmultiplier != NULL)
				{
					delete []polmultiplier;
					polmultiplier = NULL; 
				}
				
				//count the number of pollutant
				long nPos = ftell(fpin);
				char strLine[MAXLINE];
				while (fgets (strLine, MAXLINE, fpin) != NULL)
				{
					CString str(strLine);
					CStringToken strToken(str);
					CString str0 = strToken.NextToken();
					if(str0[0] != 'C' && str0[0] != 'c')
						++nPollutant;
					else
					{
						fseek (fpin, nPos, SEEK_SET);
						break;
					}
					memset (strLine, 0, MAXLINE);
				}

				//assign memory if there is any pollutant
				if (nPollutant > 0)
				{
					polmultiplier = new double[nPollutant];
					m_pPollutant = new POLLUTANT[nPollutant];
				}

				nIndex = 0;
				while (ReadDataLine(fpin, str))
				{
					CStringToken strToken(str);
					int	nID = atoi((LPCSTR)strToken.NextToken());
					CString	sName = strToken.NextToken();
					double lfMult = atof((LPCSTR)strToken.NextToken());
					int	nSedfg = atoi((LPCSTR)strToken.NextToken());
					int	nSedQual = atoi((LPCSTR)strToken.NextToken());
					double lfsand_qfr = atof((LPCSTR)strToken.NextToken());
					double lfsilt_qfr = atof((LPCSTR)strToken.NextToken());
					double lfclay_qfr = atof((LPCSTR)strToken.NextToken());

					//check if the qual is sediment associated
					if (nSedQual == 1)
						m_nSedQualFlag = 1;

					m_pPollutant[nIndex].m_nID = nID;
					m_pPollutant[nIndex].m_sName = sName;
					m_pPollutant[nIndex].m_lfMult = lfMult;
					m_pPollutant[nIndex].m_nSedfg = nSedfg;
					m_pPollutant[nIndex].m_nSedQual = nSedQual;
					m_pPollutant[nIndex].m_lfsand_qfr = lfsand_qfr;
					m_pPollutant[nIndex].m_lfsilt_qfr = lfsilt_qfr;
					m_pPollutant[nIndex].m_lfclay_qfr = lfclay_qfr;
					polmultiplier[nIndex] = lfMult;
					nIndex++;
				}

				//check for TSS
				int nSAND = 0,nSILT = 0,nCLAY = 0, nTSS = 0;
				nNWQ = nPollutant;
				for (i=0; i<nPollutant; i++)
				{
					if (m_pPollutant[i].m_nSedfg == SAND)
					{
						nSAND++;
					}
					else if (m_pPollutant[i].m_nSedfg == SILT)
					{
						nSILT++;
					}
					else if (m_pPollutant[i].m_nSedfg == CLAY)
					{
						nCLAY++;
					}
					else if (m_pPollutant[i].m_nSedfg == TSS)
					{
						nTSS++;
						nNWQ += 2; //split TSS into sand, silt, and clay
						//break;
					}
				}

				//check if TSS, SAND, SILT, or CLAY is defined more than once
				if (nTSS > 1 )
				{
					strError.Format("Sediment is defined more than once in card 705");
					return false;
				}

				if (nSAND > 1 )
				{
					strError.Format("Sediment type: Sand is defined more than once in card 705");
					return false;
				}

				if (nSILT > 1 )
				{
					strError.Format("Sediment type: Silt is defined more than once in card 705");
					return false;
				}

				if (nCLAY > 1)
				{
					strError.Format("Sediment type: Clay is defined more than once in card 705");
					return false;
				}

				//check if SAND is defined but SILT or CLAY is not defined
				if (nTSS == 1 && (nSAND > 0 || nSILT > 0 || nCLAY > 0))
				{
					strError.Format("Sediment is defined more than once in card 705");
					return false;
				}

				// check if all three classes are required

				//check if SAND is defined but SILT or CLAY is not defined
//				if (nSAND == 1 && (nSILT == 0 || nCLAY == 0))
//				{
//					strError.Format("Sediment type: Sand is defined but Silt or Clay is missing in card 705");
//					return false;
//				}

				//check if SILT is defined but SAND or CLAY is not defined
//				if (nSILT == 1 && (nSAND == 0 || nCLAY == 0))
//				{
//					strError.Format("Sediment type: Silt is defined but Sand or Clay is missing in card 705");
//					return false;
//				}

				//check if CLAY is defined but SILT or SAND is not defined
//				if (nCLAY == 1 && (nSAND == 0 || nSILT == 0))
//				{
//					strError.Format("Sediment type: Clay is defined but Sand or Silt is missing in card 705");
//					return false;
//				}
			}
			break;
		case 710:
			//required if land simulation control is external
			if (nLandSimulation == 0)
			{
				while (ReadDataLine(fpin, str))
				{
					CLandUse *pLU = new CLandUse();
					CStringToken strToken(str);
					// landuse id
					pLU->m_nID = atoi((LPCSTR)strToken.NextToken());
					// landuse type name
					pLU->m_strLanduse = strToken.NextToken();
					// impervious or not
					str = strToken.NextToken();
					pLU->m_nType = (str[0] == '0')?0:1;
					// time series data file path
					pLU->m_strFileName = strToken.NextToken();
					// make sure full path for time series data file
					if(pLU->m_strFileName.Find(":", 0) == -1)
						pLU->m_strFileName = strInputDir + pLU->m_strFileName;
					// sand fraction
					pLU->m_lfsand_fr = atof((LPCSTR)strToken.NextToken());
					// silt fraction
					pLU->m_lfsilt_fr = atof((LPCSTR)strToken.NextToken());
					// clay fraction
					pLU->m_lfclay_fr = atof((LPCSTR)strToken.NextToken());

					pLU->m_nQualNum = nPollutant + 1;	// add flow

					// add the new landuse to the list
					luList.AddTail(pLU);
				}
			}
			break;
		case 715:
			nIndex = 1;
			while (ReadDataLine(fpin, str))
			{
				int bmpClass;
				CString bmpID, bmpName, bmpType;

				CStringToken strToken(str);
				// BMP id
				bmpID = strToken.NextToken();
				// BMP name
				bmpName = strToken.NextToken();
				// BMP type
				bmpType = strToken.NextToken();

				if(bmpType.CompareNoCase("BIORETENTION") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nBIORETENTION++;
				}
				else if(bmpType.CompareNoCase("WETPOND") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nWETPOND++;
				}
				else if(bmpType.CompareNoCase("CISTERN") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nCISTERN++;
				}
				else if(bmpType.CompareNoCase("DRYPOND") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nDRYPOND++;
				}
				else if(bmpType.CompareNoCase("INFILTRATIONTRENCH") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nINFILTRATIONTRENCH++;
				}
				else if(bmpType.CompareNoCase("GREENROOF") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nGREENROOF++;
				}
				else if(bmpType.CompareNoCase("POROUSPAVEMENT") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nPOROUSPAVEMENT++;
				}
				else if(bmpType.CompareNoCase("RAINBARREL") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nRAINBARREL++;
				}
				else if(bmpType.CompareNoCase("REGULATOR") == 0)
				{
					bmpClass = CLASS_A;
					nBMPA++; // increase the count for class A
					nREGULATOR++;
				}
				else if(bmpType.CompareNoCase("SWALE") == 0)
				{
					bmpClass = CLASS_B;
					nBMPB++; // increase the count for class B
					nSWALE++;
				}
				else if(bmpType.CompareNoCase("CONDUIT") == 0)	
				{
					bmpClass = CLASS_C;
					nBMPC++; // increase the count for class C
				}
				else if(bmpType.CompareNoCase("BUFFERSTRIP") == 0)	
				{
					bmpClass = CLASS_D;
					nBMPD++; // increase the count for class D
				}
				else
				{
					bmpClass = CLASS_X;
				}

				// add the new BMP to the list
				CBMPSite *pBMP = new CBMPSite(bmpID, bmpName, bmpType, bmpClass);
				pBMP->m_lfSiteDArea = atof((LPCSTR)strToken.NextToken());
				pBMP->m_lfBMPUnit = atof((LPCSTR)strToken.NextToken());
				pBMP->m_lfDDarea = atof((LPCSTR)strToken.NextToken());

				if (nLandSimulation == 0)
				{
					int nLuID;
					nLuID = atoi((LPCSTR)strToken.NextToken());
					CLandUse *pLU = (CLandUse *) FindLandUse(nLuID);
					if (pLU == NULL)
					{
						strError.Format("Cannot find Landuse type with ID %d", nLuID);
						return false;
					}
					pBMP->m_preLU = pLU;
				}
				
				bmpsiteList.AddTail(pBMP);

				//increment the BMP index
				nIndex++;
			}

			//find the total number of unique bmp types (A or B)

			if (m_pBMPcost != NULL)
			{
				delete []m_pBMPcost;
				m_pBMPcost = NULL; 
			}

			nBMPtype = 0;
			if (nBIORETENTION > 0)
				nBMPtype++;
			if (nWETPOND > 0)
				nBMPtype++;
			if (nCISTERN > 0)
				nBMPtype++;
			if (nDRYPOND > 0)
				nBMPtype++;
			if (nINFILTRATIONTRENCH > 0)
				nBMPtype++;
			if (nGREENROOF > 0)
				nBMPtype++;
			if (nPOROUSPAVEMENT > 0)
				nBMPtype++;
			if (nRAINBARREL > 0)
				nBMPtype++;
			if (nREGULATOR > 0)
				nBMPtype++;
			if (nSWALE > 0)
				nBMPtype++;

			//assign memory here
			if (nBMPtype > 0)
			{
				m_pBMPcost = new BMPCOST[nBMPtype];
				//initialize
				nIndex = 0;
				if (nBIORETENTION > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "BIORETENTION";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nWETPOND > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "WETPOND";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nCISTERN > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "CISTERN";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nDRYPOND > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "DRYPOND";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nINFILTRATIONTRENCH > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "INFILTRATIONTRENCH";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nGREENROOF > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "GREENROOF";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nPOROUSPAVEMENT > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "POROUSPAVEMENT";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nRAINBARREL > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "RAINBARREL";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nREGULATOR > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_A;
					m_pBMPcost[nIndex].m_strBMPType = "REGULATOR";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
				if (nSWALE > 0)
				{
					m_pBMPcost[nIndex].m_nBMPClass = CLASS_B;
					m_pBMPcost[nIndex].m_strBMPType = "SWALE";
					m_pBMPcost[nIndex].m_lfCost = 0.0;
					nIndex++;
				}
			}

			// Create input and output directories for buffer strip 
			if (nBMPD > 0)
			{
				CString strDirPath;
				TRY
				{
					// create VFSMOD directories under <output directory> 
					strDirPath = strOutputDir + "VFSMOD_input";
					_mkdir(LPCSTR(strDirPath));
					strDirPath = strOutputDir + "VFSMOD_output";
					_mkdir(LPCSTR(strDirPath));
				}
				CATCH_ALL(e)
				{
					//ignore if the directory already exists
				}
				END_CATCH_ALL
			}
			break;
		case 720:	//optional card for point source timeseries
			while (ReadDataLine(fpin, str))
			{
				CSitePointSource *pSitePS = new CSitePointSource();
				CStringToken strToken(str);

				pSitePS->m_nID = atoi((LPCSTR)strToken.NextToken());
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				CBMPSite* pBMPSite = FindBMPSite(strID);
				pSitePS->m_pBMPSite = pBMPSite;
				pSitePS->m_lfMult = atof((LPCSTR)strToken.NextToken());
				pSitePS->m_strPSFile = strToken.NextToken();

				// make sure full path for time series data file
				if(pSitePS->m_strPSFile.Find(":", 0) == -1)
					pSitePS->m_strPSFile = strInputDir + pSitePS->m_strPSFile;

				pSitePS->m_lfSand = atof((LPCSTR)strToken.NextToken());
				pSitePS->m_lfSilt = atof((LPCSTR)strToken.NextToken());
				pSitePS->m_lfClay = atof((LPCSTR)strToken.NextToken());

				pSitePS->m_nQualNum = nPollutant + 1;	// add flow

				// add the new point source to the list
				sitepsList.AddTail(pSitePS);
			}
			break;
		case 721:	//optional card for break point info
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);

				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				CBMPSite* pBMPSite = FindBMPSite(strID);
				pBMPSite->m_nBreakPoints = atoi((LPCSTR)strToken.NextToken());
				pBMPSite->m_strCostFile = strToken.NextToken();
				
				// make sure full path for the file
				if(pBMPSite->m_strCostFile.Find(":", 0) == -1)
					pBMPSite->m_strCostFile = strInputDir+pBMPSite->m_strCostFile;

				if (pBMPSite->m_TradeOff != NULL)
				{
					delete []pBMPSite->m_TradeOff;
					pBMPSite->m_TradeOff = NULL;
				}

				//allocate memory for the timeseries files
				if (pBMPSite->m_nBreakPoints > 0)
					pBMPSite->m_TradeOff = new TradeOffCurve[pBMPSite->m_nBreakPoints+3]; //add init (0), predev (-1), and postdev(-2) conditions
			}
			break;
		case 722:	//optional card for break point timeseries
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);

				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				CBMPSite* pBMPSite = FindBMPSite(strID);

				if (pBMPSite->m_TradeOff == NULL)
					return false;
				
				int nBrPtIndex = atoi((LPCSTR)strToken.NextToken());
				if (nBrPtIndex > pBMPSite->m_nBreakPoints)
					return false;

				pBMPSite->m_TradeOff[nBrPtIndex+2].m_nID = nBrPtIndex;
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_lfMult = atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_strBrPtFile = strToken.NextToken();

				// make sure full path for time series data file
				if(pBMPSite->m_TradeOff[nBrPtIndex+2].m_strBrPtFile.Find(":", 0) == -1)
					pBMPSite->m_TradeOff[nBrPtIndex+2].m_strBrPtFile = strInputDir 
					+ pBMPSite->m_TradeOff[nBrPtIndex+2].m_strBrPtFile;

				pBMPSite->m_TradeOff[nBrPtIndex+2].m_lfSand = atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_lfSilt = atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_lfClay = atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_nQualNum = nPollutant + 1;	// add flow
				pBMPSite->m_TradeOff[nBrPtIndex+2].m_lfCost = 0.0;
			}
			break;
		case 725:
			if(nBMPA == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				int nExitType, nWeirType, nORelease, nPeople, nDays;
				double lfBasinWidth, lfBasinLength, lfOrificeHeight, lfOrificeDiameter, lfWeirHeight, lfWeirWidth, lfWeirAngle;
				CString strID;
				sscanf(str, "%s %lf %lf %lf %lf %d %d %d %d %d %lf %lf %lf",
					strID,
					&lfBasinWidth,
					&lfBasinLength,
					&lfOrificeHeight,
					&lfOrificeDiameter,
					&nExitType,
					&nORelease,
					&nPeople,
					&nDays,
					&nWeirType,
					&lfWeirHeight,
					&lfWeirWidth,
					&lfWeirAngle);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_A) 
				{
					strError.Format("BMP site (ID = %s) is not in class A", strID);
					return false;
				}

				BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;
				pBMP->m_lfBasinWidth      = lfBasinWidth;							
				pBMP->m_lfBasinLength     = lfBasinLength;
				pBMP->m_lfOrificeHeight   = lfOrificeHeight;
				pBMP->m_lfOrificeDiameter = lfOrificeDiameter;
				pBMP->m_nExitType         = nExitType;
				pBMP->m_nORelease         = nORelease;
				pBMP->m_nPeople           = nPeople;
				pBMP->m_nDays             = nDays;
				pBMP->m_nWeirType         = nWeirType;
				pBMP->m_lfWeirHeight      = lfWeirHeight;
				pBMP->m_lfWeirWidth       = lfWeirWidth;
				pBMP->m_lfWeirAngle       = lfWeirAngle;
			}
			break;
		case 730:
			if(nBMPA == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				//int nSiteID;
				double lfRelease[24];
				CString strID;
				sscanf(str,"%s %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf",
					strID,
					&lfRelease[0],
					&lfRelease[1],
					&lfRelease[2],
					&lfRelease[3],
					&lfRelease[4],
					&lfRelease[5],
					&lfRelease[6],
					&lfRelease[7],
					&lfRelease[8],
					&lfRelease[9],
					&lfRelease[10],
					&lfRelease[11],
					&lfRelease[12],
					&lfRelease[13],
					&lfRelease[14],
					&lfRelease[15],
					&lfRelease[16],
					&lfRelease[17],
					&lfRelease[18],
					&lfRelease[19],
					&lfRelease[20],
					&lfRelease[21],
					&lfRelease[22],
					&lfRelease[23]);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_A)
				{
					strError.Format("BMP site (ID = %s) is not in class A", strID);
					return false;
				}
				BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;
				for(i=0; i<24; i++)
					pBMP->m_lfRelease[i] = lfRelease[i];
			}
			break;
		case 735:
			if(nBMPB == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				//int nSiteID;
				double lfBasinWidth, lfBasinLength, lfMaximumDepth, lfSideSlope1, lfSideSlope2, lfSideSlope3, lfManning;
				CString strID;
				sscanf(str,"%s %lf %lf %lf %lf %lf %lf %lf",
					strID,
					&lfBasinWidth,
					&lfBasinLength,
					&lfMaximumDepth,
					&lfSideSlope1,
					&lfSideSlope2,
					&lfSideSlope3,
					&lfManning);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_B)
				{
					strError.Format("BMP site (ID = %s) is not in class B", strID);
					return false;
				}

				BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;
				pBMP->m_lfBasinWidth   = lfBasinWidth;
				pBMP->m_lfBasinLength  = lfBasinLength;
				pBMP->m_lfMaximumDepth = lfMaximumDepth;
				if (lfMaximumDepth <= 0)
				{
					str.Format("Check Maximum Depth for the site, ID = %s" ,strID);
					AfxMessageBox(str);
					return false;
				}
				pBMP->m_lfSideSlope1 = lfSideSlope1;
				if (lfSideSlope1 <= 0)
				{
					str.Format("Check Side Slope1 for the site, ID = %s" ,strID);
					AfxMessageBox(str);
					return false;
				}
				pBMP->m_lfSideSlope2 = lfSideSlope2;
				if (lfSideSlope2 <= 0)
				{
					str.Format("Check Side Slope2 for the site, ID = %s" ,strID);
					AfxMessageBox(str);
					return false;
				}
				pBMP->m_lfSideSlope3 = lfSideSlope3;
				if (lfSideSlope3 <= 0)
				{
					str.Format("Check Side Slope3 for the site, ID = %s" ,strID);
					AfxMessageBox(str);
					return false;
				}
				pBMP->m_lfManning = lfManning;
				if (lfManning <= 0)
				{
					str.Format("Check Manning's N value for the site, ID = %s" ,strID);
					AfxMessageBox(str);
					return false;
				}
			}
			break;
		case 740:	// soil properties for class A and B
			if(nBMPA+nBMPB == 0)
				return true;

			// create arrays
			if (GAInfil != NULL)	FREE(GAInfil);
			GAInfil = (TGrnAmpt *) calloc(nBMPA+nBMPB, sizeof(TGrnAmpt));
		    if ( GAInfil == NULL ) return ERR_MEMORY;

			// initialize 
			InitializeGAInfil(nBMPA+nBMPB);

			nIndex = 0;
			while (ReadDataLine(fpin, str))
			{
				int nInfiltMethod, nPolRotMethod, nPolRemMethod, nUndSwitch;
				double lfSoilDepth, lfPorosity, lfFCapacity, lfWPoint, lfVegA, lfInfilt, 
					   lfUndDepth, lfUndVoid, lfUndInfilt, lfSuction, lfHydCon, lfIMDmax;
				float  x[3];
				CString strID;				
				sscanf(str,"%s %d %d %d %lf %lf %lf %lf %lf %lf %d %lf %lf %lf %lf %lf %lf",
					strID,
					&nInfiltMethod,
					&nPolRotMethod,
					&nPolRemMethod,
					&lfSoilDepth,
					&lfPorosity,
					&lfFCapacity,
					&lfWPoint,
					&lfVegA,
					&lfInfilt,
					&nUndSwitch,
					&lfUndDepth,
					&lfUndVoid,
					&lfUndInfilt,
					&lfSuction,
					&lfHydCon,
					&lfIMDmax);

				if (lfSuction == 0)
					lfSuction = 3.0;//default value (in)
				if (lfHydCon == 0)
					lfHydCon = 0.5;//default value (in/hr)
				if (lfIMDmax == 0)
					lfIMDmax = 0.3;//default value (fraction)

				if (nInfiltMethod != 1) nInfiltMethod = 0;

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot lfind BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_nGAInfil_Index			= nIndex;
				pBMPSite->m_nInfiltMethod			= nInfiltMethod;
				pBMPSite->m_nPolRotMethod			= nPolRotMethod;
				pBMPSite->m_nPolRemMethod			= nPolRemMethod;
				pBMPSite->m_lfSoilDepth				= lfSoilDepth;
				pBMPSite->m_lfPorosity				= lfPorosity;
				pBMPSite->m_lfFCapacity				= lfFCapacity;
				pBMPSite->m_lfWPoint				= lfWPoint;
				pBMPSite->m_lfUndDepth				= lfUndDepth;
				pBMPSite->m_lfUndVoid				= lfUndVoid;
				pBMPSite->m_lfUndInfilt				= lfUndInfilt;
				pBMPSite->m_bUndSwitch				= (nUndSwitch==0)?false:true;
				pBMPSite->m_holtanParam.m_lfVegA	= lfVegA;
				pBMPSite->m_holtanParam.m_lfFInfilt	= lfInfilt;

				x[0] = lfSuction;
				x[1] = lfHydCon;
				x[2] = lfIMDmax;

				if ( !grnampt_setParams(nIndex, x) ) 
					return error_setInpError(ERR_NUMBER, "");
				
				infil_initState(nIndex,GREEN_AMPT);

				// copy data
				if (pBMPSite->m_nBMPClass == CLASS_A)
				{
					BMP_A* pBMP = (BMP_A*) pBMPSite->m_pSiteProp;
					CopyGAInfil(&GAInfil[nIndex], &pBMP->m_pGAInfil);
				}
				else if (pBMPSite->m_nBMPClass == CLASS_B)
				{
					BMP_B* pBMP = (BMP_B*) pBMPSite->m_pSiteProp;
					CopyGAInfil(&GAInfil[nIndex], &pBMP->m_pGAInfil);
				}

				nIndex++;
			}
			break;
		case 745:
			while (ReadDataLine(fpin, str))
			{
				//int nSiteID;
				double lfGrowth[12];
				CString strID;			
				sscanf(str,"%s %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf",
					strID,
					&lfGrowth[0],
					&lfGrowth[1],
					&lfGrowth[2],
					&lfGrowth[3],
					&lfGrowth[4],
					&lfGrowth[5],
					&lfGrowth[6],
					&lfGrowth[7],
					&lfGrowth[8],
					&lfGrowth[9],
					&lfGrowth[10],
					&lfGrowth[11]);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot lfind BMP site with ID %s", strID);
					return false;
				}

				for(i=0; i<12; i++)
					pBMPSite->m_holtanParam.m_lfGrowth[i] = lfGrowth[i];
			}
			break;
		case 750:	// SWMM5 CONDUITS
			if(nBMPC == 0)
				return true;

			// create arrays
			if (Link != NULL)	FREE(Link);
			Link = (TLink *)    calloc(nBMPC, sizeof(TLink));
		    if ( Link == NULL ) return ERR_MEMORY;
			
			if (Conduit != NULL)	FREE(Conduit);
			Conduit = (TConduit *)  calloc(nBMPC, sizeof(TConduit));
		    if ( Conduit == NULL )	return ERR_MEMORY;

			if (Transect != NULL)	transect_delete();
			transect_create(nBMPC);

			InitializeLinkConduitTransect(nBMPC);

			nIndex = 0;
			while (ReadDataLine(fpin, str))
			{
				int nInlet, nOutlet;
				double lfCondLength, lfManning, lfInletH, lfOutletH, lfInitflow;
				double lfEntLossCoeff, lfExtLossCoeff, lfAvgLossCoeff;
				CString strID;
				sscanf(str,"%s %d %d %lf %lf %lf %lf %lf %lf %lf %lf",
					strID,
					&nInlet,
					&nOutlet,
					&lfCondLength,
					&lfManning,
					&lfInletH,
					&lfOutletH,
					&lfInitflow,
					&lfEntLossCoeff,
					&lfExtLossCoeff,
					&lfAvgLossCoeff);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_C)
				{
					strError.Format("BMP site (ID = %s) is not in class C", strID);
					return false;
				}

				BMP_C* pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
				pBMP->m_strID = strID;
				pBMP->m_nIndex = nIndex;

				Link[nIndex].subIndex = nIndex;
				Link[nIndex].node1 = nInlet;       
				Link[nIndex].node2 = nOutlet;      
				Link[nIndex].z1 = lfInletH;        
				Link[nIndex].z2 = lfOutletH;       
				Link[nIndex].q0 = lfInitflow;      
				Link[nIndex].cLossInlet = lfEntLossCoeff;
				Link[nIndex].cLossOutlet = lfExtLossCoeff;
				Link[nIndex].cLossAvg = lfAvgLossCoeff;
				
				Conduit[nIndex].length = lfCondLength;      
				Conduit[nIndex].roughness = lfManning;      
				Conduit[nIndex].modLength = lfCondLength;   

				nIndex++;
			}
			break;
		case 755:	// SWMM5 XSECTIONS 
			if(nBMPC == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				int nBarrels;
				CString sCondType;
				CString sCondName;
				double lfGeom1, lfGeom2, lfGeom3, lfGeom4;

				CStringToken strToken(str);
				// get site id
				//nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				// conduit type
				sCondType = strToken.NextToken();
				// conduit name
				sCondName = strToken.NextToken();

				if (sCondType != "IRREGULAR")
					lfGeom1 = atof((LPCSTR)sCondName);
				else
					lfGeom1 = 0;

				// get conduit geometry
				lfGeom2 = atof((LPCSTR)strToken.NextToken());
				// get conduit geometry
				lfGeom3 = atof((LPCSTR)strToken.NextToken());
				// get conduit geometry
				lfGeom4 = atof((LPCSTR)strToken.NextToken());
				// get number of barrels
				nBarrels = atoi((LPCSTR)strToken.NextToken());

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot lfind BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_C)
				{
					strError.Format("BMP site (ID = %s) is not in class C", strID);
					return false;
				}
				BMP_C* pBMP = (BMP_C*) pBMPSite->m_pSiteProp;

				nIndex = pBMP->m_nIndex;

				pBMP->m_strCondType	= sCondType;

				// --- get code of xsection shape
				char* s = sCondType.GetBuffer(sCondType.GetLength());
				int k = findmatch(s, XsectTypeWords);
				if ( k < 0 )
				{
					strError.Format("Cannot find Conduit type: %s", sCondType);
					return false;
				}
				Link[nIndex].xsect.type = k;
				if (k == IRREGULAR)
				{
					pBMP->m_strCondName = sCondName;
					Link[nIndex].xsect.transect = nIndex;	
				}
				else
				{
					//double UCF = 1.0;
					float x[4];
					x[0] = lfGeom1;
					x[1] = lfGeom2;
					x[2] = lfGeom3;
					x[3] = lfGeom4;
					xsect_setParams(&Link[nIndex].xsect, k, x, UCF(LENGTH));
				}
				Conduit[nIndex].barrels = nBarrels;
			}
			break;

		case 760:	// SWMM5 TRANSECTS 
		//  Purpose: read parameters of a transect from a tokenized line of input data.
		//
		//  Format of transect data follows that used for HEC-2 program:
		//    NC  nLeft  nRight  nChannel
		//    X1  name  nSta  xLeftBank  xRightBank  0  0  0  xFactor  yFactor
		//    GR  Elevation  Station  ... 
			if(nBMPC == 0)
				return true;
			while (ReadDataLine(fpin, str))
			{
				int   i, k;
				int   index;              // transect index
				char tok[2];			  // input line values
				float x[10];              // parameter values
				int ntoks = 0;			  // number of tokens
				CString sCondName;
				CStringToken tmpToken(str);
				while (tmpToken.HasMoreTokens())
				{
					ntoks++;
					tmpToken.NextToken();
				}
				CStringToken strToken(str);
				CString strTemp = strToken.NextToken();
				sscanf(strTemp,"%s", &tok[0]);
				// --- match first token to a transect keyword
				char* s = strTemp.GetBuffer(strTemp.GetLength());
				k = findmatch(s, TransectKeyWords);
				if ( k < 0 )
				{
					strError.Format("Cannot find Transect line: %s", tok[0]);
					return false;
				}
				// --- read parameters associated with keyword
				switch ( k )
				{
				  // --- NC line: Manning n values
				  case 0:
				    // --- finish processing the previous transect
					if (ncount > 0)
					{
						CBMPSite *pBMPSite;
						POSITION pos;
						BMP_C* pBMP;
						// check the conduit with irregular shape having same transect name
						pos = bmpsiteList.GetHeadPosition();
						while (pos != NULL)
						{
							pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
							if (pBMPSite->m_nBMPClass == CLASS_C)
							{
								pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
								if (pBMP->m_strCondName == id && id != "")
								{
									transect_validate(Link[pBMP->m_nIndex].xsect.transect);
									xsect_setIrregXsectParams(&Link[pBMP->m_nIndex].xsect);
								}
							}
						}
					}
					// --- update total transect count
					ncount++;
					// --- read Manning's n values
					if ( ntoks < 4 ) 
					{
						strError.Format("Missing parameters on Transect line: %s", tok[0]);
						return false;
					}
					for (i = 1; i <= 3; i++)
						x[i] = atof((LPCSTR)strToken.NextToken());
					setManning(x);
					break;
				  // --- X1 line: identifies start of next transect
				  case 1:
					// --- check that transect was already added to project
					//     (by input_countObjects)
					if ( ntoks < 10 ) 
					{
						strError.Format("Missing parameters on Transect line: %s", tok[0]);
						return false;
					}
					strTemp = strToken.NextToken();
					sscanf(strTemp,"%s", &tok[1]);
					sCondName =  strTemp;
					id = strTemp;
					// --- read in rest of numerical values on data line
					for ( i = 2; i < 10; i++ )
						x[i] = atof((LPCSTR)strToken.NextToken());

					CBMPSite *pBMPSite;
					POSITION pos;
					BMP_C* pBMP;
					// check the conduit with irregular shape having same transect name
					pos = bmpsiteList.GetHeadPosition();
					while (pos != NULL)
					{
						pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
						if (pBMPSite->m_nBMPClass == CLASS_C)
						{
							pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
							if (pBMP->m_strCondName == id && id != "")
							{
								index = Link[pBMP->m_nIndex].xsect.transect; 
								// --- transfer parameter values to transect's properties
								char* s = id.GetBuffer(id.GetLength());
								setParams(index, s, x);
							}
						}
					}
					break;
				  // --- GR line: station elevation & location data
				  case 2:
					// --- check that line contains pairs of data values
					if ( (ntoks - 1) % 2 > 0 ) 
					{
						strError.Format("Cannot find pairs of data values on Transect line: %s", tok[0]);
						return false;
					}
					// --- parse each pair of Elevation-Station values
					i = 1;
					while ( i < ntoks )
					{
						x[1] = atof((LPCSTR)strToken.NextToken());
						x[2] = atof((LPCSTR)strToken.NextToken());
						addStation(x[1], x[2]);
						i += 2;
					}
					break;
				}
			}
			// --- finish processing the last transect
			if (ncount > 0)
			{
				CBMPSite *pBMPSite;
				POSITION pos;
				BMP_C* pBMP;
				// check the conduit with irregular shape having same transect name
				pos = bmpsiteList.GetHeadPosition();
				while (pos != NULL)
				{
					pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
					if (pBMPSite->m_nBMPClass == CLASS_C)
					{
						pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
						if (pBMP->m_strCondName == id && id != "")
						{
							transect_validate(Link[pBMP->m_nIndex].xsect.transect);
							xsect_setIrregXsectParams(&Link[pBMP->m_nIndex].xsect);
						}
					}
				}
			}
			break;
		case 765:
			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 765 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}
				
				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_pDecay = new double[nPollutants];
				
				j = 0;
				while (strToken.HasMoreTokens())
					pBMPSite->m_pDecay[j++] = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 766:
			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 766 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}
				
				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_pK = new double[nPollutants];
				
				j = 0;
				while (strToken.HasMoreTokens())
					pBMPSite->m_pK[j++] = atof((LPCSTR)strToken.NextToken())/8760.0;//ft/yr to ft/hr
				i++;
			}
			break;
		case 767:
			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 767 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}
				
				CStringToken strToken(str);
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_pCstar = new double[nPollutants];
				
				j = 0;
				while (strToken.HasMoreTokens())
				{
					pBMPSite->m_pCstar[j++] = atof((LPCSTR)strToken.NextToken())/LBpCFT2MGpL;//lb/ft3
				}
				if (j != nPollutants)
					return false;
				i++;
			}
			break;
		case 770:  
			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 770 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_pUndRemoval = new double[nPollutants];
				
				j = 0;
				while (strToken.HasMoreTokens())
					pBMPSite->m_pUndRemoval[j++] = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 775:
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_sediment.m_lfBEDWID		= atof((LPCSTR)strToken.NextToken());//ft
				pBMPSite->m_sediment.m_lfBEDDEP		= atof((LPCSTR)strToken.NextToken());//ft
				pBMPSite->m_sediment.m_lfBEDPOR		= atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_sediment.m_lfSAND_FRAC	= atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_sediment.m_lfSILT_FRAC	= atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_sediment.m_lfCLAY_FRAC  = atof((LPCSTR)strToken.NextToken());
			}
			break;
		case 780:	//sand
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_sand.m_lfD = atof((LPCSTR)strToken.NextToken());//in
				pBMPSite->m_sand.m_lfW = atof((LPCSTR)strToken.NextToken());//in/sec
				pBMPSite->m_sand.m_lfRHO = atof((LPCSTR)strToken.NextToken());//lb/ft3
				pBMPSite->m_sand.m_lfKSAND = atof((LPCSTR)strToken.NextToken());
				pBMPSite->m_sand.m_lfEXPSND = atof((LPCSTR)strToken.NextToken());
			}
			break;
		case 785:	//silt
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_silt.m_lfD = atof((LPCSTR)strToken.NextToken());//in
				pBMPSite->m_silt.m_lfW = atof((LPCSTR)strToken.NextToken());//in/s
				pBMPSite->m_silt.m_lfRHO = atof((LPCSTR)strToken.NextToken());//lb/ft3
				pBMPSite->m_silt.m_lfTAUCD = atof((LPCSTR)strToken.NextToken());//lb/ft2
				pBMPSite->m_silt.m_lfTAUCS = atof((LPCSTR)strToken.NextToken());//lb/ft2
				pBMPSite->m_silt.m_lfW = atof((LPCSTR)strToken.NextToken());//lb/ft2/day
			}
			break;
		case 786:	//clay
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_clay.m_lfD = atof((LPCSTR)strToken.NextToken());//in
				pBMPSite->m_clay.m_lfW = atof((LPCSTR)strToken.NextToken());//in/s
				pBMPSite->m_clay.m_lfRHO = atof((LPCSTR)strToken.NextToken());//lb/ft3
				pBMPSite->m_clay.m_lfTAUCD = atof((LPCSTR)strToken.NextToken());//lb/ft2
				pBMPSite->m_clay.m_lfTAUCS = atof((LPCSTR)strToken.NextToken());//lb/ft2
				pBMPSite->m_clay.m_lfW = atof((LPCSTR)strToken.NextToken());//lb/ft2/day
			}
			break;
		case 790:
			//required if the land simulation control is external
			if (nLandSimulation == 0)
			{
				while (ReadDataLine(fpin, str))
				{
					int nSiteLuID, nLuID;
					double lfArea;
					CString strID;
					sscanf(str, "%d %d %lf %s", &nSiteLuID, &nLuID, &lfArea, strID);

					CLandUse *pLU = (CLandUse *) FindLandUse(nLuID);
					if (pLU == NULL)
					{
						strError.Format("Cannot find Landuse type with ID %d", nLuID);
						return false;
					}

					CBMPSite* pBMPSite = FindBMPSite(strID);
					CSiteLandUse *pSiteLU = new CSiteLandUse(pLU, pBMPSite, lfArea);
					siteluList.AddTail(pSiteLU);
				}
			}
			break;
		case 795:
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				int nOutletType;
				CString strID, strID2;
				strID = strToken.NextToken();
				nOutletType = atoi((LPCSTR)strToken.NextToken());
				strID2 = strToken.NextToken();

				//sscanf(str,"%s %d %s", strID, &nOutletType, strID2);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				DS_BMPSITE *pDS = new DS_BMPSITE;
				pDS->m_nOutletType = nOutletType;
				pDS->m_pDSBMPSite = FindBMPSite(strID2);

				pBMPSite->m_dsbmpsiteList.AddTail(pDS);
			}
			break;
		case 800:
			// loading running option and cost limit
			ReadDataLine(fpin, str);
			if (str.GetLength() > 0)
			{
				CStringToken strToken(str);
				nStrategy    = atoi((LPCSTR)strToken.NextToken());
				nRunOption   = atoi((LPCSTR)strToken.NextToken());
				//lfCostLimit  = atof((LPCSTR)strToken.NextToken());
				lfStopDelta  = atof((LPCSTR)strToken.NextToken());
				lfMaxRunTime = atof((LPCSTR)strToken.NextToken());
				nSolution    = atoi((LPCSTR)strToken.NextToken());
				if (strToken.HasMoreTokens())
				{
					nTargetBreak = atoi((LPCSTR)strToken.NextToken());
					if (nTargetBreak < 1)
						nTargetBreak = 1;
				}
			}
			break;
		case 805:
			while (ReadDataLine(fpin, str))
			{
				//int nSiteID;
				double lfLinearCost, lfAreaCost, lfTotalVolumeCost, lfMediaVolumeCost,
					   lfUnderDrainVolumeCost, lfConstantCost, lfPercentCost,
					   lfLengthExp, lfAreaExp, lfTotalVolExp, lfMediaVolExp, lfUDVolExp;

				CString strID;
				sscanf(str, "%s %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf", strID, 
					   &lfLinearCost, &lfAreaCost, &lfTotalVolumeCost, &lfMediaVolumeCost, 
					   &lfUnderDrainVolumeCost, &lfConstantCost, &lfPercentCost, 
					   &lfLengthExp, &lfAreaExp, &lfTotalVolExp, &lfMediaVolExp, 
					   &lfUDVolExp);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot lfind BMP site with ID %s", strID);
					return false;
				}

				pBMPSite->m_costParam.m_lfLinearCost = lfLinearCost;			
				pBMPSite->m_costParam.m_lfAreaCost = lfAreaCost;            				
				pBMPSite->m_costParam.m_lfTotalVolumeCost = lfTotalVolumeCost;     				
				pBMPSite->m_costParam.m_lfMediaVolumeCost = lfMediaVolumeCost;     				
				pBMPSite->m_costParam.m_lfUnderDrainVolumeCost = lfUnderDrainVolumeCost;				
				pBMPSite->m_costParam.m_lfConstantCost = lfConstantCost;        				
				pBMPSite->m_costParam.m_lfPercentCost = lfPercentCost;         
				pBMPSite->m_costParam.m_lfLengthExp = lfLengthExp;         
				pBMPSite->m_costParam.m_lfAreaExp = lfAreaExp;         
				pBMPSite->m_costParam.m_lfTotalVolExp = lfTotalVolExp;         
				pBMPSite->m_costParam.m_lfMediaVolExp = lfMediaVolExp;         
				pBMPSite->m_costParam.m_lfUDVolExp = lfUDVolExp;         
			}
			break;
		case 810:
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				ADJUSTABLE_PARAM* pAdjustableParam = new ADJUSTABLE_PARAM;
				pAdjustableParam->m_strVariable = strToken.NextToken();
				pAdjustableParam->m_lfFrom      = atof((LPCSTR)strToken.NextToken());
				pAdjustableParam->m_lfTo        = atof((LPCSTR)strToken.NextToken());
				pAdjustableParam->m_lfStep      = atof((LPCSTR)strToken.NextToken());

				pBMPSite->m_adjustList.AddTail(pAdjustableParam);
				nAdjVariable++;
			}

			if (nRunOption != OPTION_NO_OPTIMIZATION) 
			{
				if (nAdjVariable == 0)
				{
					//error message and stop the application
					AfxMessageBox("No decision variables defined!\nOptimization run requires decision variables!");
					return false;
				}
			}
			break;
		case 815:
			// loading evaluation factor values
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				//int nSiteID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}

				EVALUATION_FACTOR* pEvalFactor = new EVALUATION_FACTOR;
				pEvalFactor->m_nFactorGroup = atoi((LPCSTR)strToken.NextToken());
				pEvalFactor->m_nFactorType  = atoi((LPCSTR)strToken.NextToken());

				pEvalFactor->m_nCalcDays = 0;
				pEvalFactor->m_lfThreshold = 0.0;
				pEvalFactor->m_lfConcThreshold = 0.0;	//optional
				if (pEvalFactor->m_nFactorType == MAC)
					pEvalFactor->m_nCalcDays = atoi((LPCSTR)strToken.NextToken());
				else if (pEvalFactor->m_nFactorType == FEF)
					pEvalFactor->m_lfThreshold = atof((LPCSTR)strToken.NextToken());
				else if (pEvalFactor->m_nFactorType == CEF)	//optional
					pEvalFactor->m_lfConcThreshold = atof((LPCSTR)strToken.NextToken());
				else
					strToken.NextToken();
				pEvalFactor->m_nCalcMode    = atoi((LPCSTR)strToken.NextToken());

				pEvalFactor->m_lfTarget      = 0.0;
				pEvalFactor->m_lfNextTarget  = 0.0;
				pEvalFactor->m_lfPriorFactor = 0.0;
				pEvalFactor->m_lfLowerTarget = 0.0;
				pEvalFactor->m_lfUpperTarget = 0.0;

				if (nRunOption == OPTION_TRADE_OFF_CURVE)
				{
					// only one Evaluation Factor is allowed for TradeOff Curve option
					if (nEvalFactor > 0)
					{
						//error message and stop the application
						AfxMessageBox("Only one Evaluation Factor is allowed for TradeOff Curve option!");
						return false;
					}

					pEvalFactor->m_lfLowerTarget = atof((LPCSTR)strToken.NextToken());
					pEvalFactor->m_lfUpperTarget = atof((LPCSTR)strToken.NextToken());
				}
				// Read in priority factor here (JZ)
				else if (nRunOption == OPTION_MAXIMIZE_CONTROL)
				{
					pEvalFactor->m_lfPriorFactor = atof((LPCSTR)strToken.NextToken());
				}
				else
				{
					pEvalFactor->m_lfTarget = atof((LPCSTR)strToken.NextToken());
				}

				pEvalFactor->m_strFactor    = strToken.LeftOut();
				pEvalFactor->m_lfInit       = 0.0;
				pEvalFactor->m_lfPreDev     = 0.0;
				pEvalFactor->m_lfCurrent    = 0.0;
				pEvalFactor->m_lfPostDev	= 0.0;	// (04-2005)

				pBMPSite->m_factorList.AddTail(pEvalFactor);
				nEvalFactor++;
			}
			break;
		case 820:
			// loading bmp group list 
			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				int nGroupID = atoi((LPCSTR)strToken.NextToken());

				BMP_GROUP* pBMPGroup = new BMP_GROUP;
				pBMPGroup->m_nGroupID = nGroupID;

				CStringToken strSiteIDToken(strToken.NextToken(), ",");

				//int nSiteID;
				while (strSiteIDToken.HasMoreTokens())
				{
					//nSiteID = atoi((LPCSTR)strSiteIDToken.NextToken());
					CString strID = strToken.NextToken();
					CBMPSite* pBMPSite = FindBMPSite(strID);
					if (pBMPSite == NULL)
					{
						strError.Format("Cannot find BMP site with ID %s", strID);
						return false;
					}
					pBMPGroup->m_bmpList.AddTail(pBMPSite);
				}

				pBMPGroup->m_lfTotalArea = atof((LPCSTR)strToken.NextToken());

				bmpGroupList.AddTail(pBMPGroup);
			}
			break;
		case 900:
			if(nBMPD == 0)
				return true;
			{
				ReadDataLine(fpin, str);
				sscanf(str, "%d %lf %lf %d %d %d %d",
					&nN,&lfTHETAW,&lfCR,&nMAXITER,&nNPOL,&nIELOUT,&nKPG);
			}
			break;
		case 901:
			if(nBMPD == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				CStringToken strToken(str);
				CString strID = strToken.NextToken();

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
				pBMP->m_strID     = strID;							
				pBMP->m_strName   = strToken.NextToken();
				pBMP->m_lfWidth   = atof((LPCSTR)strToken.NextToken())*0.3048;//ft->m
				pBMP->m_lfLength  = atof((LPCSTR)strToken.NextToken())*0.3048;//ft->m
				pBMP->m_nSegments = atoi((LPCSTR)strToken.NextToken());

				//assign memory here
				if (pBMP->m_pSEGMENT_D != NULL)	
					delete []pBMP->m_pSEGMENT_D;
				if (pBMP->m_nSegments > 0)
					pBMP->m_pSEGMENT_D = new SEGMENT_D[pBMP->m_nSegments];

				if (pBMP->m_pPOLLUTANT_D != NULL) 
					delete []pBMP->m_pPOLLUTANT_D;
				if (nPollutant > 0)
					pBMP->m_pPOLLUTANT_D = new POLLUTT_D[nPollutant];

				//initialize arrays
				for (i=0; i<pBMP->m_nSegments; i++)
				{
					pBMP->m_pSEGMENT_D[i].m_nSegmentID = i+1;
					pBMP->m_pSEGMENT_D[i].m_lfRNA = 0.0;
					pBMP->m_pSEGMENT_D[i].m_lfSOA = 0.0;
					pBMP->m_pSEGMENT_D[i].m_lfSX = 0.0;
				}

				for (i=0; i<nPollutant; i++)
				{
					pBMP->m_pPOLLUTANT_D[i].m_lfQUALDECAY_ads = 0.0;
					pBMP->m_pPOLLUTANT_D[i].m_lfQUALDECAY_dis = 0.0;
					pBMP->m_pPOLLUTANT_D[i].m_lfQUALSED_frac  = 0.0;
					pBMP->m_pPOLLUTANT_D[i].m_lfTEMPCORR_ads  = 0.0;
					pBMP->m_pPOLLUTANT_D[i].m_lfTEMPCORR_dis  = 0.0;
				}
			}
			break;
		case 902:
			if(nBMPD == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				int nSEGID;
				double lfSX, lfRNA, lfSOA;
				CString strID;
				sscanf(str, "%s %d %lf %lf %lf",
					strID,
					&nSEGID,
					&lfSX,
					&lfRNA,
					&lfSOA);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
				int nNumSeg = pBMP->m_nSegments;

				if (nNumSeg <= 0 || nSEGID > nNumSeg)
				{
					strError.Format("The number of segments for BMP site (ID = %s) shold be greater than 0", strID);
					return false;
				}

				if (nSEGID > nNumSeg)
				{
					strError.Format("The number of segments for BMP site (ID = %s) are greater than defined in Card 901", strID);
					return false;
				}

				pBMP->m_pSEGMENT_D[nSEGID-1].m_nSegmentID = nSEGID;
				pBMP->m_pSEGMENT_D[nSEGID-1].m_lfRNA = lfRNA;
				pBMP->m_pSEGMENT_D[nSEGID-1].m_lfSOA = lfSOA;
				pBMP->m_pSEGMENT_D[nSEGID-1].m_lfSX  = lfSX*0.3048;//ft->m
			}
			break;
		case 903:
			if(nBMPD == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				//int nVFSID;
				double lfVKS, lfSAV, lfOS, lfOI, lfSM, lfSCHK;
				CString strID;
				sscanf(str, "%s %lf %lf %lf %lf %lf %lf",
					strID,
					&lfVKS,
					&lfSAV,
					&lfOS,
					&lfOI,
					&lfSM,
					&lfSCHK);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
				pBMP->m_lfVKS   = lfVKS*0.0254/3600.0;//in/hr->m/s
				pBMP->m_lfSAV   = lfSAV*0.3048;//ft->m
				pBMP->m_lfOS    = lfOS;
				pBMP->m_lfOI	= lfOI;
				pBMP->m_lfSM	= lfSM*0.3048;//ft->m
				pBMP->m_lfSCHK  = lfSCHK;
			}
			break;
		case 904:
			if(nBMPD == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				int nICO;
				double lfSS, lfVN, lfH, lfVN2;
				CString strID;
				sscanf(str, "%s %lf %lf %lf %lf %d",
					strID,
					&lfSS,
					&lfVN,
					&lfH,
					&lfVN2,
					&nICO);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
				pBMP->m_lfSS	= lfSS*2.54;//in->cm
				pBMP->m_lfVN	= lfVN;
				pBMP->m_lfH		= lfH*30.48;//ft->cm
				pBMP->m_lfVN2	= lfVN2;
				pBMP->m_nICO	= nICO;
			}
			break;
		case 905:
			if(nBMPD == 0)
				return true;

			while (ReadDataLine(fpin, str))
			{
				int nNPART[3];
				double lfCOARSE[3], lfPOR[3], lfDP[3], lfSG[3];
				CString strID;
				sscanf(str, "%s %d %d %d %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf %lf",
					strID,
					&nNPART[0],
					&nNPART[1],
					&nNPART[2],
					&lfCOARSE[0],
					&lfCOARSE[1],
					&lfCOARSE[2],
					&lfPOR[0],
					&lfPOR[1],
					&lfPOR[2],
					&lfDP[0],
					&lfDP[1],
					&lfDP[2],
					&lfSG[0],
					&lfSG[1],
					&lfSG[2]);

				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				for (i=0; i<3; i++)
				{
					pBMP->m_nNPART[i]	= nNPART[i];
					pBMP->m_lfCOARSE[i]	= lfCOARSE[i];
					pBMP->m_lfPOR[i]	= lfPOR[i];
					pBMP->m_lfDP[i]		= lfDP[i]*2.54;//in->cm
					pBMP->m_lfSG[i]		= lfSG[i]*453.5924/28316.85;//lb/ft3->g/cm3
				}
			}
			break;
		case 906:
			if(nBMPD == 0)
				return true;

			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 906 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				//read BMPSite ID
				CStringToken strToken(str);
				//int nVFSID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				
				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				j = 0;
				while (strToken.HasMoreTokens())
					pBMP->m_pPOLLUTANT_D[j++].m_lfQUALSED_frac = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 907:
			if(nBMPD == 0)
				return true;

			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 907 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				//read BMPSite ID
				CStringToken strToken(str);
				//int nVFSID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				
				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				j = 0;
				while (strToken.HasMoreTokens())
					pBMP->m_pPOLLUTANT_D[j++].m_lfQUALDECAY_ads = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 908:
			if(nBMPD == 0)
				return true;

			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 908 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				//read BMPSite ID
				CStringToken strToken(str);
				//int nVFSID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				
				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				j = 0;
				while (strToken.HasMoreTokens())
					pBMP->m_pPOLLUTANT_D[j++].m_lfQUALDECAY_dis = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 909:
			if(nBMPD == 0)
				return true;

			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 909 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				//read BMPSite ID
				CStringToken strToken(str);
				//int nVFSID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				
				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				j = 0;
				while (strToken.HasMoreTokens())
					pBMP->m_pPOLLUTANT_D[j++].m_lfTEMPCORR_ads = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		case 910:
			if(nBMPD == 0)
				return true;

			i = 0;
			nPollutants = 0;

			while (ReadDataLine(fpin, str))
			{
				if (i == 0)
				{
					CStringToken tmpToken(str);
					while(tmpToken.HasMoreTokens())
					{
						nPollutants++;
						tmpToken.NextToken();
					}
					// first token is the BMP Site ID
					nPollutants--;

					if (nPollutants != nPollutant)
					{
						strError.Format("The number of pollutants in card 910 (%d) are different than in card 705 (%d)", nPollutants, nPollutant);
						return false;
					}
				}

				//read BMPSite ID
				CStringToken strToken(str);
				//int nVFSID = atoi((LPCSTR)strToken.NextToken());
				CString strID = strToken.NextToken();
				
				CBMPSite* pBMPSite = FindBMPSite(strID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strID);
					return false;
				}
				if (pBMPSite->m_nBMPClass != CLASS_D) 
				{
					strError.Format("BMP site (ID = %s) is not in class D", strID);
					return false;
				}

				BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

				j = 0;
				while (strToken.HasMoreTokens())
					pBMP->m_pPOLLUTANT_D[j++].m_lfTEMPCORR_dis = atof((LPCSTR)strToken.NextToken());
				i++;
			}
			break;
		default:
			break;
	}

	return true;
}

void CBMPData::SkipCommentLine(FILE *fp)
{
	char strLine[MAXLINE];
	long nStart = ftell (fp);

	while (!feof(fp))
    {
		fgets(strLine, MAXLINE, fp);
		CString str(strLine);
		str.TrimLeft();
		str.TrimRight();

		if (str.GetLength() == 0)
			continue;

		if(str[0] == 'C' || str[0] == 'c')
		{
			CStringToken strToken(str);
			CString str0 = strToken.NextToken();
			if(str0.GetLength() >= 2 && str0[1] >= '0' && str0[1] <= '9')
			{
				fseek (fp, nStart, SEEK_SET);
				return;
			}
		}
		else
		{
			fseek (fp, nStart, SEEK_SET);
			return;
		}

		nStart = ftell (fp);
	}
}

bool CBMPData::ReadDataLine(FILE *fp, CString& strData)
{
	char strLine[MAXLINE];
	strData = "";
	long nStart = ftell(fp);

	while (!feof(fp))
    {
		strLine[0] = '\0';
		fgets(strLine, MAXLINE, fp);
		CString str(strLine);
		str.TrimLeft();
		str.TrimRight();
		//skip blank line
		if (str.GetLength() == 0)
		{
			nStart = ftell(fp);
			continue;
		}

		//stop at comment line
		if(str[0] == 'C' || str[0] == 'c')
		{
			fseek (fp, nStart, SEEK_SET);
			return false;
		}
		//read data line
		else
		{
			strData = str;
			return true;
		}
	}

	return false;
}

bool CBMPData::ReadBestPopFile(int nBestPopId)
{
	int i;
	FILE *fpin = NULL;
	char strLine[MAXLINE];
	bool retVal = true;
	bool bBestPopId = false;
	CString strFilePath, strDecVar, strBMPID;

	//assume input file is already read
	//open the file for reading
	strFilePath = strOutputDir + "\\BestSolutions.out";
	fpin = fopen (strFilePath, "rt");
	if(fpin == NULL)
	{
		strError = "Cannot open file " + strFilePath + " for reading";
		return false;
	}

	//skip the first header line
	i = 1;
	while(i-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	//read the second header line
	i = 1;
	while(i-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);
	
	CString str(strLine);
	CStringToken strToken(str);
	CStringToken strToken1(str);
	
	//skip eight values (NO., TotalCost($), TotalSurfaceArea(ac), TotalExcavatnVol(ac-ft), TotalSurfStorVol(ac-ft), TotalSoilStorVol(ac-ft), TotalUdrnStorVol(ac-ft), EvalFactor)
	for (i=0; i<8; i++)
	{
		strToken.NextToken();
		strToken1.NextToken();
	}

	//skip cost for each unique bmp type
	for(i=0; i<nBMPtype; i++)
	{
		strToken.NextToken();
		strToken1.NextToken();
	}

	//find the number of decision variables
	int nDecVar = 0;
	while (strToken.HasMoreTokens())
	{
		strToken.NextToken();
		nDecVar++;
	}

	//read the parameters
	while (fgets(strLine, MAXLINE, fpin) != NULL)
	{
		CString str2(strLine);
		CStringToken strToken2(str2);
		
		//read the Best Pop ID
		int nID = atoi(LPCSTR(strToken2.NextToken()));
		if (nID == nBestPopId)
		{
			bBestPopId = true;

			//skip seven values
			for (i=0; i<7; i++)
				strToken2.NextToken();
			
			//skip cost for each unique bmp type
			for(i=0; i<nBMPtype; i++)
				strToken2.NextToken();

			//read the decision variables
			for (i=0; i<nDecVar; i++)
			{
				strDecVar = strToken1.NextToken();
				int nIndex = strDecVar.ReverseFind('_');
				strBMPID = strDecVar.Left(nIndex);
				strDecVar = strDecVar.Right(strDecVar.GetLength()-nIndex-1);
				
				CBMPSite* pBMPSite = FindBMPSite(strBMPID);
				if (pBMPSite == NULL)
				{
					strError.Format("Cannot find BMP site with ID %s", strBMPID);
					retVal = false;
					goto L01;
				}

				double* pVariable = pBMPSite->GetVariablePointer(strDecVar);
				*pVariable = atof(LPCSTR(strToken2.NextToken()));
			}

			break;
		}
	}

	if (!bBestPopId)
	{
		strError.Format("Cannot find Best Population ID %d", nBestPopId);
		retVal = false;
		goto L01;
	}
L01:
	fclose(fpin);
	return retVal;
}

bool CBMPData::ReadWeatherFile(CString strFileName)
{
	if(strFileName.GetLength() == 0)
		return false;

	FILE *fpin = NULL;
	char strLine[MAXLINE];
	bool retVal = true;
	
	// open the file for reading
	fpin = fopen (strFileName, "rt");
	if(fpin == NULL)
	{
		strError = "Cannot open file " + strFileName + " for reading";
		return false;
	}

	// get the number of records
	lRecords = GetNumberOfRecords(fpin);

	if (lRecords == 0)
		return false;

	// allocate memory
	if (pWEATHERDATA != NULL)
		delete[]pWEATHERDATA;

	pWEATHERDATA = new WEATHERDATA[lRecords];
	
	// read the data
	long ii = 0;
	double Sum = 0.0;
	while (fgets(strLine, MAXLINE, fpin) != NULL)
	{
		CString str(strLine);
		CStringToken strToken(str);
		//skip the first dummy number
		strToken.NextToken();
		//read the date
		int nYear = atoi(LPCSTR(strToken.NextToken()));
		int nMonth = atoi(LPCSTR(strToken.NextToken()));
		int nDay = atoi(LPCSTR(strToken.NextToken()));
		int nHour = atoi(LPCSTR(strToken.NextToken()));
		int nMinute = atoi(LPCSTR(strToken.NextToken()));
		if(nHour == 24)																	
		{                          
			nHour = 0;                  
			pWEATHERDATA[ii].tDATE = COleDateTime(nYear,nMonth,nDay,nHour,nMinute,0) + COleDateTimeSpan(1,0,0,0);
		}
		else
		{
			pWEATHERDATA[ii].tDATE = COleDateTime(nYear,nMonth,nDay,nHour,nMinute,0);
		}
		//read the parameter values
		double value1 = atof(LPCSTR(strToken.NextToken()));
		pWEATHERDATA[ii].lfPrec = value1;
		pWEATHERDATA[ii].lfDailyPrec = Sum;
		
		//initialize
		pWEATHERDATA[ii].bWetInt = false;

		ii++;
		if (ii >= lRecords)
			break;
	}

	//assign first record value equal to the second record
	pWEATHERDATA[0].lfPrec = pWEATHERDATA[1].lfPrec;

	fclose(fpin);
	return retVal;
}


long CBMPData::GetNumberOfRecords(FILE *fp)
{
	char strLine[MAXLINE];
	long nNumRecords = 0;

	while (fgets (strLine, MAXLINE, fp) != NULL)
    {
		CString str(strLine);
		str.MakeLower();
		// skip anoter line once found "Date/time"
		if(str.Find("date/time") != -1)
		{
			for(int i=0; i<1; i++)
			{
				if(fgets(strLine, MAXLINE, fp) == NULL)
				{
					AfxMessageBox("Check weather file");
					return -1;
				}
			}
			break;	
		}
	}

	long nStart = ftell (fp);
    // scan the file to see how many records are in the file
	while (fgets(strLine, MAXLINE, fp) != NULL)
    {
		CString str(strLine);
		str.TrimLeft();
		str.TrimRight();
		if(str.GetLength() < 3)
			continue;
		nNumRecords++;
	}

    fseek (fp, nStart, SEEK_SET);
	return nNumRecords;
}

long CBMPData::FindDataIndex(COleDateTime tCurrent)
{
	if(pWEATHERDATA == NULL)
		return -1;

	//find the year, month, day, hour for the current date
	int yr1 = tCurrent.GetYear();
	int mo1 = tCurrent.GetMonth();
	int dy1 = tCurrent.GetDay();
	int hr1 = tCurrent.GetHour();

	for(long i=0; i<lRecords; i++)
	{
		int yr2 = pWEATHERDATA[i].tDATE.GetYear();
		int mo2 = pWEATHERDATA[i].tDATE.GetMonth();
		int dy2 = pWEATHERDATA[i].tDATE.GetDay();
		int hr2 = pWEATHERDATA[i].tDATE.GetHour();

		//if(tCurrent <= pWEATHERDATA[i].tDATE)
		if((yr1 == yr2) && (mo1 == mo2) && (dy1 == dy2) && (hr1 == hr2))
		{
			return i;
		}
	}

	//can not find the data
	return -1;	
}

bool CBMPData::MarkWetIntervals(COleDateTime tStart,COleDateTime tEnd)
{
	if(pWEATHERDATA == NULL)
		return false;

    //Using rainfall data to define a wet interval
    //Summarizes average daily fecal exceedences over a wet interval

    int NumInt = 0;		//number of wet intervals
    int TSperDay = 24;	//number of timesteps per day
    int RainInt = 24;	//number of timesteps per rain interval
    int DryInt = 72;	//number of timesteps for three dry days following end of rainfall
    bool Raining = false;
	bool WetInt = false;
    bool MarkEnd = false;
    double RainLim = 0.1;	//rainfall threshold (inch) for total rain measured in rainint
	double lfTolerance = 1.0E-9;
	queue<double> qPrec;

	COleDateTime IntStart,IntEnd;
	COleDateTimeSpan tspan;
	CList<COleDateTime,COleDateTime> WetList;

	//find the start date index
	lStartIndex = FindDataIndex(tStart);
	if (lStartIndex == -1)
		return false;

	//find the end date index
	lEndIndex = FindDataIndex(tEnd);
	if (lEndIndex == -1)
		return false;
	
	double value1 = 0.0;
	int qsize = 24;	// hourly time step 
	while (qPrec.size() != qsize)
		qPrec.push(value1);
		
	double lfSum = 0.0;
	for(long i=lStartIndex; i<=lEndIndex; i++)
	{
		double value0 = qPrec.front();
		lfSum -= value0;
		if (lfSum < lfTolerance) 
			lfSum = 0.0;
		value1 = pWEATHERDATA[i].lfPrec;
		lfSum += value1;
		pWEATHERDATA[i].lfDailyPrec = lfSum;
		qPrec.pop();
		qPrec.push(value1);
		
		//mark when rain starts
        if (pWEATHERDATA[i].lfPrec > 0 && !Raining && !WetInt) 
		{
            Raining = true;
            IntStart = pWEATHERDATA[i].tDATE;
        }
        
		//mark wet interval
        if ((pWEATHERDATA[i].lfDailyPrec + lfTolerance) >= RainLim && !WetInt)
		{
            WetInt = true;
            IntEnd = pWEATHERDATA[i].tDATE + COleDateTimeSpan(DryInt/TSperDay,0,0,0);
		}
        
 		//If new rainfall occurs within the same 24-hour period, push back IntEnd
        if ((pWEATHERDATA[i].lfDailyPrec + lfTolerance) >= RainLim && pWEATHERDATA[i].lfPrec > 0)
		{
            MarkEnd = false;
		}
        
        //If the sum of rainfall over RainInt timesteps exceeds RainLim then mark as WetInt
        if ((pWEATHERDATA[i].lfDailyPrec + lfTolerance) >= RainLim && pWEATHERDATA[i].lfPrec == 0 && !MarkEnd)
		{
            WetInt = true;
            MarkEnd = true;
            IntEnd = pWEATHERDATA[i].tDATE + COleDateTimeSpan(DryInt/TSperDay-1,23,0,0);
        }
        
		//turn off raining flag if rainvol = 0 and interval is not marked as a wet interval
        if (pWEATHERDATA[i].lfDailyPrec == 0 && !WetInt)
 		{
            Raining = false;
		}
        
        //reset WetInt flag and mark interval end dates if date > IntEnd
        if (WetInt && pWEATHERDATA[i].tDATE > IntEnd)
		{
            //reset flags
            Raining = false;
            WetInt = false;
            MarkEnd = false;
            
            //this section writes the start and end dates to an array
            NumInt = NumInt + 1;
            WetList.AddTail(IntStart);
            WetList.AddTail(IntEnd);

			//get the number of wet intervals
			tspan = IntEnd - IntStart;
			lfWetDays += tspan.GetTotalDays();

			//assign the wet interval
			long lStart = FindDataIndex(IntStart);
			if (lStart == -1)
				return false;

			long lEnd = FindDataIndex(IntEnd);
			if (lEnd == -1)
				return false;

			for (long j=lStart; j<=lEnd; j++)
			{
				pWEATHERDATA[j].bWetInt = true;
				nWetInt++;
			}
		}

		// check if it is end of simulation
		if (WetInt && i==lEndIndex)
		{
            //reset flags
            Raining = false;
            WetInt = false;
            MarkEnd = false;
            
            //this section writes the start and end dates to an array
            NumInt = NumInt + 1;
            WetList.AddTail(IntStart);
            WetList.AddTail(tEnd);

			//get the number of wet intervals
			tspan = tEnd - IntStart;
			lfWetDays += tspan.GetTotalDays();
			
			//assign the wet interval
			long lStart = FindDataIndex(IntStart);
			if (lStart == -1)
				return false;

			long lEnd = FindDataIndex(tEnd);
			if (lEnd == -1)
				return false;

			for (long j=lStart; j<=lEnd; j++)
			{
				pWEATHERDATA[j].bWetInt = true;
				nWetInt++;
			}
		}
	}

	//assign the total wet periods
	nWetPeriod = NumInt;

	//check the count in list
	int nCount = WetList.GetCount();

	if (nCount > 0)
	{
		pWetPeriod = new COleDateTime[nCount];

		int nIndex = 0;
		POSITION pos = WetList.GetHeadPosition();
		while (pos)
			pWetPeriod[nIndex++] = WetList.GetNext(pos);
	}

	// Release the memory
	WetList.RemoveAll();
	while (!qPrec.empty())
		qPrec.pop();

	return true;
}

void CBMPData::ClearCheckedFlag()
{
	CBMPSite* pBMPSite;
	POSITION pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		pBMPSite->m_bChecked = false;
	}
}

// This function will be called recursively for each downstream bmp site
// until all are checked. This function is developed to be able to handle
// splitter (one bmp site flows to more than one downstream bmps site)
bool CBMPData::RoutingCycleExist(CBMPSite* pBMPSite)
{
	if (pBMPSite->m_bChecked)
		return false;

	CBMPSite* pBMPSiteDown;

	POSITION pos = pBMPSite->m_dsbmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		DS_BMPSITE* pDS = (DS_BMPSITE*) pBMPSite->m_dsbmpsiteList.GetNext(pos);
		pBMPSiteDown = pDS->m_pDSBMPSite;
		if (pBMPSiteDown == NULL || pBMPSiteDown->m_bChecked)
			continue;
		if (RoutingCycleExist(pBMPSiteDown))
			return true;
	}

	pBMPSite->m_bChecked = true;
	return false;
}

// This function will be called recursively for each upstream bmp site
// until all are checked. This function is developed to be able to handle
// splitter (one bmp site flows to more than one downstream bmps site)
void CBMPData::AddRouteNode(CBMPSite* pBMPSite)
{
	if (pBMPSite->m_bChecked)
		return;

	US_BMPSITE* pUS;
	POSITION pos = pBMPSite->m_usbmpsiteList.GetHeadPosition();

	while (pos != NULL)
	{
		pUS = (US_BMPSITE*) pBMPSite->m_usbmpsiteList.GetNext(pos);
		AddRouteNode(pUS->m_pUSBMPSite);
	}

	pBMPSite->m_bChecked = true;
	routeList.AddTail(pBMPSite);
}

bool CBMPData::PrepareDataForModel()
{
	_setmaxstdio(2048);

	//validate the conduit
	Validate_Conduit(nBMPC);

	if (!ProcessTransportData())	
		return false;

	if (!ProcessPollutantData())
		return false;

	CBMPSite *pBMPSite, *pBMPSiteDown;
	CLandUse *pLU;
	CSiteLandUse *pSiteLU;
	US_BMPSITE* pUS;
	ADJUSTABLE_PARAM* pAP;
	CSitePointSource *pSitePS;
	POSITION pos, pos1;

	// set checked status of all bmp site to false
	ClearCheckedFlag();

	// check if any cycle exists in the BMP site routing network
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (RoutingCycleExist(pBMPSite))
		{
			strError = "Routing cycle exists in the network";
			return false;
		}
	}

	// load time series data for each landuse type for external land simulation control
	if (nLandSimulation == 0)
	{
		pos = luList.GetHeadPosition();
		while (pos != NULL)
		{
			pLU = (CLandUse*) luList.GetNext(pos);
			pLU->LoadLanduseTSData(startDate, endDate, polmultiplier);
		}

		pLU = (CLandUse*) luList.GetHead();
		if (pLU == NULL)
		{
			strError = "No landuse information is loaded.";
			return false;
		}
	}
	else //internal land simulation 
	{
		//read watershed output
		pos = bmpsiteList.GetHeadPosition();
		while (pos != NULL)
		{
			pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
			CString BMPSiteID = pBMPSite->m_strID;

			//load PreLU timeseries data
			int landfg = 0;
			pBMPSite->LoadWatershedTSData(BMPSiteID, startDate, endDate, strPreLUFileName,
					             polmultiplier, landfg);
			//load MixLU timeseries data
			landfg = 1;
			pBMPSite->LoadWatershedTSData(BMPSiteID, startDate, endDate, strMixLUFileName,
					             polmultiplier, landfg);
		}
	}

	// load time series data for point sources
	int nSitePS = sitepsList.GetCount();
	if (nSitePS > 0)
	{
		pos = sitepsList.GetHeadPosition();
		while (pos != NULL)
		{
			pSitePS = (CSitePointSource*) sitepsList.GetNext(pos);
			//load point source timeseries data
			pSitePS->LoadPointsourceTSData(startDate, endDate, polmultiplier);
		}
	}

	// load time series data for tradeoff curve
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);

		if (pBMPSite->m_nBreakPoints > 0)
		{
			//read breakpoint cost data
			pBMPSite->ReadTradeOffCurveCosts();
			
			// check if it is not a decision variable
			if (pBMPSite->m_nBreakPoints == 1)
			{
				//load break point timeseries data for post-dev condition
				int BPindex = 0;
				if(!pBMPSite->LoadTradeOffCurveData(BPindex, startDate, endDate))
				{
					CString strErr;
					strErr.Format("Check Cost-Effectiveness Curve Data for BMPSite ID: %s", pBMPSite->m_strID);
					AfxMessageBox(strErr);
					return false;
				}  
				//load break point timeseries data for pre-dev condition
				BPindex = 1;
				if(!pBMPSite->LoadTradeOffCurveData(BPindex, startDate, endDate))
				{
					CString strErr;
					strErr.Format("Check Cost-Effectiveness Curve Data for BMPSite ID: %s", pBMPSite->m_strID);
					AfxMessageBox(strErr);
					return false;
				}  
				//load break point timeseries data for init condition
				BPindex = 2;
				if(!pBMPSite->LoadTradeOffCurveData(BPindex, startDate, endDate))
				{
					CString strErr;
					strErr.Format("Check Cost-Effectiveness Curve Data for BMPSite ID: %s", pBMPSite->m_strID);
					AfxMessageBox(strErr);
					return false;
				}  
				//load break point timeseries data
				BPindex = 3;
				if(!pBMPSite->LoadTradeOffCurveData(BPindex, startDate, endDate))
				{
					CString strErr;
					strErr.Format("Check Cost-Effectiveness Curve Data for BMPSite ID: %s", pBMPSite->m_strID);
					AfxMessageBox(strErr);
					return false;
				}  
			}
		}
	}

	// load climate timeseries data
	if (nETflag > 0)
		LoadClimateTSData(startDate, endDate);

	// associate site landuse to each bmp site
	if (nLandSimulation == 0)
	{
		pos = siteluList.GetHeadPosition();
		while (pos != NULL)
		{
			pSiteLU = (CSiteLandUse*) siteluList.GetNext(pos);
			pBMPSite = pSiteLU->m_pBMPSite;
			if (pBMPSite != NULL)
				pBMPSite->m_siteluList.AddTail(pSiteLU);
		}
	}

	// associate site pointsource to each bmp site
	if (nSitePS > 0)
	{
		pos = sitepsList.GetHeadPosition();
		while (pos != NULL)
		{
			pSitePS = (CSitePointSource*) sitepsList.GetNext(pos);
			pBMPSite = pSitePS->m_pBMPSite;
			if (pBMPSite != NULL)
				pBMPSite->m_sitepsList.AddTail(pSitePS);
		}
	}

	// associate upstream bmpsites to each bmp site
	// and determine the most down bmp sites
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		pos1 = pBMPSite->m_dsbmpsiteList.GetHeadPosition();
		while (pos1 != NULL)
		{
			DS_BMPSITE* pDS = (DS_BMPSITE*) pBMPSite->m_dsbmpsiteList.GetNext(pos1);
			pBMPSiteDown = pDS->m_pDSBMPSite;
			if (pBMPSiteDown != NULL)
			{
				US_BMPSITE* pUS = new US_BMPSITE;
				pUS->m_nOutletType = pDS->m_nOutletType;
				pUS->m_pUSBMPSite = pBMPSite;
				pBMPSiteDown->m_usbmpsiteList.AddTail(pUS);
			}
		}
	}

	// set checked status of all bmp site to false
	ClearCheckedFlag();

	// create routing list of bmp sites for model calculation in the correct sequence
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		AddRouteNode(pBMPSite);
	}

	//calculate the accumulated design drainage area
	pos = routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) routeList.GetNext(pos);
		pBMPSite->m_lfAccDArea = pBMPSite->m_lfSiteDArea;	// acre 

		pos1 = pBMPSite->m_usbmpsiteList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pUS = (US_BMPSITE*) pBMPSite->m_usbmpsiteList.GetNext(pos1);
			if (pUS->m_nOutletType == TOTAL || pUS->m_nOutletType == ORIFICE_CHANNEL)
				pBMPSite->m_lfAccDArea += pUS->m_pUSBMPSite->m_lfAccDArea;	// acre
		}
/*
		//check if design drainage area is greater than 0.001 acres
		if (pBMPSite->m_lfDDarea > 0.001)
		{
			pBMPSite->m_lfBMPUnit = floor(pBMPSite->m_lfAccDArea / pBMPSite->m_lfDDarea + 0.5);
			if (pBMPSite->m_lfBMPUnit < 1)
				pBMPSite->m_lfBMPUnit = 1;

			//check the decision variable
			pos1 = pBMPSite->m_adjustList.GetHeadPosition();
			while (pos1 != NULL)
			{
				pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
				if (pAP->m_strVariable.CompareNoCase("NUMUNIT") == 0)
					pAP->m_lfTo = pBMPSite->m_lfBMPUnit;
			}
		}
*/
	}

	return true;
}

bool CBMPData::OpenOutputFiles(const CString& runID)
{
	int nIndex = 1;
	POSITION pos;
	CBMPSite *pBMPSite = NULL;
	CString strFilePath;
	FILE *fp = NULL;

	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{			          
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_factorList.GetCount() > 0)
		{
			strFilePath.Format("%s_%s_%s.out", runID, pBMPSite->m_strName, pBMPSite->m_strID);
			strFilePath = strOutputDir + strFilePath;
			fp = fopen(LPCSTR(strFilePath), "wt");
			if(fp == NULL)
			{
				strError = "Cannot open file " + strFilePath + " for writing.";
				return false;
			}

			WriteFileHeader(fp, nPollutant);
			pBMPSite->m_fileOut = fp; // keep file open for outputing, need to be closed by calling CloseOutputFiles
		}
	}

	return true;
}

bool CBMPData::CloseOutputFiles()
{
	POSITION pos;
	CBMPSite *pBMPSite = NULL;

	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{			          
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_fileOut != NULL)
		{
			fclose(pBMPSite->m_fileOut);
			pBMPSite->m_fileOut = NULL;
		}
	}

	return true;
}

//create input files for VFSMOD
bool CBMPData::WriteVFSMODFiles(int nRunMode,CString strID,CString& strVFSCall)
{
	CBMPSite* pBMPSite = FindBMPSite(strID);
	BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;
	FILE *fp = NULL;

	COleDateTime tStart = startDate;
	COleDateTime tEnd = endDate;
	COleDateTimeSpan span0 = tEnd - tStart;

	long nSimHrs = (long)span0.GetTotalHours() + 24;
	CString strFileName,strInputFilePath,strOutputFilePath,strErr;

	if (nRunMode == RUN_PREDEV)
		strFileName.Format("PreDev_%s", strID);
	else
		strFileName.Format("PostDev_%s", strID);

	strInputFilePath = strOutputDir + "VFSMOD_input\\" + strFileName;
	strOutputFilePath = strOutputDir + "VFSMOD_output\\" + strFileName;
	strVFSCall = strOutputDir + "VFSMOD_input\\" + strFileName + ".prj";

	// ++++++++ Open a file called vfsinput.ikw ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".ikw"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".ikw"));
		AfxMessageBox(strErr);
		return false;
	}

	fprintf(fp, "%s\n", pBMP->m_strName);
	fprintf(fp, "%lf\n", pBMP->m_lfWidth);
	fprintf(fp, "%lf %d %lf %lf %d %d %d %d\n", pBMP->m_lfLength,nN,lfTHETAW,lfCR,nMAXITER,nNPOL,nIELOUT,nKPG);
	fprintf(fp, "%d\n",pBMP->m_nSegments);

	for (long i=0; i<pBMP->m_nSegments; i++)
	{
		fprintf(fp,"%lf %lf %lf\n", pBMP->m_pSEGMENT_D[i].m_lfSX,pBMP->m_pSEGMENT_D[i].m_lfRNA,pBMP->m_pSEGMENT_D[i].m_lfSOA);
	}

	fclose(fp);

	// ++++++++ Open a file called vfsinput.igr ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".igr"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".igr"));
		AfxMessageBox(strErr);
		return false;
	}

	fprintf(fp, "%lf %lf %lf %lf %d\n", pBMP->m_lfSS,pBMP->m_lfVN,pBMP->m_lfH,pBMP->m_lfVN2,pBMP->m_nICO);
	fclose(fp);

	// ++++++++ Open a file called vfsinput.iso ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".iso"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".iso"));
		AfxMessageBox(strErr);
		return false;
	}

	fprintf(fp, "%lf %lf %lf %lf %lf %lf\n", pBMP->m_lfVKS,pBMP->m_lfSAV,
		         pBMP->m_lfOS,pBMP->m_lfOI,pBMP->m_lfSM,pBMP->m_lfSCHK);
	fclose(fp);

	// ++++++++ Open a file called vfsinput.irn ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".irn"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".irn"));
		AfxMessageBox(strErr);
		return false;
	}

	// calculate the number of points and peak value of rainfall (m/sec)
	long nNRAIN = nSimHrs;
	double lfRPEAK = 0.0;	//assume no rainfall input to the bufferstrip
	double lfRNValue = 0.0;

	fprintf(fp, "%d %lf\n", nNRAIN, lfRPEAK);
	
	for (i=0; i<nNRAIN; i++)
	{
		fprintf(fp,"%e %e\n",i*3600.0,lfRNValue);//sec, m/sec
	}
	fclose(fp);

	// ++++++++ Open a file called vfsinput.iro ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".iro"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".iro"));
		AfxMessageBox(strErr);
		return false;
	}

	long nNBCROFF = nSimHrs;
	double lfBCROPEAK = GetPeakFlowRate(nRunMode, nNBCROFF, strID)*0.02831685;//ft3/s->m3/s
	double lfROFlow = 0.0;

	fprintf(fp, "%lf %lf\n", pBMP->m_lfWidth,pBMP->m_lfLength);//m,m
	fprintf(fp, "%d %e\n", nNBCROFF, lfBCROPEAK);//number of blocks, m3/s
		
	for (i=0; i<nNBCROFF; i++)
	{
		//check the landsimulation option
		if (nLandSimulation == 0)
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum]*pBMPSite->m_lfSiteDArea * 3630.0 / 3600.0;	// cfs
			}
			else
			{
				POSITION pos;
				CSiteLandUse *pSiteLU;
				lfROFlow = 0.0;

				pos = pBMPSite->m_siteluList.GetHeadPosition();
				while (pos != NULL)
				{
					pSiteLU = (CSiteLandUse*) pBMPSite->m_siteluList.GetNext(pos);
					lfROFlow += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum]*pSiteLU->m_lfArea * 3630.0 / 3600.0;	// cfs
				}
			}
		}
		else	//internal simulation 
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum] / 3600.0;	// cfs
			}
			else
			{
				lfROFlow = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum] / 3600.0;	// cfs
			}
		}

		fprintf(fp,"%e %e\n",i*3600.0,lfROFlow*0.02831685);//sec, m3/s
	}

	fclose(fp);


	// ++++++++ Open a file called vfsinput.isd ++++++++
	fp = fopen(LPCSTR(strInputFilePath + ".isd"), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strInputFilePath + ".isd"));
		AfxMessageBox(strErr);
		return false;
	}

	double lfCI=50.0; //Sediment concentration (g/cm3)
	fprintf(fp, "%d %lf %lf %lf \n", pBMP->m_nNPART[0],pBMP->m_lfCOARSE[0],
		lfCI,pBMP->m_lfPOR[0]);
	fprintf(fp, "%lf %lf\n", pBMP->m_lfDP[0],pBMP->m_lfSG[0]);

	//Modify the above line - Remove CI and add sediment concentrations below

/*	double lfSediment[3]; //sand, silt, and clay concentration
	
	for (i=0; i<nNBCROFF; i++)
	{
		//check the landsimulation option
		if (nLandSimulation == 0)
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum]*pBMPSite->m_lfSiteDArea * 3630.0;	// ft3/hr
				for (int j=0; j<nPollutant; j++)
				{
					if (m_pPollutant[j].m_nSedfg == TSS)
					{
						// need to split the TSS into sand, silt, and clay
						double SED_FR[3];
						SED_FR[0] = pBMPSite->m_preLU->m_lfsand_fr;
						SED_FR[1] = pBMPSite->m_preLU->m_lfsilt_fr;
						SED_FR[2] = pBMPSite->m_preLU->m_lfclay_fr;
						for (int k=0; k<3; k++)
						{
							lfSediment[k] = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea*SED_FR[k];	// lbs/hr
							lfSediment[k] /= lfROFlow;	//lb/ft3
						}
					}
					else if (m_pPollutant[j].m_nSedfg == SAND)
					{
						lfSediment[0] = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea;	// lbs/hr
						lfSediment[0] /= lfROFlow;	//lb/ft3
					}
					else if (m_pPollutant[j].m_nSedfg == SILT)
					{
						lfSediment[1] = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea;	// lbs/hr
						lfSediment[1] /= lfROFlow;	//lb/ft3
					}
					else if (m_pPollutant[j].m_nSedfg == CLAY)
					{
						lfSediment[2] = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum+1+j]*pBMPSite->m_lfSiteDArea;	// lbs/hr
						lfSediment[2] /= lfROFlow;	//lb/ft3
					}
				}
			}
			else
			{
				POSITION pos;
				CSiteLandUse *pSiteLU;

				for (int k=0; k<3; k++)
					lfSediment[k] = 0.0;

				pos = pBMPSite->m_siteluList.GetHeadPosition();
				while (pos != NULL)
				{
					pSiteLU = (CSiteLandUse*) pBMPSite->m_siteluList.GetNext(pos);
					lfROFlow = pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum]*pSiteLU->m_lfArea * 3630.0;	// ft3/hr
					for (int j=0; j<nPollutant; j++)
					{
						if (m_pPollutant[j].m_nSedfg == TSS)
						{
							// need to split the TSS into sand, silt, and clay
							double SED_FR[3];
							SED_FR[0] = pSiteLU->m_pLU->m_lfsand_fr;
							SED_FR[1] = pSiteLU->m_pLU->m_lfsilt_fr;
							SED_FR[2] = pSiteLU->m_pLU->m_lfclay_fr;
							for (k=0; k<3; k++)
							{
								lfSediment[k] += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea*SED_FR[k] / lfROFlow;	// lbs/ft3
							}
						}
						else if (m_pPollutant[j].m_nSedfg == SAND)
						{
							lfSediment[0] += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea / lfROFlow;	// lbs/ft3
						}
						else if (m_pPollutant[j].m_nSedfg == SILT)
						{
							lfSediment[1] += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea / lfROFlow;	// lbs/ft3
						}
						else if (m_pPollutant[j].m_nSedfg == CLAY)
						{
							lfSediment[2] += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum+1+j]*pSiteLU->m_lfArea / lfROFlow;	// lbs/ft3
						}
					}
				}
			}
		}
		else	//internal simulation 
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum];	// ft3/hr
				for (int j=0; j<nNWQ; j++)
				{
					if (nSedflag[j] == SAND)
					{
						lfSediment[0] = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[0] /= lfROFlow;	//lb/ft3
					}
					else if (nSedflag[j] == SILT)
					{
						lfSediment[1] = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[1] /= lfROFlow;	//lb/ft3
					}
					else if (nSedflag[j] == CLAY)
					{
						lfSediment[2] = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[2] /= lfROFlow;	//lb/ft3
					}
				}
			}
			else
			{
				lfROFlow = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum];	// ft3/hr
				for (int j=0; j<nNWQ; j++)
				{
					if (nSedflag[j] == SAND)
					{
						lfSediment[0] = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[0] /= lfROFlow;	//lb/ft3
					}
					else if (nSedflag[j] == SILT)
					{
						lfSediment[1] = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[1] /= lfROFlow;	//lb/ft3
					}
					else if (nSedflag[j] == CLAY)
					{
						lfSediment[2] = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum+1+j];	// lbs/hr
						lfSediment[2] /= lfROFlow;	//lb/ft3
					}
				}
			}
		}

		fprintf(fp,"%e %e %e %e\n",i*3600.0,lfSediment[0],lfSediment[1],lfSediment[2]);
	}
*/
	fclose(fp);

	// ++++++++ Create VFSMOD project file vfsinput.prj ++++++++
	fp = fopen(LPCSTR(strVFSCall), "wt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open input file: %s", LPCSTR(strVFSCall));
		AfxMessageBox(strErr);
		return false;
	}

	fprintf(fp, "ikw=" + strInputFilePath  + ".ikw\n");
	fprintf(fp, "iso=" + strInputFilePath  + ".iso\n");
	fprintf(fp, "igr=" + strInputFilePath  + ".igr\n");
	fprintf(fp, "isd=" + strInputFilePath  + ".isd\n");
	fprintf(fp, "irn=" + strInputFilePath  + ".irn\n");
	fprintf(fp, "iro=" + strInputFilePath  + ".iro\n");
	fprintf(fp, "og1=" + strOutputFilePath + ".og1\n");
	fprintf(fp, "og2=" + strOutputFilePath + ".og2\n");
	fprintf(fp, "ohy=" + strOutputFilePath + ".ohy\n");
	fprintf(fp, "osm=" + strOutputFilePath + ".osm\n");
	fprintf(fp, "osp=" + strOutputFilePath + ".osp\n");

	fclose(fp);

	return true;

}

double CBMPData::GetPeakFlowRate(int nRunMode,long nNBCROFF,CString strID)
{
	double lfROFlow = 0.0;
	double lfBCROPEAK = 0.0;

	POSITION pos;
	CSiteLandUse *pSiteLU = NULL;
	CBMPSite* pBMPSite = FindBMPSite(strID);

	for (long i=0; i<nNBCROFF; i++)
	{
		//check the landsimulation option
		if (nLandSimulation == 0)
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_preLU->m_pData[i*pBMPSite->m_preLU->m_nQualNum]*pBMPSite->m_lfSiteDArea * 3630.0 / 3600.0;	// cfs
			}
			else
			{
				lfROFlow = 0.0;

				pos = pBMPSite->m_siteluList.GetHeadPosition();
				while (pos != NULL)
				{
					pSiteLU = (CSiteLandUse*) pBMPSite->m_siteluList.GetNext(pos);
					lfROFlow += pSiteLU->m_pLU->m_pData[i*pSiteLU->m_pLU->m_nQualNum]*pSiteLU->m_lfArea * 3630.0 / 3600.0;	// cfs
				}
			}
		}
		else	//internal simulation 
		{
			if (nRunMode == RUN_PREDEV)
			{
				lfROFlow = pBMPSite->m_pDataPreLU[i*pBMPSite->m_nQualNum] / 3600.0;	// cfs
			}
			else
			{
				lfROFlow = pBMPSite->m_pDataMixLU[i*pBMPSite->m_nQualNum] / 3600.0;	// cfs
			}
		}

		if (lfBCROPEAK < lfROFlow)
			lfBCROPEAK = lfROFlow; 
	}

	return lfBCROPEAK;
}

bool CBMPData::ReadVFSMODFiles(int nRunMode,CString strID)
{
	FILE *fp = NULL;
	CBMPSite* pBMPSite = FindBMPSite(strID);
	BMP_D* pBMP = (BMP_D*) pBMPSite->m_pSiteProp;

	CString strFileName,strOutputFilePath,strErr;

	if (nRunMode == RUN_PREDEV)
		strFileName.Format("PreDev_%s", strID);
	else
		strFileName.Format("PostDev_%s", strID);

	strOutputFilePath = strOutputDir + "VFSMOD_output\\" + strFileName;

	// ++++++++ Open a file (flow output) called vfsout.og2 ++++++++
	fp = fopen(LPCSTR(strOutputFilePath + ".og2"), "rt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open output file: %s", LPCSTR(strOutputFilePath + ".og2"));
		AfxMessageBox(strErr);
		return false;
	}

	//skip few lines till there are ----

	//qout is the last element

	fclose(fp);

	// ++++++++ Open a file (sediment output) called vfsout.og1 ++++++++
	fp = fopen(LPCSTR(strOutputFilePath + ".og1"), "rt");

	if (fp == NULL)
	{
		strErr.Format("Unable to open output file: %s", LPCSTR(strOutputFilePath + ".og1"));
		AfxMessageBox(strErr);
		return false;
	}

	//skip few lines till there are ----

	//gso is 10th value cum.gso is 14th

	fclose(fp);

	return true;

}

bool CBMPData::RunVFSMOD(int nRunMode)
{
	POSITION pos;
	CString strVFSCall,strErr;
	CBMPSite* pBMPSite;

	pos = routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) routeList.GetNext(pos);
		if (pBMPSite->m_nBMPClass == CLASS_D)
		{
			//create input files for VFSMOD
			if(!WriteVFSMODFiles(nRunMode,pBMPSite->m_strID,strVFSCall))
			{
				strErr.Format("Error for writing VFSMOD input files for BMPSite: %s",
					           pBMPSite->m_strID);
				AfxMessageBox(strErr);
				return false;
			}

			//strVFSCall.Format("E:\\CVS\\VFSMOD_DLL\\Sample\\sample.prj");

			//call VFSMOD to run buffer strip simulation
			if (!CallVFSMOD((LPCSTR)strVFSCall))
			{
				strErr.Format("check input files for VFSMOD for BMPSite: %s",
					           pBMPSite->m_strID);
				AfxMessageBox(strErr);
				return false;
			}
		}
	}

	return true;
}

void CBMPData::WriteFileHeader(FILE *fp, int NWQ)
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

		fprintf(fp, "TT Volume          BMP volume (ft3)\n");
		fprintf(fp, "TT Stage           Water depth (ft)\n");
		fprintf(fp, "TT Inflow_t        Total inflow (cfs)\n");
		fprintf(fp, "TT Outflow_w       Weir outflow (cfs)\n");
		fprintf(fp, "TT Outflow_o       Orifice or channel outflow (cfs)\n");
		fprintf(fp, "TT Outflow_ud      Underdrain outflow (cfs)\n");
		fprintf(fp, "TT Outflow_ut      Untreated (bypass) outflow (cfs)\n");
		fprintf(fp, "TT Outflow_t       Total outflow (cfs)\n");
		fprintf(fp, "TT Infiltration    Infiltration (cfs)\n");
		fprintf(fp, "TT Perc            Percolation to underdrain storage (cfs)\n");
		fprintf(fp, "TT AET             Evapotranspiration (cfs)\n");
		fprintf(fp, "TT Seepage         Seepage to groundwater (cfs)\n");
//		fprintf(fp, "TT USstorage       Upper Soil storage (ft3)\n");
//		fprintf(fp, "TT UDstorage       Under Drain storage (ft3)\n");

	for(int i=0; i<NWQ; i++)
    {
		fprintf(fp, "TT Mass_in_%d       Mass entering the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Mass_w_%d        Mass leaving (weir outflow) the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Mass_o_%d        Mass leaving (orifice outflow) the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Mass_ud_%d       Mass leaving (underdrain outflow) the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Mass_ut_%d       Mass bypassing (untreated) the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Mass_out_%d      Mass leaving the BMP (lbs)\n", i+1);
		fprintf(fp, "TT Conc_%d          Total outflow concentration (mg/l)\n", i+1);
	}
	fputs("TT    \n", fp);
	fputs("TT-----------------------------------------------------------------------------------------\n",fp);
	fputs("TT    Date/time                      Values\n", fp);
	fflush(fp);
 }				

CString CBMPData::GetRoutingOrder()
{
	CString strValue, strOrder = "Routing Sequence: ";
	int nIndex;
	CBMPSite* pBMPSite;

	POSITION pos = routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) routeList.GetNext(pos);
		nIndex = ::FindObIndexFromList(bmpsiteList, pBMPSite) + 1;
		strValue.Format(" %d", nIndex);
		strOrder += strValue;
	}

	return strOrder;
}

bool CBMPData::ProcessPollutantData()
{
	int i, j, nIndex;

	if (nSedflag != NULL)	
		delete []nSedflag;
	nSedflag = new int[nNWQ];

	nIndex = 0;
	for (i=0; i<nPollutant; i++)
	{
		if (m_pPollutant[i].m_nSedfg == TSS)
		{
			for (j=0; j<3; j++)
				nSedflag[nIndex++] = j+1;
		}
		else
		{
			nSedflag[nIndex++] = m_pPollutant[i].m_nSedfg;
		}
	}

	if (nNWQ == nPollutant)
		return true;

	CBMPSite *pBMPSite;
	POSITION pos;

	//process decay parameters
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_pDecay == NULL)		
			continue;

		double* lfDecay = new double[nPollutant];
		for (i=0; i<nPollutant; i++)
			lfDecay[i] = pBMPSite->m_pDecay[i];

		delete []pBMPSite->m_pDecay;
		pBMPSite->m_pDecay = new double[nNWQ];

		nIndex = 0;
		for (i=0; i<nPollutant; i++)
		{
			if (m_pPollutant[i].m_nSedfg == TSS)
			{
				for (j=0; j<3; j++)
				{
					pBMPSite->m_pDecay[nIndex] = lfDecay[i];
					nIndex++;
				}
			}
			else
			{
				pBMPSite->m_pDecay[nIndex] = lfDecay[i];
				nIndex++;
			}
		}

		delete []lfDecay;
	}

	//process k' parameters
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_pK == NULL)			
			continue;

		double* lfK = new double[nPollutant];
		for (i=0; i<nPollutant; i++)
			lfK[i] = pBMPSite->m_pK[i];

		delete []pBMPSite->m_pK;
		pBMPSite->m_pK = new double[nNWQ];

		nIndex = 0;
		for (i=0; i<nPollutant; i++)
		{
			if (m_pPollutant[i].m_nSedfg == TSS)
			{
				for (j=0; j<3; j++)
				{
					pBMPSite->m_pK[nIndex] = lfK[i];
					nIndex++;
				}
			}
			else
			{
				pBMPSite->m_pK[nIndex] = lfK[i];
				nIndex++;
			}
		}

		delete []lfK;
	}

	//process C* parameters
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_pCstar == NULL)
			continue;

		double* lfCstar = new double[nPollutant];
		for (i=0; i<nPollutant; i++)
			lfCstar[i] = pBMPSite->m_pCstar[i];

		delete []pBMPSite->m_pCstar;
		pBMPSite->m_pCstar = new double[nNWQ];

		nIndex = 0;
		for (i=0; i<nPollutant; i++)
		{
			if (m_pPollutant[i].m_nSedfg == TSS)
			{
				for (j=0; j<3; j++)
				{
					pBMPSite->m_pCstar[nIndex] = lfCstar[i];
					nIndex++;
				}
			}
			else
			{
				pBMPSite->m_pCstar[nIndex] = lfCstar[i];
				nIndex++;
			}
		}

		delete []lfCstar;
	}

	//process underdrain parameters
	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_pUndRemoval == NULL)
			continue;

		double* lfUndRemoval = new double[nPollutant];
		for (i=0; i<nPollutant; i++)
			lfUndRemoval[i] = pBMPSite->m_pUndRemoval[i];

		delete []pBMPSite->m_pUndRemoval;
		pBMPSite->m_pUndRemoval = new double[nNWQ];

		nIndex = 0;
		for (i=0; i<nPollutant; i++)
		{
			if (m_pPollutant[i].m_nSedfg == TSS)
			{
				for (j=0; j<3; j++)
				{
					pBMPSite->m_pUndRemoval[nIndex] = lfUndRemoval[i];
					nIndex++;
				}
			}
			else
			{
				pBMPSite->m_pUndRemoval[nIndex] = lfUndRemoval[i];
				nIndex++;
			}
		}

		delete []lfUndRemoval;
	}

	return true;
}

// load the time series data
bool CBMPData::LoadClimateTSData(COleDateTime startDate,COleDateTime endDate)
{
	int i, j;
	char strLine[MAXLINE];
	CString str;

	FILE *fpin = NULL;
	// open the file for reading
	fpin = fopen (strClimateFileName, "rt");
	if(fpin == NULL)
		return false;

	// count the time series numbers
	COleDateTimeSpan tsSpan = endDate - startDate;
	long nTSNum = (long)tsSpan.GetTotalDays() + 1;
	
    // read first data line for starting date of the time series data
	long nStart = ftell (fpin);
	fgets(strLine, MAXLINE, fpin);
    fseek(fpin, nStart, SEEK_SET);

	int year, month, day, hour, min, sec;
	str = strLine;
	CStringToken strToken3(str);
	strToken3.NextToken(); // skip the climate station name
	str = strToken3.NextToken();
	year = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	month = atoi((LPCSTR)str);

	str = strToken3.NextToken();
	day = atoi((LPCSTR)str);
	
	COleDateTime tmStart = COleDateTime(year, month, day, 0, 0, 0);
	tsSpan = startDate - tmStart;

	// calculate how many lines we need to skip
	long nSkipLineNum = (long)tsSpan.GetTotalDays();

    // skip all lines the time stamp is before the specified start date
	while (nSkipLineNum-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	double* pData = NULL;
	if(m_pDataClimate != NULL)
		delete []m_pDataClimate;
	m_pDataClimate = new double[nTSNum*m_nNum];
	pData = m_pDataClimate;
	
    // read the data
	i = 0;

    while (!feof(fpin))
    {
		// read one line
		fgets(strLine, MAXLINE, fpin);

		// get the data
		str = strLine;
		CStringToken strToken(str);
		strToken.NextToken(); // station name
		strToken.NextToken(); // year
		strToken.NextToken(); // month
		strToken.NextToken(); // day

		for(j=0; j<m_nNum; j++)
		{
			str = strToken.NextToken();
			*(pData++) = atof((LPCSTR)str);
		}

		i++;
		if (i == nTSNum)
			break;
	}

	fclose(fpin);
	return (i == nTSNum);
}

//SWMM5
bool CBMPData::ProcessTransportData()
{
	int nIndex;
	CBMPSite *pBMPSite;
	BMP_C* pBMP;
	POSITION pos;

	pos = bmpsiteList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) bmpsiteList.GetNext(pos);
		if (pBMPSite->m_nBMPClass == CLASS_C)
		{
			pBMP = (BMP_C*) pBMPSite->m_pSiteProp;
			nIndex = pBMP->m_nIndex;

			// --- compute full flow through cross section
			if ( Link[nIndex].xsect.type == DUMMY ) 
				Conduit[nIndex].beta = 0.0;
			else
			{
				if (Conduit[nIndex].roughness > 0)
					Conduit[nIndex].beta = PHI * sqrt(fabs(Conduit[nIndex].slope)) / Conduit[nIndex].roughness;
				else
					Conduit[nIndex].beta = 0.0;
			}
			Link[nIndex].qFull = Link[nIndex].xsect.sFull * Conduit[nIndex].beta;
			Conduit[nIndex].qMax = Link[nIndex].xsect.sMax * Conduit[nIndex].beta;
			// --- set value of hasLosses flag
			if ( Link[nIndex].cLossInlet  == 0 && Link[nIndex].cLossOutlet == 0 && Link[nIndex].cLossAvg == 0) 
				Conduit[nIndex].hasLosses = FALSE;
			else 
				Conduit[nIndex].hasLosses = TRUE;

			// --- initialize water quality state
			FREE(Link[nIndex].oldQual);
			FREE(Link[nIndex].newQual);
			if (pBMP->m_pTLink.oldQual != NULL) delete []pBMP->m_pTLink.oldQual;
			if (pBMP->m_pTLink.newQual != NULL) delete []pBMP->m_pTLink.newQual;

			if (nNWQ > 0)
			{
				pBMP->m_pTLink.oldQual = new float[nNWQ];
				pBMP->m_pTLink.newQual = new float[nNWQ];
				Link[nIndex].oldQual = (float *) calloc(nNWQ, sizeof(float));
				Link[nIndex].newQual = (float *) calloc(nNWQ, sizeof(float));
			}

			for (int p = 0; p < nNWQ; p++)
			{
				Link[nIndex].oldQual[p] = 0.0;
				Link[nIndex].newQual[p] = 0.0;
			}

			// copy data
			CopyLink(&Link[nIndex], &pBMP->m_pTLink, nNWQ);
			CopyConduit(&Conduit[nIndex], &pBMP->m_pTConduit);
			CopyTransect(&Transect[nIndex], &pBMP->m_pTTransect);
		}
	}

	return true;
}

void CBMPData::OutputFileHeaderForTradeOffCurve(FILE* fp)
{
	POSITION pos, pos1;
	CBMPSite* pBMPSite;
	ADJUSTABLE_PARAM* pAP;
	EVALUATION_FACTOR* pEF;
	CString strLine, strValue;
	strLine = "NSGA-II Cost-Effectiveness Curve Solutions\n";
	strLine += "BestPop#\tSolution#\tCost($)";
//	strLine += "\tTargetValue";

	pos = routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) routeList.GetNext(pos);

		pos1 = pBMPSite->m_factorList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
			strValue.Format("\t%s_%s_%d", pBMPSite->m_strID, pEF->m_strFactor, pEF->m_nCalcMode);
			strLine += strValue;
		}
	}

	pos = routeList.GetHeadPosition();
	while (pos != NULL)
	{
		pBMPSite = (CBMPSite*) routeList.GetNext(pos);

		pos1 = pBMPSite->m_adjustList.GetHeadPosition();
		while (pos1 != NULL)
		{
			pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
			strValue.Format("\t%s_%s", pBMPSite->m_strID, pAP->m_strVariable);
			strLine += strValue;
		}
	}

//	strLine += "\tValue";

	strLine += "\n\n";
	fputs(strLine, fp);
	fflush(fp);
}

void InitializeGAInfil(int nNum)
{
	for (int nIndex=0; nIndex<nNum; nIndex++)
	{
		GAInfil[nIndex].F		= 0.0;
		GAInfil[nIndex].FU		= 0.0;
		GAInfil[nIndex].FUmax	= 0.0;
		GAInfil[nIndex].IMD		= 0.0;
		GAInfil[nIndex].IMDmax	= 0.0;
		GAInfil[nIndex].Ks		= 0.0;
		GAInfil[nIndex].L		= 0.0;
		GAInfil[nIndex].S		= 0.0;
		GAInfil[nIndex].Sat		= FALSE;
		GAInfil[nIndex].T		= 0.0;
	}
	return;
}

void CopyGAInfil(TGrnAmpt* Source, TGrnAmpt* Target)
{
	// --- copy conduit properties
	Target->F		= Source->F;		
	Target->FU		= Source->FU;		
	Target->FUmax	= Source->FUmax;	
	Target->IMD		= Source->IMD;		
	Target->IMDmax	= Source->IMDmax;	
	Target->Ks		= Source->Ks;		
	Target->L		= Source->L;		
	Target->S		= Source->S;		
	Target->Sat		= Source->Sat;		
	Target->T		= Source->T;		

	return;
}

void InitializeLinkConduitTransect(int nNum)
{
	for (int nIndex=0; nIndex<nNum; nIndex++)
	{
		// --- initialize link properties
		Link[nIndex].ID = "";				    // link ID
		Link[nIndex].type = CONDUIT;			// link type code
		Link[nIndex].subIndex = -1;			    // index of link's sub-category
		Link[nIndex].rptFlag = FALSE;		    // reporting flag
		Link[nIndex].node1 = 0;				    // start node index
		Link[nIndex].node2 = 0;				    // end node index
		Link[nIndex].z1 = 0.;				    // upstrm invert ht. above node invert (ft)
		Link[nIndex].z2 = 0.;				    // downstrm invert ht. above node invert (ft)
		Link[nIndex].q0 = 0.;				    // initial flow (cfs)
		Link[nIndex].qLimit = 0.;			    // constraint on max. flow (cfs)
		Link[nIndex].cLossInlet = 0.;		    // inlet loss coeff.
		Link[nIndex].cLossOutlet = 0.;		    // outlet loss coeff.
		Link[nIndex].cLossAvg = 0.;			    // avg. loss coeff.
		Link[nIndex].hasFlapGate = FALSE;	    // true if flap gate present
		Link[nIndex].oldFlow = 0.;			    // previous flow rate (cfs)
		Link[nIndex].newFlow = 0.;			    // current flow rate (cfs)
		Link[nIndex].oldDepth = 0.;			    // previous flow depth (ft)
		Link[nIndex].newDepth = 0.;			    // current flow depth (ft)
		Link[nIndex].oldVolume = 0.;		    // previous flow volume (ft3)
		Link[nIndex].newVolume = 0.;		    // current flow volume (ft3)
		Link[nIndex].qFull = 0.;			    // flow when full (cfs)
		Link[nIndex].setting = 0.;			    // control setting
		Link[nIndex].froude = 0.;			    // Froude number
		Link[nIndex].oldQual = NULL;		    // previous quality state
		Link[nIndex].newQual = NULL;		    // current quality state
		Link[nIndex].flowClass = -1;		    // flow classification
		Link[nIndex].dqdh = 0.;				    // change in flow w.r.t. head (ft2/sec)
		Link[nIndex].direction = -1;		    // flow direction flag
		Link[nIndex].isClosed = FALSE;		    // flap gate closed flag
		Link[nIndex].xsect.type = -1;		    // type code of cross section shape
		Link[nIndex].xsect.transect = -1;	    // index of transect (if applicable)
		Link[nIndex].xsect.yFull = 0.;		    // depth when full (ft)
		Link[nIndex].xsect.wMax = 0.;		    // width at widest point (ft)
		Link[nIndex].xsect.aFull = 0.;		    // area when full (ft2)
		Link[nIndex].xsect.rFull = 0.;		    // hyd. radius when full (ft)
		Link[nIndex].xsect.sFull = 0.;		    // section factor when full (ft^4/3)
		Link[nIndex].xsect.sMax = 0.;		    // section factor at max. flow (ft^4/3)
		Link[nIndex].xsect.yBot = 0.;		    // depth of bottom section
		Link[nIndex].xsect.aBot = 0.;		    // area of bottom section
		Link[nIndex].xsect.sBot = 0.;		    // slope of bottom section
		Link[nIndex].xsect.rBot = 0.;			// radius of bottom section

		// --- initialize conduit properties
		Conduit[nIndex].length = 0.;		    // conduit length (ft)
		Conduit[nIndex].roughness = 0.;		    // Manning's n
		Conduit[nIndex].barrels = 1;		    // number of barrels
		Conduit[nIndex].modLength = 0.;		    // modified conduit length (ft)
		Conduit[nIndex].roughFactor = 0.;	    // roughness factor for DW routing
		Conduit[nIndex].slope = 0.;			    // slope
		Conduit[nIndex].beta = 0.;			    // discharge factor
		Conduit[nIndex].qMax = 0.;			    // max. flow (cfs)
		Conduit[nIndex].a1 = 0.;			    // upstream areas (ft2)
		Conduit[nIndex].a2 = 0.;			    // downstream areas (ft2)
		Conduit[nIndex].q1 = 0.;			    // upstream flows per barrel (cfs)
		Conduit[nIndex].q2 = 0.;			    // downstream flows per barrel (cfs)
		Conduit[nIndex].q1Old = 0.;			    // previous values of q1 & q2 (cfs)
		Conduit[nIndex].q2Old = 0.;			    // previous values of q1 & q2 (cfs)
		Conduit[nIndex].superCritical = FALSE;	// super-critical flow flag
		Conduit[nIndex].hasLosses = FALSE;	    // local losses flag

		// --- initialize transect properties
		Transect[nIndex].ID = "";               // section ID
		Transect[nIndex].yFull = 0.;            // depth when full (ft)
		Transect[nIndex].aFull = 0.;            // area when full (ft2)
		Transect[nIndex].rFull = 0.;            // hyd. radius when full (ft)
		Transect[nIndex].wMax = 0.;             // width at widest point (ft)
		Transect[nIndex].sMax = 0.;             // section factor at max. flow (ft^4/3)
		Transect[nIndex].aMax = 0.;             // area at max. flow (ft2)
		Transect[nIndex].roughness = 0.;        // Manning's n
		for (int i=0; i<N_TRANSECT_TBL; i++)
		{
			Transect[nIndex].areaTbl[i] = 0.0;	// table of area v. depth
			Transect[nIndex].hradTbl[i] = 0.0;	// table of hyd. radius v. depth
			Transect[nIndex].widthTbl[i] = 0.0;	// table of top width v. depth
		}
		Transect[nIndex].nTbl = N_TRANSECT_TBL; // size of geometry tablesnTbl = N_TRANSECT_TBL;
	}
	return;
}

void CopyLink(TLink* Source, TLink* Target, int nPollutant)
{
	// --- copy link properties
	Target->ID				= Source->ID;			
	Target->type			= Source->type;			
	Target->subIndex		= Source->subIndex;		
	Target->rptFlag			= Source->rptFlag;		
	Target->node1			= Source->node1;		
	Target->node2			= Source->node2;		
	Target->z1				= Source->z1;			 
	Target->z2				= Source->z2;			 
	Target->q0				= Source->q0;			
	Target->qLimit			= Source->qLimit;		
	Target->cLossInlet		= Source->cLossInlet;	
	Target->cLossOutlet		= Source->cLossOutlet;	
	Target->cLossAvg		= Source->cLossAvg;		
	Target->hasFlapGate		= Source->hasFlapGate;	
	Target->oldFlow			= Source->oldFlow;		
	Target->newFlow			= Source->newFlow;		
	Target->oldDepth		= Source->oldDepth;		
	Target->newDepth		= Source->newDepth;		
	Target->oldVolume		= Source->oldVolume;	
	Target->newVolume		= Source->newVolume;	
	Target->qFull			= Source->qFull;		
	Target->setting			= Source->setting;		
	Target->froude			= Source->froude;		
	for (int p=0; p<nPollutant; p++)
	{
		Target->oldQual[p]	= Source->oldQual[p];		
		Target->newQual[p]	= Source->newQual[p];		
	}
	Target->flowClass		= Source->flowClass;	
	Target->dqdh			= Source->dqdh;			
	Target->direction		= Source->direction;	
	Target->isClosed		= Source->isClosed;		
	Target->xsect.type		= Source->xsect.type;
	Target->xsect.transect	= Source->xsect.transect;
	Target->xsect.yFull		= Source->xsect.yFull;	
	Target->xsect.wMax		= Source->xsect.wMax;	
	Target->xsect.aFull		= Source->xsect.aFull;	
	Target->xsect.rFull		= Source->xsect.rFull;	
	Target->xsect.sFull		= Source->xsect.sFull;	
	Target->xsect.sMax		= Source->xsect.sMax;	
	Target->xsect.yBot		= Source->xsect.yBot;	
	Target->xsect.aBot		= Source->xsect.aBot;	
	Target->xsect.sBot		= Source->xsect.sBot;	
	Target->xsect.rBot		= Source->xsect.rBot;	

	return;
}

void CopyConduit(TConduit* Source, TConduit* Target)
{
	// --- copy conduit properties
	Target->length			= Source->length;		
	Target->roughness		= Source->roughness;	
	Target->barrels			= Source->barrels;		
	Target->modLength		= Source->modLength;	
	Target->roughFactor		= Source->roughFactor;	
	Target->slope			= Source->slope;		
	Target->beta			= Source->beta;			
	Target->qMax			= Source->qMax;			
	Target->a1				= Source->a1;			
	Target->a2				= Source->a2;			
	Target->q1				= Source->q1;			
	Target->q2				= Source->q2;			
	Target->q1Old			= Source->q1Old;		
	Target->q2Old			= Source->q2Old;		
	Target->superCritical	= Source->superCritical;
	Target->hasLosses		= Source->hasLosses;	

	return;
}

void CopyTransect(TTransect* Source, TTransect* Target)
{
	// --- copy transect properties
	Target->ID				= Source->ID;		
	Target->yFull			= Source->yFull;	
	Target->aFull			= Source->aFull;	
	Target->rFull			= Source->rFull;	
	Target->wMax			= Source->wMax;		
	Target->sMax			= Source->sMax;		
	Target->aMax			= Source->aMax;		
	Target->roughness		= Source->roughness;
	Target->nTbl			= Source->nTbl;		
	for (int i=0; i<N_TRANSECT_TBL; i++)
	{
		Target->areaTbl[i]	= Source->areaTbl[i];
		Target->hradTbl[i]	= Source->hradTbl[i];
		Target->widthTbl[i]	= Source->widthTbl[i];
	}

	return;
}

int FindObIndexFromList(CObList& list, CObject* ob)
{
	int nIndex = 0;
	POSITION pos = list.GetHeadPosition();
	while (pos != NULL)
	{
		if (list.GetNext(pos) == ob)
			return nIndex;

		nIndex++;
	}
	return -1;
}

void Validate_Conduit(int nNum)
{
	//conduit_validate
	for (int nIndex=0; nIndex<nNum; nIndex++)
	{
		// --- if irreg. xsection, assign transect roughness to conduit
		if ( Link[nIndex].xsect.type == IRREGULAR )
		{
			Conduit[nIndex].roughness = Transect[Link[nIndex].xsect.transect].roughness;
		}

		// --- adjust conduit offsets for partly filled circular xsection
		if ( Link[nIndex].xsect.type == FILLED_CIRCULAR )
		{
			Link[nIndex].z1 += Link[nIndex].xsect.yBot;
			Link[nIndex].z2 += Link[nIndex].xsect.yBot;
		}

		// --- compute conduit slope 
		double elev1 = Link[nIndex].z1;
		double elev2 = Link[nIndex].z2;

		if (Conduit[nIndex].length > 0)
		{
			if ( fabs(elev1 - elev2) < MIN_DELTA_Z )
			{
				Conduit[nIndex].slope = MIN_DELTA_Z / Conduit[nIndex].length;
			}
			else Conduit[nIndex].slope = (elev1 - elev2) / Conduit[nIndex].length;
		}
		else
		{
			Conduit[nIndex].slope = 0;
		}

		double lengthFactor = 1.0;

		// --- compute modified slope, roughness & roughness factor
		Conduit[nIndex].slope /= lengthFactor;
		double roughness = Conduit[nIndex].roughness / sqrt(lengthFactor);
		Conduit[nIndex].roughFactor = GRAVITY * SQR(roughness/PHI);

		// --- compute full flow through cross section
		Conduit[nIndex].beta = 0.0;
		
		if ( Link[nIndex].xsect.type != DUMMY && roughness != 0) 
		{
			Conduit[nIndex].beta = PHI * sqrt(fabs(Conduit[nIndex].slope)) / roughness;
		}

		Link[nIndex].qFull = Link[nIndex].xsect.sFull * Conduit[nIndex].beta;
		Conduit[nIndex].qMax = Link[nIndex].xsect.sMax * Conduit[nIndex].beta;

		// --- see if flow is supercritical most of time
		//     by comparing normal & critical velocities.
		//     (factor of 0.3 is for circular pipe 95% full)
		// NOTE: this factor is used for modified Kinematic Wave routing.
		double aa = Conduit[nIndex].beta / sqrt(32.2) *
			 pow(Link[nIndex].xsect.yFull, 0.1666667) * 0.3;
		if ( aa >= 1.0 ) Conduit[nIndex].superCritical = TRUE;
		else             Conduit[nIndex].superCritical = FALSE;

		// --- set value of hasLosses flag
		if ( Link[nIndex].cLossInlet  == 0.0 &&
			 Link[nIndex].cLossOutlet == 0.0 &&
			 Link[nIndex].cLossAvg    == 0.0
		   ) Conduit[nIndex].hasLosses = FALSE;
		else Conduit[nIndex].hasLosses = TRUE;
	}
	
	return;
}


