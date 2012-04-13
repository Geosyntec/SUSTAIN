// Global.cpp: implementation of the exported function for simulation
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Global.h"
#include "BMPSite.h"
#include "BMPData.h"
#include "BMPRunner.h"
#include "BMPOptimizer.h"
#include "BMPOptimizerGA.h"
#include "ProgressWnd.h"
#include "StringToken.h"

#include "./swmm5/swmm5.h"
#include <math.h>


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

// Fetch a single random number between 0.0 and 1.0
double random_perc()
{
	// since the random number returned by rand() function is within
	// the range between 0 and 0x7fff (32767), we use 10001 to preserve
	// the floating point up to 4 digits
	return rand()%10001/10000.0;
}

// Fetch a single random integer between low and high including the bounds
int random_int(int low, int high)
{
	int res = low + (int)((high-low)*random_perc());
	if (res < low)
		res = low;
	if (res > high)
		res = high;
	return res;
}

// Fetch a single random real number between low and high including the bounds
double random_real(double low, double high)
{
	return low + (high-low)*random_perc();
}

// Fetch a single random real number between low and high with fixed increments including the bounds
double random_real_with_inc(double low, double high, double inc)
{
	double res = low + int((high-low)*random_perc()/inc+0.5)*inc;
	if (res < low)
		res = low;
	if (res > high)
		res = high;
	return res;
}

extern "C" BOOL PASCAL EXPORT StartLandSimulation(char* strLandPreDevFilePath,
												  char* strLandPostDevFilePath)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// normal function body here

	CString str1, str2;
	
	str1 = strLandPreDevFilePath;
	str1.TrimLeft();
	str1.TrimRight();
	
	str2 = strLandPostDevFilePath;
	str2.TrimLeft();
	str2.TrimRight();

	// initialize progress bar
	CProgressWnd wndProgress(NULL, "SUSTAIN_MODEL", TRUE);

	// land simulation for pre-developed scenario
	int landfg = 0;
	if(str1.GetLength() > 2)
	{
		LandSimulation(landfg, strLandPreDevFilePath, &wndProgress);
		if ( ErrorCode ) return false;
	}

	// land simulation for post-developed scenario
	landfg = 1;
	if(str2.GetLength() > 2)
	{
		LandSimulation(landfg, strLandPostDevFilePath, &wndProgress);
		if ( ErrorCode ) return false;
	}

	wndProgress.DestroyWindow();
	AfxMessageBox("Land simulation completed successfully", MB_ICONINFORMATION);
	return true;
}

extern "C" BOOL PASCAL EXPORT StartSimulation(char* strInputFilePath,char* strBestPopRun)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// normal function body here

	int nReturn = TRUE;
	CString str1, str2, strFilePath;
	
	str1 = strInputFilePath;
	str1.TrimLeft();
	str1.TrimRight();
	
	str2 = strBestPopRun;
	str2.TrimLeft();
	str2.TrimRight();
	
	// initialize progress bar
	CProgressWnd wndProgress(NULL, "SUSTAIN_MODEL", TRUE);

	if(str1.GetLength() < 2)	
		return true;

//////////////////////////////////////////////////////////////////////
//SWMM5 initialize conduit parameters
    initPointers();
    setDefaults();
    // --- create hash tables for fast retrieval of objects by ID names
    createHashTables();
//////////////////////////////////////////////////////////////////////
	
	strFilePath = str1;

	CBMPData bmpData;
	CBMPRunner bmpRunner(&bmpData);

	if (!bmpData.ReadInputFile(strFilePath))
	{
		AfxMessageBox(bmpData.strError);
		nReturn = FALSE;
		goto L001;
	}

	//optional weather data
	if (bmpData.nWeatherFile == 1)
	{
		if (!bmpData.ReadWeatherFile(bmpData.strInputDir + "weather.inp"))
			return FALSE;

		if (!bmpData.MarkWetIntervals(bmpData.startDate,bmpData.endDate))
			return FALSE;
	}

	if (!bmpData.PrepareDataForModel())
	{
		AfxMessageBox(bmpData.strError);
		nReturn = FALSE;
		goto L001;
	}

	// disabled for the first release (under testing)
	// run PreDeveloped scenario for bufferstrip simulation if necessary
//	if (!bmpData.RunVFSMOD(RUN_PREDEV))	 
//	{
//		AfxMessageBox(bmpData.strError);
//		nReturn = FALSE;
//		goto L001;
//	}

	// run PostDeveloped scenario for bufferstrip simulation if necessary
//	if (!bmpData.RunVFSMOD(RUN_POSTDEV))	 
//	{
//		AfxMessageBox(bmpData.strError);
//		nReturn = FALSE;
//		goto L001;
//	}
	
	if (!bmpData.OpenOutputFiles("Init"))// open file for time series 
	{
		AfxMessageBox(bmpData.strError);
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.OpenOutputFiles("Init", bmpData.nRunOption, RUN_INIT))// open file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	bmpRunner.pWndProgress = &wndProgress;

	bmpRunner.RunModel(RUN_INIT);

	if (!bmpData.CloseOutputFiles())	// close file for time series
	{
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.CloseOutputFiles())	// close file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	if (bmpRunner.pWndProgress->Cancelled())
	{
		bmpRunner.pWndProgress->DestroyWindow();
		AfxMessageBox("BMP simulation is cancelled");
		goto L001;
	}

	if (!bmpData.OpenOutputFiles("PreDev"))	// open file for time series
	{
		AfxMessageBox(bmpData.strError);
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.OpenOutputFiles("PreDev", bmpData.nRunOption, RUN_PREDEV))	// open file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	bmpRunner.RunModel(RUN_PREDEV);

	if (!bmpData.CloseOutputFiles())	// close file for time series
	{
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.CloseOutputFiles())	// close file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	if (bmpRunner.pWndProgress->Cancelled())
	{
		bmpRunner.pWndProgress->DestroyWindow();
		AfxMessageBox("BMP simulation is cancelled");
		goto L001;
	}

	if (!bmpData.OpenOutputFiles("PostDev"))// open file for time series
	{
		AfxMessageBox(bmpData.strError);
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.OpenOutputFiles("PostDev", bmpData.nRunOption, RUN_POSTDEV))// open file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	bmpRunner.RunModel(RUN_POSTDEV);

	if (!bmpData.CloseOutputFiles())	// close file for time series
	{
		nReturn = FALSE;
		goto L001;
	}
	if (!bmpRunner.CloseOutputFiles())	// close file for evaluation factor
	{
		nReturn = FALSE;
		goto L001;
	}

	if (bmpRunner.pWndProgress->Cancelled())
	{
		bmpRunner.pWndProgress->DestroyWindow();
		AfxMessageBox("BMP simulation is cancelled");
		goto L001;
	}

	srand((unsigned)time(NULL));   // initialize seed for random number generator

	if (bmpData.nRunOption != OPTION_NO_OPTIMIZATION)
	{
		if (bmpData.nStrategy == STRATEGY_SCATTER_SEARCH)
		{
			// run option MinimizeCost or MaximizeControl
			CBMPOptimizer bmpOptimizer(&bmpRunner);
			
			if (bmpData.nAdjVariable < 5)
			{
				bmpOptimizer.problem.b1 = 5;
				bmpOptimizer.problem.b2 = 5;
			}
			else
			{
				bmpOptimizer.problem.b1 = bmpData.nAdjVariable;
				bmpOptimizer.problem.b2 = bmpData.nAdjVariable;
			}
			bmpOptimizer.problem.PSize = 3*(bmpOptimizer.problem.b1+bmpOptimizer.problem.b2);
			bmpOptimizer.problem.LS = FALSE;

			if (bmpRunner.lInitRunTime != 0)
				bmpOptimizer.nMaxRun = int((bmpData.lfMaxRunTime*3600000)/bmpRunner.lInitRunTime)+1;
			else
				bmpOptimizer.nMaxRun = 1000;
			if (bmpOptimizer.nMaxRun < 2*bmpOptimizer.problem.PSize)
				bmpOptimizer.nMaxRun = 2*bmpOptimizer.problem.PSize;
			bmpOptimizer.nMaxIter = int(bmpOptimizer.nMaxRun / (2*bmpOptimizer.problem.PSize)) + 1;

			bmpRunner.nMaxRun = bmpOptimizer.nMaxRun;

			// run optimization for Minimum Cost or MaximumControl options
			if (bmpData.nRunOption == OPTION_MIMIMIZE_COST ||
				bmpData.nRunOption == OPTION_MAXIMIZE_CONTROL)
			{
				CString	strOutPath = bmpData.strOutputDir + "\\AllSolutions.out";
				bmpOptimizer.m_pAllSolutions = fopen(LPCSTR(strOutPath), "wt");
				if (bmpOptimizer.m_pAllSolutions == NULL)
				{
					nReturn = FALSE;
					goto L001;
				}

				bmpOptimizer.OutputFileHeader("SS - All solutions", bmpOptimizer.m_pAllSolutions);
				bmpOptimizer.nRunCounter = 0;
				bmpOptimizer.InitProblem(bmpData.nAdjVariable, bmpOptimizer.problem.b1, bmpOptimizer.problem.b2, bmpOptimizer.problem.PSize, bmpOptimizer.problem.LS);
				bmpOptimizer.InitRefSet();
				bmpOptimizer.PerformSearch();
				fclose(bmpOptimizer.m_pAllSolutions);
				bmpOptimizer.OutputBestSolutions();

				if (bmpRunner.pWndProgress->Cancelled())
				{
					bmpRunner.pWndProgress->DestroyWindow();
					AfxMessageBox("BMP simulation is cancelled");
					goto L001;
				}
				else
				{
					bmpRunner.pWndProgress->DestroyWindow();
					AfxMessageBox("BMP simulation completed successfully", MB_ICONINFORMATION);
					goto L001;
				}
			}
			// run optimization for TradeOff Curve options
			else if (bmpData.nRunOption == OPTION_TRADE_OFF_CURVE) 
			{
				CString	strOutPath = bmpData.strOutputDir + "\\AllSolutions.out";
				bmpOptimizer.m_pAllSolutions = fopen(LPCSTR(strOutPath), "wt");
				if (bmpOptimizer.m_pAllSolutions == NULL)
				{
					nReturn = FALSE;
					goto L001;
				}

				bmpOptimizer.OutputFileHeader("SS - All solutions", bmpOptimizer.m_pAllSolutions);

				// prepare output file for TradeOff Curve optimization
				strFilePath = bmpRunner.pBMPData->strOutputDir + "\\CECurve_Solutions.out";
				FILE* fp = fopen(LPCSTR(strFilePath), "wt");
				if(fp == NULL)
				{
					AfxMessageBox("Failed in creating output file "+strFilePath, MB_ICONEXCLAMATION);
					nReturn = FALSE;					
					goto L001;
				}
				bmpOptimizer.OutputFileHeaderForTradeOffCurve(fp);
				
				// initialize problem only once
				bmpOptimizer.InitProblem(bmpData.nAdjVariable, bmpOptimizer.problem.b1, bmpOptimizer.problem.b2, bmpOptimizer.problem.PSize, bmpOptimizer.problem.LS);
				
				//update the max. runs
				bmpRunner.nMaxRun *= (bmpData.nTargetBreak + 1);

				// run optimization for the target range
				for (int i=0; i<=bmpData.nTargetBreak; i++)
				{
					POSITION pos, pos1;
					pos = bmpData.routeList.GetHeadPosition();

					while (pos != NULL)
					{
						CBMPSite* pBMPSite = (CBMPSite*) bmpData.routeList.GetNext(pos);
						pos1 = pBMPSite->m_factorList.GetHeadPosition();

						while (pos1 != NULL)
						{
							EVALUATION_FACTOR* pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
							pEF->m_lfTarget = pEF->m_lfUpperTarget - i*(pEF->m_lfUpperTarget-pEF->m_lfLowerTarget)/bmpData.nTargetBreak;
							pEF->m_lfNextTarget = pEF->m_lfTarget - (pEF->m_lfUpperTarget-pEF->m_lfLowerTarget)/bmpData.nTargetBreak;
						}
					}

					bmpOptimizer.nRunCounter = 0;
					if (i == 0)
						bmpOptimizer.InitRefSet();
					else
						bmpOptimizer.ResetRefSet();
					bmpOptimizer.PerformSearch();
					bmpOptimizer.OutputBestSolutionsForTradeOffCurve(i, fp);

					if (bmpRunner.pWndProgress->Cancelled())
					{
						bmpRunner.pWndProgress->DestroyWindow();
						AfxMessageBox("BMP simulation is cancelled");
						fclose(bmpOptimizer.m_pAllSolutions);
						fclose(fp);
						goto L001;
					}
				}

				fclose(bmpOptimizer.m_pAllSolutions);
				fclose(fp);
			}
		}
		else if (bmpData.nStrategy == STRATEGY_GENETIC_ALGORITHM)
		{
			//run the best population scenarios (call from the post-processor spreadsheet)
			if(str2.GetLength() > 0)
			{
				CString strLine, strValue;
				POSITION pos, pos1;
				CBMPSite* pBMPSite;
				EVALUATION_FACTOR* pEF;
				ADJUSTABLE_PARAM* pAP;

				CStringToken strToken(str2);
				int nBestPops = 0;
				while (strToken.HasMoreTokens())
				{
					str1 = strToken.NextToken();
					nBestPops++;
				}

				bmpData.nSolution = nBestPops;

				// prepare output file for TradeOff Curve optimization
				strFilePath = bmpData.strOutputDir + "\\CECurve_Solutions.out";
				FILE* fp = fopen(LPCSTR(strFilePath), "wt");
				if(fp == NULL)
				{
					AfxMessageBox("Failed in creating output file "+strFilePath, MB_ICONEXCLAMATION);
					nReturn = FALSE;					
					goto L001;
				}
					
				bmpData.OutputFileHeaderForTradeOffCurve(fp);
				CStringToken strToken1(str2);
						
				for (int i=0; i<nBestPops; i++)
				{
					//read best population file
					int nBestPopId = atoi((LPCSTR)strToken1.NextToken());
					if (!bmpData.ReadBestPopFile(nBestPopId))
					{
						AfxMessageBox(bmpData.strError);
						nReturn = FALSE;
						goto L001;
					}

					//output time series for the best pop solution
					strValue.Format("BestPop%d", nBestPopId);
					if (!bmpData.OpenOutputFiles(strValue))
					{
						AfxMessageBox(bmpData.strError);
						nReturn = FALSE;
						goto L001;
					}

					//run model
					bmpRunner.RunModel(RUN_OUTPUT);

					//close time series for the best pop solution
					if (!bmpData.CloseOutputFiles())
					{
						AfxMessageBox(bmpData.strError);
						nReturn = FALSE;
						goto L001;
					}

					//output CECurve_Solutions
					double totalCost = 0.0;
					double output1 = 0.0;
					pos = bmpData.routeList.GetHeadPosition();
					while (pos != NULL)
					{
						pBMPSite = (CBMPSite*) bmpData.routeList.GetNext(pos);
						totalCost += pBMPSite->m_lfCost;
					}

					strLine.Format("%d\t%d", i+1, 1);
					strValue.Format("\t%lf", totalCost);
					strLine += strValue;

					//evaluation factors
					pos = bmpData.routeList.GetHeadPosition();
					while (pos != NULL)
					{
						pBMPSite = (CBMPSite*) bmpData.routeList.GetNext(pos);
						pos1 = pBMPSite->m_factorList.GetHeadPosition();
						while (pos1 != NULL)
						{
							pEF = (EVALUATION_FACTOR*) pBMPSite->m_factorList.GetNext(pos1);
							if (pEF->m_nCalcMode == CALC_PERCENT) // if the calculation mode is percentage
							{
								if (pEF->m_lfPostDev == 0.0)
									output1 = pEF->m_lfCurrent;
								else
									output1 = pEF->m_lfCurrent/pEF->m_lfPostDev*100;
							}
							else if (pEF->m_nCalcMode == CALC_VALUE) // if the calculation mode is value
							{
								output1 = pEF->m_lfCurrent;
							}
							else // if the calculation mode is scale
							{
								if (pEF->m_lfPostDev - pEF->m_lfPreDev > 0)
									output1 = (pEF->m_lfCurrent - pEF->m_lfPreDev) / (pEF->m_lfPostDev-pEF->m_lfPreDev);
								else
									output1 = pEF->m_lfCurrent;
							}
							strValue.Format("\t%lf", output1);
							strLine += strValue;
						}
					}

					//adjust parameters
					pos = bmpData.routeList.GetHeadPosition();
					while (pos != NULL)
					{
						pBMPSite = (CBMPSite*) bmpData.routeList.GetNext(pos);
						pos1 = pBMPSite->m_adjustList.GetHeadPosition();
						while (pos1 != NULL)
						{
							pAP = (ADJUSTABLE_PARAM*) pBMPSite->m_adjustList.GetNext(pos1);
							double* pVariable = pBMPSite->GetVariablePointer(pAP->m_strVariable);
							strValue.Format("\t%lf", *pVariable);
							strLine += strValue;
						}
					}

					strLine += "\n";
					fputs(strLine, fp);
					
					if (bmpRunner.pWndProgress->Cancelled())
					{
						bmpRunner.pWndProgress->DestroyWindow();
						AfxMessageBox("BMP simulation is cancelled");
						fclose(fp);
						nReturn = FALSE;
						goto L001;
					}
				}

				bmpRunner.pWndProgress->DestroyWindow();
				AfxMessageBox("BMP simulation is completed", MB_ICONINFORMATION);
				fclose(fp);
				nReturn = TRUE;
				goto L001;
			}		
			else
			{
				CBMPOptimizerGA bmpOptimizerGA(&bmpRunner);

				if (!bmpOptimizerGA.OpenOutputFiles())
				{
					AfxMessageBox(bmpData.strError);
					nReturn = FALSE;
					goto L001;
				}

				if (!bmpOptimizerGA.LoadData())
				{
					nReturn = FALSE;
					goto L001;
				}

				if (!bmpOptimizerGA.ValidateParams())
				{
					AfxMessageBox(bmpData.strError);
					nReturn = FALSE;
					goto L001;
				}

				bmpOptimizerGA.nMaxRun = bmpOptimizerGA.problem.popsize * bmpOptimizerGA.problem.ngen;
				bmpRunner.nMaxRun = bmpOptimizerGA.nMaxRun;

				if (!bmpOptimizerGA.InitProblem())
				{
					nReturn = FALSE;
					goto L001;
				}

				bmpOptimizerGA.PerformSearch();
				bmpOptimizerGA.OutputBestPopulation();
				bmpOptimizerGA.CloseOutputFiles();

				if (bmpRunner.pWndProgress->Cancelled())
				{
					bmpRunner.pWndProgress->DestroyWindow();
					AfxMessageBox("BMP simulation is cancelled");
					nReturn = FALSE;
					goto L001;
				}
			}
		}
	}

	bmpRunner.pWndProgress->DestroyWindow();
	AfxMessageBox("BMP simulation is completed", MB_ICONINFORMATION);

L001:
	//SWMM5 release memory
    deleteHashTables();
	transect_delete();
	for (int j=0; j<bmpData.nBMPC; j++)
	{
		FREE(Link[j].oldQual);
		FREE(Link[j].newQual);
	}
	FREE(Link);
    FREE(Conduit);
	FREE(GAInfil);

	return nReturn;
}

int LandSimulation(int landfg,char* strInputFilePath,CProgressWnd* pwndProgress)
{
	CString ReportFilePath(strInputFilePath);
	ReportFilePath.TrimRight("inp");
	CString OutputFilePath = ReportFilePath;
	ReportFilePath += "rpt";
	OutputFilePath += "out";
	char* strReportFilePath = ReportFilePath.GetBuffer(ReportFilePath.GetLength());
	char* strOutputFilePath = OutputFilePath.GetBuffer(OutputFilePath.GetLength());
	
	// initialize progress bar
	pwndProgress->SetRange(0, 100);			 
	pwndProgress->SetText("");
	CString strMsg, strForDdg, strE;

	COleDateTime time_i;		// time at the beginning of the simulation
	COleDateTime time_f;		// time at the end of the simulation
	COleDateTimeSpan time_dif;	// simulation run time
	SYSTEMTIME tm;				// system time
	GetLocalTime(&tm);
	time_i = COleDateTime(tm);

	long newHour, oldHour = 0;
	DateTime elapsedTime = 0.0;

	// --- open the files & read input data
	ErrorCode = 0;
	swmm_open(strInputFilePath,strReportFilePath,strOutputFilePath);

	// --- run the simulation if input data OK
	if ( !ErrorCode )
	{
		// --- initialize values
		swmm_start(TRUE);

		// --- execute each time step until elapsed time is re-set to 0
		if ( !ErrorCode )
		{
			int y, m, d;
			datetime_decodeDate(StartDateTime, &y, &m, &d);
			COleDateTime tStart(y,m,d,0,0,0);
			datetime_decodeDate(EndDateTime, &y, &m, &d);
			COleDateTime tEnd(y,m,d,0,0,0);
			COleDateTimeSpan span0 = tEnd - tStart;

			do
			{
				swmm_step(&elapsedTime);
				newHour = elapsedTime * 24.0;

				COleDateTimeSpan span = COleDateTimeSpan(0,newHour,0,0);
				COleDateTime tCurrent = tStart + span;

				int nSYear = tCurrent.GetYear();
				int nSMonth = tCurrent.GetMonth();
				int nSDay = tCurrent.GetDay();
				int nSHour = tCurrent.GetHour();

				GetLocalTime(&tm);
				time_f = COleDateTime(tm);
				time_dif = time_f - time_i;

				int dd_elap = int(time_dif.GetDays());
				int hh_elap = int(time_dif.GetHours());
				int mm_elap = int(time_dif.GetMinutes());
				int ss_elap = int(time_dif.GetSeconds());

				if ( newHour > oldHour )
				{
					oldHour = newHour;

					if (landfg == 0)
						strMsg.Format("Land Simulation:\t Pre-Development Scenario\n");
					else
						strMsg.Format("Land Simulation:\t Post-Development Scenario\n");
					strForDdg = strMsg;

					strMsg.Format("Calculating:\t %02d-%02d-%04d\n", nSMonth, nSDay, nSYear);
					strForDdg += strMsg;
					
					strE.Format("\nTime Elapsed:\t %02d:%02d:%02d:%02d\n", dd_elap, hh_elap, mm_elap, ss_elap);
					strForDdg += strE;

					double lfPart = span.GetTotalSeconds();
					double lfAll  = span0.GetTotalSeconds();
					double lfPerc = lfPart/lfAll;

					if(pwndProgress->GetSafeHwnd() != NULL && nSHour == 0)
					{
						pwndProgress->SetText(strForDdg);
						pwndProgress->SetPos((int)(lfPerc*100));
						pwndProgress->PeekAndPump();
					}

					if (pwndProgress->Cancelled())
					{
						pwndProgress->DestroyWindow();
						AfxMessageBox("BMP simulation is cancelled");
						break;
					}
				}
			} while ( elapsedTime > 0.0 && !ErrorCode );
		}

		// --- clean up
		swmm_end();
	}

	// --- report results
	swmm_report();

	// --- close the system
	swmm_close();

	//return ErrorCode;
	return ErrorCode;
}

double pet_Hamon(double lat, double cts, double tavc, double day)
{
	//calculate daily PET value based on Hamon method 1961 (in/day)

	//  PET  = cts * DYL^2 * VDSAT 
	//  DYL  = 7.63942 * [atan{-X/sqrt(-X^2+1)}+2*atan(1)]
	//    X  = 0.43481 * tan(lat * 0.017453) * cos[0.0172 * (day + 9)]
	// VDSAT = [(216.7 * VPSAT) / (tavc + 273.3)]
	// VPSAT = 6.108 * exp[(17.26939 * tavc) / (tavc + 273.3)]

	double lfPET = 0.0;
	double lfX = 0.43481*tan(lat*0.017453)*cos(0.0172*(day+9));
	double lfDYL = 7.63942*(atan(-lfX/sqrt(1-lfX*lfX))+2*atan(1));
	if (lfDYL > 24.0) lfDYL = 24.0;
	if (lfDYL < 0.0) lfDYL = 0.0;
	double lfVPSAT = 6.108 * exp((17.26939 * tavc) / (tavc + 273.3));
	double lfVDSAT = ((216.7 * lfVPSAT) / (tavc + 273.3));
	
	lfPET  = cts * lfDYL / 12.0 * lfDYL / 12.0 * lfVDSAT; 

	return lfPET;
}

int CallVFSMOD(LPCSTR strVfsProjFile)
{
	int nReturnVal = 0;

	HINSTANCE m_hViewDll;
//	m_hViewDll = AfxLoadLibrary("E:\\CVS\\VFSMOD_DLL\\Debug\\VFSDLL.dll");
	m_hViewDll = AfxLoadLibrary("VFSDLL.dll");	//DLL in sys32 folder
	
	if (!m_hViewDll)
	{
	  AfxMessageBox("Error: Cannot find component \"VFSDLL.dll\"");
	  return nReturnVal;
	}

	GETDLLVIEW GetView = (GETDLLVIEW) GetProcAddress(m_hViewDll, "StartBufferSimulation");
//	ASSERT (GetView != NULL);

   if (!GetView)
   {
      // handle the error
      FreeLibrary(m_hViewDll);
	  AfxMessageBox("Error: Cannot find DLL entry point \"StartBufferSimulation\"");
	  return nReturnVal;
   }
   else
   {
      // call the function
      nReturnVal = GetView(strVfsProjFile, strlen(strVfsProjFile));
      FreeLibrary(m_hViewDll);
   }

	return nReturnVal;
}