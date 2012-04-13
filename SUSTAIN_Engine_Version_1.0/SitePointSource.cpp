// SitePointSource.cpp: implementation of the CSitePointSource class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BMPSite.h"
#include "StringToken.h"
#include "SitePointSource.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CSitePointSource::CSitePointSource()
{
	m_nID		= -1;	// Point Source ID
	m_lfMult	= 1.0;	// The multiplier to the timeseries file
	m_lfSand	= 0.0;	// The fraction of total sediemnt which is sand
	m_lfSilt	= 0.0;	// The fraction of total sediemnt which is silt
	m_lfClay	= 0.0;	// The fraction of total sediemnt which is clay
	m_strPSFile	= "";	// The timeseries file name
	m_pBMPSite	= NULL;	// pointer to the associated BMPSite
	
	m_nQualNum = 0;
	m_nTSNum = 0;
	m_tmStart = COleDateTime(1890,1,1,0,0,0);
	m_pDataPS = NULL;
}

CSitePointSource::~CSitePointSource()
{
	if (m_pDataPS != NULL)
		delete []m_pDataPS;
}

// load the time series data
bool CSitePointSource::LoadPointsourceTSData(COleDateTime startDate, COleDateTime endDate, 
	 double* multiplier)
{
	int i, j;
	char strLine[MAXLINE];
	CString str;

	FILE *fpin = NULL;
	// open the file for reading
	fpin = fopen (m_strPSFile, "rt");
	if(fpin == NULL)
		return false;
/*
	// find the flag "label"
	m_nQualNum = 0;
	CString strFind = "label";
	while (!feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		str = strLine;
		str.MakeLower();

 		if(str.Find("hspf") != -1)
			strFind = "label   ";
		if(str.Find(strFind) != -1)
		{
			while (!feof(fpin))
			{
				fgets(strLine, MAXLINE, fpin);
				str = strLine;
				CStringToken strToken(str);
				strToken.NextToken();
				str = strToken.LeftOut();
				str.TrimLeft();
				str.TrimRight();
				if(str.GetLength() > 2)
					m_nQualNum++;
				else
					break;
			}
			break;
		}
	}
*/
    // find the flag "date/time"
	while (!feof(fpin))
	{
		fgets(strLine, MAXLINE, fpin);
		str = strLine;
		str.MakeLower();

		// skip another line once found "date/time"
		if(str.Find("date/time") != -1)
		{
			i = 1;
			while(i-- > 0 && !feof(fpin))
				fgets(strLine, MAXLINE, fpin);
			break;
		}
	}
	
	// count the time series numbers
	COleDateTimeSpan tsSpan = endDate - startDate;
	m_nTSNum = (long)tsSpan.GetTotalHours() + 24;
	
    // read first data line for starting date of the time series data
	long nStart = ftell (fpin);
	fgets(strLine, MAXLINE, fpin);
    fseek(fpin, nStart, SEEK_SET);

	int year, month, day, hour, min;
	str = strLine;
	CStringToken strToken(str);
	strToken.NextToken(); // skip the dummy number (first column)
	
	str = strToken.NextToken();
	year = atoi((LPCSTR)str);

	str = strToken.NextToken();
	month = atoi((LPCSTR)str);

	str = strToken.NextToken();
	day = atoi((LPCSTR)str);
	
	str = strToken.NextToken();
	hour = atoi((LPCSTR)str);

	str = strToken.NextToken();
	min = atoi((LPCSTR)str);

	if(hour == 24)																	
	{                          
		hour = 0; //it is the beginning of next day
		m_tmStart = COleDateTime(year, month, day, hour, min, 0) + COleDateTimeSpan(1,0,0,0);
	}
	else
	{
		m_tmStart = COleDateTime(year, month, day, hour, min, 0);
	}

	tsSpan = startDate - m_tmStart;

	// calculate how many lines we need to skip
	long nSkipLineNum = (long)tsSpan.GetTotalHours();

    // skip all lines the time stamp is before the specified start date
	while (nSkipLineNum-- > 0 && !feof(fpin))
		fgets(strLine, MAXLINE, fpin);

	if(m_pDataPS != NULL)
		delete []m_pDataPS;
	m_pDataPS = new double[m_nTSNum*m_nQualNum];

    // read the data
	i = 0;
	double *pData = m_pDataPS;

    while (!feof(fpin))
    {
		// read one line
		fgets(strLine, MAXLINE, fpin);

		// get the data
		str = strLine;
		CStringToken strToken1(str);
		strToken1.NextToken(); // serial number
		strToken1.NextToken(); // year
		strToken1.NextToken(); // month
		strToken1.NextToken(); // day
		strToken1.NextToken(); // hour
		strToken1.NextToken(); // minute

		for(j=0; j<m_nQualNum; j++)
		{
			str = strToken1.NextToken();
			if (j > 0 && multiplier != NULL)
				*(pData++) = atof((LPCSTR)str)*multiplier[j-1];	//lb
			else 
				*(pData++) = atof((LPCSTR)str);
		}

		i++;
		if (i == m_nTSNum)
			break;
	}

	fclose(fpin);
	return (i == m_nTSNum);
}
