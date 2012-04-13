// SitePointSource.h: interface for the CSitePointSource class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SITEPOINTSOURCE_H__E65B7665_7D0A_4BF4_AC9D_98A4362E03B2__INCLUDED_)
#define AFX_SITEPOINTSOURCE_H__E65B7665_7D0A_4BF4_AC9D_98A4362E03B2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBMPSite;

class CSitePointSource : public CObject  
{
public:
	CSitePointSource();
	virtual ~CSitePointSource();

public:
	int m_nID;				// Point Source ID
	double m_lfMult;		// The multiplier to the timeseries file
	double m_lfSand;		// The fraction of total sediemnt which is sand
	double m_lfSilt;		// The fraction of total sediemnt which is silt
	double m_lfClay;		// The fraction of total sediemnt which is clay
	CString m_strPSFile;	// The timeseries file name
	CBMPSite* m_pBMPSite;	// pointer to the associated BMPSite

	int m_nQualNum;
	long m_nTSNum;
	COleDateTime m_tmStart;
	double* m_pDataPS;		// pointer to the point source timeseries data 

public:
	bool LoadPointsourceTSData(COleDateTime startDate,COleDateTime endDate,double* multiplier);
};

#endif // !defined(AFX_SITEPOINTSOURCE_H__E65B7665_7D0A_4BF4_AC9D_98A4362E03B2__INCLUDED_)

