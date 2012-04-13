// LandUse.h: interface for the CLandUse class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_LANDUSE_H__0251A527_E2D4_4206_8C22_E9481E5AB851__INCLUDED_)
#define AFX_LANDUSE_H__0251A527_E2D4_4206_8C22_E9481E5AB851__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CLandUse : public CObject  
{
public:
	CLandUse();
	virtual ~CLandUse();

public:
	CString m_strLanduse;	// landuse name
	CString m_strFileName;	// time series file path
	int	m_nID;		//serial number
	int	m_nType;	//pevious or impevious
	int m_nQualNum;
	long m_nTSNum;
	double* m_pData;
	COleDateTime m_tmStart;
	double m_lfsand_fr;
	double m_lfsilt_fr;
	double m_lfclay_fr;

public:
	CString* GetPollutantName(int& nNum);
	bool LoadLanduseTSData(COleDateTime startDate,COleDateTime endDate,double* multiplier);
};

#endif // !defined(AFX_LANDUSE_H__0251A527_E2D4_4206_8C22_E9481E5AB851__INCLUDED_)
