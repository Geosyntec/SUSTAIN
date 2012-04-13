// SiteLandUse.h: interface for the CSiteLandUse class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SITELANDUSE_H__D58C417C_E57D_4149_B031_1DBFDD6DD30F__INCLUDED_)
#define AFX_SITELANDUSE_H__D58C417C_E57D_4149_B031_1DBFDD6DD30F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CLandUse;
class CBMPSite;

class CSiteLandUse : public CObject
{
public:
	CSiteLandUse();
	CSiteLandUse(CLandUse *pLU, CBMPSite *pBMPSite, double lfArea);
	virtual ~CSiteLandUse();

public:
	double m_lfArea; // area size
	CLandUse* m_pLU;
	CBMPSite* m_pBMPSite;
};

#endif // !defined(AFX_SITELANDUSE_H__D58C417C_E57D_4149_B031_1DBFDD6DD30F__INCLUDED_)
