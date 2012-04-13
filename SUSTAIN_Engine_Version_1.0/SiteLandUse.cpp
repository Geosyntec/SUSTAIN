// SiteLandUse.cpp: implementation of the CSiteLandUse class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "LandUse.h"
#include "BMPSite.h"
#include "SiteLandUse.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CSiteLandUse::CSiteLandUse()
{
	m_lfArea = 0.0;
	m_pLU = NULL;
	m_pBMPSite = NULL;
}

CSiteLandUse::CSiteLandUse(CLandUse *pLU, CBMPSite *pBMPSite, double lfArea)
{
	m_lfArea = lfArea;
	m_pLU = pLU;
	m_pBMPSite = pBMPSite;
}

CSiteLandUse::~CSiteLandUse()
{
}
