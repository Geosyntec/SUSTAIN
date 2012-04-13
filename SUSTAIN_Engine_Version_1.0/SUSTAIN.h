// SUSTAIN.h : main header file for the SUSTAIN DLL
//

#if !defined(AFX_SUSTAIN_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_)
#define AFX_SUSTAIN_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CSUSTAINApp
// See SUSTAIN.cpp for the implementation of this class
//

class CSUSTAINApp : public CWinApp
{
public:
	CSUSTAINApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSUSTAINApp)
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CSUSTAINApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SUSTAIN_H__E8EDFB62_5A44_40C2_9B1A_CD0EA86DEC14__INCLUDED_)
