// excel2.h : main header file for the EXCEL2 application
//

#if !defined(AFX_EXCEL2_H__18C6B48B_6A36_45C1_8BC1_E2E23585FE4A__INCLUDED_)
#define AFX_EXCEL2_H__18C6B48B_6A36_45C1_8BC1_E2E23585FE4A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CExcel2App:
// See excel2.cpp for the implementation of this class
//

class CExcel2App : public CWinApp
{
public:
	CExcel2App();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CExcel2App)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CExcel2App)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EXCEL2_H__18C6B48B_6A36_45C1_8BC1_E2E23585FE4A__INCLUDED_)
