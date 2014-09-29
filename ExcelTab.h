#if !defined(AFX_EXCELTAB_H__B038474C_4E40_40E4_BBD2_2825B05A7D46__INCLUDED_)
#define AFX_EXCELTAB_H__B038474C_4E40_40E4_BBD2_2825B05A7D46__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ExcelTab.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// ExcelTab dialog

class ExcelTab : public CDialog
{
// Construction
public:
	ExcelTab(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(ExcelTab)
	enum { IDD = _UNKNOWN_RESOURCE_ID_ };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(ExcelTab)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(ExcelTab)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EXCELTAB_H__B038474C_4E40_40E4_BBD2_2825B05A7D46__INCLUDED_)
