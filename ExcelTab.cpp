// ExcelTab.cpp : implementation file
//

#include "stdafx.h"
#include "excel2.h"
#include "ExcelTab.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// ExcelTab dialog


ExcelTab::ExcelTab(CWnd* pParent /*=NULL*/)
	: CDialog(ExcelTab::IDD, pParent)
{
	//{{AFX_DATA_INIT(ExcelTab)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void ExcelTab::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(ExcelTab)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(ExcelTab, CDialog)
	//{{AFX_MSG_MAP(ExcelTab)
		// NOTE: the ClassWizard will add message map macros here
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// ExcelTab message handlers
