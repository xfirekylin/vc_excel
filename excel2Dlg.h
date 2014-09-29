// excel2Dlg.h : header file
//

#if !defined(AFX_EXCEL2DLG_H__8137C962_9C05_4CE2_9131_DA2D5BE007AD__INCLUDED_)
#define AFX_EXCEL2DLG_H__8137C962_9C05_4CE2_9131_DA2D5BE007AD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Para1.h"
#include "Para2.h"
/////////////////////////////////////////////////////////////////////////////
// CExcel2Dlg dialog

class CExcel2Dlg : public CDialog
{
// Construction
public:
	CExcel2Dlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CExcel2Dlg)
	enum { IDD = IDD_EXCEL2_DIALOG };
	CTabCtrl	m_tab;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CExcel2Dlg)
	//}}AFX_VIRTUAL
    bool OpenExcelBook(CString filename);
    void NewExcelBook();
    void OpenExcelApp(void);
    void SaveExcel(void);
    void SaveAsExcel(CString filename);
    void SetCellValue(int row, int col,int Align,CString value);
    CString GetCellValue(int row, int col);;
    void SetRowHeight(int row, CString height);
        void SetColumnWidth(int col,CString width);
        CString GetColumnWidth(int col);
        CString GetRowHeight(int row);
        CString IndexToString( int row, int col ); 
        int LastLineIndex(); 
        
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;
	CPara2 m_para2;
	
	CPara1 m_para1;
	

	// Generated message map functions
	//{{AFX_MSG(CExcel2Dlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_EXCEL2DLG_H__8137C962_9C05_4CE2_9131_DA2D5BE007AD__INCLUDED_)
