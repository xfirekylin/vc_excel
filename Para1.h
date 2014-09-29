#if !defined(AFX_PARA1_H__FA33324D_CCD6_4D51_9A7A_6A19E816BECB__INCLUDED_)
#define AFX_PARA1_H__FA33324D_CCD6_4D51_9A7A_6A19E816BECB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Para1.h : header file
//
void find_rows_no_translate(const char *filename1, 
								const char *save_file,
								long file1_sheet_index,
								long file1_start_row, 
								long file1_col,
								char *col_flag);
int set_out_txt_file1_cols(const char *buf);
int set_out_txt_file2_cols(const char *buf);
int set_out_excel_file1_cols(const char *buf);
int set_out_excel_file2_cols(const char *buf);

/////////////////////////////////////////////////////////////////////////////
// CPara1 dialog

class CPara1 : public CDialog
{
// Construction
public:
	CPara1(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPara1)
	enum { IDD = IDD_PARA1 };
	CString	m_filename;
	int		m_sheet;
	int		m_startrow;
	int		m_col;
	CString	m_savefile;
	CString	m_out_cols;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPara1)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CPara1)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PARA1_H__FA33324D_CCD6_4D51_9A7A_6A19E816BECB__INCLUDED_)
