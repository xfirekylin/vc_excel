#if !defined(AFX_PARA2_H__1026C81A_8015_45CA_AC6A_9FD2E742634A__INCLUDED_)
#define AFX_PARA2_H__1026C81A_8015_45CA_AC6A_9FD2E742634A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Para2.h : header file
//

void find_rows_in_file2_same_with_file1( const char *filename1, 
												const char *filename2, 
												const char *filename3,
												const char *save_file,
												long file1_sheet_index,
												long file2_sheet_index,
												long file1_start_row, 
												long file2_start_row, 
												BOOLEAN is_only_mark,
												BOOLEAN is_write_file3,
												char *col_flag);
int set_out_txt_file1_cols(const char *buf);
int set_out_txt_file2_cols(const char *buf);
int set_out_excel_file1_cols(const char *buf);
int set_out_excel_file2_cols(const char *buf);
int set_compare_cols(long file1_col, long file2_col);

/////////////////////////////////////////////////////////////////////////////
// CPara2 dialog

class CPara2 : public CDialog
{
// Construction
public:
	CPara2(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPara2)
	enum { IDD = IDD_PARA2 };
	CString	m_file1;
	CString	m_file2;
	CString	m_file3;
	CString	m_savefile;
	int		m_file1_sheet;
	int		m_file2_sheet;
	int		m_file_startrow;
	int		m_file2_startrow;
	BOOL	m_is_savefile3;
	BOOL	m_is_only_mark;
	CString	m_out_txt_col1;
	CString	m_out_txt_col2;
	CString	m_out_excel_col1;
	CString	m_out_excel_col2;
	int		m_file1_col;
	int		m_file2_col;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPara2)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CPara2)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PARA2_H__1026C81A_8015_45CA_AC6A_9FD2E742634A__INCLUDED_)
