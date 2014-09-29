// Para1.cpp : implementation file
//

#include "stdafx.h"
#include "excel2.h"
#include "Para1.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPara1 dialog


CPara1::CPara1(CWnd* pParent /*=NULL*/)
	: CDialog(CPara1::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPara1)
	m_filename = _T("");
	m_sheet = 1;
	m_startrow = 2;
	m_col = 1;
	m_savefile = _T("result.txt");
	m_out_cols = _T("1,2");
	//}}AFX_DATA_INIT
}


void CPara1::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPara1)
	DDX_Text(pDX, IDC_EDIT1, m_filename);
	DDX_Text(pDX, IDC_EDIT2, m_sheet);
	DDX_Text(pDX, IDC_EDIT3, m_startrow);
	DDX_Text(pDX, IDC_EDIT4, m_col);
	DDX_Text(pDX, IDC_EDIT5, m_savefile);
	DDX_Text(pDX, IDC_EDIT6, m_out_cols);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPara1, CDialog)
	//{{AFX_MSG_MAP(CPara1)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPara1 message handlers

void CPara1::OnOK() 
{
	// TODO: Add extra validation here
    UpdateData(TRUE);
    
    if (0 == m_filename.GetLength()
        || 0 == m_sheet
        || 0 == m_startrow
        || 0 == m_col
        || 0 == m_savefile.GetLength()
        || 0 == m_out_cols.GetLength())
    {
		AfxMessageBox(_T("无效输入!"));
    }
    else
    {
        const char * out_cols = (LPCSTR)m_out_cols;

        if (!set_out_txt_file1_cols(out_cols))
        {
            
            AfxMessageBox(_T("输出列内容输入不正确!"));
            return;
        }
        
        find_rows_no_translate((LPCSTR)m_filename,
            (LPCSTR)m_savefile,
            m_sheet,
            m_startrow,
            m_col,
            ">"
            );
        
		AfxMessageBox(_T("已完成!"));
    }
    
	//CDialog::OnOK();
}
