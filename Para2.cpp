// Para2.cpp : implementation file
//

#include "stdafx.h"
#include "excel2.h"
#include "Para2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPara2 dialog


CPara2::CPara2(CWnd* pParent /*=NULL*/)
	: CDialog(CPara2::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPara2)
	m_file1 = _T("");
	m_file2 = _T("");
	m_file3 = _T("");
	m_savefile = _T("result.txt");
	m_file1_sheet = 1;
	m_file2_sheet = 1;
	m_file_startrow = 2;
	m_file2_startrow = 1;
	m_is_savefile3 = FALSE;
	m_is_only_mark = FALSE;
	m_out_txt_col1 = _T("");
	m_out_txt_col2 = _T("");
	m_out_excel_col1 = _T("");
	m_out_excel_col2 = _T("");
	m_file1_col = 1;
	m_file2_col = 1;
	//}}AFX_DATA_INIT
}


void CPara2::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPara2)
	DDX_Text(pDX, IDC_EDIT1, m_file1);
	DDX_Text(pDX, IDC_EDIT2, m_file2);
	DDX_Text(pDX, IDC_EDIT3, m_file3);
	DDX_Text(pDX, IDC_EDIT4, m_savefile);
	DDX_Text(pDX, IDC_EDIT5, m_file1_sheet);
	DDX_Text(pDX, IDC_EDIT6, m_file2_sheet);
	DDX_Text(pDX, IDC_EDIT7, m_file_startrow);
	DDX_Text(pDX, IDC_EDIT8, m_file2_startrow);
	DDX_Check(pDX, IDC_CHECK1, m_is_savefile3);
	DDX_Check(pDX, IDC_CHECK2, m_is_only_mark);
	DDX_Text(pDX, IDC_EDIT9, m_out_txt_col1);
	DDX_Text(pDX, IDC_EDIT10, m_out_txt_col2);
	DDX_Text(pDX, IDC_EDIT11, m_out_excel_col1);
	DDX_Text(pDX, IDC_EDIT12, m_out_excel_col2);
	DDX_Text(pDX, IDC_EDIT13, m_file1_col);
	DDX_Text(pDX, IDC_EDIT14, m_file2_col);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPara2, CDialog)
	//{{AFX_MSG_MAP(CPara2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPara2 message handlers

void CPara2::OnOK() 
{
	// TODO: Add extra validation here

    
    UpdateData(TRUE);

    if (0 == m_file1.GetLength()
        || 0 == m_file2.GetLength()
        || (m_is_savefile3 && 0 == m_file3.GetLength())
        || 0 == m_savefile.GetLength()
        || m_file1_sheet <= 0
        || m_file2_sheet <= 0
        || m_file_startrow <= 0
        || m_file2_startrow <= 0
        || m_file1_col <= 0
        || m_file2_col <= 0
        )
    {
		AfxMessageBox(_T("无效输入!"));
    }
    else
    {
        if (!m_is_only_mark)
        {
            
            const char * out_cols = (LPCSTR)m_out_txt_col1;
            
            if (!set_out_txt_file1_cols(out_cols))
            {
                
                AfxMessageBox(_T("输出列内容输入不正确!"));
                return;
            }
            
            out_cols = (LPCSTR)m_out_txt_col2;
            
            if (!set_out_txt_file2_cols(out_cols))
            {
                
                AfxMessageBox(_T("输出列内容输入不正确!"));
                return;
            }

            if (m_is_savefile3)
            {
                
                out_cols = (LPCSTR)m_out_excel_col1;
                
                if (!set_out_excel_file1_cols(out_cols))
                {
                    
                    AfxMessageBox(_T("输出列内容输入不正确!"));
                    return;
                }
                
                out_cols = (LPCSTR)m_out_excel_col2;
                
                if (!set_out_excel_file2_cols(out_cols))
                {
                    AfxMessageBox(_T("输出列内容输入不正确!"));
                    return;
                }
            }
        }

        set_compare_cols(m_file1_col, m_file2_col);
        
        find_rows_in_file2_same_with_file1((LPCSTR)m_file1,
                                            (LPCSTR)m_file2,
                                            (LPCSTR)m_file3, 
                                            (LPCSTR)m_savefile,
                                            m_file1_sheet, 
                                            m_file2_sheet, 
                                            m_file_startrow, 
                                            m_file2_startrow, 
                                            m_is_only_mark,
                                            m_is_savefile3,
                                            ">");
        
		AfxMessageBox(_T("已完成!"));
    }
//	CDialog::OnOK();
}
