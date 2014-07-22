// excel2Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "excel2.h"
#include "excel2Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#ifdef ACCESS_EXCEL_SIMPLE
	#include "excel.h"
#else
	#include "illusion_excel_file.h"
#endif

#include "comdef.h"



_Application app;
Workbooks books;
_Workbook book;
Worksheets sheets;
_Worksheet sheet;
Range range;
Range cell;
Font font;

COleVariant covOptional2((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

IllusionExcelFile acess_excel;
IllusionExcelFile acess_excel2;
IllusionExcelFile acess_excel3;
/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CExcel2Dlg dialog

CExcel2Dlg::CExcel2Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CExcel2Dlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CExcel2Dlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcel2Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CExcel2Dlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CExcel2Dlg, CDialog)
	//{{AFX_MSG_MAP(CExcel2Dlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CExcel2Dlg message handlers

BOOL CExcel2Dlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here

	
#ifdef ACCESS_EXCEL_SIMPLE
	if (::CoInitialize( NULL ) == E_INVALIDARG) 
	{ 
		AfxMessageBox(_T("��ʼ��Comʧ��!")); 
		return;
	}
	
	//��֤office�ļ��Ƿ������ȷ����
	
	if( !app.CreateDispatch(_T("Excel.Application")) )
	{
		AfxMessageBox(_T("�޷�����ExcelӦ�ã�"));
		return;
	}
#else
	acess_excel.InitExcel();

#endif
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CExcel2Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CExcel2Dlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CExcel2Dlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void get_string_by_same_value()
{
    acess_excel.OpenExcelFile("D:\\excel2\\str_table.xls");
	acess_excel.LoadSheet(1, TRUE);
	
	acess_excel2.OpenExcelFile("D:\\excel2\\tr.xls");
	acess_excel2.LoadSheet(1, TRUE);
    
	acess_excel3.OpenExcelFile("D:\\excel2\\re.xls");
	acess_excel3.LoadSheet(1, TRUE);


	int file1_rows = acess_excel.GetRowCount();
	int file2_rows = acess_excel2.GetRowCount();
    
	int file1_cur_row = 2;
	int file2_cur_row = 1;
	int file3_cur_row = 1;

    CString value2 ;

    CString value1;

    
    for (;file2_cur_row<=file2_rows; file2_cur_row++)
    {
        value2 = acess_excel2.GetCellString(file2_cur_row, 4);

        file1_cur_row = 2;
    
    	for (;file1_cur_row<=file1_rows;file1_cur_row++)
    	{
    	    value1 = acess_excel.GetCellString(file1_cur_row, 4);
            
            if (0 == value1.Compare((LPCTSTR)value2))
            {
                acess_excel3.SetCellString(file3_cur_row, 1, acess_excel.GetCellString(file1_cur_row, 1));
                acess_excel3.SetCellString(file3_cur_row, 2, acess_excel.GetCellString(file1_cur_row, 2));
                acess_excel3.SetCellString(file3_cur_row, 3, acess_excel.GetCellString(file1_cur_row, 3));
                acess_excel3.SetCellString(file3_cur_row, 4, acess_excel.GetCellString(file1_cur_row, 4));
				file3_cur_row++;
            }
    	}

    }
    
    acess_excel3.SaveasXSLFile(acess_excel3.GetOpenFileName());
}

void save_cell_to_file(void)
{
    CStdioFile myFile;

    CFileException fileException;
	int i = 1;
    if(myFile.Open("n.txt",CFile::typeText|CFile::modeCreate|CFile::modeReadWrite),&fileException)

    {



    }
    
    acess_excel.OpenExcelFile("D:\\excel2\\str_table.xls");
	acess_excel.LoadSheet(1, TRUE);
	int rows = acess_excel.GetRowCount();

    for (;i<rows;i++)
    {
        #if 0
        TRACE("%s",(LPCTSTR) acess_excel.GetCellString(i, 1));
        TRACE("\n");
        #endif
        myFile.WriteString(acess_excel.GetCellString(i, 1));
        myFile.WriteString(",");
        myFile.WriteString(acess_excel.GetCellString(i, 3));
        myFile.WriteString(",");
        myFile.WriteString(acess_excel.GetCellString(i, 4));
        myFile.WriteString(",");
        myFile.WriteString(acess_excel.GetCellString(i, 5));
        myFile.WriteString("\n");
    }
}

void get_string_by_same_value2(void)
{
    acess_excel.OpenExcelFile("D:\\excel2\\tr.xls");
    acess_excel.LoadSheet(1, TRUE);
    
    acess_excel2.OpenExcelFile("D:\\excel2\\re.xls");
    acess_excel2.LoadSheet(1, TRUE);

    int file1_rows = acess_excel.GetRowCount();
    int file2_rows = acess_excel2.GetRowCount();
    
    int file1_cur_row = 1;
    int file2_cur_row = 1;

    CString value2 ;
    CString value1;

    
    for (;file2_cur_row<=file2_rows; file2_cur_row++)
    {
        value2 = acess_excel2.GetCellString(file2_cur_row, 4);

    
        for (;file1_cur_row<=file1_rows;file1_cur_row++)
        {
            value1 = acess_excel.GetCellString(file1_cur_row, 4);
            
            if (0 == value1.Compare((LPCTSTR)value2))
            {
                acess_excel2.SetCellValue(file2_cur_row, 5, acess_excel.GetCellValue(file1_cur_row, 6));
                break;
            }
        }

    }
    
    acess_excel2.SaveasXSLFile(acess_excel2.GetOpenFileName());
}
typedef unsigned char       BOOLEAN;

typedef unsigned char       uint8;
typedef unsigned short      uint16;
typedef unsigned long int   uint32;
typedef unsigned long int   uint64;
typedef unsigned int        uint;

typedef signed char         int8;
typedef signed short        int16;
typedef signed long int     int32;

typedef enum
{
    MMICOM_SEARCH_FIRST_EQUAL,//������ǰ���һ����ͬ���˳�
    MMICOM_SEARCH_LAST_EQUAL,//��������һ����ͬ���˳�
    MMICOM_SEARCH_ANY_EQUAL,//��������һ����ͬ���˳�
}MMI_BIN_SEARCH_TYPE_E;

typedef struct _MMI_BIN_SEARCH_INFO_T
{
    int32  start_pos;
	int32  end_pos;
	void    *compare_base_data;
	MMI_BIN_SEARCH_TYPE_E search_type;
}MMI_BIN_SEARCH_INFO_T;
typedef int (* BIN_COMPARE_FUNC)(uint32 base_index, void *compare_base_data, void *list);

int MMIAPICOM_BinSearch(MMI_BIN_SEARCH_INFO_T *search_info, //[IN]
                           BIN_COMPARE_FUNC func, //[IN]
                           uint32 *pos, //[OUT]
                           void *list//[IN]
                           )
{
	BOOLEAN         is_in_list  = FALSE;    // ��¼�Ƿ��б�������ȫһ��������
	uint32	        ret_pos         = 0;        // ��¼�ҵ��ĺ���λ��
    int32          low_pos	    = search_info->start_pos;        // ���ַ���ǰһ��λ��
    int32          mid_pos	    = 0;        // ���ַ����м�λ��
   	int32          high_pos	= search_info->end_pos;        // ���ַ��ĺ�һ��λ��
	int 	        cmp_result  = 0;  // �����ַ����ıȽϽ��

    if(NULL == search_info || (NULL == func) || (NULL == pos))
    {
        return FALSE;
    } 
	//func = func_t->func;
	// ���ַ�����
	//SCI_TRACE_LOW("MMIPB_BinSearch low_pos %d, high_pos %d",low_pos,high_pos);
	while (low_pos <= high_pos)	// ������ַ�������
	{
	    // ���mid_pos
	    mid_pos = ((low_pos + high_pos) >> 1); 
        // �������ַ������бȽ�//
        cmp_result = (*func)(mid_pos,search_info->compare_base_data, list);
        // ���ݱȽϵĽ���������ҵķ�Χ
        if (cmp_result < 0)
      	{
          	low_pos = (mid_pos + 1);
      	}
      	else if (cmp_result > 0)
      	{
			if(mid_pos == 0)
			{
				break;
 			}
          	high_pos = mid_pos -1;
        }
        else // begin (0 == cmp_result)
        {
            // ��ʾ�б�������ȫһ���ļ�¼��
            // ��¼Ŀǰ��λ�ã�������Ѿ��ҵ������ļ�¼��
        	is_in_list  = TRUE;
        	ret_pos     = mid_pos;
            
            // ���Ҳ���
            // ����ʱ����Ҫ������ǰ�ң����������ң��õ����ʵĲ��ҷ�Χ  
            if (MMICOM_SEARCH_LAST_EQUAL == search_info->search_type)
        	{
                low_pos = (mid_pos + 1);
            }
            else if(MMICOM_SEARCH_FIRST_EQUAL == search_info->search_type)
            {
            	// ���п��ܲ����б���������������ǰ��һ����¼����Ҫ����Ѱ��
                if(mid_pos == 0)
				{
					break;
				}
				high_pos = (mid_pos - 1);	
            }
            else
            {
                //�ҵ�����һ����ȵľ��˳�
                break;
            }

        } // end (0 == cmp_result)
    } // end of while

    if (!is_in_list)
    {
        if(cmp_result < 0)
        {
            ret_pos = low_pos;
        }
        else
        {
            ret_pos = mid_pos;
        }
    }
	else
	{
		cmp_result = 0;
	}
    *pos = ret_pos;
    
    return cmp_result;    
}

int order_table[20000];
int order_cnt = 0;

 int CompareString(uint32 base_index, void *compare_data, void *list)
{

    int cur_item = (int)compare_data;

    int base_item = order_table[base_index];
    CString value2 ;
    CString value1;

    value1 = acess_excel.GetCellString(base_item, 1);

    return value1.Compare((LPCTSTR)acess_excel.GetCellString(cur_item, 1));
}
 
void insert_table(int end_pos,int cur_item)
{
    uint32 pos = 0;
	//uint16 i = 0;
    uint16 need_moved_num = 0;
    MMI_BIN_SEARCH_INFO_T search_info = {0};
    
    search_info.search_type = MMICOM_SEARCH_ANY_EQUAL;
    
    search_info.end_pos = end_pos -1;
    
    search_info.compare_base_data = (void *)cur_item;
    
    if (0 == MMIAPICOM_BinSearch(&search_info, (BIN_COMPARE_FUNC)CompareString, &pos, (void *)1))
    {
        int ijjj=0;
        ijjj++;
        TRACE("LINE=%d\n",cur_item);
    }
    
    if(pos != order_cnt)
    {
        need_moved_num = order_cnt - pos;
        memmove(&order_table[pos+1], 
                &order_table[pos], 
                need_moved_num * sizeof(order_table[0]));
    }
    
    order_table[pos] = cur_item;
    order_cnt++;
}

int CompareString2(uint32 base_index, void *compare_data, void *list)
{

    int cur_item = (int)compare_data;

    int base_item = order_table[base_index];
    CString value2 ;
    CString value1;

    value1 = acess_excel.GetCellString(base_item, 1);
    value2 = acess_excel2.GetCellString(cur_item, 1);

    return value1.Compare((LPCTSTR)value2);
}

int find_pos(uint32 *pos,uint32 *pos2,int cur_item)
{
    int result = 1;
    MMI_BIN_SEARCH_INFO_T search_info = {0};
    
    search_info.search_type = MMICOM_SEARCH_LAST_EQUAL;
    
    search_info.end_pos = order_cnt -1;
    
    search_info.compare_base_data = (void *)cur_item;
    
    if (0==MMIAPICOM_BinSearch(&search_info, (BIN_COMPARE_FUNC)CompareString2, pos, (void *)1))
    {
        
        search_info.search_type = MMICOM_SEARCH_FIRST_EQUAL;
        
        search_info.end_pos = *pos;
        
        search_info.compare_base_data = (void *)cur_item;
        
        MMIAPICOM_BinSearch(&search_info, (BIN_COMPARE_FUNC)CompareString2, pos2, (void *)1);

        result = 0;
    }

    return result;
}

int find_pos2(uint32 *pos,int cur_item)
{
    MMI_BIN_SEARCH_INFO_T search_info = {0};
    
    search_info.search_type = MMICOM_SEARCH_ANY_EQUAL;
    
    search_info.end_pos = order_cnt -1;
    
    search_info.compare_base_data = (void *)cur_item;
    
    return MMIAPICOM_BinSearch(&search_info, (BIN_COMPARE_FUNC)CompareString2, pos, (void *)1);
}


void get_string_by_id(void)
{
    acess_excel.OpenExcelFile("D:\\excel24\\m0.xls");
    acess_excel.LoadSheet(1, TRUE);
    
    acess_excel2.OpenExcelFile("D:\\excel24\\m1.xls");
    acess_excel2.LoadSheet(1, TRUE);

    acess_excel3.OpenExcelFile("D:\\excel24\\mn.xls");
    acess_excel3.LoadSheet(1, TRUE);

    int file1_rows = acess_excel.GetRowCount();
    int file2_rows = acess_excel2.GetRowCount();
    
    int file1_cur_row = 2;
    int file2_cur_row = 2;

    int file3_cur_row = 1;

    CString value2 ;
    CString value1;

    
    value2 = acess_excel2.GetCellString(file2_cur_row, 1);

    
    CStdioFile myFile;

    CFileException fileException;
	uint32 i = 1;
    if(myFile.Open("n.txt",CFile::typeText|CFile::modeCreate|CFile::modeReadWrite),&fileException)

    {



    }

    memset(order_table, 0 ,sizeof(order_table));

    order_table[0] = 2;
    order_cnt = 1;
    
    for(i=3;i<=file1_rows;i++)
    {
        insert_table(order_cnt, i);
    }

    return;
    
    for (;file2_cur_row<=file2_rows;file2_cur_row++)
    {
        uint32 pos=0;
        uint32 pos2 = 0;
        
        if (0 == find_pos2(&pos, file2_cur_row))
        {
        #if 1
            acess_excel2.SetCellString(file2_cur_row, 1, "xlh__xlh");
        #endif
            
        #if 0
            myFile.WriteString(acess_excel.GetCellString(order_table[pos], 1));
            
            myFile.WriteString("\n");
        #endif
        }
    }

    acess_excel2.SaveasXSLFile(acess_excel2.GetOpenFileName());
}

void get_string_by_id2(void)
{
    acess_excel.OpenExcelFile("D:\\excel24\\m3.xls");
    acess_excel.LoadSheet(1, TRUE);
    
    acess_excel2.OpenExcelFile("D:\\excel24\\m2.xls");
    acess_excel2.LoadSheet(1, TRUE);

    acess_excel3.OpenExcelFile("D:\\excel24\\mn.xls");
    acess_excel3.LoadSheet(1, TRUE);

    int file1_rows = acess_excel.GetRowCount();
    int file2_rows = acess_excel2.GetRowCount();
    
    int file1_cur_row = 2;
    int file2_cur_row = 2;

    int file3_cur_row = 1;

    CString value2 ;
    CString value1;

    
    value2 = acess_excel2.GetCellString(file2_cur_row, 1);

    
    CStdioFile myFile;

    CFileException fileException;
	uint32 i = 1;
    if(myFile.Open("n.txt",CFile::typeText|CFile::modeCreate|CFile::modeReadWrite),&fileException)

    {



    }

    memset(order_table, 0 ,sizeof(order_table));

    order_table[0] = 2;
    order_cnt = 1;
    
    for(i=3;i<=file1_rows;i++)
    {
        insert_table(order_cnt, i);
    }
    
    for (;file2_cur_row<=file2_rows;file2_cur_row++)
    {
        uint32 pos=0;
        uint32 pos2 = 0;
        
        if (0 == find_pos(&pos, &pos2,file2_cur_row))
        {
            for (i=pos2;i<=pos;i++)
            {
                //acess_excel3.SetCellValue(file3_cur_row, 1, acess_excel2.GetCellValue(file2_cur_row, 2));
                file3_cur_row++;
                
                myFile.WriteString(acess_excel.GetCellString(order_table[i], 1));
                myFile.WriteString(">");
                myFile.WriteString(acess_excel.GetCellString(order_table[i], 2));  
                myFile.WriteString(">");
                myFile.WriteString(acess_excel.GetCellString(order_table[i], 3));
                myFile.WriteString(">");
                myFile.WriteString(acess_excel.GetCellString(order_table[i], 4));
            
                myFile.WriteString("\n");
            }
        }
        else 
        {
            myFile.WriteString("\n");
        }
    }

    acess_excel3.SaveasXSLFile(acess_excel3.GetOpenFileName());
}

UINT ThreadFun(LPVOID pParam)
{  //�߳�Ҫ���õĺ���
    
    return 0;
}

void CExcel2Dlg::OnOK() 
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();

#ifdef ACCESS_EXCEL_SIMPLE
	if (!OpenExcelBook("D:\\excel2\\str_table.xls"))
	{
		TRACE("Open xls file fail");
		return;
	}

	int rows = LastLineIndex();
	int i = 0;
	
	TRACE("rows=%d",rows);

	for (i=1;i<rows;i++)
	{
		TRACE("%s",(LPCTSTR) GetCellValue(i, 1));

	}
#else
    //::AfxBeginThread(ThreadFun, NULL); 
    get_string_by_id();

#endif


}

void CExcel2Dlg::OnCancel() 
{
	// TODO: Add extra cleanup here

#ifdef ACCESS_EXCEL_SIMPLE
	app.Quit();  
#endif
	
	CDialog::OnCancel();
}


bool CExcel2Dlg::OpenExcelBook(CString filename)
{
	CFileFind filefind; 
	if( !filefind.FindFile(filename) ) 
	{ 
		AfxMessageBox(_T("�ļ�������"));
		return false;
	}
	LPDISPATCH lpDisp; //�ӿ�ָ��
	books=app.GetWorkbooks();
	lpDisp = books.Open(filename,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2
		);										//��office 2000�Ĳ�ͬ���Ǹ������ģ�ֱ���ں����������covOptional2�ɹ���
	book.AttachDispatch(lpDisp);
	sheets=book.GetSheets();
	sheet=sheets.GetItem(COleVariant((short)1));		//��Ĳ�ͬ���Ǹ������ģ�ֱ���ں����������covOptional2�ɹ���
	return true;
}
void CExcel2Dlg::NewExcelBook()
{
	books=app.GetWorkbooks();
	book=books.Add(covOptional2);
	sheets=book.GetSheets();
	sheet=sheets.GetItem(COleVariant((short)1));		//��Ĳ�ͬ���Ǹ������ģ�ֱ���ں����������covOptional2�ɹ���
}

////////////////////////////////////////////////////////////////////////
///Function:	OpenExcelApp
///Description:	��Ӧ�ó���Ҫע���Ժ����ʶ���û�Ҫ�򿪵����ĸ��ļ���
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::OpenExcelApp(void)
{
	app.SetVisible(TRUE);
	app.SetUserControl(TRUE);
}

////////////////////////////////////////////////////////////////////////
///Function:	SaveExcel
///Description:	���ڴ������ļ����������ݺ�ֱ�ӱ���
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SaveExcel(void)
{
	book.SetSaved(TRUE);
}

////////////////////////////////////////////////////////////////////////
///Function:	SaveAsExcel
///Description:	����excel�ļ�
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SaveAsExcel(CString filename)
{
	book.SaveAs(COleVariant(filename),covOptional2,
	covOptional2,covOptional2,
	covOptional2,covOptional2,(long)0,covOptional2,covOptional2,covOptional2,
	covOptional2,covOptional2);					
}


////////////////////////////////////////////////////////////////////////
///Function:	SetCellValue
///Description:	�޸ĵ�Ԫ���ڵ�ֵ
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int row ��Ԫ��������
///				int col ��Ԫ��������
///				int Align		���뷽ʽĬ��Ϊ����
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SetCellValue(int row, int col,int Align,CString value)
{
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	range.SetValue2(COleVariant(value));
	cell.AttachDispatch((range.GetItem (COleVariant(long(1)), COleVariant(long(1)))).pdispVal);
	cell.SetHorizontalAlignment(COleVariant((short)Align));
}

////////////////////////////////////////////////////////////////////////
///Function:	GetCellValue
///Description:	�õ��ĵ�Ԫ���е�ֵ
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int row ��Ԫ��������
///				int col ��Ԫ��������
///Return:		CString ��Ԫ���е�ֵ
////////////////////////////////////////////////////////////////////////
CString CExcel2Dlg::GetCellValue(int row, int col)
{
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	COleVariant rValue;
	rValue=COleVariant(range.GetValue2());
	rValue.ChangeType(VT_BSTR);
	return CString(rValue.bstrVal);
}
////////////////////////////////////////////////////////////////////////
///Function:	SetRowHeight
///Description:	�����и�
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int row ��Ԫ��������
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SetRowHeight(int row, CString height)
{
	int col = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	range.SetRowHeight(COleVariant(height));
}
////////////////////////////////////////////////////////////////////////
///Function:	SetColumnWidth
///Description:	�����п�
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int col Ҫ�����п����
///				CString ��ֵ
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SetColumnWidth(int col,CString width)
{
	int row = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	range.SetColumnWidth(COleVariant(width));
}

////////////////////////////////////////////////////////////////////////
///Function:	SetRowHeight
///Description:	�����и�
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int row ��Ԫ��������
////////////////////////////////////////////////////////////////////////
CString CExcel2Dlg::GetColumnWidth(int col)
{
	int row = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	VARIANT width = range.GetColumnWidth();
	CString strwidth;
	strwidth.Format(CString((LPCSTR)(_bstr_t)(_variant_t)width));
	return strwidth;
}

////////////////////////////////////////////////////////////////////////
///Function:	GetRowHeight
///Description:	�����и�
///Call:		IndexToString() ��(x,y)������ʽת��Ϊ��A1����ʽ�ַ���
///Input:		int row Ҫ�����иߵ���
///				CString ��ֵ
////////////////////////////////////////////////////////////////////////
CString CExcel2Dlg::GetRowHeight(int row)
{
	int col = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	VARIANT height = range.GetRowHeight();
	CString strheight;
	strheight.Format(CString((LPCSTR)(_bstr_t)(_variant_t)height));
	return strheight;
}


////////////////////////////////////////////////////////////////////////
///Function:	IndexToString
///Description:	�õ��ĵ�Ԫ����EXCEL�еĶ�λ�����ַ���
///Input:		int row ��Ԫ��������
///				int col ��Ԫ��������
///Return:		CString ��Ԫ����EXCEL�еĶ�λ�����ַ���
////////////////////////////////////////////////////////////////////////
CString CExcel2Dlg::IndexToString( int row, int col ) 
{ 
	CString strResult;
	if( col > 26 ) 
	{ 
		strResult.Format(_T("%c%c%d"),'A' + (col-1)/26-1,'A' + (col-1)%26,row);
	} 
	else 
	{ 
		strResult.Format(_T("%c%d"), 'A' + (col-1)%26,row);
	} 
	return strResult;
} 

////////////////////////////////////////////////////////////////////////
///Function:	LastLineIndex
///Description:	�õ�����ܵ�һ�����е�����
///Return:		int ���е�������
////////////////////////////////////////////////////////////////////////
int CExcel2Dlg::LastLineIndex() 
{ 
	int i,j,flag=0;
	CString str;
	for(i=1;;i++)
	{
		flag = 0;
		//����ͳ�ƣ���Ϊǰ�ж�û�����ݼ�Ϊ����
		for(j=1;j<=5;j++)
		{
			str.Format(_T("%s"),this->GetCellValue(i,j));
			if(str.Compare(_T(""))!=0)
			{
				flag = 1;
				break;
			}
			
		}
		if(flag==0)
			return i;
		
	}
}
 

