// excel2Dlg.cpp : implementation file
//

#include "stdafx.h"
#include "excel2.h"
#include "excel2Dlg.h"
#include "Para1.h"
#include "Para2.h"

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
	DDX_Control(pDX, IDC_TAB1, m_tab);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CExcel2Dlg, CDialog)
	//{{AFX_MSG_MAP(CExcel2Dlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, OnSelchangeTab1)
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
		AfxMessageBox(_T("初始化Com失败!")); 
		return;
	}
	
	//验证office文件是否可以正确运行
	
	if( !app.CreateDispatch(_T("Excel.Application")) )
	{
		AfxMessageBox(_T("无法创建Excel应用！"));
		return;
	}
#else
	acess_excel.InitExcel();

#endif
    CFont m_NewFont;
    m_NewFont.CreateFont (14, 0, 0, 0, 800, TRUE, 0, 0, 1, 0, 0, 0, 0, _T("Arial") );

    //m_tab.SetFont (&m_NewFont);

	m_tab.InsertItem(0, "没翻译字符");
	m_tab.InsertItem(1, "相同字符");

	m_para1.Create(IDD_PARA1,GetDlgItem(IDC_TAB1));
    m_para2.Create(IDD_PARA2,GetDlgItem(IDC_TAB1));
	
    //获得IDC_TABTEST客户区大小
	
    CRect rs;
    m_tab.GetClientRect(&rs);
    //调整子对话框在父窗口中的位置
    rs.top += 20;
    rs.bottom -= 20;
    rs.left += 1;
    rs.right -= 2;
    //设置子对话框尺寸并移动到指定位置
    m_para1.MoveWindow(&rs);
    m_para2.MoveWindow(&rs);
    //分别设置隐藏和显示
    m_para1.ShowWindow(1);
    m_para2.ShowWindow(0);
	
    //设置默认的选项卡
    m_tab.SetCurSel(0);
	
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
    MMICOM_SEARCH_FIRST_EQUAL,//查找最前面的一个相同的退出
    MMICOM_SEARCH_LAST_EQUAL,//查找最后的一个相同的退出
    MMICOM_SEARCH_ANY_EQUAL,//查找任意一个相同的退出
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
	BOOLEAN         is_in_list  = FALSE;    // 记录是否列表中有完全一样的姓名
	uint32	        ret_pos         = 0;        // 记录找到的合适位置
    int32          low_pos	    = search_info->start_pos;        // 二分法的前一个位置
    int32          mid_pos	    = 0;        // 二分法的中间位置
   	int32          high_pos	= search_info->end_pos;        // 二分法的后一个位置
	int 	        cmp_result  = 0;  // 两个字符串的比较结果

    if(NULL == search_info || (NULL == func) || (NULL == pos))
    {
        return FALSE;
    } 
	//func = func_t->func;
	// 二分法查找
	//SCI_TRACE_LOW("MMIPB_BinSearch low_pos %d, high_pos %d",low_pos,high_pos);
	while (low_pos <= high_pos)	// 满足二分法的条件
	{
	    // 获得mid_pos
	    mid_pos = ((low_pos + high_pos) >> 1); 
        // 将两个字符串进行比较//
        cmp_result = (*func)(mid_pos,search_info->compare_base_data, list);
        // 根据比较的结果调整查找的范围
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
            // 表示列表中有完全一样的记录。
            // 记录目前的位置，并标记已经找到这样的记录。
        	is_in_list  = TRUE;
        	ret_pos     = mid_pos;
            
            // 查找操作
            // 查找时，需要根据往前找，还是往后找，得到合适的查找范围  
            if (MMICOM_SEARCH_LAST_EQUAL == search_info->search_type)
        	{
                low_pos = (mid_pos + 1);
            }
            else if(MMICOM_SEARCH_FIRST_EQUAL == search_info->search_type)
            {
            	// 很有可能不是列表中满足条件的最前面一条记录，需要继续寻找
                if(mid_pos == 0)
				{
					break;
				}
				high_pos = (mid_pos - 1);	
            }
            else
            {
                //找到任意一个相等的就退出
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
long excel1_compare_column = 1;
long excel2_compare_column = 1;
long file1_save_txt_cols[10] = {0};
long file2_save_txt_cols[10] = {0};
long file1_save_excel_cols[10] = {0};
long file2_save_excel_cols[10] = {0};
long file1_txt_cols_cnt = 0;
long file2_txt_cols_cnt = 0;
long file1_excel_cols_cnt = 0;
long file2_excel_cols_cnt = 0;

 int CompareString(uint32 base_index, void *compare_data, void *list)
{

    int cur_item = (int)compare_data;

    int base_item = order_table[base_index];
    CString value2 ;
    CString value1;

    value1 = acess_excel.GetCellString(base_item, excel1_compare_column);

    return value1.Compare((LPCTSTR)acess_excel.GetCellString(cur_item, excel1_compare_column));
}
 
void insert_table(int end_pos,int cur_item, CStdioFile &myFile, BOOLEAN is_mark,BOOLEAN no_same)
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
        if (is_mark)
        {
            myFile.WriteString("match\n");
        }

        if (no_same)
        {
            return;
        }
    }
    else if (is_mark)
    {
        myFile.WriteString("nomatch\n");
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

    value1 = acess_excel.GetCellString(base_item, excel1_compare_column);
    value2 = acess_excel2.GetCellString(cur_item, excel2_compare_column);

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

void write_ascii_to_unicode_file(FILE *fp, char *buf)
{
    int len = strlen(buf);

    int i = 0;

    for (i=0;i<len;i++)
    {
        fputc(buf[i], fp);
        fputc(0,fp);
    }
}

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
												char *col_flag)
{
	ASSERT(NULL != filename1);
    
    acess_excel.ReleaseExcel();
	acess_excel.InitExcel();
    
    acess_excel.OpenExcelFile(filename1);
    acess_excel.LoadSheet(file1_sheet_index, TRUE);
    
	if (NULL != filename2)
	{
        acess_excel2.OpenExcelFile(filename2);
        acess_excel2.LoadSheet(file2_sheet_index, TRUE);
	}

	if (is_write_file3 && NULL != filename3)
	{
	    acess_excel3.OpenExcelFile(filename3);
	    acess_excel3.LoadSheet(1, TRUE);
	}

    int file1_rows = acess_excel.GetRowCount();
    int file2_rows = -1;
    
    long file1_cur_row = file1_start_row;
    long file2_cur_row = file2_start_row;
    long file3_cur_row = 1;

	if (NULL != filename2)
	{
        file2_rows = acess_excel2.GetRowCount();
	}

    FILE *fp2;
    CStdioFile myFile;
    uint32 i = 1;
    
	if (is_only_mark)
	{

        CFileException fileException;
        if(myFile.Open(save_file,CFile::typeText|CFile::modeCreate|CFile::modeReadWrite),&fileException)
        {

        }
	}
    else
    {
        if (NULL == (fp2=fopen(save_file,"wb")))
        {
            printf("open file error!");
            return;
        }
        
        fputc(0xff,fp2);
        fputc(0xfe,fp2);
    }

    memset(order_table, 0 ,sizeof(order_table));

    order_table[0] = file1_cur_row;
    order_cnt = 1;
    
    for(i=3;i<=file1_rows;i++)
    {
        if (NULL != filename2)
        {
            insert_table(order_cnt, i, myFile, FALSE, FALSE);
        }
        else
        {
            insert_table(order_cnt, i, myFile, TRUE, TRUE);
        }
    }

	if (is_only_mark)
	{
	    for (;file2_cur_row<=file2_rows;file2_cur_row++)
	    {
	        uint32 pos=0;
	        
	        if (0 == find_pos2(&pos, file2_cur_row))
	        {
	            //acess_excel2.SetCellString(file2_cur_row, 1, "xlh__xlh");
	            myFile.WriteString("match");
	        }
			else
			{
	            myFile.WriteString("nomatch");
			}
			myFile.WriteString("\n");
	    }
	}
	else 
	{
	    for (;file2_cur_row<=file2_rows;file2_cur_row++)
	    {
	        uint32 pos=0;
	        uint32 pos2 = 0;
	        
	        if (0 == find_pos(&pos, &pos2,file2_cur_row))
	        {
                TRACE("POS=%d,pos2=%d\n",pos,pos2);
	            for (i=pos2;i<=pos;i++)
	            {
	            	int write_col = 0;

            	#if 0
					for (;write_col<file1_txt_cols_cnt;write_col++)
					{
						myFile.WriteString(acess_excel.GetCellString(order_table[i], file1_save_txt_cols[write_col]));
						myFile.WriteString(col_flag);
					}
					
					for (write_col=0;write_col<file2_txt_cols_cnt;write_col++)
					{
						myFile.WriteString(acess_excel2.GetCellString(file2_cur_row, file2_save_txt_cols[write_col]));
						myFile.WriteString(col_flag);
					}
                    
					myFile.WriteString("\n");
            	#endif
					
                    for (;write_col<file1_txt_cols_cnt;write_col++)
                    {
                        BOOLEAN is_str = acess_excel.GetCell_is_string(order_table[i], file1_save_txt_cols[write_col]);
                        if (is_str)
                        {
                            short * data = acess_excel.GetCellunicode(order_table[i], file1_save_txt_cols[write_col]);
                            fwrite(data, wcslen((const unsigned short *)data)*2, 1, fp2);
                        }
                        else
                        {
                            CString str = acess_excel.GetCellString(order_table[i], file1_save_txt_cols[write_col]);
                            write_ascii_to_unicode_file(fp2,(char *)(LPCSTR)str);
                        }
                        
                        write_ascii_to_unicode_file(fp2,col_flag);
                    }
                    
                    for (write_col=0;write_col<file2_txt_cols_cnt;write_col++)
                    {
                        
                        BOOLEAN is_str = acess_excel2.GetCell_is_string(file2_cur_row, file2_save_txt_cols[write_col]);
                        if (is_str)
                        {
                            short * data = acess_excel.GetCellunicode(file2_cur_row, file2_save_txt_cols[write_col]);
                            fwrite(data, wcslen((const unsigned short *)data)*2, 1, fp2);
                        }
                        else
                        {
                            CString str = acess_excel.GetCellString(file2_cur_row, file2_save_txt_cols[write_col]);
                            write_ascii_to_unicode_file(fp2,(char *)(LPCSTR)str);
                        }
                        write_ascii_to_unicode_file(fp2,col_flag);
                    }
                
                    fputs("\r",fp2);
                    fputc(0,fp2);
                    fputs("\n",fp2);
                    fputc(0,fp2);
                    
	            	if (is_write_file3)
	            	{
						int file3_col = 1;
						
						for (write_col=0;write_col<file1_excel_cols_cnt;write_col++)
						{
		                	acess_excel3.SetCellValue(file3_cur_row, file3_col, acess_excel.GetCellValue(order_table[i], file1_save_excel_cols[write_col]));
							file3_col++;
						}
						
						for (write_col=0;write_col<file2_excel_cols_cnt;write_col++)
						{
		                	acess_excel3.SetCellValue(file3_cur_row, file3_col, acess_excel2.GetCellValue(file2_cur_row, file2_save_excel_cols[write_col]));
							file3_col++;
						}
						
	                	file3_cur_row++;
	            	}

	            }
	        }
        #if 0
            else
            {
                int write_col = 0;
                for (;write_col<file1_txt_cols_cnt;write_col++)
                {
                    myFile.WriteString(acess_excel2.GetCellString(file2_cur_row, file1_save_txt_cols[write_col]));
                    myFile.WriteString(col_flag);
                }
                
                myFile.WriteString("\n");
            }
        #endif
			
	    }
		
	}

	if (!is_only_mark)
	{
        fclose(fp2);
	}
    
	if (is_write_file3)
	{
    	acess_excel3.SaveasXSLFile(acess_excel3.GetOpenFileName());
	}
}

void find_rows_no_translate(const char *filename1, 
								const char *save_file,
								long file1_sheet_index,
								long file1_start_row, 
								long file1_col,
								char *col_flag)
{
	ASSERT(NULL != filename1);
    acess_excel.ReleaseExcel();
	acess_excel.InitExcel();
    
    acess_excel.OpenExcelFile(filename1);
    acess_excel.LoadSheet(file1_sheet_index, TRUE);
    

    long file1_rows = acess_excel.GetRowCount();
    
    long file1_cur_row = file1_start_row;

    FILE *fp2;
    CStdioFile myFile;
    uint32 i = 1;
    
    if (NULL == (fp2=fopen(save_file,"wb")))
    {
        printf("open file error!");
        return;
    }
    
    fputc(0xff,fp2);
    fputc(0xfe,fp2);


    for (;file1_cur_row<=file1_rows;file1_cur_row++)
    {
        CString str = acess_excel.GetCellString(file1_cur_row, file1_col);
        if (0 == str.GetLength())
        {
				long write_col = 0;
                for (;write_col<file1_txt_cols_cnt;write_col++)
                {
                    BOOLEAN is_str = acess_excel.GetCell_is_string(file1_cur_row, file1_save_txt_cols[write_col]);
                    if (is_str)
                    {
                        short * data = acess_excel.GetCellunicode(file1_cur_row, file1_save_txt_cols[write_col]);
                        fwrite(data, wcslen((const unsigned short *)data)*2, 1, fp2);
                    }
                    else
                    {
                        CString str = acess_excel.GetCellString(file1_cur_row, file1_save_txt_cols[write_col]);
                        write_ascii_to_unicode_file(fp2,(char *)(LPCSTR)str);
                    }
                    
                    write_ascii_to_unicode_file(fp2,col_flag);
                }
                
            
                fputs("\r",fp2);
                fputc(0,fp2);
                fputs("\n",fp2);
                fputc(0,fp2);
        }
		
    }
		
    fclose(fp2);
}

int set_compare_cols(long file1_col, long file2_col)
{
	excel1_compare_column = file1_col;
	excel2_compare_column = file2_col;

    return 1;
}

int set_out_cols(const char *buf, long *cnt, long *cols)
{
    int len = strlen(buf);
    int digit_len = 0;
    const char *digit_buf;
    int i = 0;

	*cnt = 0;

    if (NULL == buf)
    {
        return 1;
    }
    
    while(i<len)
    {
        if (('0'<= buf[i] && buf[i] <= '9'))
        {
            if (0 == digit_len)
            {
                digit_buf = buf+i;
            }
            digit_len++;
            i++;
        }
        else if (',' == buf[i])
        {
            if (0 < digit_len)
            {
                cols[*cnt] = atoi(digit_buf);
                (*cnt)++;
            }
            
            digit_len = 0;
            i++;
        }
        else
        {
            return 0;
        }
    }

    
    if (0 < digit_len)
    {
        cols[*cnt] = atoi(digit_buf);
        (*cnt)++;
    }

    return 1;
}

int set_out_txt_file1_cols(const char *buf)
{
    return set_out_cols(buf, &file1_txt_cols_cnt, file1_save_txt_cols);
}

int set_out_txt_file2_cols(const char *buf)
{
    return set_out_cols(buf, &file2_txt_cols_cnt, file2_save_txt_cols);
}

int set_out_excel_file1_cols(const char *buf)
{
    return set_out_cols(buf, &file1_excel_cols_cnt, file1_save_excel_cols);
}

int set_out_excel_file2_cols(const char *buf)
{
    return set_out_cols(buf, &file2_excel_cols_cnt, file2_save_excel_cols);
}


void get_string_by_id2(void)
{
	excel1_compare_column = 4;
	excel2_compare_column = 4;
	#if 0
	file1_save_txt_cols[10] = {0};
	file2_save_txt_cols[10] = {0};
	file1_save_excel_cols[10] = {0};
	file2_save_excel_cols[10] = {0};
	#endif
	file1_txt_cols_cnt = 4;
	file2_txt_cols_cnt = 0;
	file1_excel_cols_cnt = 0;
	file2_excel_cols_cnt = 1;
	
	file1_save_txt_cols[0] = 1;
	file1_save_txt_cols[1] = 2;
	file1_save_txt_cols[2] = 3;
	file1_save_txt_cols[3] = 4;
	file1_save_txt_cols[4] = 5;
    
	file2_save_excel_cols[0] = 5;
	file2_save_txt_cols[0] = 5;
	file2_save_txt_cols[1] = 2;
	file2_save_txt_cols[2] = 3;
	file2_save_txt_cols[3] = 4;
	file2_save_txt_cols[4] = 5;
    
    long file1_sheet = 1;
    long file1_col   = 33;//30
    long file1_start_row = 2;
    long file2_start_row = 1;
    BOOLEAN is_only_mark = FALSE;
    BOOLEAN is_save_file3 = TRUE;
    
	find_rows_no_translate("D:\\excel24\\str_table1.xls",
        "result.txt", 
        file1_sheet,
        file1_start_row, 
        file1_col,
        ">");
}

void get_string_by_id(void)
{
	excel1_compare_column = 4;
	excel2_compare_column = 4;
	#if 0
	file1_save_txt_cols[10] = {0};
	file2_save_txt_cols[10] = {0};
	file1_save_excel_cols[10] = {0};
	file2_save_excel_cols[10] = {0};
	#endif
	file1_txt_cols_cnt = 4;
	file2_txt_cols_cnt = 0;
	file1_excel_cols_cnt = 0;
	file2_excel_cols_cnt = 1;
	
	file1_save_txt_cols[0] = 1;
	file1_save_txt_cols[1] = 2;
	file1_save_txt_cols[2] = 3;
	file1_save_txt_cols[3] = 4;
	file1_save_txt_cols[4] = 5;
    
	file2_save_excel_cols[0] = 5;
	file2_save_txt_cols[0] = 5;
	file2_save_txt_cols[1] = 2;
	file2_save_txt_cols[2] = 3;
	file2_save_txt_cols[3] = 4;
	file2_save_txt_cols[4] = 5;
    
    long file1_sheet = 1;
    long file2_sheet = 1;
    long file1_start_row = 2;
    long file2_start_row = 1;
    BOOLEAN is_only_mark = FALSE;
    BOOLEAN is_save_file3 = TRUE;
    
	find_rows_in_file2_same_with_file1("D:\\excel24\\str_table.xls",
        "D:\\excel24\\mmm.xls",
        NULL,
        "result.txt", 
        file1_sheet,
        file2_sheet,
        file1_start_row, 
        file2_start_row,
        is_only_mark,
        is_save_file3,
        ">");
}

UINT ThreadFun(LPVOID pParam)
{  //线程要调用的函数
    
    return 0;
}

void CExcel2Dlg::OnOK() 
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();

#if 0//def ACCESS_EXCEL_SIMPLE
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
    //get_string_by_id();
    get_string_by_id2();

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
		AfxMessageBox(_T("文件不存在"));
		return false;
	}
	LPDISPATCH lpDisp; //接口指针
	books=app.GetWorkbooks();
	lpDisp = books.Open(filename,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2, covOptional2, covOptional2,
		covOptional2, covOptional2
		);										//与office 2000的不同，是个参数的，直接在后面加了两个covOptional2成功了
	book.AttachDispatch(lpDisp);
	sheets=book.GetSheets();
	sheet=sheets.GetItem(COleVariant((short)1));		//与的不同，是个参数的，直接在后面加了两个covOptional2成功了
	return true;
}
void CExcel2Dlg::NewExcelBook()
{
	books=app.GetWorkbooks();
	book=books.Add(covOptional2);
	sheets=book.GetSheets();
	sheet=sheets.GetItem(COleVariant((short)1));		//与的不同，是个参数的，直接在后面加了两个covOptional2成功了
}

////////////////////////////////////////////////////////////////////////
///Function:	OpenExcelApp
///Description:	打开应用程序（要注意以后如何识别用户要打开的是哪个文件）
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::OpenExcelApp(void)
{
	app.SetVisible(TRUE);
	app.SetUserControl(TRUE);
}

////////////////////////////////////////////////////////////////////////
///Function:	SaveExcel
///Description:	用于打开数据文件，续存数据后直接保存
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SaveExcel(void)
{
	book.SetSaved(TRUE);
}

////////////////////////////////////////////////////////////////////////
///Function:	SaveAsExcel
///Description:	保存excel文件
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
///Description:	修改单元格内的值
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int row 单元格所在行
///				int col 单元格所在列
///				int Align		对齐方式默认为居中
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
///Description:	得到的单元格中的值
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int row 单元格所在行
///				int col 单元格所在列
///Return:		CString 单元格中的值
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
///Description:	设置行高
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int row 单元格所在行
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SetRowHeight(int row, CString height)
{
	int col = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	range.SetRowHeight(COleVariant(height));
}
////////////////////////////////////////////////////////////////////////
///Function:	SetColumnWidth
///Description:	设置列宽
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int col 要设置列宽的列
///				CString 宽值
////////////////////////////////////////////////////////////////////////
void CExcel2Dlg::SetColumnWidth(int col,CString width)
{
	int row = 1;
	range=sheet.GetRange(COleVariant(IndexToString(row,col)),COleVariant(IndexToString(row,col)));
	range.SetColumnWidth(COleVariant(width));
}

////////////////////////////////////////////////////////////////////////
///Function:	SetRowHeight
///Description:	设置行高
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int row 单元格所在行
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
///Description:	设置行高
///Call:		IndexToString() 从(x,y)坐标形式转化为“A1”格式字符串
///Input:		int row 要设置行高的行
///				CString 宽值
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
///Description:	得到的单元格在EXCEL中的定位名称字符串
///Input:		int row 单元格所在行
///				int col 单元格所在列
///Return:		CString 单元格在EXCEL中的定位名称字符串
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
///Description:	得到表格总第一个空行的索引
///Return:		int 空行的索引号
////////////////////////////////////////////////////////////////////////
int CExcel2Dlg::LastLineIndex() 
{ 
	int i,j,flag=0;
	CString str;
	for(i=1;;i++)
	{
		flag = 0;
		//粗略统计，认为前列都没有数据即为空行
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
 


void CExcel2Dlg::OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	int CurSel = m_tab.GetCurSel();
	
    switch(CurSel)
    {
	case 0:
		m_para1.ShowWindow(true);
		m_para2.ShowWindow(false);
		break;
	case 1:
		m_para1.ShowWindow(false);
		m_para2.ShowWindow(true);
		break;
	default:
		;
		
		*pResult = 0;
    }
}
