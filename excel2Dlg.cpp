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
	
    get_string_by_same_value2();
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
 

