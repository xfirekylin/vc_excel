#pragma once

//OLE��ͷ�ļ�
#include "excel.h"

///
///����OLE�ķ�ʽ��EXCEL��д��
class IllusionExcelFile
{

public:

    //���캯������������
    IllusionExcelFile();
    virtual ~IllusionExcelFile();

protected:
    ///�򿪵�EXCEL�ļ�����
    CString       open_excel_file_;

    ///EXCEL BOOK���ϣ�������ļ�ʱ��
    Workbooks    excel_books_; 
    ///��ǰʹ�õ�BOOK����ǰ������ļ�
    _Workbook     excel_work_book_; 
    ///EXCEL��sheets����
    Worksheets   excel_sheets_; 
    ///��ǰʹ��sheet
    _Worksheet    excel_work_sheet_; 
    ///��ǰ�Ĳ�������
    Range        excel_current_range_; 


    ///�Ƿ��Ѿ�Ԥ������ĳ��sheet������
    BOOL          already_preload_;
    ///Create the SAFEARRAY from the VARIANT ret.
    COleSafeArray ole_safe_array_;

protected:

    ///EXCEL�Ľ���ʵ��
    static _Application excel_application_;
public:
    
    ///
    void ShowInExcel(BOOL bShow);

    ///���һ��CELL�Ƿ����ַ���
    BOOL    IsCellString(long iRow, long iColumn);
    ///���һ��CELL�Ƿ�����ֵ
    BOOL    IsCellInt(long iRow, long iColumn);

    ///�õ�һ��CELL��String
    CString GetCellString(long iRow, long iColumn);

    
    ///�õ�һ��CELL��ֵ
    COleVariant GetCellValue(long iRow, long iColumn);
    
    void SetCellValue(long irow, long icolumn,COleVariant value);
    
    ///�õ�����
    int     GetCellInt(long iRow, long iColumn);
    ///�õ�double������
    double  GetCellDouble(long iRow, long iColumn);

    ///ȡ���е�����
    int GetRowCount();
    ///ȡ���е�����
    int GetColumnCount();

    ///ʹ��ĳ��shet��shit��shit
    BOOL LoadSheet(long table_index,BOOL pre_load = FALSE);
    ///ͨ������ʹ��ĳ��sheet��
    BOOL LoadSheet(const char* sheet,BOOL pre_load = FALSE);
    ///ͨ�����ȡ��ĳ��Sheet������
    CString GetSheetName(long table_index);

    ///�õ�Sheet������
    int GetSheetCount();

    ///���ļ�
    BOOL OpenExcelFile(const char * file_name);
    ///�رմ򿪵�Excel �ļ�����ʱ���EXCEL�ļ���Ҫ
    void CloseExcelFile(BOOL if_save = FALSE);
    //���Ϊһ��EXCEL�ļ�
    void SaveasXSLFile(const CString &xls_file);
    ///ȡ�ô��ļ�������
    CString GetOpenFileName();
    ///ȡ�ô�sheet������
    CString GetLoadSheetName();
    
    ///д��һ��CELLһ��int
    void SetCellInt(long irow, long icolumn,int new_int);
    ///д��һ��CELLһ��string
    void SetCellString(long irow, long icolumn,CString new_string);
    
public:
    ///��ʼ��EXCEL OLE
    static BOOL InitExcel();
    ///�ͷ�EXCEL�� OLE
    static void ReleaseExcel();
    ///ȡ���е����ƣ�����27->AA
    static char *GetColumnName(long iColumn);
    
protected:

    //Ԥ�ȼ���
    void PreLoadSheet();
};



