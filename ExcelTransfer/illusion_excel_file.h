#ifndef ILLUSION_EXCEL_FILE_H_H
#define ILLUSION_EXCEL_FILE_H_H
/******************************************************************************************
���ļ�copy��http://blog.csdn.net/fullsail/article/details/8449448
���ڴ��룬�ҽ��������˱�����󣬶������������û�и��ġ�ʱ�䣺2016-03-20��
******************************************************************************************/

//OLE��ͷ�ļ�
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"

///
///����OLE�ķ�ʽ��EXCEL��д��
class IllusionExcelFile
{
public:
	//���캯������������
	IllusionExcelFile();
	virtual ~IllusionExcelFile();

public:
	///
	void ShowInExcel(BOOL bShow);

	///���һ��CELL�Ƿ����ַ���
	BOOL    IsCellString(long iRow, long iColumn);
	///���һ��CELL�Ƿ�����ֵ
	BOOL    IsCellInt(long iRow, long iColumn);

	///�õ�һ��CELL��String
	CString GetCellString(long iRow, long iColumn);
	///�õ�����
	int     GetCellInt(long iRow, long iColumn);
	///�õ�double������
	double  GetCellDouble(long iRow, long iColumn);

	///ȡ���е�����
	int GetRowCount();
	///ȡ���е�����
	int GetColumnCount();

	///ʹ��ĳ��sheet��
	BOOL LoadSheet(long table_index, BOOL pre_load = FALSE);
	///ͨ������ʹ��ĳ��sheet��
	BOOL LoadSheet(LPCTSTR sheet, BOOL pre_load = FALSE);
	///ͨ�����ȡ��ĳ��Sheet������
	CString GetSheetName(long table_index);

	///�õ�Sheet������
	int GetSheetCount();

	///���ļ�
	BOOL OpenExcelFile(LPCTSTR file_name);
	///�رմ򿪵�Excel �ļ�����ʱ���EXCEL�ļ���Ҫ
	void CloseExcelFile(BOOL if_save = FALSE);
	//����Ϊһ��EXCEL�ļ�
	void SaveasXSLFile(const CString &xls_file);
	///ȡ�ô��ļ�������
	CString GetOpenFileName();
	///ȡ�ô�sheet������
	CString GetLoadSheetName();

	///д��һ��CELLһ��int
	void SetCellInt(long irow, long icolumn, int new_int);
	///д��һ��CELLһ��string
	void SetCellString(long irow, long icolumn, CString new_string);

public:
	///��ʼ��EXCEL OLE
	static BOOL InitExcel();
	///�ͷ�EXCEL�� OLE
	static void ReleaseExcel();
	///ȡ���е����ƣ�����27->AA
	static char *GetColumnName(long iColumn);

protected:
	///�򿪵�EXCEL�ļ�����
	CString       open_excel_file_;

	///EXCEL BOOK���ϣ�������ļ�ʱ��
	CWorkbooks    excel_books_;
	///��ǰʹ�õ�BOOK����ǰ�������ļ�
	CWorkbook     excel_work_book_;
	///EXCEL��sheets����
	CWorksheets   excel_sheets_;
	///��ǰʹ��sheet
	CWorksheet    excel_work_sheet_;
	///��ǰ�Ĳ�������
	CRange        excel_current_range_;


	///�Ƿ��Ѿ�Ԥ������ĳ��sheet������
	BOOL          already_preload_;
	///Create the SAFEARRAY from the VARIANT ret.
	COleSafeArray ole_safe_array_;

protected:

	///EXCEL�Ľ���ʵ��
	static CApplication excel_application_;

protected:
	//Ԥ�ȼ���
	void PreLoadSheet();
};
#endif//ILLUSION_EXCEL_FILE_H_H