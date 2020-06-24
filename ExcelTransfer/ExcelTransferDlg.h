
// ExcelTransferDlg.h : 头文件

//
#include <afxdisp.h> 
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"

#pragma once


// CExcelTransferDlg 对话框
class CExcelTransferDlg : public CDialogEx
{
// 构造
public:
	CExcelTransferDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELTRANSFER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()


public:
	CString strSrcDir;     //源文件名
	CString strSmpDir;     //样本文件名
	CString strDstDir;     // 保存目录

	CApplication Application;            // Excel 应用程序接口
	CWorkbook Workbook, Workbook1;                  // 工作簿
	CWorkbooks Workbooks, Workbooks1;                // 工作簿集合
	CWorksheets Worksheets, Worksheets1;              // 工作表集合
	CWorksheet Worksheet, Worksheet1;
	CRange Range, Range1;
	LPDISPATCH lpDisp, lpDisp1;
	COleVariant vResult, vResult1;

	int Rows, Cols;   // for Excel
	int Rows1, Cols1;   // for Excel
	CString strCell;
	int nRow, nCol;   // for List

public:
	void ReadExcelFile(CString strExcleFilePath);
	void WriteExcelFile(long Row, CString strExcleFilePath);
	void WriteExcelFile();
	CString GetWorkDir();
	void InitUI();
	void GetCells(long Col, long Row);
	void Row2File(int Row, CString strDst);
	void InsertCell(int Col, int Row, CString str);

public:
	afx_msg void OnClickedButtonImport();
	afx_msg void OnClickedButtonDone();
	afx_msg void OnClickedButtonSample();
	afx_msg void OnClickedButtonPreview();
	afx_msg void OnClickedButtonSave();
	afx_msg void OnClickedButtonDel();
	afx_msg void OnClickedButtonAdd();
	CString m_edInfo;
	CListCtrl m_listSrc;
	CComboBox m_cbLeft;
	CComboBox m_cbRight;
	CComboBox m_cbTop;
	CComboBox m_cbBottom;
	CComboBox m_cbFilename1;
	CComboBox m_cbFilename2;
	CComboBox m_cbFieldRow;
	CEdit m_edList;
	afx_msg void OnSelchangeComboFieldrow();
	afx_msg void OnDestroy();
	afx_msg void OnDropdownComboFieldrow();

	afx_msg void OnDblclkListSrc(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnKillfocusEditList();

};
