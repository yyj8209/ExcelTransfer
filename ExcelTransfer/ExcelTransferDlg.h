
// ExcelTransferDlg.h : 头文件

//
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
	CString strINI1; //读取INI文件内容CString变量1
	CString strINI2; //读取INI文件内容CString变量2
	CString strINI3; //读取INI文件内容CString变量3
	CString strINI4; //读取INI文件内容CString变量4
	CString strExcel1; //读取Excel文件内容CString变量1
	CString strExcel2; //读取Excel文件内容CString变量2
	CString strINIFilePath;   //用于保存数据源INI文件路径
	CString strExcleFilePath;  //用于保存数据源Excel文件路径
	CString strOutputExcleFilePath;  //用于保存输出Excel文件路径
	CString strWorkDir;     //用于保存exe所在路径
public:
	void ReadExcelFile();
	void GetWorkDir();
public:
	afx_msg void OnClickedButtonImport();
	afx_msg void OnClickedButtonDone();
	afx_msg void OnClickedButtonSample();
	afx_msg void OnClickedButtonPreview();
	afx_msg void OnClickedButtonSave();
	afx_msg void OnClickedButtonDel();
	afx_msg void OnClickedButtonAdd();
};
