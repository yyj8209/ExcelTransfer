
// ExcelTransferDlg.h : ͷ�ļ�

//
#include <afxdisp.h> 
#include "CApplication.h"
#include "CWorkbooks.h"
#include "CWorkbook.h"
#include "CWorksheets.h"
#include "CWorksheet.h"
#include "CRange.h"

#pragma once


// CExcelTransferDlg �Ի���
class CExcelTransferDlg : public CDialogEx
{
// ����
public:
	CExcelTransferDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELTRANSFER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()


public:
	CString strSrcDir;     //Դ�ļ���
	CString strSmpDir;     //�����ļ���
	CString strDstDir;     // ����Ŀ¼

	CApplication Application;            // Excel Ӧ�ó���ӿ�
	CWorkbook Workbook, Workbook1;                  // ������
	CWorkbooks Workbooks, Workbooks1;                // ����������
	CWorksheets Worksheets, Worksheets1;              // ��������
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
