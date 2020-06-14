
// ExcelTransferDlg.h : ͷ�ļ�

//
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
	CString strINI1; //��ȡINI�ļ�����CString����1
	CString strINI2; //��ȡINI�ļ�����CString����2
	CString strINI3; //��ȡINI�ļ�����CString����3
	CString strINI4; //��ȡINI�ļ�����CString����4
	CString strExcel1; //��ȡExcel�ļ�����CString����1
	CString strExcel2; //��ȡExcel�ļ�����CString����2
	CString strINIFilePath;   //���ڱ�������ԴINI�ļ�·��
	CString strExcleFilePath;  //���ڱ�������ԴExcel�ļ�·��
	CString strOutputExcleFilePath;  //���ڱ������Excel�ļ�·��
	CString strWorkDir;     //���ڱ���exe����·��
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
