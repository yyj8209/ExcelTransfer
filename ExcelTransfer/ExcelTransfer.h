
// ExcelTransfer.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelTransferApp: 
// �йش����ʵ�֣������ ExcelTransfer.cpp
//

class CExcelTransferApp : public CWinApp
{
public:
	CExcelTransferApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelTransferApp theApp;