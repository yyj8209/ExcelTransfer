
// ExcelTransferDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExcelTransfer.h"
#include "ExcelTransferDlg.h"
#include "afxdialogex.h"

#include <iostream>
#include <atlstr.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define MAX_LISTITME_LEN 100
// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcelTransferDlg �Ի���



CExcelTransferDlg::CExcelTransferDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelTransferDlg::IDD, pParent)
	, m_edInfo(_T(""))
	, Rows(0)
	, Cols(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

}

void CExcelTransferDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT, m_edInfo);
	DDX_Control(pDX, IDC_LIST_SRC, m_listSrc);
	DDX_Control(pDX, IDC_COMBO_LEFT, m_cbLeft);
	DDX_Control(pDX, IDC_COMBO_RIGHT,m_cbRight);
	DDX_Control(pDX, IDC_COMBO_TOP, m_cbTop);
	DDX_Control(pDX, IDC_COMBO_BOTTOM, m_cbBottom);
	DDX_Control(pDX, IDC_COMBO_FILENAME1, m_cbFilename1);
	DDX_Control(pDX, IDC_COMBO_FILENAME2, m_cbFilename2);
	DDX_Control(pDX, IDC_COMBO_FIELDROW, m_cbFieldRow);
	DDX_Control(pDX, IDC_EDIT_LIST, m_edList);
	
}

BEGIN_MESSAGE_MAP(CExcelTransferDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_IMPORT, &CExcelTransferDlg::OnClickedButtonImport)
	ON_BN_CLICKED(IDC_BUTTON_DONE, &CExcelTransferDlg::OnClickedButtonDone)
	ON_BN_CLICKED(IDC_BUTTON_SAMPLE, &CExcelTransferDlg::OnClickedButtonSample)
	ON_BN_CLICKED(IDC_BUTTON_PREVIEW, &CExcelTransferDlg::OnClickedButtonPreview)
	ON_BN_CLICKED(IDC_BUTTON_SAVE, &CExcelTransferDlg::OnClickedButtonSave)
	ON_BN_CLICKED(IDC_BUTTON_DEL, &CExcelTransferDlg::OnClickedButtonDel)
	ON_BN_CLICKED(IDC_BUTTON_ADD, &CExcelTransferDlg::OnClickedButtonAdd)
	ON_CBN_SELCHANGE(IDC_COMBO_FIELDROW, &CExcelTransferDlg::OnSelchangeComboFieldrow)
	ON_WM_DESTROY()
	ON_CBN_DROPDOWN(IDC_COMBO_FIELDROW, &CExcelTransferDlg::OnDropdownComboFieldrow)

	ON_NOTIFY(NM_DBLCLK, IDC_LIST_SRC, &CExcelTransferDlg::OnDblclkListSrc)
	ON_EN_KILLFOCUS(IDC_EDIT_LIST, &CExcelTransferDlg::OnKillfocusEditList)
END_MESSAGE_MAP()


// CExcelTransferDlg ��Ϣ�������

BOOL CExcelTransferDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������
	InitUI(); 

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CExcelTransferDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelTransferDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CExcelTransferDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExcelTransferDlg::InitUI() 
{
	m_listSrc.SetExtendedStyle(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES); // ����ѡ��������
	m_listSrc.InsertColumn(0, _T("ѡȡ������Ϣ"), LVCFMT_LEFT, 100);
	m_listSrc.InsertColumn(1, _T("��д���������У�"), LVCFMT_LEFT, 110);
	m_listSrc.InsertColumn(2, _T("��д���������У�"), LVCFMT_LEFT, 110);

	m_edList.ShowWindow(SW_HIDE);
}


void CExcelTransferDlg::OnClickedButtonImport()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	strSrcDir = GetWorkDir();
	if (strSrcDir.IsEmpty())
		return;

	m_edInfo.Append(_T("Դ�ļ���") + strSrcDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
	ShellExecute(NULL, "open", strSrcDir, NULL, NULL, SW_SHOWNORMAL); 

	ReadExcelFile(strSrcDir);

	for (int k = 0; k < Cols; k++)
	{
		CString sCol;
		sCol.Format("%c",(int)'A' + k);
		m_cbLeft.AddString(sCol);
		m_cbRight.AddString(sCol);
	}
	for (int k = 0; k < Rows; k++)
	{
		CString sRow;
		sRow.Format("%d",  k + 1);
		m_cbTop.AddString(sRow);
		m_cbBottom.AddString(sRow);
		m_cbFieldRow.AddString(sRow);
	}
	m_cbLeft.SetCurSel(0);
	m_cbTop.SetCurSel(0);
	m_cbRight.SetCurSel(Rows - 1);
	m_cbBottom.SetCurSel(Cols - 1);
	m_cbFieldRow.SetCurSel(0);

}


void CExcelTransferDlg::OnClickedButtonDone()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString  strPathName;
	GetModuleFileName(NULL, strPathName.GetBuffer(256), 256);
	strPathName.ReleaseBuffer(256);
	int nPos = strPathName.ReverseFind('\\');

	if (strDstDir.IsEmpty())
		strDstDir = strPathName.Left(nPos + 1);
}


void CExcelTransferDlg::OnClickedButtonSample()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	strSmpDir = GetWorkDir();
	if (strSmpDir.IsEmpty())
		return;
	m_edInfo.Append(_T("�����ļ���") + strSmpDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
	ShellExecute(NULL, "open", strSmpDir, NULL, NULL, SW_SHOWNORMAL);
}


void CExcelTransferDlg::OnClickedButtonPreview()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	if (strDstDir.IsEmpty())
		OnClickedButtonSave();
	int m = m_cbFilename1.GetCurSel() + 1;
	int n = m_cbTop.GetCurSel() + 1;
	GetCells((long)m, (long)n);
	CString strDstFile = strDstDir + "\\" + strCell;
	m = m_cbFilename2.GetCurSel() + 1;
	GetCells((long)m, (long)n);
	strDstFile = strDstFile + "-" + strCell + ".xls";

	CopyFile((LPCSTR)strSmpDir, (LPCSTR)strDstFile, true);
//	WriteExcelFile();
	WriteExcelFile((long)n, strDstFile);
	ShellExecute(NULL, "open", strDstFile, NULL, NULL, SW_SHOWNORMAL);

}


void CExcelTransferDlg::OnClickedButtonSave()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	BROWSEINFO bi;
	CHAR Buffer[MAX_PATH];

	//��ʼ����ڲ��� bi
	bi.hwndOwner = NULL;
	bi.pidlRoot = NULL;
	bi.pszDisplayName = (LPSTR)Buffer;
	bi.lpszTitle = _T("��ѡ�����Ŀ¼");
	bi.ulFlags = BIF_EDITBOX;
	bi.lpfn = NULL;
	bi.iImage = IDR_MAINFRAME;

	LPITEMIDLIST pIDList = SHBrowseForFolder(&bi); //������ʾѡ��Ի��� 
	//ע���� �������������ڴ� �������ͷ� ��Ҫ�ֶ��ͷ�


	if (pIDList)
	{
		SHGetPathFromIDList(pIDList, (LPSTR)Buffer);
		strDstDir = Buffer;
		//GamePath = Buffer; //���ļ���·��������CString ��������
		//ȡ���ļ���·������Buffer�ռ�
		//GUI_ShowMessage(true, Buffer);
		// �ѱ������ݸ��µ��Ի���
		m_edInfo.Append(_T("����ļ�Ŀ¼��") + strDstDir);
		m_edInfo.Append(_T("\r\n"));
		UpdateData(false);

	}

	CoTaskMemFree(pIDList); //�ͷ�pIDList��ָ���ڴ�ռ�;
	TRACE("%d", pIDList);

}


void CExcelTransferDlg::OnClickedButtonDel()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CExcelTransferDlg::OnClickedButtonAdd()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

// Դ�ļ���ȡ
void CExcelTransferDlg::ReadExcelFile(CString strExcleFilePath)
{

	//strExcleFilePath = strSrcDir.Left(strSrcDir.ReverseFind('\\')) + _T("\\��վ����9��.xls");  //exe����·����ǰ·���µ�Excel�ļ�
	//strExcleFilePath = _T(".\\��վ����9��.xls");  //exe����·����ǰ·���µ�Excel�ļ�
	if (!Application.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("Error!Creat Excel Application Server Faile!"));
	}
	Workbooks = Application.get_Workbooks();   //��ȡ����������
	//COleVariant covTrue((short)TRUE);
	//COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	lpDisp = Workbooks.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	Workbook.AttachDispatch(lpDisp);
	//book = books.Add(covOptional); //��ȡ��ǰ������,��ʹ�ô���䣬��Ϊ�½�һ��EXCEL��
	Worksheets.AttachDispatch(Workbook.get_Worksheets()); //��ȡ��ǰ������ҳ�ļ���
	//�õ���ǰ��Ծsheet
	lpDisp = Workbook.get_ActiveSheet();
	Worksheet.AttachDispatch(lpDisp);
	//�������
	CRange usedRange;
	CRange mRange;
	usedRange.AttachDispatch(Worksheet.get_UsedRange());
	mRange.AttachDispatch(usedRange.get_Rows(), true);
	Rows = mRange.get_Count();
	mRange.AttachDispatch(usedRange.get_Columns(), true);
	Cols = mRange.get_Count();

	usedRange.ReleaseDispatch();
	mRange.ReleaseDispatch();
}

// Ŀ���ļ���ȡ��
void CExcelTransferDlg::WriteExcelFile(long Row, CString strExcleFilePath)
{

	//if (!Application.CreateDispatch(_T("Excel.Application")))
	//{
	//	MessageBox(_T("Error!Creat Excel Application Server Faile!"));
	//}
	Workbooks1 = Application.get_Workbooks();   //��ȡ����������
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	lpDisp1 = Workbooks1.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	Workbook1.AttachDispatch(lpDisp1);
	//book = books.Add(covOptional); //��ȡ��ǰ������,��ʹ�ô���䣬��Ϊ�½�һ��EXCEL��
	Worksheets1.AttachDispatch(Workbook1.get_Worksheets()); //��ȡ��ǰ������ҳ�ļ���
	//�õ���ǰ��Ծsheet
	lpDisp1 = Workbook1.get_ActiveSheet();
	Worksheet1.AttachDispatch(lpDisp1);
	
	//��Դ�ļ���Row�е���Ϣд��Ŀ���ļ���
	//Range1.AttachDispatch(Worksheet1.get_Cells());
	for (int i = 0; i < Cols; i++)
	{
		if (m_listSrc.GetCheck(i))
		{
			int nRowDst, nColDst;
			//CString sRow = m_listSrc.GetItemText(i, 2);
			CString sRange = m_listSrc.GetItemText(i, 1) + m_listSrc.GetItemText(i, 2);
			//char chCol[1];
			//memcpy(chCol, sCol, 1);
			//nRowDst = atoi(sRow);
			//nColDst = int(sCol[0]) - int('A') + 1;  // ����A�ļ��
			GetCells((long)i, Row);
			//Range1.put_Item(COleVariant(long(nColDst)), COleVariant(long(nRowDst)), COleVariant(strCell));
			Range1 = Worksheet1.get_Range(COleVariant(sRange), COleVariant(sRange));
			Range1.put_Value2(COleVariant(strCell));
		}
	}
	/*app1.put_Visible(false);*/
	Application.put_UserControl(TRUE);
	Workbook1.Save();
	//Workbook1.SaveCopyAs(COleVariant(strExcleFilePath));
	Workbook1.put_Saved(true);
	Application.put_DisplayAlerts(false);//�ر�ʱ����̸��ѯ���Ƿ񱣴�ĶԻ���
	Workbooks1.Close();
	Range1.ReleaseDispatch();
	Worksheet1.ReleaseDispatch();
	Worksheets1.ReleaseDispatch();
	Workbook1.ReleaseDispatch(); //�ͷŵ�ǰ������
	Workbooks1.ReleaseDispatch(); //�ͷŹ�������

	//��������������������������������
	//	��Ȩ����������ΪCSDN������liji_digital����ԭ�����£���ѭCC 4.0 BY - SA��ȨЭ�飬ת���븽��ԭ�ĳ������Ӽ���������
	//ԭ�����ӣ�https ://blog.csdn.net/liji_digital/article/details/82320287
}

CString CExcelTransferDlg::GetWorkDir()
{
	CFileDialog dlg(TRUE, _T("Excel�ļ� (*.xls; *.xlsx)|*.xls; *.xlsx|"), NULL,
		NULL, _T("Excel�ļ� (*.xls; *.xlsx)|*.xls; *.xlsx|"), NULL);
	dlg.m_ofn.lpstrTitle = _T("ѡ��Դ�ļ�");
	CString FileName = "";
	if (dlg.DoModal() == IDOK)
	{
		POSITION fileNamesPosition = dlg.GetStartPosition();
		FileName = dlg.GetPathName();
	}
	return FileName;
}


void CExcelTransferDlg::OnSelchangeComboFieldrow()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	int row = m_cbFieldRow.GetCurSel() + 1;
	m_cbFilename1.ResetContent();
	m_cbFilename2.ResetContent();
	m_listSrc.DeleteAllItems();
	for (int i = 0; i < Cols; i++)
	{
		GetCells((long)i+1, (long)row);
		m_cbFilename1.AddString(strCell);
		m_cbFilename2.AddString(strCell);
		m_listSrc.InsertItem(m_listSrc.GetItemCount(), strCell);
	}
	m_cbFilename1.SetCurSel(0);
	m_cbFilename2.SetCurSel(1);
	m_cbTop.SetCurSel(row);

}

void CExcelTransferDlg::OnDropdownComboFieldrow()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	if (m_cbFieldRow.GetCount() < 1)
	{
		this->MessageBox("��ѡ�����ݱ�");
		return;
	}
}

void CExcelTransferDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	// TODO:  �ڴ˴������Ϣ����������
	//�ͷŸ�����ע����˳������ִ�����²��裬Excel�����޷��˳�����������������ῴ����������
	Range.ReleaseDispatch();
	Worksheet.ReleaseDispatch();
	Worksheets.ReleaseDispatch();
	Workbook.ReleaseDispatch(); //�ͷŵ�ǰ������
	Workbooks.ReleaseDispatch(); //�ͷŹ�������
	Application.Quit(); //�˳�EXCEL����
	Application.ReleaseDispatch(); //�ͷ�EXCEL����

}

void CExcelTransferDlg::GetCells(long Col, long Row)
{
	// ��ȡָ����Ԫ�������
	Range.AttachDispatch(Worksheet.get_Cells());
	Range.AttachDispatch(Range.get_Item(COleVariant(Row), COleVariant(Col)).pdispVal);
	vResult = Range.get_Value2();
	if (vResult.vt == VT_BSTR){
		strCell = vResult.bstrVal;
	}
	else if (vResult.vt == VT_R8){
		strCell.Format("%d", vResult.dblVal);
	}
	else{
		strCell = "���ݴ���!";
		//this->MessageBox(strCell+"��������ѡ��");
		//return;
	}

}





void CExcelTransferDlg::OnDblclkListSrc(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	CRect rc;
	nRow = pNMListView->iItem;//���ѡ�е���  
	nCol = pNMListView->iSubItem;//���ѡ����  

	if (pNMListView->iSubItem != 0) //���ѡ���������;  
	{
		m_listSrc.GetSubItemRect(nRow, nCol, LVIR_LABEL, rc);//��������RECT��  
		m_edList.SetParent(&m_listSrc);//ת������Ϊ�б���е�����  
		m_edList.MoveWindow(rc);//�ƶ�Edit��RECT���ڵ�λ��;  
		m_edList.SetWindowText(m_listSrc.GetItemText(nRow, nCol));//���������е�ֵ����Edit�ؼ��У�  
		m_edList.ShowWindow(SW_SHOW);//��ʾEdit�ؼ���  
		m_edList.SetFocus();//����Edit����  
		m_edList.ShowCaret();//��ʾ���  
		m_edList.SetSel(-1);//������ƶ������  
	}

	*pResult = 0;
}


void CExcelTransferDlg::OnKillfocusEditList()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString sCell;
	m_edList.GetWindowText(sCell);    //�õ��û�������µ�����
	switch (nCol)
	{
	case 1:
		if (sCell.GetLength() > 1)
		{
			MessageBox("���벻�ܳ���һ���ַ���");
			sCell.Empty();
		}
		else if (sCell.Compare("z")>0 || sCell.Compare("A")<0)
		{
			MessageBox("�к�ֻ֧������ A~Z ֮����ַ���");
			sCell.Empty();
		}
		sCell.MakeUpper();
		break;
	case 2:
		if (!(sCell.SpanIncluding(_T("0123456789")) == sCell))
		{
			MessageBox("�к�ֻ֧���������֡�");
			sCell.Empty();
		}
		break;
	default:
		MessageBox("������ȷ�ĵ�Ԫ��");

	}
	m_listSrc.SetItemText(nRow, nCol, sCell);   //���ñ༭���������  
	m_edList.ShowWindow(SW_HIDE);                //Ӧ�ر༭��
}

void CExcelTransferDlg::Row2File(int Row, CString strDst)
{
}
void CExcelTransferDlg::InsertCell(int Col, int Row, CString str)
{
}


