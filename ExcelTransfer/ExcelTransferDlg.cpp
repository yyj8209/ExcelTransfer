
// ExcelTransferDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExcelTransfer.h"
#include "ExcelTransferDlg.h"
#include "afxdialogex.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


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
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelTransferDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT, m_edInfo);
	DDX_Control(pDX, IDC_LIST_SRC, m_listSrc);
	DDX_Control(pDX, IDC_LIST_DEST, m_listDst);
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
	InitListCtrl();

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

void CExcelTransferDlg::InitListCtrl()
{
	m_listSrc.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); // ����ѡ��������
	m_listDst.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES); // ����ѡ��������
	m_listSrc.InsertColumn(0, _T("�ֶε���Ϣ"), LVCFMT_LEFT, 100);
	m_listSrc.InsertItem(0,_T("�ֶε���Ϣ") );
	m_listDst.InsertColumn(0, _T("�ֶε���Ϣ"), LVCFMT_LEFT, 100);
	m_listDst.InsertColumn(1, _T("���±��λ��"), LVCFMT_LEFT, 100);

}


void CExcelTransferDlg::OnClickedButtonImport()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	strSrcDir = GetWorkDir();
	//GetDlgCtrl(IDC_EDIT)->
	m_edInfo.Append(_T("Դ�ļ���") + strSrcDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
	ShellExecute(NULL, "open", strSrcDir, NULL, NULL, SW_SHOWNORMAL); 
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
	//GetDlgCtrl(IDC_EDIT)->
	m_edInfo.Append(_T("�����ļ���") + strSmpDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
	ShellExecute(NULL, "open", strSmpDir, NULL, NULL, SW_SHOWNORMAL);
}


void CExcelTransferDlg::OnClickedButtonPreview()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
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
	bi.lpszTitle = _T("�ļ���·��ѡ��");
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

	}

	CoTaskMemFree(pIDList); //�ͷ�pIDList��ָ���ڴ�ռ�;
	TRACE("%d", pIDList);

	// �ѱ������ݸ��µ��Ի���
	m_edInfo.Append(_T("����ļ�Ŀ¼��") + strDstDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
}


void CExcelTransferDlg::OnClickedButtonDel()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void CExcelTransferDlg::OnClickedButtonAdd()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

void CExcelTransferDlg::ReadExcelFile()
{

	strExcleFilePath = strSrcDir.Left(strSrcDir.ReverseFind('\\')) + _T("\\��վ����9��.xls");  //exe����·����ǰ·���µ�Excel�ļ�
	//strExcleFilePath = _T(".\\��վ����9��.xls");  //exe����·����ǰ·���µ�Excel�ļ�
	CApplication app;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	LPDISPATCH lpDisp;
	COleVariant vResult;
	COleVariant
		covTrue((short)TRUE),
		covFalse((short)FALSE),
		covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("Error!Creat Excel Application Server Faile!"));
	}
	books = app.get_Workbooks();   //��ȡ����������

	lpDisp = books.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	book.AttachDispatch(lpDisp);
	//book = books.Add(covOptional); //��ȡ��ǰ������,��ʹ�ô���䣬��Ϊ�½�һ��EXCEL��
	sheets = book.get_Worksheets(); //��ȡ��ǰ������ҳ�ļ���
	sheet = sheets.get_Item(COleVariant((short)1)); //��ȡ��ǰ�ҳ,��һҳsheet1

	//1.��ȡ����1
	range = sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1"))); //��ȡ��Ԫ��
	vResult = COleVariant(range.get_Value2());
	strExcel1 = CString(vResult.bstrVal);

	//1.��ȡ����2
	range = sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1"))); //��ȡ��Ԫ��
	vResult = COleVariant(range.get_Value2());
	strExcel2 = CString(vResult.bstrVal);
	MessageBox(strExcel1);
	MessageBox(strExcel2);

	//�ͷŸ�����ע����˳������ִ�����²��裬Excel�����޷��˳�����������������ῴ����������
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch(); //�ͷŵ�ǰ������
	books.ReleaseDispatch(); //�ͷŹ�������
	app.Quit(); //�˳�EXCEL����
	app.ReleaseDispatch(); //�ͷ�EXCEL����
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
