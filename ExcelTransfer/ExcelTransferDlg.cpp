
// ExcelTransferDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelTransfer.h"
#include "ExcelTransferDlg.h"
#include "afxdialogex.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CExcelTransferDlg 对话框



CExcelTransferDlg::CExcelTransferDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelTransferDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelTransferDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
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


// CExcelTransferDlg 消息处理程序

BOOL CExcelTransferDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelTransferDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelTransferDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcelTransferDlg::OnClickedButtonImport()
{
	// TODO:  在此添加控件通知处理程序代码

}


void CExcelTransferDlg::OnClickedButtonDone()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonSample()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonPreview()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonSave()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonDel()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonAdd()
{
	// TODO:  在此添加控件通知处理程序代码
}

void CExcelTransferDlg::ReadExcelFile()
{
	strExcleFilePath = strWorkDir + _T("\\ExcelTest.xlsx");  //exe所在路径当前路径下的Excel文件
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
	books = app.get_Workbooks();   //获取工作薄集合

	lpDisp = books.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	book.AttachDispatch(lpDisp);
	//book = books.Add(covOptional); //获取当前工作薄,若使用此语句，则为新建一个EXCEL表
	sheets = book.get_Worksheets(); //获取当前工作薄页的集合
	sheet = sheets.get_Item(COleVariant((short)1)); //获取当前活动页,第一页sheet1

	//1.获取数据1
	range = sheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1"))); //获取单元格
	vResult = COleVariant(range.get_Value2());
	strExcel1 = CString(vResult.bstrVal);

	//1.获取数据2
	range = sheet.get_Range(COleVariant(_T("B1")), COleVariant(_T("B1"))); //获取单元格
	vResult = COleVariant(range.get_Value2());
	strExcel2 = CString(vResult.bstrVal);
	MessageBox(strExcel1);
	MessageBox(strExcel2);

	//释放各对象，注意其顺序，若不执行以下步骤，Excel进程无法退出，打开任务管理器将会看到残留进程
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch(); //释放当前工作薄
	books.ReleaseDispatch(); //释放工作薄集
	app.Quit(); //退出EXCEL程序
	app.ReleaseDispatch(); //释放EXCEL程序
}

void CExcelTransferDlg::GetWorkDir()
{
	TCHAR pFileName[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, pFileName);
	CString dir(pFileName);
	strWorkDir = dir;
}
