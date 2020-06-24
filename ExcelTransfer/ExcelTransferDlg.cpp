
// ExcelTransferDlg.cpp : 实现文件
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
	InitUI(); 

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

void CExcelTransferDlg::InitUI() 
{
	m_listSrc.SetExtendedStyle(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES); // 整行选择、网格线
	m_listSrc.InsertColumn(0, _T("选取基本信息"), LVCFMT_LEFT, 100);
	m_listSrc.InsertColumn(1, _T("填写到样本（列）"), LVCFMT_LEFT, 110);
	m_listSrc.InsertColumn(2, _T("填写到样本（行）"), LVCFMT_LEFT, 110);

	m_edList.ShowWindow(SW_HIDE);
}


void CExcelTransferDlg::OnClickedButtonImport()
{
	// TODO:  在此添加控件通知处理程序代码
	strSrcDir = GetWorkDir();
	if (strSrcDir.IsEmpty())
		return;

	m_edInfo.Append(_T("源文件：") + strSrcDir);
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
	// TODO:  在此添加控件通知处理程序代码
	CString  strPathName;
	GetModuleFileName(NULL, strPathName.GetBuffer(256), 256);
	strPathName.ReleaseBuffer(256);
	int nPos = strPathName.ReverseFind('\\');

	if (strDstDir.IsEmpty())
		strDstDir = strPathName.Left(nPos + 1);
}


void CExcelTransferDlg::OnClickedButtonSample()
{
	// TODO:  在此添加控件通知处理程序代码
	strSmpDir = GetWorkDir();
	if (strSmpDir.IsEmpty())
		return;
	m_edInfo.Append(_T("样本文件：") + strSmpDir);
	m_edInfo.Append(_T("\r\n"));
	UpdateData(false);
	ShellExecute(NULL, "open", strSmpDir, NULL, NULL, SW_SHOWNORMAL);
}


void CExcelTransferDlg::OnClickedButtonPreview()
{
	// TODO:  在此添加控件通知处理程序代码
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
	// TODO:  在此添加控件通知处理程序代码
	BROWSEINFO bi;
	CHAR Buffer[MAX_PATH];

	//初始化入口参数 bi
	bi.hwndOwner = NULL;
	bi.pidlRoot = NULL;
	bi.pszDisplayName = (LPSTR)Buffer;
	bi.lpszTitle = _T("请选择输出目录");
	bi.ulFlags = BIF_EDITBOX;
	bi.lpfn = NULL;
	bi.iImage = IDR_MAINFRAME;

	LPITEMIDLIST pIDList = SHBrowseForFolder(&bi); //调用显示选择对话框 
	//注意下 这个函数会分配内存 但不会释放 需要手动释放


	if (pIDList)
	{
		SHGetPathFromIDList(pIDList, (LPSTR)Buffer);
		strDstDir = Buffer;
		//GamePath = Buffer; //将文件夹路径保存在CString 对象里面
		//取得文件夹路径放置Buffer空间
		//GUI_ShowMessage(true, Buffer);
		// 把变量内容更新到对话框
		m_edInfo.Append(_T("输出文件目录：") + strDstDir);
		m_edInfo.Append(_T("\r\n"));
		UpdateData(false);

	}

	CoTaskMemFree(pIDList); //释放pIDList所指向内存空间;
	TRACE("%d", pIDList);

}


void CExcelTransferDlg::OnClickedButtonDel()
{
	// TODO:  在此添加控件通知处理程序代码
}


void CExcelTransferDlg::OnClickedButtonAdd()
{
	// TODO:  在此添加控件通知处理程序代码
}

// 源文件读取
void CExcelTransferDlg::ReadExcelFile(CString strExcleFilePath)
{

	//strExcleFilePath = strSrcDir.Left(strSrcDir.ReverseFind('\\')) + _T("\\新站歌舞9家.xls");  //exe所在路径当前路径下的Excel文件
	//strExcleFilePath = _T(".\\新站歌舞9家.xls");  //exe所在路径当前路径下的Excel文件
	if (!Application.CreateDispatch(_T("Excel.Application")))
	{
		MessageBox(_T("Error!Creat Excel Application Server Faile!"));
	}
	Workbooks = Application.get_Workbooks();   //获取工作薄集合
	//COleVariant covTrue((short)TRUE);
	//COleVariant covFalse((short)FALSE);
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	lpDisp = Workbooks.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	Workbook.AttachDispatch(lpDisp);
	//book = books.Add(covOptional); //获取当前工作薄,若使用此语句，则为新建一个EXCEL表
	Worksheets.AttachDispatch(Workbook.get_Worksheets()); //获取当前工作薄页的集合
	//得到当前活跃sheet
	lpDisp = Workbook.get_ActiveSheet();
	Worksheet.AttachDispatch(lpDisp);
	//获得行数
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

// 目标文件读取。
void CExcelTransferDlg::WriteExcelFile(long Row, CString strExcleFilePath)
{

	//if (!Application.CreateDispatch(_T("Excel.Application")))
	//{
	//	MessageBox(_T("Error!Creat Excel Application Server Faile!"));
	//}
	Workbooks1 = Application.get_Workbooks();   //获取工作薄集合
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	lpDisp1 = Workbooks1.Open(strExcleFilePath,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	Workbook1.AttachDispatch(lpDisp1);
	//book = books.Add(covOptional); //获取当前工作薄,若使用此语句，则为新建一个EXCEL表
	Worksheets1.AttachDispatch(Workbook1.get_Worksheets()); //获取当前工作薄页的集合
	//得到当前活跃sheet
	lpDisp1 = Workbook1.get_ActiveSheet();
	Worksheet1.AttachDispatch(lpDisp1);
	
	//将源文件第Row行的信息写入目标文件。
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
			//nColDst = int(sCol[0]) - int('A') + 1;  // 列与A的间距
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
	Application.put_DisplayAlerts(false);//关闭时不再谈出询问是否保存的对话框
	Workbooks1.Close();
	Range1.ReleaseDispatch();
	Worksheet1.ReleaseDispatch();
	Worksheets1.ReleaseDispatch();
	Workbook1.ReleaseDispatch(); //释放当前工作薄
	Workbooks1.ReleaseDispatch(); //释放工作薄集

	//————————————————
	//	版权声明：本文为CSDN博主「liji_digital」的原创文章，遵循CC 4.0 BY - SA版权协议，转载请附上原文出处链接及本声明。
	//原文链接：https ://blog.csdn.net/liji_digital/article/details/82320287
}

CString CExcelTransferDlg::GetWorkDir()
{
	CFileDialog dlg(TRUE, _T("Excel文件 (*.xls; *.xlsx)|*.xls; *.xlsx|"), NULL,
		NULL, _T("Excel文件 (*.xls; *.xlsx)|*.xls; *.xlsx|"), NULL);
	dlg.m_ofn.lpstrTitle = _T("选择源文件");
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
	// TODO:  在此添加控件通知处理程序代码
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
	// TODO:  在此添加控件通知处理程序代码
	if (m_cbFieldRow.GetCount() < 1)
	{
		this->MessageBox("请选择数据表。");
		return;
	}
}

void CExcelTransferDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	// TODO:  在此处添加消息处理程序代码
	//释放各对象，注意其顺序，若不执行以下步骤，Excel进程无法退出，打开任务管理器将会看到残留进程
	Range.ReleaseDispatch();
	Worksheet.ReleaseDispatch();
	Worksheets.ReleaseDispatch();
	Workbook.ReleaseDispatch(); //释放当前工作薄
	Workbooks.ReleaseDispatch(); //释放工作薄集
	Application.Quit(); //退出EXCEL程序
	Application.ReleaseDispatch(); //释放EXCEL程序

}

void CExcelTransferDlg::GetCells(long Col, long Row)
{
	// 读取指定单元格的数据
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
		strCell = "数据错误!";
		//this->MessageBox(strCell+"，请重新选择。");
		//return;
	}

}





void CExcelTransferDlg::OnDblclkListSrc(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO:  在此添加控件通知处理程序代码
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
	CRect rc;
	nRow = pNMListView->iItem;//获得选中的行  
	nCol = pNMListView->iSubItem;//获得选中列  

	if (pNMListView->iSubItem != 0) //如果选择的是子项;  
	{
		m_listSrc.GetSubItemRect(nRow, nCol, LVIR_LABEL, rc);//获得子项的RECT；  
		m_edList.SetParent(&m_listSrc);//转换坐标为列表框中的坐标  
		m_edList.MoveWindow(rc);//移动Edit到RECT坐在的位置;  
		m_edList.SetWindowText(m_listSrc.GetItemText(nRow, nCol));//将该子项中的值放在Edit控件中；  
		m_edList.ShowWindow(SW_SHOW);//显示Edit控件；  
		m_edList.SetFocus();//设置Edit焦点  
		m_edList.ShowCaret();//显示光标  
		m_edList.SetSel(-1);//将光标移动到最后  
	}

	*pResult = 0;
}


void CExcelTransferDlg::OnKillfocusEditList()
{
	// TODO:  在此添加控件通知处理程序代码
	CString sCell;
	m_edList.GetWindowText(sCell);    //得到用户输入的新的内容
	switch (nCol)
	{
	case 1:
		if (sCell.GetLength() > 1)
		{
			MessageBox("输入不能超过一个字符。");
			sCell.Empty();
		}
		else if (sCell.Compare("z")>0 || sCell.Compare("A")<0)
		{
			MessageBox("列号只支持输入 A~Z 之间的字符。");
			sCell.Empty();
		}
		sCell.MakeUpper();
		break;
	case 2:
		if (!(sCell.SpanIncluding(_T("0123456789")) == sCell))
		{
			MessageBox("行号只支持输入数字。");
			sCell.Empty();
		}
		break;
	default:
		MessageBox("请点击正确的单元格。");

	}
	m_listSrc.SetItemText(nRow, nCol, sCell);   //设置编辑框的新内容  
	m_edList.ShowWindow(SW_HIDE);                //应藏编辑框
}

void CExcelTransferDlg::Row2File(int Row, CString strDst)
{
}
void CExcelTransferDlg::InsertCell(int Col, int Row, CString str)
{
}


