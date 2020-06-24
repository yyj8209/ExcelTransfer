// EditListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "CADIPlatform.h"
#include "EditListCtrl.h"
#include "CodeDBHelper.h"
#include "PersistXmlParser.h"

// CEditListCtrl

IMPLEMENT_DYNAMIC(CEditListCtrl, CListCtrl)

CEditListCtrl::CEditListCtrl()
{

}

CEditListCtrl::~CEditListCtrl()
{
}


BEGIN_MESSAGE_MAP(CEditListCtrl, CListCtrl)
	ON_WM_LBUTTONDBLCLK()
	ON_CBN_SELCHANGE(IDC_CELL_COMBOBOX, &CEditListCtrl::OnCbnSelchangeCbCellComboBox)
	ON_WM_HSCROLL()
	ON_WM_VSCROLL()
END_MESSAGE_MAP()



// CEditListCtrl message handlers

// 当编辑完成后同步数据
void CEditListCtrl::DisposeEdit(void)
{
	int nIndex = GetSelectionMark();
	//同步更新处理数据
	CString sLable;
	_PersistItem* pPersistItem = (_PersistItem*)GetItemData(m_nRow);
	if (NULL == pPersistItem) return;

	if (0 == m_nCol) //edit控件
	{
		CString szLable;
		m_edit.GetWindowText(szLable);
		SetItemText(m_nRow, m_nCol, szLable);
		pPersistItem->szPersistence = szLable;

		m_edit.ShowWindow(SW_HIDE);
	}
	else if ((1 == m_nCol) || (2 == m_nCol))
	{
		m_comboBox.ShowWindow(SW_HIDE);
	}

}

// 设置当前窗口的风格
void CEditListCtrl::SetStyle(void)
{
	LONG lStyle;
	lStyle = GetWindowLong(m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; //清除显示方式
	lStyle |= LVS_REPORT; //list模式
	lStyle |= LVS_SINGLESEL; //单选
	SetWindowLong(m_hWnd, GWL_STYLE, lStyle);

	//扩展模式
	DWORD dwStyle = GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT; //选中某行使其整行高亮
	dwStyle |= LVS_EX_GRIDLINES; //网格线
	SetExtendedStyle(dwStyle); //设置扩展风格
}

void CEditListCtrl::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	CListCtrl::OnLButtonDblClk(nFlags, point);

	LVHITTESTINFO info;
	info.pt = point;
	info.flags = LVHT_ONITEMLABEL;

	if (SubItemHitTest(&info) >= 0)
	{
		m_nRow = info.iItem;
		m_nCol = info.iSubItem;

		CRect rect;
		GetSubItemRect(m_nRow, m_nCol, LVIR_LABEL, rect);

		CString strValue;
		strValue = GetItemText(m_nRow, m_nCol);

		_PersistItem* pPersistItem = (_PersistItem*)GetItemData(m_nRow);
		if (NULL == pPersistItem) return;
		//Edit单元格
		if ((0 == m_nCol))
		{
			if (pPersistItem->nFlag & 0x01)
			{
				MB_PROMPT(RES_LOAD_STRING(IDS_ST_MODIFIYPKATTRI_20170502_01));
				return;
			}

			if (NULL == m_edit)
			{
				//创建Edit控件
				// ES_WANTRETURN 使多行编辑器接收回车键输入并换行。如果不指定该风格,按回车键会选择缺省的命令按钮,这往往会导致对话框的关闭。
				m_edit.Create(WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | ES_WANTRETURN | ES_LEFT,
					CRect(0, 0, 0, 0), this, IDC_CELL_EDIT);
			}
			m_edit.MoveWindow(rect);
			m_edit.SetWindowText(strValue);
			m_edit.ShowWindow(SW_SHOW);
			m_edit.SetSel(0, -1);
			m_edit.SetFocus();
			UpdateWindow();
		}
		//ComboBox单元格
		if ((1 == m_nCol))
		{
			if (NULL == m_comboBox)
			{
				//创建Combobox控件
				//CBS_DROPDOWNLIST : 下拉式组合框，但是输入框内不能进行输入
				//WS_CLIPCHILDREN  :  其含义就是，父窗口不对子窗口区域进行绘制。默认情况下父窗口会对子窗口背景是进行绘制的，但是如果父窗口设置了WS_CLIPCHILDREN属性，父亲窗口不在对子窗口背景绘制.
				m_comboBox.Create(WS_CHILD | WS_VISIBLE | CBS_SORT | WS_BORDER | CBS_DROPDOWNLIST | WS_VSCROLL | WS_TABSTOP | CBS_AUTOHSCROLL,
					CRect(0, 0, 0, 0), this, IDC_CELL_COMBOBOX);
			}
			m_comboBox.MoveWindow(rect);

			m_comboBox.ShowWindow(SW_SHOW);
			//TODO comboBox 初始化
			LoadPlmAttri();
			//当前cell的值
			CString strCellValue = GetItemText(m_nRow, m_nCol);
			int nIndex = m_comboBox.FindStringExact(0, strCellValue);
			if (CB_ERR == nIndex)
				m_comboBox.SetCurSel(0); //TODO 设置为当前值的Item
			else
				m_comboBox.SetCurSel(nIndex);

			m_comboBox.SetFocus();
			UpdateWindow();
		}
	}
}


void CEditListCtrl::OnCbnSelchangeCbCellComboBox()
{
	//同步更新处理数据
	_PersistItem* pPersistItem = (_PersistItem*)GetItemData(m_nRow);
	if (NULL == pPersistItem) return;

	if (1 == m_nCol) //combobox控件
	{
		CString szLable;
		int nindex = m_comboBox.GetCurSel();
		m_comboBox.GetLBText(nindex, szLable);
		ATTRIFIELD* pAttriField = (ATTRIFIELD*)m_comboBox.GetItemData(nindex);
		if (NULL == pAttriField)
			return;

		//持久化的处理
		if (0 == pAttriField->szComment.CompareNoCase(ATTRIBUTE_PERSISTENCE))
		{
			pPersistItem->szPersistence = pAttriField->mapPersistenceValue[ATTRIBUTE_VALUE];
			SetItemText(m_nRow, m_nCol - 1, pPersistItem->szPersistence);
		}

		pPersistItem->szUcProperty = pAttriField->szPLMToXML;
		pPersistItem->szXmlOption = pAttriField->szXMLOpention;
		SetItemText(m_nRow, m_nCol, pAttriField->szDetail);
		m_comboBox.ShowWindow(SW_HIDE);
	}
}

//发送失效消息
void CEditListCtrl::SendInvalidateMsg()
{
	//Edit单元格
	if ((0 == m_nCol))
	{

		if (NULL == m_edit.m_hWnd) return;
		BOOL bShow = m_edit.IsWindowVisible();
		if (bShow)
			::SendMessage(m_edit.m_hWnd, WM_KILLFOCUS, (WPARAM)0, (LPARAM)0);

	}
	else if ((1 == m_nCol)) //combobox
	{
		if (NULL == m_comboBox.m_hWnd) return;

		BOOL bShow = m_comboBox.IsWindowVisible();
		if (bShow)
			::SendMessage(m_comboBox.m_hWnd, WM_KILLFOCUS, (WPARAM)0, (LPARAM)0);
	}
}

void CEditListCtrl::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	SendInvalidateMsg();

	CListCtrl::OnHScroll(nSBCode, nPos, pScrollBar);
}

void CEditListCtrl::OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	SendInvalidateMsg();

	CListCtrl::OnVScroll(nSBCode, nPos, pScrollBar);
}