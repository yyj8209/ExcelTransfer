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

// ���༭��ɺ�ͬ������
void CEditListCtrl::DisposeEdit(void)
{
	int nIndex = GetSelectionMark();
	//ͬ�����´�������
	CString sLable;
	_PersistItem* pPersistItem = (_PersistItem*)GetItemData(m_nRow);
	if (NULL == pPersistItem) return;

	if (0 == m_nCol) //edit�ؼ�
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

// ���õ�ǰ���ڵķ��
void CEditListCtrl::SetStyle(void)
{
	LONG lStyle;
	lStyle = GetWindowLong(m_hWnd, GWL_STYLE);
	lStyle &= ~LVS_TYPEMASK; //�����ʾ��ʽ
	lStyle |= LVS_REPORT; //listģʽ
	lStyle |= LVS_SINGLESEL; //��ѡ
	SetWindowLong(m_hWnd, GWL_STYLE, lStyle);

	//��չģʽ
	DWORD dwStyle = GetExtendedStyle();
	dwStyle |= LVS_EX_FULLROWSELECT; //ѡ��ĳ��ʹ�����и���
	dwStyle |= LVS_EX_GRIDLINES; //������
	SetExtendedStyle(dwStyle); //������չ���
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
		//Edit��Ԫ��
		if ((0 == m_nCol))
		{
			if (pPersistItem->nFlag & 0x01)
			{
				MB_PROMPT(RES_LOAD_STRING(IDS_ST_MODIFIYPKATTRI_20170502_01));
				return;
			}

			if (NULL == m_edit)
			{
				//����Edit�ؼ�
				// ES_WANTRETURN ʹ���б༭�����ջس������벢���С������ָ���÷��,���س�����ѡ��ȱʡ�����ť,�������ᵼ�¶Ի���Ĺرա�
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
		//ComboBox��Ԫ��
		if ((1 == m_nCol))
		{
			if (NULL == m_comboBox)
			{
				//����Combobox�ؼ�
				//CBS_DROPDOWNLIST : ����ʽ��Ͽ򣬵���������ڲ��ܽ�������
				//WS_CLIPCHILDREN  :  �京����ǣ������ڲ����Ӵ���������л��ơ�Ĭ������¸����ڻ���Ӵ��ڱ����ǽ��л��Ƶģ��������������������WS_CLIPCHILDREN���ԣ����״��ڲ��ڶ��Ӵ��ڱ�������.
				m_comboBox.Create(WS_CHILD | WS_VISIBLE | CBS_SORT | WS_BORDER | CBS_DROPDOWNLIST | WS_VSCROLL | WS_TABSTOP | CBS_AUTOHSCROLL,
					CRect(0, 0, 0, 0), this, IDC_CELL_COMBOBOX);
			}
			m_comboBox.MoveWindow(rect);

			m_comboBox.ShowWindow(SW_SHOW);
			//TODO comboBox ��ʼ��
			LoadPlmAttri();
			//��ǰcell��ֵ
			CString strCellValue = GetItemText(m_nRow, m_nCol);
			int nIndex = m_comboBox.FindStringExact(0, strCellValue);
			if (CB_ERR == nIndex)
				m_comboBox.SetCurSel(0); //TODO ����Ϊ��ǰֵ��Item
			else
				m_comboBox.SetCurSel(nIndex);

			m_comboBox.SetFocus();
			UpdateWindow();
		}
	}
}


void CEditListCtrl::OnCbnSelchangeCbCellComboBox()
{
	//ͬ�����´�������
	_PersistItem* pPersistItem = (_PersistItem*)GetItemData(m_nRow);
	if (NULL == pPersistItem) return;

	if (1 == m_nCol) //combobox�ؼ�
	{
		CString szLable;
		int nindex = m_comboBox.GetCurSel();
		m_comboBox.GetLBText(nindex, szLable);
		ATTRIFIELD* pAttriField = (ATTRIFIELD*)m_comboBox.GetItemData(nindex);
		if (NULL == pAttriField)
			return;

		//�־û��Ĵ���
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

//����ʧЧ��Ϣ
void CEditListCtrl::SendInvalidateMsg()
{
	//Edit��Ԫ��
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