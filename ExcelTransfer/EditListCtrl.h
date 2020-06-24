#pragma once
#include "ListEdit.h"
#include "ListComboBox.h"
#include "_PersistAttriMapXML.h"

#define  IDC_CELL_EDIT		0xffe0
#define  IDC_CELL_COMBOBOX	0xffe1

// CEditListCtrl

class CEditListCtrl : public CListCtrl
{
	DECLARE_DYNAMIC(CEditListCtrl)

public:
	CEditListCtrl();
	virtual ~CEditListCtrl();

protected:
	DECLARE_MESSAGE_MAP()

private:
	CListEdit m_edit;
	CListComboBox m_comboBox;
	int m_nRow; //��
	int m_nCol; //��

public:

	// ���༭��ɺ�ͬ������
	void DisposeEdit(void);

	//����ʧЧ��Ϣ
	void SendInvalidateMsg();

	// ���õ�ǰ���ڵķ��
	void SetStyle(void);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	afx_msg void OnCbnSelchangeCbCellComboBox();

	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnVScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
};