#pragma once


// CListEdit

class CListEdit : public CEdit
{
	DECLARE_DYNAMIC(CListEdit)

public:
	CListEdit();
	virtual ~CListEdit();

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnKillFocus(CWnd* pNewWnd);
protected:
	virtual void PreSubclassWindow();
};
