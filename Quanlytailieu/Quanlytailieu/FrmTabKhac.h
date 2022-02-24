#pragma once
#include "StdAfx.h"

// FrmTabKhac dialog
class FrmTabKhac : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTabKhac)

public:
	FrmTabKhac(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTabKhac();

	enum { IDD = FVIEW_RENAME_KHAC };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CButton m_rad_upper, m_rad_lower, m_rad_unsigned, m_rad_proper, m_rad_kodau;

public:
	void _SetDefault();
	void _SetUnCheckRadio();

};
