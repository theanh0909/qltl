#pragma once
#include "StdAfx.h"

// FrmTabReplace dialog
class FrmTabReplace : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTabReplace)

public:
	FrmTabReplace(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTabReplace();

	enum { IDD = FVIEW_RENAME_REPLACE };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CColorEdit m_txt_rep, m_txt_chuoi;
	CButton m_chk_case, m_chk_kodau;

public:
	void _SetDefault();

};
