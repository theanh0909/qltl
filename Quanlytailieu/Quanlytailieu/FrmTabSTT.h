#pragma once
#include "StdAfx.h"

// FrmTabSTT dialog
class FrmTabSTT : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTabSTT)

public:
	FrmTabSTT(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTabSTT();

	enum { IDD = FVIEW_RENAME_STT };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CEdit m_txt_num;
	CColorEdit m_txt_tiento, m_txt_hauto;
	CColorEdit m_lbl_vdtiento, m_lbl_vdhauto, m_lbl_vitri;
	CButton m_check_zero, m_check_end, m_check_replace;
	CNumSpinCtrl m_spin_num;

public:
	void _SetDefault();

	afx_msg void OnEnChangeTxtNum();
	afx_msg void OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult);
};
