#pragma once
#include "StdAfx.h"

// FrmTabAdd dialog
class FrmTabAdd : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTabAdd)

public:
	FrmTabAdd(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTabAdd();

	enum { IDD = FVIEW_RENAME_ADD };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CEdit m_txt_num;
	CColorEdit m_txt_chuoi;
	CColorEdit m_lbl_vdnum;
	CButton m_rad_begin, m_rad_end, m_rad_mid;
	CNumSpinCtrl m_spin_num;

public:
	void _SetDefault();
	void _SetUnCheckRadio();
	void _SetEnableNumber(bool bl = false);

	afx_msg void OnBnClickedRadBegin();
	afx_msg void OnBnClickedRadEnd();
	afx_msg void OnBnClickedRadMid();

	afx_msg void OnEnChangeTxtNum();


	afx_msg void OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult);
};
