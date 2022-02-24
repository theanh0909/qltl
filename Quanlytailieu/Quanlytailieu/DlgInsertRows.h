#pragma once
#include "StdAfx.h"

class DlgInsertRows : public CDialogEx
{
	DECLARE_DYNAMIC(DlgInsertRows)

public:
	DlgInsertRows(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgInsertRows();

	// Dialog Data
	enum { IDD = DLG_INSERT_ROWS };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CColorEdit m_txt_num;
	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;
	CNumSpinCtrl m_spin_num;

	int numInsert;
	int numMax;

	void _SetDefault();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	
	afx_msg void OnEnChangeTxtNum();
	afx_msg void OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult);	
};
