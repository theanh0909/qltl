#pragma once
#include "StdAfx.h"

class DlgSelectionCells : public CDialogEx
{
	DECLARE_DYNAMIC(DlgSelectionCells)

public:
	DlgSelectionCells(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgSelectionCells();

	// Dialog Data
	enum { IDD = DLG_SELECTION_CELLS };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;
	CButton m_radio_sel, m_radio_group, m_radio_all;

	int iSelectCell;
	bool blEnableSelection;

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();	
};
