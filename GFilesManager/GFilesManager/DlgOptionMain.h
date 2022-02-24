#pragma once
#include "pch.h"

class DlgOptionMain : public CDialogEx
{
	DECLARE_DYNAMIC(DlgOptionMain)

public:
	DlgOptionMain(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgOptionMain();

	// Dialog Data
	enum { IDD = IDD_DIALOG_OPTION };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CButton m_check_start;
	CXPStyleButtonST m_btn_ok, m_btn_cancel;
	CThemeHelperST m_ThemeHelper;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	void _SetDefault();
	void _SetLocationAndSize();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
};
