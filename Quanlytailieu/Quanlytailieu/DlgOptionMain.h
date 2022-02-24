#pragma once
#include "StdAfx.h"

class DlgOptionMain : public CDialogEx
{
	DECLARE_DYNAMIC(DlgOptionMain)

public:
	DlgOptionMain(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgOptionMain();

	// Dialog Data
	enum { IDD = DLG_OPTION_MAIN };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CColorEdit m_txt_browse;
	CComboBox m_cbb_autoupdate;
	CButton m_check_msgupdate, m_check_msgerror, m_check_delfilecsv, m_check_startwindows;
	
	CXPStyleButtonST m_btn_ok, m_btn_cancel, m_btn_register;
	CThemeHelperST m_ThemeHelper;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	_bstr_t bsConfig;
	_WorksheetPtr psConfig;
	RangePtr prConfig;

	CString pathFolder;

public:
	void _XacdinhSheetLienquan();
	void _SetDefault();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnBrowse();
	afx_msg void OnBnClickedBtnRegisterCom();
};
