#pragma once
#include "pch.h"

// DlgUpdateInfo dialog

class DlgUpdateInfo : public CDialogEx
{
	DECLARE_DYNAMIC(DlgUpdateInfo)

public:
	DlgUpdateInfo(CWnd* pParent = nullptr);   // standard constructor
	virtual ~DlgUpdateInfo();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = DLG_UPDATEINFO };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CXPStyleButtonST m_btn_logo, m_btn_update;
	CThemeHelperST m_ThemeHelper;
	
	CColorEdit m_lbl_update;	
	CColorEdit m_lblchk_name, m_lblchk_email, m_lblchk_mobile;
	CColorEdit m_txt_name, m_txt_email, m_txt_mobile, m_txt_address;

	CStatic m_lbl_name, m_lbl_email, m_lbl_mobile, m_lbl_address;

	CFunction* ff;

public:
	void SetControlDefault();
	void TextAuto();

	bool CheckValueName();
	bool CheckValueEmail();
	bool CheckValueMobile();

	afx_msg void OnBnClickedBtnUpdate();	
	afx_msg void OnEnKillfocusTxtName();
	afx_msg void OnEnKillfocusTxtEmail();
	afx_msg void OnEnKillfocusTxtMobile();
	afx_msg void OnEnKillfocusTxtAddress();

	afx_msg void OnEnSetfocusTxtName();
	afx_msg void OnEnSetfocusTxtEmail();
	afx_msg void OnEnSetfocusTxtMobile();
	afx_msg void OnEnSetfocusTxtAddress();

	afx_msg void OnNMClickHyperlink(NMHDR *pNMHDR, LRESULT *pResult);
};
