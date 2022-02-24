#pragma once
#include "StdAfx.h"

// FrmTabDelete dialog
class FrmTabDelete : public CDialogEx
{
	DECLARE_DYNAMIC(FrmTabDelete)

public:
	FrmTabDelete(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTabDelete();

	enum { IDD = FVIEW_RENAME_DELETE };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CColorEdit m_txt_chuoi;
	CButton m_rad_trim, m_rad_space, m_rad_chuoi;
	CButton m_rad_all, m_rad_begin, m_rad_end, m_rad_left, m_rad_right;

public:
	void _SetDefault();
	void _SetUnCheckRadio();
	void _SetUnCheckChildRadio();
	void _SetEnableRadio(bool bl);

	afx_msg void OnBnClickedRadTrim();
	afx_msg void OnBnClickedRadSpace();
	afx_msg void OnBnClickedRadChuoi();

};
