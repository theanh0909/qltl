#pragma once
#include "StdAfx.h"

class DlgThreadProgress : public CDialogEx
{
	DECLARE_DYNAMIC(DlgThreadProgress)

public:
	DlgThreadProgress(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgThreadProgress();

	// Dialog Data
	enum { IDD = DLG_THREADPROGRESS };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnTimer(UINT_PTR nIDEvent);

public:
	CXPStyleButtonST m_btn_stop;
	CThemeHelperST m_ThemeHelper;
	CTextProgressCtrl m_progress;	

	CString szTextPro;
	int istart, ifinish;
	int itotal;

	afx_msg void OnBnClickedBtnStop();
};
