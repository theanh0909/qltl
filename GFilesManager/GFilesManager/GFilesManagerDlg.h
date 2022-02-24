
// GFilesManagerDlg.h : header file
//

#pragma once
#include "TrayIcon.h"

#include "MyTabCtrl.h"
#include "FTab1.h"
#include "FTab2.h"
#include "FTab3.h"
#include "CDownloadFileSever.h"

// CGFilesManagerDlg dialog
class CGFilesManagerDlg : public CDialogEx
{
// Construction
public:
	CGFilesManagerDlg(CWnd* pParent = nullptr);	// standard constructor
	virtual ~CGFilesManagerDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GFILESMANAGER_DIALOG };
#endif

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnTimer(UINT_PTR nIDEvent);

public:
	CMFCButton m_btn_update;
	CXPStyleButtonST m_btn_icon, m_btn_close, m_btn_options;
	CThemeHelperST m_ThemeHelper;

	CTrayIcon* m_pTrayIcon;
	CMyTabCtrl tab_main;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;
	CDownloadFileSever *sv;

	int demTimer;
	int nVersion;
	CString strNewVersion;	
	CString szCheckRestartFiles;

public:
	afx_msg void _SetDefault();
	afx_msg void _CreateTrayIconTask();
	afx_msg void _SetLocationAndSize();
	afx_msg void _HiddenTrayIconTaskbar();
	afx_msg void _ShowIconTask(bool flag = false);

	afx_msg void OnTrayiconShow();
	afx_msg void OnTrayiconClose();

	afx_msg void OnBnClickedBtnClose();
	afx_msg void OnBnClickedBtnOptions();
	afx_msg void OnBnClickedBtnUpdate();
};
