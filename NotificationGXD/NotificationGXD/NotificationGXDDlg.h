#pragma once
#include "TrayIcon.h"
#include "TaskbarNotifier.h"

// CNotificationGXDDlg dialog
class CNotificationGXDDlg : public CDialogEx
{
// Construction
public:
	CNotificationGXDDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_NOTIFICATIONGXD_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

protected:
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	DECLARE_MESSAGE_MAP()

private:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg LRESULT OnTaskbarNotifierClicked(WPARAM wParam, LPARAM lParam);

public:
	CTrayIcon* m_pTrayIcon;
	CTaskbarNotifier m_taskbar_notifier;

	int demTimer;

public:

	afx_msg void _CreateTrayIcon();
	afx_msg void _ShowToolBarNotifier();
	afx_msg void _CreateToolBarNotifier(CWnd * pParent);

	afx_msg void OnTrayiconTaskbar();
	afx_msg void OnTrayiconShow();
	afx_msg void OnTrayiconExit();

	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
};
