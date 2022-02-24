
// NotificationGXDDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "NotificationGXD.h"
#include "NotificationGXDDlg.h"
#include "afxdialogex.h"

#include "MouseMsgHandlerImpl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define NotificationApp		L"Notification"

// CNotificationGXDDlg dialog

CNotificationGXDDlg::CNotificationGXDDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_NOTIFICATIONGXD_DIALOG, pParent)
{
	demTimer = 0;
}

void CNotificationGXDDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CNotificationGXDDlg, CDialogEx)
	ON_WM_TIMER()
	ON_WM_SYSCOMMAND()
	ON_COMMAND(ID_TRAYICON_SHOW, &CNotificationGXDDlg::OnTrayiconShow)
	ON_COMMAND(ID_TRAYICON_EXIT, &CNotificationGXDDlg::OnTrayiconExit)
	ON_MESSAGE(WM_TASKBARNOTIFIERCLICKED, OnTaskbarNotifierClicked)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()


// CNotificationGXDDlg message handlers

BOOL CNotificationGXDDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	SetLayeredWindowAttributes(RGB(255, 0, 255), 0, LWA_COLORKEY);

	_CreateTrayIcon();

	demTimer = 0;
	SetTimer(1, 1000, NULL);		// 1000ms = 1 second

	_CreateToolBarNotifier(this);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

BOOL CNotificationGXDDlg::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_ESCAPE)
		{
			OnTrayiconTaskbar();
			return TRUE;
		}
	}

	return FALSE;
}


void CNotificationGXDDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnTrayiconTaskbar();
	else CDialogEx::OnSysCommand(nID, lParam);
}

void CNotificationGXDDlg::OnTimer(UINT_PTR nIDEvent)
{
	if (demTimer == 0)
	{
		_ShowToolBarNotifier();
		HideWnd(this);
	}
	else if (demTimer == 20)
	{
		CDialog::OnCancel();
	}
	demTimer++;

	CDialogEx::OnTimer(nIDEvent);
}

void CNotificationGXDDlg::_ShowToolBarNotifier()
{
	CString szContent = L"Đã có bản cập nhật mới "
		L"\nClick để cập nhật phần mềm hoặc trên Ribbon 'Quản lý tài liệu' "
		L"chạy lệnh 'Cập nhật phần mềm' ";
	m_taskbar_notifier.Show(szContent, 1000, 8 * 1000, 2000, 1);
}

void CNotificationGXDDlg::_CreateToolBarNotifier(CWnd* pParent)
{
	m_taskbar_notifier.Create(pParent);
	m_taskbar_notifier.SetSkin(IDB_BITMAP_NOTIFIER);
	m_taskbar_notifier.SetTextFont(L"Arial", 120, TN_TEXT_NORMAL, TN_TEXT_UNDERLINE);
	m_taskbar_notifier.SetTextColor(RGB(0, 112, 192), RGB(192, 0, 0));
	m_taskbar_notifier.SetTextRect(CRect(10, 55, m_taskbar_notifier.m_nSkinWidth - 10, m_taskbar_notifier.m_nSkinHeight - 15));
}

LRESULT CNotificationGXDDlg::OnTaskbarNotifierClicked(WPARAM wParam, LPARAM lParam)
{
	MessageBox(L"OK");
	return 0;
}

void CNotificationGXDDlg::_CreateTrayIcon()
{
	try
	{
		MouseMsgHandlerPtr *pHandler = new MouseMsgHandlerPtr[3];
		MouseMsgHandlerPtr rbcHandler = new CRightMouseClickMsgHandler();
		MouseMsgHandlerPtr lbcHandler = new CLeftMouseDblClickMsgHandler();
		MouseMsgHandlerPtr mhHandler = new CMouseHoverMsgHandler(m_pTrayIcon);

		pHandler[0] = lbcHandler;
		pHandler[1] = rbcHandler;
		pHandler[2] = mhHandler;

		IconDataPtr *pIconData = new IconDataPtr[2];
		pIconData[0] = new CIconData(IDR_MAINFRAME, (LPSTR)NotificationApp, IDR_MAINFRAME);
		pIconData[1] = new CIconData(IDR_MAINFRAME, (LPSTR)NotificationApp, IDR_MAINFRAME);
		m_pTrayIcon = new CTrayIcon(pHandler, 3, pIconData, 2, 0, SECOND);

		((CMouseHoverMsgHandler*)mhHandler)->SetTrayIcon(m_pTrayIcon);
	}
	catch (...) {}
}


void CNotificationGXDDlg::OnTrayiconTaskbar()
{
	m_pTrayIcon->ShowIcon();
	HideWnd(this);
}


void CNotificationGXDDlg::OnTrayiconShow()
{
	ShowWnd(this);
}


void CNotificationGXDDlg::OnTrayiconExit()
{
	CDialogEx::OnCancel();
}


BOOL CNotificationGXDDlg::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;

	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(255, 0, 255));

	return FALSE;
}
