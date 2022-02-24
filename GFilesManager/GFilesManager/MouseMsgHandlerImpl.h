// MouseMsgHandlerImpl.h : header file

#pragma once

#include "GFilesManagerDlg.h"

#include "TrayIcon.h"
#include "TrayIconMouseMsgHandler.h"
#include "Utilities.h"

// Left mouse button double click message handler
class CLeftMouseDblClickMsgHandler: public CTrayIconMouseMsgHandler
{
public:

	// Thay sự kiện kích đúp (WM_LBUTTONDBLCLK) = Sự kiện sau khi kích nhả chuột (WM_LBUTTONDOWN)
	//CLeftMouseDblClickMsgHandler() : CTrayIconMouseMsgHandler(WM_LBUTTONDBLCLK){}
	CLeftMouseDblClickMsgHandler() : CTrayIconMouseMsgHandler(WM_LBUTTONUP){}

	void MouseMsgHandler()
	{
		POINT ptMouse;
		GetCursorPos(&ptMouse);
		jXIcon = (int)ptMouse.x;
		jYIcon = (int)ptMouse.y;

		if (!AfxGetMainWnd()->IsWindowVisible())
		{
			ShowWnd(AfxGetMainWnd());
			AfxGetMainWnd()->ActivateTopParent();
		}
		else
		{
			HideWnd(AfxGetMainWnd());
		}
	}
};

// Right mouse button click message handler
class CRightMouseClickMsgHandler: public CTrayIconMouseMsgHandler
{
public:

	CRightMouseClickMsgHandler() : CTrayIconMouseMsgHandler(WM_RBUTTONDOWN){}
	void MouseMsgHandler()
	{
		ShowPopupMenu(AfxGetMainWnd(), IDR_TRAYICON_MENU, 0);
	}
};

// Mouse hover message handler
class CMouseHoverMsgHandler: public CTrayIconMouseMsgHandler
{
public:

	CMouseHoverMsgHandler(CTrayIcon* pTrayIcon) : CTrayIconMouseMsgHandler(WM_MOUSEFIRST)
	{
		m_pTrayIcon = pTrayIcon;
	}

	void SetTrayIcon(CTrayIcon* pTrayIcon)
	{
		m_pTrayIcon = pTrayIcon;
	}

	void MouseMsgHandler()
	{
		m_pTrayIcon->HoverIcon();
	}

private:

	CTrayIcon* m_pTrayIcon;
};