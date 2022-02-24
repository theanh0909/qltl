// TrayIcon.cpp : implementation file
//

#include "pch.h"
#include "TrayIcon.h"


// CTrayIcon

IMPLEMENT_DYNAMIC(CTrayIcon, CWnd)
CTrayIcon::CTrayIcon(	MouseMsgHandlerPtr *pMouseMsgHandler, 
						unsigned int uHandlersCount, 
						IconDataPtr *pIconData,
						unsigned int uIconsCount,
						int nSelectedIconIndex = -1,
						unsigned int uElapse = SECOND )
{
	CTrayIcon::Initialize(NULL, ::RegisterWindowMessage(_T("NotifyIcon")), pMouseMsgHandler, uHandlersCount, pIconData, uIconsCount, nSelectedIconIndex, uElapse);
}

CTrayIcon::CTrayIcon(	CWnd* pWnd,
						unsigned int uNotificationMsg,
						MouseMsgHandlerPtr *pMouseMsgHandler, 
						unsigned int uHandlersCount, 
						IconDataPtr *pIconData,
						unsigned int uIconsCount,
						int nSelectedIconIndex = -1,
						unsigned int uElapse = SECOND )
{
	CTrayIcon::Initialize(pWnd, uNotificationMsg, pMouseMsgHandler, uHandlersCount, pIconData, uIconsCount, nSelectedIconIndex, uElapse);
}

void CTrayIcon::Initialize(	CWnd* pWnd,
							unsigned int uNotificationMsg,
							MouseMsgHandlerPtr *pMouseMsgHandler, 
							unsigned int uHandlersCount, 
							IconDataPtr *pIconData,
							unsigned int uIconsCount,
							int nSelectedIconIndex,
							unsigned int uElapse )
{
	// Create an invisible window for the Tray Icon
	CTrayIcon::CreateEx(0, AfxRegisterWndClass(0), _T(""), WS_POPUP, 0,0,0,0, NULL, 0);

	m_tNotifyIconData.cbSize = sizeof(NOTIFYICONDATA); 
	m_bHiddenIcon = TRUE;

	SetTargetWnd(pWnd);
	SetNotificationMsg(uNotificationMsg);
	SetIconData(pIconData, uIconsCount);
	SetSelectedIconIndex(nSelectedIconIndex);	
	SetMouseMsgHandler(pMouseMsgHandler, uHandlersCount);	

	m_uHoverTimer = m_uTimer = 0;
	SetTimerElapse(uElapse);
	m_bAnimate = false;
}

CTrayIcon::~CTrayIcon()
{
	if(m_uTimer != 0)
		KillTimer(m_uTimer);

	if(m_uHoverTimer != 0)
		KillTimer(m_uHoverTimer);

	DeleteIcon();

	for( unsigned int i = 0; i < m_uIconsCount; i++)
		delete [] m_pIconData[i];

	delete [] m_pIconData;

	for( unsigned int i = 0; i < m_uHandlersCount; i++ )	    
		delete [] m_pMouseMsgHandler[i];

	delete [] m_pMouseMsgHandler;
}

void CTrayIcon::SetTargetWnd(CWnd* pWnd)
{
	if(pWnd != NULL)
		m_pWnd = pWnd;
	else
		m_pWnd = (CWnd*)this;

	m_tNotifyIconData.hWnd = m_pWnd->GetSafeHwnd(); 
}

void CTrayIcon::SetNotificationMsg(unsigned int uNotificationMsg)
{
	// Make sure we avoid conflict with other messages
	ASSERT(uNotificationMsg >= WM_APP);

	m_uNotificationMsg = uNotificationMsg;

	m_tNotifyIconData.uFlags |= NIF_MESSAGE;
	m_tNotifyIconData.uCallbackMessage = m_uNotificationMsg;
}

void CTrayIcon::SetSelectedIconIndex(int nIconIndex)
{
	m_nSelectedIconIndex = nIconIndex;
	if(m_pIconData != NULL)
	{
		if(m_nSelectedIconIndex > -1 && m_nSelectedIconIndex < (int)m_uIconsCount)
		{
			if(m_pIconData[m_nSelectedIconIndex]->GetIconHandler() != NULL)
			{
				m_tNotifyIconData.uFlags |= NIF_ICON;
				m_tNotifyIconData.hIcon = m_pIconData[m_nSelectedIconIndex]->GetIconHandler();
				m_tNotifyIconData.uID = m_pIconData[m_nSelectedIconIndex]->GetIconID();
			}

			if(m_pIconData[m_nSelectedIconIndex]->GetToolTip() != NULL)
			{
				m_tNotifyIconData.uFlags |= NIF_TIP;	
				if (m_pIconData[m_nSelectedIconIndex]->GetToolTip()) 
					lstrcpyn((LPWSTR)m_tNotifyIconData.szTip,(LPCWSTR)m_pIconData[m_nSelectedIconIndex]->GetToolTip(), sizeof(m_tNotifyIconData.szTip)); 
				else 
					m_tNotifyIconData.szTip[0] = '\0';
			}
		}
	}
}

int CTrayIcon::GetSelectedIconIndex() const
{
	return m_nSelectedIconIndex;
}

void CTrayIcon::SetTimerElapse(unsigned int uElapse)
{
	m_uElapse = uElapse;
}

int CTrayIcon::GetTimerElapse() const
{
	return m_uElapse;
}

void CTrayIcon::SetMouseMsgHandler(MouseMsgHandlerPtr *pMouseMsgHandler, unsigned int uHandlersCount)
{
	m_pMouseMsgHandler = pMouseMsgHandler;
	m_uHandlersCount = uHandlersCount;
}

MouseMsgHandlerPtr* CTrayIcon::GetMouseMsgHandler() const
{
	return m_pMouseMsgHandler;
}

void CTrayIcon::SetIconData(IconDataPtr *pIconData, unsigned int uIconsCount)
{
	m_pIconData = pIconData;
	m_uIconsCount = uIconsCount;
}

IconDataPtr* CTrayIcon::GetIconData() const
{
	return m_pIconData;
}

BOOL CTrayIcon::IsAnimating() const
{
	return m_bAnimate;
}

BEGIN_MESSAGE_MAP(CTrayIcon, CWnd)
	ON_REGISTERED_MESSAGE(WM_TI_TASKBARCREATED, OnTaskBarCreated)
END_MESSAGE_MAP()

// CTrayIcon message handlers

LRESULT CTrayIcon::OnTaskBarCreated(WPARAM wParam, LPARAM lParam)
{
	VERIFY(AddIcon());
	return 0;
}

LRESULT CTrayIcon::OnNotifyIcon(WPARAM wParam, LPARAM lParam)
{
	unsigned int uID = (unsigned int) wParam; 
	unsigned int uMouseMsg = (unsigned int) lParam; 
 
	for(unsigned int i=0; i<m_uHandlersCount; i++)
		if ( (uMouseMsg == m_pMouseMsgHandler[i]->GetMouseMsgID()) && (uID == m_pIconData[m_nSelectedIconIndex]->GetIconID()) )
			m_pMouseMsgHandler[i]->MouseMsgHandler();

	return 0;
}

LRESULT CTrayIcon::WindowProc(unsigned int msg, WPARAM wParam, LPARAM lParam) 
{
	if (msg == m_uNotificationMsg)
		return OnNotifyIcon(wParam, lParam);
	else if(msg == WM_TIMER)
		OnTimer(wParam);

	return CWnd::WindowProc(msg, wParam, lParam);
}

BOOL CTrayIcon::AddIcon()
{
	m_tNotifyIconData.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
 
	BOOL res = Shell_NotifyIcon(NIM_ADD, &m_tNotifyIconData); 
 
	if (m_pIconData[m_nSelectedIconIndex]->GetIconHandler()) 
		DestroyIcon(m_pIconData[m_nSelectedIconIndex]->GetIconHandler()); 
 
	m_bHiddenIcon = FALSE;

	return res; 
}

BOOL CTrayIcon::DeleteIcon()
{
	if(m_bAnimate)
		StopIconAnimation();

	return Shell_NotifyIcon(NIM_DELETE, &m_tNotifyIconData);
}

BOOL CTrayIcon::HideIcon()
{
	if (m_bHiddenIcon)
		return TRUE;

	m_tNotifyIconData.uFlags = NIF_STATE;
	m_tNotifyIconData.dwState = NIS_HIDDEN;
	m_tNotifyIconData.dwStateMask = NIS_HIDDEN;
	BOOL res = m_bHiddenIcon = Shell_NotifyIcon( NIM_MODIFY, &m_tNotifyIconData);

	DeleteIcon();

	return res;
}

BOOL CTrayIcon::ShowIcon()
{
	if (!m_bHiddenIcon)
		return TRUE;

	m_tNotifyIconData.uFlags = NIF_STATE;
	m_tNotifyIconData.dwState = 0;
	m_tNotifyIconData.dwStateMask = NIS_HIDDEN;
	BOOL res = Shell_NotifyIcon ( NIM_MODIFY, &m_tNotifyIconData );
	m_bHiddenIcon = FALSE;

	AddIcon();

	return res;
}

BOOL CTrayIcon::RefreshIcon()
{
	if (m_bHiddenIcon)
		return TRUE;

	DeleteIcon();

	SetSelectedIconIndex(m_nSelectedIconIndex);

	m_tNotifyIconData.uFlags = NIF_STATE;
	m_tNotifyIconData.dwState = 0;
	m_tNotifyIconData.dwStateMask = NIS_HIDDEN;
	BOOL res = Shell_NotifyIcon ( NIM_MODIFY, &m_tNotifyIconData );
	m_bHiddenIcon = FALSE;

	AddIcon();

	return res;
}

BOOL CTrayIcon::SetHoverIcon(bool bEnable)
{
	IconDataPtr pIconData = m_pIconData[m_nSelectedIconIndex];

	if(bEnable && pIconData->GetIconID() != pIconData->GetHoverIconID())
	{
		pIconData->LoadIcon(pIconData->GetHoverIconID());		
		RefreshIcon();
		return true;
	}
	else if(!bEnable && pIconData->GetIconID() != pIconData->GetDefaultIconID())
	{
		pIconData->LoadIcon(pIconData->GetDefaultIconID());		
		RefreshIcon();
		return true;
	}

	return false;
}

BOOL CTrayIcon::HoverIcon()
{
	if( m_uHoverTimer == 0 )
	{
		SetHoverIcon(true);

		m_uHoverTimer = SetTimer(HOVER_EVENT, SECOND, 0);
	}

	return false;
}

BOOL CTrayIcon::StartIconAnimation()
{
	if(m_uTimer == 0 && m_uElapse > 0)
	{
		m_uTimer = SetTimer(ANIMATE_EVENT, m_uElapse, 0);
		if(m_uTimer > 0)
		{
			m_bAnimate = true;
			return true;
		}
	}
	
	return false;
}

BOOL CTrayIcon::StopIconAnimation()
{
	if(KillTimer(m_uTimer))
	{
		m_uTimer = 0;
		m_bAnimate = false;
		return true;
	}

	return false;
}

void CTrayIcon::OnTimer(unsigned int uIDEvent)
{
	if(uIDEvent == ANIMATE_EVENT)
	{
		int nIconIndex = m_nSelectedIconIndex + 1;
		if(nIconIndex >= (int)m_uIconsCount)
			nIconIndex = 0;
		
		HideIcon();
		SetSelectedIconIndex(nIconIndex);
		ShowIcon();
		StartIconAnimation();
	}
	else if(uIDEvent == HOVER_EVENT)
	{
		SetHoverIcon(false);
		KillTimer(m_uHoverTimer);
		m_uHoverTimer = 0;
	}

	// Call base class handler.   
	CWnd::OnTimer(uIDEvent);
}