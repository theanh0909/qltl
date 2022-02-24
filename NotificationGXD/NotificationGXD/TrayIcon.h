// CTrayIcon.h : header file

#pragma once

#include "TrayIconMouseMsgHandler.h"
#include "IconData.h"

const unsigned int SECOND = 1000;

const unsigned int HOVER_EVENT = 1;

const unsigned int ANIMATE_EVENT = 2;

const unsigned int WM_TI_TASKBARCREATED = ::RegisterWindowMessage(_T("TaskBarCreated"));

// CTrayIcon

class CTrayIcon : public CWnd
{
	DECLARE_DYNAMIC(CTrayIcon)

public:

	// Use if notification messages are to be sent to the parent window.
	CTrayIcon(	CWnd *pWnd,
				unsigned int uNotificationMsg,
				MouseMsgHandlerPtr *pMouseMsgHandler, 
				unsigned int uHandlersCount, 
				IconDataPtr *pIconData,
				unsigned int uIconsCount,
				int nSelectedIconIndex,
				unsigned int uElapse );

	// Use if notification messages are to be sent to the CTrayIcon class.
	CTrayIcon(	MouseMsgHandlerPtr *pMouseMsgHandler, 
				unsigned int uHandlersCount, 
				IconDataPtr *pIconData,
				unsigned int uIconsCount,
				int nSelectedIconIndex,
				unsigned int uElapse );

	// Sets the notification message target window
	void SetTargetWnd(CWnd *pWnd);

	// Sets the notification message identifier
	void SetNotificationMsg(unsigned int uNotificationMsg);

	// Sets the selected icon index
	void SetSelectedIconIndex(int nIconIndex);

	// Gets the selected icon index
	int GetSelectedIconIndex() const;

	// Sets the mouse message handler
	void SetMouseMsgHandler(MouseMsgHandlerPtr *pMouseMsgHandler, unsigned int uHandlersCount);

	// Gets the mouse message handler
	MouseMsgHandlerPtr* GetMouseMsgHandler() const;

	// Sets timer elapse
	void SetTimerElapse(unsigned int uElapse);

	// Gets timer elapse
	int GetTimerElapse() const;

	// Sets icon data
	void SetIconData(IconDataPtr *pIconData, unsigned int uIconsCount);

	// Gets icons data
	IconDataPtr* GetIconData() const;

	// Indicates whether animation is in progress
	BOOL IsAnimating() const;

public:

	// Destructor
	~CTrayIcon();

public:

	// Dispalys a hidden icon
	BOOL ShowIcon();

	// Hides a visible icon
	BOOL HideIcon();

	// Refreshes icon
	BOOL RefreshIcon();

	// Shows the hover icon
	BOOL HoverIcon();

	// Starts icon animation
	BOOL StartIconAnimation();

	// Stops icon animation
	BOOL StopIconAnimation();

protected:

	DECLARE_MESSAGE_MAP()

	// Adds an icon but doesn't display it
	BOOL AddIcon();

	// Deletes an added icon fro the taskbar status area
	BOOL DeleteIcon();

	// Sets or unsets a hover icon
	BOOL SetHoverIcon(bool bEnable);

public:

	// Callback handler for the taskbar created event
	LRESULT OnTaskBarCreated(WPARAM wParam, LPARAM lParam);

	// Default callback handler for taskbar notification message event
	LRESULT OnNotifyIcon(WPARAM uID, LPARAM lEvent);

	// Callback handler for the window procedure
	LRESULT WindowProc(unsigned int uMsg, WPARAM wParam, LPARAM lParam);

	// Callback handler for the timer elapsed event
	void OnTimer(unsigned int uIDEvent);	

private:

	// Default constructor
	CTrayIcon(){}

	// Initializes the CTrayIcon object
	void Initialize(	CWnd *pWnd,
						unsigned int uNotificationMsg,
						MouseMsgHandlerPtr *pMouseMsgHandler, 
						unsigned int uHandlersCount, 
						IconDataPtr *pIconData,
						unsigned int uIconsCount,
						int nSelectedIconIndex,
						unsigned int uElapse );

private:

	NOTIFYICONDATA m_tNotifyIconData; 

	CWnd *m_pWnd;

	unsigned int m_uNotificationMsg;

	IconDataPtr *m_pIconData;

	unsigned int m_uIconsCount;

	int m_nSelectedIconIndex;

	BOOL m_bHiddenIcon;

	MouseMsgHandlerPtr *m_pMouseMsgHandler;

	unsigned int m_uHandlersCount;

	unsigned int m_uTimer;

	unsigned int m_uElapse;

	unsigned int m_uHoverTimer;

	BOOL m_bAnimate;
};