// Utilities.h : header file

#pragma once

// Utilities functions

// Shows a window
BOOL ShowWnd(CWnd* pWnd)
{
	return pWnd->ShowWindow(SW_SHOW);
}

// Hides a window
BOOL HideWnd(CWnd* pWnd)
{
	return pWnd->ShowWindow(SW_HIDE);
}

// Shows a popup menu
void ShowPopupMenu(CWnd* pWnd, unsigned int uResourseID, int nSubMenuPos)
{
	try
	{
		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(uResourseID));
		HMENU pMenu = GetSubMenu(hMMenu, nSubMenuPos);
		if (pMenu)
		{
			MenuIcon mnIcon;
			mnIcon.GdiplusStartupInputConfig();

			HBITMAP hBmp;

			if (!AfxGetMainWnd()->IsWindowVisible())
			{
				ModifyMenu(pMenu, ID_TRAYICON_SHOW, MF_BYCOMMAND | MF_STRING, ID_TRAYICON_SHOW, L"Hiển thị");
				mnIcon.AddIconMenuItem(pMenu, hBmp, ID_TRAYICON_SHOW, IDI_ICON_UP);
			}
			else
			{
				ModifyMenu(pMenu, ID_TRAYICON_SHOW, MF_BYCOMMAND | MF_STRING, ID_TRAYICON_SHOW, L"Ẩn xuống");
				mnIcon.AddIconMenuItem(pMenu, hBmp, ID_TRAYICON_SHOW, IDI_ICON_DOWN);
			}

			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_TRAYICON_CLOSE, IDI_ICON_CLOSE);

			CPoint aPos;
			GetCursorPos(&aPos);
			TrackPopupMenu(pMenu, TPM_RIGHTALIGN | TPM_RIGHTBUTTON, aPos.x + 10, aPos.y - 15, 0, pWnd->GetSafeHwnd(), NULL);

			if (hBmp) ::DeleteObject(hBmp);
		}

		if (hMMenu) VERIFY(DestroyMenu(hMMenu));
	}
	catch (...) {}
}

int IsTaskbarWndVisible(int &iTaskHeight)
{
	// nPostion = -1 --> Thanh Taskbar bị ẩn
	// nPostion =  0 - Bottom | 1 - Left | 2 - Top | 3 - Right
	int nPostion = -1;
	HWND hTaskbarWnd = ::FindWindow(L"Shell_TrayWnd", NULL);
	HMONITOR hMonitor = MonitorFromWindow(hTaskbarWnd, MONITOR_DEFAULTTONEAREST);
	MONITORINFO info = { sizeof(MONITORINFO) };
	if (GetMonitorInfo(hMonitor, &info))
	{
		RECT rect;
		GetWindowRect(hTaskbarWnd, &rect);

		if (!((rect.top >= info.rcMonitor.bottom - 4) ||
			(rect.right <= 2) ||
			(rect.bottom <= 4) ||
			(rect.left >= info.rcMonitor.right - 2)))
		{
			if ((int)rect.left > 0)
			{
				// Task --> Pane Right
				nPostion = 3;
				iTaskHeight = (int)rect.left;
			}
			else
			{
				if ((int)rect.top > 0)
				{
					// Task --> Pane Bottom
					nPostion = 0;
					iTaskHeight = (int)rect.top;
				}
				else
				{
					if ((int)rect.right == 1366)
					{
						// Task --> Pane Top
						nPostion = 2;
						iTaskHeight = (int)rect.bottom;
					}
					else
					{
						// Task --> Pane Left
						nPostion = 1;
						iTaskHeight = (int)rect.right;
					}
				}
			}
		}
	}
	return nPostion;
}