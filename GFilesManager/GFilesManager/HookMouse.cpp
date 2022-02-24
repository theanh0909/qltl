#include "pch.h"
#include "HookMouse.h"

int MyHook::Messsages()
{
	while (msg.message != WM_QUIT)
	{
		if (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		Sleep(1);
	}
	UninstallHook();
	return (int)msg.wParam;
}

void MyHook::InstallHook()
{
	if(!(hook = SetWindowsHookEx(WH_MOUSE_LL, MyMouseCallback, NULL, 0))) printf_s("Error: %x \n", GetLastError());
}

void MyHook::UninstallHook()
{
	UnhookWindowsHookEx(hook);
}

LRESULT WINAPI MyMouseCallback(int nCode, WPARAM wParam, LPARAM lParam)
{
	MSLLHOOKSTRUCT * pMouseStruct = (MSLLHOOKSTRUCT *)lParam;	
	
	if(nCode == 0)
	{
		switch (wParam)
		{
			case WM_LBUTTONDOWN:
			{
				POINT ptMouse;
				GetCursorPos(&ptMouse);
				int X = (int)ptMouse.x;
				int Y = (int)ptMouse.y;

				if (!(X >= dPosX && Y >= dPosY && X <= dPosX + dWidthF && Y <= dPosY + dHeightF)
					&& (jXIcon == -1 || !(X >= jXIcon && Y >= jYIcon && X <= jXIcon + 16 && Y <= jYIcon + 16)))
				{
					AfxGetMainWnd()->ShowWindow(SW_HIDE);
				}
			}
			break;
		}
	}

	return CallNextHookEx(MyHook::Instance().hook, nCode, wParam, lParam);
}