#include "pch.h"
#include "GHookForm.h"
#include "FrmChiphi.h"
#include "FrmQuytrinh.h"
#include "FrmTieuchuan.h"


#pragma data_seg("SHARED_DATA")
HHOOK hGlobalHook = NULL;
#pragma data_seg()

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CGHookFormApp
BEGIN_MESSAGE_MAP(CGHookFormApp, CWinApp)
END_MESSAGE_MAP()

// CGHookFormApp construction
CGHookFormApp::CGHookFormApp()
{

}

// The one and only CGHookFormApp object
CGHookFormApp theApp;

// CGHookFormApp initialization
BOOL CGHookFormApp::InitInstance()
{
	CWinApp::InitInstance();

	m_FontHeaderList.DeleteObject();
	m_FontHeaderList.CreateFont(16, 0, 0, 0, FW_BOLD, false, false, 0,
		ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
		DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, L"Arial");

	return TRUE;
}

/************************************************************************************************/

void CALLBACK MyTimerProc(HWND hwnd, UINT uMsg, UINT timerId, DWORD dwTime)
{
	KillTimer(hWndMain, CONST_TIMER_MAIN);

	xl->PutStatusBar(L"Ready");
	xl->PutEnableAutoComplete(true);

	// 1 trong 3 điều kiện để hiển thị hộp thoại:
	// 1: Giá trị nhập vào phải khác giá trị ban đầu
	// 2: Phím nhập vào không phải phím ENTER
	// 3: Phím nhập vào là phím ENTER nhưng ngay trước đó sử dụng phím F2
	// (Tạm thời không xét đến trường hợp kích đúp chuột để chỉnh sửa)
	szcellAfter = _xlGIOR(pRgActive, iRowActive, iColumnActive, L"");
	if (szcellAfter != szCellBefore
		|| !blKeyEnter || (blKeyEnter && jKeyBeforeF2 == 1))
	{
		switch (nIndexForm)
		{
			case 1:
			{
				// Tra cứu quy trình
				xl->EnableCancelKey = XlEnableCancelKey(FALSE);
				AFX_MANAGE_STATE(AfxGetStaticModuleState());
				FrmQuytrinh dlg;
				dlg.szTukhoa = szcellAfter;
				dlg.DoModal();
			}
			break;

			case 2:
			{
				// Tra cứu tiêu chuẩn
				xl->EnableCancelKey = XlEnableCancelKey(FALSE);
				AFX_MANAGE_STATE(AfxGetStaticModuleState());
				FrmTieuchuan dlg;
				dlg.szTukhoa = szcellAfter;
				dlg.DoModal();
			}
			break;

			case 3:
			{
				// Tra cứu chi phí
				xl->EnableCancelKey = XlEnableCancelKey(FALSE);
				AFX_MANAGE_STATE(AfxGetStaticModuleState());
				FrmChiphi dlg;
				dlg.szTukhoa = szcellAfter;
				dlg.DoModal();
			}
			break;

			default:
				break;
		}
	}
}

__declspec(dllexport) LRESULT CALLBACK MyHookKeyboard(int nCode, WPARAM wParam, LPARAM lParam)
{
	try
	{
		bool blHook = false;
		unsigned short charBuf;
		static string szResult = "";

		CString szClass = GetClassActiveWindow();
		if (szClass == ClassMainExcel)
		{
			if (nCode >= 0 && (HC_ACTION == nCode) && !((DWORD)lParam & 0x80000000))
			{
				DWORD CTRL_key = GetAsyncKeyState(VK_CONTROL);
				if (CTRL_key == 0)
				{
					if (CheckKeyBoard(wParam))
					{
						BYTE KeyState[256];
						GetKeyboardState(KeyState);
						HKL OriginalLayout = GetKeyboardLayout(0);

						if (ToAsciiEx(wParam, UINT(HIWORD(lParam)), KeyState, &charBuf, 0, OriginalLayout) == 1)
						{
							szResult = char(charBuf);
							blHook = true;
						}
					}
				}

				CTRL_key = 0;
			}

			if (blHook)
			{
				try
				{
					CoInitialize(NULL);
					xl.GetActiveObject(ExcelApplication);
					xl->PutVisible(VARIANT_TRUE);

					_xlGetInfoSheetActive();

					POINT point;
					if (GetCursorPos(&point))
					{
						RangePtr pCell = xl->ActiveWindow->RangeFromPoint(point.x, point.y);
						if (pCell != NULL)
						{
							nIndexMouse = 0;
							nIndexForm = KiemtraDieukienHook();
							if (nIndexForm > 0)
							{
								xl->PutEnableAutoComplete(false);

								szKeyboard = (CString)szResult.c_str();
								szCellBefore = _xlGIOR(pRgActive, iRowActive, iColumnActive, L"");

								if (wParam != VK_RETURN)
								{
									CSendKeys *m_sk = new CSendKeys;
									m_sk->SendKeys(L"{DOWN}");
									delete m_sk;
								}

								SetTimer(hWndMain, CONST_TIMER_MAIN, 1, (TIMERPROC)&MyTimerProc);
							}
						}
					}

					CoUninitialize();
				}
				catch (...) {}
			}
		}
	}
	catch (...) {}

	return CallNextHookEx(hGlobalHook, nCode, wParam, lParam);
}

__declspec(dllexport) void SetGlobalHookHandle(HHOOK hHook)
{
	hGlobalHook = hHook;
}

void __stdcall CallFrmTieuchuan()
{
	try
	{
		if (_xlCreateExcel(false))
		{
			_xlGetInfoSheetActive();

			szCellBefore = _xlGIOR(pRgActive, iRowActive, iColumnActive, L"");
			szcellAfter = szCellBefore;

			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			FrmTieuchuan dlg;
			dlg.szTukhoa = szCellBefore;
			dlg.DoModal();

			CoUninitialize();
		}		
	}
	catch (...) {}
}



void __stdcall CallFrmQuytrinh()
{
	try
	{
		if (_xlCreateExcel(false))
		{
			_xlGetInfoSheetActive();

			szCellBefore = _xlGIOR(pRgActive, iRowActive, iColumnActive, L"");
			szcellAfter = szCellBefore;

			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			FrmQuytrinh dlg;
			dlg.szTukhoa = szCellBefore;
			dlg.DoModal();

			CoUninitialize();
		}
	}
	catch (...) {}
}


void __stdcall CallFrmChiphi()
{
	try
	{
		if (_xlCreateExcel(false))
		{
			_xlGetInfoSheetActive();

			szCellBefore = _xlGIOR(pRgActive, iRowActive, iColumnActive, L"");
			szcellAfter = szCellBefore;

			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			FrmChiphi dlg;
			dlg.szTukhoa = szCellBefore;
			dlg.DoModal();

			CoUninitialize();
		}
	}
	catch (...) {}
}


void __stdcall CallDownloadQuytrinh()
{
	try
	{
		if (_xlCreateExcel(false))
		{
			_xlGetInfoSheetActive();

			if (sCodeActive == L"shQuytrinh")
			{
				FrmTieuchuan fcall;
				fcall.HotroDownloadQuytrinh();
			}

			CoUninitialize();
		}		
	}
	catch (...) {}
}


void __stdcall CallOpenURLPhaply()
{
	try
	{
		if (_xlCreateExcel(false))
		{
			_xlGetInfoSheetActive();

			if (sCodeActive == L"shQuytrinh")
			{
				int iColURL = _xlGetColumn(pShActive, L"QTR_URL", -1);
				if (iColURL == 0)
				{
					iColURL = _xlGetColumn(pShActive, L"QTR_PHAPLY", -1);
					if (iColURL > 0) iColURL++;
				}

				if (iColURL > 0)
				{
					CString szval = _xlGIOR(pRgActive, iRowActive, iColURL, L"");
					if (szval != L"")
					{
						vector<CString> vecKey;
						int nKey = TackTukhoa(vecKey, szval, L";", L";", 1, 0);

						for (int i = 0; i < nKey; i++)
						{
							try
							{
								ShellExecute(NULL, L"open", vecKey[i], NULL, NULL, SW_SHOWMAXIMIZED);
							}
							catch (...) {}
						}

						vecKey.clear();
					}
				}
			}

			CoUninitialize();
		}
	}
	catch (...) {}
}