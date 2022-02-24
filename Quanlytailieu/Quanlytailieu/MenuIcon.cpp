#include "pch.h"
#include "StdAfx.h"
#include "MenuIcon.h"
#include "Quanlytailieu.h"

MenuIcon::MenuIcon()
{
}

MenuIcon::~MenuIcon()
{
}


void MenuIcon::GdiplusStartupInputConfig()
{
	try
	{
		GdiplusStartupInput gdiplusStartupInput;
		GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

		OSVERSIONINFO osi;
		osi.dwOSVersionInfoSize = sizeof(osi);
		BOOL bGV = GetVersionEx(&osi);
		ASSERT(bGV);

		bWinXP = bGV && (osi.dwMajorVersion == 5 && osi.dwMinorVersion >= 1);
		hUxTheme = LoadLibrary(_T("UxTheme.dll"));
		if (hUxTheme)
		{
			(FARPROC&)pfnGetBufferedPaintBits = GetProcAddress(hUxTheme, "GetBufferedPaintBits");
			(FARPROC&)pfnBeginBufferedPaint = GetProcAddress(hUxTheme, "BeginBufferedPaint");
			(FARPROC&)pfnEndBufferedPaint = GetProcAddress(hUxTheme, "EndBufferedPaint");
		}
		else
		{
			pfnGetBufferedPaintBits = NULL;
			pfnBeginBufferedPaint = NULL;
			pfnEndBufferedPaint = NULL;
		}
	}
	catch (...) {}
}

HBITMAP MenuIcon::IconToBitmapPARGB32(UINT nResIconID, int nIconCx, int nIconCy)
{
	HRESULT hr = E_OUTOFMEMORY;
	HBITMAP hBmp = NULL;

	if (pfnBeginBufferedPaint &&
		pfnEndBufferedPaint)
	{
		HICON hIcon = LoadIconSmart(nResIconID, nIconCx, nIconCy);
		if (hIcon)
		{
			SIZE sizIcon;
			sizIcon.cx = nIconCx;
			sizIcon.cy = nIconCy;

			RECT rcIcon = { 0, 0, nIconCx, nIconCy };

			HDC hdcDest = CreateCompatibleDC(NULL);
			if (hdcDest)
			{
				hr = Create32BitHBITMAP(hdcDest, &sizIcon, NULL, &hBmp);
				if (SUCCEEDED(hr))
				{
					hr = E_FAIL;

					HBITMAP hbmpOld = (HBITMAP)SelectObject(hdcDest, hBmp);
					if (hbmpOld)
					{
						BLENDFUNCTION bfAlpha = { AC_SRC_OVER, 0, 255, AC_SRC_ALPHA };
						BP_PAINTPARAMS paintParams = { 0 };
						paintParams.cbSize = sizeof(paintParams);
						paintParams.dwFlags = BPPF_ERASE;
						paintParams.pBlendFunction = &bfAlpha;

						HDC hdcBuffer;
						HPAINTBUFFER hPaintBuffer = pfnBeginBufferedPaint(hdcDest, &rcIcon, BPBF_DIB, &paintParams, &hdcBuffer);
						if (hPaintBuffer)
						{
							if (DrawIconEx(hdcBuffer, 0, 0, hIcon, sizIcon.cx, sizIcon.cy, 0, NULL, DI_NORMAL))
							{
								hr = ConvertBufferToPARGB32(hPaintBuffer, hdcDest, hIcon, sizIcon);
							}
							pfnEndBufferedPaint(hPaintBuffer, TRUE);
						}
						SelectObject(hdcDest, hbmpOld);
					}
				}
				VERIFY(DeleteDC(hdcDest));
			}
			VERIFY(DestroyIcon(hIcon));

			if (FAILED(hr))
			{
				if (hBmp)
				{
					VERIFY(DeleteObject(hBmp));
					hBmp = NULL;
				}
			}
		}
	}

	return hBmp;
}

HICON MenuIcon::LoadIconSmart(UINT nResIconID, int nCx, int nCy)
{
	HICON hIcon = NULL;
	HRESULT(WINAPI *pfnLoadIconWithScaleDown)(HINSTANCE, PCWSTR, int, int, HICON *) = NULL;
	HMODULE hComCtrl32 = LoadLibrary(_T("Comctl32.dll"));
	if (hComCtrl32)
	{
		(FARPROC&)pfnLoadIconWithScaleDown = GetProcAddress(hComCtrl32, "LoadIconWithScaleDown");
	}

	if (pfnLoadIconWithScaleDown)
	{
		pfnLoadIconWithScaleDown(theApp.m_hInstance, MAKEINTRESOURCE(nResIconID), nCx, nCy, &hIcon);
	}
	else
	{
		hIcon = (HICON)::LoadImage(theApp.m_hInstance, MAKEINTRESOURCE(nResIconID), IMAGE_ICON, nCx, nCy, 0);
	}

	return hIcon;
}

HRESULT MenuIcon::Create32BitHBITMAP(HDC hdc, const SIZE *psize, void **ppvBits, HBITMAP* phBmp)
{
	*phBmp = NULL;

	BITMAPINFO bmi;
	InitBitmapInfo(&bmi, sizeof(bmi), psize->cx, psize->cy, 32);

	HDC hdcUsed = hdc ? hdc : ::GetDC(NULL);
	if (hdcUsed)
	{
		*phBmp = CreateDIBSection(hdcUsed, &bmi, DIB_RGB_COLORS, ppvBits, NULL, 0);
		if (hdc != hdcUsed)
		{
			::ReleaseDC(NULL, hdcUsed);
		}
	}

	return (NULL == *phBmp) ? E_OUTOFMEMORY : S_OK;
}

void MenuIcon::InitBitmapInfo(__out_bcount(cbInfo) BITMAPINFO *pbmi, ULONG cbInfo, LONG cx, LONG cy, WORD bpp)
{
	ZeroMemory(pbmi, cbInfo);
	pbmi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	pbmi->bmiHeader.biPlanes = 1;
	pbmi->bmiHeader.biCompression = BI_RGB;

	pbmi->bmiHeader.biWidth = cx;
	pbmi->bmiHeader.biHeight = cy;
	pbmi->bmiHeader.biBitCount = bpp;
}

HRESULT MenuIcon::ConvertBufferToPARGB32(HPAINTBUFFER hPaintBuffer, HDC hdc, HICON hicon, SIZE& sizIcon)
{
	HRESULT hr;
	RGBQUAD *prgbQuad;
	int cxRow;
	if (pfnGetBufferedPaintBits)
	{
		hr = pfnGetBufferedPaintBits(hPaintBuffer, &prgbQuad, &cxRow);
		if (SUCCEEDED(hr))
		{
			ARGB *pargb = reinterpret_cast<ARGB *>(prgbQuad);
			if (!HasAlpha(pargb, sizIcon, cxRow))
			{
				ICONINFO info;
				if (GetIconInfo(hicon, &info))
				{
					if (info.hbmMask)
					{
						hr = ConvertToPARGB32(hdc, pargb, info.hbmMask, sizIcon, cxRow);
					}

					if (info.hbmColor)
						VERIFY(DeleteObject(info.hbmColor));

					if (info.hbmMask)
						VERIFY(DeleteObject(info.hbmMask));
				}
			}
		}
	}
	else
	{
		hr = E_NOTIMPL;
	}

	return hr;
}

bool MenuIcon::HasAlpha(ARGB *pargb, SIZE& sizImage, int cxRow)
{
	ULONG cxDelta = cxRow - sizImage.cx;
	for (ULONG y = sizImage.cy; y; --y)
	{
		for (ULONG x = sizImage.cx; x; --x)
		{
			if (*pargb++ & 0xFF000000)
			{
				return true;
			}
		}

		pargb += cxDelta;
	}

	return false;
}

HRESULT MenuIcon::ConvertToPARGB32(HDC hdc, ARGB *pargb, HBITMAP hbmp, SIZE& sizImage, int cxRow)
{
	BITMAPINFO bmi;
	InitBitmapInfo(&bmi, sizeof(bmi), sizImage.cx, sizImage.cy, 32);

	HRESULT hr = E_OUTOFMEMORY;
	HANDLE hHeap = GetProcessHeap();
	void *pvBits = HeapAlloc(hHeap, 0, bmi.bmiHeader.biWidth * 4 * bmi.bmiHeader.biHeight);
	if (pvBits)
	{
		hr = E_UNEXPECTED;
		if (GetDIBits(hdc, hbmp, 0, bmi.bmiHeader.biHeight, pvBits, &bmi, DIB_RGB_COLORS) == bmi.bmiHeader.biHeight)
		{
			ULONG cxDelta = cxRow - bmi.bmiHeader.biWidth;
			ARGB *pargbMask = static_cast<ARGB *>(pvBits);

			for (ULONG y = bmi.bmiHeader.biHeight; y; --y)
			{
				for (ULONG x = bmi.bmiHeader.biWidth; x; --x)
				{
					if (*pargbMask++)
					{
						*pargb++ = 0;
					}
					else
					{
						*pargb++ |= 0xFF000000;
					}
				}
				pargb += cxDelta;
			}
			hr = S_OK;
		}
		HeapFree(hHeap, 0, pvBits);
	}

	return hr;
}

void MenuIcon::PostNcDestroy()
{
	GdiplusShutdown(gdiplusToken);
	if (hUxTheme) ::FreeLibrary(hUxTheme);
	PostNcDestroy();
}

void MenuIcon::AddIconMenuItem(HMENU &pMenu, HBITMAP &hBmp, int nIndexMenu, int nIndexIcon)
{
	try
	{
		hBmp = NULL;
		CArray<MENU_ITEM_DATA> arrMenuIcons;

		MENUITEMINFO mii = { 0 };
		mii.cbSize = sizeof(mii);

		int SIZE_ICON = SIZE_MENU_ICON;

		if (!bWinXP)
		{
			hBmp = IconToBitmapPARGB32(nIndexIcon, SIZE_ICON, SIZE_ICON);
			mii.fMask = MIIM_BITMAP;
			mii.hbmpItem = hBmp;
		}
		else
		{
			MENU_ITEM_DATA mid;

			mid.dwMenuIconID = nIndexMenu;
			mid.dwResIconID = nIndexIcon;

			mid.szIcon.cx = SIZE_ICON;
			mid.szIcon.cy = SIZE_ICON;

			mid.hIcon = LoadIconSmart(mid.dwResIconID, mid.szIcon.cx, mid.szIcon.cy);
			int nIndex = arrMenuIcons.Add(mid);

			mii.hbmpItem = HBMMENU_CALLBACK;
			mii.fMask = MIIM_BITMAP | MIIM_DATA | MIIM_ID;
			mii.wID = mid.dwMenuIconID;
			mii.dwItemData = (ULONG_PTR)&arrMenuIcons.GetAt(nIndex);
		}

		SetMenuItemInfo(pMenu, nIndexMenu, FALSE, &mii);
	}
	catch (...) {}
}

LRESULT MenuIcon::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_MEASUREITEM)
	{
		if (wParam == 0)
		{
			MEASUREITEMSTRUCT* pMS = (MEASUREITEMSTRUCT*)lParam;
			if (pMS &&
				pMS->CtlType == ODT_MENU)
			{
				MENU_ITEM_DATA* pMID = (MENU_ITEM_DATA*)pMS->itemData;
				ASSERT(pMID);

				if (pMS->itemID == pMID->dwMenuIconID)
				{
					pMS->itemHeight = pMID->szIcon.cy;
					pMS->itemWidth = pMID->szIcon.cx;

					return TRUE;
				}
			}
		}
	}
	else if (message == WM_DRAWITEM)
	{
		DRAWITEMSTRUCT* pDS = (DRAWITEMSTRUCT*)lParam;
		if (pDS)
		{
			if (pDS->CtlType == ODT_MENU)
			{
				MENU_ITEM_DATA* pMID = (MENU_ITEM_DATA*)pDS->itemData;
				ASSERT(pMID);

				if (pDS->itemID == pMID->dwMenuIconID)
				{
					HICON hIcon = LoadIconSmart(pMID->dwResIconID, pMID->szIcon.cx, pMID->szIcon.cy);
					if (hIcon)
					{
						VERIFY(DrawIconEx(pDS->hDC,
							pDS->rcItem.left - GetSystemMetrics(SM_CXMENUCHECK) / 2,
							pDS->rcItem.top + (pDS->rcItem.bottom - pDS->rcItem.top - pMID->szIcon.cy) / 2,
							hIcon,
							pMID->szIcon.cx,
							pMID->szIcon.cy,
							0, NULL, DI_NORMAL));

						VERIFY(::DestroyIcon(hIcon));

						return TRUE;
					}
				}
			}
		}
	}

	return WindowProc(message, wParam, lParam);
}
