#pragma once

#include <Gdiplus.h>
using namespace Gdiplus;
#pragma comment(lib, "gdiplus.lib")

struct MENU_ITEM_DATA
{
	HICON hIcon;
	UINT dwResIconID;
	UINT dwMenuIconID;
	SIZE szIcon;

	MENU_ITEM_DATA()
	{
		hIcon = NULL;
		dwResIconID = 0;
		dwMenuIconID = 0;
		szIcon.cx = 0;
		szIcon.cy = 0;
	}
	~MENU_ITEM_DATA()
	{
		if (hIcon)
		{
			::DestroyIcon(hIcon);
			hIcon = NULL;
		}
	}
};

class MenuIcon
{
public:
	MenuIcon();
	~MenuIcon();

protected:
	static HICON LoadIconSmart(UINT nResIconID, int nCx, int nCy);
	HRESULT Create32BitHBITMAP(HDC hdc, const SIZE * psize, void ** ppvBits, HBITMAP * phBmp);
	void InitBitmapInfo(BITMAPINFO * pbmi, ULONG cbInfo, LONG cx, LONG cy, WORD bpp);
	HRESULT ConvertBufferToPARGB32(HPAINTBUFFER hPaintBuffer, HDC hdc, HICON hicon, SIZE & sizIcon);
	bool HasAlpha(ARGB *pargb, SIZE& sizImage, int cxRow);
	HRESULT ConvertToPARGB32(HDC hdc, ARGB * pargb, HBITMAP hbmp, SIZE & sizImage, int cxRow);

public:
	BOOL bWinXP;
	HMODULE hUxTheme;

	void GdiplusStartupInputConfig();
	HRESULT(WINAPI *pfnGetBufferedPaintBits)(HPAINTBUFFER, RGBQUAD **, int*);
	HPAINTBUFFER(WINAPI *pfnBeginBufferedPaint)(HDC, const RECT *, BP_BUFFERFORMAT, BP_PAINTPARAMS *, HDC *);
	HRESULT(WINAPI *pfnEndBufferedPaint)(HPAINTBUFFER, BOOL);
	HBITMAP IconToBitmapPARGB32(UINT nResIconID, int nIconCx, int nIconCy);
	void AddIconMenuItem(HMENU &pMenu, HBITMAP &hBmp, int nIndexMenu, int nIndexIcon);

protected:
	ULONG_PTR gdiplusToken;
	virtual void PostNcDestroy();	
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

};

