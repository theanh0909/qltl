#include "stdafx.h"
#include "Base64.h"
#include "Function.h"
#include "Tailieuthamkhao.h"
#include "FrmTailieuthamkhao.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CTailieuthamkhaoApp
BEGIN_MESSAGE_MAP(CTailieuthamkhaoApp, CWinApp)
END_MESSAGE_MAP()

// CTailieuthamkhaoApp construction
CTailieuthamkhaoApp::CTailieuthamkhaoApp()
{
}

CTailieuthamkhaoApp::~CTailieuthamkhaoApp()
{	
}

CTailieuthamkhaoApp theApp;

BOOL CTailieuthamkhaoApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

void __stdcall CallOpenTailieuthamkhaoDLL()
{
	try
	{
		Function ff;
		int numScale = ff._GetScaleLayoutWindow();
		if (numScale > 100) numScale -= 5;

		jScaleLayout = (double)numScale / 100;

		// Thiết lập Font Robot
		myfontCheckbot.DeleteObject();
		myfontCheckbot.CreateFont((int)(26 * jScaleLayout), 0, 0, 0, FW_NORMAL, false, false, 0,
			ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, FONTNORMAL);

		m_FontHeaderList.DeleteObject();
		m_FontHeaderList.CreateFont((int)(16 * jScaleLayout), 0, 0, 0, FW_BOLD, false, false, 0,
			ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, _T("Arial"));
	}
	catch (...) {}
}

void __stdcall OpenFrmDownloadTailieu()
{
	try
	{
		Function ff;
		ff._xlCreateExcel(1);		
		if (ff._CheckConnection() == TRUE)
		{
			xl->PutStatusBar((_bstr_t)L"Đang kiểm tra bản quyền phần mềm. Vui lòng đợi trong giây lát...");
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			FrmTailieuthamkhao pp;
			pp.DoModal();
			xl->PutStatusBar((_bstr_t)L"Ready");
		}
		else
		{
			ff._msgbox(L"Vui lòng kiểm tra kết nối mạng.", 2, 1);
		}

		ff._xlCloseExcel();
	}
	catch (...) {}
}
