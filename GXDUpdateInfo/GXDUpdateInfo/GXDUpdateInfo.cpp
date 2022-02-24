#include "pch.h"
#include "framework.h"
#include "GXDUpdateInfo.h"
#include "DlgUpdateInfo.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CGXDUpdateInfoApp

BEGIN_MESSAGE_MAP(CGXDUpdateInfoApp, CWinApp)
END_MESSAGE_MAP()


// CGXDUpdateInfoApp construction

CGXDUpdateInfoApp::CGXDUpdateInfoApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only CGXDUpdateInfoApp object

CGXDUpdateInfoApp theApp;


// CGXDUpdateInfoApp initialization

BOOL CGXDUpdateInfoApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

void __stdcall InitGXDUpdateInfo()
{
	
}

void __stdcall CallGXDUpdateInfo()
{
	CFunction* ff = new CFunction;

	try
	{
		if (ff->_xlCreateExcel(true))
		{
			int numScale = ff->_GetScaleLayoutWindow();
			if (numScale > 100) numScale -= 5;

			jScaleLayout = (double)numScale / 100;

			m_FontTitleLarge.DeleteObject();
			m_FontTitleLarge.CreateFont((int)(24 * jScaleLayout), 0, 0, 0, FW_BOLD, false, false, 0,
				ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
				DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, FontArial);

			m_FontTiteForm.DeleteObject();
			m_FontTiteForm.CreateFont((int)(16 * jScaleLayout), 0, 0, 0, FW_BOLD, false, false, 0,
				ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
				DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, FontArial);

			m_FontTiteError.DeleteObject();
			m_FontTiteError.CreateFont((int)(15 * jScaleLayout), 0, 0, 0, FW_NORMAL, true, false, 0,
				ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
				DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, FontArial);

			/***********************************************************************************************/

			DlgUpdateInfo* pCall = new DlgUpdateInfo;
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			pCall->DoModal();
			delete pCall;

			CoUninitialize();
		}
	}
	catch (...) {}

	delete ff;
}

