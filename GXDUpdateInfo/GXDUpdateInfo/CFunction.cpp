#include "pch.h"
#include "CFunction.h"

int MessageBoxTimeoutA(IN HWND hWnd, IN LPCSTR lpText, IN LPCSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds)
{
	static MSGBOXAAPI MsgBoxTOA = NULL;
	if (!MsgBoxTOA)
	{
		HMODULE hUser32 = GetModuleHandle(_T("user32.dll"));
		if (hUser32) MsgBoxTOA = (MSGBOXAAPI)GetProcAddress(hUser32, "MessageBoxTimeoutA");
		else return 0;
	}
	if (MsgBoxTOA) return MsgBoxTOA(hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
	return 0;
}

int MessageBoxTimeoutW(IN HWND hWnd, IN LPCWSTR lpText, IN LPCWSTR lpCaption,
	IN UINT uType, IN WORD wLanguageId, IN DWORD dwMilliseconds)
{
	static MSGBOXWAPI MsgBoxTOW = NULL;
	if (!MsgBoxTOW)
	{
		HMODULE hUser32 = GetModuleHandle(_T("user32.dll"));
		if (hUser32) MsgBoxTOW = (MSGBOXWAPI)GetProcAddress(hUser32, "MessageBoxTimeoutW");
		else return 0;
	}
	if (MsgBoxTOW) return MsgBoxTOW(hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
	return 0;
}

CFunction::CFunction()
{
	reg = new CRegistry;
}

CFunction::~CFunction()
{
	delete reg;
}

void CFunction::_msgbox(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);

		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void CFunction::_s(CString message, int itype, int iexcel, int itimeout)
{
	try
	{
		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

		UINT style[3] = { MB_OK | MB_ICONINFORMATION, MB_OK | MB_ICONWARNING, MB_OK | MB_ICONERROR };
		MessageBoxTimeout(NULL, message, MS_THONGBAO, style[itype], 0, itimeout);

		if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);
	}
	catch (...)
	{
		MessageBox(NULL, DF_ERROR, MS_THONGBAO, MB_OK | MB_ICONERROR);
	}
}

void CFunction::_d(int num, int itype, int iexcel, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, iexcel, itimeout);
}

void CFunction::_n(CString message, int num, int itype, int istr, int iexcel, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (istr == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, iexcel, itimeout);
}

int CFunction::_y(CString message, int istyle, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT qs[2] = { MB_YESNO, MB_YESNOCANCEL };
	UINT style[2] = { MB_ICONQUESTION, MB_ICONWARNING };

	return (int)MessageBoxTimeout(NULL, message, MS_THONGBAO,
		qs[istyle] | style[itype] | MB_DEFBUTTON1, 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);

	return 6;
}

/***********************************************************************************************/

bool CFunction::_xlCreateExcel(bool blOpenFrm)
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(ExcelApp);
		xl->PutVisible(VARIANT_TRUE);
		if (blOpenFrm) xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	}
	catch (...) {
		return false;
	}

	return true;
}


void CFunction::_xlCallFunctionDLL(CString szDllName, CString szFunction, bool blQuit)
{
	try
	{
		typedef bool(__stdcall *p)();									// ...__stdcall *p)(int opt) nếu có đối là 'opt'
		HINSTANCE loadDLL = LoadLibrary((_bstr_t)szDllName);			// Ex: UtilityQLCL9.dll

		if (loadDLL != NULL)
		{
			p pc = (p)GetProcAddress(loadDLL, (_bstr_t)szFunction);		// Ex: OpenFrmInhoso
			if (pc != NULL) pc();
		}

		if (blQuit) FreeLibrary(loadDLL);
	}
	catch (...) {}
}

int CFunction::_GetScaleLayoutWindow()
{
	try
	{
		reg->CreateKey(L"SOFTWARE\\GXDJSC\\ASCO\\Settings\\");

		int numScale = 100;
		CString szval = reg->ReadString(L"ScaleLayoutInWindows", L"");
		if (szval != L"")
		{
			numScale = _wtoi(szval);
			if (numScale < 100) numScale = 100;
			else if (numScale > 200) numScale = 200;
		}

		return numScale;
	}
	catch (...) {}
	return 100;
}