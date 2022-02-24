#include "pch.h"
#include "AutoRegisterCom.h"

void AutoRegister()
{
	try
	{
		CString szRootDir = L"C:\\Test\\";

		CString szFileDLL = L"bbbbbbbb.dll";
		CString szFileTLB = L"bbbbbbbb.tlb";

		CString szNamespace = L"bbbbbbbb";
		CString szClassName = L"cccccccc";

		CString szVersion = L"1.0.0.0";
		CString szCulture = L"neutral";
		CString szPublicKeyToken = L"null";

		CString uidCLSID			= L"{3a3f738d-238f-3a8a-82b6-d5bdcdf4ddb1}";		// [default]
		CString uidInterfaceIID		= L"{1c548269-a209-3239-9685-12493dbf3e00}";		// Iiiiiiiii
		CString uidInterfaceClass	= L"{891f75f2-1423-30b1-af08-824e35f16fca}";		// _cccccccc
		CString uidLibrary			= L"{107a0a4b-1bea-4551-9011-2bf1bd9c93e3}";		// LIBID (GIUD)
		CString uidProgrammatic		= szNamespace + L"." + szClassName;					// ProgID

/***********************************************************************************************************/





	}
	catch (...) {}
}
