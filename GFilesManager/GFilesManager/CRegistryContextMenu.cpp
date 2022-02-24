#include "pch.h"
#include "CRegistryContextMenu.h"

CRegistryContextMenu::CRegistryContextMenu()
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;
}

CRegistryContextMenu::~CRegistryContextMenu()
{
	delete ff;
	delete bb;
	delete sv;
	delete reg;
}

void CRegistryContextMenu::_StartRegContextMenu()
{
	try
	{
		CString pathFolder = ff->_GetPathFolder(szAppSmart);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		// Download file sl.exe
		CString szAppsl = pathFolder + FileSelected;
		if (ff->_FileExistsUnicode(szAppsl, 0) != 1)
		{
			CString szWin = L"32\\";
			if (ff->_Is64BitWindows()) szWin = L"64\\";

			sv->_LoadDecodeBase64();
			sv->_DownloadFile(sv->dbSeverDir + szWin + FileSelected, szAppsl);
		}
		
		CString szIcon = pathFolder + szAppSmart;
		CString szName = L"Add items with SmartStart";

		CString szContextMenu = L"";
		CString szCreate[2] = { CreateKeyCMFiles, CreateKeyCMFolders };
		for (int i = 0; i < 2; i++)
		{
			szContextMenu = bb->BaseDecodeEx(szCreate[i]);

			reg->CreateKey(szContextMenu);
			reg->WriteChar("Icon", ff->_ConvertCstringToChars(szIcon));
			reg->WriteChar("MUIVerb", ff->_ConvertCstringToChars(szName));

			if (szContextMenu.Right(1) != L"\\") szContextMenu += L"\\";
			szContextMenu += L"command";

			reg->CreateKey(szContextMenu);
			reg->WriteChar("", ff->_ConvertCstringToChars(L"\"" + szAppsl + L"\" \"%1\""));
		}

/********************************************************************************************/

		CString szAppsm = pathFolder + szAppSmart;
		szContextMenu = bb->BaseDecodeEx(CreateKeyDesktop);
		reg->CreateKey(szContextMenu);
		reg->WriteChar("Position", "Middle");
		reg->WriteChar("Icon", ff->_ConvertCstringToChars(szIcon));
		reg->WriteChar("SubCommands", "");

		// Settings
		if (szContextMenu.Right(1) != L"\\") szContextMenu += L"\\";
		CString szSub = szContextMenu + L"shell\\Settings\\";
		reg->CreateKey(szSub);
		reg->WriteChar("Icon", "SystemSettingsBroker.exe");
		reg->WriteChar("MUIVerb", "Settings");
		
		szSub += L"command";
		reg->CreateKey(szSub);
		reg->WriteChar("", ff->_ConvertCstringToChars(L"\"" + szAppsm + L"\""));

		// SmartStart
		szSub = szContextMenu + L"shell\\SmartStart\\";
		reg->CreateKey(szSub);
		reg->WriteChar("Icon", ff->_ConvertCstringToChars(szIcon));		
		reg->WriteChar("MUIVerb", "Open 'SmartStart.exe'");
		reg->WriteDword(L"CommandFlags", 32);

		szSub += L"command";
		reg->CreateKey(szSub);
		reg->WriteChar("", ff->_ConvertCstringToChars(L"\"" + szAppsm + L"\""));
	}
	catch (...) {}
}

