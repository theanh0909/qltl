#include "pch.h"
#include "CHelpRibbon.h"
#include "CDownloadFileSever.h"

CHelpRibbon::CHelpRibbon()
{
	ff = new Function;

	pathFolder = L"";
}

CHelpRibbon::~CHelpRibbon()
{
	delete ff;
}

CString CHelpRibbon::_GetLineTextHelp(int nIndex)
{
	try
	{
		CString szText = L"";

		CDownloadFileSever *sv = new CDownloadFileSever;
		sv->_LoadDecodeBase64();

		pathFolder = ff->_GetPathFolder(strNAMEDLL);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		CString szFileSaved = pathFolder + FileHelp;
		if (sv->_DownloadFile(sv->dbSeverDir + FileHelp, szFileSaved))
		{
			vector<CString> vecLine;
			int ncount = ff->_ReadMultiUnicode(szFileSaved, vecLine);
			if (ncount > 0)
			{
				for (int i = ncount - 1; i >= 0; i--)
				{
					vecLine[i].Trim();
					if (vecLine[i] == L"" || vecLine[i].Left(2) == L"//") vecLine.erase(vecLine.begin() + i);
				}				
			}

			ncount = (int)vecLine.size();
			if (nIndex > 0 && nIndex <= ncount) szText = vecLine[nIndex - 1];

			vecLine.clear();
		}

		DeleteFile(szFileSaved);

		delete sv;
		return szText;
	}
	catch (...) {}
	return L"";
}

bool CHelpRibbon::_OpenFileHelp(CString szText)
{
	try
	{
		if (szText == L"") return false;
		
		if (szText.Left(4) == L"http")
		{
			ShellExecute(NULL, L"open", szText, NULL, NULL, SW_SHOWMAXIMIZED);
		}
		else
		{
			CString szFile = pathFolder + szText;
			if (ff->_FileExistsUnicode(szFile, 0) == 1)
			{
				ShellExecute(NULL, L"open", szFile, NULL, NULL, SW_SHOWMAXIMIZED);
			}
			else return false;
		}
		return true;
	}
	catch (...) {}
	return false;
}


void CHelpRibbon::_HelpUM()
{
	try
	{
		if (!_OpenFileHelp(_GetLineTextHelp(1)))
		{
			// Tìm tệp tin tương đối
			CFileFinder _finder;
			int ncountFile = ff->_FindAllFile(_finder, pathFolder, L"*");
			if (ncountFile > 0)
			{
				CString szval = L"";
				for (int i = 0; i < ncountFile; i++)
				{
					szval = _finder.GetFilePath(i).GetFileName();
					if (szval.Find(L"HDSD") >= 0 || szval.Find(L"hdsd") >= 0)
					{
						ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
						return;
					}
				}
			}
		}
	}
	catch (...) {}
}

void CHelpRibbon::_HelpFAQ()
{
	if (!_OpenFileHelp(_GetLineTextHelp(2)))
	{
		ShellExecute(NULL, L"open", GXDVN, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void CHelpRibbon::_HelpVideo()
{
	if (!_OpenFileHelp(_GetLineTextHelp(3)))
	{
		ShellExecute(NULL, L"open", GXDVN, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void CHelpRibbon::_HelpForum()
{
	if (!_OpenFileHelp(_GetLineTextHelp(4)))
	{
		ShellExecute(NULL, L"open", GXDVN, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void CHelpRibbon::_HelpSoft()
{
	try
	{
		pathFolder = ff->_GetPathFolder(strNAMEDLL);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		vector<CString> vecLine;
		int ncount = ff->_ReadMultiUnicode(pathFolder + FileVersion, vecLine);
		if (ncount >= 5)
		{
			ff->_msgbox(L"Bạn đang sử dụng bản miễn phí - version: " + vecLine[4], 0, 0);
		}
		vecLine.clear();
	}
	catch (...) {}
}

void CHelpRibbon::_HelpSupport()
{
	if (!_OpenFileHelp(_GetLineTextHelp(6)))
	{
		// Download Teamview GXD
		ShellExecute(NULL, L"open", LinkTeamviewGXD, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void CHelpRibbon::_HelpFeedback()
{
	if (!_OpenFileHelp(_GetLineTextHelp(7)))
	{
		ShellExecute(NULL, L"open", GXDVN, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

