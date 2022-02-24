// pch.cpp: source file corresponding to the pre-compiled header

#include "pch.h"

// When you are using pre-compiled headers, this source file is necessary for compilation to succeed.

CFont m_FontHeader;

int jColumnHand = -1;
int dPosX = 0, dPosY = 0, dWidthF = 0, dHeightF = 0;
int jXIcon = -1, jYIcon = -1;
int nIndexMouse = 0;
bool blDragging = false;
CString szAppSmart = (CString)FileSmartStartApp + L".exe";

vector<MyStrList> vecAllFiles;

int _GetListMyFiles(vector<MyStrList> &vecItem)
{
	try
	{
		vecItem.clear();

		Function *ff = new Function;

		// Read file {FileMyFiles} .ini
		CString pathFolderDir = ff->_GetPathFolder(szAppSmart);
		if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

		vector<CString> vecLine;
		ff->_ReadMultiUnicode(pathFolderDir + FileMyFiles, vecLine);

		MyStrList MSLT;

		int pos = -1;
		int ncount = (int)vecLine.size();
		CString szColor = L"";

		for (int i = 0; i < ncount; i++)
		{
			vecLine[i].Replace(L"\"", L"");
			if (vecLine[i] != L"" && vecLine[i].Left(2) != L"//")
			{
				MSLT.iID = i;
				MSLT.clrBkg = NULL;

				pos = vecLine[i].Find(L"|");
				if (pos > 0)
				{
					szColor = (vecLine[i].Right(vecLine[i].GetLength() - pos - 1)).Trim();
					if (szColor.Find(L"RGB") >= 0) MSLT.clrBkg = ff->_SetColorRGB(szColor);

					vecLine[i] = (vecLine[i].Left(pos)).Trim();
				}

				MSLT.szFullpath = vecLine[i];
				vecItem.push_back(MSLT);
			}
		}

		vecLine.clear();
		delete ff;

		return (int)vecItem.size();;
	}
	catch (...) {}
	return 0;
}
