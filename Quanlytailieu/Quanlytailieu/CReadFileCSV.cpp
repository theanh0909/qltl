#include "pch.h"
#include "CReadFileCSV.h"

CReadFileCSV::CReadFileCSV()
{
	ff = new Function;
}

CReadFileCSV::~CReadFileCSV()
{
	delete ff;
}

int CReadFileCSV::_ReadLineFile(CString pathFileCSV, vector<CString> & vecItem)
{
	try
	{
		int iSize = _GetLineFile(pathFileCSV) + 1;
		wchar_t **dArray = (wchar_t **)calloc(iSize, sizeof(wchar_t *));
		wchar_t **p = (wchar_t **)calloc(iSize, sizeof(wchar_t *));
		for (int i = 0; i < iSize; i++)
		{
			*(dArray + i) = (wchar_t *)malloc(SIZE_LINE_CSV * sizeof(wchar_t));
			*(p + i) = (wchar_t *)malloc(SIZE_LINE_CSV * sizeof(wchar_t));
		}

		FILE *pReadFile;
		if (_wfopen_s(&pReadFile, pathFileCSV, L"r, ccs=UTF-16LE") != 0)
		{
			for (int i = 0; i < iSize; i++) free(*(dArray + i));
			free(dArray);
			return 0;
		}
		else
		{
			CString szval = L"";
			int iLine = 0, iNull = 0;
			while (fgetws(*(p + iLine), ROW_LINE_CSV, pReadFile)) iLine++;
			fclose(pReadFile);

			for (int i = 1; i < iLine; i++)
			{
				wchar_t *txtStr = new wchar_t[SIZE_LINE_CSV];
				wcscpy_s(*(dArray + i), wcslen(*(p + i)) + 1, *(p + i));				
				wcscpy_s(txtStr, wcslen(*(dArray + i)) + 1, *(dArray + i));
				
				szval = txtStr;
				if (szval.Trim() != L"")
				{
					// Lưu ý pushback 'txtStr' chứ không phải szval (do đã Trim để kiểm tra NULL)
					vecItem.push_back(txtStr);
					iNull = 0;
				}
				else iNull++;

				free(txtStr);

				// Trong csv xảy ra 1 tình huống dòng cuối cùng chứa dữ liệu không phải
				// là dòng cuối cùng trong csv, do đó phải có đoạn này để ngắt
				if (iNull >= 10) break;
			}
			fclose(pReadFile);
		}

		// Free memory
		for (int i = 0; i < iSize; i++)
		{
			free(*(dArray + i));
			free(*(p + i));
		}
		free(dArray);
		free(p);

		return (int)vecItem.size();
	}
	catch (...) {}
	return 0;
}


int CReadFileCSV::_GetLineFile(CString pathFileCSV)
{
	try
	{
		FILE *pFile;
		wchar_t *c = (wchar_t *)malloc(2000 * sizeof(wchar_t));
		
		int iLine = 0;
		pFile = _wfopen(pathFileCSV, L"r, ccs=UTF-16LE");
		if (pFile != NULL)
		{
			while (fgetws(c, 2000, pFile)) iLine++;
			fclose(pFile);
		}
		free(c);
		return iLine;
	}
	catch (...) {}
	return 0;
}

int CReadFileCSV::_GetItemStr(CString szItem, vector<CString> &vecLine)
{
	try
	{
		vecLine.clear();
		szItem.Replace(L"\"", L"");
		if (szItem != L"")
		{
			int pos = szItem.Find(L'\t');	// Find 'TAB'
			while (pos >= 0)
			{
				vecLine.push_back((szItem.Left(pos)).Trim());
				szItem = szItem.Right(szItem.GetLength() - pos - 1);
				pos = szItem.Find(L'\t');
			}

			if (pos == -1)
			{
				szItem.Trim();
				vecLine.push_back(szItem);
			}
		}
		return (int)vecLine.size();
	}
	catch (...) {}
	return 0;
}