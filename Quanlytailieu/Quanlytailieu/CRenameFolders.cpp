#include "pch.h"
#include "CRenameFolders.h"

CRenameFolders::CRenameFolders()
{
	ff = new Function;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
}

CRenameFolders::~CRenameFolders()
{
	delete ff;
}

void CRenameFolders::_GetDanhsachFolders()
{
	try
	{
		_XacdinhSheetLienquan();

		if (jcotFolder == 0) return;

		ff->_xlPutScreenUpdating(false);

		vector<int> vecRow;	

		int pos = -1;
		CString szval = L"", szFolder = L"";
		for (int i = irowStart; i < irowEnd; i++)
		{
			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) == 0)
			{
				// Kiểm tra xem tên thư mục có thay đổi không?
				szFolder = ff->_xlGIOR(pRgActive, i, jcotLV, L"");
				if (szFolder != L"")
				{
					szval = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
					if (szval != L"")
					{
						szval = ff->_GetNameOfPath(szval, pos, 0);
						if (szFolder != szval) vecRow.push_back(i);
					}					
				}
			}
		}

		int ncount = (int)vecRow.size();
		if (ncount > 0)
		{
			szval.Format(L"Bạn có chắc chắn đổi tên (%d) thư mục?", ncount);
			int result = ff->_y(szval, 0, 0, 0);
			if (result == 6)
			{
				RangePtr PRS;
				CString szName = L"", szType = L"", szDirOld = L"", szChange = L"";
				CString szNameOld = L"";

				for (int i = 0; i < ncount; i++)
				{
					szFolder = ff->_xlGIOR(pRgActive, vecRow[i], jcotLV, L"");
					szDirOld = ff->_xlGIOR(pRgActive, vecRow[i], jcotFolder, L"");
					szval = ff->_GetNameOfPath(szDirOld, pos, 1);					
					szChange = szval + L"\\" + szFolder;

					// Kiểm tra lại 1 lần nữa xem tên lĩnh vực có khác tên thư mục không?
					// Nguyên nhân là do trong quá trình thay đổi tên lần (i) có thể ảnh hưởng tới lần (j) tiếp theo

					szNameOld = szDirOld.Right(szDirOld.GetLength() - pos - 1);
					if (szFolder != szNameOld)
					{
						// Đổi tên thư mục trong Windows
						if ((int)_wrename(szDirOld, szChange) != 0)
						{
							PRS = pRgActive->GetItem(vecRow[i], jcotFolder);
							PRS->Interior->PutColor(rgbRed);
						}
						else
						{
							// Tiến hành tìm kiếm và đổi tên tất cả đường dẫn con nếu có
							_RenameAllChildFolder(szFolder, szDirOld, szChange);
						}
					}									
				}
			}
		}
		else
		{
			ff->_msgbox(L"Không tồn tại thư mục yêu cầu được đổi tên.", 0, 0);
		}		

		vecRow.clear();
		ff->_xlPutScreenUpdating(true);
	}
	catch (...) {}
}

void CRenameFolders::_RenameAllChildFolder(CString szNameFd, CString szDirOld, CString szChange)
{
	try
	{
		int iLen = 0, dem = 0;
		int iLOld = szDirOld.GetLength();
		int iTotal = irowEnd - irowStart + 1;
		
		RangePtr PRS;
		CString szval = L"", szDir = L"", szName = L"", szType = L"";
		CString szUpdateStatus = L"Đang tìm kiếm và thay đổi tên thư mục. Vui lòng đợi trong giây lát...";
		
		for (int i = irowStart; i < irowEnd; i++)
		{
			dem++;
			szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100*dem/iTotal), L"%");
			ff->_xlPutStatus(szval);

			szval = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
			if (szval == szDirOld)
			{
				// Đây là vị trí thư mục gốc cũ
				pRgActive->PutItem(i, jcotLV, (_bstr_t)szNameFd);
				pRgActive->PutItem(i, jcotFolder, (_bstr_t)szChange);
			}
			else
			{
				iLen = szval.GetLength();
				if (iLen > iLOld)
				{
					if (szval.Left(iLOld) == szDirOld)
					{
						// Đây là thư mục con của thư mục gốc cũ
						szval = szChange + szval.Right(iLen - iLOld);
						pRgActive->PutItem(i, jcotFolder, (_bstr_t)szval);
					}
				}
			}

			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) > 0)
			{
				// Đặt lại hyperlink
				szDir = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
				szName = ff->_xlGIOR(pRgActive, i, jcotTen, L"");
				szType = ff->_xlGIOR(pRgActive, i, jcotType, L"");
				if (szDir.Right(1) != L"\\") szDir += L"\\";

				szval = szDir + szName + L"." + szType;
				PRS = pRgActive->GetItem(i, icotFilegoc);
				ff->_xlSetHyperlink(pShActive, PRS, szval);
			}
		}

		ff->_xlPutStatus(L"Ready");
	}
	catch (...) {}
}

void CRenameFolders::_XacdinhSheetLienquan()
{
	CString szCopy = shCopy;
	jcotFolder = 0, jcotType = 0, jcotNoidung = 0;

	if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
	{
		ff->_GetThongtinSheetCategory(jcotSTT, jcotLV, icotMH, jcotTen,
			icotFilegoc, icotFileGXD, icotThLuan, icotEnd, irowTieude, irowStart, irowEnd);
	}
	else
	{
		ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
			jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
	}
}