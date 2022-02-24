#include "pch.h"
#include "CQSortData.h"

CQSortData::CQSortData()
{
	ff = new Function;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
}

CQSortData::~CQSortData()
{
	delete ff;
}

void CQSortData::_XacdinhSheetLienquan()
{
	ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
		jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
}

void CQSortData::_SapxepTatcaDulieu()
{
	try
	{
		_XacdinhSheetLienquan();
		if (jcotFolder == 0) return;	// Tính năng chỉ sử dụng tại các sheet 'Quản lý files'

		RangePtr PRS = pRgActive->GetItem(irowStart + 1, 1);
		if ((int)PRS->EntireRow->OutlineLevel <= 1)
		{
			ff->_msgbox(L"Bạn cần tạo nhóm dữ liệu trước khi sử dụng lệnh sắp xếp. "
				L"Để tạo nhóm, bạn kích chọn lệnh 'Tiện ích / Tạo nhóm dữ liệu'.", 0, 0);
			return;
		}

		ff->_xlPutScreenUpdating(false);

		bool blCheckMove = false;
		int ntong = irowEnd - irowStart + 1;
		int pos = -1, dem = 0, cong = 0, islsapxep = 0;
		CString szval = L"", szlink = L"", sznewDir = L"", szPathNew = L"", szLinhvuc = L"";
		CString szUpdateStatus = L"Đang tiến hành kiểm tra và sắp xếp lại dữ liệu. Vui lòng đợi trong giây lát...";
		for (int i = irowStart; i <= irowEnd; i++)
		{
			dem++;
			szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / ntong), L"%");
			ff->_xlPutStatus(szval);

			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) == 0)
			{
				szPathNew = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
				if (szPathNew.Right(1) != L"\\") szPathNew += L"\\";

				szLinhvuc = ff->_GetNameOfPath(szPathNew, pos, 0);
				if (szLinhvuc == L"") szLinhvuc = szPathNew;
			}
			else
			{
				cong++;
				pRgActive->PutItem(i, jcotSTT, cong);

				// Kiểm tra file có nằm đúng thư mục không?
				szval = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
				if (szval.Right(1) != L"\\") szval += L"\\";

				if (szPathNew != L"" && szval != szPathNew)
				{
					PRS = pRgActive->GetItem(i, icotFilegoc);
					szlink = ff->_xlGetHyperLink(PRS);
					if (szlink != L"")
					{
						// Kiểm tra sao chép hay di chuyển file
						blCheckMove = _CheckFileMove(i);

						sznewDir = szPathNew + ff->_GetNameOfPath(szlink, pos, 0);
						if (blCheckMove) MoveFile(szlink, sznewDir);
						else CopyFile(szlink, sznewDir, FALSE);

						// Thiết lập lại thông tin
						ff->_xlSetHyperlink(pShActive, PRS, sznewDir);

						pRgActive->PutItem(i, jcotFolder, (_bstr_t)szPathNew);
						pRgActive->PutItem(i, jcotLV, (_bstr_t)szLinhvuc);

						islsapxep++;
					}
				}
			}
		}

		ff->_xlPutStatus(L"Ready");
		ff->_xlPutScreenUpdating(true);

		if (islsapxep > 0)
		{
			szval.Format(L"Sắp xếp lại thành công (%d) dữ liệu.", islsapxep);
			ff->_msgbox(szval, 0, 0, 2000);
		}
		else
		{
			ff->_msgbox(L"Không tồn tại dữ liệu cần sắp xếp. Để sử dụng tính năng này, "
				L"bạn chọn cả dòng chứa files cần di chuyển và Cut (Ctrl + X), "
				L"sau đó dán chèn (Insert Cut) vào vị trí dòng bất kỳ tại thư mục sẽ di chuyển đến. "
				L"Chạy lệnh và kiểm tra kết quả.", 0, 0);
		}
	}
	catch (...) {}
}

bool CQSortData::_CheckFileMove(int nRow)
{
	try
	{
		int dem = 0;
		CString szID = ff->_xlGIOR(pRgActive, nRow, jcotSTT, L"");
		if (_wtoi(szID) > 0)
		{
			CString szval = L"";
			for (int i = irowStart; i <= irowEnd; i++)
			{
				szval = ff->_xlGIOR(pRgActive, nRow, jcotSTT, L"");
				if (szID == szval)
				{
					dem++;
					if (dem >= 2) return false;
				}
			}
		}
	}
	catch (...) {}
	return true;
}