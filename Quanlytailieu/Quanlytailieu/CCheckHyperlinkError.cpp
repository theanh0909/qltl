#include "pch.h"
#include "CCheckHyperlinkError.h"

CCheckHyperlinkError::CCheckHyperlinkError()
{
	ff = new Function;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;

	blCheckMaxLen = false;
}

CCheckHyperlinkError::~CCheckHyperlinkError()
{
	delete ff;
}

void CCheckHyperlinkError::_XacdinhSheetLienquan()
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

void CCheckHyperlinkError::_KiemtraLienketError()
{
	try
	{
		_XacdinhSheetLienquan();

		ff->_xlPutScreenUpdating(false);

		RangePtr PRS;
		CString szval = L"";
		vector<int> vecIError;
		int dem = 0, ichk = 0, itong = 0;
		int ntong = irowEnd - irowStart + 1;
		

		CString szUpdateStatus = L"Đang tiến hành kiểm tra liên kết lỗi. Vui lòng đợi trong giây lát...";
		for (int i = irowStart; i <= irowEnd; i++)
		{
			dem++;
			szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / ntong), L"%");
			ff->_xlPutStatus(szval);

			ichk = _CheckError(i, icotFilegoc);
			if (ichk >= 0)
			{
				itong++;
				if (ichk > 0) vecIError.push_back(ichk);
			}
			
			// Trường hợp là sheet văn bản --> Kiểm tra cả cột file Word
			if (jcotFolder == 0)
			{
				ichk = _CheckError(i, icotFileGXD);
				if (ichk >= 0)
				{
					itong++;
					if (ichk > 0) vecIError.push_back(ichk);
				}
			}
		}

		ff->_xlPutStatus(L"Ready");
		ff->_xlPutScreenUpdating(true);

		if (itong == 0)
		{
			ff->_msgbox(L"Không tồn tại liên kết lỗi tại sheet lựa chọn.", 0, 0);
			return;
		}

		int isize = (int)vecIError.size();
		
		szval.Format(L"Phát hiện (%d) liên kết lỗi. "
			L"Bạn dùng tính năng lọc (Auto Filter) tại cột 'Mở file' để kiểm tra. ", itong);

		if (blCheckMaxLen)
		{
			szval += L"\n- Màu đỏ: Liên kết không tìm thấy file -> Xóa dòng dữ liệu đó và thêm lại file."
				L"\n- Màu cam: Excel không hỗ trợ mở file có đường dẫn vượt quá 255 ký tự -> sửa lại đường dẫn hoặc tên file để nhỏ hơn 255 ký tự.";
		}

		if (isize > 0)
		{
			szval += L"\nBạn có muốn khôi phục các liên kết Hyperlink bị lỗi?";
			int result = ff->_y(szval, 0, 0, 0);
			if (result == 6)
			{
				CString szDir = L"", szName = L"", szType = L"";
				for (int i = 0; i < isize; i++)
				{
					szval = ff->_xlGIOR(pRgActive, vecIError[i], jcotSTT, L"");
					if (_wtoi(szval) > 0)
					{
						// Đặt lại hyperlink
						szDir = ff->_xlGIOR(pRgActive, vecIError[i], jcotFolder, L"");
						szName = ff->_xlGIOR(pRgActive, vecIError[i], jcotTen, L"");
						szType = ff->_xlGIOR(pRgActive, vecIError[i], jcotType, L"");
						if (szDir.Right(1) != L"\\") szDir += L"\\";

						szval = szDir + szName + L"." + szType;
						PRS = pRgActive->GetItem(vecIError[i], icotFilegoc);
						ff->_xlSetHyperlink(pShActive, PRS, szval);

						if (ff->_FileExistsUnicode(szval, 0) != 1)
						{
							PRS->Font->PutColor(rgbWhite);
							PRS->Interior->PutColor(rgbRed);
						}
					}
				}
			}
		}
		else
		{
			ff->_msgbox(szval, 1, 0);
		}
	}
	catch (...) {}
}

int CCheckHyperlinkError::_CheckError(int iRow, int iColumn)
{
	try
	{
		int iError = -1;
		RangePtr PRS = pRgActive->GetItem(iRow, iColumn);
		CString szval = ff->_xlGetHyperLink(PRS);
		if (szval != L"")
		{
			if (szval.Right(11) != L"desktop.ini")
			{
				if (szval.GetLength() >= 255)
				{
					iError = 0;
					blCheckMaxLen = true;
					PRS->Font->PutColor(rgbWhite);
					PRS->Interior->PutColor(rgbOrange);
				}
				else
				{
					if (ff->_FileExistsUnicode(szval, 0) != 1)
					{
						iError = iRow;
						PRS->Font->PutColor(rgbWhite);
						PRS->Interior->PutColor(rgbRed);
					}
					else
					{
						PRS->Font->PutColor(rgbBlue);
						PRS->Interior->PutColor(xlNone);
					}
				}
			}			
		}

		return iError;
	}
	catch (...) {}
	return -1;
}