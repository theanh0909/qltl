#include "pch.h"
#include "CMoveFiles.h"

CMoveFiles::CMoveFiles()
{
	ff = new Function;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
}

CMoveFiles::~CMoveFiles()
{
	delete ff;
}

void CMoveFiles::_XacdinhSheetLienquan()
{
	ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
		jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
}

void CMoveFiles::_DichuyenDulieu()
{
	try
	{
		_XacdinhSheetLienquan();
		if (jcotFolder == 0) return;	// Tính năng chỉ sử dụng tại các sheet 'Quản lý files'

		// Xác định vùng dữ liệu được chọn
		vector<int> vecRow;
		int isize = 0, icheck = 0;
		int iResult = ff->_xlGetAllCellSelection(0, 0);
		for (int i = 0; i < iResult; i++)
		{
			icheck = 0;
			isize = (int)vecRow.size();
			if (isize > 0)
			{
				for (int j = 0; j < isize; j++)
				{
					if (vecSelect[i].row == vecRow[j])
					{
						icheck = 1;
						break;
					}
				}
			}

			if (icheck == 0)
			{
				// Kiểm tra đường dẫn có hợp lệ không?
				if (_CheckMoveFiles(vecSelect[i].row)) vecRow.push_back(vecSelect[i].row);
			}
		}

		isize = (int)vecRow.size();
		if (isize == 0)
		{
			ff->_msgbox(L"Không tồn tại dữ liệu cần di chuyển. "
				L"Để di chuyển files đến các thư mục khác, bạn chỉ cần sao chép đường dẫn sẽ di chuyển tới "
				L"và dán thay thế đường dẫn files cần chuyển tại cột 'Đường dẫn thư mục'. Chạy lệnh và kiểm tra kết quả.", 0, 0);
			return;
		}

		ff->_xlPutScreenUpdating(false);

		if (isize > 1) ff->_QuickSortArr(vecRow, 0, isize - 1);

		CString szPathFind = ff->_xlGIOR(pRgActive, vecRow[0], jcotFolder, L"");

		CString szval = L"";
		bool blChungPath = true;
		for (int i = 1; i < isize; i++)
		{
			// Kiểm tra xem tất cả có cùng 1 đường dẫn đi tới không?
			szval = ff->_xlGIOR(pRgActive, vecRow[i], jcotFolder, L"");
			if (szPathFind != szval)
			{
				blChungPath = false;
				break;
			}
		}

		int ivtChen = 0;
		CString szLinhvuc = L"";
		CString szUpdateStatus = L"Đang tiến hành di chuyển dữ liệu. Vui lòng đợi trong giây lát...";
		
		xl->PutCutCopyMode(XlCutCopyMode(false));

		if (blChungPath)
		{
			ivtChen = _XacdinhVitriChen(szPathFind);
			if (ivtChen > 0)
			{
				// Tên lĩnh vực gốc
				szLinhvuc = ff->_xlGIOR(pRgActive, ivtChen - 1, jcotLV, L"");
				int dem = 0;

				// Tất cả dữ liệu được chọn đều di chuyển về chung 1 thư mục
				for (int i = isize - 1; i >= 0; i--)
				{
					dem++;
					szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / isize), L"%");
					ff->_xlPutStatus(szval);

					if (_DichuyenFile(vecRow[i], ivtChen, szPathFind, szLinhvuc))
					{
						// Xác định lại các vị trí còn lại
						for (int j = 0; j < i; j++)
						{
							if ((vecRow[i] > vecRow[j] && vecRow[j] > ivtChen)) vecRow[j]++;
							else if ((vecRow[i] < vecRow[j] && vecRow[j] < ivtChen)) vecRow[j]--;
						}
					}					
				}

				// Đánh lại số thứ tự
				_AutoSothutu(pShActive, 1, 0, 0);

				ff->_xlPutStatus(L"Ready");
			}
		}
		else
		{
			// Trường hợp di chuyển các files đến nhiều vị trí khác nhau
			for (int i = 0; i < isize; i++)
			{
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / isize), L"%");
				ff->_xlPutStatus(szval);

				szPathFind = ff->_xlGIOR(pRgActive, vecRow[i], jcotFolder, L"");				

				// Xác định vị trí cần xác định để chèn
				if (ivtChen == 0) ivtChen = _XacdinhVitriChen(szPathFind);
				if (ivtChen > 0)
				{
					szLinhvuc = ff->_xlGIOR(pRgActive, ivtChen - 1, jcotLV, L"");
					if (_DichuyenFile(vecRow[i], ivtChen, szPathFind, szLinhvuc))
					{
						// Xác định lại các vị trí còn lại
						for (int j = i + 1; j < isize; j++)
						{
							if ((vecRow[i] > vecRow[j] && vecRow[j] > ivtChen)) vecRow[j]++;
							else if ((vecRow[i] < vecRow[j] && vecRow[j] < ivtChen)) vecRow[j]--;
						}
					}

					// Đánh lại số thứ tự
					_AutoSothutu(pShActive, 1, 0, 0);

					ivtChen = 0;
				}				
			}

			ff->_xlPutStatus(L"Ready");
		}

		RangePtr PRS = pRgActive->GetItem(vecRow[0], jcotFolder);
		PRS->Select();

		ff->_xlPutScreenUpdating(true);
		vecRow.clear();
	}
	catch (...) {}
}

bool CMoveFiles::_DichuyenFile(int iRow, int iRChen, CString szPathFind, CString szLinhvuc)
{
	try
	{
		RangePtr PRS = pRgActive->GetItem(iRow, icotFilegoc);
		CString szlink = ff->_xlGetHyperLink(PRS);
		if (szlink != L"")
		{
			// Chèn thêm 1 dòng trước vị trí chèn (Không nên sử dung Cut -> Insert Cut --> Có thể sẽ bị treo)
			PRS = pRgActive->GetItem(iRChen, 1);
			PRS->EntireRow->Insert(xlShiftDown);

			if (iRow > iRChen) iRow++;

			// Di chuyển vị trí vừa chèn
			PRS = pRgActive->GetItem(iRow, 1);
			PRS->EntireRow->Copy();

			PRS = pRgActive->GetItem(iRChen, 1);
			PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteValues, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
			xl->PutCutCopyMode(XlCutCopyMode(false));

			// Định dạng lại
			PRS = pRgActive->GetItem(iRChen, 1);
			PRS->EntireRow->Font->PutBold(false);
			PRS->EntireRow->Font->PutItalic(false);
			PRS->EntireRow->Font->PutColor(rgbBlack);
			PRS->EntireRow->Interior->PutColor(xlNone);

			pRgActive->PutItem(iRChen, jcotLV, (_bstr_t)szLinhvuc);

			int pos = -1;
			PRS = pRgActive->GetItem(iRChen, icotFilegoc);
			if (szPathFind.Right(1) != L"\\") szPathFind += L"\\";
			CString szdir = szPathFind + ff->_GetNameOfPath(szlink, pos, 0);
			ff->_xlSetHyperlink(pShActive, PRS, szdir);

			// Di chuyển file đến vị trí mới
			MoveFile(szlink, szdir);			
		}

		// Xóa dòng cũ sau khi đã di chuyển
		PRS = pRgActive->GetItem(iRow, 1);
		PRS->EntireRow->Delete(xlShiftUp);

		// Trả lại vị trí cũ (nếu có thay đổi)
		if (iRow > iRChen) iRow--;

		return true;
	}
	catch (...) {}
	return false;
}

int CMoveFiles::_XacdinhVitriChen(CString szPathFind)
{
	try
	{
		if (szPathFind != L"")
		{
			CString szval = L"";
			for (int i = irowStart; i <= irowEnd; i++)
			{
				szval = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
				if (szval == szPathFind)
				{
					szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
					if (_wtoi(szval) == 0) return i + 1;
				}
			}
		}		
	}
	catch (...) {}
	return 0;
}

bool CMoveFiles::_CheckMoveFiles(int iRow)
{
	try
	{
		CString szval = ff->_xlGIOR(pRgActive, iRow, jcotSTT, L"");
		if (_wtoi(szval) == 0) return false;

		szval = ff->_xlGIOR(pRgActive, iRow, jcotFolder, L"").Trim();
		if (szval == L"") return false;

		if (szval.Right(1) != L"\\") szval += L"\\";

		RangePtr PRS = pRgActive->GetItem(iRow, icotFilegoc);
		CString szlink = ff->_xlGetHyperLink(PRS);
		if (szlink != L"")
		{
			int pos = -1;
			szlink = ff->_GetNameOfPath(szlink, pos, 1);
			if (szlink.Right(1) != L"\\") szlink += L"\\";
			if (szval != szlink) return true;
		}
	}
	catch (...) {}
	return false;
}

void CMoveFiles::_AutoSothutu(_WorksheetPtr pSheet, int numStart, int nRowStart, int nRowEnd)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		if (nRowStart == 0) nRowStart = irowStart;
		if (nRowEnd == 0) nRowEnd = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;

		CString szval = L"";
		for (int i = nRowStart; i <= nRowEnd; i++)
		{
			szval = (ff->_xlGIOR(pRange, i, jcotTen, L"")).Trim();
			if (szval == L"")
			{
				if (jcotType > 0)
				{
					szval = (ff->_xlGIOR(pRange, i, jcotType, L"")).Trim();
				}
			}

			if (szval != L"")
			{
				pRange->PutItem(i, jcotSTT, numStart);
				numStart++;
			}
		}
	}
	catch (...) {}
}