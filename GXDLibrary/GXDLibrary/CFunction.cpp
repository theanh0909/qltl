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

}

CFunction::~CFunction()
{

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

void CFunction::_msgupdate()
{
	_msgbox(L"Tính năng đang phát triển. Vui lòng đợi ở phiên bản cập nhật tiếp theo.", 0);
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

		pWb = xl->GetActiveWorkbook();
	}
	catch (...) {
		_msgbox(L"Lỗi khởi tạo ứng dụng Microsoft Excel. " 
			L"Nguyên nhân có thể do có Excel chạy ngầm, " 
			L"bạn vào TaskManager để tắt đi sau đó khởi động lại phần mềm.", 2);
		return false;
	}

	return true;
}

CString CFunction::_xlGetVersionExcel()
{
	CString szVersion = (LPCTSTR)xl->GetVersion();
	return szVersion;

}
void CFunction::_xlDeveloperSettings()
{
	try
	{
		CRegistry *reg = new CRegistry;
		CString szVersion = _xlGetVersionExcel();

		reg->CreateKey(L"Software\\Microsoft\\Office\\" + szVersion + L"\\Excel\\Security");
		reg->WriteDword(L"VBAWarnings", 1);
		reg->WriteDword(L"AccessVBOM", 1);

		reg->CreateKey(L"Software\\Microsoft\\Office\\" + szVersion + L"\\Common\\Security");
		reg->WriteDword(L"DisableHyperlinkWarning", 1);

		delete reg;
	}
	catch (...) {}
}


bool CFunction::_xlIsOperatingSystem64()
{
	CoInitialize(NULL);

	_ApplicationPtr objExcel;
	if (FAILED(objExcel.GetActiveObject(ExcelApp)))
	{
		if (FAILED(objExcel.CreateInstance(ExcelApp)))
		{
			return false;
		}
	}

	if (!(objExcel == NULL))
	{
		CString szOperatingSystem = (LPCTSTR)objExcel->GetOperatingSystem();
		if (szOperatingSystem.Find(L"64") >= 0) return true;
		else return false;
	}

	CoUninitialize();
	return false;
}


void CFunction::_xlPutStatus(CString szStatus)
{
	xl->PutStatusBar((_bstr_t)szStatus);
}

_bstr_t CFunction::_xlGetNameSheet(_bstr_t codename, bool blError)
{
	try
	{
		_WorksheetPtr psh;
		int ncountSheet = xl->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= ncountSheet; i++)
		{
			psh = pWb->Worksheets->Item[i];
			if (i == 1)
			{
				if ((int)psh->GetVisible() == 2)
				{
					if (psh->Name == (_bstr_t)shNegs)
					{
						_msgbox(L"Phát hiện virut NEGS.xls trong file đang sử dụng. "
							L"Virut có thể gây ảnh hưởng đến quá trình sử dụng. "
							L"Vui lòng kiểm tra lại hoặc liên hệ với "
							L"tổng đài hỗ trợ GXD 1900.0147 để được xử lý.",
							1, 1);
						return L"";
					}
				}
			}

			if (psh->CodeName == codename) return psh->Name;
		}
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy bảng tính có codename '%s'. Vui lòng kiểm tra lại.", (LPCTSTR)codename);
		if (blError) _msgbox(str, 2);
	}
	return L"";
}

int CFunction::_xlGetRow(_WorksheetPtr pSheet, CString name, int itype)
{
	try
	{
		RangePtr dRangeptr = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		return (int)dRangeptr->Cells->Row;
	}
	catch (...)
	{
		CString szval = DF_ERROR;
		int irow = 0, icolumn = 0;

		_xlCreateNewName(name, szval, irow, icolumn);

		if (irow > 0) return irow;
		if (itype == 1) _xlMsgNotName(pSheet, name);
	}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
}


int CFunction::_xlGetColumn(_WorksheetPtr pSheet, CString name, int itype)
{
	try
	{
		RangePtr dRangeptr = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		return (int)dRangeptr->Cells->Column;
	}
	catch (...)
	{
		CString szval = DF_ERROR;
		int irow = 0, icolumn = 0;

		_xlCreateNewName(name, szval, irow, icolumn);

		if (icolumn > 0) return icolumn;
		if (itype == 1) _xlMsgNotName(pSheet, name);
	}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn() + 1;
}


int CFunction::_xlGetTableColumnLast(_WorksheetPtr pSheet, int rowGet, int icolStart, int icolEnd, int itype)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		RangePtr pCell;

		int iColumnLast = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn();
		if (icolEnd == 0) icolEnd = iColumnLast;

		int colEnd = iColumnLast;

		for (int i = icolStart; i <= icolEnd; i++)
		{
			pCell = pRange->GetItem(rowGet, i);
			if ((int)pCell->Borders->GetItem(XlBordersIndex::xlEdgeRight)->GetWeight() == 2
				&& (int)pCell->Borders->GetItem(XlBordersIndex::xlEdgeRight)->GetLineStyle() == 1)
			{
				colEnd = i;
				if (itype == 1) break;
			}
		}

		return colEnd;
	}
	catch (...) {}
	return 0;
}


int CFunction::_xlGetTableRowLast(_WorksheetPtr pSheet, int colGet, int irowStart, int irowEnd, int itype)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		RangePtr pCell;

		int iRowLast = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow();
		if (irowEnd == 0) irowEnd = iRowLast;

		int rowEnd = iRowLast;

		for (int i = irowStart; i <= irowEnd; i++)
		{
			pCell = pRange->GetItem(i, colGet);
			if ((int)pCell->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->GetWeight() == 2
				&& (int)pCell->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->GetLineStyle() == 1)
			{
				rowEnd = i;
				if (itype == 1) break;
			}
		}

		return rowEnd;
	}
	catch (...) {}
	return 0;
}


double CFunction::_xlTotalWidthColumn(RangePtr pRange, int iColumnStart, int iColumnEnd)
{
	RangePtr pCell;
	double dbTotal = 0;
	for (int i = iColumnStart; i <= iColumnEnd; i++)
	{
		pCell = pRange->GetItem(1, i);
		dbTotal += (double)pCell->GetWidth();
	}

	return dbTotal + 0.71;
}


double CFunction::_xlGetWidthStringExcel(CString szNoidung, CString szFont, double iSize)
{
	int pos = -1;
	double db = 0;
	int nsize = getSizeArray(dWTimesNewRoman);
	szNoidung = _ConvertKhongDau(szNoidung);
	for (int i = 0; i < szNoidung.GetLength(); i++)
	{
		pos = szASCII.Find(szNoidung.Mid(i, 1));
		if (pos == -1) pos = 0;
		if (pos < nsize) db += (double)(iSize*dWTimesNewRoman[pos]);		
	}

	return db;
}


void CFunction::_xlAutoMergeCellNew(_WorksheetPtr pSheet, int iRowStart, int iColumnStart, int iColumnEnd, double iTotalW, int itype)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		CString szChuoi = _xlGIOR(pRange, iRowStart, iColumnStart, L"");
		if (szChuoi == L"") return;

		RangePtr pCell = pRange->GetItem(iRowStart, iColumnStart);
		CString _szFont = (LPCTSTR)(_bstr_t)pCell->Font->GetName();
		float izSize = (float)pCell->Font->GetSize();
		if (_szFont != FontTimes) return;

		if (iTotalW == 0) iTotalW = _xlTotalWidthColumn(pRange, iColumnStart, iColumnEnd);
		double dbChuoi = _xlGetWidthStringExcel(szChuoi, _szFont, izSize);

		if (dbChuoi >= iTotalW)
		{
			int nLine = (int)(dbChuoi / iTotalW) + 1;
			if (nLine > 1)
			{
				if (itype == 0) pCell->PutRowHeight(20 * nLine);
				else
				{
					if ((double)(20 * nLine) > (double)pCell->GetRowHeight()) pCell->PutRowHeight(20 * nLine);
				}

				pCell = pSheet->GetRange(pRange->GetItem(iRowStart, iColumnStart), pRange->GetItem(iRowStart, iColumnEnd));
				if ((int)pCell->GetMergeCells() != 1) pCell->PutMergeCells(1);
				if ((int)pCell->GetWrapText() != 1) pCell->PutWrapText(1);
				pCell->PutHorizontalAlignment(xlJustify);
			}
		}
		else
		{
			if (itype == 0) pCell->PutRowHeight(20);
		}
	}
	catch (...)
	{
		CString szMess = L"";
		szMess.Format(L"Lỗi tự động gộp căn chỉnh tại dòng số %d, bắt đầu từ cột '%s' đến cột '%s'. "
			L"Bạn có thể nhập từ khóa trên Google: 'auto merge + gxd' để xử lý lỗi.",
			iRowStart, _xlConvertRCtoA1(iColumnStart), _xlConvertRCtoA1(iColumnEnd));
		_msgbox(szMess, 2, 1);
	}
}


void CFunction::_xlSetHyperlink(_WorksheetPtr pSheet, RangePtr pCell, CString pathFile, CString szName,
	COLORREF clrBkg, COLORREF clrTxt, CString szFontName, int iSize, bool bCenter, bool bFormula)
{
	try
	{
		pSheet->Hyperlinks->Add(pCell, (_bstr_t)pathFile);
		pCell->PutValue2((_bstr_t)szName);
		pCell->Interior->PutColor(clrBkg);
		pCell->Font->PutColor(clrTxt);
		pCell->Font->PutName((_bstr_t)szFontName);
		pCell->Font->PutSize(iSize);
		if (bCenter) pCell->PutHorizontalAlignment(xlCenter);
	}
	catch (...) {}
}


CString CFunction::_xlGetHyperLink(RangePtr pRgSelect, int iType)
{
	try
	{
		CString szlink = L"", szsub = L"";
		int ncount = (int)pRgSelect->Hyperlinks->GetCount();
		if (ncount > 0)
		{
			szlink = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(1)->GetAddress();
			szsub = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(1)->GetSubAddress();
			if (szsub != L"") szlink += (L"#" + szsub);

			if (iType == 1) szlink.Replace(L"/", L"\\");
			else if (iType == 2) szlink.Replace(L"\\", L"/");
		}
		return szlink;
	}
	catch (...) {}
	return L"";
}


int CFunction::_xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType, int iStatus)
{
	try
	{
		vecHyper.clear();
		int ncount = (int)pRgSelect->Hyperlinks->GetCount();
		if (ncount > 0)
		{
			CString szval = L"", szlink = L"", szsub = L"";
			CString szUpdateStatus = L"Đang tiến hành lấy thông tin dữ liệu. Vui lòng đợi trong giây lát...";

			for (int i = 1; i <= ncount; i++)
			{
				szlink = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(i)->GetAddress();
				szsub = (LPCTSTR)pRgSelect->Hyperlinks->GetItem(1)->GetSubAddress();
				if (szsub != L"") szlink += (L"#" + szsub);

				if (iType == 1) szlink.Replace(L"/", L"\\");
				else if (iType == 2) szlink.Replace(L"\\", L"/");

				if (iStatus == 1)
				{
					szval.Format(L"%s (%d%s) %s", szUpdateStatus,
						(int)(100 * i / ncount), L"%", szlink);
					if (szval.GetLength() > 250) szval = szval.Left(250) + L"...";
					_xlPutStatus(szval);
				}

				vecHyper.push_back(szlink);
			}

			if (iStatus == 1) _xlPutStatus(L"Ready");
		}
		return (int)vecHyper.size();
	}
	catch (...) {}
	return 0;
}


void CFunction::_xlCreateNewName(CString szName, CString &szKpname, int &iKprow, int &iKpcol)
{
	try
	{
		

	}
	catch (...) {}
}


void CFunction::_xlMsgNotName(_WorksheetPtr pSheet, CString name)
{
	CString str = L"";
	str.Format(L"Không tồn tại name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
	_msgbox(str, 1, 1);
}

CString CFunction::_xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError)
{
	try
	{
		_bstr_t bstr = pRange->GetItem(nRow, nColumn);
		return (LPCTSTR)bstr;
	}
	catch (...) { return szError; }
}

int CFunction::_xlFindComment(_WorksheetPtr pSheet, 
	int column, int rowStart, int rowEnd, _bstr_t comment, int itype)
{
	try
	{
		int iRow = 0;
		CString szval = L"";
		RangePtr pCell, pRange = pSheet->Cells;

		if (rowStart <= 0) rowStart = 1;
		if (rowEnd <= 0) rowEnd = MAX_ROW;
		szval.Format(L"%s:%s", _xlGAR(pRange, rowStart, column, 0), _xlGAR(pRange, rowEnd, column, 0));
		if (pCell = pRange->GetRange(COleVariant(szval))->Find(comment, vtMissing, XlFindLookIn::xlComments, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false))
		{
			iRow = pCell->Cells->Row;
		}

		pRange->GetRange(COleVariant(L"A1:A1"))->Find("", vtMissing, XlFindLookIn::xlFormulas, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false);

		if (iRow > 0) return iRow;
	}
	catch (...)
	{
		if (comment != (_bstr_t)L"END")
		{
			CString str = L"";
			str.Format(L"Không tìm thấy comment '%s' trong bảng tính '%s'.", (LPCTSTR)comment, (LPCTSTR)pSheet->Name);
			if (itype == 1) _msgbox(str, 1);
		}
	}

	if (itype == -1) return 0;
	return (int)pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
}

CString CFunction::_xlGAR(RangePtr pRng, int nRow, int nColumn, int itype)
{
	try
	{
		_bstr_t szAdress = L"";
		RangePtr dRangeptr = pRng->GetItem(nRow, nColumn);

		if (itype < 0 || itype > 3) itype = 0;

		switch (itype)
		{
			case 0:
				szAdress = dRangeptr->GetAddress(false, false, XlReferenceStyle::xlA1);
				break;

			case 1:
				szAdress = dRangeptr->GetAddress(true, false, XlReferenceStyle::xlA1);
				break;

			case 2:
				szAdress = dRangeptr->GetAddress(false, true, XlReferenceStyle::xlA1);
				break;

			case 3:
				szAdress = dRangeptr->GetAddress(true, true, XlReferenceStyle::xlA1);
				break;
		}

		return (LPCTSTR)szAdress;
	}
	catch (...) {}
	return L"";
}


bool CFunction::_xlConvertA1toRC(CString szValueA1, int &nRow, int &nColumn)
{
	try
	{
		szValueA1.Trim();
		if (szValueA1 == L"") return false;
		if (_wtoi(szValueA1.Right(1)) == 0 && szValueA1.Right(1) != L"0") szValueA1 += L"1";

		CString szval = L"";
		CString sr[3] = { L"_", L"_", L"A" };
		int num[3] = { 0, 0, 1 };

		for (int i = 0; i < szValueA1.GetLength(); i++)
		{
			if (_wtoi(szValueA1.Mid(i, 1)) > 0)
			{
				nRow = _wtoi(szValueA1.Right(szValueA1.GetLength() - i));
				szval = szValueA1.Left(i);

				sr[2] = szval.Right(1);

				int len = szval.GetLength();
				if (len == 2) sr[1] = szval.Left(1);
				else if (len == 3)
				{
					sr[0] = szval.Left(1);
					sr[1] = szval.Mid(1, 1);
				}

				break;
			}
		}

		CString szABC = L"_" + scvtABC;
		if (szval != L"")
		{
			for (int i = 0; i < 3; i++)
			{
				for (int j = 0; j < 27; j++)
				{
					if (sr[i] == szABC.Mid(j, 1))
					{
						num[i] = j;
						break;
					}
				}
			}

			nColumn = 676 * num[0] + 26 * num[1] + num[2];
			if (nColumn > 0) return true;
		}
	}
	catch (...) {}
	return false;
}


CString CFunction::_xlConvertRCtoA1(int iColumn)
{
	try
	{
		CString szval = L"";
		CString szCol = L"_" + scvtABC;

		if (iColumn <= 26) return szCol.Mid(iColumn, 1);
		else if (iColumn > 26 && iColumn <= 702)
		{
			int num1 = (int)(iColumn / 26);
			int num2 = iColumn % 26;
			if (num2 == 0)
			{
				num1--;
				num2 = 26;
			}

			szval.Format(L"%s%s", szCol.Mid(num1, 1), szCol.Mid(num2, 1));
			return szval;
		}
		else
		{
			int num1 = (int)(iColumn / 676);
			int mod = iColumn % 676;

			int num2 = (int)(mod / 26);
			if (num2 == 0)
			{
				num1--;
				num2 = 26;
			}

			int num3 = mod % 26;
			if (num3 == 0)
			{
				num2--;
				num3 = 26;
			}

			szval.Format(L"%s%s%s", szCol.Mid(num1, 1), szCol.Mid(num2, 1), szCol.Mid(num3, 1));
			return szval;
		}
	}
	catch (...) {}
	return L"";
}

void CFunction::_xlFormatClearFormat(RangePtr pRange)
{
	pRange->ClearContents();
	pRange->ClearComments();
	pRange->Font->PutBold(false);
	pRange->Font->PutItalic(false);
	pRange->Font->PutColor(rgbBlack);
	pRange->Interior->PutColor(xlNone);
}

void CFunction::_xlFormatBorders(RangePtr pRange, int iStyle, bool blTop, bool blBottom)
{
	try
	{
		pRange->Borders->GetItem(XlBordersIndex::xlEdgeLeft)->PutWeight(xlThin);
		pRange->Borders->GetItem(XlBordersIndex::xlEdgeRight)->PutWeight(xlThin);
		pRange->Borders->GetItem(XlBordersIndex::xlInsideVertical)->PutWeight(xlThin);

		if (blTop == true) pRange->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
		if (blBottom == true) pRange->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);

		if (iStyle == 1)
		{
			pRange->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
		}
		else if (iStyle == 2)
		{
			pRange->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlHairline);
			if (blBottom == false) pRange->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlHairline);
		}
		else if (iStyle == 3)
		{
			pRange->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlThin);
			pRange->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutLineStyle(xlDot);

			if (blBottom == false)
			{
				pRange->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
				pRange->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutLineStyle(xlDot);
			}
		}
	}
	catch (...) {}
}

void CFunction::_xlPutScreenUpdating(bool bl)
{
	if (bl == false) xl->PutScreenUpdating(VARIANT_FALSE);
	else xl->PutScreenUpdating(VARIANT_TRUE);
}

CString CFunction::_xlGetcomment(RangePtr pRange)
{
	if (pRange->GetComment() != NULL)
	{
		return (LPCTSTR)pRange->GetComment()->Text();
	}
	return L"";
}

void CFunction::_xlPutcomment(RangePtr pRange, CString szComment)
{
	try
	{
		if (pRange->GetComment() != NULL) pRange->ClearComments();
		pRange->AddComment((_bstr_t)szComment);
	}
	catch (...) {}
}


CString CFunction::_xlGetValue(_WorksheetPtr pSheet, CString name, int itype, int inumberformat)
{
	try
	{
		if (inumberformat == 1)
		{
			RangePtr pCell = pSheet->Cells;
			return _xlGIOR(pCell, _xlGetRow(pSheet, name, 1), _xlGetColumn(pSheet, name, 0), L"");
		}
		else return (LPCTSTR)(_bstr_t)pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name)->GetValue2();
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
		if (itype == 1) _msgbox(str, 1);
	}

	return L"";
}


void CFunction::_xlPutValue(_WorksheetPtr pSheet, CString name, CString szValue, int itype)
{
	try
	{
		pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name)->PutValue2((_variant_t)szValue);
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
		if (itype == 1) _msgbox(str, 1);
	}
}


void CFunction::_xlPutValue(_WorksheetPtr pSheet, CString name, int iValue, int itype)
{
	try
	{
		CString szval = _ConvertNumberToString(iValue);
		RangePtr pCell = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		pCell->PutValue2((_variant_t)szval);
	}
	catch (...)
	{
		CString str = L"";
		str.Format(L"Không tìm thấy name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
		if (itype == 1) _msgbox(str, 1);
	}
}


void CFunction::_xlPutInputMessage(RangePtr pRange, CString szNoidung, CString szHuongdan)
{
	try
	{
		ValidationPtr Vali = pRange->Validation;

		Vali->Delete();
		Vali->Add(xlValidateInputOnly, xlValidAlertStop, xlBetween);
		Vali->PutInputTitle((_bstr_t)szHuongdan);
		Vali->PutInputMessage((_bstr_t)szNoidung);		
		Vali->PutShowError(false);
		Vali->PutShowInput(true);
	}
	catch (...) {}
}

void CFunction::_xlGetInfoSheetActive()
{
	pShActive = pWb->ActiveSheet;
	pRgActive = pShActive->Cells;
	szCodeActive = (LPCTSTR)pShActive->CodeName;
	szNameActive = (LPCTSTR)pShActive->Name;
	bkgColorActive = pShActive->Tab->GetColor();

	iRowActive = pShActive->Application->ActiveCell->GetRow();
	iColumnActive = pShActive->Application->ActiveCell->GetColumn();
}

bool CFunction::_xlCreateNewSheet(_WorksheetPtr &wPsh, int nIndex, 
	CString szCodename, CString szName, COLORREF clrTab)
{
	try
	{
		// Total = count(Sheet) + (1) This Workbook + (0) Module
		int ncount = xl->ActiveWorkbook->Sheets->Count;
		int total = ncount + 1 + 0;

		if (!(pWb->VBProject->Protection))
		{
			total = pWb->VBProject->VBComponents->GetCount();
		}

		CString szval = L"";
		vector<CString> vecWsh;
		for (int i = 1; i <= total; i++)
		{
			szval = (LPCTSTR)pWb->VBProject->VBComponents->Item(i)->Name;
			if (szval.Left(5) == L"Sheet")
			{
				vecWsh.push_back(szval);
			}
		}

		_variant_t varOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		VARIANT varAfter;
		varAfter.vt = VT_DISPATCH;

		if (nIndex == 0 || nIndex > ncount) nIndex = ncount;
		varAfter.pdispVal = pWb->Sheets->GetItem(nIndex);
		_WorksheetPtr pSheet = pWb->Worksheets->Add(varOption, varAfter, varOption, varOption);

		if (szName != L"") pSheet->PutName((_bstr_t)szName);
		if (szCodename != L"")
		{
			bool blNew = true;
			int nszieWSheet = (int)vecWsh.size();
			for (int i = 1; i <= total + 1; i++)
			{
				try
				{
					szval = (LPCTSTR)pWb->VBProject->VBComponents->Item(i)->Name;
					if (szval.Left(5) == L"Sheet")
					{
						blNew = true;
						for (int j = 0; j < nszieWSheet; j++)
						{
							if (szval == vecWsh[j])
							{
								blNew = false;
								break;
							}
						}

						if (blNew)
						{
							pWb->VBProject->VBComponents->Item(i)->PutName((_bstr_t)szCodename);
							break;
						}
					}
				}
				catch (...) {}
			}
		}

		pSheet->Tab->PutColor(clrTab);
		wPsh = pSheet;

		vecWsh.clear();

		return true;
	}
	catch (...) {}
	return false;
}

int CFunction::_xlGetAllSheet(vector<_WorksheetPtr> &wpSh, int &nVeryHidden)
{
	try
	{
		wpSh.clear();
		nVeryHidden = 0;

		_WorksheetPtr psh;
		int ncount = xl->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= ncount; i++)
		{
			psh = pWb->Worksheets->Item[i];
			if ((int)psh->GetVisible() != 2) wpSh.push_back(psh);
			else nVeryHidden++;
		}
		return (int)wpSh.size();
	}
	catch (...) {}
	return 0;
}

int CFunction::_xlGetIndexSheet(CString szCodename, vector<_WorksheetPtr> &vecPsh)
{
	try
	{
		vecPsh.clear();
		CString szval = L"";
		int nIndex = 0, nHidden = 0;	
		int ncount = _xlGetAllSheet(vecPsh, nHidden);
		for (int i = 0; i < ncount; i++)
		{
			szval = (LPCTSTR)vecPsh[i]->CodeName;
			if (szCodename == szval)
			{
				nIndex = i + nHidden + 1;
				break;
			}
		}

		return nIndex;
	}
	catch (...) {}
	return 0;
}

bool CFunction::_xlCreateNewNameSheet(CString &szName, CString &szCodename, vector<_WorksheetPtr> vecPsh)
{
	try
	{
		int ncount = (int)vecPsh.size();
		
		int dem = 1;
		bool blTrung = false;	
		CString szNew = L"", szval = L"";

		while (true)
		{
			blTrung = false;
			szNew.Format(L"%s.%d", szName, dem);
			for (int i = 0; i < ncount; i++)
			{
				szval = (LPCTSTR)vecPsh[i]->Name;
				if (szNew == szval)
				{
					blTrung = true;
					break;
				}
			}

			if (!blTrung) break;
			dem++;
		}

		szName = szNew;

		/**********************************************************************/

		dem = 1;
		while (true)
		{
			blTrung = false;
			szNew.Format(L"%s_%d", szCodename, dem);
			for (int i = 0; i < ncount; i++)
			{
				szval = (LPCTSTR)vecPsh[i]->CodeName;
				if (szNew == szval)
				{
					blTrung = true;
					break;
				}
			}

			if (!blTrung) break;
			dem++;
		}

		szCodename = szNew;

		return true;
	}
	catch (...) {}
	return false;
}


// Hàm xác định đậm tên font chữ
CString CFunction::_xlGetCellFontName(RangePtr pCell)
{
	try
	{
		return (LPCTSTR)(_bstr_t)pCell->Font->GetBold();
	}
	catch (...)
	{
		try
		{
			CharactersPtr chPtr = pCell->GetCharacters(1, 1);
			return (LPCTSTR)(_bstr_t)chPtr->Font->GetName();
		}
		catch (...) {}
	}

	return L"";
}


// Hàm xác định đậm ô cell
int CFunction::_xlGetCellSize(RangePtr pCell)
{
	try
	{
		return (int)pCell->Font->GetSize();
	}
	catch (...)
	{
		try
		{
			CharactersPtr chPtr = pCell->GetCharacters(1, 1);
			return (int)chPtr->Font->GetSize();
		}
		catch (...) {}
	}

	return 0;
}


// Hàm xác định đậm ô cell
bool CFunction::_xlGetCellBold(RangePtr pCell)
{
	try
	{
		return (bool)pCell->Font->GetBold();
	}
	catch (...) 
	{
		try
		{
			CharactersPtr chPtr = pCell->GetCharacters(1, 1);
			return (bool)chPtr->Font->GetBold();
		}
		catch (...) {}
	}

	return false;
}

// Hàm xác định kiểu nghiêng ô cell
bool CFunction::_xlGetCellItalic(RangePtr pCell)
{
	try
	{
		return (bool)pCell->Font->GetItalic();
	}
	catch (...)
	{
		try
		{
			CharactersPtr chPtr = pCell->GetCharacters(1, 1);
			return (bool)chPtr->Font->GetItalic();
		}
		catch (...) {}
	}

	return false;
}

// Hàm xác định màu chữ ô cell
COLORREF CFunction::_xlGetCellTxtColor(RangePtr pCell)
{
	try
	{
		return (bool)pCell->Font->GetColor();
	}
	catch (...)
	{
		try
		{
			CharactersPtr chPtr = pCell->GetCharacters(1, 1);
			return (COLORREF)chPtr->Font->GetColor();
		}
		catch (...) {}
	}

	return RGB(0, 0, 0);
}


int CFunction::_xlDeleteShape(_WorksheetPtr pSheet, CString sName, bool blEntire)
{
	try
	{
		int dem = 0;
		int ncount = (int)pSheet->Shapes->Count;

		if (sName != L"")
		{
			CString szval = L"";
			int ilen = sName.GetLength();
			for (int i = ncount; i >= 1; i--)
			{
				szval = (LPCTSTR)pSheet->GetShapes()->Item(i)->Name;

				if (blEntire)
				{
					if (szval == sName)
					{
						pSheet->GetShapes()->Item(i)->Delete();
						dem++;
					}
				}
				else
				{
					if (szval.Left(ilen) == sName)
					{
						pSheet->GetShapes()->Item(i)->Delete();
						dem++;
					}
				}
			}
		}
		else
		{
			for (int i = ncount; i >= 1; i--)
			{
				pSheet->GetShapes()->Item(i)->Delete();
			}

			dem = ncount;
		}

		return dem;
	}
	catch (...) {}
	return 0;
}


bool CFunction::_GetImageSize(CString szPathIMG, int *x, int *y)
{
	try
	{
		const char *fn = _ConvertCstringToChars(szPathIMG);
		FILE *f = fopen(fn, "rb");
		if (f == 0) return false;

		fseek(f, 0, SEEK_END);
		long len = ftell(f);
		fseek(f, 0, SEEK_SET);

		if (len < 24)
		{
			fclose(f);
			return false;
		}

		unsigned char buf[24]; fread(buf, 1, 24, f);

		if ((buf[0] == 0xFF && buf[1] == 0xD8 && buf[2] == 0xFF && buf[3] == 0xE0 && buf[6] == 'J' && buf[7] == 'F' && buf[8] == 'I' && buf[9] == 'F') ||
			(buf[0] == 0xFF && buf[1] == 0xD8 && buf[2] == 0xFF && buf[3] == 0xE1 && buf[6] == 'E' && buf[7] == 'x' && buf[8] == 'i' && buf[9] == 'f'))
		{
			long pos = 2;
			while (buf[2] == 0xFF)
			{
				if (buf[3] == 0xC0 || buf[3] == 0xC1 || buf[3] == 0xC2
					|| buf[3] == 0xC3 || buf[3] == 0xC9 || buf[3] == 0xCA || buf[3] == 0xCB) break;

				pos += 2 + (buf[4] << 8) + buf[5];
				if (pos + 12 > len) break;

				fseek(f, pos, SEEK_SET);
				fread(buf + 2, 1, 12, f);
			}
		}

		fclose(f);

		if (buf[0] == 0xFF && buf[1] == 0xD8 && buf[2] == 0xFF)
		{
			*y = (buf[7] << 8) + buf[8];
			*x = (buf[9] << 8) + buf[10];
			return true;
		}

		if (buf[0] == 'G' && buf[1] == 'I' && buf[2] == 'F')
		{
			*x = buf[6] + (buf[7] << 8);
			*y = buf[8] + (buf[9] << 8);
			return true;
		}

		if (buf[0] == 0x89 && buf[1] == 'P' && buf[2] == 'N' && buf[3] == 'G' && buf[4] == 0x0D && buf[5] == 0x0A && buf[6] == 0x1A && buf[7] == 0x0A
			&& buf[12] == 'I' && buf[13] == 'H' && buf[14] == 'D' && buf[15] == 'R')
		{
			*x = (buf[16] << 24) + (buf[17] << 16) + (buf[18] << 8) + (buf[19] << 0);
			*y = (buf[20] << 24) + (buf[21] << 16) + (buf[22] << 8) + (buf[23] << 0);
			return true;
		}

	}
	catch (...) {}
	return false;
}


CString CFunction::_xlCreateNewNameShape(_WorksheetPtr pSheet, CString szPrefix)
{
	try
	{
		CString szNewName = L"";
		szNewName.Format(L"%s%03d", szPrefix, 1);

		int ncount = (int)pSheet->Shapes->Count;
		if (ncount == 0) return szNewName;


		int dem = 0;
		bool bl = false;
		CString szval = L"";

		while (true)
		{
			dem++;
			if (dem == 999) break;

			szNewName.Format(L"%s%03d", szPrefix, dem);
			szNewName.MakeUpper();

			for (int i = 1; i <= ncount; i++)
			{
				bl = false;
				szval = (LPCTSTR)pSheet->GetShapes()->Item(i)->Name;
				szval.MakeUpper();
				if (szval == szNewName)
				{
					bl = true;
					break;
				}
			}

			if (!bl) break;
		}

		return szNewName;
	}
	catch (...) {}
	return L"GXD_ERROR";
}


int CFunction::_xlInsertPictureExcel(_WorksheetPtr pSheet, CString szPathIMG, int iRowBegin, int iColEnd, int numAdd)
{
	try
	{
		int qImageW = 0, qImageH = 0;
		if (_GetImageSize(szPathIMG, &qImageW, &qImageH))
		{
			RangePtr pCell;
			RangePtr pRange = pSheet->Cells;
			float WW = 0, HH = 0, PP = 0, DD = 0;

			for (int i = iColEnd; i <= iColEnd + numAdd; i++)
			{
				pCell = pRange->GetItem(1, i);
				DD = (float)pCell->GetWidth();
				WW += (float)DD;
				PP += (float)(4 * DD / 3);
			}

			if (qImageW > (int)PP)
			{
				qImageH = (int)(PP*qImageH / qImageW);
				qImageW = (int)PP;
			}

			int jHRow = 25;	// Độ cao mặc định 1 ô cell chứa ảnh
			HH = (float)(WW*qImageH / qImageW);

			int pnguyen = (int)(HH / jHRow);
			float pduw = HH - pnguyen * jHRow;
			if (pduw > (float)0) pnguyen++;

			pCell = pSheet->GetRange(pRange->GetItem(iRowBegin + 1, 1), pRange->GetItem(iRowBegin + pnguyen - 1, 1));
			pCell->EntireRow->Insert(xlShiftDown);

			pCell = pSheet->GetRange(pRange->GetItem(iRowBegin, 1), pRange->GetItem(iRowBegin + pnguyen - 1, 1));
			pCell->PutRowHeight(jHRow);

			if (pduw > (float)0)
			{
				pCell = pRange->GetItem(iRowBegin + pnguyen - 1, 1);
				pCell->PutRowHeight(pduw + 2);
			}

			CString szNewName = _xlCreateNewNameShape(pSheet, L"IMG_");
			pCell = pRange->GetItem(iRowBegin, iColEnd);
			ShapePtr ShapeP = pSheet->Shapes->AddPicture(
				(_bstr_t)szPathIMG,	Office::msoTrue, Office::msoTrue, 
				pCell->Left, pCell->Top, WW, HH);
			ShapeP->PutName((_bstr_t)szNewName);
			ShapeP->PutPlacement(xlMoveAndSize);

			return (iRowBegin + pnguyen);
		}
	}
	catch (...) {}

	return 0;
}


bool CFunction::_xlCreateCheckBox(_WorksheetPtr pSheet, int nRow, int nCol,
	CString szNameRec, CString szFormula, CString szOnAction, bool blFont, int nsize)
{
	try
	{
		RangePtr pCell = pSheet->Cells->GetItem(nRow, nCol);
		long xTop = pCell->GetTop();
		long xLeft = pCell->GetLeft();
		long xHeight = pCell->GetRowHeight();

		RangePtr pCellNext = pSheet->Cells->GetItem(nRow, nCol + 1);
		long xWidth = (long)pCellNext->GetLeft() - xLeft;

		ShapePtr ShapeP = pSheet->Shapes->AddShape(
			Office::msoShapeRectangle, xLeft + 0.5, xTop + 0.5, xWidth - 1, xHeight - 1);
		ShapeP->Line->PutVisible(Office::msoFalse);
		ShapeP->Fill->PutVisible(Office::msoFalse);

		if (szNameRec != L"") ShapeP->PutName((_bstr_t)szNameRec);
		if (szFormula != L"") pCell->PutValue2((_bstr_t)szFormula);
		if (szOnAction != L"") ShapeP->PutOnAction((_bstr_t)szOnAction);

		if (blFont)
		{
			try { pCell->Font->PutName((_bstr_t)L"Wingdings 2"); } catch (...) {}
			if (nsize > 0 && nsize < 100) pCell->Font->PutSize(nsize);
		}

		return true;
	}
	catch (...) {}
	return false;
}


bool CFunction::_xlSetNumberFormat(RangePtr pCell, CString szFormat)
{
	try
	{
		pCell->PutNumberFormat((_bstr_t)szFormat);
		return true;
	}
	catch (...) {
		_msgbox(L"Kiểu định dạng không hợp lệ. Vui lòng kiểm tra lại trong "
			L"Time & Language / Region --> 'Regional Format'. "
			L"Ví dụ, bạn có thể đổi lại thành 'English (United States)'.", 0, 0);
	}
	return false;
}


void CFunction::_xlGetDecimalSymbol(CString &szDecimal, CString &szDigitGroup, int nGetRegistry)
{
	try
	{
		if (nGetRegistry == 1)
		{
			// Sử dụng bằng Registry (Phải chạy bằng quyền Admin)
			CRegistry *reg = new CRegistry;
			reg->SetRootKey(HKEY_USERS);
			reg->CreateKey(L".DEFAULT\\Control Panel\\International");
			szDecimal = reg->ReadString(L"sDecimal", L".");
			if (szDecimal == L".") szDigitGroup = L".";
			else szDigitGroup = L",";

			reg->SetRootKey(HKEY_CURRENT_USER);
			delete reg;
		}
		else
		{
			// Sử dụng trên nền Excel
			_bstr_t bsConfig = _xlGetNameSheet(shConfig, false);
			if (bsConfig != (_bstr_t)L"")
			{
				_WorksheetPtr pShConfig = xl->Sheets->GetItem((_bstr_t)bsConfig);
				RangePtr pRgConfig = pShConfig->Cells;

				int iColumnLast = pShConfig->Cells->SpecialCells(xlCellTypeLastCell)->GetColumn();

				CString szval = _xlGIOR(pRgConfig, 1, iColumnLast, L"").Trim();
				if (szval != L"") iColumnLast++;

				RangePtr pCell = pRgConfig->GetItem(1, iColumnLast);
				if (_xlSetNumberFormat(pCell, L"General"))
				{
					pCell->PutValue2((_bstr_t)L"=1/2");

					szval = _xlGIOR(pRgConfig, 1, iColumnLast, L"");
					if (szval.Find(L".") >= 0)
					{
						szDecimal = L".";
						szDigitGroup = L",";
					}
					else
					{
						szDecimal = L",";
						szDigitGroup = L".";
					}

					pCell->ClearContents();
				}
			}
		}
	}
	catch (...) {}
}


CString CFunction::_xlAccountingFormat(
	CString szPositive, CString szNegative, CString szZero, CString szText,
	CString szDecimal, CString szDigitGroup, bool blReplace)
{
	try
	{
		if (blReplace)
		{
			CString szval = L"#" + szDigitGroup + L"#";
			szPositive.Replace(L"#.#", szval);
			szNegative.Replace(L"#.#", szval);
			szZero.Replace(L"#.#", szval);
			szText.Replace(L"#.#", szval);

			szval = L"0" + szDigitGroup + L"0";
			szPositive.Replace(L"0,0", szval);
			szNegative.Replace(L"0,0", szval);
			szZero.Replace(L"0,0", szval);
			szText.Replace(L"0,0", szval);
		}

		CString szFormat = L"";
		if (szPositive != L"") szFormat += (szPositive + L";");
		if (szNegative != L"") szFormat += (szNegative + L";");
		if (szZero != L"") szFormat += (szZero + L";");
		if (szText != L"") szFormat += szText;

		if (szFormat.Right(1) == L";") szFormat = szFormat.Left(szFormat.GetLength() - 1);

		return szFormat;
	}
	catch (...) {}
	return L"";
}


int CFunction::_xlLockSheet(_WorksheetPtr pSheet, CString szPassLock, int iLook)
{
	try
	{
		if (iLook == 1)
		{
			pSheet->PutEnableOutlining(VARIANT_TRUE);
			pSheet->Protect((_bstr_t)szPassLock);
		}
		else
		{
			// Mở khóa
			pSheet->Unprotect((_bstr_t)szPassLock);
		}

		return 1;
	}
	catch (...) {}

	return 0;
}


int CFunction::_xlLockCells(_WorksheetPtr pSheet, RangePtr pCell, CString szPassLock, int iLook)
{
	try
	{
		RangePtr pRange = pSheet->Cells;


		if (iLook == 1)
		{
			// Khóa sheet
			pRange->PutLocked(false);
			pRange->PutFormulaHidden(false);
			pCell->PutLocked(true);

			pSheet->PutEnableOutlining(VARIANT_TRUE);
			pSheet->Protect((_bstr_t)szPassLock);
		}
		else
		{
			// Mở khóa
			pSheet->Unprotect((_bstr_t)szPassLock);
		}

		return 1;
	}
	catch (...) {}

	return 0;
}


/***********************************************************************************************/

int CFunction::_CheckVersionSoftware()
{
	return _wtoi(VERSION_SOFT);
}

char* CFunction::_ConvertCstringToChars(CString szvalue)
{
	wchar_t* wszString = new wchar_t[szvalue.GetLength() + 1];
	wcscpy(wszString, szvalue);
	int num = WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
	char* chr = new char[num + 1];
	WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), chr, num, NULL, NULL);
	chr[num] = '\0';
	delete[] wszString;
	return chr;
}


CString CFunction::_ConvertCharsToCstring(char* chr)
{
	int num = MultiByteToWideChar(CP_UTF8, 0, chr, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[num + 1];
	MultiByteToWideChar(CP_UTF8, 0, chr, -1, wstr, num);
	wstr[num] = '\0';
	CString szval = L"";
	szval.Format(L"%s", wstr);
	delete[] wstr;
	return szval;
}

int CFunction::_TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1, CString symbol2, int itypeFilter, int itypeTrim)
{
	try
	{
		vecPkey.clear();

		if (sKeyValue == L"") return 0;
		if (sKeyValue.Right(1) != symbol1 && sKeyValue.Right(1) != symbol2)
		{
			sKeyValue += symbol1;
		}

		int vtri = 0, icheck = 0;
		CString szval = L"", sktra = L"";
		for (int i = 0; i < sKeyValue.GetLength(); i++)
		{
			szval = sKeyValue.Mid(i, 1);
			if ((szval == symbol1 || szval == symbol2) && i >= vtri)
			{
				icheck = 0;
				if (i > vtri) sktra = sKeyValue.Mid(vtri, i - vtri);
				else sktra = L"";

				if (itypeFilter == 1)
				{
					if ((int)vecPkey.size() > 0)
					{
						for (int j = 0; j < (int)vecPkey.size(); j++)
						{
							if (sktra == vecPkey[j])
							{
								icheck = 1;
								break;
							}
						}
					}
				}

				if (icheck == 0)
				{
					if (itypeTrim == 0) sktra.Trim();
					vecPkey.push_back(sktra);
				}

				vtri = i + 1;
			}
		}

		return (int)vecPkey.size();
	}
	catch (...) {}
	return 0;
}


int CFunction::_FindAllFile(CString pathFolder, CFileFinder &_finder, CString typeOfFile)
{
	try
	{
		CFileFinder::CFindOpts opts;
		opts.sBaseFolder = pathFolder;
		opts.sFileMask = L"*" + typeOfFile + L"*";
		opts.FindNormalFiles();

		_finder.RemoveAll();
		_finder.Find(opts);

		return _finder.GetFileCount();
	}
	catch (...) {}
	return 0;
}

int CFunction::_FindAllFileEx(CString pathFolder, vector<CString> &vecPathFiles)
{
	try
	{
		vecPathFiles.clear();

		if (pathFolder.CompareNoCase(L"") == 0) return 0;
		if (pathFolder.GetLength() < 5) return 0;

		CUtil::GetFileList(_tstring(pathFolder.GetBuffer()), L"*.*", false, vecPathFiles);

		return (int)vecPathFiles.size();
	}
	catch (...) {}
	return 0;
}

// Hàm tìm kiếm các thư mục trong thư mục cha
int CFunction::_FindAllFolder(CString szPathCha, vector<CString> &vecPathFolder)
{
	try
	{
		vecPathFolder.clear();

		if (szPathCha.Right(1) != L"\\") szPathCha += L"\\";
		CString searchString = szPathCha + L"\\*.*";

		bool bIsFolder = true;
		CString szNFolder = L"";
		WIN32_FIND_DATA FindFileData;
		HANDLE hFind = FindFirstFile(searchString, &FindFileData);
		if (hFind != INVALID_HANDLE_VALUE)
		{
			bIsFolder = (FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY;
			if (bIsFolder)
			{
				// Loại bỏ 3 folder (.)  (..) và  ($ recycle.bin)
				szNFolder = FindFileData.cFileName;
				if (szNFolder != L"." && szNFolder != L".." && szNFolder.Find(L"$") == -1)
				{
					vecPathFolder.push_back(szPathCha + szNFolder);
				}
			}

			while (FindNextFile(hFind, &FindFileData))
			{
				bIsFolder = (FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY;
				if (bIsFolder)
				{
					szNFolder = FindFileData.cFileName;
					if (szNFolder != L"." && szNFolder != L".." && szNFolder.Find(L"$") == -1)
					{
						vecPathFolder.push_back(szPathCha + szNFolder);
					}
				}
			}
		}

		FindClose(hFind);
		return (int)vecPathFolder.size();
	}
	catch (...) {}
	return 0;
}


CString CFunction::_GetPathFolder(CString fileName)
{
	try
	{
		HMODULE hModule;
		hModule = GetModuleHandle(fileName);
		TCHAR szFileName[MAX_PATH];
		DWORD dSize = MAX_PATH;
		GetModuleFileName(hModule, szFileName, dSize);
		CString szSource = szFileName;
		for (int i = szSource.GetLength() - 1; i > 0; i--)
		{
			if (szSource.GetAt(i) == '\\')
			{
				szSource = szSource.Left(szSource.GetLength() - (szSource.GetLength() - i) + 1);
				break;
			}
		}
		return szSource;
	}
	catch (...) {}
	return L"";
}


void CFunction::_WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, bool blClear, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (_FileExistsUnicode(spathFile, 0) != 1) return;
		}

		vector<CString> vecLine;
		if (!blClear) _ReadMultiUnicode(spathFile, vecLine);

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(spathFile));
		if (myfile)
		{
			myfile.imbue(loc);

			// Giá trị cũ (nếu có)
			long total = (long)vecLine.size();
			if (total > 0)
			{
				for (long i = 0; i < total; i++)
				{
					wstring ws(vecLine[i]);
					myfile << ws << "\n";
				}
			}

			// Giá trị mới
			total = (long)vecTextLine.size();
			for (long i = 0; i < total; i++)
			{
				wstring ws(vecTextLine[i]);
				myfile << ws << "\n";
			}
			myfile.close();
		}

		vecLine.clear();
	}
	catch (...) {}
}


void CFunction::_WriteMultiUnicode(CString strLine, CString spathFile, bool blClear, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (_FileExistsUnicode(spathFile, 0) != 1) return;
		}

		vector<CString> vecLine;
		if (!blClear) _ReadMultiUnicode(spathFile, vecLine);

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(spathFile));
		if (myfile)
		{
			myfile.imbue(loc);

			// Giá trị cũ (nếu có)
			long total = (long)vecLine.size();
			if (total > 0)
			{
				for (long i = 0; i < total; i++)
				{
					wstring ws(vecLine[i]);
					myfile << ws << "\n";
				}
			}

			// Giá trị mới
			wstring ws(strLine);
			myfile << ws << "\n";

			myfile.close();
		}

		vecLine.clear();
	}
	catch (...) {}
}


int CFunction::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExistsUnicode(spathFile, 0) != 1) return 0;

		wifstream readfile_(_ConvertCstringToChars(spathFile));
		readfile_.imbue(locale(locale::empty(), new codecvt_utf8<wchar_t, 0x10ffff, consume_header>));

		if (readfile_.is_open())
		{
			wstring username;
			while (getline(readfile_, username))
			{
				vecTextLine.push_back(username.c_str());
			}

			readfile_.close();
		}

		return (int)vecTextLine.size();
	}
	catch (...) {}

	return 0;
}


bool CFunction::_IsNumber(CString &szValue)
{
	try
	{
		szValue.Replace(L" ", L"");
		szValue.Replace(L",", L"");
		szValue.Replace(L".", L"");

		int nlen = szValue.GetLength();
		if (nlen == 0) return false;

		CString szCheck = scvtNumber;
		for (int i = 0; i < nlen; i++)
		{
			if (szCheck.Find(szValue.Mid(i, 1)) == -1)
			{
				return false;
			}
		}
	}
	catch (...) {}
	return true;
}

bool CFunction::_IsNumberLaMa(CString szNumber)
{
	int nLen = szNumber.GetLength();
	if (nLen == 0) return false;

	szNumber.MakeUpper();
	CString Lama = scvtLaMa;
	for (int i = 0; i < nLen; i++)
	{
		if (Lama.Find(szNumber.Mid(i, 1)) == -1)
		{
			return false;
		}
	}

	return true;
}

bool CFunction::_IsNumberLaMaCap2(CString szNumber)
{
	int pos = szNumber.Find(L".");
	if (pos > 0)
	{
		szNumber = szNumber.Left(pos);
		if (_IsNumberLaMa(szNumber))
		{
			return true;
		}
	}

	return false;
}

int CFunction::_ConvertLaMaToNumber(CString szLaMa)
{
	int nsize = getSizeArray(solama);
	for (int i = 0; i < nsize; i++)
	{
		if (szLaMa == solama[i]) return i + 1;
	}
	return 0;
}


CString CFunction::_ConvertNumberToLaMa(int number)
{
	int nsize = getSizeArray(solama);
	for (int i = 0; i < nsize; i++)
	{
		if (number == i + 1) return solama[i];
	}
	return L"";
}


CString CFunction::_ConvertUpper(CString szNoidung)
{
	if (szNoidung != L"")
	{
		szNoidung.MakeUpper();
		int len = scvtUpper.GetLength();
		for (int i = 0; i < len; i++)
		{
			szNoidung.Replace(scvtLower.Mid(i, 1), scvtUpper.Mid(i, 1));
		}
	}	

	return szNoidung;
}


CString CFunction::_ConvertLower(CString szNoidung)
{
	if (szNoidung != L"")
	{
		szNoidung.MakeLower();
		int len = scvtLower.GetLength();
		for (int i = 0; i < len; i++)
		{
			szNoidung.Replace(scvtUpper.Mid(i, 1), scvtLower.Mid(i, 1));
		}
	}	

	return szNoidung;
}


CString CFunction::_ConvertUPProper(CString szNoidung)
{
	try
	{
		if (szNoidung != L"")
		{
			szNoidung.MakeLower();
			int len = scvtLower.GetLength();
			for (int i = 0; i < len; i++)
			{
				szNoidung.Replace(scvtUpper.Mid(i, 1), scvtLower.Mid(i, 1));
			}

			CString szABC = scvtABC;
			szABC.MakeLower();
			szABC += scvtLower;
			szABC += scvtNumber;

			bool flag = false;			
			int total = szNoidung.GetLength();
			CString szval = L"", szReturn = szNoidung;

			for (int i = 0; i < total; i++)
			{
				szval = szNoidung.Mid(i, 1);
				if (i == 0 || szval == L"." || szval == L";") flag = true;

				if (flag)
				{
					if (szABC.Find(szval) >= 0)
					{
						szval.MakeUpper();
						for (int j = 0; j < len; j++)
						{
							if (szval == scvtLower.Mid(j, 1))
							{
								szval = scvtUpper.Mid(j, 1);
								break;
							}
						}

						szReturn = szNoidung.Left(i) + szval + szNoidung.Right(total - i - 1);
						szNoidung = szReturn;
						flag = false;
					}
				}
			}

			return szReturn;
		}
	}
	catch (...) {}

	return szNoidung;
}


CString CFunction::_ReplaceMatchCase(CString szNoidung, CString szOld, CString szNew, bool blCase, bool blKDau)
{
	try
	{
		if (!blCase && !blKDau)
		{
			szNoidung.Replace(szOld, szNew);
			return szNoidung;
		}

		CString szLower = szNoidung;
		if (blKDau)
		{
			szLower = _ConvertKhongDau(szLower);
			szOld = _ConvertKhongDau(szOld);

			if (blCase)
			{
				szLower.MakeLower();
				szOld.MakeLower();
			}
		}
		else
		{
			szLower = _ConvertLower(szLower);
			szOld = _ConvertLower(szOld);
		}

		int nTotal = 0;
		int nLen = szOld.GetLength();

		int pos = -1;
		CString szval = L"";
		vector<CString> vecNdung;

		while (true)
		{
			pos = szLower.Find(szOld);
			if (pos == -1) break;

			nTotal = szLower.GetLength();
			szval = szNoidung.Left(pos);
			vecNdung.push_back(szval);

			pos += nLen;
			if (pos < nTotal - 1)
			{
				szLower = szLower.Right(nTotal - pos);
				szNoidung = szNoidung.Right(nTotal - pos);
			}
			else
			{
				szNoidung = L"";
				break;
			}
		}

		szval = L"";
		int nsize = (int)vecNdung.size();
		if (nsize > 0)
		{
			for (int i = 0; i < nsize; i++)
			{
				szval += vecNdung[i];
				szval += szNew;
			}
		}
		szval += szNoidung;

		vecNdung.clear();

		return szval;
	}
	catch (...) {}
	return szNoidung;
}

CString CFunction::_ConvertNumberToString(int number, int iType)
{
	try
	{
		CString szNum = L"";
		if (iType < 0) iType = 0;
		if (iType > 6) iType = 6;

		if (iType == 0 || iType == 1) szNum.Format(L"%d", number);
		else
		{
			switch (iType)
			{
				case 2: szNum.Format(L"%02d", number); break;
				case 3: szNum.Format(L"%03d", number); break;
				case 4: szNum.Format(L"%04d", number); break;
				case 5: szNum.Format(L"%05d", number); break;
				case 6: szNum.Format(L"%06d", number); break;
			}
		}

		return szNum;
	}
	catch (...) {}
	return L"";
}


CString CFunction::_ConvertKhongDau(CString szNoidung, int iLower)
{
	if (szNoidung != L"")
	{
		int len = scvtUTF.GetLength();
		for (int i = 0; i < len; i++)
		{
			szNoidung.Replace(scvtUTF.Mid(i, 1), scvtKOD.Mid(i, 1));
		}

		if (iLower == 1) szNoidung.MakeLower();
		else if (iLower == 2) szNoidung.MakeUpper();
	}

	return szNoidung;
}


CString CFunction::_ConvertBytes(double ivalue)
{
	CString szval = L"";
	CString cs[4] = { L"KB", L"MB", L"GB", L"TB" };
	for (int i = 0; i < 4; i++)
	{
		ivalue = ivalue / 1024;
		if (ivalue < 1024)
		{
			ivalue = _RoundDouble(ivalue, 2);
			szval.Format(L"%4.2f %s", ivalue, cs[i]);
			return szval;
		}
	}

	szval.Format(L"%f bytes", ivalue);
	return szval;
}

void CFunction::_GetFileAttributesEx(CString szDir, CString &szKichthuoc, CString &szNgay)
{
	try
	{
		szKichthuoc = L"";
		szNgay = L"";

		HMODULE hLib;
		GET_FILE_ATTRIBUTES_EX func;
		FILE_INFO fInfo;

		hLib = LoadLibrary(L"Kernel32.dll");
		if (hLib != NULL)
		{
			func = (GET_FILE_ATTRIBUTES_EX)GetProcAddress(hLib, "GetFileAttributesExW");
			if (func != NULL) func(szDir, 0, &fInfo);

			SYSTEMTIME times, stLocal;
			FileTimeToSystemTime(&fInfo.ftLastWriteTime, &times);
			SystemTimeToTzSpecificLocalTime(NULL, &times, &stLocal);

			szKichthuoc = _ConvertBytes((double)fInfo.nFileSizeLow);

			if ((int)stLocal.wYear >= 2000)
			{
				szNgay.Format(L"%02d/%02d/%04d %02d:%02d:%02d",
					stLocal.wDay, stLocal.wMonth, stLocal.wYear,
					stLocal.wHour, stLocal.wMinute, stLocal.wSecond);
			}

			FreeLibrary(hLib);
		}
	}
	catch (...) {}
}


CString CFunction::_NameColumnHeading(CListCtrl &clist, int column, int itype, CString szName)
{
	try
	{
		CString strNome = L"";
		if (itype == 1) strNome = szName;

		CHeaderCtrl *pHeader = clist.GetHeaderCtrl();
		HDITEM hdi;
		hdi.mask = HDI_TEXT;
		hdi.pszText = strNome.GetBuffer(256);
		hdi.cchTextMax = 256;

		if (itype == 0) pHeader->GetItem(column, &hdi);
		else pHeader->SetItem(column, &hdi);
		strNome.ReleaseBuffer();

		return strNome;
	}
	catch (...) {}
	return L"";
}


bool CFunction::_Is64BitWindows()
{
	#if defined(_WIN64)
		return TRUE;  // 64-bit programs run only on Win64
	#elif defined(_WIN32)
		// 32-bit programs run on both 32-bit and 64-bit Windows
		// so must sniff
		BOOL f64 = FALSE;
		return IsWow64Process(GetCurrentProcess(), &f64) && f64;
	#else
		return FALSE; // Win64 does not support Win16
	#endif
}


bool CFunction::_IsProcessRunning(const wchar_t *processName)
{
	bool exists = false;
	PROCESSENTRY32 entry;
	entry.dwSize = sizeof(PROCESSENTRY32);

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);

	if (Process32First(snapshot, &entry))
		while (Process32Next(snapshot, &entry))
			if (!wcsicmp(entry.szExeFile, processName))
				exists = true;

	CloseHandle(snapshot);
	return exists;
}


int CFunction::_FileExistsUnicode(CString pathFile, int itype)
{
	if (pathFile == L"") return 0;

	CString str = L"Đường dẫn file không hợp lệ. Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		LPTSTR lpPath = pathFile.GetBuffer(MAX_PATH);
		if (PathFileExists(lpPath)) return 1;
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}
	catch (...)
	{
		if (itype == 1) _msgbox(str, 2);
		return -1;
	}

	return 0;
}


bool CFunction::_FolderExistsUnicode(CString dirName_in)
{
	DWORD dwAttrib = GetFileAttributes((LPCTSTR)dirName_in);
	return (dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}


COLORREF CFunction::_SetColorRGB(CString szColor, CString symbol, int itype)
{
	szColor.Replace(L" ", L"");
	szColor.Replace(L")", L"");
	szColor.Replace(L"RGB(", L"");	
	szColor.Replace(L".", symbol);
	szColor.Replace(L";", symbol);

	if (szColor == L"")
	{
		if (itype == 0) return RGB(255, 255, 255);
		else return RGB(0, 0, 0);
	}

	int pos = szColor.Find(symbol);
	int iRed = _wtoi(szColor.Left(pos));

	CString szval = szColor.Right(szColor.GetLength() - pos - 1);
	pos = szval.Find(symbol);
	int iGreen = _wtoi(szval.Left(pos));
	int iBlue = _wtoi(szval.Right(szval.GetLength() - pos - 1));

	BYTE bRed = iRed;
	BYTE bGreen = iGreen;
	BYTE bBlue = iBlue;

	return RGB(bRed, bGreen, bBlue);
}

CString CFunction::_GetColorRGB(COLORREF clr, CString symbol, int itype)
{
	CString szval = L"";
	szval.Format(L"%d%s%d%s%d", 
		(int)GetRValue(clr), symbol, 
		(int)GetGValue(clr), symbol, 
		(int)GetBValue(clr));

	if (itype == 1) return (L"RGB(" + szval + L")");

	return szval;
}

bool CFunction::_GetColorDlg(COLORREF &clr, CString &szColor)
{
	CColorDialog dlg;
	if (dlg.DoModal() == IDOK)
	{
		COLORREF clrRGB = dlg.GetColor();
		if (clrRGB != NULL)
		{
			clr = clrRGB;
			szColor = _GetColorRGB(clr, L",", 0);
			
			return true;
		}
	}

	return false;
}

bool CFunction::_GetValueThongso(CString szThongso, int nIndex, CString &szReturn)
{
	try
	{
		szThongso.Trim();
		if (szThongso == L"") return false;

		vector<CString> vecKey;
		int iKey = _TackTukhoa(vecKey, szThongso, L"|", L";", 0, 0);

		if (iKey >= nIndex)
		{
			szReturn = vecKey[nIndex];
			vecKey.clear();
			return true;
		}
		
		vecKey.clear();
	}
	catch (...) {}
	return false;
}


int CFunction::_IsTaskbarWndVisible(int &iTaskHeight)
{
	// nPostion = -1 --> Thanh Taskbar bị ẩn
	// nPostion =  0 - Bottom | 1 - Left | 2 - Top | 3 - Right
	int nPostion = -1;
	HWND hTaskbarWnd = ::FindWindow(L"Shell_TrayWnd", NULL);
	HMONITOR hMonitor = MonitorFromWindow(hTaskbarWnd, MONITOR_DEFAULTTONEAREST);
	MONITORINFO info = { sizeof(MONITORINFO) };
	if (GetMonitorInfo(hMonitor, &info))
	{
		RECT rect;
		GetWindowRect(hTaskbarWnd, &rect);

		if (!((rect.top >= info.rcMonitor.bottom - 4) ||
			(rect.right <= 2) ||
			(rect.bottom <= 4) ||
			(rect.left >= info.rcMonitor.right - 2)))
		{
			if ((int)rect.left > 0)
			{
				// Task --> Pane Right
				nPostion = 3;
				iTaskHeight = (int)rect.left;
			}
			else
			{
				if ((int)rect.top > 0)
				{
					// Task --> Pane Bottom
					nPostion = 0;
					iTaskHeight = (int)rect.top;
				}
				else
				{
					if ((int)rect.right == 1366)
					{
						// Task --> Pane Top
						nPostion = 2;
						iTaskHeight = (int)rect.bottom;
					}
					else
					{
						// Task --> Pane Left
						nPostion = 1;
						iTaskHeight = (int)rect.right;
					}
				}
			}
		}
	}
	return nPostion;
}


HWND CFunction::_CreateToolTip(int toolID, HWND hDlg, CString pszText)
{
	if (!toolID || !hDlg || !pszText) return FALSE;
	HWND hwndTool = GetDlgItem(hDlg, toolID);
	HWND hwndTip = CreateWindowExW(NULL, TOOLTIPS_CLASSW, NULL,
		WS_POPUP | TTS_ALWAYSTIP | TTS_BALLOON,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, hDlg, NULL, NULL, NULL);

	if (!hwndTool || !hwndTip) return (HWND)NULL;
	TOOLINFOW toolInfo = { 0 };
	toolInfo.cbSize = sizeof(toolInfo);
	toolInfo.hwnd = hDlg;
	toolInfo.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
	toolInfo.uId = (UINT_PTR)hwndTool;
	toolInfo.lpszText = (LPWSTR)(_bstr_t)pszText;

	SendMessage(hwndTip, TTM_ADDTOOLW, 0, (LPARAM)(LPTOOLINFOW)&toolInfo);
	return hwndTip;
}

CString CFunction::_FormatTienVND(CString szTien, CString szSteparator, CString szDecimal)
{
	if (szTien == L"") return L"";
	CString snguyen = szTien;
	CString stphan = L"";
	int pos = szTien.Find(szDecimal);
	if (pos > 0)
	{
		snguyen = szTien.Left(pos);
		stphan = szTien.Right(szTien.GetLength() - pos - 1);
	}

	CString szval = L"";
	int len = snguyen.GetLength();
	int pnguyen = (int)(len / 3);
	int ple = len % 3;

	if (ple > 0) szval = snguyen.Left(ple);
	if (pnguyen >= 1)
	{
		if (ple > 0 && szval != L"-") szval += szSteparator;
		for (int i = 0; i < pnguyen; i++)
		{
			szval += snguyen.Mid(ple + 3 * i, 3);
			if (i < pnguyen - 1) szval += szSteparator;
		}
	}

	if (szval == L"") szval = snguyen;
	if (stphan != L"")
	{
		szval += szDecimal;
		szval += stphan;
	}

	return szval;
}


double CFunction::_RoundDouble(double doValue, int nPrecision)
{
	static const double doBase = 10.0;
	double doComplete5, doComplete5i;

	doComplete5 = doValue * pow(doBase, (double)(nPrecision + 1));

	if (doValue < 0.0) doComplete5 -= 5.0;
	else doComplete5 += 5.0;

	doComplete5 /= doBase;
	modf(doComplete5, &doComplete5i);

	return doComplete5i / pow(doBase, (double)nPrecision);
}

CString CFunction::_GetDesktopDir()
{
	CString szDesktop = L"";
	char* szpath = new char[MAX_PATH + 1];
	if (SHGetSpecialFolderPathA(HWND_DESKTOP, szpath, CSIDL_DESKTOP, FALSE))
	{
		szDesktop = _ConvertCharsToCstring(szpath);
		if (szDesktop.Right(1) != L"\\") szDesktop += L"\\";
	}
	return szDesktop;
}


int CFunction::_RandomTime()
{
	time_t now = time(0);
	tm *ltm = localtime(&now);
	return (100 * (int)ltm->tm_min + (int)ltm->tm_sec + rand() % 1000);
}

CString CFunction::_GetTimeNow(int ifulltime)
{
	try
	{
		time_t currentTime;
		time(&currentTime);
		struct tm *localTime = localtime(&currentTime);
		int iYear = (int)localTime->tm_year;
		if (iYear < 1900) iYear += 1900;

		CString szDaynow = L"";
		szDaynow.Format(L"%02d/%02d/%d",
			(int)localTime->tm_mday, 1 + (int)localTime->tm_mon, iYear);

		if (ifulltime == 1)
		{
			CString sztimer = L"";
			sztimer.Format(L" %02d:%02d:%02d",
				(int)localTime->tm_hour, (int)localTime->tm_min, (int)localTime->tm_sec);
			szDaynow += sztimer;
		}
		return szDaynow;
	}
	catch (...) {}
	return L"";
}


CString CFunction::_FormatDate(CString szDate)
{
	try
	{
		if (szDate.GetLength() <= 10)
		{
			vector<CString> vecKey;
			CString szFormat = szDate;
			if (_TackTukhoa(vecKey, szDate, L"/", L"-", 0, 0) == 3)
			{
				// Format = dd/MM/yyyy
				szFormat.Format(L"%02d/%02d/%d",
					_wtoi(vecKey[0]), _wtoi(vecKey[1]), 2000 + _wtoi(vecKey[2].Right(2)));
			}

			vecKey.clear();
			return szFormat;
		}
	}
	catch (...) {}
	return szDate;
}


int CFunction::_CompareDate(CString date1, CString date2)
{
	try
	{
		date1 = _FormatDate(date1);
		date2 = _FormatDate(date2);

		int d1 = _wtoi(date1.Left(2));
		int d2 = _wtoi(date2.Left(2));

		int m1 = _wtoi(date1.Mid(3, 2));
		int m2 = _wtoi(date2.Mid(3, 2));

		int y1 = _wtoi(date1.Right(2));
		int y2 = _wtoi(date2.Right(2));

		if (y1 < y2 || (y1 == y2 && m1 < m2) || (y1 == y2 && m1 == m2 && d1 <= d2)) return 1;
		return 0;
	}
	catch (...) {}
	return 0;
}


CString	CFunction::_DayNextPrev(CString szDate, int num)
{
	try
	{
		szDate = _FormatDate(szDate);

		if (num == 0) return szDate;

		CString szDay = szDate;
		int dd = _wtoi(szDay.Left(2));
		int mm = _wtoi(szDay.Mid(3, 2));
		int yy = 2000 + _wtoi(szDay.Right(2));
		int ddmax = 31;

		if (num > 0)
		{
			while (true)
			{
				ddmax = 31;
				if (mm == 2)
				{
					if (yy % 400 == 0 || (yy % 4 == 0 && yy % 100 != 0)) ddmax = 29;
					else ddmax = 28;
				}
				else if (mm == 4 || mm == 6 || mm == 9 || mm == 11) ddmax = 30;

				if (num + dd > ddmax)
				{
					num -= (ddmax - dd + 1);

					dd = 1;
					if (mm < 12) mm++;
					else
					{
						mm = 1;
						yy++;
					}
				}
				else
				{
					dd += num;
					break;
				}
			}
		}
		else
		{
			num = abs(num);
			while (true)
			{
				if (dd - num < 1)
				{
					mm--;
					if (mm <= 0)
					{
						mm = 12;
						yy--;
					}

					ddmax = 31;
					if (mm == 2)
					{
						if (yy % 400 == 0 || (yy % 4 == 0 && yy % 100 != 0)) ddmax = 29;
						else ddmax = 28;
					}
					else if (mm == 4 || mm == 6 || mm == 9 || mm == 11) ddmax = 30;

					num -= dd;
					dd = ddmax;
				}
				else
				{
					dd -= num;
					break;
				}
			}
		}

		szDay.Format(L"%02d/%02d/%d", dd, mm, yy);
		return szDay;
	}
	catch (...) {}
	return szDate;
}


int CFunction::_DayValue(int day, int month, int year)
{
	if (month < 3)
	{
		year--;
		month += 12;
	}
	return 365 * year + year / 4 - year / 100 + year / 400 + (153 * month - 457) / 5 + day - 306;
}


int CFunction::_NumDay2(CString dateMax, CString dateMin, int cong)
{
	int d1 = _wtoi(dateMax.Left(2));
	int m1 = _wtoi(dateMax.Mid(3, 2));
	int y1 = 2000 + _wtoi(dateMax.Right(2));

	int d2 = _wtoi(dateMin.Left(2));
	int m2 = _wtoi(dateMin.Mid(3, 2));
	int y2 = 2000 + _wtoi(dateMin.Right(2));

	int numberOfDays = _DayValue(y1, m1, d1) - _DayValue(y2, m2, d2);
	return numberOfDays + cong;
}


CString CFunction::_IsDayOfTheWeek(CString szDate)
{
	try
	{
		if (szDate == L"") return L"";

		szDate = _FormatDate(szDate);

		int dd = _wtoi(szDate.Left(2));
		int mm = _wtoi(szDate.Mid(3, 2));
		int yy = 2000 + _wtoi(szDate.Right(2));

		int nIndex = (dd + ((153 * (mm + 12 * ((14 - mm) / 12) - 3) + 2) / 5) +
			(365 * (yy + 4800 - ((14 - mm) / 12))) +
			((yy + 4800 - ((14 - mm) / 12)) / 4) -
			((yy + 4800 - ((14 - mm) / 12)) / 100) +
			((yy + 4800 - ((14 - mm) / 12)) / 400) - 32045) % 7;

		CString cs[7] = { L"T2", L"T3", L"T4", L"T5", L"T6", L"T7", L"CN" };
		if (nIndex >= 0 && nIndex < 7) return cs[nIndex];
	}
	catch (...) {}

	return L"";
}


void CFunction::_ShellExecuteSelectFile(CString szDirFileSelect)
{
	SHELLEXECUTEINFO shExecInfo = { 0 };
	shExecInfo.cbSize = sizeof(shExecInfo);
	shExecInfo.lpFile = _T("Explorer.exe");
	shExecInfo.lpParameters = L"/Select, " + szDirFileSelect;
	shExecInfo.nShow = SW_SHOWNORMAL;
	shExecInfo.lpVerb = _T("Open");
	shExecInfo.fMask = SEE_MASK_INVOKEIDLIST | SEE_MASK_FLAG_DDEWAIT | SEE_MASK_FLAG_NO_UI;
	VERIFY(ShellExecuteEx(&shExecInfo));
}


BOOL CFunction::_CreateNewFolder(CString directoryPath)
{
	CString directoryPathClone = directoryPath;
	LPTSTR lpPath = directoryPath.GetBuffer(MAX_PATH);
	if (!PathFileExists(lpPath))
	{
		if (!CreateDirectory(lpPath, NULL) && (GetLastError() == ERROR_PATH_NOT_FOUND))
		{
			if (PathRemoveFileSpec(lpPath))
			{
				CString newPath = lpPath;
				_CreateNewFolder(newPath);
			}
		}
	}

	directoryPath.ReleaseBuffer(-1);
	return PathFileExists(directoryPathClone);
}


bool CFunction::_CallFunctionDLL(CString szDllName, CString szFunction, bool blQuit)
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
		return true;
	}
	catch (...) {}
	return false;
}

bool CFunction::_CallModuleXLA(CString szFileName, CString szModule, CString szSub)
{
	try
	{
		_variant_t varUnused((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		CString PATH_MODULEXLA = L"";
		PATH_MODULEXLA.Format(L"%s!%s.%s", szFileName, szModule, szSub);
		_variant_t szMacro(PATH_MODULEXLA);

		xl->Run(szMacro,
			varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
			varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
			varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused, varUnused,
			varUnused, varUnused, varUnused, varUnused, varUnused, varUnused);

		return true;
	}
	catch (...) {}
	return false;
}


bool CFunction::_CreateDirectories(CString pathName)
{
	try
	{
		if (pathName.Right(1) != L"\\") pathName += L"\\";
		for (int i = 0; i < pathName.GetLength(); i++)
		{
			if (pathName.Mid(i, 1) == L"\\") _CreateNewFolder(pathName.Left(i));
		}
		return true;
	}
	catch (...) {}
	return false;
}


void CFunction::_LogFileWrite(CString szLog, bool blClear, bool blOpenFile, CString szFileName)
{
	try
	{
		CString szpathFileLog = _GetDesktopDir() + szFileName;

		vector<CString> vecLine;
		if (!blClear) _ReadMultiUnicode(szpathFileLog, vecLine);

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(szpathFileLog));
		if (myfile)
		{
			myfile.imbue(loc);

			int itotal = vecLine.size();
			if (itotal > 0)
			{
				for (int i = 0; i < itotal; i++)
				{
					wstring ws(vecLine[i]);
					myfile << ws << "\n";
				}
			}

			if (szLog != "<!>")
			{
				wstring ws(szLog);
				myfile << ws << "\n";
			}

			myfile.close();
		}

		vecLine.clear();

		if (blOpenFile) ShellExecute(NULL, L"open", szpathFileLog, NULL, NULL, SW_SHOWMAXIMIZED);
	}
	catch (...) {}
}

void CFunction::_LogFileWrite(int iLog, bool blClear, bool blOpenFile, CString szFileName)
{
	CString str = _ConvertNumberToString(iLog);
	_LogFileWrite(str, blClear, blOpenFile, szFileName);
}

void CFunction::_LogFileStart(CString szStart) { _LogFileWrite(szStart, true); }
void CFunction::_LogFileEnd() { _LogFileWrite(L"<!>", false, true); }

CString CFunction::_GetTypeOfFile(CString pathFile)
{
	try
	{
		if (pathFile != L"")
		{
			if ((int)pathFile.Find(L".") > 0)
			{
				CString szval = L"";
				int nLen = pathFile.GetLength();
				for (int i = nLen - 1; i >= 0; i--)
				{
					szval = pathFile.Mid(i, 1);
					if (szval == L".") return pathFile.Right(nLen - i - 1);
					else if (szval == L"\\" || szval == L"/") return L"";
				}
			}
		}
	}
	catch (...) {}
	return L"";
}

CString CFunction::_GetNameOfPath(CString pathFile, int &pos, int ipath)
{
	try
	{
		pos = -1;
		if (pathFile != L"")
		{
			if (pathFile.Right(1) == L"\\") pathFile = pathFile.Left(pathFile.GetLength() - 1);

			int nLen = pathFile.GetLength();
			for (int i = nLen - 1; i >= 0; i--)
			{
				if (pathFile.Mid(i, 1) == L"\\" || pathFile.Mid(i, 1) == L"/")
				{
					pos = i;	// <-- Trả về vị trí phân tách

					if (ipath == -1)
					{
						// Trả về kết quả là tên file
						CString szfile = pathFile.Right(nLen - i - 1);
						int nLenName = szfile.GetLength();
						for (int j = nLenName - 1; j >= 0; j--)
						{
							if (szfile.Mid(j, 1) == L".")
							{
								return szfile.Left(j);
							}
						}
						return szfile;
					}
					else if (ipath == 0)
					{
						// Trả về kết quả là tên file + đuôi
						return pathFile.Right(nLen - i - 1);
					}

					// Trả về đường dẫn chứa file
					return pathFile.Left(i);
				}
			}
		}
	}
	catch (...) {}
	return L"";
}


void CFunction::_CopyClipboard(CString szNoidung)
{
	try
	{
		if (!OpenClipboard(NULL)) return;
		if (!EmptyClipboard()) return;

		size_t size = sizeof(TCHAR)*(1 + szNoidung.GetLength());
		HGLOBAL hResult = GlobalAlloc(GMEM_MOVEABLE, size);
		LPTSTR lptstrCopy = (LPTSTR)GlobalLock(hResult);
		memcpy(lptstrCopy, szNoidung.GetBuffer(), size);
		GlobalUnlock(hResult);
#ifndef _UNICODE
		if (::SetClipboardData(CF_TEXT, hResult) == NULL)
#else
		if (::SetClipboardData(CF_UNICODETEXT, hResult) == NULL)
#endif
		{
			GlobalFree(hResult);
			CloseClipboard();
			return;
		}
		CloseClipboard();
	}
	catch (...) {}
}

CString CFunction::_PasteClipboard()
{
	try
	{
		if (!OpenClipboard(NULL)) return L"";

		HANDLE hData;
		CString strData = L"";

#ifndef _UNICODE
		hData = ::GetClipboardData(CF_TEXT);
#else
		hData = ::GetClipboardData(CF_UNICODETEXT);
#endif

		if (hData)
		{
			WCHAR *pchData = (WCHAR*)GlobalLock(hData);
			if (pchData)
			{
				strData = pchData;
				GlobalUnlock(hData);
			}
		}

		CloseClipboard();

		return strData;
	}
	catch (...) {}

	return L"";
}


void CFunction::_QuickSortArr(int *arr, int l, int r)
{
	int b, i = l, j = r, x = arr[(l + r) / 2];
	do
	{
		while (arr[i] < x) i++;
		while (arr[j] > x) j--;
		if (i <= j)
		{
			b = arr[i];
			arr[i] = arr[j];
			arr[j] = b;
			i++;
			j--;
		}
	} while (i < j);

	if (l < j) _QuickSortArr(arr, l, j);
	if (r > i) _QuickSortArr(arr, i, r);
}