#include "pch.h"
#include "Function.h"

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

Function::Function()
{
}

Function::~Function()
{
}

void Function::_msgbox(CString message, int itype, int iexcel, int itimeout)
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

void Function::_s(CString message, int itype, int iexcel, int itimeout)
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

void Function::_d(int num, int itype, int iexcel, int itimeout)
{
	CString szval = L"";
	szval.Format(L"Num= %d", num);
	_s(szval, itype, iexcel, itimeout);
}

void Function::_n(CString message, int num, int itype, int istr, int iexcel, int itimeout)
{
	message.Trim();
	if (message.Right(1) == L"=") message = message.Left(message.GetLength() - 1);

	CString szval = L"";
	if (istr == 0) szval.Format(L"%s= %d", message, num);
	else szval.Format(L"%d= %s", num, message);

	_s(szval, itype, iexcel, itimeout);
}


int Function::_y(CString message, int istyle, int itype, int iexcel, int itimeout)
{
	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_TRUE);

	UINT qs[2] = { MB_YESNO, MB_YESNOCANCEL };
	UINT style[2] = { MB_ICONQUESTION, MB_ICONWARNING };

	return (int)MessageBoxTimeout(NULL, message, MS_THONGBAO,
		qs[istyle] | style[itype] | MB_DEFBUTTON1, 0, itimeout);

	if (iexcel == 1) xl->PutScreenUpdating(VARIANT_FALSE);

	return 6;
}

void Function::_msgupdate()
{
	_msgbox(L"Tính năng đang phát triển. Vui lòng đợi ở phiên bản cập nhật tiếp theo.", 0, 0);
}

/***********************************************************************************************/

void Function::_xlCreateExcel(int itype)
{
	CoInitialize(NULL);
	xl.GetActiveObject(ExcelApp);
	xl->PutVisible(VARIANT_TRUE);
	if (itype == 1) xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	pWb = xl->GetActiveWorkbook();
}

void Function::_xlCloseExcel()
{
	CoUninitialize();
}

void Function::_xlGetInfoSheetActive()
{
	pShActive = pWb->ActiveSheet;
	pRgActive = pShActive->Cells;
	sCodeActive = (LPCTSTR)pShActive->CodeName;
	sNameActive = (LPCTSTR)pShActive->Name;
	bkgColorActive = pShActive->Tab->GetColor();

	iRowActive = pShActive->Application->ActiveCell->GetRow();
	iColumnActive = pShActive->Application->ActiveCell->GetColumn();
}

bool Function::_xlIsOperatingSystem64()
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

	//HRESULT hr = objExcel.GetActiveObject(ExcelApp);
	if (!(/*FAILED(hr) && */objExcel == NULL))
	{
		CString szOperatingSystem = (LPCTSTR)objExcel->GetOperatingSystem();
		if (szOperatingSystem.Find(L"64") >= 0) return true;
		else return false;
	}

	CoUninitialize();
	return false;
}

CString Function::_xlGetVersionExcel()
{
	CString szVersion = (LPCTSTR)xl->GetVersion();
	return szVersion;
}

void Function::_xlDeveloperMacroSettings()
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

void Function::_xlPutScreenUpdating(bool bl)
{
	if (bl == false) xl->PutScreenUpdating(VARIANT_FALSE);
	else xl->PutScreenUpdating(VARIANT_TRUE);
}

void Function::_xlPutStatus(CString szStatus)
{
	xl->PutStatusBar((_bstr_t)szStatus);
}

_bstr_t Function::_xlGetNameSheet(_bstr_t codename, int itype)
{
	try
	{
		_WorksheetPtr psh;
		int ncount = xl->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= ncount; i++)
		{
			psh = pWb->Worksheets->Item[i];
			if (i == 1)
			{
				if ((int)psh->GetVisible() == 2)
				{
					if (!bCheckVirut)
					{
						if (psh->CodeName != (_bstr_t)shConfig)
						{
							bCheckVirut = !bCheckVirut;
							_msgbox(L"Phát hiện virut NEGS.xls trong máy tính bạn. "
								L"Virut có thể gây ảnh hưởng đến quá trình sử dụng. "
								L"Vui lòng kiểm tra lại hoặc liên hệ với "
								L"tổng đài hỗ trợ GXD 1900.0147 để được xử lý.",
								1, 1);
						}						
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
		if (itype == 1) _msgbox(str, 2);
	}
	return L"";
}

int Function::_xlFindValue(_WorksheetPtr pSheet, int column, int rowbd, int rowkt,
	_bstr_t bvalue, bool match_content, bool match_case, int ishow, int itype)
{
	try
	{
		if (rowbd <= 0) rowbd = 1;
		if (rowkt <= 0) rowkt = 1000000;

		bool blHidden = false;
		CString szval = L"A1:AZ1000000";
		RangePtr pRange = pSheet->Cells;
		RangePtr PRShow = pRange->GetItem(1, column);

		if (ishow == 1)
		{
			blHidden = PRShow->EntireColumn->GetHidden();
			if (blHidden == true) PRShow->EntireColumn->PutHidden(false);
		}
		if (column > 0) szval.Format(L"%s:%s", _xlGAR(pRange, rowbd, column, 0), _xlGAR(pRange, rowkt, column, 0));

		_variant_t vtCon = XlLookAt::xlPart;
		if (match_content == true) vtCon = XlLookAt::xlWhole;

		RangePtr PRS;
		int iRowValue = 0;
		if (PRS = pRange->GetRange(COleVariant(szval))->Find(bvalue, vtMissing, XlFindLookIn::xlValues, vtCon,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, match_case, false, false))
		{
			iRowValue = PRS->Cells->Row;
		}

		pRange->GetRange(COleVariant(L"A1:A1"))->Find("", vtMissing, XlFindLookIn::xlFormulas, XlLookAt::xlPart,
			XlSearchOrder::xlByColumns, XlSearchDirection::xlNext, false, false, false);

		if (ishow == 1) PRShow->EntireColumn->PutHidden(blHidden);

		return iRowValue;
	}
	catch (...) {}

	CString str = L"";
	str.Format(L"Không tìm thấy giá trị '%s' trong bảng tính '%s'.",
		(LPCTSTR)bvalue, (LPCTSTR)pSheet->Name);
	if (itype == 1) _msgbox(str, 1);

	return 0;
}

CString Function::_xlGetValue(_WorksheetPtr pSheet, CString name, int itype, int inumberformat)
{
	try
	{
		if (inumberformat == 1)
		{
			RangePtr pRg = pSheet->Cells;
			return _xlGIOR(pRg, _xlGetRow(pSheet, name, 1), _xlGetColumn(pSheet, name, 0), L"");
		}
		else return (LPCTSTR)(_bstr_t)pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name)->GetValue2();
	}
	catch (...)
	{
		CString szval = DF_ERROR;
		int irow = 0, icolumn = 0;

		_xlCreateNewName(name, szval, irow, icolumn);

		if (szval != DF_ERROR) return szval;
		if (itype == 1) _xlMsgNotName(pSheet, name);
	}
	return L"";
}

void Function::_xlPutValue(_WorksheetPtr pSheet, CString name, CString szvalue, int itype)
{
	try
	{
		pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name)->PutValue2((_variant_t)szvalue);
	}
	catch (...)
	{
		CString szval = DF_ERROR;
		int irow = 0, icolumn = 0;

		_xlCreateNewName(name, szval, irow, icolumn);

		if (szval == DF_ERROR)
		{
			if (itype == 1) _xlMsgNotName(pSheet, name);
		}
	}
}

void Function::_xlPutValue(_WorksheetPtr pSheet, CString name, int ivalue, int itype)
{
	try
	{
		CString szval = L"";
		szval.Format(L"%d", ivalue);
		RangePtr PRS = pSheet->Cells->GetRange((_variant_t)name, (_variant_t)name);
		PRS->PutValue2((_variant_t)szval);
	}
	catch (...)
	{
		CString szval = DF_ERROR;
		int irow = 0, icolumn = 0;

		_xlCreateNewName(name, szval, irow, icolumn);

		if (szval == DF_ERROR)
		{
			if (itype == 1) _xlMsgNotName(pSheet, name);
		}
	}
}

int Function::_xlGetRow(_WorksheetPtr pSheet, CString name, int itype)
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

int Function::_xlGetColumn(_WorksheetPtr pSheet, CString name, int itype)
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

CString Function::_xlGIOR(RangePtr pRange, int nRow, int nColumn, CString szError)
{
	try
	{
		_bstr_t bstr = pRange->GetItem(nRow, nColumn);
		return (LPCTSTR)bstr;
	}
	catch (...) { return szError; }
}

CString Function::_xlGAR(RangePtr pRng, int nRow, int nColumn, int itype)
{
	RangePtr dRangeptr = pRng->GetItem(nRow, nColumn);
	_bstr_t szAdress = dRangeptr->GetAddress(false, false, XlReferenceStyle::xlA1);
	CString kq = (LPCTSTR)szAdress;
	if (itype == 0) return kq;

	int k = 0;
	int iLen = kq.GetLength();
	for (int i = 1; i < iLen; i++)
	{
		if (_wtoi(kq.Mid(i, 1)) > 0)
		{
			k = i;
			break;
		}
	}

	CString szval = L"";
	if (itype == 1) szval.Format(L"$%s", kq);
	else if (itype == 2) szval.Format(L"%s$%s", kq.Left(k), kq.Right(iLen - k));
	else if (itype == 3) szval.Format(L"$%s$%s", kq.Left(k), kq.Right(iLen - k));

	return szval;
}

CString Function::_xlConvertRCtoA1(int iColumn)
{
	try
	{
		CString szval = L"";
		CString szCol = L"_ABCDEFGHIJKLMNOPQRSTUVWXYZ";

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

void Function::_xlMsgNotName(_WorksheetPtr pSheet, CString name)
{
	CString str = L"";
	str.Format(L"Không tồn tại name '%s' trong bảng tính '%s'.", name, (LPCTSTR)pSheet->Name);
	_msgbox(str, 1, 1);
}

void Function::_xlCreateNewName(CString szName, CString &szKpname, int &iKprow, int &iKpcol)
{
	try
	{
		CString bsName = L"";
		_WorksheetPtr pSheet;
		RangePtr pRange;
		RangePtr PRS;

		if (szName == L"CF_CHECK_COLUMN")
		{
			bsName = (LPCTSTR)_xlGetNameSheet(shConfig, 1);
			pSheet = xl->Sheets->GetItem((_bstr_t)bsName);
			pRange = pSheet->Cells;

			szKpname = L"1";
			iKprow = 11, iKpcol = 3;
			
			PRS = pRange->GetItem(iKprow, iKpcol);
			PRS->PutName((_bstr_t)szName);
			PRS->PutValue2((_bstr_t)szKpname);

			pRange->PutItem(iKprow, iKpcol - 1, 
				(_bstr_t)L"Chỉ tìm kiếm tại các cột nội dung / mô tả");
		}
		else if (szName == L"CF_CHECK_WORKBOOK")
		{
			bsName = (LPCTSTR)_xlGetNameSheet(shConfig, 1);
			pSheet = xl->Sheets->GetItem((_bstr_t)bsName);
			pRange = pSheet->Cells;

			szKpname = L"0";
			iKprow = 12, iKpcol = 3;

			PRS = pRange->GetItem(iKprow, iKpcol);
			PRS->PutName((_bstr_t)szName);
			PRS->PutValue2((_bstr_t)szKpname);

			pRange->PutItem(iKprow, iKpcol - 1,
				(_bstr_t)L"Tìm kiếm nội dung trên toàn bộ Workbook");
		}		
	}
	catch (...) {}
}

void Function::_xlSetHyperlink(_WorksheetPtr pSheet, RangePtr PRS, CString pathFile , CString szName,
	COLORREF clrBkg, COLORREF clrTxt, CString szFontName, int iSize, bool bCenter, bool bFormula)
{
	try
	{
		pSheet->Hyperlinks->Add(PRS, (_bstr_t)pathFile);
		PRS->PutValue2((_bstr_t)szName);
		PRS->Interior->PutColor(clrBkg);
		PRS->Font->PutColor(clrTxt);
		PRS->Font->PutName((_bstr_t)szFontName);
		PRS->Font->PutSize(iSize);
		if (bCenter) PRS->PutHorizontalAlignment(xlCenter);
	}
	catch (...) {}
}

int Function::_xlGetAllSheet(vector<_WorksheetPtr> &wpSheet, int &nHidden)
{
	try
	{
		nHidden = 0;
		wpSheet.clear();

		_WorksheetPtr psh;
		int ncount = xl->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= ncount; i++)
		{
			psh = pWb->Worksheets->Item[i];
			if ((int)psh->GetVisible() != 2) wpSheet.push_back(psh);
			else nHidden++;
		}
		return (int)wpSheet.size();
	}
	catch (...) {}
	return 0;
}

int Function::_xlGetIndexSheet(CString szCodename)
{
	try
	{
		int nIndex = 0;
		int nHidden = 0;
		CString szval = L"";
		vector<_WorksheetPtr> vecPsh;
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
		vecPsh.clear();
		return nIndex;
	}
	catch (...) {}
	return 0;
}

bool Function::_xlCreateNewSheet(_WorksheetPtr &wPsh, int nIndex, CString szCodename, CString szName, COLORREF clrTab)
{
	try
	{
		// Total = count(Sheet) + (1) This Workbook + (2) Module
		int ncount = xl->ActiveWorkbook->Sheets->Count;		
		int total = ncount + 3;
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
	catch (...) {
		_msgbox(L"Lỗi tạo mới sheet. Nguyên nhân có thể Excel của bạn đã bị nhiễm virus 'NEGS' "
			L"hoặc Workbook của bạn đã đặt khóa 'VBAProject', vui lòng kiểm tra lại. "
			L"Chú ý trong cài đặt 'Trust Center Settings...' bạn cần tích 2 lựa chọn "
			L"'Enable all macros' và 'Trust access to the VBA project object model'. "
			L"Vui lòng liên hệ Tổng đài GXD 1900.0147 để được hỗ trợ giải đáp.", 2);
	}
	return false;
}


CString Function::_xlGetcomment(RangePtr pRange)
{
	try
	{
		if (pRange->GetComment() != NULL)
		{
			return (LPCTSTR)pRange->GetComment()->Text();
		}
		return L"";
	}
	catch (...) {}
	return L"";
}

void Function::_xlPutcomment(RangePtr pRange, CString szComment)
{
	try
	{
		if (pRange->GetComment() != NULL) pRange->ClearComments();
		pRange->AddComment((_bstr_t)szComment);
	}
	catch (...) {}
}

void Function::_xlFormatBorders(RangePtr pRange, int iStyle, bool blTop, bool blBottom)
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

CString Function::_xlGetHyperLink(RangePtr pRgSelect, int iType)
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


int Function::_xlGetAllHyperLink(RangePtr pRgSelect, vector<CString> &vecHyper, int iType, int iStatus)
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

CString Function::_xlCheckSelection()
{
	try
	{
		RangePtr PRS = pShActive->Application->Selection;
		if ((int)PRS->GetCount() > 0)
		{
			return (LPCTSTR)PRS->GetAddress(1, 1, XlReferenceStyle::xlR1C1);
		}
	}
	catch (...) {}
	return L"";
}

int Function::_xlGetAllCellSelection(int itypeCell, int itypeArea)
{
	try
	{
		vecSelect.clear();

		// Xác định vùng lựa chọn
		CString szArray = _xlCheckSelection();

		// Vị trí cuối cùng chứa dữ liệu
		int pRowEnd = pRgActive->SpecialCells(xlCellTypeLastCell)->GetRow();
		int pColumnEnd = pRgActive->SpecialCells(xlCellTypeLastCell)->GetColumn();

		CString szval = L"";
		vector<CString> vecKey;
		int iArr = _TackTukhoa(vecKey, szArray, L",", L";", 1, 0);
		if (iArr == 0)
		{
			iArr = 1;
			szval.Format(L"R1C1:R%dC%d", pRowEnd, pColumnEnd);
			vecKey.push_back(szval);

			if ((long)(pRowEnd*pColumnEnd) >= (long)MAX_ROW)
			{
				if (_y(L"Bạn đang chọn toàn bộ bảng tính. Bạn có chắc chắn với điều này?", 0) != 6)
				{
					vecKey.clear();
					return 0;
				}
			}
		}

		if (itypeArea == 1) iArr = 1;

		int pos = 0, x1 = 0, y1 = 0, x2 = 0, y2 = 0;
		CString szleft = L"", szright = L"";
		for (int i = 0; i < iArr; i++)
		{
			x1 = vecKey[i].Find(L"R");
			y1 = vecKey[i].Find(L"C");

			if (x1 == -1 || y1 == -1)
			{
				// x1 = -1 --> lựa chọn toàn bộ 1 cột nào đó  (C:1 --> pRowEnd)
				// y1 = -1 --> lựa chọn toàn bộ 1 dòng nào đó (R:1 --> pColumnEnd)
				if (x1 == -1 && y1 >= 0) _xlSaveR1C1(vecKey[i], L"C", pRowEnd, itypeCell);
				else if (x1 >= 0 && y1 == -1)  _xlSaveR1C1(vecKey[i], L"R", pColumnEnd, itypeCell);
			}
			else
			{
				RangePtr PRS;
				CString szformula = L"";
				pos = vecKey[i].Find(L":");
				if (pos > 0)
				{
					szleft = vecKey[i].Left(pos);
					szright = vecKey[i].Right(vecKey[i].GetLength() - pos - 1);

					szleft.Replace(L"R", L"");
					pos = szleft.Find(L"C");

					x1 = _wtoi(szleft.Left(pos));
					y1 = _wtoi(szleft.Right(szleft.GetLength() - pos - 1));

					szright.Replace(L"R", L"");
					pos = szright.Find(L"C");

					x2 = _wtoi(szright.Left(pos));
					y2 = _wtoi(szright.Right(szright.GetLength() - pos - 1));

					for (int j = x1; j <= x2; j++)
					{
						for (int k = y1; k <= y2; k++)
						{
							PRS = pRgActive->GetItem(j, k);
							szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

							szval = _xlGIOR(pRgActive, j, k, L"");
							_xlSaveTotalSelection(j, k, szval, szformula, itypeCell);
						}
					}
				}
				else
				{
					vecKey[i].Replace(L"R", L"");
					pos = vecKey[i].Find(L"C");

					x1 = _wtoi(vecKey[i].Left(pos));
					y1 = _wtoi(vecKey[i].Right(vecKey[i].GetLength() - pos - 1));

					PRS = pRgActive->GetItem(x1, y1);
					szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

					szval = _xlGIOR(pRgActive, x1, y1, L"");
					_xlSaveTotalSelection(x1, y1, szval, szformula, itypeCell);
				}
			}
		}

		vecKey.clear();
		blUndovalue = true;

		return (int)vecSelect.size();
	}
	catch (...) {}
	return 0;
}

void Function::_xlSaveTotalSelection(int iRow, int iColumn, CString szChuoi, CString szFormula, int itypeCell)
{
	try
	{
		szChuoi.Trim();
		if ((itypeCell == 1 && szChuoi == L"") || (itypeCell == 2 && szChuoi != L"")) return;

		int itotal = (int)vecSelect.size();
		if (itotal > 0)
		{
			for (int i = 0; i < itotal; i++)
			{
				if (iRow == vecSelect[i].row && iColumn == vecSelect[i].column)
				{
					return;
				}
			}
		}

		MyStrSelection MSSLC;

		MSSLC.row = iRow;
		MSSLC.column = iColumn;
		MSSLC.cell.Format(L"%s%d", _xlConvertRCtoA1(iColumn), iRow);
		
		MSSLC.sheet = sNameActive;
		MSSLC.bkgcolor = bkgColorActive;
		
		MSSLC.value = szChuoi;
		MSSLC.formula = szFormula;
		

		vecSelect.push_back(MSSLC);
	}
	catch (...) {}
}

void Function::_xlSaveR1C1(CString szPkey, CString RC, int iRCEnd, int itypeCell)
{
	try
	{
		RangePtr PRS;
		int x1 = 0, x2 = 0;
		CString szval = L"", szformula = L"";

		int _pos = szPkey.Find(L":");
		if (_pos > 0)
		{
			CString szleft = szPkey.Left(_pos);
			CString szright = szPkey.Right(szPkey.GetLength() - _pos - 1);
			szleft.Replace(RC, L"");
			szright.Replace(RC, L"");
			x1 = _wtoi(szleft);
			x2 = _wtoi(szright);

			for (int j = x1; j <= x2; j++)
			{
				for (int k = 1; k <= iRCEnd; k++)
				{
					if (RC == L"R")
					{
						PRS = pRgActive->GetItem(j, k);
						szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

						szval = _xlGIOR(pRgActive, j, k, L"");
						_xlSaveTotalSelection(j, k, szval, szformula, itypeCell);
					}
					else
					{
						PRS = pRgActive->GetItem(k, j);
						szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

						szval = _xlGIOR(pRgActive, k, j, L"");
						_xlSaveTotalSelection(k, j, szval, szformula, itypeCell);
					}
				}
			}
		}
		else
		{
			szPkey.Replace(RC, L"");
			x1 = _wtoi(szPkey);

			for (int k = 1; k <= iRCEnd; k++)
			{
				if (RC == L"R")
				{
					PRS = pRgActive->GetItem(x1, k);
					szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

					szval = _xlGIOR(pRgActive, x1, k, L"");
					_xlSaveTotalSelection(x1, k, szval, szformula, itypeCell);
				}
				else
				{
					PRS = pRgActive->GetItem(k, x1);
					szformula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

					szval = _xlGIOR(pRgActive, k, x1, L"");
					_xlSaveTotalSelection(k, x1, szval, szformula, itypeCell);
				}
			}
		}
	}
	catch (...) {}
}

void Function::_xlCallFunctionDLL(CString szDllName, CString szFunction, bool blQuit)
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
	}
	catch (...) {}
}

/***********************************************************************************************/

int Function::_CheckVersionSoftware()
{
	return _wtoi(VERSION_SOFT);
}

char* Function::_ConvertCstringToChars(CString szvalue)
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

CString Function::_ConvertCharsToCstring(char* chr)
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

int Function::_FileExists(CString pathFile, int itype)
{
	CString str = L"Đường dẫn file không hợp lệ. Vui lòng kiểm tra lại.\nPath = " + pathFile;

	try
	{
		struct stat fileInfo;
		if (!stat(_ConvertCstringToChars(pathFile), &fileInfo)) return 1;
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

int Function::_FileExistsUnicode(CString pathFile, int itype)
{
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

bool Function::_FolderExistsUnicode(CString dirName_in)
{
	DWORD dwAttrib = GetFileAttributes((LPCTSTR)dirName_in);
	return (dwAttrib != INVALID_FILE_ATTRIBUTES && (dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

void Function::_ShellExecuteSelectFile(CString szDirFileSelect)
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

BOOL Function::_CreateNewFolder(CString directoryPath)
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

bool Function::_CreateDirectories(CString pathName)
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

int Function::_FindAllFile(CFileFinder &_finder, CString pathFolder, CString typeOfFile)
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

int Function::_FindAllFileEx(CString szPathCha, vector<CString> &vecPathFiles)
{
	try
	{
		vecPathFiles.clear();

		if (szPathCha.CompareNoCase(L"") == 0) return 0;
		if (szPathCha.GetLength() < 5) return 0;

		CUtil::GetFileList(_tstring(szPathCha.GetBuffer()), L"*.*", false, vecPathFiles);

		return (int)vecPathFiles.size();
	}
	catch (...) {}
	return 0;
}

// Hàm tìm kiếm các thư mục trong thư mục cha
int Function::_FindAllFolder(CString szPathCha, vector<CString> &vecPathFolder)
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

CString Function::_GetPathFolder(CString fileName)
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

bool Function::_IsProcessRunning(const wchar_t *processName)
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

int Function::_TackTukhoa(vector<CString> &vecPkey, CString sKeyValue, CString symbol1, CString symbol2, int itypeFilter, int itypeTrim)
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

int Function::_RandomTime()
{
	time_t now = time(0);
	tm *ltm = localtime(&now);
	return (100*(int)ltm->tm_min + (int)ltm->tm_sec + rand() % 1000);
}

CString Function::_GetTimeNow(int ifulltime)
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

bool Function::_CompareDate(CString dateBefore, CString dateAfter)
{
	int d1 = _wtoi(dateBefore.Left(2));
	int d2 = _wtoi(dateAfter.Left(2));

	int m1 = _wtoi(dateBefore.Mid(3, 2));
	int m2 = _wtoi(dateAfter.Mid(3, 2));

	int y1 = _wtoi(dateBefore.Right(2));
	int y2 = _wtoi(dateAfter.Right(2));

	if (y1 < y2 || (y1 == y2 && m1 < m2) || (y1 == y2 && m1 == m2 && d1 <= d2)) return true;

	return false;
}

CString	Function::_DayNextPrev(CString szdate, int num)
{
	if (szdate == L"") return L"";
	if (num == 0) return szdate;

	CString szDay = szdate;
	int dd = _wtoi(szDay.Left(2));
	int mm = _wtoi(szDay.Mid(3, 2));
	int yy = _wtoi(szDay.Right(2)) + 2000;
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

	szDay.Format(L"%02d/%02d/%04d", dd, mm, yy);
	return szDay;
}

void Function::_QuickSortArr(vector<int> &vec, int L, int R)
{
	int i = L, j = R, swap = 0;
	int mid = L + (R - L) / 2;
	int piv = vec[mid];

	while (i < R || j > L)
	{
		while (vec[i] < piv) i++;
		while (vec[j] > piv) j--;

		if (i <= j)
		{
			swap = vec[i];
			vec[i] = vec[j];
			vec[j] = swap;

			i++;
			j--;
		}
		else
		{
			if (i < R) _QuickSortArr(vec, i, R);
			if (j > L) _QuickSortArr(vec, L, j);
			return;
		}
	}
}

CString Function::_GetDesktopDir()
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

CString Function::_GetTypeOfFile(CString pathFile)
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

CString Function::_GetNameOfPath(CString pathFile, int &pos, int ipath)
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

void Function::_WriteMultiUnicode(vector<CString> vecTextLine, CString spathFile, int itype)
{
	try
	{
		if (itype == 1)
		{
			if (_FileExistsUnicode(spathFile, 0) != 1) return;
		}

		locale loc(locale(), new codecvt_utf8<wchar_t>);
		wofstream myfile(_ConvertCstringToChars(spathFile));
		if (myfile)
		{
			myfile.imbue(loc);
			long total = (long)vecTextLine.size();
			for (long i = 0; i < total; i++)
			{
				wstring ws(vecTextLine[i]);
				myfile << ws << "\n";
			}
			myfile.close();
		}
	}
	catch (...) {}
}

int Function::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
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

void Function::_LogFileWrite(CString szLog, bool blClear, bool blOpenFile, CString szFileName)
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

void Function::_LogFileWrite(int iLog, bool blClear, bool blOpenFile, CString szFileName)
{
	CString str = L"";
	str.Format(L"%d", iLog);
	_LogFileWrite(str, blClear, blOpenFile, szFileName);
}

void Function::_LogFileStart(CString szStart) { _LogFileWrite(szStart, true); }
void Function::_LogFileEnd() { _LogFileWrite(L"<!>", false, true); }

CString Function::_ConvertKhongDau(CString szNoidung)
{
	CString scvtUTF[] = {
		L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
		L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
		L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
		L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
		L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",
		L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
		L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",
		L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
		L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"đ",
		L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"Đ" };

	CString scvtKOD[] = {
		L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",L"a",
		L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",L"A",
		L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",L"o",
		L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",L"O",
		L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",L"e",
		L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",L"E",
		L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",L"u",
		L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",L"U",
		L"i",L"i",L"i",L"i",L"i",L"y",L"y",L"y",L"y",L"y",L"d",
		L"I",L"I",L"I",L"I",L"I",L"Y",L"Y",L"Y",L"Y",L"Y",L"D" };

	int nsizeArr = getSizeArray(scvtUTF);
	for (int i = 0; i < nsizeArr; i++)
	{
		szNoidung.Replace(scvtUTF[i], scvtKOD[i]);
	}

	return szNoidung;
}

// Hàm chuyển đổi sang chữ HOA có chứa ký tự Unicode
CString Function::_UpperUnicode(CString szNoidung)
{
	CString szUpper = szNoidung;
	szUpper.MakeUpper();

	CString scvtLower[] = {
		L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
		L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
		L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",
		L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",
		L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"đ" };

	CString scvtUpper[] = {
		L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
		L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
		L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
		L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
		L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"Đ" };

	int nsizeArr = getSizeArray(scvtLower);
	for (int i = 0; i < nsizeArr; i++)
	{
		szUpper.Replace(scvtLower[i], scvtUpper[i]);
	}

	return szUpper;
}

// Hàm chuyển đổi sang chữ thường có chứa ký tự Unicode
CString Function::_LowerUnicode(CString szNoidung)
{
	CString szLower = szNoidung;
	szLower.MakeLower();

	CString scvtUpper[] = {
		L"À",L"Á",L"Ả",L"Ã",L"Ạ",L"Ă",L"Ằ",L"Ắ",L"Ẳ",L"Ẵ",L"Ặ",L"Â",L"Ầ",L"Ấ",L"Ẩ",L"Ẫ",L"Ậ",
		L"Ò",L"Ó",L"Ỏ",L"Õ",L"Ọ",L"Ô",L"Ồ",L"Ố",L"Ổ",L"Ỗ",L"Ộ",L"Ơ",L"Ờ",L"Ớ",L"Ở",L"Ỡ",L"Ợ",
		L"È",L"É",L"Ẻ",L"Ẽ",L"Ẹ",L"Ê",L"Ề",L"Ế",L"Ể",L"Ễ",L"Ệ",
		L"Ù",L"Ú",L"Ủ",L"Ũ",L"Ụ",L"Ư",L"Ừ",L"Ứ",L"Ử",L"Ữ",L"Ự",
		L"Ì",L"Í",L"Ỉ",L"Ĩ",L"Ị",L"Ỳ",L"Ý",L"Ỷ",L"Ỹ",L"Ỵ",L"Đ" };

	CString scvtLower[] = {
		L"à",L"á",L"ả",L"ã",L"ạ",L"ă",L"ằ",L"ắ",L"ẳ",L"ẵ",L"ặ",L"â",L"ầ",L"ấ",L"ẩ",L"ẫ",L"ậ",
		L"ò",L"ó",L"ỏ",L"õ",L"ọ",L"ô",L"ố",L"ồ",L"ổ",L"ỗ",L"ộ",L"ơ",L"ờ",L"ớ",L"ở",L"ỡ",L"ợ",
		L"è",L"é",L"ẻ",L"ẽ",L"ẹ",L"ê",L"ề",L"ế",L"ể",L"ễ",L"ệ",
		L"ù",L"ú",L"ủ",L"ũ",L"ụ",L"ư",L"ừ",L"ứ",L"ử",L"ữ",L"ự",
		L"ì",L"í",L"ỉ",L"ĩ",L"ị",L"ỳ",L"ý",L"ỷ",L"ỹ",L"ỵ",L"đ" };

	int nsizeArr = getSizeArray(scvtUpper);
	for (int i = 0; i < nsizeArr; i++)
	{
		szLower.Replace(scvtUpper[i], scvtLower[i]);
	}

	return szLower;
}

CString Function::_ReplaceMatchCase(CString szNoidung, CString szOld, CString szNew, bool blCase, bool blKDau)
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
			szLower = _LowerUnicode(szLower);
			szOld = _LowerUnicode(szOld);
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

double Function::_RoundDouble(double doValue, int nPrecision)
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

CString Function::_ConvertBytes(double ivalue)
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

void Function::_GetFileAttributesEx(CString szDir, CString &szKichthuoc, CString &szNgay)
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

			if((int)stLocal.wYear >= 2000)
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

CString Function::_FormatTienVND(CString szTien, CString szSteparator, CString szDecimal)
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

CString Function::_ConvertRename(CString sname)
{
	if (sname == L"") return L"";

	sname.Replace(L"\\", L"");
	sname.Replace(L"/", L"");
	sname.Replace(L":", L"");
	sname.Replace(L"*", L"");
	sname.Replace(L"?", L"");
	sname.Replace(L"\"", L"");
	sname.Replace(L"<", L"");
	sname.Replace(L">", L"");
	sname.Replace(L"|", L"");	

	return sname;
}

bool Function::_CheckRegisterCOM()
{
	try
	{
		HRESULT hr = CoInitialize(NULL);
		IPropertiesPtr pILib(__uuidof(CProperties));

		if ((LPCTSTR)pILib->TestRegister() != NULL)
		{
			_msgbox(L"Đăng ký COM thành công.", 0, 0, 10000);
		}

		CoUninitialize();
		return true;
	}
	catch (...) {
		_msgbox(L"Đăng ký COM thất bại!", 2, 0);
	}
	return false;
}

CString Function::_GetProperties(CString szpath, int nIndex)
{
	try
	{
		HRESULT hr = CoInitialize(NULL);
		IPropertiesPtr pILib(__uuidof(CProperties));
		
		CString szInfo = L"";
		CString szval = szpath.Right(3);
		szval.MakeLower();
		if (szval == L"pdf")
		{
			szInfo = (LPCTSTR)pILib->GetPDFInfo((_bstr_t)szpath, nIndex);
		}
		else
		{
			szInfo = (LPCTSTR)pILib->GetOfficeInfo((_bstr_t)szpath, nIndex);
		}

		CoUninitialize();
		return szInfo;
	}
	catch (...) {}
	return L"";
}

bool Function::_SetProperties(CString szpath, int nIndex, CString strVal)
{
	try
	{
		HRESULT hr = CoInitialize(NULL);
		IPropertiesPtr pILib(__uuidof(CProperties));

		bool bl = false;
		CString sztype = szpath.Right(3);
		sztype.MakeLower();
		if (sztype == L"pdf")
		{
			bl = (bool)pILib->SetPDFInfo((_bstr_t)szpath, nIndex, (_bstr_t)strVal);
		}
		else
		{
			bl = (bool)pILib->SetOfficeInfo((_bstr_t)szpath, nIndex, (_bstr_t)strVal);
		}

		CoUninitialize();
		return bl;
	}
	catch (...) {}
	return false;
}

bool Function::_SetAllProperties(
	CString szpath, CString szTitle, CString szSubject, CString szTags, 
	CString szCategories, CString szComments, CString szAuthor, CString szCompany)
{
	try
	{
		HRESULT hr = CoInitialize(NULL);
		IPropertiesPtr pILib(__uuidof(CProperties));
		
		bool bl = false;		
		CString szval = szpath.Right(3);
		szval.MakeLower();
		if (szval == L"pdf")
		{
			szval.Format(L"%s|%s", szAuthor, szCompany);
			bl = (bool)pILib->SetPDFInfoAll((_bstr_t)szpath,
				(_bstr_t)szTitle, (_bstr_t)szSubject, (_bstr_t)szTags, 
				(_bstr_t)szCategories, (_bstr_t)szComments, (_bstr_t)szval);
		}
		else
		{
			bl = (bool)pILib->SetOfficeInfoAll(
				(_bstr_t)szpath, (_bstr_t)szTitle, (_bstr_t)szSubject, (_bstr_t)szTags, 
				(_bstr_t)szCategories, (_bstr_t)szComments, (_bstr_t)szAuthor, (_bstr_t)szCompany);
		}
		
		CoUninitialize();
		return bl;
	}
	catch (...) {}
	return false;
}

bool Function::_Is64BitWindows()
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

/*** Phần hàm viết riêng **********************************************************************/

void Function::_GetThongtinSheetCategory(int &icotSTT, int &icotLV, int &icotMH, int &icotNoidung,
	int &icotFilegoc, int &icotFileGXD, int &icotThLuan, int &icotEnd, int &irowTieude, int &irowStart, int &irowEnd)
{
	try
	{
		_bstr_t bsCategory = _xlGetNameSheet(shCategory, 0);
		_WorksheetPtr psCategory = xl->Sheets->GetItem(bsCategory);
		RangePtr prCategory = psCategory->Cells;

		icotSTT = _xlGetColumn(psCategory, L"CT_STT", 0);
		icotLV = _xlGetColumn(psCategory, L"CT_LINHVUC", 0);
		icotMH = _xlGetColumn(psCategory, L"CT_MASO", 0);
		icotNoidung = _xlGetColumn(psCategory, L"CT_NOIDUNG", 0);
		icotFilegoc = _xlGetColumn(psCategory, L"CT_FILEGOC", 0);
		icotFileGXD = _xlGetColumn(psCategory, L"CT_FILEGXD", 0);
		icotThLuan = _xlGetColumn(psCategory, L"CT_THLUAN", 0);
		icotEnd = _xlGetColumn(psCategory, L"CT_COLUMN_END", 0);

		irowTieude = _xlGetRow(psCategory, L"CT_STT", 0);
		irowStart = irowTieude + 1;

		CString szval = L"";
		for (int i = irowStart; i <= irowStart + 5; i++)
		{
			szval = _xlGIOR(prCategory, irowStart, icotSTT, L"");
			if (szval.Left(1) == L"[" && szval.Right(1) == L"]")
			{
				irowStart = i + 1;
				break;
			}
		}

		irowEnd = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
		if (irowEnd < irowStart) irowEnd = irowStart;
	}
	catch (...) {}
}

void Function::_GetThongtinSheetFManager(int &icotSTT, int &icotLV, int &icotFolder, int &icotTenFile, int &icotType,
	int &icotNoidung, int &icotFilegoc, int &icotEnd, int &irowTieude, int &irowStart, int &irowEnd)
{
	try
	{
		_bstr_t bsFManager = _xlGetNameSheet(shFManager, 0);
		_WorksheetPtr psFManager = xl->Sheets->GetItem(bsFManager);
		RangePtr prFManager = psFManager->Cells;

		icotSTT = _xlGetColumn(psFManager, L"MNF_STT", 0);
		icotLV = _xlGetColumn(psFManager, L"MNF_LINHVUC", 0);
		icotFolder = _xlGetColumn(psFManager, L"MNF_DIR", 0);
		icotTenFile = _xlGetColumn(psFManager, L"MNF_NAME", 0);
		icotType = _xlGetColumn(psFManager, L"MNF_TYPE", 0);
		icotNoidung = _xlGetColumn(psFManager, L"MNF_NOIDUNG", 0);
		icotFilegoc = _xlGetColumn(psFManager, L"MNF_LINKGOC", 0);
		icotEnd = _xlGetColumn(psFManager, L"MNF_COLUMN_END", 0);

		irowTieude = _xlGetRow(psFManager, L"MNF_STT", 0);
		irowStart = irowTieude + 1;

		CString szval = L"";
		for (int i = irowStart; i <= irowStart + 5; i++)
		{
			szval = _xlGIOR(prFManager, irowStart, icotSTT, L"");
			if (szval.Left(1) == L"[" && szval.Right(1) == L"]")
			{
				irowStart = i + 1;
				break;
			}
		}

		irowEnd = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
		if (irowEnd < irowStart) irowEnd = irowStart;
	}
	catch (...) {}
}


void Function::_DefaultSheetActive(CString szSheet)
{
	try
	{
		CString szCopy = shCopy, szFile = shFile;
		if (sCodeActive != shCategory && sCodeActive != shFManager
			&& sCodeActive.Left(szCopy.GetLength()) != szCopy
			&& sCodeActive.Left(szFile.GetLength()) != szFile)
		{
			sCodeActive = szSheet;
			sNameActive = (LPCTSTR)_xlGetNameSheet((_bstr_t)sCodeActive, 1);
			pShActive = xl->Sheets->GetItem((_bstr_t)sNameActive);
			pRgActive = pShActive->Cells;

			if ((int)pShActive->GetVisible() != -1) pShActive->PutVisible(XlSheetVisibility::xlSheetVisible);
			pShActive->Activate();

			iRowActive = pShActive->Application->ActiveCell->GetRow();
			iColumnActive = pShActive->Application->ActiveCell->GetColumn();
		}		
	}
	catch (...) {}
}

void Function::_RegistryCOMGpro()
{
	try
	{
		// Kiểm tra đã đăng ký COM chưa?
		Base64Ex *bb = new Base64Ex;
		CRegistry *reg = new CRegistry;
		
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + EditProperties;
		reg->CreateKey(szCreate);
		
		CString szval = reg->ReadString(L"RegisterCOM", L"");
		if (_wtoi(szval) != 1)
		{
			int result = _y(L"Để đọc được nội dung thông tin files "
				L"bạn cần phải đăng ký thư viện hỗ trợ trong máy tính. "
				L"Chọn YES để đăng ký.", 0, 0, 0);
			if (result == 6)
			{
				CString pathFolderDir = _GetPathFolder(strNAMEDLL);
				CString szpathFile = pathFolderDir + FileRegistry;
				if (_FileExistsUnicode(szpathFile, 1))
				{
					ShellExecute(NULL, L"open", szpathFile, NULL, NULL, SW_SHOWMAXIMIZED);
				}
			}
		}

		delete reg;
		delete bb;
	}
	catch (...) {}
}

void Function::_xlOpenFileEx(_WorksheetPtr pSheet, int nRowActive, int nColumnActive)
{
	try
	{
		CString szCopy = shCopy, szFile = shFile;
		CString szCode = (LPCTSTR)pSheet->CodeName;
		if (szCode == shCategory || szCode == shFManager
			|| szCode.Left(szCopy.GetLength()) == szCopy
			|| szCode.Left(szFile.GetLength()) == szFile
			|| szCode == shTieuchuan)
		{
			_WorksheetPtr psh;
			int icotFilegoc = 0, icotFileGXD = 0;

			if (szCode == shTieuchuan)
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shTieuchuan, 1));
				icotFilegoc = _xlGetColumn(psh, L"TCH_OPEN", -1);
			}		
			else if (szCode == shCategory || szCode.Left(szCopy.GetLength()) == szCopy)
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shCategory, 1));
				icotFilegoc = _xlGetColumn(psh, L"CT_FILEGOC", -1);
				icotFileGXD = _xlGetColumn(psh, L"CT_FILEGXD", -1);
			}
			else
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shFManager, 1));
				icotFilegoc = _xlGetColumn(psh, L"MNF_LINKGOC", -1);
			}

			int iColumnFile = icotFilegoc;
			if (icotFilegoc == 0 || nColumnActive == icotFileGXD) iColumnFile = icotFileGXD;

			RangePtr pRange = pSheet->Cells;
			RangePtr PRS = pRange->GetItem(nRowActive, iColumnFile);

			if (iColumnFile > 0)
			{
				CString szval = _xlGIOR(pRange, nRowActive, iColumnFile, L"");
				if (szval != L"")
				{
					vector<CString> vecHyper;
					int ncoutHyp = _xlGetAllHyperLink(PRS, vecHyper, 1, 0);
					if (ncoutHyp > 0)
					{
						if (_FileExistsUnicode(vecHyper[0], 0) == 1)
						{
							ShellExecute(NULL, L"open", vecHyper[0], NULL, NULL, SW_SHOWMAXIMIZED);
						}
						else
						{
							szval = _GetPathFolder(strNAMEDLL) + vecHyper[0];
							if (_FileExistsUnicode(szval, 0) == 1)
							{
								PRS = pRange->GetItem(nRowActive, iColumnFile);
								_xlSetHyperlink(pSheet, PRS, szval);

								ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
							}
							else
							{
								int jcotFolder = _xlGetColumn(pSheet, L"MNF_DIR", -1);
								if (jcotFolder > 0)
								{
									int jcotTen = _xlGetColumn(pSheet, L"MNF_NAME", 0);
									int jcotType = _xlGetColumn(pSheet, L"MNF_TYPE", 0);

									// Đặt lại hyperlink
									CString szDir = _xlGIOR(pRange, nRowActive, jcotFolder, L"");
									CString szName = _xlGIOR(pRange, nRowActive, jcotTen, L"");
									CString szType = _xlGIOR(pRange, nRowActive, jcotType, L"");
									if (szDir.Right(1) != L"\\") szDir += L"\\";

									szval = szDir + szName + L"." + szType;
									if (_FileExistsUnicode(szval, 0) == 1)
									{
										PRS = pRange->GetItem(nRowActive, iColumnFile);
										_xlSetHyperlink(pSheet, PRS, szval);

										ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
									}
								}
							}
						}
					}
					vecHyper.clear();
				}
			}
		}
	}
	catch (...) {}
}


void Function::_xlOpenFolderEx(_WorksheetPtr pSheet, int nRowActive, int nColumnActive)
{
	try
	{
		CString szCopy = shCopy, szFile = shFile;
		CString szCode = (LPCTSTR)pSheet->CodeName;		
		if (szCode == shCategory || szCode == shFManager
			|| szCode.Left(szCopy.GetLength()) == szCopy
			|| szCode.Left(szFile.GetLength()) == szFile 
			|| szCode == shTieuchuan)
		{
			_WorksheetPtr psh;
			int icotFolder = 0, icotFilegoc = 0, icotFileGXD = 0;
			
			if (szCode == shTieuchuan)
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shTieuchuan, 1));
				icotFilegoc = _xlGetColumn(psh, L"TCH_OPEN", -1);
			}
			else if (szCode == shCategory || szCode.Left(szCopy.GetLength()) == szCopy)
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shCategory, 1));
				icotFilegoc = _xlGetColumn(psh, L"CT_FILEGOC", -1);
				icotFileGXD = _xlGetColumn(psh, L"CT_FILEGXD", -1);
			}
			else
			{
				psh = xl->Sheets->GetItem(_xlGetNameSheet(shFManager, 1));
				icotFilegoc = _xlGetColumn(psh, L"MNF_LINKGOC", -1);
				icotFolder = _xlGetColumn(pSheet, L"MNF_DIR", 0);
			}

			int iColumnFile = icotFilegoc;
			if (icotFilegoc == 0 || nColumnActive == icotFileGXD) iColumnFile = icotFileGXD;

			RangePtr pRange = pSheet->Cells;
			RangePtr PRS = pRange->GetItem(nRowActive, iColumnFile);

			if (iColumnFile > 0)
			{
				CString szval = _xlGIOR(pRange, nRowActive, iColumnFile, L"");
				if (szval != L"")
				{
					vector<CString> vecHyper;
					int ncoutHyp = _xlGetAllHyperLink(PRS, vecHyper, 1, 0);
					if (ncoutHyp > 0)
					{
						if (_FileExistsUnicode(vecHyper[0], 0) == 1)
						{
							_ShellExecuteSelectFile(vecHyper[0]);
						}
						else
						{
							szval = _GetPathFolder(strNAMEDLL) + vecHyper[0];
							if (_FileExistsUnicode(szval, 0) == 1)
							{
								PRS = pRange->GetItem(nRowActive, iColumnFile);
								_xlSetHyperlink(pSheet, PRS, szval);

								_ShellExecuteSelectFile(szval);
							}
							else
							{
								if (icotFolder > 0)
								{
									int jcotTen = _xlGetColumn(pSheet, L"MNF_NAME", -1);
									int jcotType = _xlGetColumn(pSheet, L"MNF_TYPE", 0);

									// Đặt lại hyperlink
									CString szDir = _xlGIOR(pRange, nRowActive, icotFolder, L"");
									CString szName = _xlGIOR(pRange, nRowActive, jcotTen, L"");
									CString szType = _xlGIOR(pRange, nRowActive, jcotType, L"");
									if (szDir.Right(1) != L"\\") szDir += L"\\";

									szval = szDir + szName + L"." + szType;
									if (_FileExistsUnicode(szval, 0) == 1)
									{
										PRS = pRange->GetItem(nRowActive, iColumnFile);
										_xlSetHyperlink(pSheet, PRS, szval);

										_ShellExecuteSelectFile(szval);
									}
									else
									{
										ShellExecute(NULL, L"open", szDir, NULL, NULL, SW_SHOWMAXIMIZED);
									}
								}
							}
						}
					}
					vecHyper.clear();
				}
				else
				{
					if (icotFolder > 0)
					{
						szval = _xlGIOR(pRange, nRowActive, icotFolder, L"");
						ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
					}
				}
			}
		}		
	}
	catch (...) {}
}