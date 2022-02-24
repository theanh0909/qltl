#include "pch.h"
#include "DlgCopyData.h"

IMPLEMENT_DYNAMIC(DlgCopyData, CDialogEx)

DlgCopyData::DlgCopyData(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgCopyData::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;

	colLV = 1, colSL = 2;	
	iOnOK = 1;

	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 50;

	iwLV = 300, iwSL = 60, iDlgW = 0, iDlgH = 0;
}

DlgCopyData::~DlgCopyData()
{
	delete ff;
	delete bb;
	delete reg;

	vecItem.clear();
}

void DlgCopyData::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, COPYDT_CHECK_ALL, m_check_all);
	DDX_Control(pDX, COPYDT_LIST_VIEW, m_list_view);
	DDX_Control(pDX, COPYDT_BTN_OK, m_btn_ok);
	DDX_Control(pDX, COPYDT_BTN_CANCEL, m_btn_cancel);
}

BEGIN_MESSAGE_MAP(DlgCopyData, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()

	ON_BN_CLICKED(COPYDT_BTN_OK, &DlgCopyData::OnBnClickedBtnOk)
	ON_BN_CLICKED(COPYDT_BTN_CANCEL, &DlgCopyData::OnBnClickedBtnCancel)
	ON_BN_CLICKED(COPYDT_CHECK_ALL, &DlgCopyData::OnBnClickedCheckAll)
	ON_NOTIFY(LVN_ENDSCROLL, COPYDT_LIST_VIEW, &DlgCopyData::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgCopyData::OnHdnEndtrackListView)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgCopyData)
	EASYSIZE(COPYDT_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(COPYDT_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(COPYDT_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(COPYDT_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgCopyData message handlers
BOOL DlgCopyData::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_COMPARE));
	SetIcon(hIcon, FALSE);

	_LoadRegistry();
	_SetDefault();

	_LoadListView();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgCopyData::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(UPD_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			OnBnClickedBtnCancel();
			return TRUE;
		}
	}

	return FALSE;
}

void DlgCopyData::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgCopyData::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgCopyData::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

/****************** Click button Ok **********************************************************/
void DlgCopyData::OnBnClickedBtnOk()
{
	try
	{
		int ncheck = 0;
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1)
			{
				ncheck++;
				break;
			}
		}

		if (ncheck > 0)
		{
			iOnOK = 1;
			_SaveRegistry();

			CDialogEx::OnOK();

			for (int i = 0; i < ncount; i++)
			{
				if ((int)m_list_view.GetCheck(i) != 1)
				{
					// Loại bỏ các dữ liệu không được tích chọn
					vecItem[i].blEnable = false;
				}
			}
		}
		else
		{
			ff->_msgbox(L"Bạn chưa tích chọn lĩnh vực cần sao chép.", 2, 0);
		}
	}
	catch (...) {}
}

void DlgCopyData::OnBnClickedBtnCancel()
{
	iOnOK = 0;
	_SaveRegistry();

	CDialogEx::OnCancel();
}

void DlgCopyData::_SetLocationAndSize()
{
	try
	{
		if (iDlgW > 0 && iDlgH > 0)
		{
			int isx = GetSystemMetrics(SM_CXSCREEN);
			int isy = GetSystemMetrics(SM_CYSCREEN);

			if (iDlgW >= isx && iDlgH + 50 >= isy)
			{
				this->ShowWindow(SW_MAXIMIZE);
			}
			else
			{
				this->SetWindowPos(NULL, 0, 0, iDlgW, iDlgH, SWP_NOMOVE | SWP_NOZORDER);

				CRect r;
				GetWindowRect(&r);
				ScreenToClient(&r);
				MoveWindow((isx - r.Width()) / 2, (isy - r.Height()) / 2, r.right - r.left, r.bottom - r.top, TRUE);
			}
		}
	}
	catch (...) {}
}

void DlgCopyData::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CopyData;
		reg->CreateKey(szCreate);

		CString szval = L"";
		iwLV = m_list_view.GetColumnWidth(colLV);
		iwSL = m_list_view.GetColumnWidth(colSL);
		szval.Format(L"%d|%d", iwLV, iwSL);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgCopyData::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CopyData;
		reg->CreateKey(szCreate);

		int pos = -1;
		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			pos = szval.Find(L"|");
			if (pos > 0)
			{
				iwLV = _wtoi((szval.Left(pos)).Trim());
				iwSL = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}

		szval = reg->ReadString(DlgSize, L"");
		if (szval != L"")
		{
			pos = szval.Find(L"x");
			if (pos > 0)
			{
				iDlgW = _wtoi((szval.Left(pos)).Trim());
				iDlgH = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}
	}
	catch (...) {}
}

void DlgCopyData::OnBnClickedCheckAll()
{
	int icheck = m_check_all.GetCheck();
	int ncount = m_list_view.GetItemCount();
	for (int i = 0; i < ncount; i++)
	{
		m_list_view.SetCheck(i, icheck);
	}
}

void DlgCopyData::_SetDefault()
{
	m_btn_ok.SetThemeHelper(&m_ThemeHelper);
	m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

	m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_view.InsertColumn(colLV, L" Lĩnh vực", LVCFMT_LEFT, iwLV);
	m_list_view.InsertColumn(colSL, L" Số lượng", LVCFMT_CENTER, iwSL);

	m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
	m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
		LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);
}

void DlgCopyData::_DeleteAllList()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

void DlgCopyData::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if (iStopload == 0) return;

		RECT r;
		CRect rectCtrl;
		m_list_view.GetItemRect(0, &r, LVIR_BOUNDS);
		m_list_view.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;
		int b = rectCtrl.Height();
		int topIndex = m_list_view.GetTopIndex();
		int ncount = m_list_view.GetItemCount();
		if (topIndex + (int)(b / a) >= ncount)
		{
			lanshow++;
			int iz = ibuocnhay * lanshow;
			if (tslkq <= iz)
			{
				iz = tslkq;
				iStopload = 0;
			}
			else iStopload = 1;

			// Thêm dữ liệu vào list kết quả
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz, (int)m_check_all.GetCheck());
		}
	}
	catch (...) {}
}

void DlgCopyData::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 22;
	*pResult = 0;
}

int DlgCopyData::_LoadListView()
{
	try
	{
		_DeleteAllList();

		tslkq = (int)vecItem.size();

		CString str = L"";
		str.Format(L"Tích chọn tất cả (%d lĩnh vực)", tslkq);
		m_check_all.SetCheck(true);
		m_check_all.SetWindowText(str);

		lanshow = 1;
		int iz = ibuocnhay;
		if (tslkq <= iz)
		{
			iz = tslkq;
			iStopload = 0;
		}
		else iStopload = 1;

		// Thêm dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz, (int)m_check_all.GetCheck());

		return iz;
	}
	catch (...) {}
	return 0;
}

void DlgCopyData::_AddItemsListKetqua(int iBegin, int iEnd, int icheck)
{
	try
	{
		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);			
			m_list_view.SetItemText(i, colLV, vecItem[i].szLinhvuc);

			szval.Format(L"%d", (int)vecItem[i].irow.size());
			m_list_view.SetItemText(i, colSL, szval);

			m_list_view.SetCheck(i, icheck);
		}
	}
	catch (...) {}
}

void DlgCopyData::_XacdinhSheetLienquan()
{
	jcotFolder = 0, jcotType = 0, jcotNoidung = 0;

	CString szCopy = shCopy;
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

/****************** Load lĩnh vực ****************************************************************/
int DlgCopyData::_LoadAllCategory()
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		_XacdinhSheetLienquan();

		vecItem.clear();
		
		MyStrCategory MSCT;

		CString szNameAc = sNameActive;
		szNameAc = ff->_ConvertKhongDau(szNameAc);
		szNameAc.MakeLower();

		int ncount = 0, icheck = 0;
		CString szval = L"", szKodau = L"";
		for (int i = irowStart; i < irowEnd; i++)
		{
			szval = ff->_xlGIOR(pRgActive, i, jcotLV, DF_NULL).Trim();
			szval = ff->_ConvertRename(szval);
			szval.Replace(L" ", L"");
			if (szval != L"")
			{
				szKodau = ff->_ConvertKhongDau(szval);
				szKodau.MakeLower();
				if (szKodau != szNameAc)
				{
					icheck = 0;
					ncount = (int)vecItem.size();
					if (ncount > 0)
					{
						for (int j = 0; j < ncount; j++)
						{
							if (szKodau == vecItem[j].szKodau)
							{
								vecItem[j].irow.push_back(i);
								icheck = 1;
								break;
							}
						}
					}

					if (icheck == 0)
					{
						MSCT.irow.clear();
						MSCT.szLinhvuc = szval;
						MSCT.szKodau = szKodau;
						MSCT.irow.push_back(i);
						MSCT.blEnable = true;

						vecItem.push_back(MSCT);
					}
				}
			}
		}

		ncount = (int)vecItem.size();
		if (ncount > 0)
		{
			ff->_xlPutScreenUpdating(true);
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DoModal();
		}
		else
		{
			ff->_msgbox(L"Không tồn tại lĩnh vực sao chép.", 1, 0);
			return 0;
		}

		// Lọc dữ liệu cập nhật
		for (int i = ncount - 1; i >= 0; i--)
		{
			if (vecItem[i].blEnable == false)
			{
				vecItem.erase(vecItem.begin() + i);
			}
		}

		ncount = (int)vecItem.size();
		if (iOnOK == 0 || ncount == 0) return 0;

		ff->_xlPutScreenUpdating(false);

		tongCopy = 0;
		for (int i = 0; i < ncount; i++)
		{
			tongCopy += (int)vecItem[i].irow.size();
		}

		// Xác định tất cả các sheet sử dụng
		int nHidden = 0;
		vector<_WorksheetPtr> vecPsh;
		int ncountSheet = ff->_xlGetAllSheet(vecPsh, nHidden);

		MyStrSheet MSSHEET;
		vector<MyStrSheet> vecSheet;
		CString szCopy = shCopy, szFile = shFile;
		CString szCodeNameNew = szFile;
		COLORREF clrSheet = RGB(146, 208, 80);

		if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
		{
			szCodeNameNew = szCopy;
			clrSheet = RGB(218, 150, 148);
		}

		for (int i = 0; i < ncountSheet; i++)
		{
			szval = (LPCTSTR)vecPsh[i]->CodeName;
			if (szval != sCodeActive)
			{
				if (szval == shCategory || szval == shFManager
					|| szval.Left(szCopy.GetLength()) == szCopy
					|| szval.Left(szFile.GetLength()) == szFile)
				{
					MSSHEET.szName = (LPCTSTR)vecPsh[i]->Name;
					MSSHEET.szCode = szval;

					MSSHEET.szLinhvuc = MSSHEET.szName;
					MSSHEET.szKodau = ff->_ConvertKhongDau(MSSHEET.szName);
					MSSHEET.szKodau.MakeLower();

					vecSheet.push_back(MSSHEET);
				}
			}					
		}

		int nIndexSheet = ff->_xlGetIndexSheet(sCodeActive);

		int isizeSh = 0;
		int dem = 0, cd = 0, nm = 0;
		CString szNewName = L"", szNewCode = L"";
		CString szSosanh[2] = { L"", L"" };
		_WorksheetPtr pSheet;

		for (int i = 0; i < ncount; i++)
		{
			// Kiểm tra lĩnh vực đã tồn tại chưa. Nếu chưa tồn tại --> Tạo mới sheet
			icheck = 0;
			isizeSh = (int)vecSheet.size();
			for (int j = 0; j < isizeSh; j++)
			{
				if (vecItem[i].szLinhvuc == vecSheet[j].szLinhvuc)
				{
					pSheet = xl->Sheets->GetItem((_bstr_t)vecSheet[j].szName);
					icheck = 1;
					break;
				}
			}
			
			// Chưa tồn tại --> Tạo mới sheet
			if (icheck == 0)
			{
				nm = 0;
				szNewName = L"";
				while (true)
				{
					szval = vecItem[i].szLinhvuc;
					if (szval.GetLength() > 30) szval = szval.Left(30);
					if (nm > 0) szNewName.Format(L"%s (%d)", szval, nm);
					else szNewName = szval;

					szval = ff->_ConvertKhongDau(szNewName);
					szval.MakeLower();
					icheck = 0;

					for (int j = 0; j < isizeSh; j++)
					{
						if (szval == vecSheet[j].szKodau || szval == sNameActive)
						{
							icheck = 1;
							break;
						}
					}

					if (icheck == 0) break;
					nm++;
				}

				cd = 0;
				szNewCode = L"";
				while (true)
				{
					cd++;
					szNewCode.Format(L"%s%03d", szCodeNameNew, cd);

					icheck = 0;
					for (int j = 0; j < isizeSh; j++)
					{
						if (szNewCode == vecSheet[j].szCode || szNewCode == sCodeActive)
						{
							icheck = 1;
							break;
						}
					}
					
					if (icheck == 0) break;
				}

				if (ff->_xlCreateNewSheet(pSheet, nIndexSheet + i, szNewCode, szNewName, clrSheet))
				{
					// Sao chép tiêu đề sang sheet mới
					_CopyStyleCategory(pSheet);

					MSSHEET.szName = (LPCTSTR)pSheet->Name;
					MSSHEET.szCode = (LPCTSTR)pSheet->CodeName;
					MSSHEET.szLinhvuc = vecItem[i].szLinhvuc;

					vecSheet.push_back(MSSHEET);
				}
			}

			// Cập nhật dữ liệu sang các sheet
			_UpdateCategory(pSheet, i, dem);
		}

		vecPsh.clear();
		vecSheet.clear();

		ff->_xlPutStatus(L"Ready");
		ff->_msgbox(L"Sao chép dữ liệu thành công!", 0, 1, 2000);
		ff->_xlPutScreenUpdating(true);

		return 1;
	}
	catch (...) {}
	return 0;
}

void DlgCopyData::_CopyStyleCategory(_WorksheetPtr pSheet)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		pRange->PutWrapText(true);
		pRange->Rows->AutoFit();
		pRange->PutVerticalAlignment(xlCenter);
		pRange->Font->PutName(FontTimes);
		pRange->Font->PutSize(13);

		RangePtr PRS = pSheet->GetRange(pRange->GetItem(1, 1), pRange->GetItem(1, icotEnd));
		PRS->EntireColumn->PutHorizontalAlignment(xlCenter);

		RangePtr PRSCopy = pShActive->GetRange(pRgActive->GetItem(1, 1), pRgActive->GetItem(irowStart - 1, 1));
		PRSCopy->EntireRow->Copy();

		PRS = pSheet->GetRange(pRange->GetItem(1, 1), pRange->GetItem(irowStart - 1, 1));
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteValues, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteFormats, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteColumnWidths, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		xl->PutCutCopyMode(XlCutCopyMode(false));

		PRS = pRange->GetItem(irowStart - 1, 1);
		PRS->EntireRow->AutoFilter(1, vtMissing, XlAutoFilterOperator::xlAnd, vtMissing, TRUE);

		pSheet->Application->ActiveWindow->PutZoom(85);
	}
	catch (...) {}
}

bool DlgCopyData::_UpdateCategory(_WorksheetPtr pSheet, int nIndex, int &dem)
{
	try
	{
		if ((int)pSheet->GetVisible() != -1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
		pSheet->Activate();

		CString szval = L"", szDisplay = L"";
		CString szUpdateStatus = L"Đang tiến hành sao chép lĩnh vực. Vui lòng đợi trong giây lát...";

		int num = 0, pos = 0;
		int iRowLast = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
		RangePtr pRange = pSheet->Cells;

		RangePtr PRS = pRange->GetItem(1, icotMH);
		bool blHidden = PRS->EntireColumn->GetHidden();
		if (blHidden == true) PRS->EntireColumn->PutHidden(false);

		// Kiểm tra xem có bao nhiêu file mới		
		PRS = pSheet->GetRange(
			pRange->GetItem(irowStart, icotFilegoc), 
			pRange->GetItem(iRowLast, icotFilegoc));

		vector<CString> vecHyper;
		int ncoutHyp = ff->_xlGetAllHyperLink(PRS, vecHyper, 1, 1);
		int icheck = 0;

		vector<int> vecGRow;
		int ncount = (int)vecItem[nIndex].irow.size();
		for (int i = 0; i < ncount; i++)
		{
			icheck = 0;
			PRS = pRgActive->GetItem(vecItem[nIndex].irow[i], icotFilegoc);
			szval = ff->_xlGetHyperLink(PRS);
			if (szval != L"" && ncoutHyp > 0)
			{
				for (int j = 0; j < ncoutHyp; j++)
				{
					if (vecHyper[j].Find(szval) >= 0)
					{
						icheck = 1;
						break;
					}
				}
			}

			if (icheck == 0) vecGRow.push_back(vecItem[nIndex].irow[i]);
		}
		vecHyper.clear();

		PRS = pRange->GetItem(1, icotMH);
		PRS->EntireColumn->PutHidden(blHidden);

		int demNew = (int)vecGRow.size();
		if (demNew > 0)
		{
			PRS = pSheet->GetRange(pRange->GetItem(irowStart, 1), pRange->GetItem(irowStart + demNew - 1, 1));
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = pSheet->GetRange(pRange->GetItem(irowStart, 1), pRange->GetItem(irowStart + demNew - 1, 1));
			PRS->EntireRow->Font->PutName(FontTimes);
			PRS->EntireRow->Font->PutSize(13);
			PRS->EntireRow->Font->PutBold(false);
			PRS->EntireRow->Font->PutItalic(false);
			PRS->EntireRow->Font->PutColor(rgbBlack);
			PRS->EntireRow->Interior->PutColor(xlNone);

			if (jcotNoidung == 0) jcotNoidung = jcotTen;
			int nColCenter[] = { jcotTen, jcotNoidung, icotThLuan + 2, icotThLuan + 3, icotThLuan + 4 };
			int nsizearr = getSizeArray(nColCenter);
			for (int i = 0; i < nsizearr; i++)
			{
				PRS = pSheet->GetRange(pRange->GetItem(irowStart, nColCenter[i]), pRange->GetItem(irowStart + demNew - 1, nColCenter[i]));
				PRS->EntireColumn->PutHorizontalAlignment(xlLeft);
			}

			int numStart = 1;
			for (int i = 0; i < demNew; i++)
			{
				dem++;
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / tongCopy), L"%");
				ff->_xlPutStatus(szval);

				for (int j = 1; j <= icotEnd; j++)
				{
					szval = ff->_xlGIOR(pRgActive, vecGRow[i], j, DF_NULL);
					pRange->PutItem(irowStart + i, j, (_bstr_t)szval);

					if (j == icotFilegoc || j == icotFileGXD || j == icotThLuan)
					{
						szDisplay = HypDisplayOpen;
						if (j == icotThLuan) szDisplay = HypDisplayGoto;

						PRS = pRgActive->GetItem(vecGRow[i], j);
						szval = ff->_xlGetHyperLink(PRS);
						if (szval != L"")
						{
							PRS = pRange->GetItem(irowStart + i, j);
							ff->_xlSetHyperlink(pShActive, PRS, szval, szDisplay);
						}
					}
				}

				pRange->PutItem(irowStart + i, jcotSTT, numStart);
				numStart++;
			}

			// Đánh lại STT
			_AutoSothutu(pSheet, numStart, irowStart + demNew, 0);

			PRS = pSheet->GetRange(pRange->GetItem(1, 1), pRange->GetItem(1, icotThLuan));
			PRS->PutHorizontalAlignment(xlCenterAcrossSelection);

			PRS = pSheet->GetRange(pRange->GetItem(irowTieude, 1), pRange->GetItem(irowStart - 1, 1));
			PRS->EntireRow->PutHorizontalAlignment(xlCenter);

			PRS = pSheet->GetRange(pRange->GetItem(irowStart, jcotSTT),
				pRange->GetItem(irowStart + demNew - 1, icotEnd));
			ff->_xlFormatBorders(PRS, 3, true, true);
			PRS->Rows->AutoFit();

			if (demNew < ncount)
			{
				for (int i = demNew; i < ncount; i++)
				{
					dem++;
					szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / tongCopy), L"%");
					ff->_xlPutStatus(szval);
				}
			}			
		}

		PRS = pRange->GetItem(irowStart, 1);
		PRS->Select();
		
		vecGRow.clear();
		return true;
	}
	catch (...) {}
	return false;
}

void DlgCopyData::_AutoSothutu(_WorksheetPtr pSheet, int numStart, int nRowStart, int nRowEnd)
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
