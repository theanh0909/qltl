#include "pch.h"
#include "DlgRibbonTimkiem.h"

IMPLEMENT_DYNAMIC(DlgRibbonTimkiem, CDialogEx)

DlgRibbonTimkiem::DlgRibbonTimkiem(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgRibbonTimkiem::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	iCheckCol = 0, iCheckWbook = 0;
	szTukhoa = L"";

	colSheet = 1, colCell = 2, colNoidung = 3;
	iwSHE = 150, iwCEL = 80, iwND = 300;
	iDlgW = 0, iDlgH = 0;

	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 50;

	nItem = -1, nSubItem = -1;
}

DlgRibbonTimkiem::~DlgRibbonTimkiem()
{
	delete ff;
	delete bb;
	delete reg;

	vecFilter.clear();
	vecSelect.clear();

	vnameSearch.clear();
	wpSheet.clear();
}

void DlgRibbonTimkiem::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, RBTK_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, RBTK_CHECK_COL, m_check_col);
	DDX_Control(pDX, RBTK_CHECK_WB, m_check_wb);
	DDX_Control(pDX, RBTK_BTN_SEARCH, m_btn_search);
	DDX_Control(pDX, RBTK_LIST_VIEW, m_list_view);
	DDX_Control(pDX, RBTK_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, RBTK_BTN_GOTOCELL, m_btn_goto);
}

BEGIN_MESSAGE_MAP(DlgRibbonTimkiem, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(RBTK_BTN_CANCEL, &DlgRibbonTimkiem::OnBnClickedBtnCancel)	
	ON_BN_CLICKED(RBTK_BTN_GOTOCELL, &DlgRibbonTimkiem::OnBnClickedBtnGotocell)
	ON_BN_CLICKED(RBTK_BTN_SEARCH, &DlgRibbonTimkiem::OnBnClickedBtnSearch)
	ON_COMMAND(MN_RBSEARCH_GOTO, &DlgRibbonTimkiem::OnRbsearchGoto)

	ON_NOTIFY(NM_CLICK, RBTK_LIST_VIEW, &DlgRibbonTimkiem::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, RBTK_LIST_VIEW, &DlgRibbonTimkiem::OnNMRClickListView)
	ON_NOTIFY(NM_DBLCLK, RBTK_LIST_VIEW, &DlgRibbonTimkiem::OnNMDblclkListView)	
	
	ON_NOTIFY(LVN_ENDSCROLL, RBTK_LIST_VIEW, &DlgRibbonTimkiem::OnLvnEndScrollListView)
	ON_NOTIFY(LVN_KEYDOWN, RBTK_LIST_VIEW, &DlgRibbonTimkiem::OnLvnKeydownListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgRibbonTimkiem::OnHdnEndtrackListView)
	ON_BN_CLICKED(RBTK_CHECK_COL, &DlgRibbonTimkiem::OnBnClickedCheckCol)
	ON_BN_CLICKED(RBTK_CHECK_WB, &DlgRibbonTimkiem::OnBnClickedCheckWb)
	ON_COMMAND(MN_RBSEARCH_OPENFILE, &DlgRibbonTimkiem::OnRbsearchOpenfile)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgRibbonTimkiem)
	EASYSIZE(RBTK_TXT_SEARCH, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RBTK_BTN_SEARCH, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RBTK_CHECK_COL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(RBTK_CHECK_WB, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(RBTK_GRP_KQ, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(RBTK_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)	
	EASYSIZE(RBTK_BTN_GOTOCELL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(RBTK_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgRibbonTimkiem message handlers
BOOL DlgRibbonTimkiem::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_SEARCH));
	SetIcon(hIcon, FALSE);

	_LoadRegistry();
	_SetDefault();

	mnIcon.GdiplusStartupInputConfig();

	_LoadListView();
	
	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgRibbonTimkiem::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(RBTK_LIST_VIEW) 
				|| pWndWithFocus == GetDlgItem(RBTK_BTN_GOTOCELL))
			{
				OnBnClickedBtnGotocell();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(RBTK_TXT_SEARCH))
			{
				OnBnClickedBtnSearch();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(RBTK_BTN_CANCEL))
			{
				OnBnClickedBtnCancel();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(RBTK_TXT_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(RBTK_LIST_VIEW));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(RBTK_LIST_VIEW))
			{
				GotoDlgCtrl(GetDlgItem(RBTK_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			OnBnClickedBtnCancel();
			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(RBTK_LIST_VIEW))
		{
			if (chr == 0x01)
			{
				_SelectAllItems();
				return TRUE;
			}
		}
	}

	return FALSE;
}

void DlgRibbonTimkiem::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgRibbonTimkiem::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgRibbonTimkiem::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgRibbonTimkiem::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_RBSEARCH));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RBSEARCH_GOTO, IDI_ICON_ARROW);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RBSEARCH_OPENFILE, IDI_ICON_OPENFILE);

			vector<int> vecRow;
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_RBSEARCH_GOTO, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_RBSEARCH_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(RBTK_LIST_VIEW));
			pListview->GetWindowRect(&rectSubmitButton);
			if (rectSubmitButton.PtInRect(point))
			{
				TrackPopupMenu(pMenu, TPM_LEFTALIGN | TPM_TOPALIGN, point.x, point.y, 0, this->GetSafeHwnd(), NULL);
			}

			if (hBmp) ::DeleteObject(hBmp);
		}

		if (hMMenu) VERIFY(DestroyMenu(hMMenu));
	}
	catch (...) {}
}

/**********************************************************************************/

void DlgRibbonTimkiem::_Timkiemdulieu()
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		CString szval = L"";
		CString szUpdateStatus = L"Đang tìm kiếm dữ liệu. Vui lòng đợi trong giây lát...";
		ff->_xlPutStatus(szUpdateStatus);

		_XacdinhSheetLienquan();

		iCheckWbook = _wtoi(ff->_xlGetValue(psConfig, L"CF_CHECK_WORKBOOK", 1, 0));
		if (iCheckWbook != 0) iCheckWbook = 1;

		vnameSearch.clear();
		iCheckCol = _wtoi(ff->_xlGetValue(psConfig, L"CF_CHECK_COLUMN", 1, 0));
		if (iCheckCol != 0) iCheckCol = 1;
		if (iCheckCol == 1) _GetNameSearch();
	
		vecSelect.clear();

		szTukhoa = ff->_xlGetValue(psConfig, L"CF_KEY_SEARCHBOX", 1, 0).Trim();
		if (szTukhoa != L"")
		{
			int iResult = 0;
			if (iCheckWbook == 0)
			{
				if (iCheckCol == 1) iResult = _GetCellItems(pShActive);
				else
				{
					iResult = ff->_xlGetAllCellSelection(1, 0);
					if (iResult <= 1) iResult = _GetCellItems(pShActive);
				}
			}
			else
			{
				wpSheet.clear();
				int nHidden = 0;
				int nSheet = ff->_xlGetAllSheet(wpSheet, nHidden);
				if (nSheet > 0)
				{
					for (int i = 0; i < nSheet; i++)
					{
						szval.Format(L"%s (%d%s) %s",
							szUpdateStatus, (int)(100 * (i + 1) / nSheet), L"%", (LPCTSTR)wpSheet[i]->Name);
						ff->_xlPutStatus(szval);

						iResult += _GetCellItems(wpSheet[i]);
					}
				}				
			}

			if (iResult > 0)
			{
				iResult = _FilterData();
				if (iResult == 1) _GotoCell(0);
				else if (iResult > 1)
				{
					// Hiển thị hộp thoại
					ff->_xlPutScreenUpdating(true);
					xl->EnableCancelKey = XlEnableCancelKey(FALSE);
					AFX_MANAGE_STATE(AfxGetStaticModuleState());
					DoModal();
				}
				else
				{
					ff->_msgbox(L"Không tìm thấy kết quả. Vui lòng nhập lại từ khóa phù hợp.", 0, 0);
				}
			}
		}

		ff->_xlPutStatus(L"Ready");
		ff->_xlPutScreenUpdating(true);
	}
	catch (...) {}
}

long DlgRibbonTimkiem::_GetCellItems(_WorksheetPtr pSheet)
{
	try
	{
		CString szName = (LPCTSTR)pSheet->Name;
		CString szCodeName = (LPCTSTR)pSheet->CodeName;
		CString szBlacklist[] = { L"shConfig" };
		
		int nsize = getSizeArray(szBlacklist);
		if (nsize > 0)
		{
			for (int i = 0; i < nsize; i++)
			{
				if (szCodeName == szBlacklist[i]) return 0;
			}
		}

		MyStrSelection MSSLC;
		RangePtr pRange = pSheet->Cells;
		RangePtr PRS;

		int pRowEnd = pRange->SpecialCells(xlCellTypeLastCell)->GetRow();
		int pColumnEnd = pRange->SpecialCells(xlCellTypeLastCell)->GetColumn();
		CString szval = L"";

		if (iCheckCol == 1)
		{
			// Giới hạn vùng tìm kiếm là cột có tiêu đề nội dung mô tả
			int ncountName = (int)vnameSearch.size();
			if (ncountName == 0) return 0;
			
			// Xác định vị trí chứa STT	<-- Nằm cùng dòng với các cột nội dung mô tả			
			CString szTitle[] = { L"stt", L"tt", L"id" };
			nsize = getSizeArray(szTitle);
			int iTitle = 0;

			for (int i = 1; i <= 10; i++)
			{
				szval = ff->_xlGIOR(pRange, i, 1);
				szval.MakeLower();
				for (int j = 0; j < nsize; j++)
				{
					if (szval == szTitle[j])
					{
						iTitle = i;
						break;
					}
				}

				if (iTitle > 0) break;
			}

			if (iTitle > 0)
			{
				// Tiếp tục tìm tiếp tại dòng tiêu đề xem cột nào chứa nội dung/mô tả
				vector<int> vcolContent;
				for (int i = 1; i <= pColumnEnd; i++)
				{
					szval = ff->_xlGIOR(pRange, iTitle, i);
					szval = ff->_ConvertKhongDau(szval);
					szval.MakeLower();

					for (int j = 0; j < ncountName; j++)
					{
						if (szval.Find(vnameSearch[j]) >= 0)
						{
							vcolContent.push_back(i);
							break;
						}
					}
				}

				nsize = (int)vcolContent.size();
				if (nsize > 0)
				{
					for (int i = 0; i < nsize; i++)
					{
						for (int j = 1; j <= pRowEnd; j++)
						{
							szval = ff->_xlGIOR(pRange, j, vcolContent[i], L"");
							if (szval != L"")
							{
								MSSLC.row = j;
								MSSLC.column = vcolContent[i];
								MSSLC.cell.Format(L"%s%d", ff->_xlConvertRCtoA1(vcolContent[i]), j);

								MSSLC.sheet = szName;
								MSSLC.bkgcolor = pSheet->Tab->Color;

								MSSLC.value = szval;
								PRS = pRange->GetItem(j, vcolContent[i]);
								MSSLC.formula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

								vecSelect.push_back(MSSLC);
							}
						}
					}
				}
			}
		}
		else
		{
			for (int i = 1; i <= pColumnEnd; i++)
			{
				for (int j = 1; j <= pRowEnd; j++)
				{
					szval = ff->_xlGIOR(pRange, j, i, L"");
					if (szval != L"")
					{
						MSSLC.row = j;
						MSSLC.column = i;
						MSSLC.cell.Format(L"%s%d", ff->_xlConvertRCtoA1(i), j);

						MSSLC.sheet = szName;
						MSSLC.bkgcolor = pSheet->Tab->Color;

						MSSLC.value = szval;
						PRS = pRange->GetItem(j, i);
						MSSLC.formula = (LPCTSTR)(_bstr_t)PRS->GetFormula();

						vecSelect.push_back(MSSLC);
					}
				}
			}
		}		

		return (long)vecSelect.size();
	}
	catch (...) {}
	return 0;
}


void DlgRibbonTimkiem::_GotoCell(int nIndex)
{
	try
	{
		_WorksheetPtr pSheet = xl->Sheets->GetItem((_bstr_t)vecFilter[nIndex].szSheet);
		if ((int)pSheet->GetVisible() != -1) pSheet->PutVisible(XlSheetVisibility::xlSheetVisible);
		pSheet->Activate();

		int ir = vecFilter[nIndex].iRow;
		int ic = vecFilter[nIndex].iColumn;

		RangePtr pRange = pSheet->Cells;
		RangePtr PRS = pRange->GetItem(ir, ic);
		
		bool blHidden = PRS->EntireRow->GetHidden();
		if (blHidden == true) PRS->EntireRow->PutHidden(false);

		PRS->Select();

		if (ir > 1) ir--;
		if (ic > 2) ic -= 2;

		xl->ActiveWindow->PutScrollRow(ir);
		xl->ActiveWindow->PutScrollColumn(ic);
	}
	catch (...) {}
}

int DlgRibbonTimkiem::_FilterData()
{
	try
	{
		vecFilter.clear();

		CString szval = szTukhoa;
		szval = ff->_ConvertKhongDau(szval);
		szval.MakeLower();

		vector<CString> vecKey;
		int iKey = ff->_TackTukhoa(vecKey, szval, L" ", L"+", 1, 1);

		int icheck = 0;
		long itotal = (long)vecSelect.size();

		MyStrList MSList;
		for (long i = 0; i < itotal; i++)
		{
			icheck = 0;
			szval = ff->_ConvertKhongDau(vecSelect[i].value);
			szval.MakeLower();

			for (int j = 0; j < iKey; j++)
			{
				if (szval.Find(vecKey[j]) >= 0) icheck++;
				else break;
			}

			if (icheck == iKey)
			{
				MSList.szNoidung = vecSelect[i].value;
				MSList.szCell = vecSelect[i].cell;

				MSList.szSheet = vecSelect[i].sheet;
				MSList.bkgTabColor = vecSelect[i].bkgcolor;

				MSList.iColumn = vecSelect[i].column;
				MSList.iRow = vecSelect[i].row;

				MSList.iEnable = 1;

				vecFilter.push_back(MSList);
			}
		}

		vecKey.clear();
		return (int)vecFilter.size();
	}
	catch (...) {}
	return 0;
}


void DlgRibbonTimkiem::_SetDefault()
{
	m_check_col.SetCheck(iCheckCol);
	m_check_wb.SetCheck(iCheckWbook);

	m_txt_search.SubclassDlgItem(RBTK_TXT_SEARCH, this);
	m_txt_search.SetBkColor(rgbWhite);
	m_txt_search.SetTextColor(rgbBlue);
	m_txt_search.SetWindowText(szTukhoa);

	m_btn_search.SetIcon(IDI_ICON_SEARCH, 24, 24);
	m_btn_search.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_search.DrawTransparent(true);
	m_btn_search.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_goto.SetThemeHelper(&m_ThemeHelper);
	m_btn_goto.SetIcon(IDI_ICON_ARROW, 24, 24);
	m_btn_goto.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_goto.SetBtnCursor(IDC_CURSOR_HAND);

	m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_view.InsertColumn(colSheet, L" Sheet", LVCFMT_LEFT, iwSHE);
	m_list_view.InsertColumn(colCell, L" Cell", LVCFMT_CENTER, iwCEL);
	m_list_view.InsertColumn(colNoidung, L" Giá trị", LVCFMT_LEFT, iwND);

	m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
	m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
}

void DlgRibbonTimkiem::OnBnClickedBtnCancel()
{
	_SaveRegistry();

	CDialog::OnCancel();
}

void DlgRibbonTimkiem::_SetLocationAndSize()
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

void DlgRibbonTimkiem::_DeleteAllList()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

void DlgRibbonTimkiem::_XacdinhSheetLienquan()
{
	bsConfig = ff->_xlGetNameSheet(shConfig, 0);
	psConfig = xl->Sheets->GetItem(bsConfig);
	prConfig = psConfig->Cells;
}

void DlgRibbonTimkiem::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + RibbonSearch;
		reg->CreateKey(szCreate);

		CString szval = L"";
		iwSHE = m_list_view.GetColumnWidth(colSheet);
		iwCEL = m_list_view.GetColumnWidth(colCell);
		iwND = m_list_view.GetColumnWidth(colNoidung);
		szval.Format(L"%d|%d|%d", iwSHE, iwCEL, iwND);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgRibbonTimkiem::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + RibbonSearch;
		reg->CreateKey(szCreate);

		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			vector< CString> vecKey;
			int iKey = ff->_TackTukhoa(vecKey, szval, L"|", L";", 0, 1);
			if (iKey < 3) for (int i = iKey; i < 3; i++) vecKey.push_back(L"");

			iwSHE	= _wtoi(vecKey[0]);
			if (iwSHE > 200) iwSHE = 200;

			iwCEL = _wtoi(vecKey[1]);
			if (iwCEL > 120) iwCEL = 120;

			iwND = _wtoi(vecKey[2]);
			if (iwND < 300) iwND = 300;
		}

		int pos = -1;
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

int DlgRibbonTimkiem::_LoadListView()
{
	try
	{
		_DeleteAllList();

		tslkq = (int)vecFilter.size();

		lanshow = 1;
		int iz = ibuocnhay;
		if (tslkq <= iz)
		{
			iz = tslkq;
			iStopload = 0;
		}
		else iStopload = 1;

		_AddItemsListKetqua(0, iz);		// Thêm dữ liệu vào list kết quả

		if (tslkq > 0) GotoDlgCtrl(GetDlgItem(RBTK_LIST_VIEW));
		else GotoDlgCtrl(GetDlgItem(RBTK_TXT_SEARCH));

		CString szval = L"";
		szval.Format(L"Danh sách hiển thị (%d kết quả)", tslkq);
		CStatic *m_grp_ketqua = (CStatic*)GetDlgItem(RBTK_GRP_KQ);
		m_grp_ketqua->SetWindowText(szval);

		return iz;
	}
	catch (...) {}
	return 0;
}

void DlgRibbonTimkiem::_AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, colSheet, vecFilter[i].szSheet);
			m_list_view.SetItemText(i, colCell, vecFilter[i].szCell);
			m_list_view.SetItemText(i, colNoidung, vecFilter[i].szNoidung);
			
			if (vecFilter[i].bkgTabColor != rgbBlack)
			{
				m_list_view.SetCellColors(i, 0, vecFilter[i].bkgTabColor, rgbBlack);
			}
		}
	}
	catch (...) {}
}

void DlgRibbonTimkiem::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz);
		}
	}
	catch (...) {}
}

void DlgRibbonTimkiem::OnBnClickedBtnSearch()
{
	try
	{
		m_txt_search.GetWindowTextW(szTukhoa);
		szTukhoa.Trim();

		if (szTukhoa != L"")
		{
			_FilterData();
			_LoadListView();
		}
	}
	catch (...) {}
}

void DlgRibbonTimkiem::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0) 
	{
		_GotoCell(nItem);
		OnBnClickedBtnCancel();
	}
}

void DlgRibbonTimkiem::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
}

void DlgRibbonTimkiem::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
}

void DlgRibbonTimkiem::OnRbsearchGoto()
{
	int ncount = m_list_view.GetItemCount();
	if (ncount > 0)
	{
		if (nItem == -1) nItem = 0;
		
		_GotoCell(nItem);
		OnBnClickedBtnCancel();
	}	
}

int DlgRibbonTimkiem::_GetAllSelectedItems(vector<int> &vecRow)
{
	try
	{
		vecRow.clear();
		if (nItem == -1 || nSubItem == -1) return 0;

		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vecRow.push_back(i);
				}
			}

			int nsize = (int)vecRow.size();
			if (nsize == ncount)
			{
				for (int i = ncount; i < tslkq; i++) vecRow.push_back(i);
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void DlgRibbonTimkiem::_SelectAllItems()
{
	int ncount = m_list_view.GetItemCount();
	if (ncount > 0)
	{
		for (int i = 0; i < ncount; i++)
		{
			m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
		}
	}
}

void DlgRibbonTimkiem::OnBnClickedBtnGotocell()
{
	OnRbsearchGoto();
}

void DlgRibbonTimkiem::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_SPACE) OnBnClickedBtnGotocell();
	*pResult = 0;
}


void DlgRibbonTimkiem::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = COLW_FIX;
	*pResult = 0;
}

void DlgRibbonTimkiem::OnBnClickedCheckCol()
{
	_ChangeSearch();
}

void DlgRibbonTimkiem::OnBnClickedCheckWb()
{
	_ChangeSearch();
}

int DlgRibbonTimkiem::_GetNameSearch()
{
	try
	{
		int iRowCF = 0;
		int iColCF = ff->_xlGetColumn(psConfig, L"CF_COLUMN_SEARCH", -1);
		if (iColCF > 0) iRowCF = ff->_xlGetRow(psConfig, L"CF_COLUMN_SEARCH", 0);
		if (iRowCF > 0)
		{
			CString szval = L"";
			for (int i = iRowCF + 1; i < iRowCF + 100; i++)
			{
				szval = ff->_xlGIOR(prConfig, i, iColCF, L"");
				if (szval != L"")
				{
					szval = ff->_ConvertKhongDau(szval);
					szval.MakeLower();
					vnameSearch.push_back(szval);
				}
				else break;
			}
		}

		if ((int)vnameSearch.size() == 0)
		{
			vnameSearch.push_back(L"Nội dung");
			vnameSearch.push_back(L"Mô tả");
			vnameSearch.push_back(L"Tên file");
			vnameSearch.push_back(L"Danh mục");
			vnameSearch.push_back(L"Tên chi phí");
			vnameSearch.push_back(L"Căn cứ xác định");
			vnameSearch.push_back(L"Ghi chú");
		}
	}
	catch (...) {}
	return 0;
}


void DlgRibbonTimkiem::_ChangeSearch()
{
	try
	{
		vecSelect.clear();

		if ((int)vnameSearch.size() == 0) _GetNameSearch();

		iCheckWbook = m_check_wb.GetCheck();
		iCheckCol = m_check_col.GetCheck();		

		if (iCheckWbook == 0)
		{
			if (iCheckCol == 1) _GetCellItems(pShActive);
			else
			{
				int iResult = ff->_xlGetAllCellSelection(1, 0);
				if (iResult <= 1) iResult = _GetCellItems(pShActive);
			}
		}
		else
		{
			wpSheet.clear();
			int nHidden = 0;
			int nSheet = (int)wpSheet.size();
			if (nSheet == 0) nSheet = ff->_xlGetAllSheet(wpSheet, nHidden);
			if (nSheet > 0)
			{
				for (int i = 0; i < nSheet; i++) _GetCellItems(wpSheet[i]);
			}
		}

		OnBnClickedBtnSearch();
	}
	catch (...) {}
}

void DlgRibbonTimkiem::OnRbsearchOpenfile()
{
	try
	{
		if (nItem == -1) return;
		
		_WorksheetPtr pSheet = xl->Sheets->GetItem((_bstr_t)vecFilter[nItem].szSheet);

		// Hàm hỗ trợ mở file đặc biệt
		ff->_xlOpenFileEx(pSheet, vecFilter[nItem].iRow, vecFilter[nItem].iColumn);
	}
	catch (...) {}
}
