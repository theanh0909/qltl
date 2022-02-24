#include "pch.h"
#include "GHookForm.h"
#include "FrmQuytrinh.h"
#include "HookMouse.h"

// FrmQuytrinh dialog
IMPLEMENT_DYNAMIC(FrmQuytrinh, CDialogEx)

FrmQuytrinh::FrmQuytrinh(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FRM_QUYTRINH, pParent)
{
	szTukhoa = L"";
	szpathDefault = L"";

	blChange = false;

	demTimer = 0;
	icotID = 0, icotQtrinh = 1, icotName = 2;
	wcID = 0, wcQTrinh = 720, wcTen = 0;

	iUnikey = 0;
	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 20;

	icolSTT = 0, iColMaso = 0, iColNdung = 0, iColChidan = 0;
	iColURL = 0, iColLienhe = 0, iColGChu = 0, iColType = 0;

	nIndexMouse = 1;
	MyHook::Instance().InstallHook();
}

FrmQuytrinh::~FrmQuytrinh()
{
	vecQuytrinh.clear();
	vecFilter.clear();

	nIndexMouse = 0;
	MyHook::Instance().UninstallHook();
}


void FrmQuytrinh::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, QTRINH_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, QTRINH_LIST_SEARCH, m_list_view);
}


BEGIN_MESSAGE_MAP(FrmQuytrinh, CDialogEx)
	ON_WM_TIMER()
	ON_WM_ERASEBKGND()
	ON_WM_SYSCOMMAND()
	ON_NOTIFY(HDN_ENDTRACK, 0, &FrmQuytrinh::OnHdnEndtrackListView)
	ON_EN_CHANGE(QTRINH_TXT_SEARCH, &FrmQuytrinh::OnEnChangeTxtSearch)
	ON_NOTIFY(NM_DBLCLK, QTRINH_LIST_SEARCH, &FrmQuytrinh::OnNMDblclkListSearch)
	ON_NOTIFY(LVN_ENDSCROLL, QTRINH_LIST_SEARCH, &FrmQuytrinh::OnLvnEndScrollListSearch)
END_MESSAGE_MAP()

// FrmQuytrinh message handlers
BOOL FrmQuytrinh::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	//SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	//SetLayeredWindowAttributes(RGB(255, 255, 255), 0, LWA_COLORKEY);

	hWndQuytrinh = m_hWnd;

	SetDefault();
	SetLocation();
	ReadFileQuytrinh();

	if (szTukhoa != L"") SetTimerSearch();
	else TimkiemQuytrinh();

	blChange = true;
	iUnikey++;

	return TRUE;
}

BOOL FrmQuytrinh::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(QTRINH_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(QTRINH_TXT_SEARCH))
			{
				if ((int)m_list_view.GetItemCount() == 0) PutKey();
				else
				{
					SelectListItem(true);
					return TRUE;
				}
			}
		}
		else if (iWPar == VK_SPACE)
		{
			if (pWndWithFocus == GetDlgItem(QTRINH_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(QTRINH_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(QTRINH_LIST_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(QTRINH_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_DOWN)
		{
			if (pWndWithFocus == GetDlgItem(QTRINH_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(QTRINH_LIST_SEARCH))
			{
				CheckArrowDownNext();
			}
		}
		else if (iWPar == VK_UP)
		{
			if (pWndWithFocus == GetDlgItem(QTRINH_LIST_SEARCH))
			{
				vector<int> vecItems;
				int ncount = GetSelectItems(vecItems);
				if (ncount == 1 && vecItems[0] == 0)
				{
					GotoDlgCtrl(GetDlgItem(QTRINH_TXT_SEARCH));
					return TRUE;
				}

				vecItems.clear();
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

void FrmQuytrinh::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

BOOL FrmQuytrinh::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(192, 192, 192));
	return FALSE;
}

void FrmQuytrinh::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;

	if (demTimer >= 2)
	{
		demTimer = 0;
		KillTimer(CONST_TIMER_QUYTRINH);

		TimkiemQuytrinh();
	}

	CDialogEx::OnTimer(nIDEvent);
}

void FrmQuytrinh::SetTimerSearch()
{
	demTimer = 0;
	SetTimer(CONST_TIMER_QUYTRINH, 100, NULL);
}

void FrmQuytrinh::PutKey()
{
	try
	{
		nIndexMouse = 0;
		//MyHook::Instance().UninstallHook();
		CDialogEx::OnOK();

		CString szval = L"";
		m_txt_search.GetWindowTextW(szval);
		pRgActive->PutItem(iRowActive, iColumnActive, (_bstr_t)szval);
	}
	catch (...) {}
}

void FrmQuytrinh::OnBnClickedBtnOk()
{
	try
	{
		if ((int)m_list_view.GetItemCount() == 0)
		{
			PutKey();
			return;
		}

		/********************************************************************************************/

		vector<int> vecItems;
		int nSelect = GetSelectItems(vecItems);
		
		// Giới hạn chỉ cho phép chọn từng quy trình
		if (nSelect > 1) nSelect = 1;

		if (nSelect > 0)
		{
			nIndexMouse = 0;
			//MyHook::Instance().UninstallHook();
			CDialogEx::OnOK();

			try
			{
				xl->PutScreenUpdating(false);
				xl->PutDisplayAlerts(false);

				// Xác định các cột dữ liệu (bắt đầu --> kết thúc)
				icolSTT = _xlGetColumn(pShActive, L"QTR_STT", -1);
				iColMaso = _xlGetColumn(pShActive, L"QTR_MASO", -1);
				iColNdung = _xlGetColumn(pShActive, L"QTR_NOIDUNG", -1);
				iColChidan = _xlGetColumn(pShActive, L"QTR_CHIDAN", -1);
				iColURL = _xlGetColumn(pShActive, L"QTR_URL", -1);
				iColLienhe = _xlGetColumn(pShActive, L"QTR_LIENHE", -1);
				iColGChu = _xlGetColumn(pShActive, L"QTR_GCHU", -1);
				iColType = _xlGetColumn(pShActive, L"QTR_TYPE", -1);

				if (icolSTT > 0 && iColMaso > 0 && iColLienhe > 0 && iColType > 0)
				{
					int nID = _wtoi(m_list_view.GetItemText(vecItems[0], icotID));
					CopyQuytrinh(nID);
				}

				xl->PutDisplayAlerts(true);
				xl->PutScreenUpdating(true);
			}
			catch (...) {}
		}

		vecItems.clear();
	}
	catch (...) {}
}

void FrmQuytrinh::OnBnClickedBtnCancel()
{
	PutKey();
	CDialogEx::OnCancel();

	CSendKeys *m_sk = new CSendKeys;
	m_sk->SendKeys(L"{UP}");
	delete m_sk;
}

void FrmQuytrinh::SetDefault()
{
	try
	{
		m_txt_search.SubclassDlgItem(QTRINH_TXT_SEARCH, this);
		m_txt_search.SetBkColor(RGB(255, 255, 255));
		m_txt_search.SetTextColor(RGB(0, 0, 255));
		m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm quy trình...");

		m_txt_search.SetWindowText(szTukhoa);
		m_txt_search.PostMessage(EM_SETSEL, -1);

		m_list_view.InsertColumn(icotID, L"", LVCFMT_CENTER, wcID);
		m_list_view.InsertColumn(icotQtrinh, L"Quy trình", LVCFMT_LEFT, wcQTrinh);
		m_list_view.InsertColumn(icotName, L"", LVCFMT_CENTER, wcTen);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(
			LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);
	}
	catch (...) {}
}

void FrmQuytrinh::SetLocation()
{
	try
	{
		int iX = 0, iY = 0;
		RangePtr pCell = pRgActive->GetItem(iRowActive + 1, iColumnActive);
		bool blFlag = _xlGetPointWindowsOfCell(pCell, iX, iY);
		
		if (blFlag)
		{
			// Vị trí ô cell active
			dPosX = iX;
			dPosY = iY;
		}
		else
		{
			POINT p;
			if (GetCursorPos(&p))
			{
				// Vị trí trỏ chuột
				dPosX = p.x;
				dPosY = p.y;
			}
		}

		if (dPosX > 0 && dPosY > 0)
		{
			CRect r;
			GetWindowRect(&r);
			ScreenToClient(&r);

			// Kích thước hộp thoại
			dWidthF = r.Width();
			dHeightF = r.Height();

			if (!blFlag)
			{
				// Xác định kích thước màn hình
				int isx = GetSystemMetrics(SM_CXSCREEN);
				int isy = GetSystemMetrics(SM_CYSCREEN);

				// Xác định lại kích thước màn hình khi bỏ qua Taskpanel
				int iTaskHeight = 0;
				int nPostion = IsTaskbarWndVisible(iTaskHeight);
				if (nPostion == 0 && iTaskHeight > 0) isy = iTaskHeight;

				// Xác định lại vị trí trỏ chuột nếu hộp thoại tràn màn hình
				if (dPosX + dWidthF + 20 > isx) dPosX -= (dWidthF + 30);
				else dPosX += 50;

				if (dPosY + (int)(2 * dHeightF / 3) > isy) dPosY = isy - dHeightF;
				else dPosY -= ((int)(dHeightF / 3));
			}

			MoveWindow(dPosX, dPosY, r.right - r.left, r.bottom - r.top, TRUE);
		}
	}
	catch (...) {}
}

void FrmQuytrinh::ReplaceTukhoa(CString &szKey)
{
	szKey.Replace(L" ", L"");
	szKey.Replace(L"-", L"");
	szKey.Replace(L":", L"");
	szKey.Replace(L"/", L"");
	szKey.Replace(L"(", L"");
	szKey.Replace(L")", L"");
}

int FrmQuytrinh::LocDulieuQuytrinh()
{
	try
	{
		CString szval = szTukhoa;
		szval = ConvertKhongDau(szval);
		szval.MakeLower();

		vector<CString> vecKey;
		int iKey = TackTukhoa(vecKey, szval, L" ", L"+", 1, 1);
		if (iKey == 1) ReplaceTukhoa(vecKey[0]);

		int icheck = 0;
		long itotal = (long)vecQuytrinh.size();
		for (long i = 0; i < itotal; i++)
		{
			icheck = 0;
			szval = ConvertKhongDau(vecQuytrinh[i].szTimkiem);
			szval.MakeLower();

			for (int j = 0; j < iKey; j++)
			{
				if (szval.Find(vecKey[j]) >= 0) icheck++;
				else break;
			}

			if (icheck == iKey)
			{
				vecFilter.push_back(vecQuytrinh[i]);
			}
		}

		vecKey.clear();
		return (int)vecFilter.size();
	}
	catch (...) {}
	return 0;
}


void FrmQuytrinh::TimkiemQuytrinh()
{
	try
	{
		vecFilter.clear();
		if ((int)vecQuytrinh.size() > 0)
		{
			m_txt_search.GetWindowTextW(szTukhoa);
			szTukhoa.Trim();

			if (szTukhoa != L"") LocDulieuQuytrinh();
			else vecFilter = vecQuytrinh;
		}

		LoadListView();
	}
	catch (...) {}
}

void FrmQuytrinh::DeleteAllList()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

int FrmQuytrinh::LoadListView()
{
	try
	{
		tslkq = (int)vecFilter.size();

		DeleteAllList();

		lanshow = 1;
		int iz = ibuocnhay;
		if (tslkq <= iz)
		{
			iz = tslkq;
			iStopload = 0;
		}
		else iStopload = 1;

		AddItemsListKetqua(0, iz);		// Thêm dữ liệu vào list kết quả

		CString szval = L"";
		szval.Format(L"Quy trình (%d kết quả)", tslkq);
		NameColumnHeading(m_list_view, icotQtrinh, 1, szval);

		return iz;
	}
	catch (...) {}
	return 0;
}

void FrmQuytrinh::AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			szval.Format(L"%d", vecFilter[i].nID);
			m_list_view.SetItemText(i, icotID, szval);
			m_list_view.SetItemText(i, icotQtrinh, vecFilter[i].szQuytrinh);
			m_list_view.SetItemText(i, icotName, vecFilter[i].szTenFile);

			if (i % 2 == 0) m_list_view.SetRowColors(i, RGB(245, 255, 135), RGB(0, 0, 0));
		}
	}
	catch (...) {}
}

int FrmQuytrinh::ReadFileQuytrinh()
{
	try
	{
		vecQuytrinh.clear();
		szpathDefault = GetPathFolder(FILENAMEDLL);

		CString szQuytrinhMDB = szpathDefault + FolderQuytrinh + FileQuytrinhMDB;
		if (FileExistsUnicode(szQuytrinhMDB) == 1)
		{
			ObjConn.OpenAccessDB(szQuytrinhMDB, L"", L"");

			CString szval = L"";
			MyStrQuytrinh MSTC;
			SqlString = L"SELECT * FROM tblQuytrinh ORDER BY Maso ASC;";
			CRecordset *Recset = ObjConn.Execute(SqlString);
			while (!Recset->IsEOF())
			{
				Recset->GetFieldValue(L"Maso", MSTC.szMa);
				Recset->GetFieldValue(L"TenQuytrinh", MSTC.szQuytrinh);
				Recset->GetFieldValue(L"TenFile", MSTC.szTenFile);

				if (MSTC.szMa != L"")
				{
					MSTC.szTimkiem.Format(L"%s|%s|%s",
						MSTC.szMa, MSTC.szQuytrinh, MSTC.szTenFile);

					MSTC.nID = (int)vecQuytrinh.size();
					vecQuytrinh.push_back(MSTC);
				}

				Recset->MoveNext();
			}
			Recset->Close();

			ObjConn.CloseDatabase();
		}

		return (int)vecQuytrinh.size();
	}
	catch (...) {}
	return 0;
}

void FrmQuytrinh::OnLvnEndScrollListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	ShowResultNext();
}

void FrmQuytrinh::OnEnChangeTxtSearch()
{
	if (iUnikey == 1)
	{
		CString szval = L"";
		m_txt_search.GetWindowTextW(szval);
		if (szval.GetLength() == 2)
		{
			m_txt_search.SetWindowText(ConvertUnikey(szval));
			m_txt_search.SetSel(-1);
		}

		iUnikey++;
		if (iUnikey > 2) iUnikey = 2;
	}

	if (blChange) SetTimerSearch();
}

int FrmQuytrinh::GetSelectItems(vector<int> &vecItem)
{
	try
	{
		vecItem.clear();

		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vecItem.push_back(i);
				}
			}
		}

		return (int)vecItem.size();
	}
	catch (...) {}
	return 0;
}

void FrmQuytrinh::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == icotID) pNMListHeader->pitem->cxy = 0;
	*pResult = 0;
}

void FrmQuytrinh::OnNMDblclkListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedBtnOk();
}

void FrmQuytrinh::SelectListItem(bool blEnter)
{
	GotoDlgCtrl(GetDlgItem(QTRINH_LIST_SEARCH));

	int ncount = m_list_view.GetItemCount();
	if (ncount > 0)
	{
		int selected = 0;
		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			while (pos)
			{
				selected = m_list_view.GetNextSelectedItem(pos);
				break;
			}
		}

		m_list_view.SetItemState(selected, LVIS_SELECTED, LVIS_SELECTED);

		if (blEnter && ncount == 1)
		{
			OnBnClickedBtnOk();
		}
	}
}

bool FrmQuytrinh::CheckArrowDownNext()
{
	try
	{
		int ncount = m_list_view.GetItemCount();
		if (ncount > 0)
		{
			int selected = -1;
			POSITION pos = m_list_view.GetFirstSelectedItemPosition();
			if (pos != NULL)
			{
				while (pos)
				{
					selected = m_list_view.GetNextSelectedItem(pos);
				}
			}

			if (selected == ncount - 1)
			{
				ShowResultNext();
				return true;
			}
		}
	}
	catch (...) {}
	return false;
}

void FrmQuytrinh::ShowResultNext()
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
			AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz);
		}
	}
	catch (...) {}
}

void FrmQuytrinh::CopyQuytrinh(int nID)
{
	try
	{
		// Kiểm tra file có tồn tại hay không?
		CString szTenFile = vecQuytrinh[nID].szTenFile;

		CString szFullXLSX = szpathDefault + FolderQuytrinh + szTenFile;
		if (FileExistsUnicode(szFullXLSX) != 1)
		{
			_msgbox(L"Không tồn tại đường dẫn chứa nội dung quy trình được chọn. "
				L"Vui lòng kiểm tra lại. \nPath= " + szFullXLSX, 0, 2);
			return;
		}

		_variant_t varOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

		xl->PutStatusBar(L"Đang tìm kiếm dữ liệu quy trình tương ứng. Vui lòng đợi trong giây lát...");

		_msgbox(L"1", 0, 0);
		_ApplicationPtr xlImport;
		if (FAILED(xlImport.CreateInstance(ExcelApplication)))
		{
			_msgbox(L"Không thể khởi tạo Excel.", 0, 2);
			return;
		}
		_WorkbookPtr pWbook;
		_msgbox(L"2", 0, 0);
		_msgbox(L"szFullXLSX= " + szFullXLSX, 0, 0);
		try
		{
			pWbook = xlImport->Workbooks->Open((_bstr_t)szFullXLSX,
				varOption, varOption, varOption, varOption, varOption, varOption,
				varOption, varOption, varOption, varOption, varOption, varOption);

		}
		catch (...) {
			_msgbox(L"#Error552", 0, 0);
		}
		

		if (pWbook == NULL)
		{
			_msgbox(L"Không tìm thấy Workbook tương ứng.", 0, 2);
			return;
		}
		_msgbox(L"3", 0, 0);
		int nIndex = 0;
		_WorksheetPtr pSh;
		int ncount = xlImport->ActiveWorkbook->Sheets->Count;
		for (int i = 1; i <= ncount; i++)
		{
			try
			{
				pSh = pWbook->Worksheets->GetItem(i);
				if ((int)pSh->GetVisible() != 2)
				{
					nIndex = i;
					break;
				}
			}
			catch (...) {}
		}
		_msgbox(L"4", 0, 0);
		if (nIndex == 0)
		{
			_msgbox(L"Không xác định đươc Worksheet tương ứng.", 0, 2);
			return;
		}
		_msgbox(L"5", 0, 0);
		_WorksheetPtr pSheet = pWbook->Sheets->Item[nIndex];
		if (pSheet == NULL)
		{
			_msgbox(L"Không tồn tại Worksheet tương ứng.", 0, 2);
			return;
		}
		_msgbox(L"6", 0, 0);
		RangePtr pRange = pSheet->Cells;
		if (pRange == NULL) return;
		_msgbox(L"7", 0, 0);
		xl->PutStatusBar(L"Đang tiến hành đọc dữ liệu quy trình. Vui lòng đợi trong giây lát...");

		CString szID = vecQuytrinh[nID].szMa;
		szID.MakeLower();

		int iRowBegin = _xlGetRow(pSheet, L"QTR_MASO", 0) + 1;
		int iRowEnd = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;

		int ncotNdung = _xlGetColumn(pSheet, L"QTR_NOIDUNG", -1);
		int ncotChidan = _xlGetColumn(pSheet, L"QTR_CHIDAN", -1);
		int ncotURL = _xlGetColumn(pSheet, L"QTR_URL", -1);
		int ncotLienhe = _xlGetColumn(pSheet, L"QTR_LIENHE", -1);
		int ncotGChu = _xlGetColumn(pSheet, L"QTR_GCHU", -1);
		int ncotType = _xlGetColumn(pSheet, L"QTR_TYPE", -1);

		CString szval = L"";
		int vtStart = 0, vtEnd = 0;
		for (int i = iRowBegin; i < iRowEnd; i++)
		{
			szval = _xlGIOR(pRange, i, iColMaso, L"");
			if (szval != L"")
			{
				szval.MakeLower();
				if (szval == szID)
				{
					vtStart = i;
					break;
				}
			}
		}

		if (vtStart > 0)
		{
			for (int i = vtStart + 1; i < iRowEnd; i++)
			{
				szval = _xlGIOR(pRange, i, iColMaso, L"");
				if (szval != L"")
				{
					vtEnd = i - 1;
					break;
				}
			}

			if (vtEnd == 0) vtEnd = iRowEnd - 1;
		}

		if (vtEnd > 0)
		{
			int slchen = vtEnd - vtStart + 1;

			RangePtr pCell = pShActive->GetRange(
				pRgActive->GetItem(iRowActive + 1, 1), 
				pRgActive->GetItem(iRowActive + slchen, 1));
			pCell->EntireRow->Insert(xlShiftDown);

			pCell = pSheet->GetRange(pRange->GetItem(vtStart, icolSTT),	pRange->GetItem(vtEnd, iColLienhe));
			pCell->Copy();

			RangePtr pRgPaste = pShActive->GetRange(
				pRgActive->GetItem(iRowActive, icolSTT),
				pRgActive->GetItem(iRowActive + slchen - 1, iColLienhe));
			pRgPaste->PasteSpecial(XlPasteType::xlPasteValues,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);
			pRgPaste->PasteSpecial(XlPasteType::xlPasteFormats,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);

			xlImport->PutCutCopyMode(XlCutCopyMode(false));

			/*********************************************************************/

			pCell = pSheet->GetRange(pRange->GetItem(vtStart, ncotGChu), pRange->GetItem(vtEnd, ncotGChu));
			pCell->Copy();

			pRgPaste = pShActive->GetRange(
				pRgActive->GetItem(iRowActive, iColGChu),
				pRgActive->GetItem(iRowActive + slchen - 1, iColGChu));
			pRgPaste->PasteSpecial(XlPasteType::xlPasteValues,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);
			pRgPaste->PasteSpecial(XlPasteType::xlPasteFormats,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);

			xlImport->PutCutCopyMode(XlCutCopyMode(false));

			/*********************************************************************/

			pCell = pSheet->GetRange(pRange->GetItem(vtStart, ncotType), pRange->GetItem(vtEnd, ncotType));
			pCell->Copy();

			pRgPaste = pShActive->GetRange(
				pRgActive->GetItem(iRowActive, iColType),
				pRgActive->GetItem(iRowActive + slchen - 1, iColType));
			pRgPaste->PasteSpecial(XlPasteType::xlPasteValues,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);
			pRgPaste->PasteSpecial(XlPasteType::xlPasteFormats,
				XlPasteSpecialOperation::xlPasteSpecialOperationNone);

			xlImport->PutCutCopyMode(XlCutCopyMode(false));

			/*********************************************************************/

			int icotGet[] = { ncotNdung, ncotChidan, ncotURL, ncotGChu };
			int icotPut[] = { iColNdung, iColChidan, iColURL, iColGChu };
			int iSizeArr = getSizeArray(icotPut);

			int dem = iRowActive;
			xl->PutStatusBar(L"Đang tiến hành kiểm tra liên kết. Vui lòng đợi trong giây lát...");
			for (int i = vtStart; i <= vtEnd; i++)
			{
				for (int j = 0; j < iSizeArr; j++)
				{
					if (icotGet[j] > 0 && icotPut[j] > 0)
					{
						szval = _xlGIOR(pRange, i, icotGet[j], L"");
						if (szval.GetLength() >= 255) pRgActive->PutItem(dem, icotPut[j], (_bstr_t)szval);
					}
				}

				dem++;
			}
			
			/*********************************************************************/

			pCell = pRgActive->GetItem(iRowActive, iColMaso);
			pCell->Activate();
			pCell->Select();
		}

		pWbook->Close(VARIANT_FALSE);
		xlImport->Quit();

		xl->PutStatusBar(L"Ready");
	}
	catch (...) {
		_msgbox(L"#Error566", 0, 0);
	}
}