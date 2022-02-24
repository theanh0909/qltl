#include "pch.h"
#include "GHookForm.h"
#include "FrmChiphi.h"
#include "HookMouse.h"

// FrmChiphi dialog
IMPLEMENT_DYNAMIC(FrmChiphi, CDialogEx)

FrmChiphi::FrmChiphi(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FRM_CHIPHI, pParent)
{
	szTukhoa = L"";
	szpathDefault = L"";

	blChange = false;

	demTimer = 0;
	icotID = 0, icotChiphi = 1, icotTable = 2;
	wcID = 0, wcChiphi = 720, wcTable = 0;

	iUnikey = 0;
	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 20;

	icolSTT = 0, iColNdung = 0, iColGChu = 0;

	nIndexMouse = 3;
	MyHook::Instance().InstallHook();
}

FrmChiphi::~FrmChiphi()
{
	vecChiphi.clear();
	vecFilter.clear();

	nIndexMouse = 0;
	MyHook::Instance().UninstallHook();
}


void FrmChiphi::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CHIPHI_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, CHIPHI_LIST_SEARCH, m_list_view);
}


BEGIN_MESSAGE_MAP(FrmChiphi, CDialogEx)
	ON_WM_TIMER()
	ON_WM_ERASEBKGND()
	ON_WM_SYSCOMMAND()
	ON_NOTIFY(HDN_ENDTRACK, 0, &FrmChiphi::OnHdnEndtrackListView)
	ON_EN_CHANGE(CHIPHI_TXT_SEARCH, &FrmChiphi::OnEnChangeTxtSearch)
	ON_NOTIFY(NM_DBLCLK, CHIPHI_LIST_SEARCH, &FrmChiphi::OnNMDblclkListSearch)
	ON_NOTIFY(LVN_ENDSCROLL, CHIPHI_LIST_SEARCH, &FrmChiphi::OnLvnEndScrollListSearch)
END_MESSAGE_MAP()

// FrmChiphi message handlers
BOOL FrmChiphi::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	//SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	//SetLayeredWindowAttributes(RGB(255, 255, 255), 0, LWA_COLORKEY);

	hWndChiphi = m_hWnd;

	SetDefault();
	SetLocation();
	ReadFileChiphi();

	if (szTukhoa != L"") SetTimerSearch();
	else TimkiemChiphi();

	blChange = true;
	iUnikey++;

	return TRUE;
}

BOOL FrmChiphi::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(CHIPHI_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(CHIPHI_TXT_SEARCH))
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
			if (pWndWithFocus == GetDlgItem(CHIPHI_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(CHIPHI_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(CHIPHI_LIST_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(CHIPHI_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_DOWN)
		{
			if (pWndWithFocus == GetDlgItem(CHIPHI_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(CHIPHI_LIST_SEARCH))
			{
				CheckArrowDownNext();
			}
		}
		else if (iWPar == VK_UP)
		{
			if (pWndWithFocus == GetDlgItem(CHIPHI_LIST_SEARCH))
			{
				vector<int> vecItems;
				int ncount = GetSelectItems(vecItems);
				if (ncount == 1 && vecItems[0] == 0)
				{
					GotoDlgCtrl(GetDlgItem(CHIPHI_TXT_SEARCH));
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

void FrmChiphi::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

BOOL FrmChiphi::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(192, 192, 192));
	return FALSE;
}

void FrmChiphi::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;

	if (demTimer >= 2)
	{
		demTimer = 0;
		KillTimer(CONST_TIMER_CHIPHI);

		TimkiemChiphi();
	}

	CDialogEx::OnTimer(nIDEvent);
}

void FrmChiphi::SetTimerSearch()
{
	demTimer = 0;
	SetTimer(CONST_TIMER_CHIPHI, 100, NULL);
}

void FrmChiphi::PutKey()
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


void FrmChiphi::OnBnClickedBtnOk()
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
				_bstr_t bs5540 = (LPCTSTR)_xlGetNameSheet(shChiphi, 0);
				if (bs5540 != (_bstr_t)L"")
				{
					_WorksheetPtr pSh5540 = xl->Sheets->GetItem((_bstr_t)bs5540);

					xl->PutScreenUpdating(false);

					// Xác định các cột dữ liệu (bắt đầu --> kết thúc)
					icolSTT = _xlGetColumn(pSh5540, L"M5540_STT", -1);
					iColNdung = _xlGetColumn(pSh5540, L"M5540_NDUNG", -1);
					iColGChu = _xlGetColumn(pSh5540, L"M5540_GCHU", -1);

					if (icolSTT > 0 && iColNdung > 0 && iColGChu > 0)
					{
						int nID = _wtoi(m_list_view.GetItemText(vecItems[0], icotID));
						PutChiphi(nID);
					}

					xl->PutScreenUpdating(true);
				}
			}
			catch (...) {}
		}

		vecItems.clear();
	}
	catch (...) {}
}

void FrmChiphi::OnBnClickedBtnCancel()
{
	PutKey();
	CDialogEx::OnCancel();

	CSendKeys *m_sk = new CSendKeys;
	m_sk->SendKeys(L"{UP}");
	delete m_sk;
}

void FrmChiphi::SetDefault()
{
	try
	{
		m_txt_search.SubclassDlgItem(CHIPHI_TXT_SEARCH, this);
		m_txt_search.SetBkColor(RGB(255, 255, 255));
		m_txt_search.SetTextColor(RGB(0, 0, 255));
		m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm quy trình...");

		szTukhoa.Replace(DF_RIGHTCLICK, L"");
		szTukhoa.Trim();

		m_txt_search.SetWindowText(szTukhoa);
		m_txt_search.PostMessage(EM_SETSEL, -1);

		m_list_view.InsertColumn(icotID, L"", LVCFMT_CENTER, wcID);
		m_list_view.InsertColumn(icotChiphi, L"Nội dung chi phí", LVCFMT_LEFT, wcChiphi);
		m_list_view.InsertColumn(icotTable, L"", LVCFMT_CENTER, wcTable);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(
			LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);
	}
	catch (...) {}
}

void FrmChiphi::SetLocation()
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

void FrmChiphi::ReplaceTukhoa(CString &szKey)
{
	szKey.Replace(L" ", L"");
	szKey.Replace(L"-", L"");
	szKey.Replace(L":", L"");
	szKey.Replace(L"/", L"");
	szKey.Replace(L"(", L"");
	szKey.Replace(L")", L"");
}

int FrmChiphi::LocDulieuChiphi()
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
		long itotal = (long)vecChiphi.size();
		for (long i = 0; i < itotal; i++)
		{
			icheck = 0;
			szval = ConvertKhongDau(vecChiphi[i].szTimkiem);
			szval.MakeLower();

			for (int j = 0; j < iKey; j++)
			{
				if (szval.Find(vecKey[j]) >= 0) icheck++;
				else break;
			}

			if (icheck == iKey)
			{
				vecFilter.push_back(vecChiphi[i]);
			}
		}

		vecKey.clear();
		return (int)vecFilter.size();
	}
	catch (...) {}
	return 0;
}


void FrmChiphi::TimkiemChiphi()
{
	try
	{
		vecFilter.clear();
		if ((int)vecChiphi.size() > 0)
		{
			m_txt_search.GetWindowTextW(szTukhoa);
			szTukhoa.Replace(DF_RIGHTCLICK, L"");
			szTukhoa.Trim();

			if (szTukhoa != L"") LocDulieuChiphi();
			else vecFilter = vecChiphi;
		}

		LoadListView();
	}
	catch (...) {}
}

void FrmChiphi::DeleteAllList()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

int FrmChiphi::LoadListView()
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
		szval.Format(L"Nội dung chi phí (%d kết quả)", tslkq);
		NameColumnHeading(m_list_view, icotChiphi, 1, szval);

		return iz;
	}
	catch (...) {}
	return 0;
}

void FrmChiphi::AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			szval.Format(L"%d", vecFilter[i].nID);
			m_list_view.SetItemText(i, icotID, szval);
			m_list_view.SetItemText(i, icotChiphi, vecFilter[i].szChiphi);
			m_list_view.SetItemText(i, icotTable, vecFilter[i].szTable);

			if (i % 2 == 0) m_list_view.SetRowColors(i, RGB(245, 255, 135), RGB(0, 0, 0));
		}
	}
	catch (...) {}
}

int FrmChiphi::ReadFileChiphi()
{
	try
	{
		vecChiphi.clear();
		szpathDefault = GetPathFolder(FILENAMEDLL);

		CString szChiphiMDB = szpathDefault + FolderChiphi + FileChiphiMDB;
		if (FileExistsUnicode(szChiphiMDB) == 1)
		{
			ObjConn.OpenAccessDB(szChiphiMDB, L"", L"");

			MyStrChiphi MSTC;
			CString szval = L"";			
			SqlString = L"SELECT * FROM tblData ORDER BY AID ASC;";
			CRecordset *Recset = ObjConn.Execute(SqlString);
			while (!Recset->IsEOF())
			{
				Recset->GetFieldValue(L"AID", MSTC.szMa);
				Recset->GetFieldValue(L"AContent", MSTC.szChiphi);
				Recset->GetFieldValue(L"ATable", MSTC.szTable);
				Recset->GetFieldValue(L"AFormula", MSTC.szFormula);
				Recset->GetFieldValue(L"ALegalGrounds", MSTC.szLegalGr);

				if (MSTC.szMa != L"")
				{
					MSTC.szTimkiem.Format(L"%s|%s", MSTC.szMa, MSTC.szChiphi);

					MSTC.nID = (int)vecChiphi.size();
					vecChiphi.push_back(MSTC);
				}

				Recset->MoveNext();
			}
			Recset->Close();

			ObjConn.CloseDatabase();
		}

		return (int)vecChiphi.size();
	}
	catch (...) {}
	return 0;
}

void FrmChiphi::OnLvnEndScrollListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	ShowResultNext();
}

void FrmChiphi::OnEnChangeTxtSearch()
{
	if (iUnikey == 1)
	{
		CString szval = L"";
		m_txt_search.GetWindowTextW(szval);
		szval.Replace(DF_RIGHTCLICK, L"");
		szval.Trim();
		
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

int FrmChiphi::GetSelectItems(vector<int> &vecItem)
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

void FrmChiphi::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == icotID) pNMListHeader->pitem->cxy = 0;
	*pResult = 0;
}

void FrmChiphi::OnNMDblclkListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedBtnOk();
}

void FrmChiphi::SelectListItem(bool blEnter)
{
	GotoDlgCtrl(GetDlgItem(CHIPHI_LIST_SEARCH));

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

bool FrmChiphi::CheckArrowDownNext()
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

void FrmChiphi::ShowResultNext()
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

void FrmChiphi::PutChiphi(int nID)
{
	try
	{
		_bstr_t bsDMTL = (LPCTSTR)_xlGetNameSheet(shDMTL, 0);
		if (bsDMTL != (_bstr_t)L"")
		{
			_WorksheetPtr pShDMTL = xl->Sheets->GetItem((_bstr_t)bsDMTL);
			RangePtr pRgDMTL = pShDMTL->Cells;

			int nColTalble = _xlGetColumn(pShDMTL, L"TL_BANG", -1);
			if (nColTalble > 0)
			{
				_bstr_t bs5540 = (LPCTSTR)_xlGetNameSheet(shChiphi, 0);
				if (bs5540 != (_bstr_t)L"")
				{
					_WorksheetPtr pSh5540 = xl->Sheets->GetItem((_bstr_t)bs5540);

					int icotSTT = _xlGetColumn(pSh5540, L"M5540_STT", -1);
					int icotNdung = _xlGetColumn(pSh5540, L"M5540_NDUNG", -1);
					int icotGTDT = _xlGetColumn(pSh5540, L"M5540_GTDT", -1);
					int icotCD_GT = _xlGetColumn(pSh5540, L"M5540_CD_GT", -1);
					int icotCD_TL = _xlGetColumn(pSh5540, L"M5540_CD_TL", -1);
					int icotCT_GT = _xlGetColumn(pSh5540, L"M5540_CT_GT", -1);
					int icotCT_TL = _xlGetColumn(pSh5540, L"M5540_CT_TL", -1);
					int icotTLNS = _xlGetColumn(pSh5540, L"M5540_TLNS", -1);
					int icotHSK1 = _xlGetColumn(pSh5540, L"M5540_HSK1", -1);
					int icotHSK2 = _xlGetColumn(pSh5540, L"M5540_HSK2", -1);
					int icotTT_DM = _xlGetColumn(pSh5540, L"M5540_TT_DM", -1);

					if (icotNdung > 0)
					{
						pRgActive->PutItem(iRowActive, icotNdung, (_bstr_t)vecChiphi[nID].szChiphi);
					}

					RangePtr pCell;
					CString szval = L"";
					int nIndex = _wtoi(vecChiphi[nID].szTable);

					if (nIndex > 0)
					{
						szval.Format(L"BANG%02d", nIndex);
						int nRowX = _xlFindValue(pShDMTL, nColTalble, 0, 0, (_bstr_t)szval, 0, false, false);
						if (nRowX > 0)
						{
							nRowX += 2;

							if (icotGTDT > 0)
							{
								szval.Format(L"='%s'!%s/10^9",
									(LPCTSTR)sNameActive, _xlGAR(pRgActive, iRowActive, icotGTDT, 0));
								pRgDMTL->PutItem(nRowX, nColTalble + 5, (_bstr_t)szval);
							}

							if (icotCD_GT > 0)
							{
								szval.Format(L"='%s'!%s", (LPCTSTR)bsDMTL, _xlGAR(pRgDMTL, nRowX, nColTalble, 0));
								pRgActive->PutItem(iRowActive, icotCD_GT, (_bstr_t)szval);
							}

							if (icotCD_TL > 0)
							{
								szval.Format(L"='%s'!%s/100", (LPCTSTR)bsDMTL, _xlGAR(pRgDMTL, nRowX, nColTalble + 2, 0));
								pRgActive->PutItem(iRowActive, icotCD_TL, (_bstr_t)szval);
							}

							if (icotCT_GT > 0)
							{
								szval.Format(L"='%s'!%s", (LPCTSTR)bsDMTL, _xlGAR(pRgDMTL, nRowX, nColTalble + 1, 0));
								pRgActive->PutItem(iRowActive, icotCT_GT, (_bstr_t)szval);
							}

							if (icotCT_TL > 0)
							{
								szval.Format(L"='%s'!%s/100", (LPCTSTR)bsDMTL, _xlGAR(pRgDMTL, nRowX, nColTalble + 3, 0));
								pRgActive->PutItem(iRowActive, icotCT_TL, (_bstr_t)szval);
							}

							if (icotTLNS > 0)
							{
								szval.Format(L"='%s'!%s/100", (LPCTSTR)bsDMTL, _xlGAR(pRgDMTL, nRowX, nColTalble + 4, 0));
								pRgActive->PutItem(iRowActive, icotTLNS, (_bstr_t)szval);
							}
						}

						if (icotTT_DM > 0 && icotGTDT > 0 && icotTLNS > 0 && icotHSK2 > 0)
						{
							szval.Format(L"=PRODUCT(%s,%s:%s)",
								_xlGAR(pRgActive, iRowActive, icotGTDT, 0),
								_xlGAR(pRgActive, iRowActive, icotTLNS, 0),
								_xlGAR(pRgActive, iRowActive, icotHSK2, 0));
							pRgActive->PutItem(iRowActive, icotTT_DM, (_bstr_t)szval);
						}
					}
					else
					{
						if (icotCD_GT > 0)
						{
							pCell = pRgActive->GetItem(iRowActive, icotCD_GT);
							pCell->PutHorizontalAlignment(xlLeft);
							pCell->PutValue2((_bstr_t)vecChiphi[nID].szFormula);
						}

						if (icotCD_TL > 0) pRgActive->PutItem(iRowActive, icotCD_TL, (_bstr_t)L"");
						if (icotCT_GT > 0) pRgActive->PutItem(iRowActive, icotCT_GT, (_bstr_t)L"");
						if (icotCT_TL > 0) pRgActive->PutItem(iRowActive, icotCT_TL, (_bstr_t)L"");
						if (icotTLNS > 0) pRgActive->PutItem(iRowActive, icotTLNS, (_bstr_t)L"");
						if (icotTT_DM > 0) pRgActive->PutItem(iRowActive, icotTT_DM, (_bstr_t)L"");
					}

					int nRowEnd = _xlFindComment(pShActive, icotSTT, 0, 0, (_bstr_t)L"END", -1);
					if (iRowActive + 1 == nRowEnd)
					{
						pCell = pShActive->GetRange(
							pRgActive->GetItem(nRowEnd, 1), pRgActive->GetItem(nRowEnd + 4, 1));
						pCell->EntireRow->Insert(xlShiftDown);
						nRowEnd += 5;

						pCell = pShActive->GetRange(
							pRgActive->GetItem(iRowActive, icolSTT), pRgActive->GetItem(nRowEnd - 1, iColGChu));
						pCell->Borders->GetItem(XlBordersIndex::xlInsideHorizontal)->PutWeight(xlHairline);
					}

					if (icotSTT > 0)
					{
						pRgActive->PutItem(iRowActive, icotSTT, 1);

						int nRowStart = 2 + _xlGetRow(pSh5540, L"M5540_STT", 0);

						int dem = 0;
						for (int i = nRowStart; i < nRowEnd; i++)
						{
							szval = _xlGIOR(pRgActive, i, icotSTT, L"");
							if (_wtoi(szval) > 0)
							{
								dem++;
								pRgActive->PutItem(iRowActive, icotSTT, dem);
							}
						}
					}

					pCell = pRgActive->GetItem(iRowActive, 1);
					pCell->EntireRow->Rows->AutoFit();
				}				
			}
		}
	}
	catch (...) {}
}