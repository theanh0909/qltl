#include "pch.h"
#include "GHookForm.h"
#include "FrmTieuchuan.h"
#include "HookMouse.h"

// FrmTieuchuan dialog
IMPLEMENT_DYNAMIC(FrmTieuchuan, CDialogEx)

FrmTieuchuan::FrmTieuchuan(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FRM_TIEUCHUAN, pParent)
{
	szTukhoa = L"";
	szpathDefault = L"";

	blChange = false;

	demTimer = 0;
	icotIcon = 0, icotName = 1, icotID = 2;
	wcIcon = 22, wcTen = 720, wcID = 0;

	iUnikey = 0;
	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 20;

	nIndexMouse = 2;
	MyHook::Instance().InstallHook();
}

FrmTieuchuan::~FrmTieuchuan()
{
	vecTieuchuan.clear();
	vecFilter.clear();

	nIndexMouse = 0;
	MyHook::Instance().UninstallHook();	
}

void FrmTieuchuan::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TCHUAN_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, TCHUAN_LIST_SEARCH, m_list_view);
}

BEGIN_MESSAGE_MAP(FrmTieuchuan, CDialogEx)
	ON_WM_TIMER()
	ON_WM_ERASEBKGND()
	ON_WM_SYSCOMMAND()
	ON_NOTIFY(HDN_ENDTRACK, 0, &FrmTieuchuan::OnHdnEndtrackListView)
	ON_EN_CHANGE(TCHUAN_TXT_SEARCH, &FrmTieuchuan::OnEnChangeTxtSearch)		
	ON_NOTIFY(NM_DBLCLK, TCHUAN_LIST_SEARCH, &FrmTieuchuan::OnNMDblclkListSearch)
	ON_NOTIFY(LVN_ENDSCROLL, TCHUAN_LIST_SEARCH, &FrmTieuchuan::OnLvnEndScrollListSearch)
END_MESSAGE_MAP()

// FrmTieuchuan message handlers
BOOL FrmTieuchuan::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	//SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	//SetLayeredWindowAttributes(RGB(255, 255, 255), 0, LWA_COLORKEY);

	hWndTieuchuan = m_hWnd;

	SetDefault();
	SetLocation();
	ReadFileTieuchuan();

	if (szTukhoa != L"") SetTimerSearch();
	else TimkiemTieuchuan();

	blChange = true;
	iUnikey++;

	return TRUE;
}

BOOL FrmTieuchuan::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(TCHUAN_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TCHUAN_TXT_SEARCH))
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
			if (pWndWithFocus == GetDlgItem(TCHUAN_LIST_SEARCH))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(TCHUAN_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TCHUAN_LIST_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(TCHUAN_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_DOWN)
		{
			if (pWndWithFocus == GetDlgItem(TCHUAN_TXT_SEARCH))
			{
				SelectListItem(false);
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TCHUAN_LIST_SEARCH))
			{
				CheckArrowDownNext();
			}
		}
		else if (iWPar == VK_UP)
		{
			if (pWndWithFocus == GetDlgItem(TCHUAN_LIST_SEARCH))
			{
				vector<int> vecItems;
				int ncount = GetSelectItems(vecItems);
				if (ncount == 1 && vecItems[0] == 0)
				{
					GotoDlgCtrl(GetDlgItem(TCHUAN_TXT_SEARCH));
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

void FrmTieuchuan::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

BOOL FrmTieuchuan::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(192, 192, 192));
	return FALSE;
}

void FrmTieuchuan::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;

	if (demTimer >= 2)
	{
		demTimer = 0;
		KillTimer(CONST_TIMER_TIEUCHUAN);

		TimkiemTieuchuan();
	}

	CDialogEx::OnTimer(nIDEvent);
}

void FrmTieuchuan::SetTimerSearch()
{
	demTimer = 0;
	SetTimer(CONST_TIMER_TIEUCHUAN, 100, NULL);
}


void FrmTieuchuan::PutKey()
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


void FrmTieuchuan::OnBnClickedBtnOk()
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
		
		// Giới hạn chỉ cho phép chọn từng tiêu chuẩn
		if (nSelect > 1) nSelect = 1;

		if (nSelect > 0)
		{
			nIndexMouse = 0;
			MyHook::Instance().UninstallHook();
			CDialogEx::OnOK();

			try
			{
				if (!FolderExistsUnicode(szpathDefault + FolderThuvienTC))
				{
					// Tạo thư mục thư viện lưu file
					CreateDirectories(szpathDefault + FolderThuvienTC);
				}

				xl->PutScreenUpdating(false);
				RangePtr pCell;

				// Chèn dòng nếu số lượng thêm lớn hơn 1
				if (nSelect > 1)
				{
					pCell = pShActive->GetRange(
						pRgActive->GetItem(iRowActive + 1, 1), 
						pRgActive->GetItem(iRowActive + nSelect - 1, 1));
					pCell->EntireRow->Insert(xlShiftDown);
				}

				int nID = 0;
				CString szLink = L"";

				int iColTen = _xlGetColumn(pShActive, L"TCH_TENTC", -1);
				int iColPLoai = _xlGetColumn(pShActive, L"TCH_PLOAI", -1);
				int iColGChu = _xlGetColumn(pShActive, L"TCH_GCHU", -1);
				int iColLink = _xlGetColumn(pShActive, L"TCH_OPEN", -1);

				sv = new CDownloadFileSever;
				sv->_LoadDecodeBase64();

				bool blLink = false;
				for (int i = 0; i < nSelect; i++)
				{
					nID = _wtoi(m_list_view.GetItemText(vecItems[i], icotID));

					pCell = pRgActive->GetItem(iRowActive + i, iColumnActive);
					pCell->PutValue2((_bstr_t)vecTieuchuan[nID].szMa);

					if (iColTen > 0)
					{
						pCell = pRgActive->GetItem(iRowActive + i, iColTen);
						pCell->PutValue2((_bstr_t)vecTieuchuan[nID].szTen);
					}

					if (iColPLoai > 0)
					{
						pCell = pRgActive->GetItem(iRowActive + i, iColPLoai);
						pCell->PutValue2((_bstr_t)vecTieuchuan[nID].szPloai);
					}

					if (iColGChu > 0)
					{
						pCell = pRgActive->GetItem(iRowActive + i, iColGChu);
						pCell->PutValue2((_bstr_t)vecTieuchuan[nID].szGchu);
					}

					if (iColLink > 0)
					{
						szLink = DownloadFileTieuchuan(vecTieuchuan[nID].szDownload);
						if (szLink != L"")
						{
							blLink = true;
							pCell = pRgActive->GetItem(iRowActive + i, iColLink);
							_xlSetHyperlink(pShActive, pCell, szLink);
						}
					}

					pCell = pRgActive->GetItem(iRowActive + i, iColTen);
					pCell->Interior->PutColor(xlNone);
					if (blLink) pCell->Font->PutColor(rgbBlack);
					else pCell->Font->PutColor(rgbRed);
				}

				xl->PutScreenUpdating(true);
				delete sv;
			}
			catch (...) {}
		}

		vecItems.clear();
	}
	catch (...) {}
}

void FrmTieuchuan::OnBnClickedBtnCancel()
{
	PutKey();
	CDialogEx::OnCancel();

	CSendKeys *m_sk = new CSendKeys;
	m_sk->SendKeys(L"{UP}");
	delete m_sk;
}

void FrmTieuchuan::SetDefault()
{
	try
	{
		m_txt_search.SubclassDlgItem(TCHUAN_TXT_SEARCH, this);
		m_txt_search.SetBkColor(RGB(255, 255, 255));
		m_txt_search.SetTextColor(RGB(0, 0, 255));
		m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm tiêu chuẩn...");

		m_txt_search.SetWindowText(szTukhoa);
		m_txt_search.PostMessage(EM_SETSEL, -1);

		m_list_view.InsertColumn(icotIcon, L"", LVCFMT_CENTER, wcIcon);
		m_list_view.InsertColumn(icotName, L"Tiêu chuẩn", LVCFMT_LEFT, wcTen);
		m_list_view.InsertColumn(icotID, L"", LVCFMT_CENTER, wcID);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(
			LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);
	}
	catch (...) {}
}

void FrmTieuchuan::SetLocation()
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

void FrmTieuchuan::ReplaceTukhoa(CString &szKey)
{
	szKey.Replace(L" ", L"");
	szKey.Replace(L"-", L"");
	szKey.Replace(L":", L"");
	szKey.Replace(L"/", L"");
	szKey.Replace(L"(", L"");
	szKey.Replace(L")", L"");
}

int FrmTieuchuan::LocDulieuTieuchuan()
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
		long itotal = (long)vecTieuchuan.size();
		for (long i = 0; i < itotal; i++)
		{
			icheck = 0;
			szval = ConvertKhongDau(vecTieuchuan[i].szTimkiem);
			szval.MakeLower();

			for (int j = 0; j < iKey; j++)
			{
				if (szval.Find(vecKey[j]) >= 0) icheck++;
				else break;
			}

			if (icheck == iKey)
			{
				vecFilter.push_back(vecTieuchuan[i]);
			}
		}

		vecKey.clear();
		return (int)vecFilter.size();
	}
	catch (...) {}
	return 0;
}


void FrmTieuchuan::TimkiemTieuchuan()
{
	try
	{
		vecFilter.clear();
		if ((int)vecTieuchuan.size() > 0)
		{
			m_txt_search.GetWindowTextW(szTukhoa);
			szTukhoa.Trim();

			if (szTukhoa != L"") LocDulieuTieuchuan();
			else vecFilter = vecTieuchuan;
		}

		LoadListView();
	}
	catch (...) {}
}

void FrmTieuchuan::DeleteAllList()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

void FrmTieuchuan::DeleteListIcon()
{
	m_imageList.DeleteImageList();
	ASSERT(m_imageList.GetSafeHandle() == NULL);
	m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);
}

int FrmTieuchuan::LoadListView()
{
	try
	{
		tslkq = (int)vecFilter.size();

		DeleteAllList();
		DeleteListIcon();

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
		szval.Format(L"Tiêu chuẩn (%d kết quả)", tslkq);
		NameColumnHeading(m_list_view, icotName, 1, szval);

		return iz;
	}
	catch (...) {}
	return 0;
}


void FrmTieuchuan::AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		//HICON hLnkDown = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_DOWNLOAD));
		HICON hLnkNull = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_SEARCH));

		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			szval.Format(L"%d", vecFilter[i].nID);
			m_list_view.SetItemText(i, icotID, szval);
			m_list_view.SetItemText(i, icotName, vecFilter[i].szTen);

			//if (vecFilter[i].szDownload != L"") m_imageList.Add(hLnkDown);
			//else m_imageList.Add(hLnkNull);

			szval = GetTypeOfFile(vecFilter[i].szDownload);
			if (szval != L"")
			{
				if (szval.Left(2) != L"*.") szval = L"*." + szval;
				HICON hi = m_sysImageList.GetFileIcon(szval);
				m_imageList.Add(hi);
			}
			else m_imageList.Add(hLnkNull);

			if (i % 2 == 0) m_list_view.SetRowColors(i, RGB(245, 255, 135), RGB(0, 0, 0));
		}

		// Load Image vào list dữ liệu
		m_list_view.SetImageList(&m_imageList);

		for (int i = iBegin; i < iEnd; i++)
		{
			// Load Image icon
			// Trường hợp icotIcon > 0 --> phải xóa icon ở cột 0 (set: nImage = -1)
			//m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, -1, 0, 0, 0);
			m_list_view.SetItem(i, icotIcon, LVIF_IMAGE, NULL, i, 0, 0, 0);
		}
	}
	catch (...) {}
}

int FrmTieuchuan::ReadFileTieuchuan()
{
	try
	{
		vecTieuchuan.clear();
		szpathDefault = GetPathFolder(FILENAMEDLL);

		CString szTieuchuanMDB = szpathDefault + FolderTieuchuan + FileTieuchuanMDB;
		if (FileExistsUnicode(szTieuchuanMDB) == 1)
		{
			ObjConn.OpenAccessDB(szTieuchuanMDB, L"", L"");

			CString szval = L"";
			MyStrTieuchuan MSTC;
			SqlString = L"SELECT * FROM tblTieuchuan ORDER BY Maso ASC;";
			CRecordset *Recset = ObjConn.Execute(SqlString);
			while (!Recset->IsEOF())
			{
				Recset->GetFieldValue(L"Maso", MSTC.szMa);
				Recset->GetFieldValue(L"Tieuchuan", MSTC.szTen);
				Recset->GetFieldValue(L"Phanloai", MSTC.szPloai);
				Recset->GetFieldValue(L"Ghichu", MSTC.szGchu);
				Recset->GetFieldValue(L"Download", MSTC.szDownload);

				if (MSTC.szMa != L"")
				{
					szval = MSTC.szMa;
					ReplaceTukhoa(szval);

					MSTC.szTimkiem.Format(L"%s|%s|%s|%s|%s", 
						szval, MSTC.szMa, MSTC.szTen, MSTC.szPloai, MSTC.szGchu);

					MSTC.nID = (int)vecTieuchuan.size();
					vecTieuchuan.push_back(MSTC);
				}

				Recset->MoveNext();
			}
			Recset->Close();

			ObjConn.CloseDatabase();
		}

		return (int)vecTieuchuan.size();
	}
	catch (...) {}
	return 0;
}

void FrmTieuchuan::OnLvnEndScrollListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	ShowResultNext();
}

void FrmTieuchuan::OnEnChangeTxtSearch()
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

int FrmTieuchuan::GetSelectItems(vector<int> &vecItem)
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

CString FrmTieuchuan::DownloadFileTieuchuan(CString szFileTchuan)
{
	try
	{
		if (szFileTchuan != L"")
		{
			CString szpathFile = szpathDefault + FolderThuvienTC + szFileTchuan;

			// Kiểm tra file có tồn tại trong thư mục gốc không?
			if (FileExistsUnicode(szpathFile) == 1) return szpathFile;

			// Kiểm tra file có tồn tại trên sever không?
			CString severTieuchuan = sv->dbSeverMain + L"tieuchuan/";
			if (sv->_DownloadFile(severTieuchuan + szFileTchuan, szpathFile)) return szpathFile;
		}		
	}
	catch (...) {}
	return L"";
}


CString FrmTieuchuan::DownloadFileQuytrinh(CString szFileQtrinh)
{
	try
	{
		if (szFileQtrinh != L"")
		{
			if (szFileQtrinh.Left(1) == L"/" || szFileQtrinh.Left(1) == L"\\")
			{
				szFileQtrinh = szFileQtrinh.Right(szFileQtrinh.GetLength() - 1);
			}

			CString szpathFile = szpathDefault + FolderThuvienQT + szFileQtrinh;

			int pos = -1;
			CString szFolder = GetNameOfPath(szpathFile, pos, 1);
			CreateDirectories(szFolder);

			// Kiểm tra file có tồn tại trong thư mục gốc không?
			if (FileExistsUnicode(szpathFile) == 1) return szpathFile;

			// Kiểm tra file có tồn tại trên sever không?
			CString severQuytrinh = sv->dbSeverMain + L"quytrinh/BieuMau/";
			if (sv->_DownloadFile(severQuytrinh + szFileQtrinh, szpathFile)) return szpathFile;
		}		
	}
	catch (...) {}
	return L"";
}

void FrmTieuchuan::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == icotID) pNMListHeader->pitem->cxy = 0;
	else if (pSubItem == icotIcon) pNMListHeader->pitem->cxy = 22;
	*pResult = 0;
}

void FrmTieuchuan::OnNMDblclkListSearch(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnBnClickedBtnOk();
}

void FrmTieuchuan::SelectListItem(bool blEnter)
{
	GotoDlgCtrl(GetDlgItem(TCHUAN_LIST_SEARCH));

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

bool FrmTieuchuan::CheckArrowDownNext()
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

void FrmTieuchuan::ShowResultNext()
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


void FrmTieuchuan::HotroDownloadQuytrinh()
{
	try
	{
		int iColDownload = _xlGetColumn(pShActive, L"QTR_GOC", -1);
		if (iColDownload > 0)
		{
			CString szval = _xlGIOR(pRgActive, iRowActive, iColDownload, L"");
			if (szval != L"")
			{
				RangePtr pCell = pRgActive->GetItem(iRowActive, iColDownload);
				if (szval != L"Mở file")
				{
					szpathDefault = GetPathFolder(FILENAMEDLL);
					if (!FolderExistsUnicode(szpathDefault + FolderThuvienQT))
					{
						// Tạo thư mục thư viện lưu file
						CreateDirectories(szpathDefault + FolderThuvienQT);
					}

					int result = _y(L"Bạn có chắc chắn tải dữ liệu biểu mẫu?", 0, 0, 0);
					if (result == 6)
					{
						xl->PutStatusBar((_bstr_t)L"Đang tải dữ liệu. Vui lòng đợi trong giây lát...");

						sv = new CDownloadFileSever;
						sv->_LoadDecodeBase64();

						CString szLink = DownloadFileQuytrinh(szval);
						if (szLink != L"")
						{
							_xlSetHyperlink(pShActive, pCell, szLink);

							int col1 = _xlGetColumn(pShActive, L"QTR_FOLDER", -1);
							int col2 = _xlGetColumn(pShActive, L"QTR_DIR", -1);
							int col3 = _xlGetColumn(pShActive, L"QTR_FILE", -1);

							int pos = -1;
							CString szFolder = GetNameOfPath(szLink, pos, 1);
							CString szFileName = szLink.Right(szLink.GetLength() - pos - 1);

							pRgActive->PutItem(iRowActive, col2, (_bstr_t)szFolder);
							pRgActive->PutItem(iRowActive, col3, (_bstr_t)szFileName);
							pRgActive->PutItem(iRowActive, col1, (_bstr_t)GetNameOfPath(szLink, pos, 0));
						}

						delete sv;
						xl->PutStatusBar((_bstr_t)L"Ready");
					}
				}
				else
				{
					vector<CString> vecHyper;
					int ncoutHyp = _xlGetAllHyperLink(pCell, vecHyper, 1, 1);
					if (ncoutHyp > 0) ShellExecuteSelectFile(vecHyper[0]);
					vecHyper.clear();
				}
			}
		}
	}
	catch (...) {}
}