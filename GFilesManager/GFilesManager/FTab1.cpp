#include "pch.h"
#include "FTab1.h"

IMPLEMENT_DYNAMIC(FTab1, CDialogEx)

FTab1::FTab1(CWnd* pParent /*=NULL*/)
	: CDialogEx(FTab1::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	iGFLNum = 1, iGFLTen = 2, iGFLPath = 3;
	iwColNum = 40, iwColTen = 200, iwColPath = 120;

	jColumnHand = iGFLPath;
	blDragging = false;

	iStopload = 1, lanshow = 1, ibuocnhay = 50;
	tongkq = 0, tongfilter = 0;

	nItem = -1, nSubItem = -1;

	pathFolderDir = L"";
	szCheckAddFiles = L"";
}

FTab1::~FTab1()
{
	delete ff;
	delete bb;
	delete reg;

	vecItem.clear();
	vecFilter.clear();
	vecAllFiles.clear();
}

void FTab1::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TAB1_BTN_ADD, m_btn_add);
	DDX_Control(pDX, TAB1_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, TAB1_LIST_VIEW, m_list_view);
}

BEGIN_MESSAGE_MAP(FTab1, CDialogEx)
	ON_WM_TIMER()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_COMMAND(MNTAB1_OPENFILE, &FTab1::OnOpenfile)
	ON_COMMAND(MNTAB1_OPENFOLDER, &FTab1::OnOpenfolder)
	ON_COMMAND(MNTAB1_ADDFILES, &FTab1::OnAddfiles)
	ON_COMMAND(MNTAB1_DELETEFILES, &FTab1::OnDeletefiles)
	ON_COMMAND(MNTAB1_MOVEUP, &FTab1::OnMoveup)
	ON_COMMAND(MNTAB1_MOVEDOWN, &FTab1::OnMovedown)
	ON_COMMAND(MNTAB1_REVERSE, &FTab1::OnReverse)
	ON_COMMAND(MNTAB1_BGKCOLOR, &FTab1::OnBgkcolor)

	ON_BN_CLICKED(TAB1_BTN_ADD, &FTab1::OnBnClickedBtnAdd)

	ON_NOTIFY(NM_CLICK, TAB1_LIST_VIEW, &FTab1::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, TAB1_LIST_VIEW, &FTab1::OnNMRClickListView)
	ON_NOTIFY(NM_DBLCLK, TAB1_LIST_VIEW, &FTab1::OnNMDblclkListView)	
	ON_NOTIFY(LVN_ENDSCROLL, TAB1_LIST_VIEW, &FTab1::OnLvnEndScrollListView)	
	ON_NOTIFY(LVN_KEYDOWN, TAB1_LIST_VIEW, &FTab1::OnLvnKeydownListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &FTab1::OnHdnEndtrackListView)
END_MESSAGE_MAP()

// FTab1 message handlers
BOOL FTab1::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	mnIcon.GdiplusStartupInputConfig();

	pathFolderDir = ff->_GetPathFolder(szAppSmart);

	_SetDefault();	

	szCheckAddFiles = bb->BaseDecodeEx(CreateKeySMStart);

	OnBnClickedBtnSearch();

	SetTimer(CONST_TIMER_ID, 100, NULL);	// 1000ms = 1 second
	
	return TRUE;
}

BOOL FTab1::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(TAB1_TXT_SEARCH))
			{
				OnBnClickedBtnSearch();
				return TRUE;
			}			
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(TAB1_TXT_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(TAB1_LIST_VIEW));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TAB1_LIST_VIEW))
			{
				GotoDlgCtrl(GetDlgItem(TAB1_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_UP)
		{
			if (GetKeyState(VK_CONTROL) < 0)
			{
				OnMoveup();
				return TRUE;
			}
		}
		else if (iWPar == VK_DOWN)
		{
			if (GetKeyState(VK_CONTROL) < 0)
			{
				OnMovedown();
				return TRUE;
			}
		}		
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(TAB1_LIST_VIEW))
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

void FTab1::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialogEx::OnSysCommand(nID, lParam);
}

void FTab1::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_TAB));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;

			vector<int> vecRow;

			CString szval = L"";
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				int ncount = m_list_view.GetItemCount();
				if (nSelect < ncount)
				{
					szval.Format(L"Xóa (%d files) khỏi danh sách (Delete)", nSelect);
				}
				else
				{
					szval.Format(L"Xóa toàn bộ files khỏi danh sách (Delete)");
				}

				ModifyMenu(pMenu, MNTAB1_DELETEFILES, MF_BYCOMMAND | MF_STRING, MNTAB1_DELETEFILES, szval);
			}

			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_ADDFILES, IDI_ICON_ADD);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_DELETEFILES, IDI_ICON_DELETE);
			
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_MOVEUP, IDI_ICON_UP);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_MOVEDOWN, IDI_ICON_DOWN);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_REVERSE, IDI_ICON_REVERSE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MNTAB1_BGKCOLOR, IDI_ICON_COLOR);

			if (nItem == -1 || nSubItem == -1)
			{
				EnableMenuItem(pMenu, MNTAB1_DELETEFILES, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MNTAB1_BGKCOLOR, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			if (nItem == -1 || nSubItem == -1 || nSelect >= 2)
			{
				EnableMenuItem(pMenu, MNTAB1_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MNTAB1_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

				EnableMenuItem(pMenu, MNTAB1_MOVEUP, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MNTAB1_MOVEDOWN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			if (nSelect != 2) EnableMenuItem(pMenu, MNTAB1_REVERSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

			if (nItem == 0) EnableMenuItem(pMenu, MNTAB1_MOVEUP, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			else if (nItem + 1 == (int)m_list_view.GetItemCount()) EnableMenuItem(pMenu, MNTAB1_MOVEDOWN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(TAB1_LIST_VIEW));
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

void FTab1::_SetDefault()
{
	try
	{
		m_txt_search.SubclassDlgItem(TAB1_TXT_SEARCH, this);
		m_txt_search.SetBkColor(RGBWHITE);
		m_txt_search.SetTextColor(RGBBLUE);
		m_txt_search.SetWindowText(L"");
		m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm...");

		m_btn_add.SetThemeHelper(&m_ThemeHelper);
		m_btn_add.SetIcon(IDI_ICON_ADD2, 16, 16);
		m_btn_add.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_add.SetBtnCursor(IDC_CURSOR_HAND);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
		m_list_view.InsertColumn(iGFLNum, L"", LVCFMT_CENTER, iwColNum);
		m_list_view.InsertColumn(iGFLTen, L"Tên file", LVCFMT_LEFT, iwColTen);
		m_list_view.InsertColumn(iGFLPath, L"Đường dẫn", LVCFMT_LEFT, iwColPath);

		m_list_view.SetColumnColors(iGFLPath, RGBWHITE, RGBBLUE);
		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeader);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);
	}
	catch (...) {}
}

int FTab1::_LocDulieuNangcao()
{
	try
	{
		vecFilter.clear();

		CString szTukhoa = L"";
		m_txt_search.GetWindowTextW(szTukhoa);
		szTukhoa.Trim();

		// Xác định tổng số dữ liệu trong file .ini
		int nsize = (int)vecAllFiles.size();
		if (nsize > 0)
		{
			vecItem = vecAllFiles;
			vecAllFiles.clear();
		}
		else
		{
			if (szTukhoa == L"") _GetListMyFiles(vecItem);
		}

		tongkq = (int)vecItem.size();
		if (tongkq > 0)
		{
			if (szTukhoa != L"")
			{
				szTukhoa = ff->_ConvertKhongDau(szTukhoa);
				szTukhoa.MakeLower();

				vector<CString> vecKey;
				int iKey = ff->_TackTukhoa(vecKey, szTukhoa, L" ", L"+", 1, 0);

				CString szval = L"";
				int pos = -1, icheck = 0;

				for (int i = 0; i < tongkq; i++)
				{
					icheck = 0;
					szval = ff->_GetNameOfPath(vecItem[i].szFullpath, pos, -1);
					szval = ff->_ConvertKhongDau(szval);
					szval.MakeLower();

					for (int j = 0; j < iKey; j++)
					{
						if (szval.Find(vecKey[j]) >= 0) icheck++;
						else break;
					}

					if (icheck == iKey)
					{
						vecFilter.push_back(vecItem[i]);
					}
				}

				vecKey.clear();
			}
			else
			{
				for (int i = 0; i < tongkq; i++) vecFilter.push_back(vecItem[i]);
			}
		}

		tongfilter = (int)vecFilter.size();
		return tongfilter;
	}
	catch (...) {}
	return 0;
}

int FTab1::_GetAllSelectedItems(vector<int> &vecRow)
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
				for (int i = ncount; i < tongfilter; i++) vecRow.push_back(i);
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void FTab1::_SelectAllItems()
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

void FTab1::_ClearAllItems()
{
	m_list_view.DeleteAllItems();
	ASSERT(m_list_view.GetItemCount() == 0);
}

void FTab1::_ClearAllImages()
{
	m_imageList.DeleteImageList();
	ASSERT(m_imageList.GetSafeHandle() == NULL);
	m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);
}

void FTab1::_LoadListImages()
{
	try
	{
		_ClearAllImages();

		HICON hLnkAdd = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_ADD));
		m_imageList.Add(hLnkAdd);

		HICON hLnkShortCut = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_SHORTCUT));
		HICON hLnkFolder = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_OPENFOLDER));
		HICON hLnkNull = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_NULL));

		CString szval = L"";
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			szval = m_list_view.GetItemText(i, iGFLPath);

			// Kiểm tra là thư mục hay file
			if (!ff->_FolderExistsUnicode(szval))
			{
				szval = ff->_GetTypeOfFile(szval);
				if (szval != L"")
				{
					if (szval.Left(2) != L"*.") szval = L"*." + szval;
					if (szval != L"*.lnk")
					{
						HICON hi = m_sysImageList.GetFileIcon(szval);
						m_imageList.Add(hi);
					}
					else m_imageList.Add(hLnkShortCut);
				}
				else m_imageList.Add(hLnkNull);
			}			
			else m_imageList.Add(hLnkFolder);
		}

		// Load Image vào list dữ liệu
		m_list_view.SetImageList(&m_imageList);
	}
	catch (...) {}
}

void FTab1::_LoadListMyFiles()
{
	try
	{
		_ClearAllItems();		// Xóa toàn bộ dữ liệu
		_ClearAllImages();		// Xóa toàn bộ list-images

		lanshow = 1;
		int iz = ibuocnhay;
		if (tongfilter <= iz)
		{
			iz = tongfilter;
			iStopload = 0;
		}
		else iStopload = 1;

		// Add dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz);
	}
	catch (...) {}
}

void FTab1::_Timkiemdulieu()
{
	
}

void FTab1::_AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		int dem = 0, ncount = 0;
		if (iBegin == 0)
		{
			HICON hLnkAdd = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_ADD));
			m_imageList.Add(hLnkAdd);
		}
		else
		{
			ncount = m_list_view.GetItemCount();
			POSITION pos = m_list_view.GetFirstSelectedItemPosition();
			if (pos != NULL)
			{
				for (int i = 0; i < m_list_view.GetItemCount(); i++)
				{
					if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;
				}
			}
		}

		HICON hLnkShortCut = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_SHORTCUT));
		HICON hLnkFolder = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_OPENFOLDER));
		HICON hLnkNull = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_NULL));

		int pos = -1;
		CString szval = L"";

		for (int i = iBegin; i < iEnd; i++)
		{
			// ID
			szval.Format(L"%d", vecFilter[i].iID);
			m_list_view.InsertItem(i, szval, 0);

			// STT
			szval.Format(L"%d", i + 1);
			m_list_view.SetItemText(i, iGFLNum, szval);

			m_list_view.SetItemText(i, iGFLTen, ff->_GetNameOfPath(vecFilter[i].szFullpath, pos, 0));
			m_list_view.SetItemText(i, iGFLPath, vecFilter[i].szFullpath);

			// Kiểm tra là thư mục hay file
			if (!ff->_FolderExistsUnicode(vecFilter[i].szFullpath))
			{
				// Type file
				szval = ff->_GetTypeOfFile(vecFilter[i].szFullpath);
				if (szval != L"")
				{
					if (szval.Left(2) != L"*.") szval = L"*." + szval;
					if (szval != L"*.lnk")
					{
						HICON hi = m_sysImageList.GetFileIcon(szval);
						m_imageList.Add(hi);
					}
					else m_imageList.Add(hLnkShortCut);
				}
				else m_imageList.Add(hLnkNull);
			}
			else m_imageList.Add(hLnkFolder);

			// Color
			if (vecFilter[i].clrBkg != NULL)
			{
				m_list_view.SetRowColors(i, vecFilter[i].clrBkg, RGBBLACK);
				m_list_view.SetCellColors(i, iGFLPath, vecFilter[i].clrBkg, RGBBLUE);
			}

			if (iBegin > 0 && dem == ncount)
			{
				m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
			}
		}

		// Load Image vào list dữ liệu
		m_list_view.SetImageList(&m_imageList);

		for (int i = iBegin; i < iEnd; i++)
		{
			// Load Image icon
			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, i + 1, 0, 0, 0);
		}

		m_list_view.SetColumnWidth(iGFLNum, LVSCW_AUTOSIZE_USEHEADER);
		m_list_view.SetColumnWidth(iGFLPath, LVSCW_AUTOSIZE_USEHEADER);
		iwColNum = m_list_view.GetColumnWidth(iGFLNum);
	}
	catch (...) {}
}

void FTab1::OnBnClickedBtnSearch()
{
	// Lọc dữ liệu
	tongfilter = _LocDulieuNangcao();
	
	// Load dữ liệu lên listview
	_LoadListMyFiles();

	m_txt_search.SetSel(0, -1);
}

void FTab1::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0)
	{
		if (nSubItem == iGFLPath) OnNMDblclkListView(pNMHDR, pResult);
	}
}

void FTab1::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
}

void FTab1::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szval = m_list_view.GetItemText(nItem, iGFLPath);
		if (szval != L"")
		{
			ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
		}
	}
}

void FTab1::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = COLW_FIX;
	else if (pSubItem == iGFLNum) pNMListHeader->pitem->cxy = iwColNum;
	*pResult = 0;
}

void FTab1::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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
			long iz = ibuocnhay * lanshow;

			if (tongfilter <= iz)
			{
				iz = tongfilter;
				iStopload = 0;
			}
			else iStopload = 1;

			// Thêm dữ liệu vào list kết quả
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz);

			m_list_view.EnsureVisible(ncount + (int)(b / a) - 5, FALSE);
		}
	}
	catch (...) {}
}


void FTab1::OnOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void FTab1::OnOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}

void FTab1::OnAddfiles()
{
	try
	{
		CMultiFileOpenDialog dlgOpenFile(TRUE);
		dlgOpenFile.GetOFN().Flags |= OFN_ALLOWMULTISELECT;
		dlgOpenFile.SetUseExtBuffer(TRUE);
		if (IDOK == dlgOpenFile.DoModal())
		{
			MyStrList MSLT;
			vector< MyStrList> vecIns;

			int icheck = -1;
			CString szval = L"";
			POSITION nPos = dlgOpenFile.GetStartPosition();
			while (NULL != nPos)
			{
				icheck = 0;
				szval = dlgOpenFile.GetNextPathName(nPos);

				for (int i = 0; i < (int)vecItem.size(); i++)
				{
					if (szval == vecItem[i].szFullpath)
					{
						icheck = 1;
						break;
					}
				}

				if (icheck == 0)
				{
					MSLT.iID = 0;
					MSLT.clrBkg = NULL;
					MSLT.szFullpath = szval;
					vecIns.push_back(MSLT);
				}				
			}

			if ((int)vecIns.size() > 0)
			{
				int nIndex = 0;
				if (nItem >= 0) nIndex = _wtoi(m_list_view.GetItemText(nItem, 0));
				vecItem.insert(vecItem.begin() + nIndex, vecIns.begin(), vecIns.end());
			}

			vecIns.clear();
			
			_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
			OnBnClickedBtnSearch();		// Cập nhật lại list
		}

		AfxGetMainWnd()->ShowWindow(SW_SHOW);
		AfxGetMainWnd()->ActivateTopParent();
	}
	catch (...) {}
}

void FTab1::OnDeletefiles()
{
	try
	{
		vector<int> vecRow;
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 0)
		{
			int nIndex = -1;
			int ncount = m_list_view.GetItemCount();
			if (nSelect < ncount)
			{
				CString szpath = L"";
				for (int i = nSelect - 1; i >= 0; i--)
				{
					szpath = m_list_view.GetItemText(vecRow[i], iGFLPath);
					if (szpath != L"")
					{
						// ID
						nIndex = _wtoi(m_list_view.GetItemText(vecRow[i], 0));
						vecItem.erase(vecItem.begin() + nIndex);
					}
				}

				_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector

				OnBnClickedBtnSearch();		// Cập nhật lại list
			}
			else
			{
				ncount = (int)vecFilter.size();
				for (int i = ncount - 1; i >= 0; i--)
				{
					vecItem.erase(vecItem.begin() + vecFilter[i].iID);
				}

				_ClearAllItems();			// Xóa toàn bộ dữ liệu
				_ClearAllImages();			// Xóa toàn bộ list-images

				_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
			}
		}
	}
	catch (...) {}
}

void FTab1::_XacdinhLaidulieu()
{
	try
	{
		MyStrList MSLT;
		vector<MyStrList> vecLuu;
		int ncount = (int)vecItem.size();
		for (int i = 0; i < ncount; i++) vecLuu.push_back(vecItem[i]);

		vecItem.clear();

		int nIndex = -1;
		CString szval = L"";
		ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			nIndex = _wtoi(m_list_view.GetItemText(i, 0));

			MSLT.iID = i;
			MSLT.szFullpath = vecLuu[nIndex].szFullpath;
			MSLT.clrBkg = vecLuu[nIndex].clrBkg;

			vecItem.push_back(MSLT);

			// ID
			szval.Format(L"%d", i);
			m_list_view.SetItemText(i, 0, szval);

			// STT
			szval.Format(L"%d", i + 1);
			m_list_view.SetItemText(i, iGFLNum, szval);

			// Color
			if (MSLT.clrBkg != NULL)
			{
				m_list_view.SetRowColors(i, MSLT.clrBkg, RGBBLACK);
				m_list_view.SetCellColors(i, iGFLPath, MSLT.clrBkg, RGBBLUE);
			}

			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, i + 1, 0, 0, 0);
		}

		vecLuu.clear();
	}
	catch (...) {}
}

void FTab1::_SaveAllVectorItems()
{
	try
	{
		vector<CString> vecText;

		tongkq = (int)vecItem.size();
		if (tongkq > 0)
		{
			CString szval = L"", szcolor = L"";
			for (int i = 0; i < tongkq; i++)
			{
				szval = vecItem[i].szFullpath;
				if (vecItem[i].clrBkg != NULL)
				{
					szcolor = ff->_GetColorRGB(vecItem[i].clrBkg);
					if (szcolor != L"RGB(255,255,255)")
					{
						szval += L"|";
						szval += szcolor;
					}					
				}
				vecText.push_back(szval);
			}
		}

		if ((int)vecText.size() == 0) vecText.push_back(L"");
		ff->_WriteMultiUnicode(vecText, pathFolderDir + FileMyFiles, 0);

		vecText.clear();
	}
	catch (...) {}
}

void FTab1::OnBnClickedBtnAdd()
{
	OnAddfiles();
}

bool FTab1::_IsFilterData()
{
	if ((int)vecFilter.size() < (int)vecItem.size())
	{
		MessageBox(L"Vui lòng bỏ lọc trước khi thay đổi vị trí dữ liệu.", 
			L"Hướng dẫn", MB_ICONERROR | MB_OK);

		AfxGetMainWnd()->ShowWindow(SW_SHOW);
		AfxGetMainWnd()->ActivateTopParent();

		return true;
	}
	return false;
}

void FTab1::OnMoveup()
{
	if (!_IsFilterData())
	{
		if (nItem > 0)
		{
			UpdateData();

			// Name
			CString szval = L"";
			int iArr[] = { iGFLTen, iGFLPath };
			int ncount = sizeof(iArr) / sizeof(iArr[0]);

			// Color
			int nIndex0 = _wtoi(m_list_view.GetItemText(nItem, 0));
			COLORREF clr0 = vecItem[nIndex0].clrBkg;
			if (clr0 == NULL) clr0 = RGBWHITE;

			int nIndex1 = _wtoi(m_list_view.GetItemText(nItem - 1, 0));
			COLORREF clr1 = vecItem[nIndex1].clrBkg;
			if (clr1 == NULL) clr1 = RGBWHITE;

			vecItem[nIndex0].szFullpath = m_list_view.GetItemText(nItem - 1, iGFLPath);
			vecItem[nIndex1].szFullpath = m_list_view.GetItemText(nItem, iGFLPath);

			vecItem[nIndex0].clrBkg = clr1;
			vecItem[nIndex1].clrBkg = clr0;

			for (int i = 0; i < ncount; i++)
			{
				szval = m_list_view.GetItemText(nItem, iArr[i]);
				m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem - 1, iArr[i]));
				m_list_view.SetItemText(nItem - 1, iArr[i], szval);
			}

			m_list_view.SetRowColors(nItem, clr1, RGBBLACK);
			m_list_view.SetRowColors(nItem - 1, clr0, RGBBLACK);
			
			m_list_view.SetCellColors(nItem, iGFLPath, clr1, RGBBLUE);
			m_list_view.SetCellColors(nItem - 1, iGFLPath, clr0, RGBBLUE);

			m_list_view.SetItem(nItem, 0, LVIF_IMAGE, NULL, nItem + 1, 0, 0, 0);
			m_list_view.SetItem(nItem - 1, 0, LVIF_IMAGE, NULL, nItem, 0, 0, 0);

			// Vị trí lựa chọn mới
			m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
			m_list_view.SetItemState(nItem - 1, LVIS_SELECTED, LVIS_SELECTED);

			nItem--;

			_LoadListImages();			// Load lại list-images
			_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
		}
	}
}

void FTab1::OnMovedown()
{
	if (!_IsFilterData())
	{
		int ncount = m_list_view.GetItemCount();
		if (nItem < ncount - 1)
		{
			UpdateData();

			// Name
			CString szval = L"";
			int iArr[] = { iGFLTen, iGFLPath };
			ncount = sizeof(iArr) / sizeof(iArr[0]);

			// Color
			int nIndex0 = _wtoi(m_list_view.GetItemText(nItem, 0));
			COLORREF clr0 = vecItem[nIndex0].clrBkg;
			if (clr0 == NULL) clr0 = RGBWHITE;

			int nIndex1 = _wtoi(m_list_view.GetItemText(nItem + 1, 0));
			COLORREF clr1 = vecItem[nIndex1].clrBkg;
			if (clr1 == NULL) clr1 = RGBWHITE;

			vecItem[nIndex0].szFullpath = m_list_view.GetItemText(nItem + 1, iGFLPath);
			vecItem[nIndex1].szFullpath = m_list_view.GetItemText(nItem, iGFLPath);

			vecItem[nIndex0].clrBkg = clr1;
			vecItem[nIndex1].clrBkg = clr0;

			for (int i = 0; i < ncount; i++)
			{
				szval = m_list_view.GetItemText(nItem, iArr[i]);
				m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem + 1, iArr[i]));
				m_list_view.SetItemText(nItem + 1, iArr[i], szval);
			}

			m_list_view.SetRowColors(nItem, clr1, RGBBLACK);
			m_list_view.SetRowColors(nItem + 1, clr0, RGBBLACK);

			m_list_view.SetCellColors(nItem, iGFLPath, clr1, RGBBLUE);
			m_list_view.SetCellColors(nItem + 1, iGFLPath, clr0, RGBBLUE);
			
			m_list_view.SetItem(nItem + 1, 0, LVIF_IMAGE, NULL, nItem + 2, 0, 0, 0);
			m_list_view.SetItem(nItem, 0, LVIF_IMAGE, NULL, nItem + 1, 0, 0, 0);

			// Vị trí lựa chọn mới
			m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
			m_list_view.SetItemState(nItem + 1, LVIS_SELECTED, LVIS_SELECTED);

			nItem++;

			_LoadListImages();			// Load lại list-images
			_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
		}
	}
}

void FTab1::OnReverse()
{
	if (!_IsFilterData())
	{
		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int dem = 0;
			int iItems[2] = { 0, 0 };
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iItems[dem] = i;
					if (dem == 1) break;
					dem++;
				}
			}

			if (dem != 1) return;

			UpdateData();

			// Name
			CString szval = L"";
			int iArr[] = { iGFLTen, iGFLPath };
			ncount = sizeof(iArr) / sizeof(iArr[0]);

			// Color
			int nIndex0 = _wtoi(m_list_view.GetItemText(iItems[0], 0));
			COLORREF clr0 = vecItem[nIndex0].clrBkg;
			if (clr0 == NULL) clr0 = RGBWHITE;

			int nIndex1 = _wtoi(m_list_view.GetItemText(iItems[1], 0));
			COLORREF clr1 = vecItem[nIndex1].clrBkg;
			if (clr1 == NULL) clr1 = RGBWHITE;

			vecItem[nIndex0].szFullpath = m_list_view.GetItemText(iItems[1], iGFLPath);
			vecItem[nIndex1].szFullpath = m_list_view.GetItemText(iItems[0], iGFLPath);

			vecItem[nIndex0].clrBkg = clr1;
			vecItem[nIndex1].clrBkg = clr0;

			for (int i = 0; i < ncount; i++)
			{
				szval = m_list_view.GetItemText(iItems[0], iArr[i]);
				m_list_view.SetItemText(iItems[0], iArr[i], m_list_view.GetItemText(iItems[1], iArr[i]));
				m_list_view.SetItemText(iItems[1], iArr[i], szval);
			}

			m_list_view.SetRowColors(iItems[0], clr1, RGBBLACK);
			m_list_view.SetRowColors(iItems[1], clr0, RGBBLACK);

			m_list_view.SetCellColors(iItems[0], iGFLPath, clr1, RGBBLUE);
			m_list_view.SetCellColors(iItems[1], iGFLPath, clr0, RGBBLUE);

			m_list_view.SetItem(iItems[0], 0, LVIF_IMAGE, NULL, iItems[0] + 1, 0, 0, 0);
			m_list_view.SetItem(iItems[1], 0, LVIF_IMAGE, NULL, iItems[1] + 1, 0, 0, 0);

			// Vị trí lựa chọn mới
			m_list_view.SetItemState(iItems[0], 0, LVIS_SELECTED);
			m_list_view.SetItemState(iItems[1], LVIS_SELECTED, LVIS_SELECTED);

			_LoadListImages();			// Load lại list-images
			_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
		}
	}
}

void FTab1::OnBgkcolor()
{
	UpdateData();

	if (nItem >= 0 && nSubItem >= 0)
	{
		vector<int> vecRow;

		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 0)
		{
			CColorDialog dlg;
			if (dlg.DoModal() == IDOK)
			{
				COLORREF clr = dlg.GetColor();
				if (clr != NULL)
				{
					int nIndex = -1;
					for (int i = 0; i < nSelect; i++)
					{
						m_list_view.SetRowColors(vecRow[i], clr, RGBBLACK);
						m_list_view.SetCellColors(vecRow[i], iGFLPath, clr, RGBBLUE);

						nIndex = _wtoi(m_list_view.GetItemText(vecRow[i], 0));
						if (clr != RGBWHITE) vecItem[nIndex].clrBkg = clr;
						else vecItem[nIndex].clrBkg = NULL;						
					}

					_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
				}
			}			
		}

		vecRow.clear();
	}	
}

void FTab1::OnTimer(UINT_PTR nIDEvent)
{
	// Chú ý: Giá trị 'CheckAddFiles' --> Liên hệ với DLL sl (GMSelect.exe)
	reg->CreateKey(szCheckAddFiles);
	int iCheckAdd = _wtoi(reg->ReadString(L"CheckAddFiles", L""));
	if (iCheckAdd != 0)
	{
		reg->WriteChar("CheckAddFiles", "0");
		m_txt_search.SetWindowText(L"");
		OnBnClickedBtnSearch();

		AfxGetMainWnd()->ShowWindow(SW_SHOW);
		AfxGetMainWnd()->ActivateTopParent();
	}
	else
	{
		if (blDragging)
		{
			blDragging = false;

			if (!_IsFilterData())
			{
				UpdateData();

				_LoadListImages();			// Load lại list-images
				_XacdinhLaidulieu();		// Xác định lại dữ liệu thay đổi
				_SaveAllVectorItems();		// Lưu dữ liệu mới thêm vào vector
			}
		}
	}	

	CDialogEx::OnTimer(nIDEvent);
}


void FTab1::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_DELETE) OnDeletefiles();
	*pResult = 0;
}
