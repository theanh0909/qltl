#include "pch.h"
#include "DlgRenameFiles.h"
#include "DlgSelectionCells.h"

IMPLEMENT_DYNAMIC(DlgRenameFiles, CDialogEx)

DlgRenameFiles::DlgRenameFiles(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgRenameFiles::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;

	iStopload = 1, lanshow = 1, ibuocnhay = 50;
	tongkq = 0, tongfilter = 0;

	iGFLType = 1, iGFLNmNew = 2, iGFLNmOld = 3, iGFLPath = 4;

	jColumnHand = iGFLPath;

	nItem = -1, nSubItem = -1;
	iKeyESC = 0;

	pathFolderDir = L"";
	szBefore = L"", szAfter = L"";

	bLookInSubdirectories = true;

	nTotalCol = getSizeArray(iwCol);
	iwCol[1] = 120, iwCol[2] = 150, iwCol[3] = 150, iwCol[4] = 200;
	iDlgW = 0, iDlgH = 0;	
}

DlgRenameFiles::~DlgRenameFiles()
{
	delete ff;
	delete bb;
	delete reg;

	vecItem.clear();
	vecFilter.clear();

	jColumnHand = -1;
}

void DlgRenameFiles::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, RNMF_BTN_OK, m_btn_ok);
	DDX_Control(pDX, RNMF_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, RNMF_BTN_SEARCH, m_btn_search);
	DDX_Control(pDX, RNMF_CHECK_ALL, m_check_all);
	DDX_Control(pDX, RNMF_CHECK_NEWS, m_check_new);
	DDX_Control(pDX, RNMF_LIST_VIEW, m_list_view);
	DDX_Control(pDX, RNMF_EDIT_VIEW, m_edit_rename);
	DDX_Control(pDX, RNMF_LBL_TOTALCHECK, m_lbl_total);
	DDX_Control(pDX, RNMF_TXT_TOTALFILES, m_edtTotalFiles);
	DDX_Control(pDX, RNMF_TXT_SEARCH, m_editSearchFiles);
	DDX_Control(pDX, RNMF_TXT_TYPEFILE, m_edtWildCard);
	DDX_Control(pDX, RNMF_LBL_HELP, m_lbl_help);
}

BEGIN_MESSAGE_MAP(DlgRenameFiles, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(RNMF_BTN_OK, &DlgRenameFiles::OnBnClickedBtnOk)
	ON_BN_CLICKED(RNMF_BTN_CANCEL, &DlgRenameFiles::OnBnClickedBtnCancel)
	ON_BN_CLICKED(RNMF_CHECK_ALL, &DlgRenameFiles::OnBnClickedCheckAll)
	ON_BN_CLICKED(RNMF_CHECK_NEWS, &DlgRenameFiles::OnBnClickedCheckNews)
	ON_BN_CLICKED(RNMF_BTN_SEARCH, &DlgRenameFiles::OnBnClickedBtnSearch)
	
	ON_EN_SETFOCUS(RNMF_TXT_SEARCH, &DlgRenameFiles::OnEnSetfocusTxtSearch)
	ON_EN_SETFOCUS(RNMF_TXT_TYPEFILE, &DlgRenameFiles::OnEnSetfocusTxtTypefile)
	ON_EN_SETFOCUS(RNMF_EDIT_VIEW, &DlgRenameFiles::OnEnSetfocusEditView)
	ON_EN_KILLFOCUS(RNMF_EDIT_VIEW, &DlgRenameFiles::OnEnKillfocusEditView)

	ON_COMMAND(MN_RENAME_RFILES, &DlgRenameFiles::OnRenameRfiles)
	ON_COMMAND(MN_RENAME_OPENFILE, &DlgRenameFiles::OnRenameOpenfile)
	ON_COMMAND(MN_RENAME_OPENFOLDER, &DlgRenameFiles::OnRenameOpenfolder)
	ON_COMMAND(MN_RENAME_DELETEFILES, &DlgRenameFiles::OnRenameDeletefiles)
	ON_COMMAND(MN_RENAME_CHECK, &DlgRenameFiles::OnRenameCheck)
	ON_COMMAND(MN_RENAME_UNCHECK, &DlgRenameFiles::OnRenameUncheck)
	ON_COMMAND(MN_RENAME_CHECKALL, &DlgRenameFiles::OnRenameCheckall)
	ON_COMMAND(MN_RENAME_COPY, &DlgRenameFiles::OnRenameCopy)

	ON_NOTIFY(NM_CLICK, RNMF_LIST_VIEW, &DlgRenameFiles::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, RNMF_LIST_VIEW, &DlgRenameFiles::OnNMRClickListView)
	ON_NOTIFY(NM_DBLCLK, RNMF_LIST_VIEW, &DlgRenameFiles::OnNMDblclkListView)
	ON_NOTIFY(LVN_KEYDOWN, RNMF_LIST_VIEW, &DlgRenameFiles::OnLvnKeydownListView)
	ON_NOTIFY(LVN_ENDSCROLL, RNMF_LIST_VIEW, &DlgRenameFiles::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgRenameFiles::OnHdnEndtrackListView)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgRenameFiles)
	EASYSIZE(RNMF_LBL_SEARCH, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_TXT_SEARCH, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_TXT_TYPEFILE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_BTN_SEARCH, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_CHECK_NEWS, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_LBL_TOTALCHECK, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_LBL_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_TXT_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(RNMF_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(RNMF_LBL_HELP, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(RNMF_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(RNMF_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgRenameFiles message handlers
BOOL DlgRenameFiles::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_RENAME));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_SetDefault();
	_LocdulieuNangcao();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgRenameFiles::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(RNMF_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(RNMF_TXT_SEARCH))
			{
				OnBnClickedBtnSearch();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(RNMF_TXT_TYPEFILE))
			{
				OnBnClickedBtnSearch();
				GotoDlgCtrl(GetDlgItem(RNMF_TXT_SEARCH));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			if (iKeyESC == 0) OnBnClickedBtnCancel();
			else
			{
				szBefore = m_list_view.GetItemText(CLRow, CLColumn);
				m_edit_rename.SetWindowTextW(szBefore);
				GotoDlgCtrl(GetDlgItem(RNMF_LIST_VIEW));
			}
			iKeyESC = 0;

			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(RNMF_LIST_VIEW))
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

void DlgRenameFiles::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgRenameFiles::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgRenameFiles::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgRenameFiles::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_RENAME));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;

			vector<int> vecRow;

			CString szval = L"";
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				szval.Format(L"Đổi tên (%d) files", nSelect);
				ModifyMenu(pMenu, MN_RENAME_RFILES, MF_BYCOMMAND | MF_STRING, MN_RENAME_RFILES, szval);
			}

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_RFILES, IDI_ICON_UPDATE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_COPY, IDI_ICON_COPYCLIP);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_DELETEFILES, IDI_ICON_DELETE);

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_CHECK, IDI_ICON_OK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_UNCHECK, IDI_ICON_CANCEL);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_RENAME_CHECKALL, IDI_ICON_TODOLIST);

			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_RENAME_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_RENAME_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(RNMF_LIST_VIEW));
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

/****************************************************************************************/

void DlgRenameFiles::_GetDanhsachFiles(int iSelectCell)
{
	try
	{
		if (iSelectCell >= 0)
		{
			ff->_xlPutScreenUpdating(false);

			_XacdinhSheetLienquan();

			if (iSelectCell == 0)
			{
				int ncount = (int)vecSelect.size();
				irowStart = vecSelect[0].row;
				irowEnd = vecSelect[ncount - 1].row + 1;
			}
			else if (iSelectCell == 1)
			{
				_GetGroupRowStartEnd(irowStart, irowEnd);
			}

			RangePtr PRS;
			MyStrFiles MSFILE;

			int id = 0, pos = -1, dem = 0;
			int ncount = irowEnd - irowStart;

			CString szval = L"", szlink = L"";
			CString szUpdateStatus = L"Đang tiến hành kiểm tra files. Vui lòng đợi trong giây lát...";

			for (int i = irowStart; i < irowEnd; i++)
			{
				dem++;
				szval.Format(L"%s (%d%s) %s", szUpdateStatus,
					(int)(100 * dem / ncount), L"%", szlink);
				if (szval.GetLength() > 250) szval = szval.Left(250) + L"...";
				ff->_xlPutStatus(szval);

				PRS = pRgActive->GetItem(i, icotFilegoc);
				szval = ff->_xlGetHyperLink(PRS);				
				if (szval != L"")
				{
					MSFILE.szHyperlink = szval;
					MSFILE.sznmOld = ff->_GetNameOfPath(MSFILE.szHyperlink, pos, -1);

					if (jcotNoidung > 0) MSFILE.sznmNew = ff->_xlGIOR(pRgActive, i, jcotTen, L"");
					else MSFILE.sznmNew = MSFILE.sznmOld;

					MSFILE.sznmNew = ff->_ConvertRename(MSFILE.sznmNew);
					MSFILE.iEnable = 1;
					MSFILE.iRow = i;
					MSFILE.iID = id;
					id++;

					vecItem.push_back(MSFILE);
				}
			}

			ff->_xlPutStatus(L"Ready");
			ff->_xlPutScreenUpdating(true);

			tongkq = (long)vecItem.size();
			if (tongkq > 0)
			{
				// Hiển thị hộp thoại
				xl->EnableCancelKey = XlEnableCancelKey(FALSE);
				AFX_MANAGE_STATE(AfxGetStaticModuleState());
				DoModal();
			}
			else
			{
				ff->_msgbox(L"Không tồn tại dữ liệu chỉnh sửa. Vui lòng kiểm tra lại.", 1, 0);
			}
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_GetGroupRowStartEnd(int &iRowStart, int &iRowEnd)
{
	try
	{
		CString szval = L"";
		for (int i = iRowActive; i >= iRowStart; i--)
		{
			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) == 0)
			{
				szval = ff->_xlGIOR(pRgActive, i, jcotLV, L"");
				if (szval != L"")
				{
					iRowStart = i;
					for (int j = iRowActive + 1; j <= iRowEnd; j++)
					{
						szval = ff->_xlGIOR(pRgActive, j, jcotSTT, L"");
						if (_wtoi(szval) == 0)
						{
							szval = ff->_xlGIOR(pRgActive, j, jcotLV, L"");
							if (szval != L"")
							{
								iRowEnd = j;
								return;
							}
						}
					}
				}
			}
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_LocdulieuNangcao()
{
	try
	{
		iKeyESC = 0;
		vecFilter.clear();

		long nsize = (long)vecItem.size();
		int icheck = m_check_new.GetCheck();
		if (icheck == 1)
		{
			for (long i = 0; i < nsize; i++)
			{
				if (vecItem[i].sznmNew != vecItem[i].sznmOld)
				{
					vecFilter.push_back(vecItem[i]);
				}
			}
		}
		else
		{
			for (long i = 0; i < tongkq; i++) vecFilter.push_back(vecItem[i]);
		}

		// Lọc dữ liệu theo đuôi file tìm kiếm
		_FilterTypeOfFile(vecFilter);

		// Lọc dữ liệu theo từ khóa tìm kiếm
		_FilterSearchText(vecFilter);

/****** Load dữ liệu *********************************************************/

		tongfilter = (long)vecFilter.size();

		m_check_all.SetCheck(1);

		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);

		// Delete list image icon
		m_imageList.DeleteImageList();
		ASSERT(m_imageList.GetSafeHandle() == NULL);
		m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);

		lanshow = 1;
		long iz = ibuocnhay;
		if (tongfilter <= iz)
		{
			iz = tongfilter;
			iStopload = 0;
		}
		else iStopload = 1;

		// Thêm dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz, (int)m_check_all.GetCheck());

		_SetStatusKetqua(tongfilter);

		CString szval = L"";
		szval.Format(L"%ld", tongfilter);
		m_edtTotalFiles.SetWindowText(ff->_FormatTienVND(szval, L",", L"."));
		m_editSearchFiles.SetSel(0, -1);
	}
	catch (...) {}
}

void DlgRenameFiles::_AddItemsListKetqua(int iBegin, int iEnd, int icheck)
{
	try
	{
		int dem = 0, ncount = 0;
		if (iBegin > 0)
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
		HICON hLnkNull = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_NULL));

		int pos = -1;
		CString szval = L"";
		CString szSize = L"", szNgay = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, iGFLNmNew, vecFilter[i].sznmNew);
			m_list_view.SetItemText(i, iGFLNmOld, vecFilter[i].sznmOld);
			m_list_view.SetItemText(i, iGFLPath, vecFilter[i].szHyperlink);
			
			szval = ff->_GetTypeOfFile(vecFilter[i].szHyperlink);
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
			

			if (vecFilter[i].sznmNew == L"")
			{
				m_list_view.SetItemText(CLRow, iGFLType, L"Tên trống");
				m_list_view.SetCellColors(CLRow, iGFLType, rgbWhite, rgbRed);
			}
			else
			{
				if (vecFilter[i].sznmNew != vecFilter[i].sznmOld)
				{
					m_list_view.SetItemText(i, iGFLType, L"Đang sửa");
					m_list_view.SetCellColors(i, iGFLType, RGB(0, 112, 192), rgbWhite);
				}
			}

			m_list_view.SetCheck(i, icheck);

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
			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, i, 0, 0, 0);
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_SetDefault()
{
	try
	{
		m_lbl_total.SubclassDlgItem(RNMF_LBL_TOTALCHECK, this);
		m_lbl_total.SetTextColor(rgbBlue);

		m_lbl_help.SubclassDlgItem(RNMF_LBL_HELP, this);
		m_lbl_help.SetTextColor(rgbBlue);

		m_check_all.SetCheck(1);

		m_edtTotalFiles.SubclassDlgItem(RNMF_TXT_TOTALFILES, this);
		m_edtTotalFiles.SetBkColor(rgbWhite);
		m_edtTotalFiles.SetTextColor(rgbRed);

		m_editSearchFiles.SetCueBanner(L"Nhập từ khóa tìm kiếm...");
		m_editSearchFiles.SubclassDlgItem(RNMF_TXT_SEARCH, this);
		m_editSearchFiles.SetBkColor(rgbWhite);
		m_editSearchFiles.SetTextColor(rgbBlue);

		m_edtWildCard.SubclassDlgItem(RNMF_TXT_TYPEFILE, this);
		m_edtWildCard.SetBkColor(rgbWhite);
		m_edtWildCard.SetTextColor(rgbBlue);
		m_edtWildCard.SetCueBanner(L"Một hoặc nhiều đuôi files. Ví dụ: xlsx doc,...");

		m_edit_rename.SubclassDlgItem(RNMF_EDIT_VIEW, this);
		m_edit_rename.SetBkColor(rgbWhite);
		m_edit_rename.SetTextColor(rgbBlue);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, 40);
		m_list_view.InsertColumn(iGFLType, L" Trạng thái", LVCFMT_CENTER, iwCol[1]);
		m_list_view.InsertColumn(iGFLNmNew, L" Tên thay thế", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLNmOld, L" Tên file gốc", LVCFMT_LEFT, iwCol[3]);
		m_list_view.InsertColumn(iGFLPath, L" Đường dẫn gốc", LVCFMT_LEFT, iwCol[4]);

		m_list_view.SetColumnEditor(iGFLNmNew, &m_edit_rename);
		m_list_view.SetColumnColors(iGFLPath, rgbWhite, rgbBlue);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);

/**********************************************************************************/

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_search.SetIcon(IDI_ICON_SEARCH, 24, 24);
		m_btn_search.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_search.DrawTransparent(true);
		m_btn_search.SetBtnCursor(IDC_CURSOR_HAND);

/**********************************************************************************/

	}
	catch (...) {}
}

void DlgRenameFiles::OnBnClickedBtnOk()
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

		if (ncheck > 0) _ClickRenameFiles(1);
		else ff->_msgbox(L"Bạn chưa tích chọn dữ liệu cần thêm.", 2, 0);
	}
	catch (...) {}
}

void DlgRenameFiles::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

void DlgRenameFiles::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + RenameFiles;
		reg->CreateKey(szCreate);

		CString szval = L"", szcol = L"";
		for (int i = 0; i < nTotalCol; i++)
		{
			szcol.Format(L"%d|", m_list_view.GetColumnWidth(i));
			szval += szcol;
		}

		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgRenameFiles::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + RenameFiles;
		reg->CreateKey(szCreate);

		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			vector<CString> vecKey;
			ff->_TackTukhoa(vecKey, szval, L"|", L";", 0, 1);

			int ncount = (int)vecKey.size();
			if (ncount < nTotalCol)
			{
				for (int i = ncount; i < nTotalCol; i++) vecKey.push_back(L"");
			}

			for (int i = 0; i < nTotalCol; i++) iwCol[i] = _wtoi(vecKey[i]);

			vecKey.clear();
		}

		szval = reg->ReadString(DlgSize, L"");
		if (szval != L"")
		{
			int pos = szval.Find(L"x");
			if (pos > 0)
			{
				iDlgW = _wtoi((szval.Left(pos)).Trim());
				iDlgH = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_SetLocationAndSize()
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

void DlgRenameFiles::_XacdinhSheetLienquan()
{
	try
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
	catch (...) {}
}

void DlgRenameFiles::_SelectAllItems()
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

void DlgRenameFiles::_SetStatusKetqua(int num)
{
	CString str = L"";
	if (num > 1) str.Format(L"Bạn đã chọn (%d) files / ", num);
	else str.Format(L"Bạn đã chọn (%d) file / ", num);
	m_lbl_total.SetWindowText(str);
}

void DlgRenameFiles::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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
			int icheck = (int)m_check_all.GetCheck();
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz, icheck);

			m_list_view.EnsureVisible(ncount + (int)(b / a) - 5, FALSE);
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_FilterTypeOfFile(vector<MyStrFiles> &vecFt)
{
	try
	{
		if (tongkq == 0) return;

		CString strWildCard = L"";
		m_edtWildCard.GetWindowText(strWildCard);
		strWildCard.Replace(L"*", L"");
		strWildCard.Replace(L".", L"");
		strWildCard.Replace(L" ", L"");
		strWildCard.MakeLower();
		strWildCard.Trim();

		if (strWildCard != L"")
		{
			CString szType = L"";
			for (long i = 0; i < tongkq; i++)
			{
				szType = ff->_GetTypeOfFile(vecItem[i].szHyperlink).MakeLower();
				if (szType != L"")
				{
					if (strWildCard.Find(szType) >= 0) vecFt.push_back(vecItem[i]);
				}
			}
		}
	}
	catch (...) {}
}

void DlgRenameFiles::_FilterSearchText(vector<MyStrFiles> &vecFt)
{
	try
	{
		tongfilter = (long)vecFilter.size();
		if (tongfilter == 0) return;

		CString szTukhoa = L"";
		m_editSearchFiles.GetWindowText(szTukhoa);
		szTukhoa.Trim();
		if (szTukhoa != L"")
		{
			szTukhoa = ff->_ConvertKhongDau(szTukhoa);
			szTukhoa.MakeLower();

			vector<CString> vecKey;
			int iKey = ff->_TackTukhoa(vecKey, szTukhoa, L" ", L"+", 1, 0);

			int icheck = 0;
			CString szval = L"";
			for (long i = tongfilter - 1; i >= 0; i--)
			{
				szval = ff->_ConvertKhongDau(vecFt[i].sznmOld);
				szval.MakeLower();

				icheck = 0;
				for (int j = 0; j < iKey; j++)
				{
					if (szval.Find(vecKey[j]) == -1) break;
					icheck++;
				}

				if (icheck < iKey) vecFt.erase(vecFt.begin() + i);
			}

			vecKey.clear();
		}		
	}
	catch (...) {}
}


void DlgRenameFiles::OnBnClickedCheckAll()
{
	iKeyESC = 0;
	int icheck = m_check_all.GetCheck();
	int ncount = m_list_view.GetItemCount();
	for (int i = 0; i < ncount; i++) m_list_view.SetCheck(i, icheck);

	if (tongfilter > ncount)
	{
		for (int i = ncount; i < tongfilter; i++) vecFilter[i].iEnable = icheck;
	}

	if (icheck == 1) _SetStatusKetqua(tongfilter);
	else _SetStatusKetqua(0);
}

void DlgRenameFiles::OnBnClickedCheckNews()
{
	_LocdulieuNangcao();
}

void DlgRenameFiles::OnBnClickedBtnSearch()
{
	_LocdulieuNangcao();
}

void DlgRenameFiles::OnEnSetfocusTxtSearch()
{
	iKeyESC = 0;
}

void DlgRenameFiles::OnEnSetfocusTxtTypefile()
{
	iKeyESC = 0;
}

void DlgRenameFiles::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0)
	{
		if (nSubItem == 0)
		{
			bool bl = !m_list_view.GetCheck(nItem);
			int dem = _GetCountItemsCheck();

			if (bl) dem++;
			else dem--;

			_SetStatusKetqua(dem);
		}
		else if (nSubItem == iGFLPath) OnNMDblclkListView(pNMHDR, pResult);
	}
}

void DlgRenameFiles::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}

void DlgRenameFiles::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szval = m_list_view.GetItemText(nItem, iGFLPath);
		if (szval != L"") ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void DlgRenameFiles::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_TAB || pLVKeyDow->wVKey == VK_F2)
	{
		if (nItem != -1 && nSubItem != -1)
		{
			CLRow = nItem, CLColumn = nSubItem;
			m_list_view.DisplayEditor(CLRow, CLColumn);
		}
	}
	else if (pLVKeyDow->wVKey == VK_DELETE) OnRenameDeletefiles();

	*pResult = 0;
}

void DlgRenameFiles::OnEnSetfocusEditView()
{
	iKeyESC = 1;
}

int DlgRenameFiles::_GetCountItemsCheck()
{
	try
	{
		int dem = 0;
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1) dem++;
		}

		if (tongfilter > ncount)
		{
			if ((int)m_check_all.GetCheck() == 1) dem += (tongfilter - ncount);
		}
		return dem;
	}
	catch (...) {}
	return 0;
}

void DlgRenameFiles::_GetItemSelect(int icheck)
{
	try
	{
		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int dem = 0;
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					m_list_view.SetCheck(i, icheck);
					dem++;
				}
			}

			if (dem == ncount)
			{
				if (icheck == 1) _SetStatusKetqua(tongfilter);
				else _SetStatusKetqua(0);
			}
			else
			{
				dem = _GetCountItemsCheck();
				_SetStatusKetqua(dem);
			}
		}
	}
	catch (...) {}
}

void DlgRenameFiles::OnEnKillfocusEditView()
{
	try
	{
		szBefore = m_list_view.GetItemText(CLRow, CLColumn);
		m_edit_rename.GetWindowTextW(szAfter);

		szAfter.Trim(); szBefore.Trim();	

		if (szAfter == szBefore) return;

		if (CLColumn == iGFLNmNew)
		{
			int nIndex = vecFilter[CLRow].iID;
			vecItem[nIndex].sznmNew = szAfter;

			if (szAfter != L"")
			{
				CString szNameOld = m_list_view.GetItemText(CLRow, iGFLNmOld);
				if (szNameOld != szAfter)
				{
					m_list_view.SetItemText(CLRow, iGFLType, L"Đang sửa");
					m_list_view.SetCellColors(CLRow, iGFLType, RGB(0, 112, 192), rgbWhite);
				}
				else
				{
					m_list_view.SetItemText(CLRow, iGFLType, L"");
					m_list_view.SetCellColors(CLRow, iGFLType, rgbWhite, rgbBlack);
				}
			}
			else
			{
				m_list_view.SetItemText(CLRow, iGFLType, L"Tên trống");
				m_list_view.SetCellColors(CLRow, iGFLType, rgbWhite, rgbRed);
			}
		}		
	}
	catch (...) {}
}


void DlgRenameFiles::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 40;
	*pResult = 0;
}

void DlgRenameFiles::_ClickRenameFiles(int iOnClickBtn)
{
	try
	{
		vector<int> vecRow;

		int nSelect = 0;
		if (iOnClickBtn == 0) nSelect = _GetAllSelectedItems(vecRow);
		else nSelect = _GetAllCheckedItems(vecRow);

		if (nSelect > 0)
		{
			if (iOnClickBtn == 1)
			{
				_SaveRegistry();
				CDialogEx::OnOK();
			}

			ff->_xlPutScreenUpdating(false);

			RangePtr PRS;
			int pos = -1, nIndex = -1;
			CString szval = L"", szdir = L"", sznameOld = L"", sznameNew = L"", sztype = L"";
			CString szUpdateStatus = L"Đang đổi tên files. Vui lòng đợi trong giây lát...";
			for (int i = 0; i < nSelect; i++)
			{
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / nSelect), L"%");
				ff->_xlPutStatus(szval);

				sznameNew = m_list_view.GetItemText(vecRow[i], iGFLNmNew);
				sznameNew = ff->_ConvertRename(sznameNew);

				sznameOld = m_list_view.GetItemText(vecRow[i], iGFLNmOld);
				if (sznameNew != L"" && sznameNew != sznameOld)
				{
					szdir = m_list_view.GetItemText(vecRow[i], iGFLPath);
					sztype = ff->_GetTypeOfFile(szdir);

					szval = ff->_GetNameOfPath(szdir, pos, 1);
					if (szval.Right(1) != L"\\") szval += L"\\";
					szval += (sznameNew + L"." + sztype);

					if ((int)_wrename(szdir, szval) != 0)
					{
						m_list_view.SetItemText(vecRow[i], iGFLType, L"Lỗi");
						m_list_view.SetCellColors(vecRow[i], iGFLType, RGB(192, 0, 0), rgbWhite);
					}
					else
					{
						m_list_view.SetItemText(vecRow[i], iGFLNmOld, sznameNew);
						m_list_view.SetItemText(vecRow[i], iGFLPath, szval);

						m_list_view.SetItemText(vecRow[i], iGFLType, L"Hoàn thành");
						m_list_view.SetCellColors(vecRow[i], iGFLType, RGB(0, 176, 80), rgbWhite);

						nIndex = vecFilter[vecRow[i]].iID;
						vecItem[nIndex].sznmOld = sznameNew;
						vecItem[nIndex].szHyperlink = szval;

						PRS = pRgActive->GetItem(vecItem[nIndex].iRow, icotFilegoc);
						ff->_xlSetHyperlink(pShActive, PRS, szval);

						if (jcotNoidung > 0)
						{
							pRgActive->PutItem(vecItem[nIndex].iRow, jcotTen, (_bstr_t)sznameNew);
						}
					}
				}
			}

			ff->_xlPutStatus(L"Ready");
			if (iOnClickBtn == 1) ff->_msgbox(L"Đổi tên files thành công!", 0, 1, 2000);
			ff->_xlPutScreenUpdating(true);
		}

		vecRow.clear();
	}
	catch (...) {}
}

int DlgRenameFiles::_GetAllSelectedItems(vector<int> &vecRow)
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

int DlgRenameFiles::_GetAllCheckedItems(vector<int> &vecRow)
{
	try
	{
		vecRow.clear();
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if (m_list_view.GetCheck(i)) vecRow.push_back(i);
		}

		if ((int)m_check_all.GetCheck() == 1)
		{
			if (tongfilter > ncount)
			{
				for (int i = ncount; i < tongfilter; i++) vecRow.push_back(i);
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void DlgRenameFiles::OnRenameRfiles()
{
	_ClickRenameFiles(0);
}

void DlgRenameFiles::OnRenameOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}


void DlgRenameFiles::OnRenameOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}


void DlgRenameFiles::OnRenameDeletefiles()
{
	try
	{
		vector<int> vecRow;
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 0)
		{
			CString szval = L"";
			szval.Format(L"Dữ liệu (%d) file", nSelect);
			if (nSelect > 1) szval += L"s";
			szval += L" sẽ bị xóa hoàn toàn khỏi máy tính. Bạn có chắc chắn xóa dữ liệu được chọn?";

			int result = ff->_y(szval, 0, 0, 0);
			if (result == 6)
			{
				CString szval = L"";
				m_edtTotalFiles.GetWindowTextW(szval);
				szval.Replace(L".", L"");
				long itotal = _wtol(szval);
				int dem = 0;

				RangePtr PRS;
				int nIndex = -1;
				for (int i = nSelect - 1; i >= 0; i--)
				{
					CString szpath = m_list_view.GetItemText(vecRow[i], iGFLPath);
					if (szpath != L"")
					{
						DeleteFile(szpath);
						m_list_view.DeleteItem(vecRow[i]);

						nIndex = vecFilter[vecRow[i]].iID;
						PRS = pRgActive->GetItem(vecItem[nIndex].iRow, 1);
						PRS->EntireRow->Delete(xlShiftUp);

						vecItem.erase(vecItem.begin() + nIndex);
						vecFilter.erase(vecFilter.begin() + vecRow[i]);
						for (int j = vecRow[i] + 1; j < (int)vecFilter.size(); j++) vecFilter[j].iID--;

						dem++;
					}
				}

				itotal -= dem;
				tongkq -= dem;
				tongfilter -= dem;
				
				if (itotal < 0) itotal = 0;
				szval.Format(L"%ld", itotal);
				m_edtTotalFiles.SetWindowText(ff->_FormatTienVND(szval, L",", L"."));

				m_lbl_total.GetWindowTextW(szval);
				szval.Replace(L"Bạn đã chọn (", L"");
				int pos = szval.Find(L")");
				if (pos > 0)
				{
					itotal = _wtoi(szval);
					itotal -= dem;
					_SetStatusKetqua(itotal);
				}
			}
		}

		vecRow.clear();
	}
	catch (...) {}
}


void DlgRenameFiles::OnRenameCheck()
{
	_GetItemSelect(1);
}


void DlgRenameFiles::OnRenameUncheck()
{
	_GetItemSelect(0);
}


void DlgRenameFiles::OnRenameCheckall()
{
	_SelectAllItems();
}


void DlgRenameFiles::OnRenameCopy()
{
	try
	{
		vector<int> vecRow;
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 0)
		{
			CString szval = L"";
			for (int i = 0; i < nSelect; i++)
			{
				szval = m_list_view.GetItemText(vecRow[i], iGFLNmOld);
				m_list_view.SetItemText(vecRow[i], iGFLNmNew, szval);

				m_list_view.SetItemText(vecRow[i], iGFLType, L"");
				m_list_view.SetCellColors(vecRow[i], iGFLType, rgbWhite, rgbBlack);
			}
		}
		vecRow.clear();
	}
	catch (...) {}
}

void DlgRenameFiles::_RightClickRenameFiles(int itype)
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		_XacdinhSheetLienquan();

		// Xác định vùng được chọn
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

			if (icheck == 0) vecRow.push_back(vecSelect[i].row);
		}

		isize = (int)vecRow.size();
		if (isize > 1) ff->_QuickSortArr(vecRow, 0, isize - 1);
		
		CString szval = L"";
		CString szUpdateStatus = L"Đang tiến hành đổi tên và cập nhật thông tin files. Vui lòng đợi trong giây lát...";
		for (int i = 0; i < isize; i++)
		{
			if (itype == 1)
			{
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / isize), L"%");
				ff->_xlPutStatus(szval);
			}			

			// Cập nhật tên & thông tin file 
			_UpdatePropertiesFile(vecRow[i], itype);
		}

		vecRow.clear();
		vecSelect.clear();

		ff->_xlPutStatus(L"Ready");
		ff->_xlPutScreenUpdating(true);
	}
	catch (...) {}
}

void DlgRenameFiles::_UpdatePropertiesFile(int nRow, int itype)
{
	try
	{
		CString szval= ff->_xlGIOR(pRgActive, nRow, jcotSTT, L"");
		if (_wtoi(szval) <= 0) return;
		
		RangePtr PRS = pRgActive->GetItem(nRow, icotFilegoc);
		szval = ff->_xlGetHyperLink(PRS);
		if (szval != L"")
		{
			int pos = -1;
			CString sztype = ff->_GetTypeOfFile(szval);
			CString szpathInfo = szval;

			CString szNameOld = ff->_GetNameOfPath(szval, pos, -1);
			CString szNameNew = ff->_xlGIOR(pRgActive, nRow, jcotTen, L"");
			szNameNew = ff->_ConvertRename(szNameNew);
			if (szNameNew != L"" && szNameOld != szNameNew)
			{
				// Thay đổi tên file
				CString szdirNew = szval.Left(pos) + L"\\" + szNameNew + L"." + sztype;
				if ((int)_wrename(szval, szdirNew) != 0)
				{
					PRS = pRgActive->GetItem(nRow, jcotSTT);
					PRS->Interior->PutColor(rgbRed);
					PRS->Font->PutColor(rgbWhite);
				}
				else
				{
					szpathInfo = szdirNew;
					PRS = pRgActive->GetItem(nRow, jcotSTT);
					PRS->Interior->PutColor(RGB(0, 176, 80));
					PRS->Font->PutColor(rgbWhite);

					PRS = pRgActive->GetItem(nRow, icotFilegoc);
					ff->_xlSetHyperlink(pShActive, PRS, szdirNew);
				}
			}

			// Thay đổi thông tin file
			if (itype == 1)
			{
				if (jcotNoidung > 0)
				{
					szval = sztype.Left(3);
					szval.MakeLower();
					if (szval == L"pdf" || szval == L"doc" || szval == L"dot"
						|| szval == L"xls" || szval == L"xlt" || szval == L"ppt" || szval == L"pot")
					{
						CString szInfoOld = ff->_GetProperties(szpathInfo, 1);
						CString szInfoNew = ff->_xlGIOR(pRgActive, nRow, jcotNoidung, L"");

						if (szInfoNew != szInfoOld)
						{
							if (ff->_SetProperties(szpathInfo, 1, szInfoNew))
							{
								PRS = pRgActive->GetItem(nRow, jcotSTT);
								PRS->Interior->PutColor(RGB(0, 140, 246));
								PRS->Font->PutColor(rgbWhite);
							}
						}
					}
				}
			}			
		}
	}
	catch (...) {}
}