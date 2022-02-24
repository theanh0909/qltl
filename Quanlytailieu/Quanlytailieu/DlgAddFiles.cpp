#include "pch.h"
#include "DlgAddFiles.h"

IMPLEMENT_DYNAMIC(DlgAddFiles, CDialogEx)

DlgAddFiles::DlgAddFiles(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgAddFiles::IDD, pParent)
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

	iGFLType = 1, iGFLTen = 2, iGFLSize = 3, iGFLModified = 4, iGFLPath = 5;
	nItem = -1, nSubItem = -1;

	jColumnHand = iGFLPath;

	iChkDefault = 0;
	bChkLookInDir = true;
	szTargetDirectory = L"", szPathCopyMove = L"";
	pathFolderDir = L"";

	iRowPut = 0;
	szpathDir = L"";
	bDeletefile = false;

	nTotalCol = getSizeArray(iwCol);
	iwCol[1] = 0, iwCol[2] = 400, iwCol[3] = 120, iwCol[4] = 150, iwCol[5] = 400;
	iDlgW = 0, iDlgH = 0;
}

DlgAddFiles::~DlgAddFiles()
{
	delete ff;
	delete bb;
	delete reg;

	vecGroups.clear();
	vecstrFileList.clear();

	jColumnHand = -1;
}

void DlgAddFiles::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, ADDDATA_SPLB_OK, m_splbtn_ok);
	DDX_Control(pDX, ADDDATA_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, ADDDATA_BTN_SEARCH, m_btn_search);

	DDX_Control(pDX, ADDDATA_TXT_BROWSE, m_edtTargetDirectory);
	DDX_Control(pDX, ADDDATA_TXT_TYPEFILE, m_edtWildCard);
	DDX_Control(pDX, ADDDATA_TXT_TOTALFILES, m_edtTotalFiles);
	
	DDX_Control(pDX, ADDDATA_RADIO_DIR, m_chkLookInDir);
	DDX_Control(pDX, ADDDATA_RADIO_SUBDIR, m_chkLookInSubdirectories);
	DDX_Control(pDX, ADDDATA_CHECK_TREEFOLDER, m_check_tree_folder);
	DDX_Control(pDX, ADDDATA_CHECK_GROUP, m_check_group);

	DDX_Control(pDX, ADDDATA_TXT_SEARCH, m_editSearchFiles);
	DDX_Control(pDX, ADDDATA_CHECK_ALL, m_check_all);
	DDX_Control(pDX, ADDDATA_LIST_VIEW, m_list_view);
	DDX_Control(pDX, ADDDATA_LBL_TOTALCHECK, m_lbl_total);

	DDX_Control(pDX, ADDDATA_CBB_CELL, m_cbb_cell);
	DDX_Control(pDX, ADDDATA_CBB_PATH, m_cbb_path);
	DDX_Control(pDX, ADDDATA_TXT_PATHMOVE, m_txt_pathMove);
}

BEGIN_MESSAGE_MAP(DlgAddFiles, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(ADDDATA_BTN_CANCEL, &DlgAddFiles::OnBnClickedBtnCancel)
	ON_BN_CLICKED(ADDDATA_BTN_BROWSE, &DlgAddFiles::OnBnClickedBtnBrowse)
	ON_BN_CLICKED(ADDDATA_BTN_SEARCH, &DlgAddFiles::OnBnClickedBtnSearch)
	ON_BN_CLICKED(ADDDATA_RADIO_DIR, &DlgAddFiles::OnBnClickedRadioDir)
	ON_BN_CLICKED(ADDDATA_RADIO_SUBDIR, &DlgAddFiles::OnBnClickedRadioSubdir)
	ON_BN_CLICKED(ADDDATA_CHECK_ALL, &DlgAddFiles::OnBnClickedCheckAll)	
	ON_BN_CLICKED(ADDDATA_SPLB_OK, &DlgAddFiles::OnBnClickedSplbOk)
	ON_BN_CLICKED(ADDDATA_CHECK_GROUP, &DlgAddFiles::OnBnClickedCheckGroup)

	ON_COMMAND(MN_ADDATA_ADDFILES, &DlgAddFiles::OnAddataAddfiles)
	ON_COMMAND(MN_ADDATA_COPYFILES, &DlgAddFiles::OnAddataCopyfiles)
	ON_COMMAND(MN_ADDATA_MOVEFILES, &DlgAddFiles::OnAddataMovefiles)
	ON_COMMAND(MN_ADDATA_OPENFILE, &DlgAddFiles::OnAddataOpenfile)
	ON_COMMAND(MN_ADDATA_OPENFOLDER, &DlgAddFiles::OnAddataOpenfolder)
	ON_COMMAND(MN_ADDATA_DELFILES, &DlgAddFiles::OnAddataDelfiles)

	ON_COMMAND(MN_ADDATA_ADDFILES2, &DlgAddFiles::OnAddataAddfiles2)
	ON_COMMAND(MN_ADDATA_COPYFILES2, &DlgAddFiles::OnAddataCopyfiles2)
	ON_COMMAND(MN_ADDATA_MOVEFILES2, &DlgAddFiles::OnAddataMovefiles2)

	ON_NOTIFY(LVN_ENDSCROLL, ADDDATA_LIST_VIEW, &DlgAddFiles::OnLvnEndScrollListView)
	ON_NOTIFY(NM_CLICK, ADDDATA_LIST_VIEW, &DlgAddFiles::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, ADDDATA_LIST_VIEW, &DlgAddFiles::OnNMRClickListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgAddFiles::OnHdnEndtrackListView)	
	ON_CBN_SELCHANGE(ADDDATA_CBB_PATH, &DlgAddFiles::OnCbnSelchangeCbbPath)
	ON_BN_CLICKED(ADDDATA_BTN_PATHMOVE, &DlgAddFiles::OnBnClickedBtnPathmove)
	ON_NOTIFY(NM_DBLCLK, ADDDATA_LIST_VIEW, &DlgAddFiles::OnNMDblclkListView)
	ON_EN_SETFOCUS(ADDDATA_TXT_PATHMOVE, &DlgAddFiles::OnEnSetfocusTxtPathmove)
	ON_COMMAND(MN_ADDATA_CHECK, &DlgAddFiles::OnAddataCheck)
	ON_COMMAND(MN_ADDATA_UNCHECK, &DlgAddFiles::OnAddataUncheck)
	ON_COMMAND(MN_ADDATA_CHECKALL, &DlgAddFiles::OnAddataCheckall)
	ON_NOTIFY(LVN_KEYDOWN, ADDDATA_LIST_VIEW, &DlgAddFiles::OnLvnKeydownListView)	
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgAddFiles)
	EASYSIZE(ADDDATA_LBL_BROWSE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_TXT_BROWSE, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_BTN_BROWSE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(ADDDATA_RADIO_DIR, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_RADIO_SUBDIR, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_LBL_TOTALCHECK, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_LBL_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_TXT_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(ADDDATA_LINEVER, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(ADDDATA_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_LBL_SEARCH, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_TXT_SEARCH, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_TXT_TYPEFILE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(ADDDATA_BTN_SEARCH, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(ADDDATA_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(ADDDATA_CBB_CELL, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(ADDDATA_CBB_PATH, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(ADDDATA_TXT_PATHMOVE, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(ADDDATA_BTN_PATHMOVE, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)

	EASYSIZE(ADDDATA_CHECK_TREEFOLDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(ADDDATA_CHECK_GROUP, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)

	EASYSIZE(ADDDATA_SPLB_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(ADDDATA_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgAddFiles message handlers
BOOL DlgAddFiles::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_ADDLIST));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_XacdinhSheetLienquan();
	_SetDefault();
	_Timkiemdulieu();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgAddFiles::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(ADDDATA_SPLB_OK))
			{
				OnBnClickedSplbOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_BROWSE))
			{
				_Timkiemdulieu();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_SEARCH))
			{
				OnBnClickedBtnSearch();
				GotoDlgCtrl(GetDlgItem(ADDDATA_TXT_SEARCH));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_TYPEFILE))
			{
				OnBnClickedBtnSearch();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_BROWSE))
			{
				GotoDlgCtrl(GetDlgItem(ADDDATA_TXT_SEARCH));
			}
			else if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(ADDDATA_TXT_TYPEFILE));
			}
			else if (pWndWithFocus == GetDlgItem(ADDDATA_TXT_TYPEFILE))
			{
				GotoDlgCtrl(GetDlgItem(ADDDATA_TXT_BROWSE));
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
		if (pWndWithFocus == GetDlgItem(ADDDATA_LIST_VIEW))
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

void DlgAddFiles::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgAddFiles::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgAddFiles::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgAddFiles::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_ADDDATA));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;

			vector<int> vecRow;

			CString szval = L"";
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				szval.Format(L"Thêm (%d files) vào danh mục", nSelect);
				ModifyMenu(pMenu, MN_ADDATA_ADDFILES, MF_BYCOMMAND | MF_STRING, MN_ADDATA_ADDFILES, szval);
			}
			
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_ADDFILES, IDI_ICON_ADD);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_COPYFILES, IDI_ICON_COPY);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_MOVEFILES, IDI_ICON_SEND);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_DELFILES, IDI_ICON_DELETE);
			
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_CHECK, IDI_ICON_OK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_UNCHECK, IDI_ICON_CANCEL);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_ADDATA_CHECKALL, IDI_ICON_TODOLIST);

			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_ADDATA_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_ADDATA_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(ADDDATA_LIST_VIEW));
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

void DlgAddFiles::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + AddData;
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
	
		szval.Format(L"%d", (int)m_chkLookInDir.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgCheckLookInDir), ff->_ConvertCstringToChars(szval));

		iChkDefault = m_cbb_cell.GetCurSel();
		if (iChkDefault != 1) iChkDefault = 0;
		szval.Format(L"%d", iChkDefault);
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckDefault"), ff->_ConvertCstringToChars(szval));

		m_edtTargetDirectory.GetWindowTextW(szTargetDirectory);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgDirectory), 
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szTargetDirectory, 0)));

		m_txt_pathMove.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgPathCopyMove), 
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szval, 0)));
	
		szval.Format(L"%d", (int)m_check_tree_folder.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckTreeFolder"), ff->_ConvertCstringToChars(szval));
	
		szval.Format(L"%d", (int)m_check_group.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckGroup"), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgAddFiles::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + AddData;
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

		int icheck = 0;
		szval = reg->ReadString(DlgCheckLookInDir, L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
			bChkLookInDir = (bool)icheck;
		}

		szval = reg->ReadString(L"CheckDefault", L"");
		if (szval != L"")
		{
			iChkDefault = _wtoi(szval);
			if (iChkDefault != 1) iChkDefault = 0;
		}

		szval = bb->BaseDecodeEx(reg->ReadString(DlgDirectory, L""));
		if (szval != L"") szTargetDirectory = szval;
		
		szval = bb->BaseDecodeEx(reg->ReadString(DlgPathCopyMove, L""));
		if (szval != L"") szPathCopyMove = szval;

		icheck = 1;
		szval = reg->ReadString(L"CheckTreeFolder", L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
		}
		m_check_tree_folder.SetCheck(icheck);

		icheck = 1;
		szval = reg->ReadString(L"CheckGroup", L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
		}
		m_check_group.SetCheck(icheck);
	}
	catch (...) {}
}

void DlgAddFiles::_SetLocationAndSize()
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

void DlgAddFiles::OnBnClickedBtnCancel()
{
	_SaveRegistry();

	CDialogEx::OnCancel();
}

void DlgAddFiles::_SetDefault()
{
	try
	{
		CString szval = ff->_GetDesktopDir();
		if (szTargetDirectory == L"") szTargetDirectory = szval;		
		m_edtTargetDirectory.SubclassDlgItem(ADDDATA_TXT_BROWSE, this);
		m_edtTargetDirectory.SetBkColor(rgbWhite);
		m_edtTargetDirectory.SetTextColor(rgbBlue);
		m_edtTargetDirectory.SetWindowText(szTargetDirectory);
		m_edtTargetDirectory.SetCueBanner(L"Chọn đường dẫn thư mục chứa files cần thêm...");

		if (szPathCopyMove == L"") szPathCopyMove = pathFolderDir;
		m_txt_pathMove.SubclassDlgItem(ADDDATA_TXT_PATHMOVE, this);
		m_txt_pathMove.SetBkColor(rgbWhite);
		m_txt_pathMove.SetTextColor(rgbBlue);
		m_txt_pathMove.SetWindowText(szPathCopyMove);
		m_txt_pathMove.SetCueBanner(L"Chọn đường dẫn lưu trữ files được sao chép hoặc di chuyển...");

		m_edtWildCard.SubclassDlgItem(ADDDATA_TXT_TYPEFILE, this);
		m_edtWildCard.SetBkColor(rgbWhite);
		m_edtWildCard.SetTextColor(rgbBlue);
		m_edtWildCard.SetCueBanner(L"Một hoặc nhiều đuôi files. Ví dụ: xlsx doc,...");

		m_edtTotalFiles.SubclassDlgItem(ADDDATA_TXT_TOTALFILES, this);
		m_edtTotalFiles.SetBkColor(rgbWhite);
		m_edtTotalFiles.SetTextColor(rgbRed);

		m_editSearchFiles.SetCueBanner(L"Nhập từ khóa tìm kiếm...");
		m_editSearchFiles.SubclassDlgItem(ADDDATA_TXT_SEARCH, this);
		m_editSearchFiles.SetBkColor(rgbWhite);
		m_editSearchFiles.SetTextColor(rgbBlue);

		m_lbl_total.SubclassDlgItem(ADDDATA_LBL_TOTALCHECK, this);
		m_lbl_total.SetTextColor(rgbBlue);

		m_chkLookInDir.SetCheck(bChkLookInDir);
		m_chkLookInSubdirectories.SetCheck(!bChkLookInDir);

		m_cbb_cell.AddString(L"Đổ dữ liệu tại dòng đầu tiên");
		m_cbb_cell.AddString(L"Đổ dữ liệu tại dòng được chỉ định");
		m_cbb_cell.SetCurSel(iChkDefault);

		m_check_all.SetCheck(1);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, 40);
		m_list_view.InsertColumn(iGFLType, L" Trạng thái", LVCFMT_CENTER, 0 /*iwCol[1]*/);
		m_list_view.InsertColumn(iGFLTen, L" Tên file", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLSize, L" Kích cỡ", LVCFMT_RIGHT, iwCol[3]);
		m_list_view.InsertColumn(iGFLModified, L" Ngày sửa đổi", LVCFMT_CENTER, iwCol[4]);
		m_list_view.InsertColumn(iGFLPath, L" Đường dẫn", LVCFMT_LEFT, iwCol[5]);

		m_list_view.SetColumnColors(iGFLPath, rgbWhite, rgbBlue);
		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);

/**********************************************************************************/

		m_btn_search.SetIcon(IDI_ICON_SEARCH, 24, 24);
		m_btn_search.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_search.DrawTransparent(true);
		m_btn_search.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_splbtn_ok.SetDropDownMenu(IDR_MENU_ADDDATA, 1);

/**********************************************************************************/

		int icolGet = 0;
		bool blCate = false;
		CString szCopy = shCopy;
		if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
		{
			blCate = true;
			icolGet = jcotLV;
		}
		else
		{
			if (jcotFolder > 0) icolGet = jcotFolder;
		}

		vector<CString> vecPath;
		int ncountPath = _XacdinhDanhsachThumuc(vecPath, icolGet);
		if (ncountPath > 0)
		{
			if(blCate) m_cbb_path.SetCueBanner(L"Kích chọn thư mục sử dụng để lưu trữ...");
			else m_cbb_path.SetCueBanner(L"Kích chọn đường dẫn sử dụng để lưu trữ...");

			for (int i = 0; i < ncountPath; i++)
			{
				m_cbb_path.AddString(vecPath[i]);
			}

			/*CString szpath = L"";
			CString szNameDir = ff->_xlGIOR(pRgActive, irowStart, icolGet, L"");
			if (szNameDir != L"")
			{
				for (int i = 0; i < ncountPath; i++)
				{
					m_cbb_path.GetLBText(i, szval);
					if (szval == szNameDir)
					{
						m_cbb_path.SetCurSel(i);

						if (blCate) szpath = pathFolderDir + szNameDir;
						else szpath = szNameDir;

						if (szpath.Right(1) != L"\\") szpath += L"\\";
						m_txt_pathMove.SetWindowText(szpath);

						break;
					}
				}
			}*/
		}
		else m_cbb_path.EnableWindow(0);

		vecPath.clear();
	}
	catch (...) {}
}

void DlgAddFiles::OnBnClickedBtnBrowse()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần tìm kiếm...";
	if (szTargetDirectory != L"") m_dlg.m_ofn.lpstrInitialDir = szTargetDirectory;
	if (m_dlg.DoModal() == IDOK)
	{
		szTargetDirectory = m_dlg.GetPathName();
		m_edtTargetDirectory.SetWindowText(szTargetDirectory);
		_GetAllFileInFolder(szTargetDirectory);
	}
}

void DlgAddFiles::_GetAllFileInFolder(CString szDir)
{
	try
	{
		bDeletefile = false;
		if (szDir.CompareNoCase(L"") == 0) return;
		if (szDir.GetLength() < 5) return;

		tongkq = 0;
		vecstrFileList.clear();

		if ((int)m_chkLookInSubdirectories.GetCheck() == 1)
		{
			vector<CString> vecFolder;
			CUtil::GetFolderList(_tstring(szDir.GetBuffer()), vecFolder);

			int ncount = (int)vecFolder.size();
			for (int i = 0; i < ncount; i++)
			{
				CUtil::GetFileList(_tstring(vecFolder[i].GetBuffer()), L"*.*", false, vecstrFileList);
			}

			vecFolder.clear();
		}
		else
		{
			CUtil::GetFileList(_tstring(szDir.GetBuffer()), L"*.*", false, vecstrFileList);
		}

		tongkq = (long)vecstrFileList.size();

		_LocdulieuNangcao();
	}
	catch (...) {}
}

void DlgAddFiles::_FilterTypeOfFile(vector<MyStrFilter> &vecFt)
{
	try
	{
		if (tongkq == 0) return;

		CString szval = L"";
		CString strWildCard = L"";
		m_edtWildCard.GetWindowText(strWildCard);
		strWildCard.Replace(L"*", L"");
		strWildCard.Replace(L".", L"");
		strWildCard.Replace(L" ", L"");
		strWildCard.MakeLower();
		strWildCard.Trim();

		MyStrFilter MSFT;

		if (strWildCard == L"")
		{
			for (long i = 0; i < tongkq; i++)
			{
				MSFT.szItem = vecstrFileList[i];
				MSFT.iEnable = 1;

				vecFt.push_back(MSFT);
			}
		}
		else
		{
			CString szType = L"";
			for (long i = 0; i < tongkq; i++)
			{
				szval = vecstrFileList[i];
				szType = ff->_GetTypeOfFile(szval).MakeLower();
				if (szType != L"")
				{
					if (strWildCard.Find(szType) >= 0)
					{
						MSFT.szItem.Format(L"%s", szval);
						MSFT.iEnable = 1;

						vecFt.push_back(MSFT);
					}
				}
			}
		}
	}
	catch (...) {}
}

void DlgAddFiles::_FilterSearchText(vector<MyStrFilter> &vecFt)
{
	try
	{
		tongfilter = (int)vecFilter.size();
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
				if (vecFt[i].szItem.Find(L"~$") >= 0 
					|| vecFt[i].szItem.Right(4) == L".lnk")
				{
					vecFt.erase(vecFt.begin() + i);
				}
				else
				{
					szval = ff->_ConvertKhongDau(vecFt[i].szItem);
					szval.MakeLower();

					icheck = 0;
					for (int j = 0; j < iKey; j++)
					{
						if (szval.Find(vecKey[j]) == -1) break;
						icheck++;
					}

					if (icheck < iKey) vecFt.erase(vecFt.begin() + i);
				}
			}

			vecKey.clear();
		}
		else
		{
			for (long i = tongfilter - 1; i >= 0; i--)
			{
				if (vecFt[i].szItem.Find(L"~$") >= 0
					|| vecFt[i].szItem.Right(4) == L".lnk")
				{
					vecFt.erase(vecFt.begin() + i);
				}
			}
		}		
	}
	catch (...) {}
}

void DlgAddFiles::_Timkiemdulieu()
{
	m_edtTargetDirectory.GetWindowText(szTargetDirectory);
	_GetAllFileInFolder(szTargetDirectory);
}

void DlgAddFiles::_LocdulieuNangcao()
{
	try
	{
		vecFilter.clear();

		// Lọc dữ liệu theo đuôi files tìm kiếm
		_FilterTypeOfFile(vecFilter);

		// Lọc dữ liệu theo từ khóa tìm kiếm
		_FilterSearchText(vecFilter);

/****** Load dữ liệu *********************************************************/

		tongfilter = (int)vecFilter.size();

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

void DlgAddFiles::OnBnClickedBtnSearch()
{
	if (tongkq == 0 || bDeletefile == true) _Timkiemdulieu();
	else _LocdulieuNangcao();
}

void DlgAddFiles::OnBnClickedRadioDir()
{
	_Timkiemdulieu();
}

void DlgAddFiles::OnBnClickedRadioSubdir()
{
	_Timkiemdulieu();
}

void DlgAddFiles::OnBnClickedCheckAll()
{
	int icheck = m_check_all.GetCheck();
	int ncount = m_list_view.GetItemCount();
	for (int i = 0; i < ncount; i++) m_list_view.SetCheck(i, icheck);

	if (tongfilter > ncount)
	{
		for (int i = ncount; i < tongfilter; i++) vecFilter[i].iEnable = icheck;
	}

	if(icheck == 1) _SetStatusKetqua(tongfilter);
	else _SetStatusKetqua(0);
}

void DlgAddFiles::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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

void DlgAddFiles::_AddItemsListKetqua(int iBegin, int iEnd, int icheck)
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

			szval = ff->_GetNameOfPath(vecFilter[i].szItem, pos, 0);
			m_list_view.SetItemText(i, iGFLTen, szval);

			// Thông tin files
			ff->_GetFileAttributesEx(vecFilter[i].szItem, szSize, szNgay);

			m_list_view.SetItemText(i, iGFLSize, szSize);
			m_list_view.SetItemText(i, iGFLModified, szNgay);

			m_list_view.SetItemText(i, iGFLPath, vecFilter[i].szItem);

			szval = ff->_GetTypeOfFile(vecFilter[i].szItem);

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

void DlgAddFiles::_AutoSothutu(_WorksheetPtr pSheet, int numStart, int nRowStart, int nRowEnd)
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

int DlgAddFiles::_GetAllSelectedItems(vector<int> &vecRow)
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

int DlgAddFiles::_GetAllCheckedItems(vector<int> &vecRow)
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

void DlgAddFiles::_AddFilesCategory(int iOnClickBtn, int iMoveCopy)
{
	try
	{
		vector<int> vecRow;

		int nSelect = 0;
		if (iOnClickBtn == 0) nSelect = _GetAllSelectedItems(vecRow);
		else nSelect = _GetAllCheckedItems(vecRow);

		if (nSelect > 0)
		{
			int pos = -1;
			CString szval = L"";

			m_edtTargetDirectory.GetWindowTextW(szTargetDirectory);
			int iLenTarget = szTargetDirectory.GetLength();
			int iLookInSub = (int)m_chkLookInSubdirectories.GetCheck();
			int iChkTreeFolder = (int)m_check_tree_folder.GetCheck();

			CString pFolder = L"";
			m_txt_pathMove.GetWindowTextW(pFolder);
			
			pFolder.Trim();
			if (pFolder == L"") pFolder = pathFolderDir;
			if (pFolder.Right(1) != L"\\") pFolder += L"\\";
			CString szNameDir = ff->_GetNameOfPath(pFolder, pos, 0);

			if (iMoveCopy > 0)
			{
				if (iChkTreeFolder == 0)
				{
					if (pFolder == pathFolderDir)
					{
						szNameDir = L"__CopyFiles";
						pFolder += (szNameDir + L"\\");
					}
				}
			}
			else szNameDir = ff->_GetNameOfPath(szTargetDirectory, pos, 0);

			if (iMoveCopy > 0) ff->_CreateDirectories(pFolder);

			// Kiểm tra vị trí đổ dữ liệu
			iChkDefault = (int)m_cbb_cell.GetCurSel();
			if (iChkDefault == 0) iRowPut = irowStart;
			else
			{
				iRowPut = iRowActive;
				if (iRowPut < irowStart) iRowPut = irowStart;
				else
				{
					// Xác định đã tồn tại dữ liệu chưa?
					szval = ff->_xlGIOR(pRgActive, irowStart, jcotLV, L"");
					if (szval == L"") szval = ff->_xlGIOR(pRgActive, irowStart + 1 , jcotLV, L"");
					if (szval == L"") iRowPut = irowStart;
				}
			}

			// Kiểm tra trạng thái check group			
			int iCheckGroup = (int)m_check_group.GetCheck();
			if (iCheckGroup == 1)
			{
				// Đối với trường hợp checkgroup mà vị trí đổ tại vị trí active
				// Thì cần phải xét xem vị trí đổ có kẹp giữa trong group nào không?
				// Nếu nằm trong group nào đó thì bỏ qua không tạo group nữa.
				if (iChkDefault == 1)
				{
					if (iRowPut > irowStart)
					{
						int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
						for (int i = iRowPut + 1; i <= iRowLast; i++)
						{
							szval = ff->_xlGIOR(pRgActive, i, jcotLV);
							if (szval != L"")
							{
								iCheckGroup = 0;
								break;
							}
						}
					}					
				}
			}

			if (iOnClickBtn == 1)
			{
				_SaveRegistry();
				CDialogEx::OnOK();
			}
			
			ff->_xlPutScreenUpdating(false);

			RangePtr PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, 1), pRgActive->GetItem(iRowPut + nSelect - 1, 1));
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, 1), pRgActive->GetItem(iRowPut + nSelect - 1, 1));
			PRS->EntireRow->Font->PutName(FontTimes);
			PRS->EntireRow->Font->PutSize(13);
			PRS->EntireRow->Font->PutBold(false);
			PRS->EntireRow->Font->PutItalic(false);
			PRS->EntireRow->Font->PutColor(rgbBlack);
			PRS->EntireRow->Interior->PutColor(xlNone);

			MyStrGroups MSFILE;
			MSFILE.szLinhvuc = L"";

			vecGroups.clear();

			int total = 0, iCheckChildFoler = 1;
			CString szChilDir = L"", szNameMove = szNameDir;
			CString szdir = L"", szname = L"", szndung = L"", sztype = L"", szcheck = L"";
			CString szUpdateStatus = L"Đang thêm mới dữ liệu. Vui lòng đợi trong giây lát...";
			for (int i = 0; i < nSelect; i++)
			{
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / nSelect), L"%");
				ff->_xlPutStatus(szval);

				szChilDir = vecFilter[vecRow[i]].szItem;
				szChilDir = szChilDir.Right(szChilDir.GetLength() - iLenTarget);
				if (szChilDir.Left(1) == L"\\") szChilDir = szChilDir.Right(szChilDir.GetLength() - 1);

				szname = ff->_GetNameOfPath(vecFilter[vecRow[i]].szItem, pos, -1);
				pRgActive->PutItem(iRowPut + i, jcotTen, (_bstr_t)szname);

				sztype = ff->_GetTypeOfFile(vecFilter[vecRow[i]].szItem);
				if (jcotType > 0) pRgActive->PutItem(iRowPut + i, jcotType, (_bstr_t)sztype);
				if (jcotNoidung > 0)
				{
					szndung = L"";
					szval = sztype.Left(3);
					szval.MakeLower();
					if (szval == L"pdf" || szval == L"doc" || szval == L"dot" 
						|| szval == L"xls" || szval == L"xlt" || szval == L"ppt" || szval == L"pot")
					{
						szndung = ff->_GetProperties(vecFilter[vecRow[i]].szItem, 1);
					}

					if (szndung == L"")
					{
						szndung = szname;
						szndung.Replace(L"-", L" ");
						szndung.Replace(L"_", L" ");
					}

					pRgActive->PutItem(iRowPut + i, jcotNoidung, (_bstr_t)szndung);
				}

				if (jcotFolder > 0 && pos > 0)
				{
					if (iMoveCopy > 0)
					{
						if (iChkTreeFolder == 1)
						{
							szdir = ff->_GetNameOfPath(pFolder + szChilDir, pos, 1);
						}
						else szdir = pFolder + szname + L"." + sztype;
					}
					else szdir = vecFilter[vecRow[i]].szItem.Left(pos);
					
					pRgActive->PutItem(iRowPut + i, jcotFolder, (_bstr_t)szdir);
				}

				szNameDir = szNameMove;
				if (iMoveCopy == 0 || iChkTreeFolder == 1)
				{
					pos = szChilDir.Find(L"\\");
					if (pos > 0) szNameDir = szChilDir.Left(pos);
				}

				if (iCheckChildFoler == 1)
				{
					if (jcotFolder > 0 && pos > 0)
					{
						szNameDir = ff->_GetNameOfPath(szdir, pos, 0);
					}
				}

				pRgActive->PutItem(iRowPut + i, jcotLV, (_bstr_t)szNameDir);

				PRS = pRgActive->GetItem(iRowPut + i, icotFilegoc);
				if (iMoveCopy > 0)
				{
					if (iChkTreeFolder == 1)
					{
						szdir = pFolder + szChilDir;						
						ff->_CreateDirectories(ff->_GetNameOfPath(szdir, pos, 1));

						if (iMoveCopy == 2) MoveFile(vecFilter[vecRow[i]].szItem, szdir);
						else CopyFile(vecFilter[vecRow[i]].szItem, szdir, FALSE);
					}
					else
					{
						szdir = pFolder + szname + L"." + sztype;
						if (iMoveCopy == 2) MoveFile(vecFilter[vecRow[i]].szItem, szdir);
						else CopyFile(vecFilter[vecRow[i]].szItem, szdir, FALSE);
					}

					ff->_xlSetHyperlink(pShActive, PRS, szdir);
				}
				else
				{
					ff->_xlSetHyperlink(pShActive, PRS, vecFilter[vecRow[i]].szItem);
				}

				if (iCheckGroup == 1)
				{
					if (total == 0 || szNameDir != MSFILE.szLinhvuc)
					{
						// Thêm mới
						MSFILE.szLinhvuc = szNameDir;
						MSFILE.szPathFolder = L"";
						if (jcotFolder > 0)
						{
							MSFILE.szPathFolder = ff->_xlGIOR(pRgActive, iRowPut + i, jcotFolder, L"");
							if (MSFILE.szPathFolder == L"")
							{
								MSFILE.szPathFolder = ff->_xlGIOR(pRgActive, iRowPut + i + 1, jcotFolder, L"");
							}
						}

						MSFILE.iRBegin = iRowPut + i;
						MSFILE.iREnd = iRowPut + i;
						MSFILE.iEnable = 1;
						vecGroups.push_back(MSFILE);

						total++;
					}
					else
					{
						// Cộng dồn giá trị
						vecGroups[total - 1].iREnd = iRowPut + i;
					}
				}
			}

			_AutoSothutu(pShActive, 1, 0, 0);

			PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, jcotTen),
				pRgActive->GetItem(iRowPut + nSelect - 1, jcotTen));
			PRS->PutHorizontalAlignment(xlLeft);

			if (jcotNoidung > 0)
			{
				PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, jcotNoidung),
					pRgActive->GetItem(iRowPut + nSelect - 1, jcotNoidung));
				PRS->PutHorizontalAlignment(xlLeft);
			}
			
			if (jcotFolder > 0)
			{
				PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, jcotFolder),
					pRgActive->GetItem(iRowPut + nSelect - 1, jcotFolder));
				PRS->PutHorizontalAlignment(xlLeft);
			}

			PRS = pShActive->GetRange(pRgActive->GetItem(iRowPut, jcotSTT),
				pRgActive->GetItem(iRowPut + nSelect - 1, icotEnd));
			ff->_xlFormatBorders(PRS, 3, true, true);
			PRS->Rows->AutoFit();

/********** Bổ sung group dữ liệu ******************************************************************/

			int nCountGrp = (int)vecGroups.size();
			if (nCountGrp > 0)
			{
				pShActive->Outline->PutAutomaticStyles(false);
				pShActive->Outline->PutSummaryRow(xlSummaryAbove);
				pShActive->Outline->PutSummaryColumn(xlSummaryOnLeft);

				int dem = 0;
				szUpdateStatus = L"Đang tiến hành tạo nhóm dữ liệu. Vui lòng đợi trong giây lát...";
				for (int i = nCountGrp - 1; i >= 0; i--)
				{
					dem++;
					szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * dem / nCountGrp), L"%");
					ff->_xlPutStatus(szval);

					_CreateGroupItems(i);	// <-- Tạo group dữ liệu
				}

				vecGroups.clear();
				//pShActive->Outline->ShowLevels(1, 1);
			}

			if (iOnClickBtn == 1) ff->_msgbox(L"Thêm mới dữ liệu thành công!", 0, 1, 2000);
			ff->_xlPutScreenUpdating(true);
			ff->_xlPutStatus(L"Ready");

			PRS = pRgActive->GetItem(iRowPut, 1);
			PRS->Select();
		}

		vecRow.clear();
	}
	catch (...) {}
}

void DlgAddFiles::_CreateGroupItems(int nItem)
{
	try
	{
		int Ibg = vecGroups[nItem].iRBegin;
		int Iend = vecGroups[nItem].iREnd;

		RangePtr PRS = pShActive->GetRange(pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Iend, 1));
		/*if ((int)PRS->EntireRow->OutlineLevel >= 2) */PRS->EntireRow->Rows->ClearOutline();

		// Chèn thêm 1 dòng trước nhóm
		CString szID = ff->_xlGIOR(pRgActive, Ibg - 1, jcotSTT, L"");
		if (_wtoi(szID) > 0 || Ibg == irowStart)
		{
			PRS = pRgActive->GetItem(Ibg + 1, 1);
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = pRgActive->GetItem(Ibg, 1);
			PRS->EntireRow->Copy(pRgActive->GetItem(Ibg + 1, 1));
			xl->PutCutCopyMode(XlCutCopyMode(false));

			PRS = pShActive->GetRange(
				pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Ibg, icotEnd));
			PRS->EntireRow->ClearContents();

			_SetStyleGroup(PRS);

			pRgActive->PutItem(Ibg, jcotLV, (_bstr_t)vecGroups[nItem].szLinhvuc);
			if (jcotFolder > 0) pRgActive->PutItem(Ibg, jcotFolder, (_bstr_t)vecGroups[nItem].szPathFolder);

			Ibg++; Iend++;
		}
		else
		{
			szID = ff->_xlGIOR(pRgActive, Ibg - 1, jcotLV, L"");
			if (szID == L"")
			{
				pRgActive->PutItem(Ibg - 1, jcotLV, (_bstr_t)vecGroups[nItem].szLinhvuc);
				if (jcotFolder > 0) pRgActive->PutItem(Ibg - 1, jcotFolder, (_bstr_t)vecGroups[nItem].szPathFolder);
			}

			PRS = pShActive->GetRange(
				pRgActive->GetItem(Ibg - 1, 1), pRgActive->GetItem(Ibg - 1, icotEnd));
			_SetStyleGroup(PRS);
		}

		PRS = pShActive->GetRange(pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Iend, 1));
		PRS->EntireRow->Rows->Group();
	}
	catch (...) {}
}

void DlgAddFiles::_SetStyleGroup(RangePtr PRS)
{
	PRS->Font->PutBold(true);
	PRS->Font->PutItalic(false);
	PRS->Font->PutColor(rgbBlack);
	PRS->Interior->PutColor(RGB(153, 255, 204));
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
}


void DlgAddFiles::_SelectAllItems()
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

void DlgAddFiles::OnAddataAddfiles()
{
	_AddFilesCategory(0, 0);
}

void DlgAddFiles::OnAddataCopyfiles()
{
	_AddFilesCategory(0 , 1);
}


void DlgAddFiles::OnAddataMovefiles()
{
	_AddFilesCategory(0, 2);
}


void DlgAddFiles::OnAddataOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}


void DlgAddFiles::OnAddataOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}

void DlgAddFiles::OnAddataDelfiles()
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
				m_edtTotalFiles.GetWindowTextW(szval);
				szval.Replace(L".", L"");
				long itotal = _wtol(szval);
				int dem = 0;

				for (int i = nSelect - 1; i >= 0; i--)
				{
					CString szpath = m_list_view.GetItemText(vecRow[i], iGFLPath);
					if (szpath != L"")
					{
						DeleteFile(szpath);
						m_list_view.DeleteItem(vecRow[i]);
						vecFilter.erase(vecFilter.begin() + vecRow[i]);

						dem++;
					}
				}

				bDeletefile = true;

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

void DlgAddFiles::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
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

void DlgAddFiles::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}

void DlgAddFiles::_XacdinhSheetLienquan()
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

int DlgAddFiles::_XacdinhDanhsachThumuc(vector<CString> &vecPath, int icolGet)
{
	try
	{
		vecPath.clear();

		CString szval = L"";
		int ncount = 0, icheck = 0;
		for (int i = irowStart; i <= irowEnd; i++)
		{
			szval = ff->_xlGIOR(pRgActive, i, icolGet, L"");
			if (szval != L"")
			{
				icheck = 0;
				ncount = (int)vecPath.size();
				if (ncount > 0)
				{
					for (int j = 0; j < ncount; j++)
					{
						if (szval == vecPath[j])
						{
							icheck = 1;
							break;
						}
					}
				}

				if (icheck == 0) vecPath.push_back(szval);
			}
		}

		return (int)vecPath.size();
	}
	catch (...) {}
	return 0;
}

void DlgAddFiles::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 40;
	else if (pSubItem == iGFLType) pNMListHeader->pitem->cxy = 0;
	*pResult = 0;
}

void DlgAddFiles::OnBnClickedSplbOk()
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
			CString szval = L"";
			m_splbtn_ok.GetWindowTextW(szval);
			szval = szval.Left(1);

			int iTypeMoveCopy = 0;							// Thêm
			if (szval == L"S") iTypeMoveCopy = 1;			// Sao chép
			else if (szval == L"D") iTypeMoveCopy = 2;		// Di chuyển

			_AddFilesCategory(1, iTypeMoveCopy);
		}
		else ff->_msgbox(L"Bạn chưa tích chọn dữ liệu cần thêm.", 2, 0);
	}
	catch (...) {}
}


void DlgAddFiles::OnAddataAddfiles2()
{
	m_splbtn_ok.SetWindowText(L"Thêm vào danh mục");
}


void DlgAddFiles::OnAddataCopyfiles2()
{
	m_splbtn_ok.SetWindowText(L"Sao chép & thêm vào danh mục");
}

void DlgAddFiles::OnAddataMovefiles2()
{
	m_splbtn_ok.SetWindowText(L"Di chuyển & thêm vào danh mục");
}

void DlgAddFiles::OnCbnSelchangeCbbPath()
{
	CString szNameDir = L"";
	int ncountPath = m_cbb_path.GetCount();
	if (ncountPath > 0)
	{
		int nIndex = m_cbb_path.GetCurSel();
		if (nIndex >= 0)
		{
			m_cbb_path.GetLBText(nIndex, szNameDir);

			CString szpath = L"";
			CString szCopy = shCopy;
			if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
			{
				szpath = pathFolderDir + szNameDir + L"\\";
			}
			else szpath = szNameDir;

			if (szpath.Right(1) != L"\\") szpath += L"\\";
			m_txt_pathMove.SetWindowText(szpath);
		}
	}
}

void DlgAddFiles::OnBnClickedBtnPathmove()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục sử dụng để lưu trữ...";
	if (szPathCopyMove != L"") m_dlg.m_ofn.lpstrInitialDir = szPathCopyMove;
	if (m_dlg.DoModal() == IDOK)
	{
		szPathCopyMove = m_dlg.GetPathName();
		m_txt_pathMove.SetWindowText(szPathCopyMove);
	}
}


void DlgAddFiles::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szval = m_list_view.GetItemText(nItem, iGFLPath);
		if (szval != L"") ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void DlgAddFiles::OnEnSetfocusTxtPathmove()
{
	m_txt_pathMove.ShowBalloonTip(L"Hướng dẫn", 
		L"Đường dẫn chứa thư mục sẽ lưu files được sao chép hoặc di chuyển. "
		L"\nGiá trị này chỉ có tác dụng khi chạy lệnh sao chép hoặc di chuyển files.", TTI_INFO);
}

int DlgAddFiles::_GetCountItemsCheck()
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

void DlgAddFiles::_GetItemSelect(int icheck)
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

void DlgAddFiles::_SetStatusKetqua(int num)
{
	CString str = L"";
	if (num > 1) str.Format(L"Bạn đã chọn (%d) files / ", num);
	else str.Format(L"Bạn đã chọn (%d) file / ", num);
	m_lbl_total.SetWindowText(str);
}

void DlgAddFiles::OnAddataCheck()
{
	_GetItemSelect(1);
}


void DlgAddFiles::OnAddataUncheck()
{
	_GetItemSelect(0);
}


void DlgAddFiles::OnAddataCheckall()
{
	_SelectAllItems();
}

void DlgAddFiles::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_SPACE) OnAddataAddfiles();
	if (pLVKeyDow->wVKey == VK_DELETE) OnAddataDelfiles();
	*pResult = 0;
}


void DlgAddFiles::OnBnClickedCheckGroup()
{
	if ((int)m_check_group.GetCheck() == 0)
	{
		ff->_msgbox(L"Khi bỏ tích 'Tạo nhóm (groups)' "
			L"bạn sẽ không sử dụng được 1 số tính năng như sắp xếp dữ liệu, "
			L"thay đổi tên thư mục,... Tuy nhiên, bạn cũng có thể tạo nhóm sau "
			L"bằng cách sử dụng lệnh 'Tiện ích / Tạo nhóm dữ liệu'.", 1, 0);
	}
}
