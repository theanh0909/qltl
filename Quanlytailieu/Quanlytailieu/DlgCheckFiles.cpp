#include "pch.h"
#include "DlgCheckFiles.h"

IMPLEMENT_DYNAMIC(DlgCheckFiles, CDialogEx)

DlgCheckFiles::DlgCheckFiles(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgCheckFiles::IDD, pParent)
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

	jColumnHand = iGFLPath;

	nItem = -1, nSubItem = -1;
	slnew = 0;

	szTargetDirectory = L"";
	pathFolderDir = L"";
	iLenPathDir = 0;

	bChkLookInDir = true;
	bChkLink = false, bChkNew = true;
	bDeletefile = false;
	bSetGroup = false;

	nTotalCol = getSizeArray(iwCol);
	iwCol[1] = 100, iwCol[2] = 400, iwCol[3] = 120, iwCol[4] = 150, iwCol[5] = 400;
	iDlgW = 0, iDlgH = 0;
}

DlgCheckFiles::~DlgCheckFiles()
{
	delete ff;
	delete bb;
	delete reg;

	vecHyper.clear();
	vecstrFileList.clear();

	jColumnHand = -1;
}

void DlgCheckFiles::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CHKFS_BTN_OK, m_btn_ok);
	DDX_Control(pDX, CHKFS_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, CHKFS_CBB_PATH, m_cbb_path);
	DDX_Control(pDX, CHKFS_CHECK_ALL, m_check_all);
	DDX_Control(pDX, CHKFS_CHECK_NEWS, m_check_new);
	DDX_Control(pDX, CHKFS_CHECK_LINK, m_check_link);
	DDX_Control(pDX, CHKFS_LIST_VIEW, m_list_view);
	DDX_Control(pDX, CHKFS_LBL_TOTALCHECK, m_lbl_total);
	DDX_Control(pDX, CHKFS_TXT_TOTALFILES, m_edtTotalFiles);
	DDX_Control(pDX, CHKFS_TXT_BROWSE, m_edtTargetDirectory);
	DDX_Control(pDX, CHKFS_RADIO_DIR, m_chkLookInDir);
	DDX_Control(pDX, CHKFS_RADIO_SUBDIR, m_chkLookInSubdirectories);
}

BEGIN_MESSAGE_MAP(DlgCheckFiles, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(CHKFS_BTN_OK, &DlgCheckFiles::OnBnClickedBtnOk)
	ON_BN_CLICKED(CHKFS_BTN_CANCEL, &DlgCheckFiles::OnBnClickedBtnCancel)
	ON_BN_CLICKED(CHKFS_BTN_BROWSE, &DlgCheckFiles::OnBnClickedBtnBrowse)

	ON_BN_CLICKED(CHKFS_CHECK_ALL, &DlgCheckFiles::OnBnClickedCheckAll)
	ON_BN_CLICKED(CHKFS_CHECK_NEWS, &DlgCheckFiles::OnBnClickedCheckNews)
	ON_CBN_SELCHANGE(CHKFS_CBB_PATH, &DlgCheckFiles::OnCbnSelchangeCbbPath)
	ON_BN_CLICKED(CHKFS_CHECK_LINK, &DlgCheckFiles::OnBnClickedCheckLink)

	ON_COMMAND(MN_CHKFL_ADDFILES, &DlgCheckFiles::OnChkflAddfiles)
	ON_COMMAND(MN_CHKFL_GOTOSHEET, &DlgCheckFiles::OnChkflGotosheet)
	ON_COMMAND(MN_CHKFL_OPENFILE, &DlgCheckFiles::OnChkflOpenfile)
	ON_COMMAND(MN_CHKFL_OPENFOLDER, &DlgCheckFiles::OnChkflOpenfolder)
	ON_COMMAND(MN_CHKFL_DELFILES, &DlgCheckFiles::OnChkflDelfiles)	
	ON_COMMAND(MN_CHKFL_CHECK, &DlgCheckFiles::OnChkflCheck)
	ON_COMMAND(MN_CHKFL_UNCHECK, &DlgCheckFiles::OnChkflUncheck)
	ON_COMMAND(MN_CHKFL_CHECKALL, &DlgCheckFiles::OnChkflCheckall)

	ON_NOTIFY(NM_CLICK, CHKFS_LIST_VIEW, &DlgCheckFiles::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, CHKFS_LIST_VIEW, &DlgCheckFiles::OnNMRClickListView)
	ON_NOTIFY(NM_DBLCLK, CHKFS_LIST_VIEW, &DlgCheckFiles::OnNMDblclkListView)

	ON_NOTIFY(LVN_ENDSCROLL, CHKFS_LIST_VIEW, &DlgCheckFiles::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgCheckFiles::OnHdnEndtrackListView)
	
	ON_NOTIFY(LVN_KEYDOWN, CHKFS_LIST_VIEW, &DlgCheckFiles::OnLvnKeydownListView)
	ON_BN_CLICKED(CHKFS_RADIO_DIR, &DlgCheckFiles::OnBnClickedRadioDir)
	ON_BN_CLICKED(CHKFS_RADIO_SUBDIR, &DlgCheckFiles::OnBnClickedRadioSubdir)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgCheckFiles)
	EASYSIZE(CHKFS_CBB_PATH, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	
	EASYSIZE(CHKFS_TXT_BROWSE, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHKFS_BTN_BROWSE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(CHKFS_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CHKFS_CHECK_NEWS, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	
	EASYSIZE(CHKFS_LBL_TOTALCHECK, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHKFS_LBL_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHKFS_TXT_TOTALFILES, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHKFS_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	
	EASYSIZE(CHKFS_RADIO_DIR, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(CHKFS_RADIO_SUBDIR, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	
	EASYSIZE(CHKFS_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(CHKFS_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgCheckFiles message handlers
BOOL DlgCheckFiles::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_TODOLIST));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";
	iLenPathDir = pathFolderDir.GetLength();

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_SetDefault();
	_DanhsachHyperlink();
	_Timkiemdulieu();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgCheckFiles::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(CHKFS_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(CHKFS_TXT_BROWSE))
			{
				_Timkiemdulieu();
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
		if (pWndWithFocus == GetDlgItem(CHKFS_LIST_VIEW))
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

void DlgCheckFiles::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgCheckFiles::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}


void DlgCheckFiles::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgCheckFiles::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_CHECKFILES));
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
				ModifyMenu(pMenu, MN_CHKFL_ADDFILES, MF_BYCOMMAND | MF_STRING, MN_CHKFL_ADDFILES, szval);
			}

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_ADDFILES, IDI_ICON_ADD);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_GOTOSHEET, IDI_ICON_MOVE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_DELFILES, IDI_ICON_DELETE);

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_CHECK, IDI_ICON_OK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_UNCHECK, IDI_ICON_CANCEL);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_CHKFL_CHECKALL, IDI_ICON_TODOLIST);

			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_CHKFL_GOTOSHEET, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_CHKFL_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_CHKFL_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			szval = m_list_view.GetItemText(nItem, iGFLType);
			if (szval != DF_GANLINK)
			{
				EnableMenuItem(pMenu, MN_CHKFL_GOTOSHEET, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(CHKFS_LIST_VIEW));
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

void DlgCheckFiles::OnBnClickedBtnOk()
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

		if (ncheck > 0) _AddFilesCategory(1);
		else ff->_msgbox(L"Bạn chưa tích chọn dữ liệu cần thêm.", 2, 0);
	}
	catch (...) {}
}


void DlgCheckFiles::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}


void DlgCheckFiles::OnBnClickedBtnBrowse()
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

void DlgCheckFiles::_GetDanhsachFiles(int iSelectCell)
{
	try
	{
		if (iSelectCell >= 0)
		{
			_XacdinhSheetLienquan();

			bSetGroup = false;
			if (iSelectCell == 0)
			{
				int ncount = (int)vecSelect.size();
				irowStart = vecSelect[0].row;
				irowEnd = vecSelect[ncount - 1].row + 1;
			}
			else if (iSelectCell == 1)
			{
				bSetGroup = true;
				_GetGroupRowStartEnd(irowStart, irowEnd);
			}
			
			// Hiển thị hộp thoại
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DoModal();
		}
	}
	catch (...) {}
}

void DlgCheckFiles::_GetAllFileInFolder(CString szDir)
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

		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);

		// Delete list image icon
		m_imageList.DeleteImageList();
		ASSERT(m_imageList.GetSafeHandle() == NULL);
		m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);

		_LocdulieuNangcao();

		CString szval = L"";
		szval.Format(L"%ld", tongfilter);
		m_edtTotalFiles.SetWindowText(ff->_FormatTienVND(szval, L",", L"."));

		int ichkNew = m_check_new.GetCheck();
		if (ichkNew == 0) m_check_all.SetCheck(0);
	}
	catch (...) {}
}

void DlgCheckFiles::_SetDefault()
{
	try
	{
		m_lbl_total.SubclassDlgItem(CHKFS_LBL_TOTALCHECK, this);
		m_lbl_total.SetTextColor(rgbBlue);

		if (szTargetDirectory == L"") szTargetDirectory = pathFolderDir;
		if (szTargetDirectory.Right(1) == L"\\")
		{
			szTargetDirectory = szTargetDirectory.Left(szTargetDirectory.GetLength() - 1);
		}

		m_edtTargetDirectory.SubclassDlgItem(CHKFS_TXT_BROWSE, this);
		m_edtTargetDirectory.SetBkColor(rgbWhite);
		m_edtTargetDirectory.SetTextColor(rgbBlue);
		m_edtTargetDirectory.SetWindowText(szTargetDirectory);
		m_edtTargetDirectory.SetCueBanner(L"Chọn đường dẫn thư mục cần kiểm tra files...");

		m_edtTotalFiles.SubclassDlgItem(CHKFS_TXT_TOTALFILES, this);
		m_edtTotalFiles.SetBkColor(rgbWhite);
		m_edtTotalFiles.SetTextColor(rgbRed);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, 40);
		m_list_view.InsertColumn(iGFLType, L" Trạng thái", LVCFMT_CENTER, iwCol[1]);
		m_list_view.InsertColumn(iGFLTen, L" Tên file", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLSize, L" Kích cỡ", LVCFMT_RIGHT, iwCol[3]);
		m_list_view.InsertColumn(iGFLModified, L" Ngày sửa đổi", LVCFMT_CENTER, iwCol[4]);
		m_list_view.InsertColumn(iGFLPath, L" Đường dẫn", LVCFMT_LEFT, iwCol[5]);

		m_list_view.SetColumnColors(iGFLPath, rgbWhite, rgbBlue);
		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);

		m_check_new.SetCheck(bChkNew);
		m_check_link.SetCheck(bChkLink);

		// Không xét qua registry nữa mà xét theo tùy chọn <bSetGroup>
		if (bSetGroup) bChkLookInDir = true;
		else bChkLookInDir = false;

		m_chkLookInDir.SetCheck(bChkLookInDir);
		m_chkLookInSubdirectories.SetCheck(!bChkLookInDir);

/**********************************************************************************/

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

/**********************************************************************************/

		int icolGet = 0;
		bool blCategory = false;
		CString szval = L"", szCopy = shCopy;
		if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
		{
			blCategory = true;
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
			if (blCategory) m_cbb_path.SetCueBanner(L"Kích chọn thư mục sử dụng để lưu trữ...");
			else m_cbb_path.SetCueBanner(L"Kích chọn đường dẫn sử dụng để lưu trữ...");

			for (int i = 0; i < ncountPath; i++)
			{
				m_cbb_path.AddString(vecPath[i]);
			}
		}
		else m_cbb_path.EnableWindow(0);

		int ncount = (int)m_cbb_path.GetCount();
		if (ncount > 0)
		{
			int icheck = 0;
			for (int i = 0; i < ncount; i++)
			{
				m_cbb_path.GetLBText(i, szval);
				if (szTargetDirectory == szval || szTargetDirectory == pathFolderDir + szval)
				{
					if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
					{
						m_edtTargetDirectory.SetWindowText(pathFolderDir + szval);
					}
					else
					{
						m_edtTargetDirectory.SetWindowText(szval);
					}
					
					icheck = 1;
					break;
				}
			}

			if (icheck == 0)
			{
				m_cbb_path.GetLBText(0, szTargetDirectory);

				if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
				{
					m_edtTargetDirectory.SetWindowText(pathFolderDir + szTargetDirectory);
				}
				else
				{
					m_edtTargetDirectory.SetWindowText(szTargetDirectory);
				}
			}			
		}

		vecPath.clear();
	}
	catch (...) {}
}

void DlgCheckFiles::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CheckFiles;
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

		m_edtTargetDirectory.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgDirectory), 
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szval, 0)));

		szval.Format(L"%d", (int)m_chkLookInDir.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgCheckLookInDir), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_check_link.GetCheck());
		reg->WriteChar("CheckHyperLink", ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_check_new.GetCheck());
		reg->WriteChar("CheckFileNew", ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgCheckFiles::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CheckFiles;
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

		szval = bb->BaseDecodeEx(reg->ReadString(DlgDirectory, L""));
		if (szval != L"") szTargetDirectory = szval;

		int icheck = 0;
		szval = reg->ReadString(DlgCheckLookInDir, L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
			bChkLookInDir = (bool)icheck;
		}

		szval = reg->ReadString(L"CheckHyperLink", L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
			bChkLink = (bool)icheck;
		}

		szval = reg->ReadString(L"CheckFileNew", L"");
		if (szval != L"")
		{
			icheck = _wtoi(szval);
			if (icheck != 1) icheck = 0;
			bChkNew = (bool)icheck;
		}
	}
	catch (...) {}
}

void DlgCheckFiles::_SetLocationAndSize()
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

void DlgCheckFiles::_GetGroupRowStartEnd(int &iRowStart, int &iRowEnd)
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

void DlgCheckFiles::_XacdinhSheetLienquan()
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

void DlgCheckFiles::_DanhsachHyperlink()
{
	try
	{
		vecHyper.clear();
		int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
		RangePtr PRS = pShActive->GetRange(
			pRgActive->GetItem(irowStart, icotFilegoc),
			pRgActive->GetItem(iRowLast, icotFilegoc));

		ff->_xlGetAllHyperLink(PRS, vecHyper, 1, 1);
	}
	catch (...) {}
}

void DlgCheckFiles::_LocdulieuNangcao()
{
	try
	{
		slnew = 0;
		tongfilter = 0;
		vecFilter.clear();

		if (tongkq == 0) return;

		MyStrFilter MSFT;
		int ichkNew = m_check_new.GetCheck();
		int ichkLink = m_check_link.GetCheck();

		if (ichkNew == 0 && ichkLink == 0) return;
		else if (ichkNew == 1 && ichkLink == 0)
		{
			for (long i = 0; i < tongkq; i++)
			{
				MSFT.szItem = vecstrFileList[i];
				if (!(MSFT.szItem.Find(L"~$") >= 0 || MSFT.szItem.Right(4) == L".lnk"))
				{
					if (_CheckFileNew(MSFT.szItem) == 1)
					{
						slnew++;
						MSFT.iEnable = 1;
						MSFT.iFileNew = 1;
						vecFilter.push_back(MSFT);
					}
				}
			}
		}
		else if (ichkNew == 0 && ichkLink == 1)
		{
			for (long i = 0; i < tongkq; i++)
			{
				MSFT.szItem = vecstrFileList[i];
				if (!(MSFT.szItem.Find(L"~$") >= 0 || MSFT.szItem.Right(4) == L".lnk"))
				{
					if (_CheckFileNew(MSFT.szItem) == 0)
					{
						MSFT.iEnable = 1;
						MSFT.iFileNew = 0;
						vecFilter.push_back(MSFT);
					}
				}
			}
		}
		else
		{
			for (long i = 0; i < tongkq; i++)
			{
				MSFT.szItem = vecstrFileList[i];
				if (!(MSFT.szItem.Find(L"~$") >= 0 || MSFT.szItem.Right(4) == L".lnk"))
				{
					MSFT.iEnable = 1;
					MSFT.iFileNew = _CheckFileNew(MSFT.szItem);
					if (MSFT.iFileNew == 1) slnew++;
					vecFilter.push_back(MSFT);
				}
			}
		}

/****** Load dữ liệu *********************************************************/

		tongfilter = (int)vecFilter.size();

		m_check_all.SetCheck(1);

		lanshow = 1;
		long iz = ibuocnhay;
		if (tongfilter <= iz)
		{
			iz = tongfilter;
			iStopload = 0;
		}
		else iStopload = 1;

		// Thêm dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz, ichkNew, ichkLink);

		if (ichkNew == 1) _SetStatusKetqua(slnew);
		else _SetStatusKetqua(0);
	}
	catch (...) {}
}

void DlgCheckFiles::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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

			int ichkNew = m_check_new.GetCheck();
			int ichkLink = m_check_link.GetCheck();

			// Thêm dữ liệu vào list kết quả
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz, ichkNew, ichkLink);

			m_list_view.EnsureVisible(ncount + (int)(b / a) - 5, FALSE);
		}
	}
	catch (...) {}
}

void DlgCheckFiles::_AddItemsListKetqua(int iBegin, int iEnd, int ichkNew, int ichkLink)
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

			if (ichkNew == 1)
			{
				if (vecFilter[i].iFileNew == 1) m_list_view.SetCheck(i, 1);
				else m_list_view.SetCheck(i, 0);
			}
			else m_list_view.SetCheck(i, 0);

			if (iBegin > 0 && dem == ncount)
			{
				m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
			}

			// Kiểm tra tệp tin đã được gán hyperlink chưa
			_CheckFileHyperlink(i);
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

void DlgCheckFiles::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 40;
	if (pSubItem == iGFLType) pNMListHeader->pitem->cxy = 100;
	*pResult = 0;
}

void DlgCheckFiles::_Timkiemdulieu()
{
	CString szpath = L"";
	m_edtTargetDirectory.GetWindowText(szpath);
	_GetAllFileInFolder(szpath);
}

void DlgCheckFiles::_SelectAllItems()
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

int DlgCheckFiles::_CheckFileNew(CString szpathFile)
{
	CString szRelatively = szpathFile.Right(szpathFile.GetLength() - iLenPathDir);
	int ncoutHyp = (int)vecHyper.size();

	for (int i = 0; i < ncoutHyp; i++)
	{
		if (szpathFile == vecHyper[i] || szRelatively == vecHyper[i]) return 0;
	}

	return 1;
}

void DlgCheckFiles::_CheckFileHyperlink(int nIndex)
{
	try
	{
		if (vecFilter[nIndex].iFileNew == 1)
		{
			m_list_view.SetItemText(nIndex, iGFLType, DF_FILEMOI);
			m_list_view.SetCellColors(nIndex, iGFLType, RGB(0, 176, 80), rgbWhite);
		}
		else
		{
			m_list_view.SetItemText(nIndex, iGFLType, DF_GANLINK);
			m_list_view.SetCellColors(nIndex, iGFLType, rgbWhite, rgbBlue);
		}
	}
	catch (...) {}
}

int DlgCheckFiles::_XacdinhDanhsachThumuc(vector<CString> &vecPath, int icolGet)
{
	try
	{
		vecPath.clear();

		CString szval = L"";

		if (bSetGroup)
		{
			// Trường hợp chọn group --> Chỉ lấy đường dẫn tại group được chọn
			szval = ff->_xlGIOR(pRgActive, irowStart, icolGet, L"");
			if (szval != L"") vecPath.push_back(szval);
		}
		else
		{
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
		}

		return (int)vecPath.size();
	}
	catch (...) {}
	return 0;
}

void DlgCheckFiles::OnCbnSelchangeCbbPath()
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

			szTargetDirectory = szpath;
			m_edtTargetDirectory.SetWindowText(szTargetDirectory);
			_GetAllFileInFolder(szTargetDirectory);
		}
	}
}

void DlgCheckFiles::OnBnClickedCheckAll()
{
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

void DlgCheckFiles::OnBnClickedCheckNews()
{
	m_edtTargetDirectory.GetWindowText(szTargetDirectory);
	_GetAllFileInFolder(szTargetDirectory);
}

void DlgCheckFiles::OnBnClickedCheckLink()
{
	m_edtTargetDirectory.GetWindowText(szTargetDirectory);
	_GetAllFileInFolder(szTargetDirectory);
}

void DlgCheckFiles::_PutStyleCell(RangePtr PRS)
{
	PRS->EntireRow->Font->PutName(FontTimes);
	PRS->EntireRow->Font->PutSize(13);
	PRS->EntireRow->Font->PutBold(false);
	PRS->EntireRow->Font->PutItalic(false);
	PRS->EntireRow->Font->PutColor(rgbBlack);
	PRS->EntireRow->Interior->PutColor(xlNone);
}

int DlgCheckFiles::_GetRowGroups(CString szpathFile)
{
	try
	{
		CString szval = L"";
		int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
		for (int i = irowStart; i <= iRowLast; i++)
		{
			szval = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
			if (szval == szpathFile)
			{
				szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
				if (_wtoi(szval) > 0)
				{
					RangePtr PRS = pRgActive->GetItem(i, 1);
					PRS->EntireRow->Insert(xlShiftDown);

					PRS = pRgActive->GetItem(i, 1);
					_PutStyleCell(PRS);

					return i;
				}
			}
		}		
	}
	catch (...) {}
	return irowStart;
}

void DlgCheckFiles::_AddFilesCategory(int iOnClickBtn)
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
			m_edtTargetDirectory.GetWindowTextW(szTargetDirectory);
			CString szTargetGoc = ff->_GetNameOfPath(szTargetDirectory, pos, 0);
			CString szNameDir = szTargetGoc;
			
			int iLenTarget = szTargetDirectory.GetLength();
			int iLookInSub = (int)m_chkLookInSubdirectories.GetCheck();

			if (iOnClickBtn == 1)
			{
				_SaveRegistry();
				CDialogEx::OnOK();
			}

			ff->_xlPutScreenUpdating(false);

			RangePtr PRS;
			if (jcotFolder == 0)
			{
				PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, 1), pRgActive->GetItem(irowStart + nSelect - 1, 1));
				PRS->EntireRow->Insert(xlShiftDown);

				PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, 1), pRgActive->GetItem(irowStart + nSelect - 1, 1));
				_PutStyleCell(PRS);
			}			

			int iRowChen = 0;
			CString szval = L"", szdir = L"", szname = L"", szndung = L"", sztype = L"";
			CString szUpdateStatus = L"Đang thêm mới dữ liệu. Vui lòng đợi trong giây lát...";
			for (int i = 0; i < nSelect; i++)
			{
				if (iOnClickBtn == 0)
				{
					m_list_view.SetItemText(vecRow[i], iGFLType, DF_GANLINK);
					m_list_view.SetCellColors(vecRow[i], iGFLType, rgbWhite, rgbBlue);
				}

				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / nSelect), L"%");
				ff->_xlPutStatus(szval);

				szname = ff->_GetNameOfPath(vecFilter[vecRow[i]].szItem, pos, -1);
				szdir = vecFilter[vecRow[i]].szItem.Left(pos);

				// Xác định vị trí đổ chèn dữ liệu vào nhóm
				if (jcotFolder > 0) iRowChen = _GetRowGroups(szdir);
				else iRowChen = irowStart + i;

				pRgActive->PutItem(iRowChen, jcotTen, (_bstr_t)szname);

				if (jcotFolder > 0 && pos > 0)
				{
					szdir = vecFilter[vecRow[i]].szItem.Left(pos);
					pRgActive->PutItem(iRowChen, jcotFolder, (_bstr_t)szdir);
				}

				sztype = ff->_GetTypeOfFile(vecFilter[vecRow[i]].szItem);
				if (jcotType > 0) pRgActive->PutItem(iRowChen, jcotType, (_bstr_t)sztype);
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

					pRgActive->PutItem(iRowChen, jcotNoidung, (_bstr_t)szndung);
				}

				szval = ff->_GetNameOfPath(vecFilter[vecRow[i]].szItem, pos, 1);
				szNameDir = ff->_GetNameOfPath(szval, pos, 0);
				pRgActive->PutItem(iRowChen, jcotLV, (_bstr_t)szNameDir);

				PRS = pRgActive->GetItem(iRowChen, icotFilegoc);
				szval = vecFilter[vecRow[i]].szItem;

				ff->_xlSetHyperlink(pShActive, PRS, szval);
			}

			_AutoSothutu(pShActive, 1, 0, 0);

			PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, jcotTen),
				pRgActive->GetItem(irowStart + nSelect - 1, jcotTen));
			PRS->PutHorizontalAlignment(xlLeft);

			if (jcotNoidung > 0)
			{
				PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, jcotNoidung),
					pRgActive->GetItem(irowStart + nSelect - 1, jcotNoidung));
				PRS->PutHorizontalAlignment(xlLeft);
			}

			if (jcotFolder > 0)
			{
				PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, jcotFolder),
					pRgActive->GetItem(irowStart + nSelect - 1, jcotFolder));
				PRS->PutHorizontalAlignment(xlLeft);
			}

			PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, jcotSTT),
				pRgActive->GetItem(irowStart + nSelect - 1, icotEnd));
			ff->_xlFormatBorders(PRS, 3, true, true);
			PRS->Rows->AutoFit();

			iRowActive += nSelect;
			PRS = pRgActive->GetItem(iRowActive, 1);
			PRS->Select();

			ff->_xlPutStatus(L"Ready");
			if (iOnClickBtn == 1) ff->_msgbox(L"Cập nhật dữ liệu mới thành công!", 0, 1, 2000);
			ff->_xlPutScreenUpdating(true);
		}

		vecRow.clear();
	}
	catch (...) {}
}

int DlgCheckFiles::_GetAllSelectedItems(vector<int> &vecRow)
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

int DlgCheckFiles::_GetAllCheckedItems(vector<int> &vecRow)
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

void DlgCheckFiles::_AutoSothutu(_WorksheetPtr pSheet, int numStart, int nRowStart, int nRowEnd)
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

void DlgCheckFiles::OnChkflAddfiles()
{
	_AddFilesCategory(0);
}


void DlgCheckFiles::OnChkflGotosheet()
{
	try
	{
		if (nItem >= 0 && nSubItem >= 0)
		{
			// Xác định vị trí chứa hyperlink
			CString szval = m_list_view.GetItemText(nItem, iGFLType);
			if (szval == DF_GANLINK)
			{
				szval = m_list_view.GetItemText(nItem, iGFLPath);
				if (szval != L"")
				{
					RangePtr PRS;
					CString szlink = L"";
					int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
					int ivtri = 0;

					for (int i = irowStart; i < iRowLast; i++)
					{
						PRS = pRgActive->GetItem(i, icotFilegoc);
						szlink = ff->_xlGetHyperLink(PRS);
						if (szlink != L"" && szlink == szval)
						{
							ivtri = i;
							break;
						}
					}

					if (ivtri > 0)
					{
						_SaveRegistry();
						CDialogEx::OnOK();

						int iColumn = jcotTen;
						PRS = pRgActive->GetItem(ivtri, iColumn);

						bool blHidden = PRS->EntireRow->GetHidden();
						if (blHidden == true) PRS->EntireRow->PutHidden(false);

						PRS->Select();

						if (ivtri > 1) ivtri--;
						if (iColumn > 2) iColumn -= 2;

						xl->ActiveWindow->PutScrollRow(ivtri);
						xl->ActiveWindow->PutScrollColumn(iColumn);
					}
					else
					{
						ff->_msgbox(L"Không tìm thấy vị trí file tương ứng trong bảng tính.", 0, 0);
					}
				}
			}
		}
	}
	catch (...) {}
}


void DlgCheckFiles::OnChkflOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}


void DlgCheckFiles::OnChkflOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}


void DlgCheckFiles::OnChkflDelfiles()
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

void DlgCheckFiles::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
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

void DlgCheckFiles::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}

void DlgCheckFiles::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
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

int DlgCheckFiles::_GetCountItemsCheck()
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

void DlgCheckFiles::_GetItemSelect(int icheck)
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

void DlgCheckFiles::_SetStatusKetqua(int num)
{
	CString str = L"";
	if (num > 1) str.Format(L"Bạn đã chọn (%d) files / ", num);
	else str.Format(L"Bạn đã chọn (%d) file / ", num);
	m_lbl_total.SetWindowText(str);
}

void DlgCheckFiles::OnChkflCheck()
{
	_GetItemSelect(1);
}


void DlgCheckFiles::OnChkflUncheck()
{
	_GetItemSelect(0);
}


void DlgCheckFiles::OnChkflCheckall()
{
	_SelectAllItems();
}


void DlgCheckFiles::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_SPACE) OnChkflAddfiles();
	if (pLVKeyDow->wVKey == VK_DELETE) OnChkflDelfiles();
	*pResult = 0;
}


void DlgCheckFiles::OnBnClickedRadioDir()
{
	_Timkiemdulieu();
}


void DlgCheckFiles::OnBnClickedRadioSubdir()
{
	_Timkiemdulieu();
}
