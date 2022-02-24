#include "pch.h"
#include "DlgAutoRename.h"

IMPLEMENT_DYNAMIC(DlgAutoRename, CDialogEx)

DlgAutoRename::DlgAutoRename(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgAutoRename::IDD, pParent)
{
	nIndexVec = 0;
	nCountList = 0;
	nIndexSplit = 0;
	iTabDefault = 0;
	szTargetDirectory = L"";

	blDragging = false;
	iGFLTenNew = 1, iGFLTenOld = 2, iGFLPath = 3;

	nTotalCol = getSizeArray(iwCol);
	iwCol[0] = 40, iwCol[1] = 200, iwCol[2] = 200, iwCol[3] = 100;

	nItem = -1, nSubItem = -1;
	iTypeApply = 0;
	iKeyESC = 0;

	szBefore = L"", szAfter = L"";

	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
}

DlgAutoRename::~DlgAutoRename()
{
	vecList.clear();
	vecItems.clear();
	
	delete m_frmSTT;
	delete m_frmAdd;
	delete m_frmDel;
	delete m_frmRep;
	delete m_frmKhac;

	delete ff;
	delete bb;
	delete reg;
}

void DlgAutoRename::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, AUTORM_TABMAIN, m_tab_main);
	DDX_Control(pDX, AUTORM_SPLIT_BROWSE, m_splbtn_browse);
	DDX_Control(pDX, AUTORM_TXT_BROWSE, m_txt_browse);
	DDX_Control(pDX, AUTORM_EDIT_LIST, m_edit_autorename);
	DDX_Control(pDX, AUTORM_LIST_VIEW, m_list_view);
	DDX_Control(pDX, AUTORM_LBL_TOTAL, m_lbl_total);
	DDX_Control(pDX, AUTORM_LBL_HDAN, m_lbl_hdan);	
	DDX_Control(pDX, AUTORM_BTN_APDUNG, m_btn_apply);
	DDX_Control(pDX, AUTORM_BTN_RESET, m_btn_reset);
	DDX_Control(pDX, AUTORM_BTN_UNDO, m_btn_undo);
	DDX_Control(pDX, AUTORM_BTN_REDO, m_btn_redo);
	DDX_Control(pDX, AUTORM_BTN_OK, m_btn_ok);
	DDX_Control(pDX, AUTORM_BTN_CANCEL, m_btn_cancel);
}

BEGIN_MESSAGE_MAP(DlgAutoRename, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_WM_TIMER()

	ON_BN_CLICKED(AUTORM_BTN_OK, &DlgAutoRename::OnBnClickedBtnOk)
	ON_BN_CLICKED(AUTORM_BTN_CANCEL, &DlgAutoRename::OnBnClickedBtnCancel)
	ON_BN_CLICKED(AUTORM_SPLIT_BROWSE, &DlgAutoRename::OnBnClickedSplitBrowse)
	ON_BN_CLICKED(AUTORM_BTN_APDUNG, &DlgAutoRename::OnBnClickedBtnApdung)
	ON_BN_CLICKED(AUTORM_BTN_RESET, &DlgAutoRename::OnBnClickedBtnReset)
	ON_BN_CLICKED(AUTORM_BTN_UNDO, &DlgAutoRename::OnBnClickedBtnUndo)
	ON_BN_CLICKED(AUTORM_BTN_REDO, &DlgAutoRename::OnBnClickedBtnRedo)
	ON_BN_CLICKED(AUTORM_CHECK_ALL, &DlgAutoRename::OnBnClickedCheckAll)
	ON_COMMAND(MN_AUTORN_FOLDER, &DlgAutoRename::OnAutornFolder)
	ON_COMMAND(MN_AUTORN_FILES, &DlgAutoRename::OnAutornFiles)	
	ON_COMMAND(MN_AUTORN_APPLY, &DlgAutoRename::OnAutornApply)
	ON_COMMAND(MN_AUTORN_OPENFILE, &DlgAutoRename::OnAutornOpenfile)
	ON_COMMAND(MN_AUTORN_OPENFOLDER, &DlgAutoRename::OnAutornOpenfolder)
	ON_COMMAND(MN_AUTORN_UP, &DlgAutoRename::OnAutornUp)
	ON_COMMAND(MN_AUTORN_DOWN, &DlgAutoRename::OnAutornDown)
	ON_COMMAND(MN_AUTORN_SEL, &DlgAutoRename::OnAutornSel)
	ON_COMMAND(MN_AUTORN_UNSEL, &DlgAutoRename::OnAutornUnsel)
	ON_COMMAND(MN_AUTORN_SELALL, &DlgAutoRename::OnAutornSelall)
	ON_COMMAND(MN_AUTORN_REVERSE, &DlgAutoRename::OnAutornReverse)
	ON_COMMAND(MN_AUTORN_CHFILES, &DlgAutoRename::OnAutornChfiles)
	ON_COMMAND(MN_AUTORN_CHFOLDER, &DlgAutoRename::OnAutornChfolder)
	ON_EN_SETFOCUS(AUTORM_EDIT_LIST, &DlgAutoRename::OnEnSetfocusEditList)
	ON_EN_KILLFOCUS(AUTORM_EDIT_LIST, &DlgAutoRename::OnEnKillfocusEditList)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgAutoRename::OnHdnEndtrackListView)	
	ON_NOTIFY(NM_CLICK, AUTORM_LIST_VIEW, &DlgAutoRename::OnNMClickListView)
	ON_NOTIFY(NM_DBLCLK, AUTORM_LIST_VIEW, &DlgAutoRename::OnNMDblclkListView)
	ON_NOTIFY(TCN_SELCHANGE, AUTORM_TABMAIN, &DlgAutoRename::OnTcnSelchangeTabmain)
	ON_EN_SETFOCUS(AUTORM_TXT_BROWSE, &DlgAutoRename::OnEnSetfocusTxtBrowse)
	ON_NOTIFY(LVN_KEYDOWN, AUTORM_LIST_VIEW, &DlgAutoRename::OnLvnKeydownListView)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgAutoRename)
	EASYSIZE(AUTORM_LBL_PATH, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(AUTORM_TXT_BROWSE, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(AUTORM_SPLIT_BROWSE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(AUTORM_TABMAIN, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(AUTORM_LBL_TOTAL, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(AUTORM_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_APDUNG, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_RESET, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_UNDO, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_REDO, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_LBL_HDAN, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(AUTORM_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgAutoRename message handlers
BOOL DlgAutoRename::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_RENAME));
	SetIcon(hIcon, FALSE);

	_LoadRegistry();
	_SetDefault();

	mnIcon.GdiplusStartupInputConfig();

	_Timkiemdulieu();

	SetTimer(CONST_TIMER_CLISTEX, 100, NULL);	// 1000ms = 1 second
	
	hWndAutoRename = m_hWnd;
	
	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgAutoRename::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(AUTORM_TXT_BROWSE))
			{
				m_txt_browse.GetWindowTextW(szTargetDirectory);
				szTargetDirectory.Trim();

				if (szTargetDirectory != L"") _Timkiemdulieu();
				else OnBnClickedSplitBrowse();

				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(AUTORM_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			if (iKeyESC == 0) OnBnClickedBtnCancel();
			else
			{
				szBefore = m_list_view.GetItemText(CLRow, CLColumn);
				m_edit_autorename.SetWindowTextW(szBefore);
				GotoDlgCtrl(GetDlgItem(AUTORM_LIST_VIEW));
			}
			iKeyESC = 0;

			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(AUTORM_LIST_VIEW))
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

void DlgAutoRename::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgAutoRename::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgAutoRename::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgAutoRename::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_AUTORENAME));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;

			vector<int> vecRow;

			int nSelect = _GetAllSelectedItems(vecRow);
			
			CString szAppy = L"", szFile = L"", szFolder = L"";
			if (nIndexSplit == 0)
			{
				szFile = L"Mở thư mục được chọn";
				szAppy = L"Áp dụng cho thư mục được chọn";				
				szFolder = L"Mở thư mục chứa thư mục được chọn";

				if (nSelect >= 2)
				{
					szAppy.Format(L"Áp dụng cho (%d) thư mục được chọn", nSelect);
				}

				ModifyMenu(pMenu, MN_AUTORN_APPLY, MF_BYCOMMAND | MF_STRING, MN_AUTORN_APPLY, szAppy);
				ModifyMenu(pMenu, MN_AUTORN_OPENFILE, MF_BYCOMMAND | MF_STRING, MN_AUTORN_OPENFILE, szFile);
				ModifyMenu(pMenu, MN_AUTORN_OPENFOLDER, MF_BYCOMMAND | MF_STRING, MN_AUTORN_OPENFOLDER, szFolder);
			}
			else
			{
				if (nSelect >= 2)
				{
					szAppy.Format(L"Áp dụng cho (%d) files được chọn", nSelect);
					ModifyMenu(pMenu, MN_AUTORN_APPLY, MF_BYCOMMAND | MF_STRING, MN_AUTORN_APPLY, szAppy);
				}

				EnableMenuItem(pMenu, MN_AUTORN_CHFILES, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_AUTORN_CHFILES, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_APPLY, IDI_ICON_CHECK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_OPENFOLDER, IDI_ICON_OPENFOLDER);

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_UP, IDI_ICON_UP);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_DOWN, IDI_ICON_DOWN);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_REVERSE, IDI_ICON_REVERSE);
			
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_CHFILES, IDI_ICON_SEARCH);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_CHFOLDER, IDI_ICON_SEARCH);

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_SEL, IDI_ICON_OK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_UNSEL, IDI_ICON_CANCEL);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_AUTORN_SELALL, IDI_ICON_TODOLIST);

			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_AUTORN_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_AUTORN_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

				EnableMenuItem(pMenu, MN_AUTORN_UP, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_AUTORN_DOWN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			if (nSelect != 2) EnableMenuItem(pMenu, MN_AUTORN_REVERSE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

			if (nItem == 0) EnableMenuItem(pMenu, MN_AUTORN_UP, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			else if (nItem + 1 == nCountList) EnableMenuItem(pMenu, MN_AUTORN_DOWN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);

			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(AUTORM_LIST_VIEW));
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

void DlgAutoRename::_SetDefault()
{
	try
	{
		_CreateTabMain();

		if (nIndexSplit == 0)
		{
			m_splbtn_browse.SetWindowText(L"Thư mục");
			m_btn_apply.SetWindowText(L"Áp dụng thư mục được chọn");
		}
		else
		{
			m_splbtn_browse.SetWindowText(L"Files");
			m_btn_apply.SetWindowText(L"Áp dụng files được tích chọn");
		}

		CButton *m_checkAll = (CButton*)GetDlgItem(AUTORM_CHECK_ALL);
		m_checkAll->SetCheck(true);

		m_edit_autorename.SubclassDlgItem(AUTORM_EDIT_LIST, this);
		m_edit_autorename.SetBkColor(rgbWhite);
		m_edit_autorename.SetTextColor(rgbBlue);

		m_txt_browse.SubclassDlgItem(AUTORM_TXT_BROWSE, this);
		m_txt_browse.SetBkColor(rgbWhite);
		m_txt_browse.SetTextColor(rgbBlue);
		m_txt_browse.SetWindowText(szTargetDirectory);
		m_txt_browse.SetCueBanner(L"Chọn đường dẫn thư mục chứa cần chỉnh sửa...");

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_splbtn_browse.SetDropDownMenu(IDR_MENU_AUTORENAME, 1);

		m_btn_apply.SetThemeHelper(&m_ThemeHelper);
		m_btn_apply.SetIcon(IDI_ICON_CHECK, 48, 48);
		m_btn_apply.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_apply.SetFont(&m_FontHeaderList);
		m_btn_apply.SetWindowText(L"Áp dụng thư mục được tích chọn");

		m_btn_reset.SetThemeHelper(&m_ThemeHelper);
		m_btn_reset.SetIcon(IDI_ICON_PROPERTIES, 24, 24);
		m_btn_reset.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_reset.SetTooltipText(L"Khôi phục lại tên dữ liệu gốc");

		m_btn_undo.SetThemeHelper(&m_ThemeHelper);
		m_btn_undo.SetIcon(IDI_ICON_UNDO, 24, 24);
		m_btn_undo.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_undo.SetTooltipText(L"Hoàn tác tên cũ đã sửa");

		m_btn_redo.SetThemeHelper(&m_ThemeHelper);
		m_btn_redo.SetIcon(IDI_ICON_REDO, 24, 24);
		m_btn_redo.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_redo.SetTooltipText(L"Quay trở lại tên mới sửa tiếp theo");
		m_btn_redo.SetAlign(2, 1);

		m_lbl_total.SubclassDlgItem(AUTORM_LBL_TOTAL, this);
		m_lbl_total.SetTextColor(rgbBlue);
		m_lbl_total.SetFont(&m_FontHeaderList);
		m_lbl_total.SetWindowText(L"");

		m_lbl_hdan.SubclassDlgItem(AUTORM_LBL_HDAN, this);
		m_lbl_hdan.SetTextColor(rgbBlue);
		m_lbl_hdan.SetFont(&m_FontItalic);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, 40);
		m_list_view.InsertColumn(iGFLTenNew, L" Tên file (sau chỉnh sửa)", LVCFMT_LEFT, iwCol[1]);
		m_list_view.InsertColumn(iGFLTenOld, L" Tên file (gốc)", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLPath, L" Đường dẫn", LVCFMT_LEFT, iwCol[3]);

		m_list_view.SetColumnEditor(iGFLTenNew, &m_edit_autorename);
		m_list_view.SetColumnColors(iGFLPath, rgbWhite, rgbBlue);
		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);
	}
	catch (...) {}
}

void DlgAutoRename::_CreateTabMain()
{
	try
	{
		m_tab_main.InsertItem(0, L"STT");
		m_tab_main.InsertItem(1, L"Thêm");
		m_tab_main.InsertItem(2, L"Xóa");
		m_tab_main.InsertItem(3, L"Thay thế");
		m_tab_main.InsertItem(4, L"Khác");

		CRect rec;
		m_tab_main.GetClientRect(&rec);
		rec.left += 1;
		rec.top += 24;
		rec.right -= 3;
		rec.bottom -= 2;

		m_tab_main.SetFont(&m_FontHeaderList);

		m_frmSTT = new FrmTabSTT(&m_tab_main);
		m_frmSTT->Create(FrmTabSTT::IDD, &m_tab_main);
		m_frmSTT->MoveWindow(&rec);

		m_frmAdd = new FrmTabAdd(&m_tab_main);
		m_frmAdd->Create(FrmTabAdd::IDD, &m_tab_main);
		m_frmAdd->MoveWindow(&rec);

		m_frmDel = new FrmTabDelete(&m_tab_main);
		m_frmDel->Create(FrmTabDelete::IDD, &m_tab_main);
		m_frmDel->MoveWindow(&rec);

		m_frmRep = new FrmTabReplace(&m_tab_main);
		m_frmRep->Create(FrmTabReplace::IDD, &m_tab_main);
		m_frmRep->MoveWindow(&rec);

		m_frmKhac = new FrmTabKhac(&m_tab_main);
		m_frmKhac->Create(FrmTabKhac::IDD, &m_tab_main);
		m_frmKhac->MoveWindow(&rec);

		// Mặc định luôn để tab đầu tiên (iTabDefault = 0)
		m_tab_main.SetCurSel(iTabDefault);

		m_frmSTT->ShowWindow(SW_SHOW);

		_LoadRegistryTab();
	}
	catch (...) {}
}


void DlgAutoRename::OnTcnSelchangeTabmain(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	iTabDefault = (int)m_tab_main.GetCurSel();
	switch (iTabDefault)
	{
		case 0:
		{
			m_frmSTT->ShowWindow(SW_SHOW);	// m_frmSTT = 0
			m_frmAdd->ShowWindow(SW_HIDE);
			m_frmDel->ShowWindow(SW_HIDE);
			m_frmRep->ShowWindow(SW_HIDE);
			m_frmKhac->ShowWindow(SW_HIDE);
		}
		break;

		case 1:
		{
			m_frmSTT->ShowWindow(SW_HIDE);
			m_frmAdd->ShowWindow(SW_SHOW);	// m_frmAdd = 1
			m_frmDel->ShowWindow(SW_HIDE);
			m_frmRep->ShowWindow(SW_HIDE);
			m_frmKhac->ShowWindow(SW_HIDE);
		}
		break;

		case 2:
		{
			m_frmSTT->ShowWindow(SW_HIDE);
			m_frmAdd->ShowWindow(SW_HIDE);
			m_frmDel->ShowWindow(SW_SHOW);	// m_frmDel = 2
			m_frmRep->ShowWindow(SW_HIDE);
			m_frmKhac->ShowWindow(SW_HIDE);
		}
		break;

		case 3:
		{
			m_frmSTT->ShowWindow(SW_HIDE);
			m_frmAdd->ShowWindow(SW_HIDE);
			m_frmDel->ShowWindow(SW_HIDE);
			m_frmRep->ShowWindow(SW_SHOW);	// m_frmRep = 3
			m_frmKhac->ShowWindow(SW_HIDE);
		}
		break;

		case 4:
		{
			m_frmSTT->ShowWindow(SW_HIDE);
			m_frmAdd->ShowWindow(SW_HIDE);
			m_frmDel->ShowWindow(SW_HIDE);
			m_frmRep->ShowWindow(SW_HIDE);
			m_frmKhac->ShowWindow(SW_SHOW);	// m_frmKhac = 4
		}
		break;


		default:
			break;
	}

	*pResult = 0;
}

void DlgAutoRename::OnBnClickedBtnOk()
{
	MyStrItems MSITEMS;
	vector< MyStrItems> vecRename;

	if (nCountList > 0)
	{
		for (int i = 0; i < nCountList; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1)
			{
				MSITEMS.szFullpath = m_list_view.GetItemText(i, iGFLPath);				
				MSITEMS.szName = m_list_view.GetItemText(i, iGFLTenOld);
				MSITEMS.szNew = m_list_view.GetItemText(i, iGFLTenNew);
				MSITEMS.szType = L"";
				MSITEMS.nIndex = i;

				vecRename.push_back(MSITEMS);
			}
		}
	}

	int nsize = (int)vecRename.size();
	if (nsize == 0)
	{
		ff->_msgbox(L"Bạn chưa tích chọn dữ liệu cần đổi tên.", 2, 0);
		return;
	}

	CString szval = L"";
	szval.Format(L"Bạn có chắc chắn đổi tên (%d) dữ liệu được tích chọn?", nsize);
	int result = ff->_y(szval, 0, 0, 0);
	if (result != 6) return;	

/******************************************************************************/

	int pos = -1;
	CString szDir = ff->_GetNameOfPath(vecRename[0].szFullpath, pos, 1);
	if (szDir.Right(1) != L"\\") szDir += L"\\";

	int dem = 0, nID = 0;
	CString szDirOld = L"", szChange = L"";
	CString szUpdateStatus = L"Đang tiến hành đổi tên dữ liệu. Vui lòng đợi trong giây lát...";
	for (int i = 0; i < nsize; i++)
	{
		szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100*(i + 1)/nsize), L"%");
		ff->_xlPutStatus(szval);

		szDirOld = szDir + vecRename[i].szName;
		szChange = szDir + vecRename[i].szNew;

		if (nIndexSplit == 1)
		{
			szval = ff->_GetTypeOfFile(vecRename[i].szFullpath);
			szDirOld += (L"." + szval);
			szChange += (L"." + szval);
		}

		nID = vecRename[i].nIndex;

		if ((int)_wrename(szDirOld, szChange) == 0)
		{
			m_list_view.SetCellColors(nID, iGFLTenNew, rgbWhite, rgbBlack);
			m_list_view.SetItemText(nID, iGFLTenNew, vecRename[i].szNew);
			m_list_view.SetItemText(nID, iGFLPath, szChange);

			dem++;
		}
		else
		{
			m_list_view.SetCellColors(nID, iGFLTenNew, rgbWhite, rgbRed);
		}
	}

	szval.Format(L"Thay đổi thành công (%d/%d) dữ liệu. \nBạn có muốn mở thư mục kiểm tra?", dem, nsize);
	result = ff->_y(szval, 0, 0, 0);
	if (result == 6)
	{
		ShellExecute(NULL, L"open", szDir, NULL, NULL, SW_SHOWMAXIMIZED);
	}

	ff->_xlPutStatus(L"Ready");	
	vecRename.clear();
}

void DlgAutoRename::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

void DlgAutoRename::OnBnClickedSplitBrowse()
{
	iKeyESC = 0;
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần chỉnh sửa...";
	if (szTargetDirectory != L"") m_dlg.m_ofn.lpstrInitialDir = szTargetDirectory;
	if (m_dlg.DoModal() == IDOK)
	{
		szTargetDirectory = m_dlg.GetPathName();
		m_txt_browse.SetWindowText(szTargetDirectory);
		
		_Timkiemdulieu();
	}
}

void DlgAutoRename::_Timkiemdulieu()
{
	try
	{
		m_txt_browse.GetWindowTextW(szTargetDirectory);
		szTargetDirectory.Trim();

		if (szTargetDirectory != L"")
		{
			if (nIndexSplit == 0)
			{
				nCountList = _TimkiemThumuc(szTargetDirectory);
			}
			else
			{
				nCountList = _TimkiemFiles(szTargetDirectory);
			}

			_LoadAllKetqua();
			_SetTotalKetqua();
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnAutornFolder()
{
	iKeyESC = 0;
	nIndexSplit = 0;
	m_splbtn_browse.SetWindowText(L"Thư mục");
	m_btn_apply.SetWindowText(L"Áp dụng thư mục được chọn");

	m_txt_browse.GetWindowTextW(szTargetDirectory);
	szTargetDirectory.Trim();

	if (szTargetDirectory != L"")
	{
		nCountList = _TimkiemThumuc(szTargetDirectory);

		_LoadAllKetqua();
		_SetTotalKetqua();
	}
	else
	{
		OnBnClickedSplitBrowse();
	}	
}

void DlgAutoRename::OnAutornFiles()
{
	iKeyESC = 0;
	nIndexSplit = 1;
	m_splbtn_browse.SetWindowText(L"Files");
	m_btn_apply.SetWindowText(L"Áp dụng files được tích chọn");

	m_txt_browse.GetWindowTextW(szTargetDirectory);
	szTargetDirectory.Trim();

	if (szTargetDirectory != L"")
	{
		nCountList = _TimkiemFiles(szTargetDirectory);

		_LoadAllKetqua();
		_SetTotalKetqua();
	}
	else
	{
		OnBnClickedSplitBrowse();
	}	
}

int DlgAutoRename::_TimkiemThumuc(CString szPath)
{
	try
	{
		vecList.clear();
		vecItems.clear();

		if (szPath == L"")
		{
			m_txt_browse.GetWindowTextW(szPath);
			szPath.Trim();
			if (szPath == L"") return 0;
		}

		vector<CString> vecFolder;
		ff->_FindAllFolder(szPath, vecFolder);

		int ncountFolder = (int)vecFolder.size();
		if (ncountFolder > 0)
		{
			int pos = -1;
			MyStrItems MSITEMS;
			vector<CString> vecArr;
			for (int i = 0; i < ncountFolder; i++)
			{
				MSITEMS.szFullpath = vecFolder[i];
				MSITEMS.szName = ff->_GetNameOfPath(MSITEMS.szFullpath, pos, 0);
				MSITEMS.szType = L"";

				if (MSITEMS.szName.Find(L"~$") == -1)
				{
					vecItems.push_back(MSITEMS);
					vecArr.push_back(MSITEMS.szName);
				}
			}

			vecList.push_back(vecArr);
		}

		vecFolder.clear();

		nIndexVec = 1;

		return (int)vecItems.size();
	}
	catch (...) {}
	return 0;
}

int DlgAutoRename::_TimkiemFiles(CString szPath)
{
	try
	{
		vecList.clear();
		vecItems.clear();

		if (szPath == L"")
		{
			m_txt_browse.GetWindowTextW(szPath);
			szPath.Trim();
			if (szPath == L"") return 0;
		}

		vector<CString> vecFiles;
		int ncountFile = ff->_FindAllFileEx(szPath, vecFiles);
		if (ncountFile > 0)
		{
			int pos = -1;
			MyStrItems MSITEMS;
			vector<CString> vecArr;
			for (int i = 0; i < ncountFile; i++)
			{
				MSITEMS.szFullpath = vecFiles[i];
				MSITEMS.szName = ff->_GetNameOfPath(MSITEMS.szFullpath, pos, -1);
				MSITEMS.szType = ff->_GetTypeOfFile(MSITEMS.szFullpath);

				if (MSITEMS.szName.Find(L"~$") == -1
					&& MSITEMS.szType != L"lnk")
				{
					vecItems.push_back(MSITEMS);
					vecArr.push_back(MSITEMS.szName);
				}
			}

			vecList.push_back(vecArr);
		}

		vecFiles.clear();

		nIndexVec = 1;

		return (int)vecItems.size();
	}
	catch (...) {}
	return 0;
}

int DlgAutoRename::_GetAllSelectedItems(vector<int> &vecRow)
{
	try
	{
		vecRow.clear();
		if (nItem == -1 || nSubItem == -1) return 0;

		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			for (int i = 0; i < nCountList; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vecRow.push_back(i);
				}
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void DlgAutoRename::_SelectAllItems()
{
	if (nCountList > 0)
	{
		for (int i = 0; i < nCountList; i++)
		{
			m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
		}
	}
}

void DlgAutoRename::_DeleteAllImages()
{
	m_imageList.DeleteImageList();
	ASSERT(m_imageList.GetSafeHandle() == NULL);
	m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);
}

void DlgAutoRename::_DeleteListKetqua()
{
	if ((int)m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}	
}

void DlgAutoRename::_LoadAllKetqua()
{
	try
	{
		_DeleteAllImages();
		_DeleteListKetqua();

		int ncount = (int)vecItems.size();
		if (ncount == 0) return;

		HICON hLnkShortCut = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_SHORTCUT));
		HICON hLnkFolder = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_OPENFOLDER));
		HICON hLnkNull = LoadIcon(AfxGetResourceHandle(), MAKEINTRESOURCE(IDI_ICON_NULL));

		CString szval = L"";
		for (int i = 0; i < ncount; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, iGFLTenNew, vecItems[i].szName);
			m_list_view.SetItemText(i, iGFLTenOld, vecItems[i].szName);
			m_list_view.SetItemText(i, iGFLPath, vecItems[i].szFullpath);
			m_list_view.SetCheck(i, true);

			if (vecItems[i].szType != L"")
			{
				// Type file
				szval = vecItems[i].szType;
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

		for (int i = 0; i < ncount; i++)
		{
			// Load Image icon
			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, i, 0, 0, 0);
		}

		m_list_view.SetColumnWidth(iGFLPath, LVSCW_AUTOSIZE_USEHEADER);
	}
	catch (...) {}
}

void DlgAutoRename::_SetTotalKetqua()
{
	CString szval = L"";
	if (nIndexSplit == 0)
	{
		szval.Format(L"Tổng số: %d thư mục", nCountList);
	}
	else
	{
		szval.Format(L"Tổng số: %d file", nCountList);
		if (nCountList > 1) szval += L"s";
	}

	m_lbl_total.SetWindowText(szval);
}

void DlgAutoRename::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 40;
	*pResult = 0;
}


void DlgAutoRename::OnTimer(UINT_PTR nIDEvent)
{
	if (blDragging)
	{
		blDragging = false;
		UpdateData();
	}

	CDialogEx::OnTimer(nIDEvent);
}


void DlgAutoRename::OnBnClickedBtnApdung()
{
	iTabDefault = m_tab_main.GetCurSel();
	switch (iTabDefault)
	{
		case 0: _RunFuncSothutu();	break;
		case 1: _RunFuncAdd();		break;
		case 2: _RunFuncDelete();	break;
		case 3: _RunFuncReplace();	break;
		case 4: _RunFuncKhac();		break;
	}
}

void DlgAutoRename::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
}

void DlgAutoRename::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem == iGFLPath)
	{
		CString szval = m_list_view.GetItemText(nItem, iGFLPath);
		if (szval != L"")
		{
			ShellExecute(NULL, L"open", szval, NULL, NULL, SW_SHOWMAXIMIZED);
		}
	}
}

void DlgAutoRename::OnAutornApply()
{
	iTypeApply = 1;

	OnBnClickedBtnApdung();

	iTypeApply = 0;
}


void DlgAutoRename::OnAutornOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}


void DlgAutoRename::OnAutornOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}

void DlgAutoRename::OnAutornUp()
{
	try
	{
		if (nItem > 0)
		{
			UpdateData();

			CString szval = L"";
			int iArr[] = { iGFLTenNew, iGFLTenOld, iGFLPath };
			int ncount = getSizeArray(iArr);

			for (int i = 0; i < ncount; i++)
			{
				szval = m_list_view.GetItemText(nItem, iArr[i]);
				m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem - 1, iArr[i]));
				m_list_view.SetItemText(nItem - 1, iArr[i], szval);
			}

			m_list_view.SetItem(nItem, 0, LVIF_IMAGE, NULL, nItem + 1, 0, 0, 0);
			m_list_view.SetItem(nItem - 1, 0, LVIF_IMAGE, NULL, nItem, 0, 0, 0);

			m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
			m_list_view.SetItemState(nItem - 1, LVIS_SELECTED, LVIS_SELECTED);

			nItem--;
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnAutornDown()
{
	try
	{
		if (nItem < nCountList - 1)
		{
			UpdateData();

			CString szval = L"";
			int iArr[] = { iGFLTenNew, iGFLTenOld, iGFLPath };
			int nsize = getSizeArray(iArr);

			for (int i = 0; i < nsize; i++)
			{
				szval = m_list_view.GetItemText(nItem, iArr[i]);
				m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem + 1, iArr[i]));
				m_list_view.SetItemText(nItem + 1, iArr[i], szval);
			}

			m_list_view.SetItem(nItem + 1, 0, LVIF_IMAGE, NULL, nItem + 2, 0, 0, 0);
			m_list_view.SetItem(nItem, 0, LVIF_IMAGE, NULL, nItem + 1, 0, 0, 0);

			m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
			m_list_view.SetItemState(nItem + 1, LVIS_SELECTED, LVIS_SELECTED);

			nItem++;
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnAutornReverse()
{
	POSITION pos = m_list_view.GetFirstSelectedItemPosition();
	if (pos != NULL)
	{
		int dem = 0;
		int iItems[2] = { 0, 0 };
		for (int i = 0; i < nCountList; i++)
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

		CString szval = L"";
		int iArr[] = { iGFLTenNew, iGFLTenOld, iGFLPath };
		int nsize = getSizeArray(iArr);

		for (int i = 0; i < nsize; i++)
		{
			szval = m_list_view.GetItemText(iItems[0], iArr[i]);
			m_list_view.SetItemText(iItems[0], iArr[i], m_list_view.GetItemText(iItems[1], iArr[i]));
			m_list_view.SetItemText(iItems[1], iArr[i], szval);
		}

		m_list_view.SetItem(iItems[0], 0, LVIF_IMAGE, NULL, iItems[0] + 1, 0, 0, 0);
		m_list_view.SetItem(iItems[1], 0, LVIF_IMAGE, NULL, iItems[1] + 1, 0, 0, 0);

		m_list_view.SetItemState(iItems[0], 0, LVIS_SELECTED);
		m_list_view.SetItemState(iItems[1], LVIS_SELECTED, LVIS_SELECTED);
	}
}

void DlgAutoRename::OnAutornSel()
{
	_GetItemSelect(1);
}

void DlgAutoRename::OnAutornUnsel()
{
	_GetItemSelect(0);
}

void DlgAutoRename::OnAutornSelall()
{
	_SelectAllItems();
}

void DlgAutoRename::OnAutornChfiles()
{
	if (nIndexSplit != 0) return;

	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"")
		{
			m_txt_browse.SetWindowText(szpath);
			OnAutornFiles();
		}
	}
}

void DlgAutoRename::OnAutornChfolder()
{
	if (nIndexSplit != 0) return;

	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"")
		{
			m_txt_browse.SetWindowText(szpath);
			OnAutornFolder();
		}
	}
}

void DlgAutoRename::OnBnClickedCheckAll()
{
	iKeyESC = 0;
	CButton *m_checkAll = (CButton*)GetDlgItem(AUTORM_CHECK_ALL);
	int nCheck = m_checkAll->GetCheck();

	for (int i = 0; i < nCountList; i++)
	{
		m_list_view.SetCheck(i, nCheck);
	}
}

void DlgAutoRename::_GetItemSelect(int icheck)
{
	try
	{
		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int dem = 0;
			for (int i = 0; i < nCountList; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					m_list_view.SetCheck(i, icheck);
					dem++;
				}
			}
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnBnClickedBtnReset()
{
	int result = ff->_y(L"Bạn có chắc chắn khôi phục lại các tên gốc?", 0, 0, 0);
	if (result == 6)
	{
		vecList.clear();

		CString szval = L"";
		vector<CString> vecArr;
		for (int i = 0; i < nCountList; i++)
		{
			szval = m_list_view.GetItemText(i, iGFLTenOld);
			vecArr.push_back(szval);

			m_list_view.SetItemText(i, iGFLTenNew, szval);
		}

		vecList.push_back(vecArr);

		nIndexVec = 1;
	}	
}

void DlgAutoRename::OnBnClickedBtnUndo()
{
	try
	{
		int nsize = (int)vecList.size();
		if (nsize > 0)
		{
			if (nIndexVec > 0)
			{
				nIndexVec--;

				for (int i = 0; i < nCountList; i++)
				{
					m_list_view.SetItemText(i, iGFLTenNew, vecList[nIndexVec][i]);
				}
			}
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnBnClickedBtnRedo()
{
	try
	{
		int nsize = (int)vecList.size();
		if (nsize > 0)
		{
			if (nIndexVec < nsize - 1)
			{
				nIndexVec++;

				for (int i = 0; i < nCountList; i++)
				{
					m_list_view.SetItemText(i, iGFLTenNew, vecList[nIndexVec][i]);
				}
			}
		}
	}
	catch (...) {}
}

void DlgAutoRename::_RunFuncSothutu()
{
	try
	{
		CString strNum = L"";
		CString szTiento = L"", szHauto = L"";

		m_frmSTT->UpdateData();
		m_frmSTT->m_txt_num.GetWindowTextW(strNum);
		m_frmSTT->m_txt_tiento.GetWindowTextW(szTiento);
		m_frmSTT->m_txt_hauto.GetWindowTextW(szHauto);

		int iChkZero = m_frmSTT->m_check_zero.GetCheck();
		int iChkEnd = m_frmSTT->m_check_end.GetCheck();
		int iChkRep = m_frmSTT->m_check_replace.GetCheck();

		strNum.Trim();
		int iStart = _wtoi(strNum);

		vector<CString> vecBk, vecArr;
		CString szval = L"", name = L"";

		for (int i = 0; i < nCountList; i++)
		{
			name = m_list_view.GetItemText(i, iGFLTenNew);
			vecBk.push_back(name);
			szval = name;

			if ((iTypeApply == 0 && (int)m_list_view.GetCheck(i) == 1)
				|| (iTypeApply == 1 && m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED))
			{
				// Kiểm tra có điền thêm giá trị '0'
				if (iChkZero == 0)
				{
					szval.Format(L"%s%d%s", szTiento, iStart, szHauto);
				}
				else
				{
					if (nCountList < 100)
					{
						szval.Format(L"%s%02d%s", szTiento, iStart, szHauto);
					}
					else if (nCountList >= 100 && nCountList < 1000)
					{
						szval.Format(L"%s%03d%s", szTiento, iStart, szHauto);
					}
					else
					{
						szval.Format(L"%s%04d%s", szTiento, iStart, szHauto);
					}
				}

				// Kiểm tra có thay thế hoàn toàn không?
				if (iChkRep == 0)
				{
					// Kiểm tra thêm vào đầu hay cuối
					if (iChkEnd == 0) szval += name;
					else szval = name + szval;
				}

				m_list_view.SetItemText(i, iGFLTenNew, szval);
				name = szval;

				iStart++;
			}

			vecArr.push_back(name);
		}

		// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
		if (nIndexVec < (int)vecList.size() - 1)
		{
			vecList.push_back(vecBk);
		}
		
		vecList.push_back(vecArr);

		nIndexVec = (int)vecList.size() - 1;

		vecArr.clear();
		vecBk.clear();
	}
	catch (...) {}
}

void DlgAutoRename::_RunFuncAdd()
{
	try
	{
		int iRad = 0;
		int ivtAdd = 0;
		CString szChuoi = L"";

		m_frmAdd->UpdateData();
		m_frmAdd->m_txt_chuoi.GetWindowTextW(szChuoi);

		if (szChuoi == L"") return;

		if ((int)m_frmAdd->m_rad_begin.GetCheck() == 1) iRad = 0;
		else if ((int)m_frmAdd->m_rad_end.GetCheck() == 1) iRad = 1;
		else if ((int)m_frmAdd->m_rad_mid.GetCheck() == 1) iRad = 2;

		if (iRad == 2)
		{
			CString strNum = L"";
			m_frmAdd->m_txt_num.GetWindowTextW(strNum);

			strNum.Trim();
			ivtAdd = _wtoi(strNum);
		}

		vector<CString> vecBk, vecArr;
		CString szval = L"", name = L"";
		int nLen = 0;

		for (int i = 0; i < nCountList; i++)
		{
			name = m_list_view.GetItemText(i, iGFLTenNew);
			vecBk.push_back(name);
			szval = name;

			if ((iTypeApply == 0 && (int)m_list_view.GetCheck(i) == 1)
				|| (iTypeApply == 1 && m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED))
			{
				if (iRad == 0) szval = szChuoi + name;
				else if (iRad == 1) szval = name + szChuoi;
				else if (iRad == 2)
				{
					if (ivtAdd > 0)
					{
						nLen = name.GetLength();
						if (ivtAdd > nLen) ivtAdd = nLen;

						szval = name.Left(ivtAdd) + szChuoi + name.Right(nLen - ivtAdd);
					}
				}

				m_list_view.SetItemText(i, iGFLTenNew, szval);
				name = szval;
			}

			vecArr.push_back(name);
		}

		// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
		if (nIndexVec < (int)vecList.size() - 1)
		{
			vecList.push_back(vecBk);
		}

		vecList.push_back(vecArr);

		nIndexVec = (int)vecList.size() - 1;

		vecArr.clear();
		vecBk.clear();
	}
	catch (...) {}
}


void DlgAutoRename::_RunFuncDelete()
{
	try
	{
		int iRad = 0;
		int iChild = 0;
		CString szChuoi = L"";

		m_frmDel->UpdateData();
		m_frmDel->m_txt_chuoi.GetWindowTextW(szChuoi);

		if ((int)m_frmDel->m_rad_trim.GetCheck() == 1) iRad = 0;
		else if ((int)m_frmDel->m_rad_space.GetCheck() == 1) iRad = 1;
		else if ((int)m_frmDel->m_rad_chuoi.GetCheck() == 1) iRad = 2;

		if (iRad == 2)
		{
			if (szChuoi == L"") return;

			if ((int)m_frmDel->m_rad_all.GetCheck() == 1) iChild = 0;
			else if ((int)m_frmDel->m_rad_begin.GetCheck() == 1) iChild = 1;
			else if ((int)m_frmDel->m_rad_end.GetCheck() == 1) iChild = 2;
			else if ((int)m_frmDel->m_rad_left.GetCheck() == 1) iChild = 3;
			else if ((int)m_frmDel->m_rad_right.GetCheck() == 1) iChild = 4;
		}

		vector<CString> vecBk, vecArr;
		CString szval = L"", name = L"";
		int nLen = 0, nTotal = 0, pos = -1;

		for (int i = 0; i < nCountList; i++)
		{
			name = m_list_view.GetItemText(i, iGFLTenNew);
			vecBk.push_back(name);
			szval = name;

			if ((iTypeApply == 0 && (int)m_list_view.GetCheck(i) == 1)
				|| (iTypeApply == 1 && m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED))
			{
				if (iRad == 0) szval.Trim();
				else if (iRad == 1)
				{
					while (true)
					{
						if (szval.Find(L"  ") == -1) break;
						szval.Replace(L"  ", L" ");
					}					
				}
				else if (iRad == 2)
				{
					nLen = szChuoi.GetLength();
					nTotal = szval.GetLength();

					switch (iChild)
					{
						case 0:
						{
							szval.Replace(szChuoi, L"");
						}
						break;

						case 1:
						{
							if (nTotal >= nLen)
							{
								if (szval.Left(nLen) == szChuoi)
								{
									szval = szval.Right(nTotal - nLen);
								}
							}
						}
						break;

						case 2:
						{
							if (nTotal >= nLen)
							{
								if (szval.Right(nLen) == szChuoi)
								{
									szval = szval.Left(nTotal - nLen);
								}
							}
						}
						break;

						case 3:
						{
							pos = szval.Find(szChuoi);
							if (pos > 0) szval = szval.Right(nTotal - pos - 1);
						}
						break;

						case 4:
						{
							pos = -1;
							for (int i = nTotal - nLen; i >= 0; i--)
							{
								if (szval.Mid(i, nLen) == szChuoi)
								{
									pos = i;
									break;
								}
							}

							if (pos >= 0)
							{
								pos += nLen;
								if (pos < nTotal - 1) szval = szval.Left(pos - 1);
							}
						}
						break;
					}
				}

				m_list_view.SetItemText(i, iGFLTenNew, szval);
				name = szval;
			}

			vecArr.push_back(name);
		}

		// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
		if (nIndexVec < (int)vecList.size() - 1)
		{
			vecList.push_back(vecBk);
		}

		vecList.push_back(vecArr);

		nIndexVec = (int)vecList.size() - 1;

		vecArr.clear();
		vecBk.clear();
	}
	catch (...) {}
}


void DlgAutoRename::_RunFuncReplace()
{
	try
	{
		CString szRep = L"", szChuoi = L"";

		m_frmRep->UpdateData();
		m_frmRep->m_txt_rep.GetWindowTextW(szRep);
		m_frmRep->m_txt_chuoi.GetWindowTextW(szChuoi);

		if (szRep == L"") return;

		int iChkCase = (int)m_frmRep->m_chk_case.GetCheck();
		int iChkKdau = (int)m_frmRep->m_chk_kodau.GetCheck();

		vector<CString> vecBk, vecArr;
		CString szval = L"", name = L"";

		for (int i = 0; i < nCountList; i++)
		{
			name = m_list_view.GetItemText(i, iGFLTenNew);
			vecBk.push_back(name);
			szval = name;

			if ((iTypeApply == 0 && (int)m_list_view.GetCheck(i) == 1)
				|| (iTypeApply == 1 && m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED))
			{
				szval = ff->_ReplaceMatchCase(szval, szRep, szChuoi, iChkCase, iChkKdau);
				m_list_view.SetItemText(i, iGFLTenNew, szval);
				name = szval;
			}

			vecArr.push_back(name);
		}

		// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
		if (nIndexVec < (int)vecList.size() - 1)
		{
			vecList.push_back(vecBk);
		}

		vecList.push_back(vecArr);

		nIndexVec = (int)vecList.size() - 1;

		vecArr.clear();
		vecBk.clear();
	}
	catch (...) {}
}


void DlgAutoRename::_RunFuncKhac()
{
	try
	{
		CString szRep = L"", szChuoi = L"";

		m_frmKhac->UpdateData();

		int nRadio = 0;
		if ((int)m_frmKhac->m_rad_upper.GetCheck() == 1) nRadio = 0;
		else if ((int)m_frmKhac->m_rad_lower.GetCheck() == 1) nRadio = 1;
		else if ((int)m_frmKhac->m_rad_unsigned.GetCheck() == 1) nRadio = 2;
		else if ((int)m_frmKhac->m_rad_proper.GetCheck() == 1) nRadio = 3;
		else if ((int)m_frmKhac->m_rad_kodau.GetCheck() == 1) nRadio = 4;

		vector<CString> vecBk, vecArr;
		CString szval = L"", name = L"";
		int nLen = 0;

		for (int i = 0; i < nCountList; i++)
		{
			name = m_list_view.GetItemText(i, iGFLTenNew);
			vecBk.push_back(name);
			szval = name;

			if ((iTypeApply == 0 && (int)m_list_view.GetCheck(i) == 1)
				|| (iTypeApply == 1 && m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED))
			{
				switch (nRadio)
				{
					case 0:
					{
						szval = ff->_UpperUnicode(szval);
					}
					break;

					case 1:
					{
						szval = ff->_LowerUnicode(szval);
					}
					break;

					case 2:
					{
						szval.Trim();
						nLen = szval.GetLength();

						CString szU2 = szval.Left(1);
						szU2 = ff->_UpperUnicode(szU2);

						CString szL2 = L"";
						if (nLen > 1)
						{
							szL2 = szval.Right(nLen - 1);
							szL2 = ff->_LowerUnicode(szL2);
						}
						szval = szU2 + szL2;
					}
					break;

					case 3:
					{
						vector<CString> vecKey;
						int ncount = ff->_TackTukhoa(vecKey, szval, L".", L".", 0, 1);
						if (ncount > 0)
						{
							szval = L"";
							CString szU3 = L"", szL3 = L"";
							for (int i = 0; i < ncount; i++)
							{
								vecKey[i].Trim();
								nLen = vecKey[i].GetLength();

								szU3 = vecKey[i].Left(1);
								szU3 = ff->_UpperUnicode(szU3);

								szL3 = L"";
								if (nLen > 1)
								{
									szL3 = vecKey[i].Right(nLen - 1);
									szL3 = ff->_LowerUnicode(szL3);
								}

								szval += (szU3 + szL3);
								if (i < ncount - 1) szval += L".";
							}
						}

						vecKey.clear();
					}
					break;

					case 4:
					{
						szval = ff->_ConvertKhongDau(szval);
					}
					break;
				}

				m_list_view.SetItemText(i, iGFLTenNew, szval);
				name = szval;
			}

			vecArr.push_back(name);
		}

		// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
		if (nIndexVec < (int)vecList.size() - 1)
		{
			vecList.push_back(vecBk);
		}

		vecList.push_back(vecArr);

		nIndexVec = (int)vecList.size() - 1;

		vecArr.clear();
		vecBk.clear();
	}
	catch (...) {}
}

void DlgAutoRename::OnEnSetfocusEditList()
{
	iKeyESC = 1;
}

void DlgAutoRename::OnEnKillfocusEditList()
{
	try
	{
		szBefore = m_list_view.GetItemText(CLRow, CLColumn);
		m_edit_autorename.GetWindowTextW(szAfter);

		szAfter.Trim(); szBefore.Trim();

		if (szAfter == szBefore) return;

		if (CLColumn == iGFLTenNew)
		{
			vector<CString> vecBk, vecArr;

			CString szval = L"";
			for (int i = 0; i < nCountList; i++)
			{
				szval = m_list_view.GetItemText(i, iGFLTenNew);

				if (i != CLRow)
				{
					vecBk.push_back(szval);
					vecArr.push_back(szval);
				}
				else
				{
					vecBk.push_back(szBefore);
					vecArr.push_back(szAfter);
				}				
			}

			// Trường hợp vị trí Undo không khớp với vị trí cuối cùng
			if (nIndexVec < (int)vecList.size() - 1)
			{
				vecList.push_back(vecBk);
			}

			vecList.push_back(vecArr);

			nIndexVec = (int)vecList.size() - 1;
		}
	}
	catch (...) {}
}

void DlgAutoRename::OnEnSetfocusTxtBrowse()
{
	iKeyESC = 0;
}

void DlgAutoRename::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
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
}

void DlgAutoRename::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename;
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

		nIndexSplit = 0;
		szval = reg->ReadString(L"SplitButton", L"");
		if (szval != L"")
		{
			nIndexSplit = _wtoi(szval);
			if (nIndexSplit != 1) nIndexSplit = 0;
		}

		/*iTabDefault = 0;
		szval = reg->ReadString(L"TabDefault", L"");
		if (szval != L"")
		{
			iTabDefault = _wtoi(szval);
			if (iTabDefault < 0 || iTabDefault > 4) iTabDefault = 0;
		}*/

		// Mặc định luôn để tab đầu tiên (iTabDefault = 0)
		iTabDefault = 0;
	}
	catch (...) {}
}

void DlgAutoRename::_LoadRegistryTab()
{
	try
	{

/*** Tab 1 - Số thứ tự ************************************************************************************/

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabNumber;
		reg->CreateKey(szCreate);

		m_frmSTT->m_txt_tiento.SetWindowText(reg->ReadString(L"Prefix", L""));
		m_frmSTT->m_txt_hauto.SetWindowText(reg->ReadString(L"Suffixes", L""));

		int num = _wtoi(reg->ReadString(L"NumStart", L""));
		if (num < 0 || num > 10000) num = 1;
		m_frmSTT->m_spin_num.SetPos(num);

		num = _wtoi(reg->ReadString(L"CheckZero", L""));
		if (num !=0) num = 1;
		m_frmSTT->m_check_zero.SetCheck(num);

		num= _wtoi(reg->ReadString(L"CheckEnd", L""));
		if (num!= 0) num= 1;
		m_frmSTT->m_check_end.SetCheck(num);

		num = _wtoi(reg->ReadString(L"CheckReplace", L""));
		if (num != 0) num = 1;
		m_frmSTT->m_check_replace.SetCheck(num);

/*** Tab 2 - Add ******************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabAdd;
		reg->CreateKey(szCreate);

		m_frmAdd->m_txt_chuoi.SetWindowText(reg->ReadString(L"NewString", L""));

		num = _wtoi(reg->ReadString(L"Radio", L""));
		if (num < 0 || num > 2) num = 0;
		
		switch (num)
		{
			case 0: 
			{
				m_frmAdd->_SetUnCheckRadio();
				m_frmAdd->m_rad_begin.SetCheck(1);

				m_frmAdd->_SetEnableNumber(false);
			}
			break;

			case 1:
			{
				m_frmAdd->_SetUnCheckRadio();
				m_frmAdd->m_rad_end.SetCheck(1);

				m_frmAdd->_SetEnableNumber(false);
			}
			break;

			case 2:
			{
				m_frmAdd->_SetUnCheckRadio();
				m_frmAdd->m_rad_mid.SetCheck(1);

				m_frmAdd->_SetEnableNumber(true);
			}
			break;
		}

		num = _wtoi(reg->ReadString(L"Position", L""));
		if (num < 1 || num > 10000) num = 1;
		m_frmAdd->m_spin_num.SetPos(num);

/*** Tab 3 - Delete ***************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabDelete;
		reg->CreateKey(szCreate);

		m_frmDel->m_txt_chuoi.SetWindowText(reg->ReadString(L"NewString", L""));

		num = _wtoi(reg->ReadString(L"Radio", L""));
		if (num < 0 || num > 2) num = 0;

		switch (num)
		{
			case 0:
			{
				m_frmDel->_SetUnCheckRadio();
				m_frmDel->m_rad_trim.SetCheck(1);
				
				m_frmDel->_SetEnableRadio(false);
			}
			break;

			case 1:
			{
				m_frmDel->_SetUnCheckRadio();
				m_frmDel->m_rad_space.SetCheck(1);

				m_frmDel->_SetEnableRadio(false);
			}
			break;

			case 2:
			{
				m_frmDel->_SetUnCheckRadio();
				m_frmDel->m_rad_chuoi.SetCheck(1);

				m_frmDel->_SetEnableRadio(true);
			}
			break;
		}

		num = _wtoi(reg->ReadString(L"RadioChild", L""));
		if (num < 0 || num > 4) num = 0;

		switch (num)
		{
			case 0:
			{
				m_frmDel->_SetUnCheckChildRadio();
				m_frmDel->m_rad_all.SetCheck(1);
			}
			break;

			case 1:
			{
				m_frmDel->_SetUnCheckChildRadio();
				m_frmDel->m_rad_begin.SetCheck(1);
			}
			break;

			case 2:
			{
				m_frmDel->_SetUnCheckChildRadio();
				m_frmDel->m_rad_end.SetCheck(1);
			}
			break;

			case 3:
			{
				m_frmDel->_SetUnCheckChildRadio();
				m_frmDel->m_rad_left.SetCheck(1);
			}
			break;

			case 4:
			{
				m_frmDel->_SetUnCheckChildRadio();
				m_frmDel->m_rad_right.SetCheck(1);
			}
			break;
		}

/*** Tab 4 - Replace **************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabReplace;
		reg->CreateKey(szCreate);

		m_frmRep->m_txt_rep.SetWindowText(reg->ReadString(L"OldString", L""));
		m_frmRep->m_txt_chuoi.SetWindowText(reg->ReadString(L"NewString", L""));

		num = _wtoi(reg->ReadString(L"MatchCase", L""));
		if (num != 0) num = 1;
		m_frmRep->m_chk_case.SetCheck(num);

		num = _wtoi(reg->ReadString(L"Unsigned", L""));
		if (num != 0) num = 1;
		m_frmRep->m_chk_kodau.SetCheck(num);

/*** Tab 5 - Khác *****************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabOther;
		reg->CreateKey(szCreate);

		num = _wtoi(reg->ReadString(L"Radio", L""));
		if (num < 0 || num > 4) num = 0;

		switch (num)
		{
			case 0:
			{
				m_frmKhac->_SetUnCheckRadio();
				m_frmKhac->m_rad_upper.SetCheck(1);
			}
			break;

			case 1:
			{
				m_frmKhac->_SetUnCheckRadio();
				m_frmKhac->m_rad_lower.SetCheck(1);
			}
			break;

			case 2:
			{
				m_frmKhac->_SetUnCheckRadio();
				m_frmKhac->m_rad_unsigned.SetCheck(1);
			}
			break;

			case 3:
			{
				m_frmKhac->_SetUnCheckRadio();
				m_frmKhac->m_rad_proper.SetCheck(1);
			}
			break;

			case 4:
			{
				m_frmKhac->_SetUnCheckRadio();
				m_frmKhac->m_rad_kodau.SetCheck(1);
			}
			break;
		}
	}
	catch (...) {}
}

void DlgAutoRename::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename;
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

		m_txt_browse.GetWindowTextW(szTargetDirectory);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgDirectory),
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szTargetDirectory, 0)));

		szval.Format(L"%d", nIndexSplit);
		reg->WriteChar(ff->_ConvertCstringToChars(L"SplitButton"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", iTabDefault);
		reg->WriteChar(ff->_ConvertCstringToChars(L"TabDefault"), ff->_ConvertCstringToChars(szval));
	
/*** Tab 1 - Số thứ tự ************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabNumber;
		reg->CreateKey(szCreate);

		m_frmSTT->m_txt_num.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"NumStart"), ff->_ConvertCstringToChars(szval));

		m_frmSTT->m_txt_tiento.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Prefix"), ff->_ConvertCstringToChars(szval));

		m_frmSTT->m_txt_hauto.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Suffixes"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_frmSTT->m_check_zero.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckZero"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_frmSTT->m_check_end.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckEnd"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_frmSTT->m_check_replace.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"CheckReplace"), ff->_ConvertCstringToChars(szval));

/*** Tab 2 - Add ******************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabAdd;
		reg->CreateKey(szCreate);

		m_frmAdd->m_txt_chuoi.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"NewString"), ff->_ConvertCstringToChars(szval));

		int iRadAdd = 0;
		if ((int)m_frmAdd->m_rad_begin.GetCheck() == 1) iRadAdd = 0;
		else if ((int)m_frmAdd->m_rad_end.GetCheck() == 1) iRadAdd = 1;
		else if ((int)m_frmAdd->m_rad_mid.GetCheck() == 1) iRadAdd = 2;

		szval.Format(L"%d", iRadAdd);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Radio"), ff->_ConvertCstringToChars(szval));

		CString strVitri = L"";
		m_frmAdd->m_txt_num.GetWindowTextW(strVitri);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Position"), ff->_ConvertCstringToChars(strVitri));

/*** Tab 3 - Delete ***************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabDelete;
		reg->CreateKey(szCreate);

		m_frmDel->m_txt_chuoi.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"NewString"), ff->_ConvertCstringToChars(szval));

		int iRadDel = 0;
		if ((int)m_frmDel->m_rad_trim.GetCheck() == 1) iRadDel = 0;
		else if ((int)m_frmDel->m_rad_space.GetCheck() == 1) iRadDel = 1;
		else if ((int)m_frmDel->m_rad_chuoi.GetCheck() == 1) iRadDel = 2;

		szval.Format(L"%d", iRadDel);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Radio"), ff->_ConvertCstringToChars(szval));

		int iRadChild = 0;
		if ((int)m_frmDel->m_rad_all.GetCheck() == 1) iRadChild = 0;
		else if ((int)m_frmDel->m_rad_begin.GetCheck() == 1) iRadChild = 1;
		else if ((int)m_frmDel->m_rad_end.GetCheck() == 1) iRadChild = 2;
		else if ((int)m_frmDel->m_rad_left.GetCheck() == 1) iRadChild = 3;
		else if ((int)m_frmDel->m_rad_right.GetCheck() == 1) iRadChild = 4;

		szval.Format(L"%d", iRadChild);
		reg->WriteChar(ff->_ConvertCstringToChars(L"RadioChild"), ff->_ConvertCstringToChars(szval));

/*** Tab 4 - Replace **************************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabReplace;
		reg->CreateKey(szCreate);

		m_frmRep->m_txt_rep.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"OldString"), ff->_ConvertCstringToChars(szval));

		m_frmRep->m_txt_chuoi.GetWindowTextW(szval);
		reg->WriteChar(ff->_ConvertCstringToChars(L"NewString"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_frmRep->m_chk_case.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"MatchCase"), ff->_ConvertCstringToChars(szval));

		szval.Format(L"%d", (int)m_frmRep->m_chk_kodau.GetCheck());
		reg->WriteChar(ff->_ConvertCstringToChars(L"Unsigned"), ff->_ConvertCstringToChars(szval));

/*** Tab 5 - Khác *****************************************************************************************/
	
		szCreate = bb->BaseDecodeEx(CreateKeySettings) + AutoRename + TabOther;
		reg->CreateKey(szCreate);

		int iRadKhac = 0;
		if ((int)m_frmKhac->m_rad_upper.GetCheck() == 1) iRadKhac = 0;
		else if ((int)m_frmKhac->m_rad_lower.GetCheck() == 1) iRadKhac = 1;
		else if ((int)m_frmKhac->m_rad_unsigned.GetCheck() == 1) iRadKhac = 2;
		else if ((int)m_frmKhac->m_rad_proper.GetCheck() == 1) iRadKhac = 3;
		else if ((int)m_frmKhac->m_rad_kodau.GetCheck() == 1) iRadKhac = 4;
		
		szval.Format(L"%d", iRadAdd);
		reg->WriteChar(ff->_ConvertCstringToChars(L"Radio"), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgAutoRename::_SetLocationAndSize()
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
