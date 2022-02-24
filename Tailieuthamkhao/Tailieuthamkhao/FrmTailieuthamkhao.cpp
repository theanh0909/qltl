#include "stdafx.h"
#include "Base64.h"
#include "Tailieuthamkhao.h"
#include "FrmCheckRobot.h"
#include "FrmTailieuthamkhao.h"

#define ALLCOMBOBOX					L"- Tất cả phân loại -"

#define IDC_COLUMN_POPUP			3000
#define IDC_COLUMN_KEY_POPUP		3001
#define IDC_COLUMN_TEN_POPUP		3002
#define IDC_COLUMN_PLOAI_POPUP		3003
#define IDC_COLUMN_MOTA_POPUP		3004
#define IDC_COLUMN_TYPE_POPUP		3005
#define IDC_COLUMN_LOCAL_POPUP		3006

static const int CONST_TIMER_TAILIEUTHAMKHAO = 1;

// FrmTailieuthamkhao dialog
IMPLEMENT_DYNAMIC(FrmTailieuthamkhao, CDialogEx)

FrmTailieuthamkhao::FrmTailieuthamkhao(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DLG_TAILIEUTHAMKHAO, pParent)
{
	jUserID = -1;
	nItem = -1, nSubItem = -1;
	iTopdownload = 0, iMaxdownload = 10;

	demTimer = 0;
	blChangeKey = false;

	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 50;		

	clmKey = 1, clmIcon =2, clmTen = 3, clmPloai = 4;
	clmMota = 5, clmTrangthai = 6, clmDuongdan = 7, clmID = 8;

	szCopyMulti = L"";
	spathSaveDownload = L"";
	spathLocal = L"", szpathDownloadTL = L"";
	szFileCategory = L"", szFolderCategory = L"", szContentCaterory = L"";
	szpathListdown = L"";
}

FrmTailieuthamkhao::~FrmTailieuthamkhao()
{
	vecDSDulieu.clear();
	vecDSKetqua.clear();
	vecCategory.clear();
	vecListDownload.clear();
}

void FrmTailieuthamkhao::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, DWTL_TXT_KEY, m_txtTimkiem);
	DDX_Control(pDX, DWTL_BTN_OPENFD, m_btn_open);
	DDX_Control(pDX, DWTL_BTN_SAVE, m_btn_browse);
	DDX_Control(pDX, DWTL_BTN_CLOSE, m_btn_close);	
	DDX_Control(pDX, DWTL_CHECK_FOLDER, m_checkFolder);
	DDX_Control(pDX, DWTL_CHECK_COPY, m_checkCopy);
	DDX_Control(pDX, DWTL_CBB_DULIEU, m_cbbDulieu);
	DDX_Control(pDX, DWTL_CBB_PLOAI, m_cbbPhanloai);
	DDX_Control(pDX, DWTL_LIST_KQUA, m_listdsdownload);
	DDX_Control(pDX, DWTL_TXT_BROWSE, m_txtPathBrowse);
	DDX_Control(pDX, DWTL_BTN_DOWNLOAD, m_btn_download);	
}

BEGIN_MESSAGE_MAP(FrmTailieuthamkhao, CDialogEx)
	ON_WM_TIMER()
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()	
	
	ON_BN_CLICKED(DWTL_BTN_DOWNLOAD, &FrmTailieuthamkhao::OnBnClickedBtnDownload)	
	ON_BN_CLICKED(DWTL_BTN_OPENFD, &FrmTailieuthamkhao::OnBnClickedBtnOpenfd)
	ON_BN_CLICKED(DWTL_BTN_CLOSE, &FrmTailieuthamkhao::OnBnClickedBtnClose)	
	ON_BN_CLICKED(DWTL_CHECK_FOLDER, &FrmTailieuthamkhao::OnBnClickedCheckFolder)
	ON_BN_CLICKED(DWTL_CHECK_COPY, &FrmTailieuthamkhao::OnBnClickedCheckCopy)
	ON_BN_CLICKED(DWTL_BTN_SAVE, &FrmTailieuthamkhao::OnBnClickedBtnSave)

	ON_CBN_SELCHANGE(DWTL_CBB_DULIEU, &FrmTailieuthamkhao::OnCbnSelchangeCbbDulieu)
	ON_CBN_SELCHANGE(DWTL_CBB_PLOAI, &FrmTailieuthamkhao::OnCbnSelchangeCbbPloai)

	ON_NOTIFY(NM_CLICK, DWTL_LIST_KQUA, &FrmTailieuthamkhao::OnNMClickListKqua)
	ON_NOTIFY(NM_DBLCLK, DWTL_LIST_KQUA, &FrmTailieuthamkhao::OnNMDblclkListKqua)
	ON_NOTIFY(NM_RCLICK, DWTL_LIST_KQUA, &FrmTailieuthamkhao::OnNMRClickListKqua)
	ON_NOTIFY(LVN_ENDSCROLL, DWTL_LIST_KQUA, &FrmTailieuthamkhao::OnLvnEndScrollListKqua)
	ON_NOTIFY(HDN_ENDTRACK, 0, &FrmTailieuthamkhao::OnHdnEndtrackListKqua)

	ON_COMMAND(ID_DW_DOWNLOAD, &FrmTailieuthamkhao::OnDwDownload)
	ON_COMMAND(ID_DW_DWANDVIEW, &FrmTailieuthamkhao::OnDwDwandview)
	ON_COMMAND(ID_DW_OPENFILE, &FrmTailieuthamkhao::OnDwOpenfile)
	ON_COMMAND(ID_DW_OPENFOLDER, &FrmTailieuthamkhao::OnDwOpenfolder)
	ON_COMMAND(ID_DW_COPYND, &FrmTailieuthamkhao::OnDwCopyNoidung)

	ON_COMMAND(IDC_COLUMN_POPUP, &FrmTailieuthamkhao::OnDwShowAll)
	ON_COMMAND(IDC_COLUMN_KEY_POPUP, &FrmTailieuthamkhao::OnDwShowKey)
	ON_COMMAND(IDC_COLUMN_TEN_POPUP, &FrmTailieuthamkhao::OnDwShowNoidung)
	ON_COMMAND(IDC_COLUMN_PLOAI_POPUP, &FrmTailieuthamkhao::OnDwShowPhanloai)
	ON_COMMAND(IDC_COLUMN_MOTA_POPUP, &FrmTailieuthamkhao::OnDwShowMota)
	ON_COMMAND(IDC_COLUMN_TYPE_POPUP, &FrmTailieuthamkhao::OnDwShowTrangthai)
	ON_COMMAND(IDC_COLUMN_LOCAL_POPUP, &FrmTailieuthamkhao::OnDwShowPath)

	ON_MESSAGE(WM_LISTCTRLEX_HEADERRIGHTCLICK, OnHeaderRightClick)	
	ON_EN_CHANGE(DWTL_TXT_KEY, &FrmTailieuthamkhao::OnEnChangeTxtKey)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(FrmTailieuthamkhao)
	EASYSIZE(DWTL_CBB_DULIEU, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DWTL_CBB_PLOAI, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DWTL_TXT_KEY, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)	
	EASYSIZE(DWTL_LBL_1, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DWTL_LBL_2, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DWTL_LBL_3, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(DWTL_LIST_KQUA, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_BTN_OPENFD, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(DWTL_TXT_BROWSE, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_BTN_SAVE, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_CHECK_FOLDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_CHECK_COPY, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_BTN_DOWNLOAD, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(DWTL_BTN_CLOSE, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// FrmTailieuthamkhao message handlers
BOOL FrmTailieuthamkhao::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_DOWNLOAD));
	SetIcon(hIcon, FALSE);
	//SetBackgroundColor(RGBBlueOffice);

	jcolDownload[0] = 25 * jScaleLayout;
	jcolDownload[1] = 95 * jScaleLayout;
	jcolDownload[2] = 22 * jScaleLayout;
	jcolDownload[3] = 500 * jScaleLayout;
	jcolDownload[4] = 200 * jScaleLayout;
	jcolDownload[5] = 300 * jScaleLayout;
	jcolDownload[6] = 85 * jScaleLayout;
	jcolDownload[7] = 0;
	jcolDownload[8] = 0;

	// Xác định đường dẫn mặc định
	ff._SetPathDefault();

	// Base-code
	GetBaseDecodeEx();

	// Tạo thư mục lưu trữ dữ liệu
	CString szFolder = L"";
	szFolder.Format(L"%s%s", spathDefault, FOLDERSAVE);
	ff._CreateNewFolder(szFolder);

	szpathListdown.Format(L"%s\\%s", szFolder, FILELISTDOWNLOAD);
	szpathListdown.Replace(L"\\\\",L"\\");

	// Tạo tệp tin lưu trữ dữ liệu download
	if (ff._FileExists(szpathListdown, 0) != 1)
	{
		if (CreateFileListDownload(szpathListdown) != 1)
		{
			ff._msgbox(L"Khởi tạo thư viện thất bại.", 2);
			CDialogEx::OnCancel();
			return TRUE;
		}
	}

	// Thiết lập các giá trị mặc định
	SetDefault();
	StyleListdulieu();

	// Tài khoản sử dụng --> 0: Miễn phí | 1: Khóa mềm | 2: Khóa cứng
	jUserID = CheckLicense();

	xl->PutStatusBar((_bstr_t)L"Đang kiểm tra thư viện và phân loại dữ liệu...");

	// Đọc tất cả dữ liệu đã tải về
	GetAllListdownload();

	// Tải và đọc dữ liệu category
	DownloadFileCategory();
	LoadComboboxBodulieu();

	xl->PutStatusBar((_bstr_t)L"Đang load dữ liệu. Vui lòng đợi trong giây lát...");

	// Xác định số lượng dữ liệu
	GetAllDulieu();
	LocPhanloaidulieu();

	xl->PutStatusBar((_bstr_t)L"Ready");

	// Sử dụng icon trong menu-content
	mnIcon.GdiplusStartupInputConfig();

	INIT_EASYSIZE;

	// Xác định kích thước vị trí hộp thoại
	SetLocationAndSize();

	SetTimerSearch();
	blChangeKey = true;

	return TRUE;
}

BOOL FrmTailieuthamkhao::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(DWTL_TXT_KEY))
			{
				GotoDlgCtrl(GetDlgItem(DWTL_LIST_KQUA));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(DWTL_LIST_KQUA))
			{
				GotoDlgCtrl(GetDlgItem(DWTL_TXT_KEY));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			OnBnClickedBtnClose();
			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(DWTL_LIST_KQUA))
		{
			if (chr == 0x03)
			{
				// Paste dữ liệu Clipboard
				if (nItem >= 0 && nSubItem >= 0)
				{
					OnDwCopyNoidung();
					return TRUE;
				}
			}
		}
	}

	return FALSE;
}

void FrmTailieuthamkhao::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void FrmTailieuthamkhao::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(1106, 515, fwSide, pRect);
}

void FrmTailieuthamkhao::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnClose();
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTailieuthamkhao::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		/*if (nItem == -1 || nSubItem == -1) return;

		CMenu _menu;
		_menu.LoadMenu(IDR_MENU_DOWNLOADTL);

		CRect rectSubmitButton;
		CListCtrlEx *pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(DWTL_LIST_KQUA));
		pClist->GetWindowRect(&rectSubmitButton);

		CMenu *_pMenu = _menu.GetSubMenu(0);
		ASSERT(_pMenu);

		int dem = 0;
		long nRow = -1;
		int ncount = m_listdsdownload.GetItemCount();
		for (long i = 0; i < (long)m_listdsdownload.GetItemCount(); i++)
		{
			if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				dem++;
				if (nRow == -1) nRow = i;
				if (dem == 2) break;
			}
		}

		if (ncount == 0 || (ncount > 0 && nRow == -1) || dem == 2)
		{
			_pMenu->EnableMenuItem(ID_DW_DOWNLOAD, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(ID_DW_DWANDVIEW, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(ID_DW_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(ID_DW_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		if (ncount == 0 || (ncount > 0 && nRow == -1))
		{
			_pMenu->EnableMenuItem(ID_DW_COPYND, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		CString szpath = m_listdsdownload.GetItemText(nRow, clmDuongdan);
		if (szpath == L"")
		{
			_pMenu->EnableMenuItem(ID_DW_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			_pMenu->EnableMenuItem(ID_DW_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
		}

		if (rectSubmitButton.PtInRect(point))
		{
			_pMenu->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON, point.x, point.y, this);
		}*/

/******************************************************************************/

		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_DOWNLOADTL));
		HMENU pMenu = GetSubMenu(hMMenu, 0);

		if (pMenu)
		{
			HBITMAP hBmp;
			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_DW_DOWNLOAD, IDI_ICON_DOWN_W10);
			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_DW_DWANDVIEW, IDI_ICON_DOWN_VIEW);
			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_DW_OPENFILE, IDI_ICON_EDIT_FILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_DW_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, ID_DW_COPYND, IDI_ICON_COPY);

			CRect rectSubmitButton;
			CListCtrlEx *pClist = reinterpret_cast<CListCtrlEx *>(GetDlgItem(DWTL_LIST_KQUA));
			pClist->GetWindowRect(&rectSubmitButton);

			if (rectSubmitButton.PtInRect(point))
			{
				int dem = 0;
				long nRow = -1;
				int ncount = m_listdsdownload.GetItemCount();
				for (long i = 0; i < (long)m_listdsdownload.GetItemCount(); i++)
				{
					if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
					{
						dem++;
						if (nRow == -1) nRow = i;
						if (dem == 2) break;
					}
				}

				if (ncount == 0 || (ncount > 0 && nRow == -1) || dem == 2)
				{
					EnableMenuItem(pMenu, ID_DW_DOWNLOAD, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
					EnableMenuItem(pMenu, ID_DW_DWANDVIEW, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
					EnableMenuItem(pMenu, ID_DW_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
					EnableMenuItem(pMenu, ID_DW_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				}

				if (ncount == 0 || (ncount > 0 && nRow == -1))
				{
					EnableMenuItem(pMenu, ID_DW_COPYND, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				}

				CString szpath = m_listdsdownload.GetItemText(nRow, clmDuongdan);
				if (szpath == L"")
				{
					EnableMenuItem(pMenu, ID_DW_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
					EnableMenuItem(pMenu, ID_DW_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				}

				TrackPopupMenu(pMenu, TPM_LEFTALIGN | TPM_TOPALIGN , point.x, point.y, 0, this->GetSafeHwnd(), NULL);
			}

			if (hBmp) ::DeleteObject(hBmp);
		}

		if (hMMenu) VERIFY(DestroyMenu(hMMenu));
	}
	catch (...) {}
}

void FrmTailieuthamkhao::OnBnClickedBtnDownload()
{
	RightClickDownload(0);
}

void FrmTailieuthamkhao::OnCbnSelchangeCbbDulieu()
{
	int nIndex = m_cbbDulieu.GetCurSel();
	if (nIndex >= 0)
	{
		szFileCategory = vecCategory[nIndex].filename;
		szFolderCategory = vecCategory[nIndex].szfolder;
		szContentCaterory = vecCategory[nIndex].szcontent;

		// Xác định lại số lượng dữ liệu
		GetAllDulieu();
		LocPhanloaidulieu();
		Timkiemdulieu();
	}
}

void FrmTailieuthamkhao::OnCbnSelchangeCbbPloai()
{
	Timkiemdulieu();
}

void FrmTailieuthamkhao::Timkiemdulieu()
{
	CString szKeyTK = L"";
	m_txtTimkiem.GetWindowTextW(szKeyTK);
	szKeyTK.Trim();

	// Xác định giá trị combobox
	CString szPloai = L"";
	int nIndex = m_cbbPhanloai.GetCurSel();
	if (nIndex > 0) m_cbbPhanloai.GetLBText(nIndex, szPloai);

	int nTotal = LoadDulieusudung(szKeyTK, szPloai);
}

void FrmTailieuthamkhao::OnBnClickedBtnOpenfd()
{
	try
	{
		m_txtPathBrowse.GetWindowTextW(spathLocal);
		if (spathLocal != L"")
		{
			ff._CreateNewFolder(spathLocal);
			ShellExecuteW(NULL, L"open", spathLocal, NULL, NULL, SW_SHOWMAXIMIZED);
		}
	}
	catch (...) {}
}

void FrmTailieuthamkhao::OnBnClickedBtnSave()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần lưu...";
	if (m_dlg.DoModal() == IDOK)
	{
		CString strTargetDirectory = m_dlg.GetPathName();
		m_txtPathBrowse.SetWindowText(strTargetDirectory);
	}
}

void FrmTailieuthamkhao::OnBnClickedBtnClose()
{
	// Đường dẫn lưu trữ
	m_txtPathBrowse.GetWindowTextW(spathLocal);
	if (spathLocal != L"")
	{
		spathSaveDownload = spathLocal;
		ff._WriteFileText(spathSaveDownload, szpathDownloadTL, 0);
	}

	// Kích thước hộp thoại
	CRect rectCtrl;
	GetWindowRect(&rectCtrl);
	irWidthDOWN = rectCtrl.Width();
	irHeightDOWN = rectCtrl.Height();

	// Độ rộng các cột dữ liệu
	for (int i = 1; i <= 6; i++) jcolDownload[i] = m_listdsdownload.GetColumnWidth(i);

	CDialogEx::OnCancel();
}


void FrmTailieuthamkhao::OnBnClickedCheckFolder()
{
	iSaveCreateFolder = m_checkFolder.GetCheck();
}

void FrmTailieuthamkhao::OnBnClickedCheckCopy()
{
	szCopyMulti = L"";
	iSaveCopyMulti = m_checkCopy.GetCheck();
}

int FrmTailieuthamkhao::CreateFileListDownload(CString spathFile)
{
	try
	{
		SqlString.Format(L"CREATE_DB=\"%s\" General\0", spathFile);
		if (::SQLConfigDataSource(NULL, ODBC_ADD_DSN, ODBC_Driver, SqlString))
		{
			if (ff._FileExists(spathFile, 1) != 1) return 0;

			CDatabase Database;
			SqlString.Format(L"ODBC;DRIVER={%s};DSN='';DBQ=%s", ODBC_Driver, spathFile);
			Database.Open(NULL, false, false, SqlString);
			Database.ExecuteSQL(
				L"CREATE TABLE tblListdown("
				L"Filename VARCHAR(255),"
				L"Fullpath LONGTEXT,"
				L"Datedown VARCHAR(10),"
				L"Download LONGTEXT"
				L");");
			Database.Close();

			return 1;
		}
	}
	catch (...) {}
	return 0;
}

void FrmTailieuthamkhao::GetBaseDecodeEx()
{
	// Connection Sever
	szBSever = bb.BaseDecodeEx(LOCALSEVER);
	szBFolderdown = bb.BaseDecodeEx(FOLDER_SEVER);
	szBFilecategory = bb.BaseDecodeEx(FILE_CATEGORY);
	szBFilesave = bb.BaseDecodeEx(FILE_SAVE);

	// Check License	
	szCreateKey = bb.BaseDecodeEx(CreateKeyRun);
	szFileDll = bb.BaseDecodeEx(LicenseDll);
	szFunction = bb.BaseDecodeEx(LicenseFunc);
}

int FrmTailieuthamkhao::CheckLicense()
{
	try
	{
		int ikey = -1;
		typedef bool(__stdcall *p)();
		HINSTANCE loadDLL = LoadLibrary((_bstr_t)szFileDll);
		if (loadDLL != NULL)
		{
			p pc = (p)GetProcAddress(loadDLL, (_bstr_t)szFunction);
			if (pc != NULL) ikey = pc();
		}

		FreeLibrary(loadDLL);

		if (ikey == 1) ikey = 2;
		else if (ikey == 2) ikey = 1;
		else if (ikey == 3) ikey = 0;

		return ikey;
	}
	catch (...) {}
	return -1;
}

void FrmTailieuthamkhao::SetDefault()
{
	m_txtTimkiem.SubclassDlgItem(DWTL_TXT_KEY, this);
	m_txtTimkiem.SetBkColor(RGBWhite);
	m_txtTimkiem.SetTextColor(RGBBlue);

	m_txtPathBrowse.SubclassDlgItem(DWTL_TXT_BROWSE, this);
	m_txtPathBrowse.SetBkColor(RGBWhite);
	m_txtPathBrowse.SetTextColor(RGBBlue);

	m_txtTimkiem.SetCueBanner(L"Nhập từ khóa tìm kiếm vào đây...");
	m_cbbPhanloai.AdjustDroppedWidth();
	m_cbbPhanloai.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);
	
	m_checkFolder.SetCheck(iSaveCreateFolder);
	m_checkCopy.SetCheck(iSaveCopyMulti);

	int izoom = 8;
	int isize = 24 * jScaleLayout;
	short shBtnColor = 30;

	/*m_btn_search.SetIcon(IDI_ICON_SEARCH, isize, isize);
	m_btn_search.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);
	m_btn_search.SetBtnCursor(IDC_CURSOR_HAND);
	m_btn_search.DrawTransparent();
	m_btn_search.DrawBorder(FALSE);*/

	/*m_btn_open.SetIcon(IDI_ICON_BROWSE, isize + izoom, isize + izoom, IDI_ICON_BROWSE, isize, isize);
	m_btn_open.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);
	m_btn_open.DrawTransparent();

	m_btn_download.SetIcon(IDI_ICON_DOWNLOAD, isize + izoom, isize + izoom, IDI_ICON_DOWNLOAD, isize, isize);
	m_btn_download.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);
	m_btn_download.DrawTransparent();

	izoom = 8, isize = 32;
	m_btn_close.SetIcon(IDI_ICON_CLOSE, isize + izoom, isize + izoom, IDI_ICON_CLOSE, isize, isize);
	m_btn_close.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);
	m_btn_close.DrawTransparent();*/

	m_btn_open.SetThemeHelper(&m_ThemeHelper);
	m_btn_open.SetIcon(IDI_ICON_BROWSE, isize, isize);
	m_btn_open.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);

	m_btn_download.SetThemeHelper(&m_ThemeHelper);
	m_btn_download.SetIcon(IDI_ICON_DOWNLOAD_BASIC, isize, isize);
	m_btn_download.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);

	m_btn_close.SetThemeHelper(&m_ThemeHelper);
	m_btn_close.SetIcon(IDI_ICON_CLOSE, isize, isize);
	m_btn_close.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);

	szpathDownloadTL.Format(L"%s%s", spathDefault, PATH_DOWNLOAD);
	CString szval = ff._ReadFileText(szpathDownloadTL);
	if (szval != L"") spathSaveDownload = szval;
	
	if (spathSaveDownload.GetLength() <= 1)
	{
		spathSaveDownload.Format(L"%s%s\\", spathDefault, FOLDERDOWNLOAD);
	}

	m_txtPathBrowse.SetWindowText(spathSaveDownload);

	spathLocal = spathSaveDownload;
}

void FrmTailieuthamkhao::StyleListdulieu()
{
	m_listdsdownload.GetHeaderCtrl()->SetFont(&m_FontHeaderList);

	m_listdsdownload.InsertColumn(0, L"", LVCFMT_CENTER, jcolDownload[0]);
	m_listdsdownload.InsertColumn(clmKey, L"Tài khoản", LVCFMT_LEFT, jcolDownload[1]);
	m_listdsdownload.InsertColumn(clmIcon, L"", LVCFMT_CENTER, jcolDownload[2]);
	m_listdsdownload.InsertColumn(clmTen, L"Nội dung tài liệu", LVCFMT_LEFT, jcolDownload[3]);
	m_listdsdownload.InsertColumn(clmPloai, L"Phân loại", LVCFMT_LEFT, jcolDownload[4]);
	m_listdsdownload.InsertColumn(clmMota, L"Mô tả", LVCFMT_LEFT, jcolDownload[5]);
	m_listdsdownload.InsertColumn(clmTrangthai, L"Trạng thái", LVCFMT_CENTER, jcolDownload[6]);
	m_listdsdownload.InsertColumn(clmDuongdan, L"Đường dẫn", LVCFMT_LEFT, jcolDownload[7]);
	m_listdsdownload.InsertColumn(clmID, _T("ID"), LVCFMT_CENTER, jcolDownload[8]);

	m_listdsdownload.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);
}

void FrmTailieuthamkhao::DeleteAllCombobox()
{
	m_cbbPhanloai.ResetContent();
	ASSERT(m_cbbPhanloai.GetCount() == 0);
}

void FrmTailieuthamkhao::DeleteAllListdulieu()
{
	if (m_listdsdownload.GetItemCount() > 0)
	{
		m_listdsdownload.DeleteAllItems();
		ASSERT(m_listdsdownload.GetItemCount() == 0);
	}
}

int FrmTailieuthamkhao::GetAllListdownload()
{
	try
	{
		ObjConn.OpenAccessDB(szpathListdown, L"", L"");

		MyStructListDown MSListdownload;
		SqlString = L"SELECT * FROM tblListdown ORDER BY Filename ASC;";
		CRecordset *Recset = ObjConn.Execute(SqlString);
		while (!Recset->IsEOF())
		{
			Recset->GetFieldValue(L"Fullpath", MSListdownload.szPath);
			Recset->GetFieldValue(L"Download", MSListdownload.szBase);
			MSListdownload.szBase = bb.BaseDecodeEx(MSListdownload.szBase);
			
			// Lưu trữ thông tin dữ liệu được download
			vecListDownload.push_back(MSListdownload);
			Recset->MoveNext();
		}
		Recset->Close();
		ObjConn.CloseDatabase();

		return (int)vecListDownload.size();
	}
	catch (...) {}
	return 0;
}

int FrmTailieuthamkhao::DownloadFileCategory()
{
	try
	{
		CString szFile = TENFILESUDUNG;
		if (szFile == L"NULL") return 0;

		CString szLinkURL = szBSever + szBFolderdown + szBFilecategory;
		CString szpathSave = spathDefault + szBFilesave;

		DeleteUrlCacheEntry(szLinkURL);
		URLDownloadToFileW(0, szLinkURL, szpathSave, 0, 0);
		if (ff._FileExists(szpathSave, 0) == 1)
		{
			vector<CString> vecLine;
			ff._ReadMultiUnicode(szpathSave, vecLine);

			DeleteFile(szpathSave);

			if ((int)vecLine.size() > 0)
			{
				int pos = 0;
				int iLen = szFile.GetLength();
				CString szval = L"", sztenfile = L"";
				MyStrCategory MSCategory;				

				for (int i = 0; i < (int)vecLine.size(); i++)
				{
					vecLine[i].Trim();
					if (vecLine[i] != L"" && vecLine[i].Left(2) != L"//")
					{
						szval = vecLine[i];
						pos = szval.Find(L"|");
						if (pos > 0)
						{
							sztenfile = szval.Left(pos);
							sztenfile.Trim();
							if (sztenfile.Left(iLen) == szFile)
							{
								MSCategory.filename = sztenfile;
								MSCategory.szfolder = L"";
								MSCategory.szcontent = L"";

								szval = szval.Right(szval.GetLength() - pos - 1);
								szval.Trim();
								pos = szval.Find(L"|");
								if (pos > 0)
								{
									MSCategory.szfolder = szval.Left(pos);
									MSCategory.szfolder.Trim();

									MSCategory.szcontent = szval.Right(szval.GetLength() - pos - 1);
									MSCategory.szcontent.Trim();

									if (MSCategory.filename != L"" &&
										MSCategory.szfolder != L"" && MSCategory.szcontent != L"")
									{
										// Lưu trữ thông tin tệp tin sử dụng
										vecCategory.push_back(MSCategory);
									}
								}
							}
						}
					}
				}
			}

			// Xóa vector trung gian
			vecLine.clear();
		}

		return (int)vecCategory.size();
	}
	catch (...) {}

	return 0;
}


void FrmTailieuthamkhao::LoadComboboxBodulieu()
{
	try
	{
		if ((int)vecCategory.size() > 0)
		{
			CString szFile = TENFILESUDUNG;
			int iLen = szFile.GetLength();
			for (int i = 0; i < (int)vecCategory.size(); i++)
			{
				if (vecCategory[i].filename.Left(iLen) == szFile)
				{
					// Thêm dữ liệu vào combobox
					m_cbbDulieu.AddString(vecCategory[i].szcontent);										
				}
			}

			// Xác định tệp tin sử dụng mặc định
			szFileCategory = vecCategory[0].filename;
			szFolderCategory = vecCategory[0].szfolder;
			szContentCaterory = vecCategory[0].szcontent;

			m_cbbDulieu.SetCurSel(0);
		}
	}
	catch (...) {}
}

int FrmTailieuthamkhao::GetAllDulieu()
{
	try
	{
		vecDSDulieu.clear();
		vector<CString> vecDSTotal;

		// Xóa text-search
		m_txtTimkiem.SetWindowText(L"");

		// Xóa toàn bộ list image icon
		int nsizeIcon = 16 * jScaleLayout;
		m_imageList.DeleteImageList();
		ASSERT(m_imageList.GetSafeHandle() == NULL);
		m_imageList.Create(nsizeIcon, nsizeIcon, ILC_COLOR32, 0, 0);
		
		CBitmap bm;
		bm.LoadBitmap(IDB_BITMAP_CHECKDOWN);
		m_imageList.Add(&bm, LVSIL_SMALL);

		// Đọc toàn bộ dữ liệu từ tệp tin được tải xuống
		CheckTotalFiles(vecDSTotal);
		if ((int)vecDSTotal.size() > 0)
		{
			CString sztype = L"";
			int ipquyen = 0, iLen = 0, dem = 0;;

			MyStructDownload MSDownload;

			vector<CString> vecLine;

			// Lọc dữ liệu phân quyền
			for (int i = 0; i < (int)vecDSTotal.size(); i++)
			{
				vecDSTotal[i].Trim();
				if (vecDSTotal[i] != L"" && vecDSTotal[i].Left(2) != L"//")
				{
					MSDownload.iNumID = dem;

					ff._TackTukhoa(vecLine, vecDSTotal[i], L"|", L"|", 0, 0);
					for (int j = 0; j < 5; j++) vecLine.push_back(L"");

					MSDownload.sNoidung = vecLine[0];
					MSDownload.sPhanloai = vecLine[1];
					MSDownload.sLinkdown = vecLine[2];
					MSDownload.sMota = vecLine[3];
					MSDownload.sLocal = L"";

					MSDownload.sTypeFile = L"";
					iLen = vecLine[2].GetLength();
					for (int j = iLen - 1; j > 0; j--)
					{
						if (vecLine[2].Mid(j, 1) == _T("."))
						{
							MSDownload.sTypeFile = vecLine[2].Right(iLen - j - 1);
							break;
						}
					}

					// Hình ảnh đại diện
					sztype = L"*." + MSDownload.sTypeFile;
					HICON hI = m_sysImageList.GetFileIcon(sztype);
					m_imageList.Add(hI);

					// Phân quyền
					ipquyen = _wtoi(vecLine[4]);
					if (ipquyen < 0 || ipquyen>2) ipquyen = 0;

					// Lấy giá trị ngẫu nhiên để test
					//ipquyen = rand() % 3;

					MSDownload.iPhanquyen = ipquyen;

					// Lưu vào dữ liệu
					vecDSDulieu.push_back(MSDownload);
					dem++;

					// Xóa vector trung gian
					vecLine.clear();
				}
			}
		}		

		vecDSTotal.clear();

		return (int)vecDSDulieu.size();
	}
	catch (...) {}

	return 0;
}

int FrmTailieuthamkhao::CheckTotalFiles(vector<CString> &vecTotal)
{
	try
	{
		CString szLinkURL = szBSever + szBFolderdown + szFileCategory;
		CString szpathSave = spathDefault + szBFilesave;

		DeleteUrlCacheEntry(szLinkURL);
		URLDownloadToFileW(0, szLinkURL, szpathSave, 0, 0);
		if (ff._FileExists(szpathSave, 0) == 1)
		{
			ff._ReadMultiUnicode(szpathSave, vecTotal);
			DeleteFile(szpathSave);
		}

		return (int)vecTotal.size();
	}
	catch (...) {}
	return 0;
}

void FrmTailieuthamkhao::LocPhanloaidulieu()
{
	try
	{
		DeleteAllCombobox();

		vector<CString> vecLine;
		vector<CString> vecPhanloai;
		m_cbbPhanloai.AddString(ALLCOMBOBOX);

		if ((int)vecDSDulieu.size() > 0)
		{
			int icheck = 0;
			for (int i = 0; i < (int)vecDSDulieu.size(); i++)
			{
				if (vecDSDulieu[i].sPhanloai != L"")
				{
					ff._TackTukhoa(vecLine, vecDSDulieu[i].sPhanloai, L";", L";", 1, 0);

					for (int j = 0; j < (int)vecLine.size(); j++)
					{
						icheck = 0;
						if ((int)vecPhanloai.size() > 0)
						{
							for (int k = 0; k < (int)vecPhanloai.size(); k++)
							{
								if (vecLine[j] == vecPhanloai[k])
								{
									icheck = 1;
									break;
								}
							}
						}

						if (icheck == 0)
						{
							vecPhanloai.push_back(vecLine[j]);
							m_cbbPhanloai.AddString(vecLine[j]);
						}
					}

					// Xóa vector trung gian
					vecLine.clear();
				}				
			}			
		}

		m_cbbPhanloai.SetCurSel(0);
		vecPhanloai.clear();
	}
	catch (...) {}
}


int FrmTailieuthamkhao::LoadDulieusudung(CString szTk, CString szPl)
{
	try
	{
		// Xóa toàn bộ list dữ liệu
		DeleteAllListdulieu();
		vecDSKetqua.clear();		

		int nTotal = (int)vecDSDulieu.size();
		if (nTotal > 0)
		{
			// Xác định kiểu tìm kiếm
			int itype = 0;
			if (szTk != L"" && szPl == L"") itype = 1;
			else if (szTk == L"" && szPl != L"") itype = 2;
			else if (szTk != L"" && szPl != L"") itype = 3;

			if (itype == 0)
			{
				for (int i = 0; i < nTotal; i++)
				{
					vecDSKetqua.push_back(vecDSDulieu[i]);
				}
			}
			else if (itype == 1)
			{
				for (int i = 0; i < nTotal; i++)
				{
					if (ff._FindKhongdau(vecDSDulieu[i].sNoidung, szTk) == 1)
					{
						vecDSKetqua.push_back(vecDSDulieu[i]);
					}
				}
			}
			else if (itype == 2)
			{
				for (int i = 0; i < nTotal; i++)
				{
					if (vecDSDulieu[i].sPhanloai.Find(szPl) >= 0)
					{
						vecDSKetqua.push_back(vecDSDulieu[i]);
					}
				}
			}
			else if (itype == 3)
			{
				for (int i = 0; i < nTotal; i++)
				{
					if (vecDSDulieu[i].sPhanloai.Find(szPl) >= 0)
					{
						if (ff._FindKhongdau(vecDSDulieu[i].sNoidung, szTk) == 1)
						{
							vecDSKetqua.push_back(vecDSDulieu[i]);
						}
					}
				}
			}

/****** Hiển thị kết quả **************************************************/

			tslkq = (int)vecDSKetqua.size();

			lanshow = 1;
			int iz = ibuocnhay;
			if (tslkq <= iz)
			{
				iz = tslkq;
				iStopload = 0;
			}
			else iStopload = 1;

			// Load Image vào list dữ liệu
			m_listdsdownload.SetImageList(&m_imageList);

			// Thêm dữ liệu vào list kết quả
			AddItemsListKetqua(0, iz);

			CString szval = L"";
			szval.Format(L"Nội dung tài liệu (%d kết quả)", tslkq);
			_NameColumnHeading(m_listdsdownload, clmTen, 1, szval);

			return iz;
		}		
	}
	catch (...) {}

	return 0;
}


CString FrmTailieuthamkhao::_NameColumnHeading(CListCtrl &clist, int column, int itype, CString szName)
{
	try
	{
		CString strNome = L"";
		if (itype == 1) strNome = szName;

		CHeaderCtrl *pHeader = clist.GetHeaderCtrl();
		HDITEM hdi;
		hdi.mask = HDI_TEXT;
		hdi.pszText = strNome.GetBuffer(256);
		hdi.cchTextMax = 256;

		if (itype == 0) pHeader->GetItem(column, &hdi);
		else pHeader->SetItem(column, &hdi);
		strNome.ReleaseBuffer();

		return strNome;
	}
	catch (...) {}
	return L"";
}


void FrmTailieuthamkhao::SetColorType(int numRow)
{
	try
	{
		CString szID = m_listdsdownload.GetItemText(numRow, clmID);
		int nID = _wtoi(szID);
		if (nID < (int)vecDSDulieu.size())
		{
			if (vecDSDulieu[nID].sLocal != L"")
			{
				for (int i = clmTen; i <= clmMota; i++)
				{
					m_listdsdownload.SetCellColors(numRow,i , RGBYellowCanary, RGBBlack);
				}				
			}
			else
			{
				if (vecDSDulieu[nID].sLinkdown == L"")
				{
					m_listdsdownload.SetRowColors(numRow, RGBWhite, RGBRed);
				}
				else
				{
					CString szval = vecDSDulieu[nID].sNoidung.Left(5);
					szval.MakeLower();

					// Lấy giá trị ngẫu nhiên để test
					//CString sztest[4] = { L"hot",L"new",L"upd",L"" };
					//szval = sztest[rand() % 4];

					if (szval.Find(L"hot") >= 0)
					{
						m_listdsdownload.SetItemText(numRow, clmTrangthai, L"HAY!!");
						m_listdsdownload.SetCellColors(numRow, clmTrangthai, RGBStrongRed, RGBWhite);
					}
					else if (szval.Find(L"new") >= 0)
					{
						m_listdsdownload.SetItemText(numRow, clmTrangthai, L"MỚI");
						m_listdsdownload.SetCellColors(numRow, clmTrangthai, RGBPigmentGreen, RGBWhite);
					}
					else if (szval.Find(L"upd") >= 0)
					{
						m_listdsdownload.SetItemText(numRow, clmTrangthai, L"Cập nhật");
						m_listdsdownload.SetCellColors(numRow, clmTrangthai, RGBStrongBlue, RGBWhite);
					}
				}
			}
		}		
	}
	catch (...) {}
}


void FrmTailieuthamkhao::OnNMClickListKqua(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	*pResult = 0;
}


void FrmTailieuthamkhao::OnNMDblclkListKqua(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem == clmKey)
	{
		CString szpath = m_listdsdownload.GetItemText(nItem, clmDuongdan);
		if (szpath != L"") OnDwOpenfile();
	}

	*pResult = 0;
}


void FrmTailieuthamkhao::OnNMRClickListKqua(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	*pResult = 0;
}


void FrmTailieuthamkhao::OnLvnEndScrollListKqua(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if (iStopload == 0) return;

		RECT r;
		CRect rectCtrl;
		m_listdsdownload.GetItemRect(0, &r, LVIR_BOUNDS);
		m_listdsdownload.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;
		int b = rectCtrl.Height();
		int topIndex = m_listdsdownload.GetTopIndex();
		int ncount = m_listdsdownload.GetItemCount();
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
			AddItemsListKetqua(ibuocnhay * (lanshow - 1), iz);

			if (lanshow % 4 == 0) ff._s(L"Bạn có thể nhập từ khóa để hiển thị kết quả chính xác hơn.", 0, 1, 5000);
		}
	}
	catch (...) {}
}

void FrmTailieuthamkhao::OnHdnEndtrackListKqua(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = jcolDownload[0];
	else if (pSubItem == clmIcon) pNMListHeader->pitem->cxy = jcolDownload[2];
	else if (pSubItem == clmID) pNMListHeader->pitem->cxy = 0;

	*pResult = 0;
}

void FrmTailieuthamkhao::SetLocationAndSize()
{
	try
	{
		if (irWidthDOWN > 0 && irHeightDOWN > 0)
		{
			int isx = GetSystemMetrics(SM_CXSCREEN);
			int isy = GetSystemMetrics(SM_CYSCREEN);

			if (irWidthDOWN >= isx && irHeightDOWN + 50 >= isy)
			{
				this->ShowWindow(SW_MAXIMIZE);
			}
			else
			{
				this->SetWindowPos(NULL, 0, 0, irWidthDOWN, irHeightDOWN, SWP_NOMOVE | SWP_NOZORDER);

				CRect r;
				GetWindowRect(&r);
				ScreenToClient(&r);
				MoveWindow((isx - r.Width()) / 2, (isy - r.Height()) / 2, r.right - r.left, r.bottom - r.top, TRUE);
			}
		}
	}
	catch (...) {}
}

void FrmTailieuthamkhao::AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		int iLen = 0;
		int ipquyen = 0;
		CString szval = L"";
		CString szTennd = L"";

		int numID = 0;
		int islFileLuu = (int)vecListDownload.size();

		for (int i = iBegin; i < iEnd; i++)
		{
			numID = vecDSKetqua[i].iNumID;
			m_listdsdownload.InsertItem(i, _T(""), 0);
			
			szval.Format(L"%d", numID);
			m_listdsdownload.SetItemText(i, clmID, szval);

			szTennd = vecDSKetqua[i].sNoidung;
			szval = szTennd.Left(5);
			szval.MakeLower();
			if (szval == L"[hot]" || szval == L"[new]" || szval == L"[upd]")
			{
				szTennd = szTennd.Right(szTennd.GetLength() - 5);
				szTennd.Trim();
			}

			m_listdsdownload.SetItemText(i, clmTen, szTennd);
			m_listdsdownload.SetItemText(i, clmPloai, vecDSKetqua[i].sPhanloai);
			m_listdsdownload.SetItemText(i, clmMota, vecDSKetqua[i].sMota);

			// Load hình ảnh đại diện
			m_listdsdownload.SetItem(i, clmIcon, LVIF_IMAGE, NULL, numID + 3, 0, 0, 0);

			// Kiểm tra đường dẫn có trong CSDL chưa
			if (islFileLuu > 0)
			{
				for (int j = 0; j < islFileLuu; j++)
				{
					if (vecDSKetqua[i].sLinkdown == vecListDownload[j].szBase)
					{
						vecDSKetqua[i].sLocal = vecListDownload[j].szPath;
						break;
					}
				}
			}

			if (vecDSKetqua[i].sLocal != L"")
			{
				m_listdsdownload.SetItemText(i, clmDuongdan, vecDSKetqua[i].sLocal);
				m_listdsdownload.SetItemText(i, clmKey, L"Đã tải xuống");
				m_listdsdownload.SetItem(i, 0, LVIF_IMAGE, NULL, 2, 0, 0, 0);
			}
			else
			{
				ipquyen = vecDSKetqua[i].iPhanquyen;
				if (ipquyen == 1)
				{
					m_listdsdownload.SetItemText(i, clmKey, L"Khóa mềm");
					m_listdsdownload.SetCellColors(i, clmKey, RGBWhite, RGBBlue);
				}
				else if (ipquyen == 2)
				{
					m_listdsdownload.SetItemText(i, clmKey, L"Khóa cứng");
					m_listdsdownload.SetCellColors(i, clmKey, RGBWhite, RGBStrongRed);
				}
				else
				{
					m_listdsdownload.SetItemText(i, clmKey, L"Miễn phí");
				}

				if (ipquyen > jUserID) m_listdsdownload.SetItem(i, 0, LVIF_IMAGE, NULL, 0, 0, 0, 0);
				else m_listdsdownload.SetItem(i, 0, LVIF_IMAGE, NULL, 1, 0, 0, 0);
			}			

			// Định dạng màu dữ liệu
			SetColorType(i);
		}
	}
	catch (...) {}
}


int FrmTailieuthamkhao::AutoCheckRobot()
{
	jCheckRotot = 0;

	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	FrmCheckRobot pp;
	pp.DoModal();

	return jCheckRotot;
}

int FrmTailieuthamkhao::CheckDownload(int numRow)
{
	CString szID = m_listdsdownload.GetItemText(numRow, clmID);
	int nID = _wtoi(szID);
	if (nID < (int)vecDSDulieu.size())
	{
		if (vecDSDulieu[nID].iPhanquyen <= jUserID) return 1;
	}
	
	return 0;
}

void FrmTailieuthamkhao::RightClickDownload(int itype)
{
	try
	{
		if (iTopdownload >= iMaxdownload)
		{
			if (AutoCheckRobot() == 0) return;
			iTopdownload = 0;
		}

		if (jUserID == -1)
		{
			ff._msgbox(L"Không thể xác nhận bản quyền phần mềm nên bạn không thể tải tài liệu.", 2);
			return;
		}

		int tongdown = 0;
		int ncount = m_listdsdownload.GetItemCount();
		if (ncount == 0) return;

		m_txtPathBrowse.GetWindowTextW(spathLocal);
		if (spathLocal == L"") spathLocal.Format(L"%s%s\\", spathDefault, FOLDERDOWNLOAD);

		CString szval = L"";
		int iDownloadmax = 1;
		if (jUserID == 2) iDownloadmax = 10;

		for (int i = 0; i < ncount; i++)
		{
			if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) tongdown++;
			if (tongdown > iDownloadmax)
			{
				if (jUserID == 2) szval.Format(L"%s (%d) %s", L"Mỗi lượt tải không quá", iDownloadmax, L"tệp tin. Vui lòng giảm bớt tích chọn.");
				else szval = L"Bạn chỉ có thể tải từng tệp tin một. Nâng cấp lên khóa cứng để có thể tải nhiều tệp tin cùng lúc.";
				ff._msgbox(szval, 1);
				return;
			}
		}

		if (tongdown > 0)
		{
			CString szTitle = L"";
			this->GetWindowTextW(szTitle);

			int iselect = 0;
			int inotdown = 0;
			int dem = 0, iDown = 0;
			for (int i = 0; i < ncount; i++)
			{
				if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					iselect++;
					iTopdownload++;
					if (iTopdownload >= iMaxdownload)
					{
						if (AutoCheckRobot() == 0) return;
						iTopdownload = 0;
					}

					szval.Format(L"%s(%d%s)", L"Đang tải tệp tin. Vui lòng đợi trong giây lát... ", (int)(100 * (dem + 1) / tongdown), L"%");
					this->SetWindowText(szval);

					if (dem > 0) Sleep(1000);
					if (CheckDownload(i) == 1)
					{
						iDown += DownloadFile(i, itype, 1);
					}
					else inotdown = 1;

					dem++;
				}

				if (dem > iDownloadmax) break;
			}

			if (iDown > 0)
			{
				if (itype == 0)
				{
					ff._msgbox(L"Tải dữ liệu thành công!", 0, 0, 3000);
				}
			}
			else
			{
				if (iselect == 1 && inotdown == 1)
				{
					ff._msgbox(L"Tài khoản của bạn không được cấp quyền để tải tài liệu này. "
						L"Bạn có thể tham khảo cột 'Tài khoản' để biết có thể tải được các tài liệu nào.", 1);
				}
			}
			
			this->SetWindowText(szTitle);
		}
		else
		{
			ff._msgbox(L"Bạn chưa tích chọn dữ liệu cần tải. Vui lòng kiểm tra lại.", 1);
		}
	}
	catch (...) {}
}


int FrmTailieuthamkhao::DownloadFile(int numRow, int itype, int iMess)
{
	try
	{
		CString szID = m_listdsdownload.GetItemText(numRow, clmID);
		int nID = _wtoi(szID);
		if (nID < (int)vecDSDulieu.size())
		{
			if (vecDSDulieu[nID].sLinkdown == L"")
			{
				if (iMess == 1)
				{
					ff._msgbox(L"Các dòng màu đỏ chưa cập nhật đường dẫn tải dữ liệu. "
						L"Vui lòng liên hệ với tổng đài hỗ trợ GXD 1900.0147 để được xử lý.");
				}
				return 0;
			}

			if (ff._CheckConnection() == TRUE)
			{
				CString szName = ff._TackTenFile(vecDSDulieu[nID].sLinkdown, 1);
				if (szName != L"")
				{
					CString szRep = L"";
					CString szpathSave = spathLocal + szName;

					// Xác định các thư mục cần tạo					
					if (iSaveCreateFolder == 1) szRep = Taothumuccon(vecDSDulieu[nID].sLinkdown);
					if (szRep != L"") szpathSave = szRep + szName;
					szpathSave.Replace(L"\\\\", L"\\");

					if (ff._FileExists(szpathSave, 0) != 1)
					{
						CString szLinkURL = szBSever + vecDSDulieu[nID].sLinkdown;

						DeleteUrlCacheEntry(szLinkURL);
						URLDownloadToFileW(0, szLinkURL, szpathSave, 0, 0);
						if (ff._FileExists(szpathSave, 0) != 1)
						{
							DeleteUrlCacheEntry(vecDSDulieu[nID].sLinkdown);
							URLDownloadToFileW(0, vecDSDulieu[nID].sLinkdown, szpathSave, 0, 0);
							if (ff._FileExists(szpathSave, 0) != 1)
							{
								if (iMess == 1)
								{
									ff._msgbox(L"Đường dẫn không tồn tại, không đúng hoặc đã hết hạn.", 2);
								}
								return 0;
							}
						}
					}

					Savepathdownload(szName, szpathSave, vecDSDulieu[nID].sLinkdown);
					SetDownloadColor(numRow, szpathSave);

					if (itype == 1)
					{
						try
						{
							if (2 == (int)ShellExecuteW(NULL, L"open", szpathSave, NULL, NULL, SW_SHOWMAXIMIZED))
							{
								ff._msgbox(L"Tệp tin không tồn tại hoặc đã bị xóa. Vui lòng kiểm tra lại.", 2);
							}
						}
						catch (...) {}
					}

					return 1;
				}
			}
		}
	}
	catch (...) {}
	return 0;
}


CString FrmTailieuthamkhao::Taothumuccon(CString szpath)
{
	try
	{
		CString szthuvienrep = szBFolderdown + szFolderCategory;

		szpath.Replace(L" ", L"");
		szpath.Replace(szthuvienrep, L"");
		szpath = ff._ConvertKhongDau(szpath);

		CString szFolder = L"";
		int nlen = szpath.GetLength();
		for (int i = nlen - 1; i >= 0; i--)
		{
			if (szpath.Mid(i, 1) == L"\\" || szpath.Mid(i, 1) == L"/")
			{
				szFolder = szpath.Left(i);
				break;
			}
		}

		CString szlocal = spathLocal;
		if (szFolder != L"")
		{
			if (szlocal.Right(1) != L"\\") szlocal += L"\\";

			vector<CString> vecLine;
			ff._TackTukhoa(vecLine, szFolder, L"\\", L"/", 0, 0);
			for (int i = 0; i < (int)vecLine.size(); i++)
			{
				szlocal += vecLine[i];
				szlocal += L"\\";
				ff._CreateNewFolder(szlocal);
			}

			// Xóa vector trung gian
			vecLine.clear();
		}
		return szlocal;
	}
	catch (...) {}
	return L"";
}

void FrmTailieuthamkhao::SetDownloadColor(int numRow, CString szpth)
{
	CString szID = m_listdsdownload.GetItemText(numRow, clmID);
	int nID = _wtoi(szID);
	if (nID < (int)vecDSDulieu.size())
	{
		for (int i = clmTen; i <= clmMota; i++)
		{
			m_listdsdownload.SetCellColors(numRow, i, RGBYellowCanary, RGBBlack);
		}

		m_listdsdownload.SetItemText(numRow, clmKey, L"Đã tải xuống");
		m_listdsdownload.SetItemText(numRow, clmDuongdan, szpth);
		vecDSDulieu[nID].sLocal = szpth;

		m_listdsdownload.SetItem(numRow, 0, LVIF_IMAGE, NULL, 2, 0, 0, 0);
	}
}

void FrmTailieuthamkhao::Savepathdownload(CString szname, CString szpth, CString szURL)
{
	try
	{
		ObjConn.OpenAccessDB(szpathListdown, L"", L"");

		CString szdem = L"";
		szURL = bb.BaseEncodeEx(szURL, 0);
		SqlString.Format(L"SELECT COUNT (*) AS iDem FROM tblListdown WHERE Download LIKE '%s';", szURL);
		CRecordset* Recset = ObjConn.Execute(SqlString);
		Recset->GetFieldValue(L"iDem", szdem);
		Recset->Close();

		CString szDay = ff._GetTimeNow(0);
		if (_wtoi(szdem) > 0)
		{
			SqlString.Format(L"UPDATE tblListdown SET Filename='%s',Fullpath='%s',Datedown='%s' "
				L"WHERE Download LIKE '%s';", szname, szpth, szDay, szURL);
		}
		else
		{
			SqlString.Format(L"INSERT INTO tblListdown (Filename,Fullpath,Datedown,Download) "
				L"VALUES ('%s','%s','%s','%s');", szname, szpth, szDay, szURL);
		}

		ObjConn.ExecuteRB(SqlString);
		ObjConn.CloseDatabase();
	}
	catch (...) {}
}


void FrmTailieuthamkhao::OnDwDownload()
{
	RightClickDownload(0);
}


void FrmTailieuthamkhao::OnDwDwandview()
{
	RightClickDownload(1);
}

void FrmTailieuthamkhao::OnDwOpenfile()
{
	try
	{
		int ncount = m_listdsdownload.GetItemCount();
		if (ncount == 0) return;

		int iselected = -1;
		for (int i = 0; i < ncount; i++)
		{
			if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iselected = i;
				break;
			}
		}

		if (iselected >= 0)
		{
			CString szpath = m_listdsdownload.GetItemText(iselected, clmDuongdan);
			if (szpath != L"")
			{
				try
				{
					if (2 == (int)ShellExecuteW(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED))
					{
						ff._msgbox(L"Tệp tin không tồn tại hoặc đã bị xóa. Vui lòng kiểm tra lại.", 2);
					}
				}
				catch (...) {}
			}
		}
	}
	catch (...) {}
}

void FrmTailieuthamkhao::OnDwOpenfolder()
{
	try
	{
		int ncount = m_listdsdownload.GetItemCount();
		if (ncount == 0) return;

		int iselected = -1;
		for (int i = 0; i < ncount; i++)
		{
			if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				iselected = i;
				break;
			}
		}

		if (iselected >= 0)
		{
			CString szpath = m_listdsdownload.GetItemText(iselected, clmDuongdan);
			if (szpath != L"")
			{
				CString szFolder = L"";
				int nlen = szpath.GetLength();
				for (int i = nlen - 1; i >= 0; i--)
				{
					if (szpath.Mid(i, 1) == L"\\")
					{
						szFolder = szpath.Left(i);
						break;
					}
				}

				try
				{
					if (szFolder != L"")
					{
						ShellExecuteW(NULL, L"open", szFolder, NULL, NULL, SW_SHOWMAXIMIZED);
					}
				}
				catch (...) {}
			}
		}
	}
	catch (...) {}
}

void FrmTailieuthamkhao::OnDwCopyNoidung()
{
	if (nSubItem == -1) return;

	CString szval = L"";
	CString szNoidung = L"";
	int ncount = m_listdsdownload.GetItemCount();
	if (ncount > 0)
	{
		for (int i = 0; i < ncount; i++)
		{
			if (m_listdsdownload.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
			{
				szval = m_listdsdownload.GetItemText(i, nSubItem);
				szNoidung += szval;
				szNoidung += L"\n";
			}
		}

		szNoidung = szNoidung.Left(szNoidung.GetLength() - 1);

		if (iSaveCopyMulti == 1)
		{
			if (szCopyMulti != L"") szCopyMulti += L"\n";
			szCopyMulti += szNoidung;			
			ff._CopyData(szCopyMulti);
		}
		else
		{
			ff._CopyData(szNoidung);
		}		
	}

	
}

void FrmTailieuthamkhao::OnDwShowAll()
{
	int icw[9] = { 25, 95, 22, 500, 200, 300, 85, 0, 0 };
	for (int i = 1; i <= 6; i++)
	{
		ShowColumnWidth(i, icw[i], 1);
	}
}

void FrmTailieuthamkhao::OnDwShowKey()
{
	ShowColumnWidth(clmKey, 95, 0);
}

void FrmTailieuthamkhao::OnDwShowNoidung()
{
	ShowColumnWidth(clmTen, 500, 0);
}

void FrmTailieuthamkhao::OnDwShowPhanloai()
{
	ShowColumnWidth(clmPloai, 200, 0);
}

void FrmTailieuthamkhao::OnDwShowMota()
{
	ShowColumnWidth(clmMota, 300, 0);
}

void FrmTailieuthamkhao::OnDwShowTrangthai()
{
	ShowColumnWidth(clmTrangthai, 85, 0);
}

void FrmTailieuthamkhao::OnDwShowPath()
{
	ShowColumnWidth(clmDuongdan, 300, 0);
}

void FrmTailieuthamkhao::ShowColumnWidth(int nColumn, int nWidth, int itype)
{
	LVCOLUMN col;
	col.mask = LVCF_WIDTH;
	col.cx = nWidth;

	if (itype == 0)
	{
		int icheck = m_listdsdownload.GetColumnWidth(nColumn);
		if (icheck > 0) col.cx = 0;
	}

	m_listdsdownload.SetColumn(nColumn, &col);
}

LRESULT FrmTailieuthamkhao::OnHeaderRightClick(WPARAM wParam, LPARAM lParam)
{
	try
	{
		nItem = -1;
		nSubItem = -1;

		ASSERT((HWND)wParam == m_listdsdownload.GetSafeHwnd());

		CListCtrlEx::CHeaderRightClick *hit = (CListCtrlEx::CHeaderRightClick*) lParam;

		CMenu m_Menu;
		m_Menu.CreateMenu();

		CMenu subMenu;
		subMenu.CreatePopupMenu();
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_POPUP, L"Hiển thị mặc định");
		subMenu.AppendMenuW(MF_SEPARATOR);

		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_KEY_POPUP, L"Tài khoản");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_TEN_POPUP, L"Nội dung tài liệu");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_PLOAI_POPUP, L"Phân loại");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_MOTA_POPUP, L"Mô tả");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_TYPE_POPUP, L"Trạng thái");
		subMenu.AppendMenuW(MF_STRING, IDC_COLUMN_LOCAL_POPUP, L"Đường dẫn");

		if ((int)m_listdsdownload.GetColumnWidth(clmKey) > 0) subMenu.CheckMenuItem(IDC_COLUMN_KEY_POPUP, MF_CHECKED);
		if ((int)m_listdsdownload.GetColumnWidth(clmTen) > 0) subMenu.CheckMenuItem(IDC_COLUMN_TEN_POPUP, MF_CHECKED);
		if ((int)m_listdsdownload.GetColumnWidth(clmPloai) > 0) subMenu.CheckMenuItem(IDC_COLUMN_PLOAI_POPUP, MF_CHECKED);
		if ((int)m_listdsdownload.GetColumnWidth(clmMota) > 0) subMenu.CheckMenuItem(IDC_COLUMN_MOTA_POPUP, MF_CHECKED);
		if ((int)m_listdsdownload.GetColumnWidth(clmTrangthai) > 0) subMenu.CheckMenuItem(IDC_COLUMN_TYPE_POPUP, MF_CHECKED);
		if ((int)m_listdsdownload.GetColumnWidth(clmDuongdan) > 0) subMenu.CheckMenuItem(IDC_COLUMN_LOCAL_POPUP, MF_CHECKED);

		m_Menu.AppendMenuW(MF_POPUP, reinterpret_cast<UINT_PTR>(subMenu.m_hMenu), L"");

		CMenu* popup = m_Menu.GetSubMenu(0);
		popup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON,
			hit->m_mousePos.x, hit->m_mousePos.y, this);

		return 0;
	}
	catch (...) {}

	return 0;
}


void FrmTailieuthamkhao::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;

	if (demTimer >= 2)
	{
		demTimer = 0;
		KillTimer(CONST_TIMER_TAILIEUTHAMKHAO);

		Timkiemdulieu();
	}

	CDialogEx::OnTimer(nIDEvent);
}


void FrmTailieuthamkhao::SetTimerSearch()
{
	demTimer = 0;
	SetTimer(CONST_TIMER_TAILIEUTHAMKHAO, 100, NULL);
}

void FrmTailieuthamkhao::OnEnChangeTxtKey()
{
	if (blChangeKey) SetTimerSearch();
}
