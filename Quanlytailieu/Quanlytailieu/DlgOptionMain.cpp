#include "pch.h"
#include "DlgOptionMain.h"

IMPLEMENT_DYNAMIC(DlgOptionMain, CDialogEx)

DlgOptionMain::DlgOptionMain(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgOptionMain::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	pathFolder = L"";
}

DlgOptionMain::~DlgOptionMain()
{
	delete ff;
	delete bb;
	delete reg;
}

void DlgOptionMain::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, OPTM_BTN_OK, m_btn_ok);
	DDX_Control(pDX, OPTM_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, OPTM_BTN_REGISTER_COM, m_btn_register);
	DDX_Control(pDX, OPTM_TXT_BROWSE, m_txt_browse);
	DDX_Control(pDX, OPTM_CBB_AUTOUPDATE, m_cbb_autoupdate);
	DDX_Control(pDX, OPTM_CHECK_MSGUPDATE, m_check_msgupdate);
	DDX_Control(pDX, OPTM_CHECK_DEL_FILECSV, m_check_delfilecsv);
	DDX_Control(pDX, OPTM_CHECK_START_WINDOWS, m_check_startwindows);
}

BEGIN_MESSAGE_MAP(DlgOptionMain, CDialogEx)
	ON_WM_SYSCOMMAND()

	ON_BN_CLICKED(OPTM_BTN_OK, &DlgOptionMain::OnBnClickedBtnOk)
	ON_BN_CLICKED(OPTM_BTN_CANCEL, &DlgOptionMain::OnBnClickedBtnCancel)
	ON_BN_CLICKED(OPTM_BTN_BROWSE, &DlgOptionMain::OnBnClickedBtnBrowse)
	ON_BN_CLICKED(OPTM_BTN_REGISTER_COM, &DlgOptionMain::OnBnClickedBtnRegisterCom)
END_MESSAGE_MAP()

// DlgOptionMain message handlers
BOOL DlgOptionMain::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_SUPPORT));
	SetIcon(hIcon, FALSE);

	_XacdinhSheetLienquan();
	_SetDefault();

	return TRUE;
}

BOOL DlgOptionMain::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(OPTM_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
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

void DlgOptionMain::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void DlgOptionMain::OnBnClickedBtnOk()
{
	try
	{
		CDialogEx::OnOK();

		m_txt_browse.GetWindowTextW(pathFolder);
		ff->_xlPutValue(psConfig, L"CF_DIR_DEFAULT", pathFolder, 0);

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
		reg->CreateKey(szCreate);

		CString szval = L"";
		int iAutoUpdate = m_cbb_autoupdate.GetCurSel();
		if (iAutoUpdate != 1 && iAutoUpdate != 2) iAutoUpdate = 0;
		szval.Format(L"%d", iAutoUpdate);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgAutoUpdate), ff->_ConvertCstringToChars(szval));

		int iDelFileCSV = m_check_delfilecsv.GetCheck();
		szval.Format(L"%d", iDelFileCSV);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgDeleteFileCSV), ff->_ConvertCstringToChars(szval));

/*****************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
		reg->CreateKey(szCreate);

		int iCheckUpSoft = m_check_msgupdate.GetCheck();
		szval.Format(L"%d", iCheckUpSoft);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgCheckUpdateSoft), ff->_ConvertCstringToChars(szval));
	
/*****************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySMStart) + UploadData;
		reg->CreateKey(szCreate);

		int iCheckStart = m_check_startwindows.GetCheck();
		szval.Format(L"%d", iCheckStart);
		reg->WriteChar(ff->_ConvertCstringToChars(StartAppWithWindows),
			ff->_ConvertCstringToChars(szval));

/*****************************************************************************/
	
		szCreate = bb->BaseDecodeEx(CreateKeyRun);
		reg->CreateKey(szCreate);

		if (iCheckStart == 1)
		{
			szval.Format(_T("%s%s.exe"), pathFolder, FileSmartStartApp);
			reg->WriteChar(
				ff->_ConvertCstringToChars(FileSmartStartApp),
				ff->_ConvertCstringToChars(szval));
		}
		else reg->DeleteValue(FileSmartStartApp);
	}
	catch (...) {}
}

void DlgOptionMain::OnBnClickedBtnCancel()
{
	CDialogEx::OnCancel();
}

void DlgOptionMain::_XacdinhSheetLienquan()
{
	bsConfig = ff->_xlGetNameSheet(shConfig, 0);
	psConfig = xl->Sheets->GetItem(bsConfig);
	prConfig = psConfig->Cells;
}

void DlgOptionMain::_SetDefault()
{
	try
	{
		pathFolder = ff->_xlGetValue(psConfig, L"CF_DIR_DEFAULT", 0, 0).Trim();
		m_txt_browse.SetWindowText(pathFolder);

		m_txt_browse.SubclassDlgItem(OPTM_TXT_BROWSE, this);
		m_txt_browse.SetBkColor(rgbWhite);
		m_txt_browse.SetTextColor(rgbBlue);
		m_txt_browse.SetCueBanner(L"Kích chọn đường dẫn thư mục gốc...");
		m_txt_browse.SetWindowText(pathFolder);

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
		reg->CreateKey(szCreate);

		int iAutoUpdate = 1;
		CString szval = reg->ReadString(DlgAutoUpdate, L"");
		if (szval != L"") iAutoUpdate = _wtoi(szval.Trim());
		if (iAutoUpdate != 1 && iAutoUpdate != 2) iAutoUpdate = 0;

		m_cbb_autoupdate.AddString(L"Tự động cập nhật dữ liệu");
		m_cbb_autoupdate.AddString(L"Luôn hỏi trước khi cập nhật");
		m_cbb_autoupdate.AddString(L"Không cập nhật");
		m_cbb_autoupdate.SetCurSel(iAutoUpdate);

		int iDelFileCSV = 0;
		szval = reg->ReadString(DlgDeleteFileCSV, L"");
		if (szval != L"") iDelFileCSV = _wtoi(szval.Trim());
		if (iDelFileCSV != 1) iDelFileCSV = 0;
		m_check_delfilecsv.SetCheck(iDelFileCSV);

/*****************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySMStart) + UploadData;
		reg->CreateKey(szCreate);

		int iCheckStart = 1;
		szval = reg->ReadString(StartAppWithWindows, L"");
		if (szval != L"") iCheckStart = _wtoi(szval.Trim());
		if (iCheckStart != 1) iCheckStart = 0;
		m_check_startwindows.SetCheck(iCheckStart);

/*****************************************************************************/

		szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
		reg->CreateKey(szCreate);

		int iCheckUpSoft = 0;
		szval = reg->ReadString(DlgCheckUpdateSoft, L"");
		if (szval != L"") iCheckUpSoft = _wtoi(szval.Trim());
		if (iCheckUpSoft != 1) iCheckUpSoft = 0;
		m_check_msgupdate.SetCheck(iCheckUpSoft);

/*****************************************************************************/

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_register.SetThemeHelper(&m_ThemeHelper);
		m_btn_register.SetIcon(IDI_ICON_PROPERTIES, 24, 24);
		m_btn_register.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_register.SetBtnCursor(IDC_CURSOR_HAND);
	}
	catch (...) {}
}

void DlgOptionMain::OnBnClickedBtnBrowse()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần lưu...";
	if (m_dlg.DoModal() == IDOK)
	{
		CString strTargetDirectory = m_dlg.GetPathName();
		m_txt_browse.SetWindowText(strTargetDirectory);
	}
}

void DlgOptionMain::OnBnClickedBtnRegisterCom()
{
	try
	{
		Function *ff = new Function;
		ff->_RegistryCOMGpro();
		ff->_CheckRegisterCOM();
		delete ff;
	}
	catch (...) {}
}
