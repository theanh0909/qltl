#include "pch.h"
#include "DlgOptionMain.h"

IMPLEMENT_DYNAMIC(DlgOptionMain, CDialogEx)

DlgOptionMain::DlgOptionMain(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgOptionMain::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
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
	DDX_Control(pDX, OPT_BTN_OK, m_btn_ok);
	DDX_Control(pDX, OPT_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, OPT_CHECK_STARTWINDOWS, m_check_start);

}

BEGIN_MESSAGE_MAP(DlgOptionMain, CDialogEx)
	ON_WM_SYSCOMMAND()

	ON_BN_CLICKED(OPT_BTN_OK, &DlgOptionMain::OnBnClickedBtnOk)
	ON_BN_CLICKED(OPT_BTN_CANCEL, &DlgOptionMain::OnBnClickedBtnCancel)
END_MESSAGE_MAP()

// DlgOptionMain message handlers
BOOL DlgOptionMain::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_DUTOAN));
	SetIcon(hIcon, FALSE);

	_SetDefault();
	_SetLocationAndSize();

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
			if(pWndWithFocus==GetDlgItem(OPT_BTN_OK))
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
	CDialogEx::OnOK();

	CString szCreate = bb->BaseDecodeEx(CreateKeySMStart) + UploadData;
	reg->CreateKey(szCreate);

	CString szval = L"";
	int iCheckStart = m_check_start.GetCheck();
	szval.Format(L"%d", iCheckStart);
	reg->WriteChar(ff->_ConvertCstringToChars(StartAppWithWindows),
		ff->_ConvertCstringToChars(szval));

	szCreate = bb->BaseDecodeEx(CreateKeyRun);
	reg->CreateKey(szCreate);

	if (iCheckStart == 1)
	{
		CString pathFolder = ff->_GetPathFolder(szAppSmart);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";
		szval.Format(_T("%s%s.exe"), pathFolder, FileSmartStartApp);
		reg->WriteChar(
			ff->_ConvertCstringToChars(FileSmartStartApp),
			ff->_ConvertCstringToChars(szval));
	}
	else reg->DeleteValue(FileSmartStartApp);
}


void DlgOptionMain::OnBnClickedBtnCancel()
{
	CDialogEx::OnCancel();
}

void DlgOptionMain::_SetLocationAndSize()
{
	try
	{
		int isx = GetSystemMetrics(SM_CXSCREEN);
		int isy = GetSystemMetrics(SM_CYSCREEN);

		CRect r;
		GetWindowRect(&r);
		ScreenToClient(&r);
		MoveWindow((isx - r.Width()) / 2, (isy - r.Height()) / 2, r.right - r.left, r.bottom - r.top, TRUE);
	}
	catch (...) {}
}

void DlgOptionMain::_SetDefault()
{
	try
	{
		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

/**********************************************************************************/

		CString szCreate = bb->BaseDecodeEx(CreateKeySMStart) + UploadData;
		reg->CreateKey(szCreate);
		
		int iCheckStart = 1;
		CString szval = reg->ReadString(StartAppWithWindows, L"");
		if (szval != L"") iCheckStart = _wtoi(szval.Trim());
		if (iCheckStart != 1) iCheckStart = 0;
		m_check_start.SetCheck(iCheckStart);
	}
	catch (...) {}
}