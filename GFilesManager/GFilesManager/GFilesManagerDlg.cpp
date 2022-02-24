
// GFilesManagerDlg.cpp : implementation file
//

#include "pch.h"

#include "GFilesManager.h"
#include "GFilesManagerDlg.h"
#include "DlgOptionMain.h"
#include "CCheckVersion.h"

#include "MouseMsgHandlerImpl.h"
#include "HookMouse.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CGFilesManagerDlg dialog

CGFilesManagerDlg::CGFilesManagerDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GFILESMANAGER_DIALOG, pParent)
{
	demTimer = 0;
	nVersion = 0;
	strNewVersion = L"";	

	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;

	MyHook::Instance().InstallHook();
}

CGFilesManagerDlg::~CGFilesManagerDlg()
{
	delete ff;
	delete bb;
	delete sv;
	delete reg;

	m_pTrayIcon->HideIcon();
	MyHook::Instance().UninstallHook();	
}

void CGFilesManagerDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, OPMAIN_TAB_MAIN, tab_main);
	DDX_Control(pDX, OPMAIN_BTN_ICON, m_btn_icon);	
	DDX_Control(pDX, OPMAIN_BTN_CLOSE, m_btn_close);
	DDX_Control(pDX, OPMAIN_BTN_OPTIONS, m_btn_options);
	DDX_Control(pDX, OPMAIN_BTN_UPDATE, m_btn_update);
}

BEGIN_MESSAGE_MAP(CGFilesManagerDlg, CDialogEx)
	ON_WM_TIMER()
	ON_WM_SYSCOMMAND()
	ON_COMMAND(ID_TRAYICON_SHOW, &CGFilesManagerDlg::OnTrayiconShow)
	ON_COMMAND(ID_TRAYICON_CLOSE, &CGFilesManagerDlg::OnTrayiconClose)
	ON_BN_CLICKED(OPMAIN_BTN_CLOSE, &CGFilesManagerDlg::OnBnClickedBtnClose)
	ON_BN_CLICKED(OPMAIN_BTN_OPTIONS, &CGFilesManagerDlg::OnBnClickedBtnOptions)
	ON_BN_CLICKED(OPMAIN_BTN_UPDATE, &CGFilesManagerDlg::OnBnClickedBtnUpdate)	
END_MESSAGE_MAP()

// CGFilesManagerDlg message handlers
BOOL CGFilesManagerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(RGBWHITE);

	_SetDefault();
	_ShowIconTask();
	_CreateTrayIconTask();
	_HiddenTrayIconTaskbar();
	_SetLocationAndSize();	

	szCheckRestartFiles = bb->BaseDecodeEx(CreateKeySMStart);

	demTimer = 0;
	SetTimer(CONST_TIMER_ID, 500, NULL);	// 1000ms = 1 second

	return TRUE;
}

void CGFilesManagerDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) _HiddenTrayIconTaskbar();
	else CDialog::OnSysCommand(nID, lParam);
}

BOOL CGFilesManagerDlg::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_ESCAPE)
		{
			_HiddenTrayIconTaskbar();
		}
	}

	return FALSE;
}

void CGFilesManagerDlg::_SetDefault()
{
	try
	{
		CStatic *m_lbl_app = (CStatic*)GetDlgItem(OPMAIN_LBL_NAMEAPP);
		m_lbl_app->SetFont(&m_FontHeader);

		int iTabDefault = 0;
		tab_main.InsertItem(0, L"Bắt đầu ");
		tab_main.InsertItem(1, L"Việc cần làm ");
		tab_main.InsertItem(2, L"Thông báo ");

		tab_main.SetCurSel(iTabDefault);
		tab_main.Init(iTabDefault);

		m_btn_icon.SetThemeHelper(&m_ThemeHelper);
		m_btn_icon.SetIcon(IDI_ICON_DUTOAN, 16, 16);
		m_btn_icon.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_icon.SetBtnCursor(IDC_CURSOR_HAND);
		m_btn_icon.DrawTransparent(false);

		m_btn_close.SetThemeHelper(&m_ThemeHelper);
		m_btn_close.SetIcon(IDI_ICON_CLOSE2, 16, 16, IDI_ICON_CLOSE, 16, 16);
		m_btn_close.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_close.SetBtnCursor(IDC_CURSOR_HAND);
		m_btn_close.DrawTransparent(false);

		m_btn_options.SetThemeHelper(&m_ThemeHelper);
		m_btn_options.SetIcon(IDI_ICON_GEAR2, 16, 16, IDI_ICON_GEAR, 16, 16);
		m_btn_options.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_options.SetBtnCursor(IDC_CURSOR_HAND);
		m_btn_options.DrawTransparent(false);

		m_btn_update.SetFont(&m_FontHeader);
		m_btn_update.SetFaceColor(RGBNAVYBLUE);
		m_btn_update.SetTextColor(RGBWHITE);
		m_btn_update.ShowWindow(0);

		// Thiết lập trong Properties: Cursor Type = Hand + Style = No Border
		//HCURSOR hcrHand = LoadCursor(AfxGetResourceHandle(), MAKEINTRESOURCE(IDC_CURSOR_HAND));
		//m_btn_update.SetMouseCursor(hcrHand);
	}
	catch (...) {}
}

void CGFilesManagerDlg::_CreateTrayIconTask()
{
	try
	{
		MouseMsgHandlerPtr *pHandler = new MouseMsgHandlerPtr[3];
		MouseMsgHandlerPtr rbcHandler = new CRightMouseClickMsgHandler();
		MouseMsgHandlerPtr lbcHandler = new CLeftMouseDblClickMsgHandler();
		MouseMsgHandlerPtr mhHandler = new CMouseHoverMsgHandler(m_pTrayIcon);

		pHandler[0] = lbcHandler;
		pHandler[1] = rbcHandler;
		pHandler[2] = mhHandler;

		IconDataPtr *pIconData = new IconDataPtr[2];
		pIconData[0] = new CIconData(IDR_MAINFRAME, (LPSTR)TitleNameApp, IDR_MAINFRAME);
		pIconData[1] = new CIconData(IDR_MAINFRAME, (LPSTR)TitleNameApp, IDR_MAINFRAME);
		m_pTrayIcon = new CTrayIcon(pHandler, 3, pIconData, 2, 0, SECOND);

		((CMouseHoverMsgHandler*)mhHandler)->SetTrayIcon(m_pTrayIcon);
	}
	catch (...) {}
}

void CGFilesManagerDlg::OnTrayiconShow()
{
	if (!AfxGetMainWnd()->IsWindowVisible()) ShowWnd(this);
	else HideWnd(this);	
}

void CGFilesManagerDlg::OnTrayiconClose()
{
	CDialogEx::OnCancel();
}

void CGFilesManagerDlg::_HiddenTrayIconTaskbar()
{
	m_pTrayIcon->ShowIcon();
	HideWnd(this);
}

void CGFilesManagerDlg::_ShowIconTask(bool flag)
{
	static CRect TmpRect;
	if (flag)
	{
		ModifyStyleEx(WS_EX_TOOLWINDOW, WS_EX_APPWINDOW);
		AfxGetApp()->GetMainWnd()->ShowWindow(SW_SHOW);
		AfxGetMainWnd()->MoveWindow(TmpRect.left, TmpRect.top, TmpRect.Width(), TmpRect.Height(), true);
	}
	else
	{
		GetWindowRect(&TmpRect);
		ModifyStyleEx(WS_EX_APPWINDOW, WS_EX_TOOLWINDOW);
		AfxGetApp()->GetMainWnd()->ShowWindow(SW_HIDE);
		AfxGetMainWnd()->MoveWindow(-TmpRect.right, -TmpRect.bottom, TmpRect.Width(), TmpRect.Height(), true);
	}
}

void CGFilesManagerDlg::_SetLocationAndSize()
{
	try
	{
		CRect crt;
		GetWindowRect(&crt);
		ScreenToClient(&crt);

		dWidthF = crt.Width();
		dHeightF = crt.Height();

		int iSx = GetSystemMetrics(SM_CXSCREEN);
		int iSy = GetSystemMetrics(SM_CYSCREEN);

		int iTaskHeight = 0;
		int nPostion = IsTaskbarWndVisible(iTaskHeight);
		switch (nPostion)
		{
			case -1:
			{
				// Task --> Hidden
				dPosX = iSx - dWidthF - 5;
				dPosY = iSy - dHeightF - 5;
			}
			break;

			case 0:
			{
				// Task --> Pane Bottom
				dPosX = iSx - dWidthF - 5;
				dPosY = iTaskHeight - dHeightF - 5;
			}
			break;

			case 1:
			{
				// Task --> Pane Left
				dPosX = iTaskHeight + 5;
				dPosY = iSy - dHeightF - 5;
			}
			break;

			case 2:
			{
				// Task --> Pane Top
				dPosX = iSx - dWidthF - 5;
				dPosY = iTaskHeight + 5;
			}
			break;

			case 3:
			{
				// Task --> Pane Right
				dPosX = iTaskHeight - dWidthF - 5;
				dPosY = iSy - dHeightF - 5;
			}
			break;

			default:
				break;
		}

		this->MoveWindow(dPosX, dPosY, crt.right - crt.left, crt.bottom - crt.top, TRUE);
	}
	catch (...) {}
}

void CGFilesManagerDlg::OnBnClickedBtnUpdate()
{
	int result = ff->_y(L"Bạn có chắc chắn cập nhật phiên bản mới?", 0, 0, 0);
	if (result == 6)
	{
		CString pathFolder = ff->_GetPathFolder(szAppSmart);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		CString szWin = L"32\\";
		if (ff->_Is64BitWindows()) szWin = L"64\\";

		sv->_LoadDecodeBase64();
		CString szFileApp = pathFolder + FileUpdatesm;
		if (sv->_DownloadFile(sv->dbSeverDir + szWin + FileUpdatesm, szFileApp))
		{
			// Mở file 'updatess.exe' để chạy cập nhật
			ShellExecute(NULL, L"open", szFileApp, NULL, NULL, SW_SHOWMAXIMIZED);
		}
	}

	AfxGetMainWnd()->ShowWindow(SW_SHOW);
	AfxGetMainWnd()->ActivateTopParent();
}

void CGFilesManagerDlg::OnBnClickedBtnClose()
{
	_HiddenTrayIconTaskbar();
}

void CGFilesManagerDlg::OnBnClickedBtnOptions()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	DlgOptionMain p;
	p.DoModal();

	AfxGetMainWnd()->ShowWindow(SW_SHOW);
	AfxGetMainWnd()->ActivateTopParent();
}

void CGFilesManagerDlg::OnTimer(UINT_PTR nIDEvent)
{
	// Chú ý: Giá trị 'CheckAddFiles' --> Liên hệ với DLL sl (GMSelect.exe)
	reg->CreateKey(szCheckRestartFiles);
	int iCheckRestart = _wtoi(reg->ReadString(L"CheckRestartFiles", L""));
	if (iCheckRestart != 0)
	{
		reg->WriteChar("CheckRestartFiles", "0");
		OnTrayiconClose();
	}

	demTimer++;

	// 60 phút kiểm tra update 1 lần
	if (demTimer % 7200 == 1)
	{
		strNewVersion = L"";
		CCheckVersion *ck = new CCheckVersion;
		nVersion = ck->_CompareVersion(strNewVersion);
		delete ck;

		if (strNewVersion != L"")
		{
			m_btn_update.ShowWindow(nVersion);
		}
	}

	if (nVersion == 1)
	{
		if (demTimer % 2 == 0)
		{
			m_btn_update.ShowWindow(1);
			m_btn_update.SetFaceColor(RGB(192, 0, 0));
		}
		else
		{
			m_btn_update.ShowWindow(0);
		}
	}

	if (demTimer == 7200) demTimer = 0;

	CDialogEx::OnTimer(nIDEvent);
}

