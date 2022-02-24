#include "pch.h"
#include "GXDUpdateInfo.h"
#include "DlgUpdateInfo.h"

#define RGBWHITE		RGB(255, 255, 255)
#define RGBBLACK		RGB(0, 0, 0)
#define RGBRED			RGB(255, 0, 0)
#define RGBBLUE			RGB(0, 0, 255)
#define RGBYELLOW		RGB(255, 255, 240)

IMPLEMENT_DYNAMIC(DlgUpdateInfo, CDialogEx)

DlgUpdateInfo::DlgUpdateInfo(CWnd* pParent /*=nullptr*/)
	: CDialogEx(DLG_UPDATEINFO, pParent)
{

}

DlgUpdateInfo::~DlgUpdateInfo()
{
}

void DlgUpdateInfo::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, UPINFO_LBL_UPDATE, m_lbl_update);
	DDX_Control(pDX, UPINFO_LBL_NAME, m_lbl_name);
	DDX_Control(pDX, UPINFO_LBL_EMAIL, m_lbl_email);
	DDX_Control(pDX, UPINFO_LBL_MOBILE, m_lbl_mobile);
	DDX_Control(pDX, UPINFO_LBL_ADDRESS, m_lbl_address);
	
	DDX_Control(pDX, UPINFO_LBLCHECK_NAME, m_lblchk_name);
	DDX_Control(pDX, UPINFO_LBLCHECK_EMAIL, m_lblchk_email);
	DDX_Control(pDX, UPINFO_LBLCHECK_MOBILE, m_lblchk_mobile);

	DDX_Control(pDX, UPINFO_TXT_NAME, m_txt_name);
	DDX_Control(pDX, UPINFO_TXT_EMAIL, m_txt_email);
	DDX_Control(pDX, UPINFO_TXT_MOBILE, m_txt_mobile);
	DDX_Control(pDX, UPINFO_TXT_ADDRESS, m_txt_address);

	DDX_Control(pDX, UPINFO_BTN_LOGO, m_btn_logo);
	DDX_Control(pDX, UPINFO_BTN_UPDATE, m_btn_update);
}


BEGIN_MESSAGE_MAP(DlgUpdateInfo, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(UPINFO_BTN_UPDATE, &DlgUpdateInfo::OnBnClickedBtnUpdate)
	ON_EN_KILLFOCUS(UPINFO_TXT_NAME, &DlgUpdateInfo::OnEnKillfocusTxtName)
	ON_EN_KILLFOCUS(UPINFO_TXT_EMAIL, &DlgUpdateInfo::OnEnKillfocusTxtEmail)
	ON_EN_KILLFOCUS(UPINFO_TXT_MOBILE, &DlgUpdateInfo::OnEnKillfocusTxtMobile)
	ON_EN_KILLFOCUS(UPINFO_TXT_ADDRESS, &DlgUpdateInfo::OnEnKillfocusTxtAddress)
	ON_EN_SETFOCUS(UPINFO_TXT_NAME, &DlgUpdateInfo::OnEnSetfocusTxtName)
	ON_EN_SETFOCUS(UPINFO_TXT_EMAIL, &DlgUpdateInfo::OnEnSetfocusTxtEmail)
	ON_EN_SETFOCUS(UPINFO_TXT_MOBILE, &DlgUpdateInfo::OnEnSetfocusTxtMobile)
	ON_EN_SETFOCUS(UPINFO_TXT_ADDRESS, &DlgUpdateInfo::OnEnSetfocusTxtAddress)	
	ON_NOTIFY(NM_CLICK, UPINFO_HYPERLINK, &DlgUpdateInfo::OnNMClickHyperlink)
END_MESSAGE_MAP()


// DlgUpdateInfo message handlers
BOOL DlgUpdateInfo::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetControlDefault();

	return TRUE;
}

BOOL DlgUpdateInfo::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN || iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(UPINFO_BTN_UPDATE))
			{
				if (iWPar == VK_RETURN)
				{
					OnBnClickedBtnUpdate();
					return TRUE;
				}
			}
			else if (pWndWithFocus == GetDlgItem(UPINFO_TXT_NAME))
			{
				GotoDlgCtrl(GetDlgItem(UPINFO_TXT_EMAIL));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(UPINFO_TXT_EMAIL))
			{
				GotoDlgCtrl(GetDlgItem(UPINFO_TXT_MOBILE));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(UPINFO_TXT_MOBILE))
			{
				GotoDlgCtrl(GetDlgItem(UPINFO_TXT_ADDRESS));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(UPINFO_TXT_ADDRESS))
			{
				GotoDlgCtrl(GetDlgItem(UPINFO_TXT_NAME));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(UPINFO_TXT_MOBILE))
		{
			if (!((chr >= 0x30 && chr <= 0x39)
				|| chr == 0x08 || chr == 0x1b || chr == 0x7f))
			{
				m_txt_mobile.ShowBalloonTip(L"Hướng dẫn",
					L"Vui lòng nhập giá trị là số [0-9]", TTI_WARNING);
				return TRUE;
			}
		}
	}

	return FALSE;
}


void DlgUpdateInfo::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) return;
	else CDialog::OnSysCommand(nID, lParam);
}


void DlgUpdateInfo::OnBnClickedBtnUpdate()
{
	try
	{
		if (!CheckValueName())
		{
			ff->_msgbox(L"Bạn chưa nhập họ và tên.", 2, 0);
			return;
		}

		if (!CheckValueEmail())
		{
			ff->_msgbox(
				L"Địa chỉ Email không được để trống và phải đúng định dạng. "
				L"Ví dụ: nguyenvannam@gmail.com.", 2, 0);
			return;
		}

		if (!CheckValueMobile())
		{
			ff->_msgbox(L"Vui lòng kiểm tra số điện thoại đã nhập.", 2, 0);
			return;
		}

		CDialogEx::OnOK();

		/***********************************************************************************************/

		CString szval = L"";
		wchar_t* strName = new wchar_t[255];		
		m_txt_name.GetWindowTextW(szval);
		szval.Trim();
		wcscpy(strName, szval);

		wchar_t* strEmail = new wchar_t[255];
		m_txt_email.GetWindowTextW(szval);
		szval.Trim();
		wcscpy(strEmail, szval);

		wchar_t* strMobile = new wchar_t[255];
		m_txt_mobile.GetWindowTextW(szval);
		szval.Trim();
		wcscpy(strMobile, szval);

		wchar_t* strAddress = new wchar_t[255];
		m_txt_address.GetWindowTextW(szval);
		szval.Trim();
		wcscpy(strAddress, szval);

		typedef bool(__stdcall *p)(
			_variant_t szValName, 
			_variant_t szValAdress,
			_variant_t szValEmail,
			_variant_t szValPhone);
		HINSTANCE loadDLL = LoadLibrary((_bstr_t)L"QLDA.dll");

		if (loadDLL != NULL)
		{
			p pc = (p)GetProcAddress(loadDLL, (_bstr_t)L"UpdateInfoAppNoDialog");
			if (pc != NULL)
			{
				ff->_msgbox(L"Call UpdateInfoAppNoDialog", 0, 0);
				pc(L"AAAAA", L"a@b.com", L"0123455667", L"null");
			}
		}

		FreeLibrary(loadDLL);

		
	}
	catch (...) {}	
}


void DlgUpdateInfo::SetControlDefault()
{
	try
	{
		int nsizeIcon = 64 * jScaleLayout;
		m_btn_logo.SetThemeHelper(&m_ThemeHelper);
		m_btn_logo.SetIcon(IDI_ICON_LOGO, nsizeIcon, nsizeIcon);
		m_btn_logo.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_logo.DrawTransparent();
		m_btn_logo.DrawBorder(FALSE);

		nsizeIcon = 24 * jScaleLayout;
		m_btn_update.SetThemeHelper(&m_ThemeHelper);
		m_btn_update.SetIcon(IDI_ICON_OK, nsizeIcon, nsizeIcon);
		m_btn_update.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);

		m_lbl_update.SubclassDlgItem(UPINFO_LBL_UPDATE, this);
		m_lbl_update.SetTextColor(RGBRED);
		m_lbl_update.SetFont(&m_FontTitleLarge);

		m_lbl_name.SetFont(&m_FontTiteForm);
		m_lbl_email.SetFont(&m_FontTiteForm);
		m_lbl_mobile.SetFont(&m_FontTiteForm);
		m_lbl_address.SetFont(&m_FontTiteForm);		

		m_txt_name.SubclassDlgItem(UPINFO_TXT_NAME, this);
		m_txt_name.SetBkColor(RGBWHITE);
		m_txt_name.SetTextColor(RGBBLUE);
		m_txt_name.SetCueBanner(L"Nguyễn Văn Nam");

		m_lblchk_name.SubclassDlgItem(UPINFO_LBLCHECK_NAME, this);
		m_lblchk_name.SetTextColor(RGBRED);
		m_lblchk_name.SetFont(&m_FontTiteError);

		m_txt_email.SubclassDlgItem(UPINFO_TXT_EMAIL, this);
		m_txt_email.SetBkColor(RGBWHITE);
		m_txt_email.SetTextColor(RGBBLUE);
		m_txt_email.SetCueBanner(L"nguyenvannam@gmail.com");

		m_lblchk_email.SubclassDlgItem(UPINFO_LBLCHECK_EMAIL, this);
		m_lblchk_email.SetTextColor(RGBRED);
		m_lblchk_email.SetFont(&m_FontTiteError);

		m_txt_mobile.SubclassDlgItem(UPINFO_TXT_MOBILE, this);
		m_txt_mobile.SetBkColor(RGBWHITE);
		m_txt_mobile.SetTextColor(RGBBLUE);
		m_txt_mobile.SetCueBanner(L"0984.555.999");

		m_lblchk_mobile.SubclassDlgItem(UPINFO_LBLCHECK_MOBILE, this);
		m_lblchk_mobile.SetTextColor(RGBRED);
		m_lblchk_mobile.SetFont(&m_FontTiteError);

		m_txt_address.SubclassDlgItem(UPINFO_TXT_ADDRESS, this);
		m_txt_address.SetBkColor(RGBWHITE);
		m_txt_address.SetTextColor(RGBBLUE);
		m_txt_address.SetCueBanner(L"Số 124 Nguyễn Ngọc Nại, P.Khương Mai, Q.Thanh Xuân, Hà Nội");		
	}
	catch (...) {}
}

void DlgUpdateInfo::OnEnKillfocusTxtName()
{
	m_txt_name.SubclassDlgItem(UPINFO_TXT_NAME, this);
	m_txt_name.SetBkColor(RGBWHITE);
	m_txt_name.SetTextColor(RGBBLUE);

	if (CheckValueName()) m_lblchk_name.ShowWindow(0);
	else m_lblchk_name.ShowWindow(1);

	TextAuto();
}


void DlgUpdateInfo::OnEnKillfocusTxtEmail()
{
	m_txt_email.SubclassDlgItem(UPINFO_TXT_EMAIL, this);
	m_txt_email.SetBkColor(RGBWHITE);
	m_txt_email.SetTextColor(RGBBLUE);

	if (CheckValueEmail()) m_lblchk_email.ShowWindow(0);
	else m_lblchk_email.ShowWindow(1);
}


void DlgUpdateInfo::OnEnKillfocusTxtMobile()
{
	m_txt_mobile.SubclassDlgItem(UPINFO_TXT_MOBILE, this);
	m_txt_mobile.SetBkColor(RGBWHITE);
	m_txt_mobile.SetTextColor(RGBBLUE);

	if (CheckValueMobile()) m_lblchk_mobile.ShowWindow(0);
	else m_lblchk_mobile.ShowWindow(1);
}


void DlgUpdateInfo::OnEnKillfocusTxtAddress()
{
	m_txt_address.SubclassDlgItem(UPINFO_TXT_ADDRESS, this);
	m_txt_address.SetBkColor(RGBWHITE);
	m_txt_address.SetTextColor(RGBBLUE);

	CString szval = L"";
	m_txt_address.GetWindowTextW(szval);

	if (szval.GetLength() > 255)
	{
		szval = szval.Left(255);
		m_txt_address.SetWindowText(szval);
	}
}


void DlgUpdateInfo::OnEnSetfocusTxtName()
{
	m_txt_name.SubclassDlgItem(UPINFO_TXT_NAME, this);
	m_txt_name.SetBkColor(RGBYELLOW);
	m_txt_name.SetTextColor(RGBBLUE);
}


void DlgUpdateInfo::OnEnSetfocusTxtEmail()
{
	m_txt_email.SubclassDlgItem(UPINFO_TXT_EMAIL, this);
	m_txt_email.SetBkColor(RGBYELLOW);
	m_txt_email.SetTextColor(RGBBLUE);
}


void DlgUpdateInfo::OnEnSetfocusTxtMobile()
{
	m_txt_mobile.SubclassDlgItem(UPINFO_TXT_MOBILE, this);
	m_txt_mobile.SetBkColor(RGBYELLOW);
	m_txt_mobile.SetTextColor(RGBBLUE);
}


void DlgUpdateInfo::OnEnSetfocusTxtAddress()
{
	m_txt_address.SubclassDlgItem(UPINFO_TXT_ADDRESS, this);
	m_txt_address.SetBkColor(RGBYELLOW);
	m_txt_address.SetTextColor(RGBBLUE);
}


bool DlgUpdateInfo::CheckValueName()
{
	try
	{
		CString szval = L"";
		m_txt_name.GetWindowTextW(szval);
		szval.Trim();

		if (szval != L"")
		{
			if (szval.GetLength() > 100)
			{
				szval = szval.Left(100);
				m_txt_name.SetWindowText(szval);
			}

			return true;
		}
	}
	catch (...) {}
	return false;
}


bool DlgUpdateInfo::CheckValueEmail()
{
	try
	{
		CString szval = L"";
		m_txt_email.GetWindowTextW(szval);
		szval.Trim();
		if (szval != L"")
		{
			if (szval.GetLength() > 50)
			{
				szval = szval.Left(50);
				m_txt_email.SetWindowText(szval);
			}

			const regex pattern("(\\w+)(\\.|_)?(\\w*)@(\\w+)(\\.(\\w+))+");
			wchar_t* wszString = new wchar_t[szval.GetLength() + 1];
			wcscpy(wszString, szval);
			int num = ::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), NULL, 0, NULL, NULL);
			char* utf8 = new char[num + 1];
			::WideCharToMultiByte(CP_UTF8, NULL, wszString, wcslen(wszString), utf8, num, NULL, NULL);
			utf8[num] = '\0';
			string strEmail(utf8);
			bool blResult = regex_match(strEmail, pattern);

			delete[] wszString;
			delete[] utf8;

			return blResult;
		}
	}
	catch (...) {}
	return false;
}


bool DlgUpdateInfo::CheckValueMobile()
{
	try
	{
		CString szval = L"";
		m_txt_mobile.GetWindowTextW(szval);
		szval.Replace(L" ", L"");

		if (szval != L"")
		{
			if (szval.GetLength() > 20)
			{
				szval = szval.Left(20);
				m_txt_mobile.SetWindowText(szval);
			}

			CString str = L"";
			int nlen = szval.GetLength();
			for (int i = 0; i < nlen; i++)
			{
				str = szval.Mid(i, 1);
				if (!(str == L"0" || _wtoi(str) > 0)) return false;
			}

			return true;
		}
	}
	catch (...) {}
	return false;
}


void DlgUpdateInfo::OnNMClickHyperlink(NMHDR *pNMHDR, LRESULT *pResult)
{
	PNMLINK pNM_Link = (PNMLINK)pNMHDR;
	::ShellExecute(m_hWnd, _T("open"), pNM_Link->item.szUrl, NULL, NULL, SW_SHOWNORMAL);
	*pResult = 0;
}


void DlgUpdateInfo::TextAuto()
{
	try
	{
		CString szval = L"";
		m_txt_name.GetWindowTextW(szval);
		if (szval == L"leogxd")
		{
			m_txt_name.SetWindowText(L"Nguyễn Tiến Thành");
			m_txt_email.SetWindowText(L"tienthanh@giaxaydung.com");
			m_txt_mobile.SetWindowText(L"0978508551");
			m_txt_address.SetWindowText(L"Số 124 Nguyễn Ngọc Nại, P.Khương Mai, Q.Thanh Xuân, Hà Nội");
			GotoDlgCtrl(GetDlgItem(UPINFO_BTN_UPDATE));
		}
	}
	catch (...) {}
}