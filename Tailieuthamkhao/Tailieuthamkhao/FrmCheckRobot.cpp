#include "stdafx.h"
#include "Function.h"
#include "FrmCheckRobot.h"

// FrmCheckRobot dialog
IMPLEMENT_DYNAMIC(FrmCheckRobot, CDialogEx)

FrmCheckRobot::FrmCheckRobot(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DLG_CHECKBOT, pParent)
{
	iCheckpass = 0;
	iChangetxt = 0;
	szChNum = _T("0123456789");
}

FrmCheckRobot::~FrmCheckRobot()
{
}

void FrmCheckRobot::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATIC_BOT_KEY, m_lbl_key);
	DDX_Control(pDX, IDC_EDIT_CODE_TXT_1, m_txt_key[0]);
	DDX_Control(pDX, IDC_EDIT_CODE_TXT_2, m_txt_key[1]);
	DDX_Control(pDX, IDC_EDIT_CODE_TXT_3, m_txt_key[2]);
	DDX_Control(pDX, IDC_EDIT_CODE_TXT_4, m_txt_key[3]);
	DDX_Control(pDX, IDC_EDIT_CODE_TXT_5, m_txt_key[4]);
	DDX_Control(pDX, IDC_BUTTON_CODE_RESET, m_btn_reset);
	DDX_Control(pDX, IDC_BUTTON_CODE_OK, m_btn_ok);
}

BEGIN_MESSAGE_MAP(FrmCheckRobot, CDialogEx)
	ON_WM_TIMER()
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(IDC_BUTTON_CODE_OK, &FrmCheckRobot::OnBnClickedButtonCodeOk)
	ON_BN_CLICKED(IDC_BUTTON_CODE_RESET, &FrmCheckRobot::OnBnClickedButtonCodeReset)

	ON_EN_KILLFOCUS(IDC_EDIT_CODE_TXT_1, &FrmCheckRobot::OnEnKillfocusEditCodeTxt1)
	ON_EN_KILLFOCUS(IDC_EDIT_CODE_TXT_2, &FrmCheckRobot::OnEnKillfocusEditCodeTxt2)
	ON_EN_KILLFOCUS(IDC_EDIT_CODE_TXT_3, &FrmCheckRobot::OnEnKillfocusEditCodeTxt3)
	ON_EN_KILLFOCUS(IDC_EDIT_CODE_TXT_4, &FrmCheckRobot::OnEnKillfocusEditCodeTxt4)
	ON_EN_KILLFOCUS(IDC_EDIT_CODE_TXT_5, &FrmCheckRobot::OnEnKillfocusEditCodeTxt5)
	ON_EN_CHANGE(IDC_EDIT_CODE_TXT_1, &FrmCheckRobot::OnEnChangeEditCodeTxt1)
	ON_EN_CHANGE(IDC_EDIT_CODE_TXT_2, &FrmCheckRobot::OnEnChangeEditCodeTxt2)
	ON_EN_CHANGE(IDC_EDIT_CODE_TXT_3, &FrmCheckRobot::OnEnChangeEditCodeTxt3)
	ON_EN_CHANGE(IDC_EDIT_CODE_TXT_4, &FrmCheckRobot::OnEnChangeEditCodeTxt4)
	ON_EN_CHANGE(IDC_EDIT_CODE_TXT_5, &FrmCheckRobot::OnEnChangeEditCodeTxt5)
END_MESSAGE_MAP()


// FrmCheckRobot message handlers
BOOL FrmCheckRobot::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_ROBOT));
	SetIcon(hIcon, FALSE);
	SetBackgroundColor(RGBBlueOffice);

	for (int i = 0; i < 5; i++)
	{
		m_txt_key[i].SetBkColor(RGBWhite);
		m_txt_key[i].SetTextColor(RGBBlue);
		m_txt_key[i].SetFont(&myfontCheckbot, TRUE);
	}

	short shBtnColor = 30;
	int isize = 24 * jScaleLayout;	
	m_btn_reset.SetIcon(IDI_ICON_RESET, isize, isize);
	m_btn_reset.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);

	m_btn_ok.SetIcon(IDI_ICON_CHECK, isize, isize);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, shBtnColor);

	iSecmax = 30, demTimer = 0;
	SetTimer(1, 1000, NULL);	// 1000ms = 1 second

	CString szval = L"";
	szval.Format(L"%s%d giây)", L"Thời gian còn lại (", iSecmax);
	this->SetWindowText(szval);

	m_lbl_key.SetWindowText(_GetRandowText());

	return TRUE;
}

BOOL FrmCheckRobot::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN || iWPar == VK_TAB || iWPar == VK_RIGHT)
		{
			if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_1))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_2));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_2))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_3));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_3))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_4));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_4))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_5));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_5))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_1));
				return TRUE;
			}
		}
		else if (iWPar == VK_LEFT)
		{
			if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_1))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_5));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_2))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_1));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_3))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_2));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_4))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_3));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_5))
			{
				GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_4));
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
		if (chr != 0x08 && chr != 0x30 && chr != 0x31 && chr != 0x32 && chr != 0x33 && chr != 0x34 &&
			chr != 0x35 && chr != 0x36 && chr != 0x37 && chr != 0x38 && chr != 0x39)
		{
			if (pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_1)
				|| pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_2)
				|| pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_3)
				|| pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_4)
				|| pWndWithFocus == GetDlgItem(IDC_EDIT_CODE_TXT_5))
			{
				return TRUE;
			}
		}
	}

	return FALSE;
}

void FrmCheckRobot::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void FrmCheckRobot::OnBnClickedButtonCodeOk()
{
	CString szval = L"";
	CString szCheck = L"";
	for (int i = 0; i < 5; i++)
	{
		m_txt_key[i].GetWindowTextW(szval);
		szCheck += szval;
	}

	m_lbl_key.GetWindowTextW(szval);
	int pos = szval.Find(L":");
	szval = szval.Right(szval.GetLength() - pos - 1);
	szval.Trim();

	if (szCheck != szval)
	{
		iCheckpass++;
		if (iCheckpass >= 5) OnBnClickedBtnCancel();

		int ivt = 0;
		for (int i = 0; i < 5; i++)
		{
			if (szCheck.Mid(i, 1) != szval.Mid(i, 1))
			{
				ivt = i;
				break;
			}
		}

		switch (ivt)
		{
		case 0:
			GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_1));
			break;

		case 1:
			GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_2));
			break;

		case 2:
			GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_3));
			break;

		case 3:
			GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_4));
			break;

		case 4:
			GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_5));
			break;

		default:
			break;
		}
	}
	else
	{
		jCheckRotot = 1;
		CDialogEx::OnOK();
	}
}


void FrmCheckRobot::OnBnClickedButtonCodeReset()
{
	if (iChangetxt <= 5)
	{
		m_lbl_key.SetWindowText(_GetRandowText());
		GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_1));
		iChangetxt++;
	}
}

void FrmCheckRobot::OnBnClickedBtnCancel()
{
	CDialogEx::OnCancel();
}

CString FrmCheckRobot::_GetRandowText()
{
	Function ff;
	int iRand = ff._RandomTime();
	CString szval = L"Nhập chuỗi giá trị sau: ";
	for (int i = 0; i < 5; i++)
	{
		szval += szChNum.Mid((iRand + rand()) % 10, 1);
	}

	return szval;
}


void FrmCheckRobot::OnTimer(UINT_PTR nIDEvent)
{
	demTimer++;
	if (demTimer >= iSecmax) OnBnClickedBtnCancel();

	CString szval = L"";
	szval.Format(L"%s%d giây)", L"Thời gian còn lại (", iSecmax - demTimer);
	this->SetWindowText(szval);

	CDialogEx::OnTimer(nIDEvent);
}


void FrmCheckRobot::KillfocusEditCode(int nIndex)
{
	CString szval = L"";
	m_txt_key[nIndex].GetWindowTextW(szval);
	if (szval.GetLength() >= 2) m_txt_key[nIndex].SetWindowText(szval.Left(1));
}

void FrmCheckRobot::OnEnKillfocusEditCodeTxt1()
{
	KillfocusEditCode(0);
}


void FrmCheckRobot::OnEnKillfocusEditCodeTxt2()
{
	KillfocusEditCode(1);
}


void FrmCheckRobot::OnEnKillfocusEditCodeTxt3()
{
	KillfocusEditCode(2);
}


void FrmCheckRobot::OnEnKillfocusEditCodeTxt4()
{
	KillfocusEditCode(3);
}


void FrmCheckRobot::OnEnKillfocusEditCodeTxt5()
{
	KillfocusEditCode(4);
}


void FrmCheckRobot::OnEnChangeEditCodeTxt1()
{
	GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_2));
}


void FrmCheckRobot::OnEnChangeEditCodeTxt2()
{
	GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_3));
}


void FrmCheckRobot::OnEnChangeEditCodeTxt3()
{
	GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_4));
}


void FrmCheckRobot::OnEnChangeEditCodeTxt4()
{
	GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_5));
}


void FrmCheckRobot::OnEnChangeEditCodeTxt5()
{
	GotoDlgCtrl(GetDlgItem(IDC_EDIT_CODE_TXT_1));
}
