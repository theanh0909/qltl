#include "pch.h"
#include "DlgInsertRows.h"

IMPLEMENT_DYNAMIC(DlgInsertRows, CDialogEx)

DlgInsertRows::DlgInsertRows(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgInsertRows::IDD, pParent)
{
	numInsert = 0;
	numMax = 9999;
}

DlgInsertRows::~DlgInsertRows()
{

}

void DlgInsertRows::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, ISTROW_TXT_NUM, m_txt_num);
	DDX_Control(pDX, ISTROW_SPIN_NUM, m_spin_num);
	DDX_Control(pDX, ISTROW_BTN_OK, m_btn_ok);
	DDX_Control(pDX, ISTROW_BTN_CANCEL, m_btn_cancel);
}

BEGIN_MESSAGE_MAP(DlgInsertRows, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(ISTROW_BTN_OK, &DlgInsertRows::OnBnClickedBtnOk)
	ON_BN_CLICKED(ISTROW_BTN_CANCEL, &DlgInsertRows::OnBnClickedBtnCancel)
	ON_NOTIFY(UDN_DELTAPOS, ISTROW_SPIN_NUM, &DlgInsertRows::OnDeltaposSpinNum)
	ON_EN_CHANGE(ISTROW_TXT_NUM, &DlgInsertRows::OnEnChangeTxtNum)
END_MESSAGE_MAP()

// DlgInsertRows message handlers
BOOL DlgInsertRows::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_ADD));
	SetIcon(hIcon, FALSE);

	_SetDefault();

	return TRUE;
}

BOOL DlgInsertRows::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(ISTROW_BTN_OK) 
				|| pWndWithFocus == GetDlgItem(ISTROW_TXT_NUM))
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
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(ISTROW_TXT_NUM))
		{
			if (chr != 0x30 && chr != 0x31 && chr != 0x32 && chr != 0x33 && chr != 0x34 &&
				chr != 0x35 && chr != 0x36 && chr != 0x37 && chr != 0x38 && chr != 0x39)
			{
				m_txt_num.ShowBalloonTip(L"Hướng dẫn",
					L"Vui lòng nhập giá trị là số [0-9]", TTI_WARNING);
				return TRUE;
			}
		}
	}

	return FALSE;
}

void DlgInsertRows::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgInsertRows::OnBnClickedBtnOk()
{
	CString szval = L"";
	m_txt_num.GetWindowTextW(szval);
	numInsert = _wtoi(szval);
	if (numInsert > numMax) numInsert = numMax;

	CDialog::OnOK();
}

void DlgInsertRows::OnBnClickedBtnCancel()
{
	numInsert = 0;
	CDialogEx::OnCancel();
}


void DlgInsertRows::OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult)
{
	GotoDlgCtrl(GetDlgItem(ISTROW_TXT_NUM));
}


void DlgInsertRows::OnEnChangeTxtNum()
{
	CString szval = L"";
	m_txt_num.GetWindowTextW(szval);
	numInsert = _wtoi(szval);
	if (numInsert > numMax)
	{
		numInsert = numMax;
		szval.Format(L"%d", numMax);
		m_txt_num.SetWindowText(szval);
		m_txt_num.SetSel(-1);
	}
}

void DlgInsertRows::_SetDefault()
{
	m_txt_num.SubclassDlgItem(ISTROW_TXT_NUM, this);
	m_txt_num.SetBkColor(rgbWhite);
	m_txt_num.SetTextColor(rgbBlue);

	m_spin_num.SetDecimalPlaces(0);
	m_spin_num.SetTrimTrailingZeros(FALSE);
	m_spin_num.SetRangeAndDelta(1, numMax, 10);
	m_spin_num.SetPos(10);

	m_btn_ok.SetThemeHelper(&m_ThemeHelper);
	m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);
}
