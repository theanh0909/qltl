#include "pch.h"
#include "FrmTabSTT.h"

// FrmTabSTT dialog
IMPLEMENT_DYNAMIC(FrmTabSTT, CDialogEx)

FrmTabSTT::FrmTabSTT(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FVIEW_RENAME_STT, pParent)
{

}

FrmTabSTT::~FrmTabSTT()
{
}

void FrmTabSTT::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TABSTT_TXT_NUM, m_txt_num);
	DDX_Control(pDX, TABSTT_TXT_TIENTO, m_txt_tiento);
	DDX_Control(pDX, TABSTT_TXT_HAUTO, m_txt_hauto);
	DDX_Control(pDX, TABSTT_LBL_VDTIENTO, m_lbl_vdtiento);
	DDX_Control(pDX, TABSTT_LBL_VDHAUTO, m_lbl_vdhauto);
	DDX_Control(pDX, TABSTT_LBL_VITRI, m_lbl_vitri);
	DDX_Control(pDX, TABSTT_SPIN_NUM, m_spin_num);
	DDX_Control(pDX, TABSTT_CHECK_ZERO, m_check_zero);
	DDX_Control(pDX, TABSTT_CHECK_END, m_check_end);
	DDX_Control(pDX, TABSTT_CHECK_REPLACE, m_check_replace);
}

BEGIN_MESSAGE_MAP(FrmTabSTT, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_EN_CHANGE(TABSTT_TXT_NUM, &FrmTabSTT::OnEnChangeTxtNum)
	ON_NOTIFY(UDN_DELTAPOS, TABSTT_SPIN_NUM, &FrmTabSTT::OnDeltaposSpinNum)
END_MESSAGE_MAP()

// FrmTabSTT message handlers
BOOL FrmTabSTT::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(rgbWhite);

	_SetDefault();

	return TRUE;
}

BOOL FrmTabSTT::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN || iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(TABSTT_TXT_NUM))
			{
				GotoDlgCtrl(GetDlgItem(TABSTT_TXT_TIENTO));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TABSTT_TXT_TIENTO))
			{
				GotoDlgCtrl(GetDlgItem(TABSTT_TXT_HAUTO));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TABSTT_TXT_HAUTO))
			{
				GotoDlgCtrl(GetDlgItem(TABSTT_TXT_NUM));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(TABSTT_TXT_NUM))
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

void FrmTabSTT::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTabSTT::_SetDefault()
{
	m_txt_tiento.SubclassDlgItem(TABSTT_TXT_TIENTO, this);
	m_txt_tiento.SetTextColor(rgbBlue);
	m_txt_tiento.SetBkColor(rgbWhite);	
	m_txt_tiento.SetCueBanner(L"Ví dụ: ký tự '#'");

	m_txt_hauto.SubclassDlgItem(TABSTT_TXT_HAUTO, this);
	m_txt_hauto.SetTextColor(rgbBlue);
	m_txt_hauto.SetBkColor(rgbWhite);
	m_txt_hauto.SetCueBanner(L"Ví dụ: ký tự '. '");	
	m_txt_hauto.SetWindowText(L". ");

	m_lbl_vdtiento.SubclassDlgItem(TABSTT_LBL_VDTIENTO, this);
	m_lbl_vdtiento.SetTextColor(rgbBlue);
	m_lbl_vdtiento.SetFont(&m_FontItalic);

	m_lbl_vdhauto.SubclassDlgItem(TABSTT_LBL_VDHAUTO, this);
	m_lbl_vdhauto.SetTextColor(rgbBlue);
	m_lbl_vdhauto.SetFont(&m_FontItalic);

	m_lbl_vitri.SubclassDlgItem(TABSTT_LBL_VITRI, this);
	m_lbl_vitri.SetTextColor(rgbBlue);
	m_lbl_vitri.SetFont(&m_FontItalic);

	m_spin_num.SetDecimalPlaces(0);
	m_spin_num.SetTrimTrailingZeros(FALSE);
	m_spin_num.SetRangeAndDelta(0, 10000, 1);
	m_spin_num.SetPos(1);

	m_check_zero.SetCheck(1);
	m_check_end.SetCheck(0);
	m_check_replace.SetCheck(0);
}

void FrmTabSTT::OnEnChangeTxtNum()
{
	CString szval = L"";
	m_txt_num.GetWindowTextW(szval);
	int num = _wtoi(szval);
	if (num > 10000)
	{
		m_txt_num.SetWindowText(L"10000");
		m_txt_num.SetSel(-1);
	}
}

void FrmTabSTT::OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult)
{
	GotoDlgCtrl(GetDlgItem(TABSTT_TXT_NUM));
}
