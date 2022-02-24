#include "pch.h"
#include "FrmTabAdd.h"

// FrmTabAdd dialog
IMPLEMENT_DYNAMIC(FrmTabAdd, CDialogEx)

FrmTabAdd::FrmTabAdd(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FVIEW_RENAME_ADD, pParent)
{

}

FrmTabAdd::~FrmTabAdd()
{
}

void FrmTabAdd::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TABADD_TXT_CHUOI, m_txt_chuoi);
	DDX_Control(pDX, TABADD_RAD_BEGIN, m_rad_begin);
	DDX_Control(pDX, TABADD_RAD_END, m_rad_end);	
	DDX_Control(pDX, TABADD_RAD_MID, m_rad_mid);
	DDX_Control(pDX, TABADD_TXT_NUM, m_txt_num);
	DDX_Control(pDX, TABADD_SPIN_NUM, m_spin_num);
	DDX_Control(pDX, TABADD_LBL_VIDUNUM, m_lbl_vdnum);
}

BEGIN_MESSAGE_MAP(FrmTabAdd, CDialogEx)
	ON_WM_SYSCOMMAND()	
	ON_BN_CLICKED(TABADD_RAD_BEGIN, &FrmTabAdd::OnBnClickedRadBegin)
	ON_BN_CLICKED(TABADD_RAD_END, &FrmTabAdd::OnBnClickedRadEnd)
	ON_BN_CLICKED(TABADD_RAD_MID, &FrmTabAdd::OnBnClickedRadMid)
	ON_EN_CHANGE(TABADD_TXT_NUM, &FrmTabAdd::OnEnChangeTxtNum)
	ON_NOTIFY(UDN_DELTAPOS, TABADD_SPIN_NUM, &FrmTabAdd::OnDeltaposSpinNum)
END_MESSAGE_MAP()

// FrmTabAdd message handlers
BOOL FrmTabAdd::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(rgbWhite);

	_SetDefault();

	return TRUE;
}

BOOL FrmTabAdd::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN || iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(TABADD_TXT_NUM))
			{
				GotoDlgCtrl(GetDlgItem(TABADD_TXT_CHUOI));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TABADD_TXT_CHUOI))
			{
				GotoDlgCtrl(GetDlgItem(TABADD_TXT_NUM));
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
		if (pWndWithFocus == GetDlgItem(TABADD_TXT_NUM))
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

void FrmTabAdd::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTabAdd::OnEnChangeTxtNum()
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

void FrmTabAdd::OnDeltaposSpinNum(NMHDR *pNMHDR, LRESULT *pResult)
{
	GotoDlgCtrl(GetDlgItem(TABADD_TXT_NUM));
}

void FrmTabAdd::_SetDefault()
{
	m_txt_chuoi.SubclassDlgItem(TABADD_TXT_CHUOI, this);
	m_txt_chuoi.SetTextColor(rgbBlue);
	m_txt_chuoi.SetBkColor(rgbWhite);
	m_txt_chuoi.SetCueBanner(L"Ví dụ: 'Gxd.vn- '");

	m_lbl_vdnum.SubclassDlgItem(TABADD_LBL_VIDUNUM, this);
	m_lbl_vdnum.SetTextColor(rgbBlue);
	m_lbl_vdnum.SetFont(&m_FontItalic);

	m_spin_num.SetDecimalPlaces(0);
	m_spin_num.SetTrimTrailingZeros(FALSE);
	m_spin_num.SetRangeAndDelta(1, 10000, 1);
	m_spin_num.SetPos(1);

	m_rad_begin.SetCheck(1);
	m_rad_end.SetCheck(0);
	m_rad_mid.SetCheck(0);

	_SetEnableNumber(false);
}

void FrmTabAdd::OnBnClickedRadBegin()
{
	_SetEnableNumber(false);
}

void FrmTabAdd::OnBnClickedRadEnd()
{
	_SetEnableNumber(false);
}

void FrmTabAdd::OnBnClickedRadMid()
{
	_SetEnableNumber(true);
}

void FrmTabAdd::_SetUnCheckRadio()
{
	m_rad_begin.SetCheck(0);
	m_rad_end.SetCheck(0);
	m_rad_mid.SetCheck(0);
}

void FrmTabAdd::_SetEnableNumber(bool bl)
{
	m_txt_num.EnableWindow(bl);
	m_spin_num.EnableWindow(bl);
}
