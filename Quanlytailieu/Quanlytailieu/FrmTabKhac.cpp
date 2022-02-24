#include "pch.h"
#include "FrmTabKhac.h"

// FrmTabKhac dialog
IMPLEMENT_DYNAMIC(FrmTabKhac, CDialogEx)

FrmTabKhac::FrmTabKhac(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FVIEW_RENAME_KHAC, pParent)
{

}

FrmTabKhac::~FrmTabKhac()
{
}

void FrmTabKhac::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TABKHAC_RAD_UPPER, m_rad_upper);
	DDX_Control(pDX, TABKHAC_RAD_LOWER, m_rad_lower);
	DDX_Control(pDX, TABKHAC_RAD_UNSIGNED, m_rad_unsigned);
	DDX_Control(pDX, TABKHAC_RAD_PROPER, m_rad_proper);
	DDX_Control(pDX, TABKHAC_RAD_KODAU, m_rad_kodau);
}

BEGIN_MESSAGE_MAP(FrmTabKhac, CDialogEx)
	ON_WM_SYSCOMMAND()
END_MESSAGE_MAP()

// FrmTabKhac message handlers
BOOL FrmTabKhac::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(rgbWhite);

	_SetDefault();

	return TRUE;
}

BOOL FrmTabKhac::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_ESCAPE)
		{
			PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
			return TRUE;
		}
	}

	return FALSE;
}

void FrmTabKhac::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTabKhac::_SetDefault()
{
	_SetUnCheckRadio();
	m_rad_upper.SetCheck(1);
}

void FrmTabKhac::_SetUnCheckRadio()
{
	m_rad_upper.SetCheck(0);
	m_rad_lower.SetCheck(0);
	m_rad_unsigned.SetCheck(0);
	m_rad_proper.SetCheck(0);
	m_rad_kodau.SetCheck(0);
}
