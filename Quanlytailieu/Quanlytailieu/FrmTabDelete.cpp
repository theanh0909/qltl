#include "pch.h"
#include "FrmTabDelete.h"

// FrmTabDelete dialog
IMPLEMENT_DYNAMIC(FrmTabDelete, CDialogEx)

FrmTabDelete::FrmTabDelete(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FVIEW_RENAME_DELETE, pParent)
{

}

FrmTabDelete::~FrmTabDelete()
{
}

void FrmTabDelete::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TABDEL_TXT_CHUOI, m_txt_chuoi);
	DDX_Control(pDX, TABDEL_RAD_TRIM, m_rad_trim);
	DDX_Control(pDX, TABDEL_RAD_SPACE, m_rad_space);
	DDX_Control(pDX, TABDEL_RAD_CHUOI, m_rad_chuoi);
	DDX_Control(pDX, TABDEL_RAD_ALL, m_rad_all);
	DDX_Control(pDX, TABDEL_RAD_BEGIN, m_rad_begin);
	DDX_Control(pDX, TABDEL_RAD_END, m_rad_end);
	DDX_Control(pDX, TABDEL_RAD_LEFT, m_rad_left);
	DDX_Control(pDX, TABDEL_RAD_RIGHT, m_rad_right);
}

BEGIN_MESSAGE_MAP(FrmTabDelete, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(TABDEL_RAD_TRIM, &FrmTabDelete::OnBnClickedRadTrim)
	ON_BN_CLICKED(TABDEL_RAD_SPACE, &FrmTabDelete::OnBnClickedRadSpace)
	ON_BN_CLICKED(TABDEL_RAD_CHUOI, &FrmTabDelete::OnBnClickedRadChuoi)
END_MESSAGE_MAP()

// FrmTabDelete message handlers
BOOL FrmTabDelete::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(rgbWhite);

	_SetDefault();

	return TRUE;
}

BOOL FrmTabDelete::PreTranslateMessage(MSG* pMsg)
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

void FrmTabDelete::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTabDelete::OnBnClickedRadTrim()
{
	_SetEnableRadio(false);
}


void FrmTabDelete::OnBnClickedRadSpace()
{
	_SetEnableRadio(false);
}

void FrmTabDelete::OnBnClickedRadChuoi()
{
	_SetEnableRadio(true);
}

void FrmTabDelete::_SetUnCheckRadio()
{
	m_rad_trim.SetCheck(0);
	m_rad_space.SetCheck(0);
	m_rad_chuoi.SetCheck(0);
}

void FrmTabDelete::_SetUnCheckChildRadio()
{
	m_rad_all.SetCheck(0);
	m_rad_begin.SetCheck(0);
	m_rad_end.SetCheck(0);
	m_rad_left.SetCheck(0);
	m_rad_right.SetCheck(0);
}

void FrmTabDelete::_SetEnableRadio(bool bl)
{
	m_txt_chuoi.EnableWindow(bl);
	m_rad_all.EnableWindow(bl);
	m_rad_begin.EnableWindow(bl);
	m_rad_end.EnableWindow(bl);
	m_rad_left.EnableWindow(bl);
	m_rad_right.EnableWindow(bl);
}

void FrmTabDelete::_SetDefault()
{
	m_txt_chuoi.SubclassDlgItem(TABDEL_TXT_CHUOI, this);
	m_txt_chuoi.SetTextColor(rgbBlue);
	m_txt_chuoi.SetBkColor(rgbWhite);
	m_txt_chuoi.SetCueBanner(L"Ví dụ: 'Gxd.vn- '");

	m_rad_trim.SetCheck(0);
	m_rad_space.SetCheck(0);
	m_rad_chuoi.SetCheck(1);

	m_rad_all.SetCheck(1);
	m_rad_begin.SetCheck(0);
	m_rad_end.SetCheck(0);
	m_rad_left.SetCheck(0);
	m_rad_right.SetCheck(0);

	_SetEnableRadio(true);
}
