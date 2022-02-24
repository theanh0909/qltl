#include "pch.h"
#include "FrmTabReplace.h"

// FrmTabReplace dialog
IMPLEMENT_DYNAMIC(FrmTabReplace, CDialogEx)

FrmTabReplace::FrmTabReplace(CWnd* pParent /*=nullptr*/)
	: CDialogEx(FVIEW_RENAME_REPLACE, pParent)
{

}

FrmTabReplace::~FrmTabReplace()
{
}

void FrmTabReplace::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, TABREP_TXT_REP, m_txt_rep);
	DDX_Control(pDX, TABREP_TXT_CHUOI, m_txt_chuoi);
	DDX_Control(pDX, TABREP_CHK_CASE, m_chk_case);
	DDX_Control(pDX, TABREP_CHK_KODAU, m_chk_kodau);

}

BEGIN_MESSAGE_MAP(FrmTabReplace, CDialogEx)
	ON_WM_SYSCOMMAND()
END_MESSAGE_MAP()

// FrmTabReplace message handlers
BOOL FrmTabReplace::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	SetBackgroundColor(rgbWhite);

	_SetDefault();

	return TRUE;
}

BOOL FrmTabReplace::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN || iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(TABREP_TXT_REP))
			{
				GotoDlgCtrl(GetDlgItem(TABREP_TXT_CHUOI));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(TABREP_TXT_CHUOI))
			{
				GotoDlgCtrl(GetDlgItem(TABREP_TXT_REP));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
			return TRUE;
		}
	}

	return FALSE;
}

void FrmTabReplace::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) PostMessageA(hWndAutoRename, WM_CLOSE, 0, 0);
	else CDialog::OnSysCommand(nID, lParam);
}

void FrmTabReplace::_SetDefault()
{
	m_txt_rep.SubclassDlgItem(TABREP_TXT_REP, this);
	m_txt_rep.SetTextColor(rgbBlue);
	m_txt_rep.SetBkColor(rgbWhite);
	m_txt_rep.SetCueBanner(L"Ví dụ: 'GXd.vn- '");

	m_txt_chuoi.SubclassDlgItem(TABREP_TXT_CHUOI, this);
	m_txt_chuoi.SetTextColor(rgbBlue);
	m_txt_chuoi.SetBkColor(rgbWhite);
	m_txt_chuoi.SetCueBanner(L"Ví dụ: 'GXD'");

	m_chk_case.SetCheck(1);
	m_chk_kodau.SetCheck(0);
}
