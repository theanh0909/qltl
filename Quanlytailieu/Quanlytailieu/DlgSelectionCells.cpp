#include "pch.h"
#include "DlgSelectionCells.h"

IMPLEMENT_DYNAMIC(DlgSelectionCells, CDialogEx)

DlgSelectionCells::DlgSelectionCells(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgSelectionCells::IDD, pParent)
{
	iSelectCell = -1;
	blEnableSelection = false;
}

DlgSelectionCells::~DlgSelectionCells()
{

}

void DlgSelectionCells::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, SELCEL_RAD_SEL, m_radio_sel);
	DDX_Control(pDX, SELCEL_RAD_GROUP, m_radio_group);
	DDX_Control(pDX, SELCEL_RAD_ALL, m_radio_all);
	DDX_Control(pDX, SELCEL_BTN_OK, m_btn_ok);
	DDX_Control(pDX, SELCEL_BTN_CANCEL, m_btn_cancel);
}

BEGIN_MESSAGE_MAP(DlgSelectionCells, CDialogEx)
	ON_WM_SYSCOMMAND()

	ON_BN_CLICKED(SELCEL_BTN_CANCEL, &DlgSelectionCells::OnBnClickedBtnCancel)
	ON_BN_CLICKED(SELCEL_BTN_OK, &DlgSelectionCells::OnBnClickedBtnOk)
END_MESSAGE_MAP()

// DlgSelectionCells message handlers
BOOL DlgSelectionCells::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_SEND));
	SetIcon(hIcon, FALSE);

	int ncount = (int)vecSelect.size();

	if (!blEnableSelection)
	{
		if (vecSelect[0].row == vecSelect[ncount - 1].row) m_radio_group.SetCheck(1);
		else m_radio_sel.SetCheck(1);
	}
	else
	{
		m_radio_sel.EnableWindow(0);
		m_radio_group.SetCheck(1);
	}

	m_btn_ok.SetThemeHelper(&m_ThemeHelper);
	m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

	return TRUE;
}

BOOL DlgSelectionCells::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(SELCEL_BTN_OK))
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

void DlgSelectionCells::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgSelectionCells::OnBnClickedBtnCancel()
{
	CDialogEx::OnCancel();
}

void DlgSelectionCells::OnBnClickedBtnOk()
{
	CDialogEx::OnOK();

	if ((int)m_radio_all.GetCheck() == 1) iSelectCell = 2;
	if ((int)m_radio_group.GetCheck() == 1) iSelectCell = 1;
	if ((int)m_radio_sel.GetCheck() == 1) iSelectCell = 0;
}
