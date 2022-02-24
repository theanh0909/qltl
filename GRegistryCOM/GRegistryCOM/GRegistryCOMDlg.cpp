#include "pch.h"
#include "framework.h"
#include "GRegistryCOM.h"
#include "GRegistryCOMDlg.h"
#include "afxdialogex.h"

// CGRegistryCOMDlg dialog
CGRegistryCOMDlg::CGRegistryCOMDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GREGISTRYCOM_DIALOG, pParent)
{

}

void CGRegistryCOMDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CGRegistryCOMDlg, CDialogEx)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()

// CGRegistryCOMDlg message handlers
BOOL CGRegistryCOMDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	SetLayeredWindowAttributes(RGB(255, 0, 255), 0, LWA_COLORKEY);

	return TRUE;
}

BOOL CGRegistryCOMDlg::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(255, 0, 255));
	return FALSE;
}
