#include "pch.h"
#include "framework.h"
#include "SetupGXD.h"
#include "SetupGXDDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CSetupGXDDlg dialog

CSetupGXDDlg::CSetupGXDDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SETUPGXD_DIALOG, pParent)
{

}

void CSetupGXDDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CSetupGXDDlg, CDialogEx)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()


// CSetupGXDDlg message handlers

BOOL CSetupGXDDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	SetLayeredWindowAttributes(RGB(255, 0, 255), 0, LWA_COLORKEY);

	return TRUE;
}

BOOL CSetupGXDDlg::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(255, 0, 255));
	return FALSE;
}