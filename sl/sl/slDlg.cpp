#include "pch.h"
#include "framework.h"
#include "sl.h"
#include "slDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CslDlg dialog
CslDlg::CslDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SL_DIALOG, pParent)
{
	
}

void CslDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CslDlg, CDialogEx)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()


// CslDlg message handlers

BOOL CslDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	SetLayeredWindowAttributes(RGB(255, 0, 255), 0, LWA_COLORKEY);

	return TRUE;
}

BOOL CslDlg::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(255, 0, 255));
	return FALSE;
}

