
// QLTL GXDDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "QLTL GXD.h"
#include "QLTL GXDDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CQLTLGXDDlg dialog



CQLTLGXDDlg::CQLTLGXDDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_QLTLGXD_DIALOG, pParent)
{

}

void CQLTLGXDDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CQLTLGXDDlg, CDialogEx)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()


// CQLTLGXDDlg message handlers

BOOL CQLTLGXDDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	SetWindowLong(m_hWnd, GWL_EXSTYLE, GetWindowLong(m_hWnd, GWL_EXSTYLE) ^ WS_EX_LAYERED);
	SetLayeredWindowAttributes(RGB(255, 0, 255), 0, LWA_COLORKEY);

	return TRUE;
}

BOOL CQLTLGXDDlg::OnEraseBkgnd(CDC* pDC)
{
	CRect clientRect;
	GetClientRect(&clientRect);
	pDC->FillSolidRect(clientRect, RGB(255, 0, 255));
	return FALSE;
}