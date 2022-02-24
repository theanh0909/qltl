#include "pch.h"
#include "FTab2.h"

IMPLEMENT_DYNAMIC(FTab2, CDialogEx)

FTab2::FTab2(CWnd* pParent /*=NULL*/)
	: CDialogEx(FTab2::IDD, pParent)
{

}

FTab2::~FTab2()
{

}

void FTab2::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);


}

BEGIN_MESSAGE_MAP(FTab2, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
END_MESSAGE_MAP()

// FTab2 message handlers
BOOL FTab2::OnInitDialog()
{
	CDialogEx::OnInitDialog();


	return TRUE;
}

BOOL FTab2::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{

	}

	return FALSE;
}

void FTab2::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialog::OnSysCommand(nID, lParam);
}

void FTab2::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{

}
