#include "pch.h"
#include "FTab3.h"

IMPLEMENT_DYNAMIC(FTab3, CDialogEx)

FTab3::FTab3(CWnd* pParent /*=NULL*/)
	: CDialogEx(FTab3::IDD, pParent)
{

}

FTab3::~FTab3()
{

}

void FTab3::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);


}

BEGIN_MESSAGE_MAP(FTab3, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

END_MESSAGE_MAP()

// FTab3 message handlers
BOOL FTab3::OnInitDialog()
{
	CDialogEx::OnInitDialog();


	return TRUE;
}

BOOL FTab3::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_ESCAPE)
		{

		}
	}

	return FALSE;
}

void FTab3::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialog::OnSysCommand(nID, lParam);
}

void FTab3::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{

}
