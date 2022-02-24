#include "pch.h"
#include "DlgThreadProgress.h"

IMPLEMENT_DYNAMIC(DlgThreadProgress, CDialogEx)

DlgThreadProgress::DlgThreadProgress(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgThreadProgress::IDD, pParent)
{
	szTextPro = L"";
	istart = 1, ifinish = 100;
	itotal = 100;
}

DlgThreadProgress::~DlgThreadProgress()
{
	
}

void DlgThreadProgress::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, THREAD_PROGRESS, m_progress);
	DDX_Control(pDX, THREAD_BTN_STOP, m_btn_stop);
}

BEGIN_MESSAGE_MAP(DlgThreadProgress, CDialogEx)
	ON_WM_TIMER()
	ON_BN_CLICKED(THREAD_BTN_STOP, &DlgThreadProgress::OnBnClickedBtnStop)
END_MESSAGE_MAP()

// DlgThreadProgress message handlers
BOOL DlgThreadProgress::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_btn_stop.SetThemeHelper(&m_ThemeHelper);
	m_btn_stop.SetIcon(IDI_ICON_PAUSE, 24, 24);
	m_btn_stop.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_stop.SetBtnCursor(IDC_CURSOR_HAND);

	m_progress.AlignText(DT_CENTER);
	m_progress.SetBarColor(RGB(0, 176, 80));
	m_progress.SetBarBkColor(rgbWhite);
	m_progress.SetTextColor(rgbBlack);
	m_progress.SetTextBkColor(rgbWhite);
	m_progress.SetWindowText(szTextPro);

	itotal = ifinish - istart + 1;
	m_progress.SetRange(0, itotal);
	m_progress.SetPos(0);

	SetTimer(1, 100, NULL);	// 1000ms = 1 second

	return TRUE;
}

void DlgThreadProgress::OnTimer(UINT_PTR nIDEvent)
{
	m_progress.SetPos(demThPro);
	if (demThPro >= itotal) CDialogEx::OnOK();
	CDialogEx::OnTimer(nIDEvent);
}

void DlgThreadProgress::OnBnClickedBtnStop()
{
	bStopThread = true;
	CDialogEx::OnCancel();
}
