
// GetPropertiesFilesDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "GetPropertiesFiles.h"
#include "GetPropertiesFilesDlg.h"
#include "afxdialogex.h"

#include "FileFilter.h"

#import "D:\CODENEW\_CSH\GLibrarycs\GLibrarycs\bin\Debug\GLibrarycs.tlb"
using namespace GLibrarycs;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CGetPropertiesFilesDlg dialog



CGetPropertiesFilesDlg::CGetPropertiesFilesDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GETPROPERTIESFILES_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CGetPropertiesFilesDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CGetPropertiesFilesDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CGetPropertiesFilesDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CGetPropertiesFilesDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CGetPropertiesFilesDlg message handlers

BOOL CGetPropertiesFilesDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	CEdit* m_txt_path = (CEdit*)(GetDlgItem(IDC_EDIT1));
	m_txt_path->SetWindowText(L"C:\\Users\\LEO\\Desktop\\Doc1.docx");


	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CGetPropertiesFilesDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CGetPropertiesFilesDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CGetPropertiesFilesDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CGetPropertiesFilesDlg::OnBnClickedButton1()
{
	try
	{
		AFileFilter	aff(L"Microsoft Word Document|*.docx|All Files|*.*||");
		CFileDialog dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_NOCHANGEDIR,
			L"Microsoft Word Document|*.doc;*.docx|All Files|*.*||");
		dlgFile.m_ofn.lpstrTitle = _T("Select file document...");

		if (dlgFile.DoModal() != IDOK) return;
		CString szPathFile = dlgFile.GetPathName();
		if (szPathFile != L"")
		{
			CEdit* m_txt_path = (CEdit*)(GetDlgItem(IDC_EDIT1));
			m_txt_path->SetWindowText(szPathFile);
		}
	}
	catch (...) {}
}


void CGetPropertiesFilesDlg::OnBnClickedButton2()
{
	try
	{
		HRESULT hr = CoInitialize(NULL);
		ILibrarycsPtr pILib(__uuidof(CLibrarycs));

		CString szval = L"";
		CEdit* m_txt_path = (CEdit*)(GetDlgItem(IDC_EDIT1));
		m_txt_path->GetWindowTextW(szval);
		szval.Trim();
		if (szval != L"")
		{
			szval = (LPCTSTR)pILib->GetProperties((_bstr_t)szval);
			CEdit* m_txt_result = (CEdit*)(GetDlgItem(IDC_EDIT2));
			m_txt_result->SetWindowText(szval);
		}

		CoUninitialize();
	}
	catch (...) {}	
}
