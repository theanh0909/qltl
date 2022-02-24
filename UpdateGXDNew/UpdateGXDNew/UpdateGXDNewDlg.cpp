#include "pch.h"
#include "framework.h"
#include "UpdateGXDNew.h"
#include "UpdateGXDNewDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define MAX_LEN					1024
//#define FileAppRun			L"QLTL GXD.exe"
#define FileAppRun				L"QLDAGXD.exe"
#define FILENAMEDLL				L"GDocmanager.dll"
#define ExcelApp				L"Excel.Application"

#define FileHelp				L"help.ini"
#define FileVersion				L"version.ini"
#define FileLogUpdate			L"LogUpdate.txt"
#define FileListUpdate			L"UpdateFiles"

//#define CreateKeySettings		L"WUZkVFJ4VFpYUjBhVzVuYzF4VgR=DJTV=VPOVYTbKlBZSFJURlhFZFlSRXBUUTFGVHg9UlFF"
#define CreateKeySettings		L"WUZkUVJ4VFpYUjBhVzVuYzF4VgR=DJTV=VPOVYTbKlBZSFJURlhFZFlSRXBUUTFFVHg9UlFV"
#define CreateKeySMStart		L"YwWE5kWUowWEZObGRIUnBibWR6ZzRUlVW4=E6REGVRVTAJ51ZZa1ZjUjFoRVNsTkRYRlhZTj10SU"

#define UpldateSoft				L"UpldateSoft"
#define UploadData				L"UploadData"
#define DlgCheckUpdateSoft		L"CheckUpdateSoft"
#define DlgVisibleCheckbox		L"VisibleCheckbox"

#define FileSmartStartApp		L"SmartStart.exe"
#define FileVersionsm			L"versionsm.ini"
#define Version					L"Version"

// CUpdateGXDNewDlg dialog
CUpdateGXDNewDlg::CUpdateGXDNewDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_UPDATEGXDNEW_DIALOG, pParent)
{
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;

	szOffice = L"32";
	pathFolder = L"";
	iCheckUpSoft = -1;

	m_myFont.DeleteObject();
	VERIFY(m_myFont.CreateFont(
		22,							// nHeight
		0,							// nWidth
		0,							// nEscapement
		0,							// nOrientation
		FW_NORMAL,					// nWeight
		TRUE,						// bItalic
		FALSE,						// bUnderline
		0,							// cStrikeOut
		ANSI_CHARSET,				// nCharSet
		OUT_DEFAULT_PRECIS,			// nOutPrecision
		CLIP_DEFAULT_PRECIS,		// nClipPrecision
		DEFAULT_QUALITY,			// nQuality
		DEFAULT_PITCH | FF_SWISS,	// nPitchAndFamily
		_T("MS Shell Dlg")));		// lpszFacename
}

CUpdateGXDNewDlg::~CUpdateGXDNewDlg()
{
	delete bb;
	delete sv;
	delete reg;
}

void CUpdateGXDNewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BUTTON_OK, m_btn_ok);
	DDX_Control(pDX, IDC_BUTTON_CANCEL, m_btn_cancel);
	DDX_Control(pDX, IDC_TXT_NOIDUNG, m_txt_noidung);
	DDX_Control(pDX, IDC_LBL_WARNING, m_lbl_warning);
	DDX_Control(pDX, IDC_CHECK_SKIPUPDATE, m_check_skipup);
	DDX_Control(pDX, IDC_PROGRESS_UPDATE, m_progressUpdate);
}

BEGIN_MESSAGE_MAP(CUpdateGXDNewDlg, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(IDC_BUTTON_OK, &CUpdateGXDNewDlg::OnBnClickedButtonOk)
	ON_BN_CLICKED(IDC_BUTTON_CANCEL, &CUpdateGXDNewDlg::OnBnClickedButtonCancel)
	ON_BN_CLICKED(IDC_CHECK_SKIPUPDATE, &CUpdateGXDNewDlg::OnBnClickedCheckSkipupdate)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(CUpdateGXDNewDlg)
	EASYSIZE(IDC_TXT_NOIDUNG, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(IDC_LINE_VER, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(IDC_LBL_WARNING, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(IDC_CHECK_SKIPUPDATE, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(IDC_BUTTON_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(IDC_BUTTON_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(IDC_PROGRESS_UPDATE, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// CUpdateGXDNewDlg message handlers

BOOL CUpdateGXDNewDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MAINFRAME));
	SetIcon(hIcon, FALSE);

	_SetDefault();
	_LoadLogUpdate();

	INIT_EASYSIZE;

	return TRUE;  // return TRUE  unless you set the focus to a control
}

BOOL CUpdateGXDNewDlg::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(IDC_BUTTON_OK))
			{
				OnBnClickedButtonOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			OnBnClickedButtonCancel();
			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		switch (chr)
		{
		case 0x01:
		{
			if (pWndWithFocus == GetDlgItem(IDC_TXT_NOIDUNG))
			{
				m_txt_noidung.SetSel(0, -1);
				return TRUE;
			}

			break;
		}

		default:
			break;
		}
	}

	return FALSE;
}

void CUpdateGXDNewDlg::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void CUpdateGXDNewDlg::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void CUpdateGXDNewDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedButtonCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void CUpdateGXDNewDlg::OnBnClickedButtonOk()
{
	try
	{
		CoInitialize(NULL);

		_ApplicationPtr objExcel;
		if (FAILED(objExcel.GetActiveObject(ExcelApp)))
		{
			if (FAILED(objExcel.CreateInstance(ExcelApp)))
			{
				CoUninitialize();
				MessageBoxW(L"Có lỗi trong quá trình khởi động Excel. "
					L"Bạn hãy cài lại office hoặc thay thế một phiên bản office khác và thử lại", 
					L"Hãy kiểm tra lại Office", MB_ICONERROR);
				return;
			}
		}

		szOffice = L"";

		//HRESULT hr = objExcel.GetActiveObject(ExcelApp);
		if (!(/*FAILED(hr) && */objExcel == NULL))
		{
			CString szOperatingSystem = (LPCTSTR)objExcel->GetOperatingSystem();
			if (szOperatingSystem.Find(L"64") >= 0) szOffice = L"64";
			else szOffice = L"32";

			int result = (int)MessageBox(L"Yêu cầu đóng tất cả ứng dụng Microsoft Excel trước khi tiến hành cập nhật phần mềm. "
				L"Chọn [YES/NO] để tự động đóng hết ứng dụng Excel và tự động lưu (Y) hoặc không lưu (N).",
				L"Thông báo", MB_YESNOCANCEL | MB_ICONQUESTION | MB_DEFBUTTON1);
			if (result == 6 || result == 7)
			{
				_WorkbookPtr Wb;
				CString szName = L"", szType = L"";
				int ncount = objExcel->Application->Workbooks->GetCount();
				for (int i = 1; i <= ncount; i++)
				{
					Wb = objExcel->Application->Workbooks->GetItem(i);

					if (result == 6)
					{
						szName = (LPCTSTR)Wb->Name;
						szType = szName.Right(5);
						if (szType.Find(L".") >= 0 && szType.Find(L"xltm") == -1)
						{
							Wb->Save();
						}						
					}

					Wb->Close(false);
				}

				objExcel->Quit();
				_QuitApplication(L"EXCEL.EXE");
			}
			else return;
		}

		CoUninitialize();

		if (szOffice != L"")
		{
			CString szFileSaved = pathFolder + FileListUpdate + szOffice + L".txt";
			if (sv->_DownloadFile(sv->dbSeverDir + FileListUpdate + szOffice + L".txt", szFileSaved))
			{
				vector<CString> vecLine;
				int nTotal = _ReadMultiUnicode(szFileSaved, vecLine);
				if (nTotal > 0)
				{
					m_btn_ok.ShowWindow(0);
					m_btn_cancel.ShowWindow(0);
					m_check_skipup.ShowWindow(0);
					m_progressUpdate.SetWindowText(L"Đang tiến hành cập nhật dữ liệu...");
					m_progressUpdate.ShowWindow(1);

					m_progressUpdate.SetRange(0, nTotal - 1);
					m_progressUpdate.SetPos(0);

					m_txt_noidung.SetSel(-1);

					int pos = -1;
					int iLenDir = sv->dbSeverDir.GetLength();

					CString szval = L"";
					CString szFull = L"", szDownload = L"", szNext = L"", szNewFile = L"";

					for (int i = 0; i < nTotal; i++)
					{
						if (vecLine[i] != L"")
						{
							szFull = sv->dbSeverMain + vecLine[i];
							pos = szFull.Find(L"|");
							if (pos > 0) szFull = szFull.Left(pos);
							szFull.Trim();

							szNext = szFull.Right(szFull.GetLength() - iLenDir);
							szNext.Replace(L"/", L"\\");

							if (szNext.Left(2) == L"32" || szNext.Left(2) == L"64")
							{
								szNext = szNext.Right(szNext.GetLength() - 2);
							}

							if (szNext.Left(1) == L"\\") szNext = szNext.Right(szNext.GetLength() - 1);
							_CreateDirectories(pathFolder + _GetNameOfPath(szNext, 1));
							szDownload = pathFolder + szNext;

							sv->_DownloadFile(szFull, szDownload);
						}

						m_progressUpdate.SetPos(i);
					}

					m_progressUpdate.ShowWindow(0);
					m_check_skipup.ShowWindow(1);
					m_btn_cancel.ShowWindow(1);
					m_btn_ok.ShowWindow(1);
				}

				vecLine.clear();
				DeleteFile(szFileSaved);
			}

			_SaveRegistry();

			if (_IsProcessRunning(FileSmartStartApp))
			{
				// Chú ý: Giá trị 'CheckRestartFiles' --> Liên hệ với DLL GFilesManager (SmartStart.exe)
				reg->CreateKey(bb->BaseDecodeEx(CreateKeySMStart));
				reg->WriteChar("CheckRestartFiles", "1");
			}

			CDialogEx::OnOK();

			if (MessageBox(
				L"Cập nhật dữ liệu thành công. Bạn có muốn tự động mở phần mềm?.",
				L"Thông báo", MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON1) == 6)
			{
				_DownloadFileSmartStart();
				ShellExecute(NULL, L"open", pathFolder + FileAppRun, NULL, NULL, SW_SHOWMAXIMIZED);
			}			

			delete sv;
		}
		else
		{
			MessageBox(L"Không xác định được phiên bản Excel. "
				L"Vui lòng liên hệ Tổng đài 1900.0147 để được trợ giúp",
				L"Lỗi cập nhật", MB_OK | MB_ICONERROR);
		}		
	}
	catch (...) {}
}


void CUpdateGXDNewDlg::_DownloadFileSmartStart()
{
	try
	{
		// Tải dữ liệu
		sv->_LoadDecodeBase64();

		CString szFileApp = pathFolder + FileSmartStartApp;
		if (sv->_DownloadFile(sv->dbSeverDirSM + szOffice + FileSmartStartApp, szFileApp))
		{
			// Ghi thông tin version vào Registry
			CString szFileVersion = pathFolder + FileVersionsm;
			if (sv->_DownloadFile(sv->dbSeverDirSM + FileVersionsm, szFileVersion))
			{
				vector<CString> vecLine;
				int ncount = _ReadMultiUnicode(szFileVersion, vecLine);
				if (ncount > 0)
				{
					CString szCreate = bb->BaseDecodeEx(CreateKeySMStart);
					reg->CreateKey(szCreate);
					reg->WriteChar(
						_ConvertCstringToChars(Version),
						_ConvertCstringToChars(vecLine[0]));
				}
				vecLine.clear();
				DeleteFile(szFileVersion);
			}
		}
	}
	catch (...) {}
}

void CUpdateGXDNewDlg::OnBnClickedButtonCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

CString CUpdateGXDNewDlg::_GetPathFolder(CString fileName)
{
	try
	{
		HMODULE hModule;
		hModule = GetModuleHandle(fileName);
		TCHAR szFileName[MAX_LEN];
		DWORD dSize = MAX_LEN;
		GetModuleFileName(hModule, szFileName, dSize);
		CString szSource = szFileName;
		for (int i = szSource.GetLength() - 1; i > 0; i--)
		{
			if (szSource.GetAt(i) == '\\')
			{
				szSource = szSource.Left(szSource.GetLength() - (szSource.GetLength() - i) + 1);
				break;
			}
		}
		return szSource;
	}
	catch (...) {}
	return L"";
}

int CUpdateGXDNewDlg::_ReadMultiUnicode(CString spathFile, vector<CString> &vecTextLine)
{
	try
	{
		vecTextLine.clear();
		if (_FileExistsUnicode(spathFile) != 1) return 0;

		wifstream readfile_(_ConvertCstringToChars(spathFile));
		readfile_.imbue(locale(locale::empty(), new codecvt_utf8<wchar_t, 0x10ffff, consume_header>));

		if (readfile_.is_open())
		{
			wstring username;
			while (getline(readfile_, username))
			{
				vecTextLine.push_back(username.c_str());
			}

			readfile_.close();
		}

		return (int)vecTextLine.size();
	}
	catch (...) {}

	return 0;
}

BOOL CUpdateGXDNewDlg::_CreateNewFolder(CString directoryPath)
{
	CString directoryPathClone = directoryPath;
	LPTSTR lpPath = directoryPath.GetBuffer(MAX_PATH);
	if (!PathFileExists(lpPath))
	{
		if (!CreateDirectory(lpPath, NULL) && (GetLastError() == ERROR_PATH_NOT_FOUND))
		{
			if (PathRemoveFileSpec(lpPath))
			{
				CString newPath = lpPath;
				_CreateNewFolder(newPath);
			}
		}
	}

	directoryPath.ReleaseBuffer(-1);
	return PathFileExists(directoryPathClone);
}

bool CUpdateGXDNewDlg::_CreateDirectories(CString pathName)
{
	try
	{
		if (pathName.Right(1) != L"\\") pathName += L"\\";
		for (int i = 0; i < pathName.GetLength(); i++)
		{
			if (pathName.Mid(i, 1) == L"\\") _CreateNewFolder(pathName.Left(i));
		}
		return true;
	}
	catch (...) {}
	return false;
}

CString CUpdateGXDNewDlg::_GetNameOfPath(CString pathFile, int ipath)
{
	try
	{
		if (pathFile != L"")
		{
			if (pathFile.Right(1) == L"\\") pathFile = pathFile.Left(pathFile.GetLength() - 1);

			int nLen = pathFile.GetLength();
			for (int i = nLen - 1; i >= 0; i--)
			{
				if (pathFile.Mid(i, 1) == L"\\" || pathFile.Mid(i, 1) == L"/")
				{
					if (ipath == -1)
					{
						// Trả về kết quả là tên file
						CString szfile = pathFile.Right(nLen - i - 1);
						int nLenName = szfile.GetLength();
						for (int j = nLenName - 1; j >= 0; j--)
						{
							if (szfile.Mid(j, 1) == L".")
							{
								return szfile.Left(j);
							}
						}
						return szfile;
					}
					else if (ipath == 0)
					{
						// Trả về kết quả là tên file + đuôi
						return pathFile.Right(nLen - i - 1);
					}

					// Trả về đường dẫn chứa file
					return pathFile.Left(i);
				}
			}
		}
	}
	catch (...) {}
	return L"";
}

void CUpdateGXDNewDlg::_SetDefault()
{
	try
	{
		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_progressUpdate.AlignText(DT_CENTER);
		m_progressUpdate.SetBarColor(RGB(0, 176, 80));
		m_progressUpdate.SetBarBkColor(rgbWhite);
		m_progressUpdate.SetTextColor(rgbBlack);
		m_progressUpdate.SetTextBkColor(rgbWhite);
		m_progressUpdate.ShowWindow(0);

		m_txt_noidung.SubclassDlgItem(IDC_TXT_NOIDUNG, this);
		m_txt_noidung.SetTextColor(RGB(0, 0, 255));
		m_txt_noidung.SetBkColor(RGB(255, 255, 255));

		m_lbl_warning.SubclassDlgItem(IDC_LBL_WARNING, this);
		m_lbl_warning.SetTextColor(RGB(240, 0, 0));
		m_lbl_warning.SetFont(&m_myFont);

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
		reg->CreateKey(szCreate);

		int iVisibleCheck = 0;
		CString szval = reg->ReadString(DlgVisibleCheckbox, L"");
		if (szval != L"") iVisibleCheck = _wtoi(szval.Trim());
		if (iVisibleCheck != 1) iVisibleCheck = 0;
		if (iVisibleCheck == 1) m_check_skipup.ShowWindow(0);
	}
	catch (...) {}
}

void CUpdateGXDNewDlg::_LoadLogUpdate()
{
	try
	{
		pathFolder = _GetPathFolder(FILENAMEDLL);
		if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";

		sv->_LoadDecodeBase64();

		vector<CString> vecLine;
		CString szFileSaved = pathFolder + FileHelp;
		if (sv->_DownloadFile(sv->dbSeverDir + FileHelp, szFileSaved))
		{
			int ncount = _ReadMultiUnicode(szFileSaved, vecLine);
			if (ncount > 0)
			{
				for (int i = ncount - 1; i >= 0; i--)
				{
					vecLine[i].Trim();
					if (vecLine[i] == L"" || vecLine[i].Left(2) == L"//") vecLine.erase(vecLine.begin() + i);
				}
			}

			ncount = (int)vecLine.size();
			if (ncount >= 5) this->SetWindowText(L"Cập nhật phiên bản mới: " + vecLine[4]);

			DeleteFile(szFileSaved);
		}
		vecLine.clear();

		szFileSaved = pathFolder + FileLogUpdate;
		if (sv->_DownloadFile(sv->dbSeverDir + FileLogUpdate, szFileSaved))
		{
			CString szval = L"";
			int nTotal = _ReadMultiUnicode(szFileSaved, vecLine);
			if (nTotal > 0)
			{
				for (int i = 0; i < nTotal; i++)
				{
					szval += vecLine[i];
					szval += L"\r\n";
				}
			}

			vecLine.clear();
			DeleteFile(szFileSaved);
			m_txt_noidung.SetWindowText(szval);
		}
	}
	catch (...) {}
}

void CUpdateGXDNewDlg::OnBnClickedCheckSkipupdate()
{
	iCheckUpSoft = (int)m_check_skipup.GetCheck();
	if (iCheckUpSoft == 1)
	{
		MessageBox(L"Để hiển thị lại thông báo có phiên bản cập nhật mới. "
			L"Bạn vào Thiết lập tùy chọn -> Tích chọn 'Luôn luôn hiển thị thông báo cập nhật phần mềm, dữ liệu mới'.", 
			L"Hướng dẫn", MB_OK | MB_ICONINFORMATION);
	}
}

void CUpdateGXDNewDlg::_QuitApplication(CString szAppName)
{
	try
	{
		DWORD cbNeeded;
		DWORD aProcesses[1024];
		if (EnumProcesses(aProcesses, sizeof(aProcesses), &cbNeeded))
		{
			HANDLE tmpH;
			unsigned int ii;
			CString szNameProc = L"";
			DWORD iTotalProc = cbNeeded / sizeof(DWORD);
			for (ii = 0; ii < iTotalProc; ii++)
			{
				szNameProc = _GetDirModules(aProcesses[ii], false);
				if (szNameProc != L"")
				{
					szNameProc = _GetNameOfPath(szNameProc);
					if (szNameProc == szAppName)
					{
						tmpH = OpenProcess(PROCESS_ALL_ACCESS, TRUE, aProcesses[ii]);
						if (tmpH != NULL) TerminateProcess(tmpH, 0);
					}
				}
			}
		}
	}
	catch (...) {}
}


void CUpdateGXDNewDlg::_SaveRegistry()
{
	try
	{
		if ((int)m_check_skipup.IsWindowVisible() == 1)
		{
			if (iCheckUpSoft != -1)
			{
				CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
				reg->CreateKey(szCreate);

				iCheckUpSoft = 1;
				CString szval = L"";
				if ((int)m_check_skipup.GetCheck() == 1) iCheckUpSoft = 0;
				szval.Format(L"%d", iCheckUpSoft);
				reg->WriteChar(_ConvertCstringToChars(DlgCheckUpdateSoft), _ConvertCstringToChars(szval));
			}
		}			
	}
	catch (...) {}
}


