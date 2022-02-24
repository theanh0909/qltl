#include "pch.h"
#include "DlgCreateNewTemp.h"

IMPLEMENT_DYNAMIC(DlgCreateNewTemp, CDialogEx)

DlgCreateNewTemp::DlgCreateNewTemp(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgCreateNewTemp::IDD, pParent)
{
	ff = new Function;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;
}

DlgCreateNewTemp::~DlgCreateNewTemp()
{
	delete ff;
}

void DlgCreateNewTemp::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CRENW_TXT_NAME, m_txt_name);
	DDX_Control(pDX, CRENW_BTN_OK, m_btn_ok);
	DDX_Control(pDX, CRENW_BTN_CANCEL, m_btn_cancel);
}

BEGIN_MESSAGE_MAP(DlgCreateNewTemp, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_BN_CLICKED(CRENW_BTN_OK, &DlgCreateNewTemp::OnBnClickedBtnOk)
	ON_BN_CLICKED(CRENW_BTN_CANCEL, &DlgCreateNewTemp::OnBnClickedBtnCancel)
	ON_EN_CHANGE(CRENW_TXT_NAME, &DlgCreateNewTemp::OnEnChangeTxtName)
END_MESSAGE_MAP()

// DlgCreateNewTemp message handlers
BOOL DlgCreateNewTemp::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_ADD_FILE));
	SetIcon(hIcon, FALSE);

	m_btn_ok.SetThemeHelper(&m_ThemeHelper);
	m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

	CString szval = L"";
	szval.Format(L"Sheet%d", (int)ff->_RandomTime());
	
	m_txt_name.SubclassDlgItem(CRENW_TXT_NAME, this);
	m_txt_name.SetBkColor(rgbWhite);
	m_txt_name.SetTextColor(rgbBlue);
	m_txt_name.SetCueBanner(L"Đặt tên bảng tính mẫu mới vào đây");
	m_txt_name.SetWindowText(szval);

	return TRUE;
}

BOOL DlgCreateNewTemp::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(CRENW_BTN_OK) 
				|| pWndWithFocus == GetDlgItem(CRENW_TXT_NAME))
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
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(CRENW_TXT_NAME))
		{
			if (chr == 0x5C || chr == 0x2F || chr == 0x3A || chr == 0x2A || 
				chr == 0x3F || chr == 0x22 || chr == 0x3C || chr == 0x3e || chr == 0x7C)
			{
				m_txt_name.ShowBalloonTip(L"Hướng dẫn",
					L"Đặt tên không chứa ký tự đặc biệt "
					L"( \\ / : * ? \" < > | ) và tối đa 30 ký tự.", TTI_INFO);
				return TRUE;
			}
		}
	}

	return FALSE;
}

void DlgCreateNewTemp::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void DlgCreateNewTemp::OnBnClickedBtnOk()
{
	CString szName = L"";
	m_txt_name.GetWindowTextW(szName);
	szName = ff->_ConvertRename(szName);
	if (szName != L"")
	{
		int icheck = 0;
		int nHidden = 0;
		CString szval = L"";
		vector<_WorksheetPtr> vecPsh;
		int ncountSheet = ff->_xlGetAllSheet(vecPsh, nHidden);
		for (int i = 0; i < ncountSheet; i++)
		{
			szval = (LPCTSTR)vecPsh[i]->Name;
			if (szval == szName)
			{
				icheck = 1;
				break;
			}
		}
		vecPsh.clear();

		if (icheck == 0) _CreateNewSheet(szName);
		else
		{
			ff->_msgbox(L"Tên bảng tính đã tồn tại.", 0, 0);
			GotoDlgCtrl(GetDlgItem(CRENW_TXT_NAME));
			return;
		}
	}
	else
	{
		ff->_msgbox(L"Bạn chưa đặt tên cho bảng tính mới.", 0, 2);
		GotoDlgCtrl(GetDlgItem(CRENW_TXT_NAME));
		return;
	}

	CDialogEx::OnOK();
}


void DlgCreateNewTemp::OnBnClickedBtnCancel()
{
	CDialogEx::OnCancel();
}


void DlgCreateNewTemp::_XacdinhSheetLienquan()
{
	bsFManager = ff->_xlGetNameSheet(shFManager, 0);
	psFManager = xl->Sheets->GetItem((_bstr_t)bsFManager);
	prFManager = psFManager->Cells;

	if ((int)psFManager->GetVisible() != -1) psFManager->PutVisible(XlSheetVisibility::xlSheetVisible);
	psFManager->Activate();

	ff->_GetThongtinSheetFManager(icotSTT, icotLV, icotFolder, icotTen, icotType,
		icotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
}


bool DlgCreateNewTemp::_CreateNewSheet(CString szNameSheet)
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		_XacdinhSheetLienquan();

		// Xác định tất cả các sheet sử dụng
		int nHidden = 0;
		vector<_WorksheetPtr> vecPsh;
		int ncountSheet = ff->_xlGetAllSheet(vecPsh, nHidden);

		MyStrSheet MSSHEET;
		vector<MyStrSheet> vecSheet;
		for (int i = 0; i < ncountSheet; i++)
		{
			MSSHEET.szName = (LPCTSTR)vecPsh[i]->Name;
			MSSHEET.szCode = (LPCTSTR)vecPsh[i]->CodeName;

			MSSHEET.szLinhvuc = MSSHEET.szName;
			MSSHEET.szKodau = ff->_ConvertKhongDau(MSSHEET.szName);
			MSSHEET.szKodau.MakeLower();

			vecSheet.push_back(MSSHEET);
		}

		int nIndexSheet = ff->_xlGetIndexSheet(shFManager);
		int isizeSh = (int)vecSheet.size();

		int cd = 0, nm = 0, icheck = 0;
		CString szval = L"", szNewName = L"", szNewCode = L"";

		while (true)
		{
			szval = szNameSheet;
			if (szval.GetLength() > 30) szval = szval.Left(30);
			if (nm > 0) szNewName.Format(L"%s (%d)", szval, nm);
			else szNewName = szval;

			szval = ff->_ConvertKhongDau(szNewName);
			szval.MakeLower();
			icheck = 0;

			for (int j = 0; j < isizeSh; j++)
			{
				if (szval == vecSheet[j].szKodau)
				{
					icheck = 1;
					break;
				}
			}

			if (icheck == 0) break;
			nm++;
		}

		while (true)
		{
			cd++;
			szNewCode.Format(L"%s%03d", shFile, cd);

			icheck = 0;
			for (int j = 0; j < isizeSh; j++)
			{
				if (szNewCode == vecSheet[j].szCode)
				{
					icheck = 1;
					break;
				}
			}

			if (icheck == 0) break;
		}

		_WorksheetPtr pSheet;
		if (ff->_xlCreateNewSheet(pSheet, nIndexSheet, szNewCode, szNewName, RGB(146, 208, 80)))
		{
			_CopyStyleCategory(pSheet);
			pSheet->Activate();
		}

		vecPsh.clear();
		vecSheet.clear();

		ff->_xlPutScreenUpdating(true);
	}
	catch (...) {}
	return false;
}

void DlgCreateNewTemp::_CopyStyleCategory(_WorksheetPtr pSheet)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		pRange->PutWrapText(true);
		pRange->Rows->AutoFit();
		pRange->PutVerticalAlignment(xlCenter);
		pRange->Font->PutName(FontTimes);
		pRange->Font->PutSize(13);

		RangePtr PRS = pSheet->GetRange(pRange->GetItem(1, 1), pRange->GetItem(1, icotEnd));
		PRS->EntireColumn->PutHorizontalAlignment(xlCenter);

		RangePtr PRSCopy = psFManager->GetRange(prFManager->GetItem(1, 1), prFManager->GetItem(irowStart - 1, 1));
		PRSCopy->EntireRow->Copy();

		PRS = pSheet->GetRange(pRange->GetItem(1, 1), pRange->GetItem(irowStart - 1, 1));
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteValues, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteFormats, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		PRS->EntireRow->PasteSpecial(XlPasteType::xlPasteColumnWidths, XlPasteSpecialOperation::xlPasteSpecialOperationNone);
		xl->PutCutCopyMode(XlCutCopyMode(false));

		PRS = pRange->GetItem(irowStart - 1, 1);
		PRS->EntireRow->AutoFilter(1, vtMissing, XlAutoFilterOperator::xlAnd, vtMissing, TRUE);

		pSheet->Application->ActiveWindow->PutZoom(85);

		PRS = pRange->GetItem(irowStart, 1);
		PRS->Select();
	}
	catch (...) {}
}


void DlgCreateNewTemp::OnEnChangeTxtName()
{
	CString szval = L"";
	m_txt_name.GetWindowTextW(szval);
	if (szval.GetLength() > 30)
	{
		m_txt_name.SetWindowText(szval.Left(30));
		m_txt_name.SetSel(-1);
	}
}
