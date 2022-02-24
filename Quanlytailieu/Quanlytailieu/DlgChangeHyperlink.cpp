#include "pch.h"
#include "DlgChangeHyperlink.h"

IMPLEMENT_DYNAMIC(DlgChangeHyperlink, CDialogEx)

DlgChangeHyperlink::DlgChangeHyperlink(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgChangeHyperlink::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	pathFolderDir = L"";
	szDirOld = L"", szDirNew = L"";

	iColumnPath = 0, iColumnFile = 0;
	iGFLType = 1, iGFLPathOld = 2, iGFLPathNew = 3;

	nTotalCol = getSizeArray(iwCol);
	iwCol[1] = 0, iwCol[2] = 350, iwCol[3] = 350;
	iDlgW = 0, iDlgH = 0;
	nItem = -1, nSubItem = -1;
}

DlgChangeHyperlink::~DlgChangeHyperlink()
{
	delete ff;
	delete bb;
	delete reg;

	vecREP.clear();
}

void DlgChangeHyperlink::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CHLINK_TXT_OLD, m_txt_dirOld);
	DDX_Control(pDX, CHLINK_TXT_NEW, m_txt_dirNew);
	DDX_Control(pDX, CHLINK_LIST_VIEW, m_list_view);
	DDX_Control(pDX, CHLINK_BTN_OK, m_btn_ok);
	DDX_Control(pDX, CHLINK_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, CHLINK_BTN_ADD, m_btn_add);
	DDX_Control(pDX, CHLINK_BTN_DEL, m_btn_del);
	DDX_Control(pDX, CHLINK_BTN_UP, m_btn_up);
	DDX_Control(pDX, CHLINK_BTN_DOWN, m_btn_down);
	DDX_Control(pDX, CHLINK_LBL_HDAN, m_lbl_hdan);
}

BEGIN_MESSAGE_MAP(DlgChangeHyperlink, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(CHLINK_BTN_OK, &DlgChangeHyperlink::OnBnClickedBtnOk)
	ON_BN_CLICKED(CHLINK_BTN_CANCEL, &DlgChangeHyperlink::OnBnClickedBtnCancel)
	ON_BN_CLICKED(CHLINK_BTN_OLD, &DlgChangeHyperlink::OnBnClickedBtnOld)
	ON_BN_CLICKED(CHLINK_BTN_NEW, &DlgChangeHyperlink::OnBnClickedBtnNew)
	ON_BN_CLICKED(CHLINK_BTN_ADD, &DlgChangeHyperlink::OnBnClickedBtnAdd)	
	ON_BN_CLICKED(CHLINK_BTN_DEL, &DlgChangeHyperlink::OnBnClickedBtnDel)
	ON_BN_CLICKED(CHLINK_BTN_UP, &DlgChangeHyperlink::OnBnClickedBtnUp)
	ON_BN_CLICKED(CHLINK_BTN_DOWN, &DlgChangeHyperlink::OnBnClickedBtnDown)
	ON_COMMAND(MN_REP_ADD, &DlgChangeHyperlink::OnRepAdd)
	ON_COMMAND(MN_REP_DEL, &DlgChangeHyperlink::OnRepDel)
	ON_COMMAND(MN_REP_UP, &DlgChangeHyperlink::OnRepUp)
	ON_COMMAND(MN_REP_DOWN, &DlgChangeHyperlink::OnRepDown)
	ON_NOTIFY(NM_CLICK, CHLINK_LIST_VIEW, &DlgChangeHyperlink::OnNMClickListView)
	ON_NOTIFY(NM_DBLCLK, CHLINK_LIST_VIEW, &DlgChangeHyperlink::OnNMDblclkListView)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgChangeHyperlink)
	EASYSIZE(CHLINK_LBL_OLD, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_LBL_NEW, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	
	EASYSIZE(CHLINK_TXT_OLD, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_TXT_NEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)

	EASYSIZE(CHLINK_BTN_OLD, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_BTN_NEW, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	
	EASYSIZE(CHLINK_BTN_ADD, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_BTN_DEL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_BTN_UP, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_BTN_DOWN, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)

	EASYSIZE(CHLINK_LBL_HDAN, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(CHLINK_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	
	EASYSIZE(CHLINK_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(CHLINK_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgChangeHyperlink message handlers
BOOL DlgChangeHyperlink::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_REPLACE));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_SetDefault();
	_LoadListReplace();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgChangeHyperlink::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(CHLINK_BTN_OK))
			{
				OnBnClickedBtnOk();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(CHLINK_TXT_OLD))
			{
				GotoDlgCtrl(GetDlgItem(CHLINK_TXT_NEW));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(CHLINK_TXT_NEW))
			{
				GotoDlgCtrl(GetDlgItem(CHLINK_TXT_OLD));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			OnBnClickedBtnCancel();
			return TRUE;
		}
	}	

	return FALSE;
}

void DlgChangeHyperlink::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgChangeHyperlink::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgChangeHyperlink::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgChangeHyperlink::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_REPLACE));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_REP_ADD, IDI_ICON_ADD);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_REP_DEL, IDI_ICON_DELETE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_REP_UP, IDI_ICON_UP);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_REP_DOWN, IDI_ICON_DOWN);
			
			vector<int> vecRow;
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_REP_UP, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_REP_DOWN, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(CHLINK_LIST_VIEW));
			pListview->GetWindowRect(&rectSubmitButton);
			if (rectSubmitButton.PtInRect(point))
			{
				TrackPopupMenu(pMenu, TPM_LEFTALIGN | TPM_TOPALIGN, point.x, point.y, 0, this->GetSafeHwnd(), NULL);
			}

			if (hBmp) ::DeleteObject(hBmp);
		}

		if (hMMenu) VERIFY(DestroyMenu(hMMenu));
	}
	catch (...) {}
}

void DlgChangeHyperlink::OnBnClickedBtnOk()
{
	try
	{
		vecREP.clear();

		MyStrRep MSREP;
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1)
			{
				MSREP.szOLD = m_list_view.GetItemText(i, iGFLPathOld);
				MSREP.szNEW = m_list_view.GetItemText(i, iGFLPathNew);

				if (MSREP.szOLD != L"" || MSREP.szNEW != L"")
				{
					vecREP.push_back(MSREP);
				}			
			}
		}

		int nsize = (int)vecREP.size();
		if (nsize == 0)
		{
			ff->_msgbox(L"Bạn chưa lựa chọn đường dẫn thay thế. Vui lòng kiểm tra lại.", 0, 2);
			return;
		}

		_SaveRegistry();
		_SaveListReplace();

		CDialogEx::OnOK();

/************************************************************************************/

		_WorksheetPtr psh;
		int irowStart = 1, irowEnd = 1;
		iColumnPath = 0, iColumnFile = 0;

		CString szCode = sCodeActive;
		CString szCopy = shCopy, szFile = shFile;

		if (szCode == shTieuchuan)
		{
			psh = xl->Sheets->GetItem(ff->_xlGetNameSheet(shTieuchuan, 1));
			iColumnFile = ff->_xlGetColumn(psh, L"TCH_OPEN", -1);
			irowStart = ff->_xlGetRow(psh, L"TCH_OPEN", 0) + 1;
		}
		else if (szCode == shCategory || szCode.Left(szCopy.GetLength()) == szCopy)
		{
			psh = xl->Sheets->GetItem(ff->_xlGetNameSheet(shCategory, 1));
			iColumnFile = ff->_xlGetColumn(psh, L"CT_FILEGOC", -1);
			irowStart = ff->_xlGetRow(psh, L"CT_FILEGOC", 0) + 1;
		}
		else
		{
			psh = xl->Sheets->GetItem(ff->_xlGetNameSheet(shFManager, 1));
			iColumnPath = ff->_xlGetColumn(psh, L"MNF_DIR", -1);
			iColumnFile = ff->_xlGetColumn(psh, L"MNF_LINKGOC", -1);
			irowStart = ff->_xlGetRow(psh, L"MNF_LINKGOC", 0) + 1;
		}

		irowEnd = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;

		CString szUpdateStatus = L"Đang tiến hành thay thế link. Vui lòng đợi trong giây lát...";
		ff->_xlPutScreenUpdating(false);

		RangePtr PRS;

		int dem = 0;
		ncount = irowEnd - irowStart;
		CString szval = L"";

		for (int i = irowStart; i < irowEnd; i++)
		{
			dem++;
			szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100*dem/ncount), L"%");
			ff->_xlPutStatus(szval);

			_ReplaceHyperlink(i);
		}
		
		ff->_xlPutScreenUpdating(true);
		ff->_xlPutStatus(L"Ready");
		vecREP.clear();
	}
	catch (...) {}
}

void DlgChangeHyperlink::_ReplaceHyperlink(int nRow)
{
	try
	{
		RangePtr PRS;
		CString szval = ff->_xlGIOR(pRgActive, nRow, iColumnFile, L"");
		if (iColumnPath == 0 && szval == L"") return;
		
		// Replace hyperlink
		int nsize = (int)vecREP.size();
		PRS = pRgActive->GetItem(nRow, iColumnFile);
		CString szHyper = ff->_xlGetHyperLink(PRS, 1);
		if (szHyper != L"")
		{
			for (int i = 0; i < nsize; i++)
			{
				if (vecREP[i].szOLD != L"")
				{
					szHyper.Replace(vecREP[i].szOLD, vecREP[i].szNEW);
				}
				else
				{
					szHyper = vecREP[i].szNEW + szHyper;
				}
			}

			ff->_xlSetHyperlink(pShActive, PRS, szHyper);

			if (ff->_FileExistsUnicode(szHyper, 0) != 1)
			{
				PRS->Font->PutColor(rgbWhite);
				PRS->Interior->PutColor(rgbRed);
			}
		}

		// Replace path
		if (iColumnPath > 0)
		{
			szval = ff->_xlGIOR(pRgActive, nRow, iColumnPath, L"");
			for (int i = 0; i < nsize; i++)
			{
				if (vecREP[i].szOLD != L"")
				{
					szval.Replace(vecREP[i].szOLD, vecREP[i].szNEW);
				}
				else
				{
					szval = vecREP[i].szNEW + szval;
				}
			}

			pRgActive->PutItem(nRow, iColumnPath, (_bstr_t)szval);
		}
	}
	catch (...) {}
}

void DlgChangeHyperlink::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	_SaveListReplace();

	CDialogEx::OnCancel();
}

void DlgChangeHyperlink::OnBnClickedBtnOld()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần tìm kiếm...";
	if (szDirOld != L"") m_dlg.m_ofn.lpstrInitialDir = szDirOld;
	if (m_dlg.DoModal() == IDOK)
	{
		szDirOld = m_dlg.GetPathName();
		m_txt_dirOld.SetWindowText(szDirOld);
	}
}

void DlgChangeHyperlink::OnBnClickedBtnNew()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn thư mục cần tìm kiếm...";
	if (szDirNew != L"") m_dlg.m_ofn.lpstrInitialDir = szDirNew;
	if (m_dlg.DoModal() == IDOK)
	{
		szDirNew = m_dlg.GetPathName();
		m_txt_dirNew.SetWindowText(szDirNew);
	}
}

void DlgChangeHyperlink::_SetDefault()
{
	try
	{
		CString szval = ff->_GetDesktopDir();

		m_lbl_hdan.SubclassDlgItem(CHLINK_LBL_HDAN, this);
		m_lbl_hdan.SetFont(&m_FontHeaderList);
		m_lbl_hdan.SetTextColor(rgbBlue);

		if (szDirOld == L"") szDirOld = szval;
		m_txt_dirOld.SubclassDlgItem(CHLINK_TXT_OLD, this);
		m_txt_dirOld.SetBkColor(rgbWhite);
		m_txt_dirOld.SetTextColor(rgbBlue);
		m_txt_dirOld.SetWindowText(szDirOld);
		m_txt_dirOld.SetCueBanner(L"Chọn đường dẫn thư mục cũ (Có thể bỏ trống)...");

		if (szDirNew == L"") szDirNew = szval;
		m_txt_dirNew.SubclassDlgItem(CHLINK_TXT_NEW, this);
		m_txt_dirNew.SetBkColor(rgbWhite);
		m_txt_dirNew.SetTextColor(rgbBlue);
		m_txt_dirNew.SetWindowText(szDirNew);
		m_txt_dirNew.SetCueBanner(L"Chọn đường dẫn thư mục mới (Có thể bỏ trống)...");

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
		m_list_view.InsertColumn(iGFLType, L" Trạng thái", LVCFMT_CENTER, 0/*iwCol[1]*/);
		m_list_view.InsertColumn(iGFLPathOld, L" Đường dẫn cũ", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLPathNew, L" Đường dẫn mới", LVCFMT_LEFT, iwCol[3]);

		m_list_view.SetColumnColors(iGFLPathOld, rgbWhite, rgbRed);
		m_list_view.SetColumnColors(iGFLPathNew, rgbWhite, rgbBlue);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);

/**********************************************************************************/

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_add.SetThemeHelper(&m_ThemeHelper);
		m_btn_add.SetIcon(IDI_ICON_ADD, 24, 24);
		m_btn_add.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_add.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_del.SetThemeHelper(&m_ThemeHelper);
		m_btn_del.SetIcon(IDI_ICON_DELETE, 24, 24);
		m_btn_del.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_del.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_up.SetThemeHelper(&m_ThemeHelper);
		m_btn_up.SetIcon(IDI_ICON_UP, 24, 24);
		m_btn_up.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_up.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_down.SetThemeHelper(&m_ThemeHelper);
		m_btn_down.SetIcon(IDI_ICON_DOWN, 24, 24);
		m_btn_down.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_down.SetBtnCursor(IDC_CURSOR_HAND);
	}
	catch (...) {}
}

void DlgChangeHyperlink::_SaveListReplace()
{
	try
	{
		CString szval = L"";
		CString szpathSave = pathFolderDir + FileReplace;
		int ncount = m_list_view.GetItemCount();

		vector<CString> vecLine;
		for (int i = 0; i < ncount; i++)
		{
			szval.Format(L"%s|%s|%d",
				m_list_view.GetItemText(i, iGFLPathOld),
				m_list_view.GetItemText(i, iGFLPathNew),
				(int)m_list_view.GetCheck(i));

			vecLine.push_back(szval);
		}

		ff->_WriteMultiUnicode(vecLine, szpathSave, 0);
	}
	catch (...) {}
}

void DlgChangeHyperlink::_LoadListReplace()
{
	try
	{
		CString szval = L"";
		CString szpathSave = pathFolderDir + FileReplace;

		vector<CString> vecLine;
		ff->_ReadMultiUnicode(szpathSave, vecLine);

		int dem = 0;
		int nCheck = 0;
		vector<CString> vecKey;
		int ncount = (int)vecLine.size();
		if (ncount > 0)
		{
			for (int i = 0; i < ncount; i++)
			{
				vecLine[i].Trim();
				if (vecLine[i] != L"")
				{
					vecKey.clear();
					ff->_TackTukhoa(vecKey, vecLine[i], L"|", L"|", 0, 1);

					if ((int)vecKey.size() < 2)
					{
						for (int j = (int)vecKey.size(); j < 3; j++)
						{
							vecKey.push_back(L"");
						}
					}

					nCheck = _wtoi(vecKey[2]);
					if (nCheck != 0) nCheck = 1;

					m_list_view.InsertItem(dem, L"", 0);
					m_list_view.SetItemText(dem, iGFLPathOld, vecKey[0]);
					m_list_view.SetItemText(dem, iGFLPathNew, vecKey[1]);
					m_list_view.SetCheck(dem, nCheck);

					dem++;
				}
			}			
		}
		
		for (int i = 0; i < ncount; i++)
		{
			szval.Format(L"%s|%s|%d",
				m_list_view.GetItemText(i, iGFLPathOld),
				m_list_view.GetItemText(i, iGFLPathOld),
				(int)m_list_view.GetCheck(i));

			vecLine.push_back(szval);
		}

		ff->_WriteMultiUnicode(vecLine, szpathSave, 0);
	}
	catch (...) {}
}

void DlgChangeHyperlink::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CheckHyperink;
		reg->CreateKey(szCreate);

		CString szval = L"", szcol = L"";
		for (int i = 0; i < nTotalCol; i++)
		{
			szcol.Format(L"%d|", m_list_view.GetColumnWidth(i));
			szval += szcol;
		}

		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));

		m_txt_dirOld.GetWindowTextW(szDirOld);
		reg->WriteChar(ff->_ConvertCstringToChars(L"DirectoryOld"),
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szDirOld, 0)));

		m_txt_dirNew.GetWindowTextW(szDirNew);
		reg->WriteChar(ff->_ConvertCstringToChars(L"DirectoryNew"),
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szDirNew, 0)));
	}
	catch (...) {}
}

void DlgChangeHyperlink::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CheckHyperink;
		reg->CreateKey(szCreate);

		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			vector<CString> vecKey;
			ff->_TackTukhoa(vecKey, szval, L"|", L";", 0, 1);

			int ncount = (int)vecKey.size();
			if (ncount < nTotalCol)
			{
				for (int i = ncount; i < nTotalCol; i++) vecKey.push_back(L"");
			}

			for (int i = 0; i < nTotalCol; i++) iwCol[i] = _wtoi(vecKey[i]);

			vecKey.clear();
		}

		szval = reg->ReadString(DlgSize, L"");
		if (szval != L"")
		{
			int pos = szval.Find(L"x");
			if (pos > 0)
			{
				iDlgW = _wtoi((szval.Left(pos)).Trim());
				iDlgH = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}

		szval = bb->BaseDecodeEx(reg->ReadString(L"DirectoryOld", L""));
		if (szval != L"") szDirOld = szval;

		szval = bb->BaseDecodeEx(reg->ReadString(L"DirectoryNew", L""));
		if (szval != L"") szDirNew = szval;
	}
	catch (...) {}
}

void DlgChangeHyperlink::_SetLocationAndSize()
{
	try
	{
		if (iDlgW > 0 && iDlgH > 0)
		{
			int isx = GetSystemMetrics(SM_CXSCREEN);
			int isy = GetSystemMetrics(SM_CYSCREEN);

			if (iDlgW >= isx && iDlgH + 50 >= isy)
			{
				this->ShowWindow(SW_MAXIMIZE);
			}
			else
			{
				this->SetWindowPos(NULL, 0, 0, iDlgW, iDlgH, SWP_NOMOVE | SWP_NOZORDER);

				CRect r;
				GetWindowRect(&r);
				ScreenToClient(&r);
				MoveWindow((isx - r.Width()) / 2, (isy - r.Height()) / 2, r.right - r.left, r.bottom - r.top, TRUE);
			}
		}
	}
	catch (...) {}
}


void DlgChangeHyperlink::OnBnClickedBtnAdd()
{
	try
	{
		m_txt_dirOld.GetWindowTextW(szDirOld);
		m_txt_dirNew.GetWindowTextW(szDirNew);

		szDirOld.Trim(); szDirNew.Trim();

		if (szDirOld == L""  && szDirNew == L"") return;

		// Kiểm tra dữ liệu đã trùng chưa?
		CString szval = szDirOld + L"|" + szDirNew;
		CString szcheck = L"";

		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			szcheck.Format(L"%s|%s", 
				m_list_view.GetItemText(i, iGFLPathOld),
				m_list_view.GetItemText(i, iGFLPathNew));

			if (szval == szcheck) return;
		}

		m_list_view.InsertItem(ncount, L"", 0);
		m_list_view.SetItemText(ncount, iGFLPathOld, szDirOld);
		m_list_view.SetItemText(ncount, iGFLPathNew, szDirNew);
		m_list_view.SetCheck(ncount, 1);
	}
	catch (...) {}
}

void DlgChangeHyperlink::OnRepAdd()
{
	OnBnClickedBtnAdd();
}

void DlgChangeHyperlink::OnRepDel()
{
	OnBnClickedBtnDel();
}

void DlgChangeHyperlink::OnRepUp()
{
	OnBnClickedBtnUp();
}

void DlgChangeHyperlink::OnRepDown()
{
	OnBnClickedBtnDown();
}

void DlgChangeHyperlink::OnBnClickedBtnDel()
{
	vector<int> vecRow;
	int nSelect = _GetAllSelectedItems(vecRow);
	if (nSelect > 0)
	{
		for (int i = nSelect - 1; i >= 0; i--)
		{
			m_list_view.DeleteItem(vecRow[i]);
		}
	}

	vecRow.clear();
}

void DlgChangeHyperlink::OnBnClickedBtnUp()
{
	if (nItem > 0)
	{
		CString szval = L"";
		int iArr[] = { iGFLPathOld, iGFLPathNew };
		int ncount = sizeof(iArr) / sizeof(iArr[0]);

		for (int i = 0; i < ncount; i++)
		{
			szval = m_list_view.GetItemText(nItem, iArr[i]);
			m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem - 1, iArr[i]));
			m_list_view.SetItemText(nItem - 1, iArr[i], szval);
		}

		m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
		m_list_view.SetItemState(nItem - 1, LVIS_SELECTED, LVIS_SELECTED);

		nItem--;
	}
}

void DlgChangeHyperlink::OnBnClickedBtnDown()
{
	int ncount = m_list_view.GetItemCount();
	if (nItem < ncount - 1)
	{
		CString szval = L"";
		int iArr[] = { iGFLPathOld, iGFLPathNew };
		ncount = sizeof(iArr) / sizeof(iArr[0]);

		for (int i = 0; i < ncount; i++)
		{
			szval = m_list_view.GetItemText(nItem, iArr[i]);
			m_list_view.SetItemText(nItem, iArr[i], m_list_view.GetItemText(nItem + 1, iArr[i]));
			m_list_view.SetItemText(nItem + 1, iArr[i], szval);
		}

		m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
		m_list_view.SetItemState(nItem + 1, LVIS_SELECTED, LVIS_SELECTED);

		nItem++;
	}
}

int DlgChangeHyperlink::_GetAllSelectedItems(vector<int> &vecRow)
{
	try
	{
		vecRow.clear();
		if (nItem == -1 || nSubItem == -1) return 0;

		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					vecRow.push_back(i);
				}
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}


void DlgChangeHyperlink::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;
}


void DlgChangeHyperlink::OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}
