#include "pch.h"
#include "DlgStructureFolder.h"

IMPLEMENT_DYNAMIC(DlgStructureFolder, CDialogEx)

DlgStructureFolder::DlgStructureFolder(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgStructureFolder::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	pathFolderDir = L"";
	szTargetDirectory = L"";
	szTargetOutput = L"";

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;

	tclName = 0, tclDescribe = 1, tclPath = 2, tclNew = 3;
	nItem = -1, nSubItem = -1;
	iDlgW = 0, iDlgH = 0;

	nTotalCol = getSizeArray(iwCol);
	iwCol[0] = 500, iwCol[1] = 400, iwCol[2] = 200;
}

DlgStructureFolder::~DlgStructureFolder()
{
	delete ff;
	delete bb;
	delete reg;

	vecDir.clear();
	vecCreate.clear();
}

void DlgStructureFolder::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, STRC_TXT_PATH, m_cbb_file);
	DDX_Control(pDX, STRC_TXT_NAME, m_txt_name);
	DDX_Control(pDX, STRC_BTN_TEMP, m_btn_temp);	
	DDX_Control(pDX, STRC_LBL_HDAN, m_lbl_hdan);
	DDX_Control(pDX, STRC_TXT_OUTPUT, m_txt_output);	
	DDX_Control(pDX, STRC_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, STRC_BTN_OK, m_btn_ok);
}

BEGIN_MESSAGE_MAP(DlgStructureFolder, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(STRC_BTN_OK, &DlgStructureFolder::OnBnClickedBtnOk)
	ON_BN_CLICKED(STRC_BTN_CANCEL, &DlgStructureFolder::OnBnClickedBtnCancel)
	ON_BN_CLICKED(STRC_BTN_BROWSE, &DlgStructureFolder::OnBnClickedBtnBrowse)
	ON_BN_CLICKED(STRC_BTN_TEMP, &DlgStructureFolder::OnBnClickedBtnTemp)
	ON_BN_CLICKED(STRC_BTN_OUTPUT, &DlgStructureFolder::OnBnClickedBtnOutput)
	ON_CBN_SELCHANGE(STRC_TXT_PATH, &DlgStructureFolder::OnCbnSelchangeTxtPath)
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgStructureFolder)
	EASYSIZE(STRC_BTN_BROWSE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(STRC_TXT_PATH, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(STRC_TXT_NAME, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(STRC_BTN_TEMP, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(STRC_LBL_HDAN, ES_KEEPSIZE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)	
	EASYSIZE(STRC_LINE_TOP, ES_BORDER, ES_BORDER, ES_BORDER, ES_KEEPSIZE, 0)
	EASYSIZE(STRC_TREE_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(STRC_LINE_BOTTOM, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(STRC_LBL_OUTPUT, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(STRC_TXT_OUTPUT, ES_BORDER, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(STRC_BTN_OUTPUT, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(STRC_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(STRC_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgStructureFolder message handlers
BOOL DlgStructureFolder::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_CHART));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_XacdinhSheetLienquan();
	_SetDefault();
	_LoadCombobox(szTargetDirectory);

	// Load dữ liệu mẫu tự tạo
	OnBnClickedBtnTemp();

	INIT_EASYSIZE;

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgStructureFolder::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(STRC_TXT_PATH))
			{
				CString szPathFile = L"";
				m_cbb_file.GetWindowTextW(szPathFile);
				if (szPathFile != L"")
				{
					if (!_LoadFileStructure(szTargetDirectory)) m_lbl_hdan.ShowWindow(0);
				}
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(STRC_BTN_OK))
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

	return FALSE;
}

void DlgStructureFolder::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgStructureFolder::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

void DlgStructureFolder::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgStructureFolder::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{

}


void DlgStructureFolder::OnBnClickedBtnOk()
{
	try
	{
		CString szval = L"";
		m_txt_output.GetWindowTextW(szval);
		szval.Trim();

		if (szval == L"")
		{
			ff->_msgbox(L"Bạn chưa chọn đường dẫn thư mục để lưu cấu trúc.", 2, 0);
			return;
		}

		if (szval.Right(1) != L"\\") szval += L"\\";
		szTargetOutput = szval;

		_SaveRegistry();
		CDialogEx::OnOK();

/****** Tiến hành tạo cấu trúc cây thư mục **************************************/

		HTREEITEM hItem;
		vector<CString> vecKey;
		CTreeCtrl& treeParent = m_TreeWnd.GetTreeCtrl();
		hItem = treeParent.GetFirstVisibleItem();

		CString szFolderPr = L"";
		m_txt_name.GetWindowTextW(szFolderPr);
		if (szFolderPr == L"")
		{
			szFolderPr = L"STRUCTURE";
			szval = treeParent.GetItemText(hItem);
			ff->_TackTukhoa(vecKey, szval, L"\t", L"\t", 0, 1);
			if ((int)vecKey.size() > 0)
			{
				szFolderPr = vecKey[0];
				szFolderPr.Trim();
			}

			vecKey.clear();
		}		

		int ncount = (int)vecCreate.size();
		if (ncount > 0)
		{
			CString szUpdateStatus = L"Đang tiến hành tạo cấu trúc dữ liệu. Vui lòng đợi trong giây lát...";

			ff->_CreateDirectories(szTargetOutput + szFolderPr);

			for (int i = 0; i < ncount; i++)
			{
				szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / ncount), L"%");

				ff->_xlPutStatus(szval);
				ff->_CreateNewFolder(szTargetOutput + vecCreate[i]);
			}

			_TaoSheetDulieuMoi(szFolderPr);	// Tạo mới sheet dữ liệu

			vecCreate.clear();
			ff->_xlPutStatus(L"Ready");
			ff->_msgbox(L"Tạo cấu trúc thư mục dữ liệu thành công.", 0, 0, 5000);
		}		
	}
	catch (...) {}
}


void DlgStructureFolder::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

void DlgStructureFolder::_SetDefault()
{
	try
	{
		CString szval = ff->_GetDesktopDir();
		if (szTargetDirectory == L"") szTargetDirectory = szval;
		
		m_cbb_file.AdjustDroppedWidth();
		m_cbb_file.SetMode(CComboBoxExt::MODE_AUTOCOMPLETE);		
		m_cbb_file.SetCueBanner(L"Chọn đường dẫn thư mục chứa file mẫu...");

		m_txt_name.SubclassDlgItem(STRC_TXT_NAME, this);
		m_txt_name.SetBkColor(rgbWhite);
		m_txt_name.SetTextColor(rgbBlue);
		m_txt_name.SetCueBanner(L"Đặt tên sheet (tên mẫu)...");

		m_lbl_hdan.SubclassDlgItem(STRC_LBL_HDAN, this);
		m_lbl_hdan.SetFont(&m_FontHeaderList);
		m_lbl_hdan.SetTextColor(rgbBlue);

		if (szTargetOutput == L"") szTargetOutput = ff->_GetDesktopDir();
		m_txt_output.SubclassDlgItem(STRC_TXT_OUTPUT, this);
		m_txt_output.SetBkColor(rgbWhite);
		m_txt_output.SetTextColor(rgbBlue);
		m_txt_output.SetWindowText(szTargetOutput);
		m_txt_output.SetCueBanner(L"Chọn đường dẫn thư mục chứa cấu trúc cần tạo...");

/**********************************************************************************/

		m_btn_ok.SetThemeHelper(&m_ThemeHelper);
		m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);
		
		m_btn_temp.SetThemeHelper(&m_ThemeHelper);
		m_btn_temp.SetIcon(IDI_ICON_TREE_STRUCT, 24, 24);
		m_btn_temp.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_temp.SetBtnCursor(IDC_CURSOR_HAND);

/**********************************************************************************/

		// Xác định vị trí treeview
		CRect rcTreeWnd;
		CWnd* pPlaceholderWnd = GetDlgItem(STRC_TREE_VIEW);
		pPlaceholderWnd->GetWindowRect(&rcTreeWnd);
		ScreenToClient(&rcTreeWnd);
		pPlaceholderWnd->DestroyWindow();

		// Tạo multi tree
		m_TreeWnd.CreateEx(WS_EX_CLIENTEDGE, NULL, NULL, WS_CHILD | WS_VISIBLE, 
			rcTreeWnd, this, STRC_TREE_VIEW);

		CTreeCtrl& tree = m_TreeWnd.GetTreeCtrl();
		CHeaderCtrl& header = m_TreeWnd.GetHeaderCtrl();

		DWORD dwStyle = GetWindowLong(tree, GWL_STYLE);
		dwStyle |= TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_FULLROWSELECT | TVS_DISABLEDRAGDROP;
		SetWindowLong(tree, GWL_STYLE, dwStyle);

		// Icon hiển thị
		m_imageTree.Create(16, 16, ILC_MASK | ILC_COLOR32, 3, 3);
		tree.SetImageList(&m_imageTree, LVSIL_NORMAL);
		m_imageTree.Add(AfxGetApp()->LoadIcon(IDI_ICON_MAIN));
		m_imageTree.Add(AfxGetApp()->LoadIcon(IDI_ICON_OPENFOLDER));
		m_imageTree.Add(AfxGetApp()->LoadIcon(IDI_ICON_OPENFILE));

		HDITEM hditem;
		hditem.mask = HDI_TEXT | HDI_WIDTH | HDI_FORMAT;
		hditem.fmt = HDF_LEFT | HDF_STRING;

		hditem.cxy = iwCol[0];
		hditem.pszText = (_bstr_t)L"Cấu trúc thư mục";
		header.InsertItem(tclName, &hditem);

		hditem.cxy = iwCol[1];
		hditem.pszText = (_bstr_t)L"Mô tả";
		header.InsertItem(tclDescribe, &hditem);

		hditem.cxy = iwCol[2];
		hditem.pszText = (_bstr_t)L"Đường dẫn";
		header.InsertItem(tclPath, &hditem);

		hditem.cxy = 0;
		hditem.pszText = (_bstr_t)L"Tạo mới";
		header.InsertItem(tclNew, &hditem);

		header.SetFont(&m_FontHeaderList);

		m_TreeWnd.UpdateColumns();
		tree.EnableWindow(0);
	}
	catch (...) {}
}

void DlgStructureFolder::_XacdinhSheetLienquan()
{
	jcotFolder = 0, jcotType = 0, jcotNoidung = 0;

	bsFManager = ff->_xlGetNameSheet(shFManager, 0);
	psFManager = xl->Sheets->GetItem((_bstr_t)bsFManager);
	prFManager = psFManager->Cells;

	if ((int)psFManager->GetVisible() != -1) psFManager->PutVisible(XlSheetVisibility::xlSheetVisible);
	psFManager->Activate();

	ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
		jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
}

void DlgStructureFolder::_SetLocationAndSize()
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

void DlgStructureFolder::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + DirStructure;
		reg->CreateKey(szCreate);

		CString szval = L"", szcol = L"";
		CHeaderCtrl& header = m_TreeWnd.GetHeaderCtrl();
		
		HDITEM hditem;
		hditem.mask = HDI_WIDTH;

		header.GetItem(tclName, &hditem);
		szcol.Format(L"%d|", (int)hditem.cxy);
		szval += szcol;

		header.GetItem(tclDescribe, &hditem);
		szcol.Format(L"%d|", (int)hditem.cxy);
		szval += szcol;

		header.GetItem(tclPath, &hditem);
		szcol.Format(L"%d", (int)hditem.cxy);
		szval += szcol;

		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));

		m_cbb_file.GetWindowTextW(szTargetDirectory);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgDirectory),
			ff->_ConvertCstringToChars(bb->BaseEncodeEx(szTargetDirectory, 0)));
	}
	catch (...) {}
}

void DlgStructureFolder::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + DirStructure;
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

		szval = bb->BaseDecodeEx(reg->ReadString(DlgDirectory, L""));
		if (szval != L"") szTargetDirectory = szval;
	}
	catch (...) {}
}

void DlgStructureFolder::_LoadTreeView(CString szFile)
{
	try
	{
		vecCreate.clear();

		CTreeCtrl& treeParent = m_TreeWnd.GetTreeCtrl();
		treeParent.EnableWindow(true);
		treeParent.DeleteAllItems();

		int ntotal = (int)vecDir.size();
		if (ntotal == 0) return;

		// Xác định cấp cao nhất của cấu trúc
		int iLevelMax = 0;
		for (int i = 0; i < ntotal; i++)
		{
			if (vecDir[i].iLevel > iLevelMax) iLevelMax = vecDir[i].iLevel;
		}

		HTREEITEM *arrHTree = new HTREEITEM[iLevelMax + 2];

		CString szNameTree = L"";
		szNameTree.Format(L"%s\t%s\t%s\t%s", szFile, L"", L"", szFile);
		arrHTree[0] = treeParent.InsertItem(szNameTree, 0, 0, TVI_ROOT/*, TVI_SORT*/);

		vector<CString> vecKey;
		CString szval = L"", szNew = L"";		
		int nLevel = 1, nFolder = 1, nImage = 1;
		for (int i = 0; i < ntotal; i++)
		{
			if (vecDir[i].iLevel > 0)
			{
				nFolder = vecDir[i].iLevel;
				nLevel = nFolder;
				nImage = 1;

				szNew = L"";
				szval = treeParent.GetItemText(arrHTree[nLevel - 1]);
				ff->_TackTukhoa(vecKey, szval, L"\t", L"\t", 0, 1);
				if ((int)vecKey.size() >= 4)
				{
					if (vecKey[3] != L"") szNew.Format(L"%s\\%s", vecKey[3], vecDir[i].szFolder);
				}

				szNameTree.Format(L"%s\t%s\t%s\t%s",
					vecDir[i].szFolder, vecDir[i].szNdung, L"", szNew);

				if (szNew != L"") vecCreate.push_back(szNew);
			}
			else
			{
				nLevel = nFolder + 1;
				nImage = 2;

				szNameTree.Format(L"%s\t%s\t%s\t%s",
					vecDir[i].szFileName, vecDir[i].szNdung, vecDir[i].szLink, L"");
			}
			
			if (nLevel == 1) treeParent.Expand(arrHTree[nLevel], TVE_EXPAND);

			arrHTree[nLevel] = treeParent.InsertItem(
				szNameTree, nImage, nImage, arrHTree[nLevel - 1]/*, TVI_SORT*/);
		}

		treeParent.Expand(arrHTree[1], TVE_EXPAND);
		treeParent.Expand(arrHTree[0], TVE_EXPAND);

		vecKey.clear();
		delete[] arrHTree;		
	}
	catch (...) {}
}

void DlgStructureFolder::OnBnClickedBtnBrowse()
{
	try
	{
		AFileFilter	aff(L"Excel Files|*.xlsx|All Files|*.*||");
		CFileDialog dlgFile(TRUE, NULL, NULL, OFN_HIDEREADONLY | OFN_NOCHANGEDIR,
			L"Excel Files|*.xls;*.xlsm;*.xlsx|All Files|*.*||");
		dlgFile.m_ofn.lpstrTitle = _T("Lựa chọn file mẫu tạo cấu trúc thư mục...");

		if (dlgFile.DoModal() != IDOK) return;
		CString szPathFile = dlgFile.GetPathName();
		if (szPathFile != L"")
		{
			szTargetDirectory = szPathFile;
			_LoadCombobox(szTargetDirectory);

			if(!_LoadFileStructure(szTargetDirectory)) m_lbl_hdan.ShowWindow(0);
		}
	}
	catch (...) {}
}


bool DlgStructureFolder::_LoadFileStructure(CString szFile)
{
	try
	{
		m_lbl_hdan.SetWindowText(L"Đang tiến hành đọc dữ liệu. Vui lòng đợi trong giây lát...");
		m_lbl_hdan.ShowWindow(1);

		_ApplicationPtr _pApplication;
		if (FAILED(_pApplication.CreateInstance(ExcelApp))) return false;

		_variant_t varOption((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		_WorkbookPtr _pWBook = _pApplication->Workbooks->Open((_bstr_t)szFile,
			varOption, varOption, varOption, varOption, varOption, varOption,
			varOption, varOption, varOption, varOption, varOption, varOption);

		if (_pWBook == NULL) return false;
		_WorksheetPtr pSheet = _pWBook->Sheets->Item[1];
		if (pSheet == NULL) return false;
		if ((int)pSheet->GetVisible() == 2) return false;

		RangePtr pRange = pSheet->Cells;
		if (pRange == NULL) return false;

		vecDir.clear();

		// Kiểm tra có phải file mẫu tạo thư mục dữ liệu không?
		if (ff->_xlGetcomment(pRange->GetItem(1, 1)) != L"")
		{
			// Các cột dữ liệu được chỉ định lấy dữ liệu
			int icLevel = 2, icFolder = 3, icFileName = 4, icNoidung = 5, icHyperlink = 6;

			MyStrDirStructure MSDS;

			int dem = 0;
			int iRowLast = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
			for (int i = 2; i <= iRowLast; i++)
			{
				MSDS.iLevel = _wtoi(ff->_xlGIOR(pRange, i, icLevel, L""));
				MSDS.szFolder = ff->_xlGIOR(pRange, i, icFolder, L"").Trim();
				MSDS.szFileName = ff->_xlGIOR(pRange, i, icFileName, L"").Trim();
				MSDS.szNdung = ff->_xlGIOR(pRange, i, icNoidung, L"").Trim();
				MSDS.szLink = ff->_xlGIOR(pRange, i, icHyperlink, L"").Trim();

				if (MSDS.szFileName == L"") MSDS.szFileName = L"( + Thêm File )";
				if (MSDS.iLevel == 0 && MSDS.szFolder == L"" && MSDS.szNdung == L"")
				{
					if (dem++ == 50) break;
				}
				else
				{
					dem = 0;
					vecDir.push_back(MSDS);
				}
			}
		}
		else
		{
			ff->_msgbox(L"Vui lòng chọn đúng mẫu tạo cấu trúc thư mục. "
				L"Mẫu phải có ký hiệu comment tại vị trí đầu tiên của bảng tính", 0, 0);
		}

		_pWBook->Close(VARIANT_FALSE);
		_pApplication->Quit();

		int pos = -1;
		m_lbl_hdan.ShowWindow(0);
		_LoadTreeView(ff->_GetNameOfPath(szFile, pos, -1));
		return true;
	}
	catch (...) {}
	return false;
}

void DlgStructureFolder::OnBnClickedBtnTemp()
{
	try
	{
		vecDir.clear();
		MyStrDirStructure MSDS;

		int iLevel[] = { 
			1,2,0,0,0,0,0,0,2,3,0,0,0,3,0,0,0,3,2,3,0,0,0,3,0,0,0,
			3,2,3,0,0,0,3,0,0,0,3,1,2,0,0,0,0,2,0,2,0,2,3,0,2,3,0,
			3,0,3,0,3,0,2,0,2,0,2,0,2,0,2,0,2,0,1,2,0,2,0
		};

		CString szNoidung[] = {
			L"GIAI DOAN CHUAN BI DU AN",
			L"Van ban chuan bi dau tu xay dung",
			L"Quyết định chủ trương đầu tư, giấy cấp chứng nhận đăng ký đầu tư theo quy định của Luật đầu tư 118/2015/NĐ-CP ngày 12 tháng 11 năm 2015 ",
			L"Các biên bản thỏa thuận về: Bãi đổ vật liệu phế thải, bãi đúc cấu kiện bê tông, địa điểm xây dựng công trình",
			L"Quyết định phê duyệt báo cáo đánh giá tác động môi trường",
			L"Quyết định phê duyệt nhiệm vụ chuẩn bị đầu tư dự án",
			L"Quyết định giao nhiệm vụ quản lý thực hiện dự án và ủy quyền cho CĐT",
			L"Quyết định phê duyệt dự án đầu tư xây dựng công trình",
			L"Bao cao nghien cuu tien kha thi dau tu xay dung",
			L"Lap bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế cơ sở",
			L"Sơ bộ tổng mức đầu tư",
			L"Tham dinh bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế cơ sở",
			L"Sơ bộ tổng mức đầu tư",
			L"Phe duyet bao cao",
			L"Bao cao nghien cuu kha thi dau tu xay dung",
			L"Lap bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế cơ sở",
			L"Tổng mức đầu tư",
			L"Tham dinh bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế cơ sở",
			L"Tổng mức đầu tư",
			L"Phe duyet bao cao",
			L"Bao cao kinh te ky thuat",
			L"Lap bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế bản vẽ thi công",
			L"Tổng mức đầu tư",
			L"Tham dinh bao cao",
			L"Thuyết minh dự án",
			L"Thiết kế bản vẽ thi công",
			L"Tổng mức đầu tư",
			L"Phe duyet bao cao",
			L"GIAI DOAN THUC HIEN DU AN",
			L"Ho so du an",
			L"Quyết định phê duyệt dự toán gói thầu và đơn vị thực hiện",
			L"Quyết định phê duyệt kết quả đấu thầu: Khảo sát, lập dự án",
			L"Quyết định phê duyệt kết quả lựa chọn nhà thầu: Khảo sát, tổng dự toán",
			L"Quyết định phê duyệt kết quả lựa chọn nhà thầu: Thi công xây lắp",
			L"Thuc hien thu tuc ve dat dai",
			L"Giao đất hoặc thuê đất",
			L"Chuan bi mat bang ra pha bom min",
			L"Chuẩn bị mặt bằng xây dựng, rà phá bom mìn",
			L"Khao sat xay dung",
			L"Du toan",
			L"Dự toán Khảo sát",
			L"Du toan",
			L"Lap",
			L"Dự toán xây dựng công trình",
			L"Tham tra",
			L"Dự toán xây dựng công trình sau thẩm tra",
			L"Tham dinh",
			L"Dự toán xây dựng công trình sau thẩm định",
			L"Phe duyet",
			L"Dự toán xây dựng công trình phê duyệt",
			L"Giay phep xay dung",
			L"",
			L"Dau thau va lua chon nha thau",
			L"",
			L"Ho so hoan cong",
			L"",
			L"Ho so giam sat",
			L"",
			L"Ho so thanh toan",
			L"",
			L"Ban giao, chay thu va cong viec khac",
			L"",
			L"GIAI DOAN KET THUC XAY DUNG",
			L"Quyet toan hop dong",
			L"",
			L"Bao hanh cong trinh",
			L""
		};

		// Chú ý số lượng phần tử của 2 mảng 'iLevel' & 'szNoidung' phải khớp nhau.
		int nTotal = getSizeArray(iLevel);

		CString szval = L"";
		for (int i = 0; i < nTotal; i++)
		{
			MSDS.iLevel = iLevel[i];

			if (iLevel[i] > 0)
			{
				MSDS.szFolder = szNoidung[i];
				MSDS.szFileName = L"";
				MSDS.szNdung = L"";
				MSDS.szLink = L"";
			}
			else
			{
				MSDS.szFolder = L"";
				MSDS.szFileName = L"( + Thêm file )";
				MSDS.szNdung = szNoidung[i];
				MSDS.szLink = L"";
			}

			vecDir.push_back(MSDS);
		}

		_LoadTreeView(L"Cấu trúc mẫu");
	}
	catch (...) {}
}


void DlgStructureFolder::OnBnClickedBtnOutput()
{
	CFolderPickerDialog m_dlg;
	m_dlg.m_ofn.lpstrTitle = L"Kích chọn đường dẫn thư mục chứa cấu trúc cần tạo...";
	if (szTargetOutput != L"") m_dlg.m_ofn.lpstrInitialDir = szTargetOutput;
	if (m_dlg.DoModal() == IDOK)
	{
		szTargetOutput = m_dlg.GetPathName();
		m_txt_output.SetWindowText(szTargetOutput);
	}
}

void DlgStructureFolder::_TaoSheetDulieuMoi(CString szFileName)
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		CString szName = ff->_ConvertRename(szFileName);
		if (szName.GetLength() > 27) szName = szName.Left(27);

		int dem = 0;
		int icheck = 0;
		CString szval = L"";
		vector<_WorksheetPtr> vecPsh;

		int nHidden = 0;
		int ncountSheet = ff->_xlGetAllSheet(vecPsh, nHidden);
		CString szNew = szName;		
		while (true)
		{
			icheck = 0;
			for (int i = 0; i < ncountSheet; i++)
			{
				szval = (LPCTSTR)vecPsh[i]->Name;
				if (szval == szNew)
				{
					icheck = 1;
					break;
				}
			}
			if (icheck == 0) break;

			dem++;
			szNew.Format(L"%s_%d", szName, dem);
		}
		vecPsh.clear();

		_CreateNewSheet(szNew);	

		ff->_xlPutScreenUpdating(true);
	}
	catch (...) {}
}

bool DlgStructureFolder::_CreateNewSheet(CString szNameSheet)
{
	try
	{
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

		int nIndexSheet = 1 + ff->_xlGetIndexSheet(shFManager);
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

			RangePtr pRange = pSheet->Cells;
			RangePtr PRS;

			vector<int> vecVtGroup;

			// Đổ dữ liệu ra sheet mới được tạo
			int dem = 0;
			int ncount = (int)vecDir.size();
			for (int i = 0; i < ncount; i++)
			{
				pRange->PutItem(irowStart + i, jcotLV, (_bstr_t)vecDir[i].szFolder);
				if (jcotNoidung > 0)
				{
					pRange->PutItem(irowStart + i, jcotNoidung, (_bstr_t)vecDir[i].szNdung);
				}

				if (jcotFolder > 0)
				{
					if (vecDir[i].iLevel > 0)
					{
						szval = szTargetOutput + vecCreate[dem];
						pRange->PutItem(irowStart + i, jcotFolder, (_bstr_t)szval);
						dem++;
					}
					
				}

				if (vecDir[i].iLevel == 1)
				{
					vecVtGroup.push_back(irowStart + i);
					PRS = pSheet->GetRange(pRange->GetItem(irowStart + i, 1),	pRange->GetItem(irowStart + i, icotEnd));
					PRS->Interior->PutColor(RGB(153, 255, 204));
					PRS->Font->PutBold(true);
				}
				else if (vecDir[i].iLevel == 2)
				{
					PRS = pSheet->GetRange(pRange->GetItem(irowStart + i, 1), pRange->GetItem(irowStart + i, icotEnd));
					PRS->Interior->PutColor(RGB(255, 204, 204));
					PRS->Font->PutBold(true);
				}
			}

			PRS = pSheet->GetRange(pRange->GetItem(irowStart, jcotLV),
				pRange->GetItem(irowStart + ncount - 1, jcotLV));
			PRS->PutWrapText(false);
			PRS->PutHorizontalAlignment(xlLeft);			

			PRS = pSheet->GetRange(pRange->GetItem(irowStart, jcotTen),
				pRange->GetItem(irowStart + ncount - 1, jcotTen));
			PRS->PutHorizontalAlignment(xlLeft);

			if (jcotNoidung > 0)
			{
				PRS = pSheet->GetRange(pRange->GetItem(irowStart, jcotNoidung),
					pRange->GetItem(irowStart + ncount - 1, jcotNoidung));
				PRS->PutHorizontalAlignment(xlLeft);
			}

			if (jcotFolder > 0)
			{
				PRS = pSheet->GetRange(pRange->GetItem(irowStart, jcotFolder),
					pRange->GetItem(irowStart + ncount - 1, jcotFolder));
				PRS->PutWrapText(false);
				PRS->PutHorizontalAlignment(xlLeft);
			}

			PRS = pSheet->GetRange(pRange->GetItem(irowStart, 1),
				pRange->GetItem(irowStart + ncount - 1, icotEnd));
			ff->_xlFormatBorders(PRS, 1, true, true);
			PRS->Rows->AutoFit();

			vecVtGroup.push_back(irowStart + ncount);

			ncount = (int)vecVtGroup.size();
			if (ncount >= 2)
			{
				pSheet->Outline->PutAutomaticStyles(false);
				pSheet->Outline->PutSummaryRow(xlSummaryAbove);
				pSheet->Outline->PutSummaryColumn(xlSummaryOnLeft);

				for (int i = ncount - 1; i > 0; i--)
				{
					PRS = pSheet->GetRange(
						pRange->GetItem(vecVtGroup[i - 1] + 1, 1), 
						pRange->GetItem(vecVtGroup[i] - 1, 1));
					PRS->EntireRow->Rows->Group();
				}
			}
		}

		vecPsh.clear();
		vecSheet.clear();
	}
	catch (...) {}
	return false;
}

void DlgStructureFolder::_CopyStyleCategory(_WorksheetPtr pSheet)
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

void DlgStructureFolder::_LoadCombobox(CString szPath)
{
	try
	{
		m_cbb_file.ResetContent();
		ASSERT(m_cbb_file.GetCount() == 0);

		int pos = -1;
		CString szDir = ff->_GetNameOfPath(szPath, pos, 1);

		CString str = L"";
		CFileFinder _finder;
		int ncountFile = ff->_FindAllFile(_finder, szDir, L"xlsx");
		if (ncountFile > 0)
		{
			for (int i = 0; i < ncountFile; i++)
			{
				str = _finder.GetFilePath(i).GetLongPath();
				m_cbb_file.AddString(str);
			}
		}
		_finder.RemoveAll();

		m_cbb_file.SetWindowText(szPath);
	}
	catch (...) {}
}

void DlgStructureFolder::OnCbnSelchangeTxtPath()
{
	CString szval = L"";
	m_cbb_file.GetWindowTextW(szval);
	if (szval != L"")
	{
		szTargetDirectory = szval;
		if (!_LoadFileStructure(szTargetDirectory)) m_lbl_hdan.ShowWindow(0);
	}	
}
