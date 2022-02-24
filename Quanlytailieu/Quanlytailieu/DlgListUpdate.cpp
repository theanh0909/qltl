#include "pch.h"
#include "DlgListUpdate.h"

IMPLEMENT_DYNAMIC(DlgListUpdate, CDialogEx)

DlgListUpdate::DlgListUpdate(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgListUpdate::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;
	sv = new CDownloadFileSever;
	rfile = new CReadFileCSV;

	szpathDefault = L"";

	iAutoUpdate = 1;
	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;

	colType = 1, colCate = 2;
	colRGB = 1, colIcon = 2, colMH = 3, colND = 4;
	iOnOK = 1;

	tslkq = 0, lanshow = 0;
	iStopload = 0, ibuocnhay = 50;

	iwMH = 150, iwND = 450, iDlgW = 0, iDlgH = 0;

	bFSever = false;
	nItem = -1, nSubItem = -1;
	nItemCate = -1, nSubItemCate = -1;
}

DlgListUpdate::~DlgListUpdate()
{
	delete ff;
	delete bb;
	delete sv;
	delete reg;
	delete rfile;

	vecItem.clear();
	vecLinhvuc.clear();
}

void DlgListUpdate::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, UPD_BTN_OK, m_btn_ok);
	DDX_Control(pDX, UPD_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, UPD_LIST_CATE, m_list_cate);
	DDX_Control(pDX, UPD_LIST_VIEW, m_list_view);
	DDX_Control(pDX, UPD_LBL_TOTAL, m_lbl_total);
	DDX_Control(pDX, UPD_CHECK_ALL, m_check_all);
	DDX_Control(pDX, UPD_CHECK_SKIP, m_check_skip);
}

BEGIN_MESSAGE_MAP(DlgListUpdate, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SIZING()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()

	ON_BN_CLICKED(UPD_BTN_OK, &DlgListUpdate::OnBnClickedBtnOk)
	ON_BN_CLICKED(UPD_BTN_CANCEL, &DlgListUpdate::OnBnClickedBtnCancel)
	ON_BN_CLICKED(UPD_CHECK_ALL, &DlgListUpdate::OnBnClickedCheckAll)
	
	ON_COMMAND(MN_LSUP_CHECK, &DlgListUpdate::OnLsupCheck)
	ON_COMMAND(MN_LSUP_UNCHECK, &DlgListUpdate::OnLsupUncheck)
	ON_COMMAND(MN_LSUP_CHECKALL, &DlgListUpdate::OnLsupCheckall)

	ON_NOTIFY(NM_CLICK, UPD_LIST_CATE, &DlgListUpdate::OnNMClickListCate)
	ON_NOTIFY(NM_RCLICK, UPD_LIST_CATE, &DlgListUpdate::OnNMRClickListCate)
	ON_NOTIFY(NM_CLICK, UPD_LIST_VIEW, &DlgListUpdate::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, UPD_LIST_VIEW, &DlgListUpdate::OnNMRClickListView)

	ON_NOTIFY(LVN_ENDSCROLL, UPD_LIST_VIEW, &DlgListUpdate::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgListUpdate::OnHdnEndtrackListView)	
END_MESSAGE_MAP()

BEGIN_EASYSIZE_MAP(DlgListUpdate)
	EASYSIZE(UPD_LBL_TOTAL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(UPD_CHECK_ALL, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, 0)
	EASYSIZE(UPD_CHECK_SKIP, ES_BORDER, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(UPD_LIST_CATE, ES_BORDER, ES_BORDER, ES_KEEPSIZE, ES_BORDER, 0)
	EASYSIZE(UPD_LIST_VIEW, ES_BORDER, ES_BORDER, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(UPD_BTN_OK, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
	EASYSIZE(UPD_BTN_CANCEL, ES_KEEPSIZE, ES_KEEPSIZE, ES_BORDER, ES_BORDER, 0)
END_EASYSIZE_MAP

// DlgListUpdate message handlers
BOOL DlgListUpdate::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_UPDATE));
	SetIcon(hIcon, FALSE);

	mnIcon.GdiplusStartupInputConfig();

	_LoadRegistry();
	_SetDefault();
	
	_LoadListCate();
	_LoadListView();

	INIT_EASYSIZE;

	_SetLocationAndSize();
		
	return TRUE;
}

BOOL DlgListUpdate::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(UPD_BTN_OK))
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
		if (pWndWithFocus == GetDlgItem(UPD_LIST_VIEW))
		{
			if (chr == 0x01)
			{
				_SelectAllItems();
				return TRUE;
			}
		}
	}

	return FALSE;
}

void DlgListUpdate::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgListUpdate::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_LISTUPDATE));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			HBITMAP hBmp;

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_LSUP_CHECK, IDI_ICON_OK);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_LSUP_UNCHECK, IDI_ICON_CANCEL);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_LSUP_CHECKALL, IDI_ICON_TODOLIST);

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(UPD_LIST_VIEW));
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

void DlgListUpdate::OnSize(UINT nType, int cx, int cy)
{
	CDialog::OnSize(nType, cx, cy);
	UPDATE_EASYSIZE;
}

void DlgListUpdate::OnSizing(UINT fwSide, LPRECT pRect)
{
	CDialog::OnSizing(fwSide, pRect);
	EASYSIZE_MINSIZE(320, 240, fwSide, pRect);
}

/****************** Click button Ok **********************************************************/
void DlgListUpdate::OnBnClickedBtnOk()
{
	try
	{
		int ncheck = 0;
		int ncount = m_list_cate.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_cate.GetCheck(i) == 1)
			{
				ncheck++;
				break;
			}
		}

		if (ncheck == 0)
		{
			ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if ((int)m_list_view.GetCheck(i) == 1)
				{
					ncheck++;
					break;
				}
			}
		}

		if (ncheck > 0)
		{
			iOnOK = 1;
			_SaveRegistry();

			CDialogEx::OnOK();

			for (int i = 0; i < ncount; i++)
			{
				vecItem[i].iEnable = (int)m_list_view.GetCheck(i);
			}

			if (tslkq > ncount)
			{
				int icheck = 0;
				for (int i = ncount; i < tslkq; i++)
				{
					icheck = _GetCheckItemVisible(vecItem[i].vSubItem[0], -1);
					if (icheck == 1) vecItem[i].iEnable = 1;
					else vecItem[i].iEnable = 0;					
				}
			}
		}
		else
		{
			ff->_msgbox(L"Bạn chưa tích chọn dữ liệu cần cập nhật.", 2, 0);
		}
	}
	catch (...) {}
}

void DlgListUpdate::_SetLocationAndSize()
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

void DlgListUpdate::OnBnClickedBtnCancel()
{
	iOnOK = 0;
	_SaveRegistry();

	CDialogEx::OnCancel();
}

void DlgListUpdate::_SaveRegistry()
{
	try
	{
		if ((int)m_check_skip.GetCheck() == 1)
		{
			// Luôn bỏ qua update( iAutoUpdate = 2)
			CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
			reg->CreateKey(szCreate);
			reg->WriteChar(ff->_ConvertCstringToChars(DlgAutoUpdate), ff->_ConvertCstringToChars(L"2"));
		}

		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
		reg->CreateKey(szCreate);

		CString szval = L"";
		iwMH = m_list_view.GetColumnWidth(colMH);
		iwND = m_list_view.GetColumnWidth(colND);
		szval.Format(L"%d|%d", iwMH, iwND);
		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize),	ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgListUpdate::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
		reg->CreateKey(szCreate);

		int pos = -1;
		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			pos = szval.Find(L"|");
			if (pos > 0)
			{
				iwMH = _wtoi((szval.Left(pos)).Trim());
				iwND = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}

		szval = reg->ReadString(DlgSize, L"");
		if (szval != L"")
		{
			pos = szval.Find(L"x");
			if (pos > 0)
			{
				iDlgW = _wtoi((szval.Left(pos)).Trim());
				iDlgH = _wtoi((szval.Right(szval.GetLength() - pos - 1)).Trim());
			}
		}
	}
	catch (...) {}
}

void DlgListUpdate::OnBnClickedCheckAll()
{
	int icheck = m_check_all.GetCheck();
	int ncount = m_list_cate.GetItemCount();
	for (int i = 0; i < ncount; i++)
	{
		m_list_cate.SetCheck(i, icheck);
	}

	ncount = m_list_view.GetItemCount();
	for (int i = 0; i < ncount; i++)
	{
		m_list_view.SetCheck(i, icheck);
	}

	if (icheck == 1) _SetStatusKetqua(tslkq);
	else _SetStatusKetqua(0);
}

void DlgListUpdate::_SetDefault()
{
	m_lbl_total.SubclassDlgItem(UPD_LBL_TOTAL, this);
	m_lbl_total.SetTextColor(rgbBlue);

	m_btn_ok.SetThemeHelper(&m_ThemeHelper);
	m_btn_ok.SetIcon(IDI_ICON_OK, 24, 24);
	m_btn_ok.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_ok.SetBtnCursor(IDC_CURSOR_HAND);

	m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
	m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
	m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
	m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

	m_list_cate.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_cate.InsertColumn(colType, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_cate.InsertColumn(colCate, L"Lĩnh vực", LVCFMT_LEFT, 150);

	m_list_cate.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
	m_list_cate.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
		LVS_EX_LABELTIP | LVS_EX_CHECKBOXES);
	
	m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_view.InsertColumn(colRGB, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_view.InsertColumn(colIcon, L"", LVCFMT_CENTER, COLW_FIX);
	m_list_view.InsertColumn(colMH, L" Số hiệu", LVCFMT_LEFT, iwMH);
	m_list_view.InsertColumn(colND, L" Nội dung", LVCFMT_LEFT, iwND);

	m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
	m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | 
		LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_SUBITEMIMAGES);
}

int DlgListUpdate::_LoadListCate()
{
	try
	{
		_DeleteAllListCate();

		int nsize = (int)vecLinhvuc.size();
		for (int i = 0; i < nsize; i++)
		{
			m_list_cate.InsertItem(i, L"", 0);
			m_list_cate.SetCellColors(i, colType, vecLinhvuc[i].clrBkg, RGB(0, 0, 0));
			m_list_cate.SetItemText(i, colCate, vecLinhvuc[i].szTenlv);
			m_list_cate.SetCheck(i, 1);
		}

		return nsize;
	}
	catch (...) {}
	return 0;
}

void DlgListUpdate::_SetStatusKetqua(int num)
{
	CString str = L"";
	if (num > 1) str.Format(L"Tổng số dữ liệu được chọn: %d (files)", num);
	else str.Format(L"Tổng số dữ liệu được chọn: %d (file)", num);
	m_lbl_total.SetWindowText(str);

	if (num == tslkq) m_check_all.SetCheck(1);
	else m_check_all.SetCheck(0);
}

int DlgListUpdate::_LoadListView()
{
	try
	{
		_DeleteAllListView();

		m_check_all.SetCheck(true);

		tslkq = (int)vecItem.size();
		_SetStatusKetqua(tslkq);

		lanshow = 1;
		int iz = ibuocnhay;
		if (tslkq <= iz)
		{
			iz = tslkq;
			iStopload = 0;
		}
		else iStopload = 1;

		// Delete list image icon
		m_imageList.DeleteImageList();
		ASSERT(m_imageList.GetSafeHandle() == NULL);
		m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);

		// Thêm dữ liệu vào list kết quả
		int icheck = (int)m_check_all.GetCheck();
		_AddItemsListKetqua(0, iz, icheck);

		return iz;
	}
	catch (...) {}
	return 0;
}

COLORREF DlgListUpdate::_GetColorLinhvuc(CString szLinhvuc)
{
	try
	{
		int nsize = (int)vecLinhvuc.size();
		for (int i = 0; i < nsize; i++)
		{
			if (szLinhvuc == vecLinhvuc[i].szTenlv)
			{
				return vecLinhvuc[i].clrBkg;
			}
		}
	}
	catch (...) {}
	return rgbWhite;
}

void DlgListUpdate::_AddItemsListKetqua(int iBegin, int iEnd, int itype, int icheck)
{
	try
	{
		int dem = 0, ncount = 0;
		if (iBegin > 0)
		{
			ncount = m_list_view.GetItemCount();
			POSITION pos = m_list_view.GetFirstSelectedItemPosition();
			if (pos != NULL)
			{
				for (int i = 0; i < m_list_view.GetItemCount(); i++)
				{
					if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED) dem++;
				}
			}
		}

		int ncheckCate = -1;
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, colMH, vecItem[i].vSubItem[1]);		// Số hiệu

			if ((int)vecItem[i].vSubItem.size() >= 2)
			{
				m_list_view.SetItemText(i, colND, vecItem[i].vSubItem[2]);	// Tên - Nội dung
			}

			// Setcolor lĩnh vực
			m_list_view.SetCellColors(i, colRGB, _GetColorLinhvuc(vecItem[i].vSubItem[0]), RGB(0, 0, 0));

			// File gốc không có link sẽ báo đỏ
			if (vecItem[i].vSubItem[8] == L"") m_list_view.SetRowColors(i, rgbWhite, rgbRed);

			HICON hI = m_sysImageList.GetFileIcon(L"*." + ff->_GetTypeOfFile(vecItem[i].vSubItem[8]));
			m_imageList.Add(hI);

			if (itype == 0) m_list_view.SetCheck(i, icheck);
			else
			{
				ncheckCate = _GetCheckItemVisible(vecItem[i].vSubItem[0], -1);
				if (ncheckCate == -1) ncheckCate = icheck;
				m_list_view.SetCheck(i, ncheckCate);
			}

			if (iBegin > 0 && dem == ncount)
			{
				m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
			}
		}

		// Load Image vào list dữ liệu
		m_list_view.SetImageList(&m_imageList);

		for (int i = iBegin; i < iEnd; i++)
		{
			// Load hình ảnh đại diện
			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, -1, 0, 0, 0);
			m_list_view.SetItem(i, colIcon, LVIF_IMAGE, NULL, i, 0, 0, 0);
		}
	}
	catch (...) {}
}

void DlgListUpdate::_XacdinhSheetLienquan()
{
	bsConfig = ff->_xlGetNameSheet(shConfig, 0);
	psConfig = xl->Sheets->GetItem(bsConfig);
	prConfig = psConfig->Cells;

	ff->_GetThongtinSheetCategory(icotSTT, icotLV, icotMH, icotNoidung,
		icotFilegoc, icotFileGXD, icotThLuan, icotEnd, irowTieude, irowStart, irowEnd);
}

/****************** Check Update ****************************************************************/
int DlgListUpdate::_CheckUpdateData(bool blUpdate)
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		_XacdinhSheetLienquan();

		CString szval = L"";
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UploadData;
		reg->CreateKey(szCreate);

		iAutoUpdate = 1;	// iAutoUpdate = 0 Tự động | 1= Luôn hỏi | 2 = Bỏ qua
		if (!blUpdate)
		{
			szval = reg->ReadString(DlgAutoUpdate, L"");
			if (szval != L"") iAutoUpdate = _wtoi(szval.Trim());
			if (iAutoUpdate != 1 && iAutoUpdate != 2) iAutoUpdate = 0;
			if (iAutoUpdate == 2) return 0;
		}

		int iSLUpdate = 0;
		szpathDefault = _SetDirDefault();
		if (szpathDefault != L"")
		{
			bFSever = false;
			CFileFinder _finder;
			vector<CString> vecLine;

			int ncountFile = ff->_FindAllFile(_finder, szpathDefault + FolderVanban, L"csv");
			if (ncountFile > 0)
			{
				for (int i = 0; i < ncountFile; i++)
				{
					rfile->_ReadLineFile(_finder.GetFilePath(i).GetLongPath(), vecLine);
				}
			}
			else
			{
				// Tạm thời không tự động tải từ trên sever về nếu không tìm thấy
				/*sv->_LoadDecodeBase64();
				CString szFileSaved = szpathDefault + FolderVanban + FileUpdateCSV;
				if (sv->_DownloadFile(sv->dbSeverData + FolderVanban + FileUpdateCSV, szFileSaved))
				{
					bFSever = true;
					rfile->_ReadLineFile(szFileSaved, vecLine);
				}*/

				int result = ff->_y(L"Trong máy tính của bạn chưa tồn tại bộ dữ liệu Quản lý dự án nào. "
					L"Để tải và sử dụng, bạn chọn [YES] để đi đến trang chia sẻ kho dữ liệu phần mềm "
					L"Quản lý dự án GXD và làm theo hướng dẫn. ", 0, 0, 0);
				if (result == 6)
				{
					vecLine.clear();
					sv->_LoadDecodeBase64();
					CString szFileSaved = szpathDefault + FileHelp;
					if (sv->_DownloadFile(sv->dbSeverDir + FileHelp, szFileSaved))
					{
						vecLine.clear();
						int nLine = ff->_ReadMultiUnicode(szFileSaved, vecLine);
						if (nLine > 0)
						{
							for (int i = nLine - 1; i >= 0; i--)
							{
								vecLine[i].Trim();
								if (vecLine[i] == L"" || vecLine[i].Left(2) == L"//") vecLine.erase(vecLine.begin() + i);
							}
						}

						nLine = (int)vecLine.size();
						if (nLine >= 8)
						{
							ShellExecute(NULL, L"open", vecLine[7], NULL, NULL, SW_SHOWMAXIMIZED);
						}
					}

					DeleteFile(szFileSaved);
					vecLine.clear();
					return 0;
				}
			}

			int ncount = (int)vecLine.size();
			if (ncount > 0)
			{
				int num = 0, icheck = 0;
				int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;

				RangePtr PRS = pShActive->GetRange(
					pRgActive->GetItem(irowStart, icotFilegoc),
					pRgActive->GetItem(iRowLast, icotFilegoc));

				vector<CString> vecHyper;
				int ncoutHyp = ff->_xlGetAllHyperLink(PRS, vecHyper, 1, 1);

				vecItem.clear();
				vecLinhvuc.clear();

				COLORREF clrCreate[] = { 
					RGB(0,176,80),RGB(54,96,146),RGB(150,54,52),RGB(118,147,60),RGB(166,166,166),
					RGB(250,191,143),RGB(153,255,102),RGB(255,255,102),RGB(0,176,240),RGB(0,102,0),
					RGB(220,20,60),RGB(255,160,122),RGB(226,106,16),RGB(149,179,215),RGB(0,51,102),
					RGB(46,139,87),RGB(102,255,204),RGB(255,192,0),RGB(218,150,148),RGB(238,232,170),
					RGB(102,204,255),RGB(188,143,143),RGB(51,204,51),RGB(255,153,255),RGB(205,133,63),
					RGB(144,238,144),RGB(165,42,42),RGB(72,61,139),RGB(177,160,199),RGB(49,134,155),
					RGB(83,141,213),RGB(216,228,188),RGB(96,73,122),RGB(255,255,0),RGB(148,138,84),
					RGB(102,102,255),RGB(191,191,191),RGB(204,0,0),RGB(0,51,204),RGB(64,64,64),
					RGB(0,154,208),RGB(255,51,153),RGB(0,51,0),RGB(153,102,255),RGB(102,0,51),
					RGB(102,0,204),RGB(153,0,153),RGB(8,130,184),RGB(199,21,133),RGB(255,102,0)};
				int totalColor = getSizeArray(clrCreate);

				// Kiểm tra files đã tồn tại hay chưa?
				MyStrCSV MSCSV;
				MyStrLinhvuc MSLVUC;
				CString szval = L"", szcheck = L"";
				for (int i = 0; i < ncount; i++)
				{
					num = rfile->_GetItemStr(vecLine[i], MSCSV.vSubItem);
					if (num > 0)
					{
						icheck = 0;

						// Nếu file gốc không tồn tại --> lấy bằng file word gxd
						if (MSCSV.vSubItem[8] == L"") MSCSV.vSubItem[8] = MSCSV.vSubItem[9];
						if (MSCSV.vSubItem[8] != L"")
						{
							szval = MSCSV.vSubItem[8];
							szval.Replace(L"/", L"\\");
							if (ncoutHyp > 0)
							{
								for (int j = 0; j < ncoutHyp; j++)
								{
									if (vecHyper[j].Find(szval) >= 0)
									{
										icheck = 1;
										break;
									}
								}
							}
						}

						if (icheck == 0)
						{
							num = (int)vecLinhvuc.size();
							if (num > 0)
							{
								for (int j = 0; j < num; j++)
								{
									if (MSCSV.vSubItem[0] == vecLinhvuc[j].szTenlv)
									{
										icheck = 1;
										break;
									}
								}
							}
							
							if (icheck == 0)
							{
								MSLVUC.szTenlv = MSCSV.vSubItem[0];
								MSLVUC.clrBkg = clrCreate[num%totalColor];
								vecLinhvuc.push_back(MSLVUC);
							}

							MSCSV.iEnable = 1;
							vecItem.push_back(MSCSV);
						}
					}
				}
				vecHyper.clear();
				
				ncount = (int)vecItem.size();
				if (ncount > 0)
				{
					if (iAutoUpdate == 1)
					{
						// Luôn khỏi trước khi Update	
						ff->_xlPutScreenUpdating(true);
						xl->EnableCancelKey = XlEnableCancelKey(FALSE);
						AFX_MANAGE_STATE(AfxGetStaticModuleState());
						DoModal();

						// Lọc dữ liệu cập nhật
						for (int i = ncount - 1; i >= 0; i--)
						{
							if (vecItem[i].iEnable == 0)
							{
								vecItem.erase(vecItem.begin() + i);
							}
						}

						ncount = (int)vecItem.size();
						if (iOnOK == 0 || ncount == 0) return 0;
					}

					ff->_xlPutScreenUpdating(false);

					PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, 1), pRgActive->GetItem(irowStart + ncount - 1, 1));
					PRS->EntireRow->Insert(xlShiftDown);

					PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, 1), pRgActive->GetItem(irowStart + ncount - 1, 1));
					PRS->EntireRow->Font->PutName(FontTimes);
					PRS->EntireRow->Font->PutSize(13);
					PRS->EntireRow->Font->PutBold(false);
					PRS->EntireRow->Font->PutItalic(false);
					PRS->EntireRow->Font->PutColor(rgbBlack);
					PRS->EntireRow->Interior->PutColor(xlNone);

					int nColCenter[] = { icotNoidung, icotNoidung + 6, icotThLuan + 1, icotEnd };
					int nsizearr = getSizeArray(nColCenter);
					for (int i = 0; i < nsizearr; i++)
					{
						PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, nColCenter[i]), pRgActive->GetItem(irowStart + ncount - 1, nColCenter[i]));
						PRS->EntireColumn->PutHorizontalAlignment(xlLeft);
					}

					CString szUpdateStatus = L"Đang tiến hành ";
					if (bFSever) szUpdateStatus += L"tải và ";
					szUpdateStatus+= L"cập nhật dữ liệu. Vui lòng đợi trong giây lát...";

					int iTTcot[] = { icotLV, icotMH, icotNoidung, 
						icotNoidung + 1, icotNoidung + 2, icotNoidung + 3, icotNoidung + 4, icotNoidung + 5,
						icotFilegoc, icotFileGXD, icotThLuan };
					nsizearr = getSizeArray(iTTcot);

					int pos = -1, numStart = 1;
					CString szDir = L"", szpath = L"";
					CString szTenfile = L"";

					iSLUpdate = ncount;
					for (int i = 0; i < ncount; i++)
					{
						szTenfile = ff->_GetNameOfPath(vecItem[i].vSubItem[8], pos, 0);
						szval.Format(L"%s (%d%s) %s", szUpdateStatus,
							(int)(100 * i / ncount), L"%", szTenfile);
						if (szval.GetLength() > 250) szval = szval.Left(250) + L"...";
						ff->_xlPutStatus(szval);

						pRgActive->PutItem(irowStart + i, icotSTT, numStart);

						numStart++;

						for (int j = 0; j < (int)vecItem[i].vSubItem.size(); j++)
						{
							if (vecItem[i].vSubItem[j] != L"")
							{
								PRS = pRgActive->GetItem(irowStart + i, iTTcot[j]);
								PRS->PutValue2((_bstr_t)vecItem[i].vSubItem[j]);

								if (iTTcot[j] == icotFilegoc || iTTcot[j] == icotFileGXD)
								{
									szDir = szpathDefault + ff->_GetNameOfPath(vecItem[i].vSubItem[j], pos, 1);
									ff->_CreateDirectories(szDir);

									szpath = szpathDefault + vecItem[i].vSubItem[j];

									// Tạm thời comment không cho tải
									/*if (bFSever)
									{
										// Download files từ trên sever xuống
										sv->_DownloadFile(sv->dbSeverData + vecItem[i].vSubItem[j], szpath);
									}*/

									ff->_xlSetHyperlink(pShActive, PRS, szpath);

									if (ff->_FileExistsUnicode(szpath, 0) != 1)
									{
										PRS->Font->PutColor(rgbWhite);
										PRS->Interior->PutColor(rgbRed);
									}
								}
								else if (iTTcot[j] == icotThLuan)
								{
									if (vecItem[i].vSubItem[j].Find(L"http") >= 0)
									{
										ff->_xlSetHyperlink(pShActive, PRS, vecItem[i].vSubItem[j], HypDisplayGoto);
									}
									else
									{
										if (bFSever)
										{
											// Download files từ trên sever xuống
											sv->_DownloadFile(sv->dbSeverData + vecItem[i].vSubItem[j], szpath);
										}

										ff->_xlSetHyperlink(pShActive, PRS, szpath, HypDisplayGoto);
									}
								}
							}
						}
					}

					// Đánh lại STT
					_AutoSothutu(pShActive, numStart, irowStart + ncount, 0);

					PRS = pShActive->GetRange(pRgActive->GetItem(1, 1), pRgActive->GetItem(1, icotThLuan));
					PRS->PutHorizontalAlignment(xlCenterAcrossSelection);

					PRS = pShActive->GetRange(pRgActive->GetItem(irowTieude, 1), pRgActive->GetItem(irowStart - 1, 1));
					PRS->EntireRow->PutHorizontalAlignment(xlCenter);

					PRS = pShActive->GetRange(pRgActive->GetItem(irowStart, icotSTT),
						pRgActive->GetItem(irowStart + ncount - 1, icotEnd));
					ff->_xlFormatBorders(PRS, 3, true, true);
					PRS->Rows->AutoFit();
					
					ff->_xlPutStatus(L"Ready");
					ff->_msgbox(L"Cập nhật dữ liệu thành công!", 0, 1, 2000);
					ff->_xlPutScreenUpdating(true);
				}

				vecItem.clear();
			}

			int iDelFileCSV = 0;
			szval = reg->ReadString(DlgDeleteFileCSV, L"");
			if (szval != L"") iDelFileCSV = _wtoi(szval.Trim());
			if (iDelFileCSV != 1) iDelFileCSV = 0;
			if (iDelFileCSV == 1)
			{
				for (int i = 0; i < ncountFile; i++)
				{
					DeleteFile(_finder.GetFilePath(i).GetFileName());
				}
			}
			_finder.RemoveAll();

			vecLine.clear();			
		}

		if (iSLUpdate > 0) return 1;
	}
	catch (...) {}
	return 0;
}

void DlgListUpdate::_AutoSothutu(_WorksheetPtr pSheet, int numStart, int nRowStart, int nRowEnd)
{
	try
	{
		RangePtr pRange = pSheet->Cells;
		if (nRowStart == 0) nRowStart = irowStart;
		if (nRowEnd == 0) nRowEnd = pSheet->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;

		CString szval = L"";
		for (int i = nRowStart; i <= nRowEnd; i++)
		{
			szval = (ff->_xlGIOR(pRange, i, icotNoidung, L"")).Trim();
			if (szval != L"")
			{
				pRange->PutItem(i, icotSTT, numStart);
				numStart++;
			}
		}
	}
	catch (...) {}
}

CString DlgListUpdate::_SetDirDefault()
{
	try
	{
		CString pathFolder = L"";
		if (bsConfig != (_bstr_t)L"")
		{
			pathFolder = ff->_xlGetValue(psConfig, L"CF_DIR_DEFAULT", 0, 0).Trim();
			if (pathFolder == L"")
			{
				pathFolder = ff->_GetPathFolder(strNAMEDLL);
				if (pathFolder != L"")
				{
					if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";
					ff->_xlPutValue(psConfig, L"CF_DIR_DEFAULT", pathFolder, 0);
				}
			}
			else
			{
				if (pathFolder.Right(1) != L"\\") pathFolder += L"\\";
			}
		}
		return pathFolder;
	}
	catch (...) {}
	return L"";
}

void DlgListUpdate::_DeleteAllListCate()
{
	if (m_list_cate.GetItemCount() > 0)
	{
		m_list_cate.DeleteAllItems();
		ASSERT(m_list_cate.GetItemCount() == 0);
	}
}

void DlgListUpdate::_DeleteAllListView()
{
	if (m_list_view.GetItemCount() > 0)
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);
	}
}

void DlgListUpdate::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	try
	{
		if (iStopload == 0) return;

		RECT r;
		CRect rectCtrl;
		m_list_view.GetItemRect(0, &r, LVIR_BOUNDS);
		m_list_view.GetWindowRect(&rectCtrl);
		int a = r.bottom - r.top;
		int b = rectCtrl.Height();
		int topIndex = m_list_view.GetTopIndex();
		int ncount = m_list_view.GetItemCount();
		if (topIndex + (int)(b / a) >= ncount)
		{
			lanshow++;
			int iz = ibuocnhay * lanshow;
			if (tslkq <= iz)
			{
				iz = tslkq;
				iStopload = 0;
			}
			else iStopload = 1;

			// Thêm dữ liệu vào list kết quả
			int icheck = (int)m_check_all.GetCheck();
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz, icheck);
		}
	}
	catch (...) {}
}

int DlgListUpdate::_GetCheckItemVisible(CString szLinhvuc, int itype)
{
	try
	{
		CString szval = L"";
		int nsize = m_list_cate.GetItemCount();
		for (int i = 0; i < nsize; i++)
		{
			szval = m_list_cate.GetItemText(i, colCate);
			if (szval == szLinhvuc)
			{
				if (itype >= 0) return itype;
				else return (int)m_list_cate.GetCheck(i);
			}
		}
	}
	catch (...) {}
	return -1;
}

void DlgListUpdate::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0 || pSubItem == 1) pNMListHeader->pitem->cxy = COLW_FIX;
	*pResult = 0;
}

void DlgListUpdate::OnNMClickListCate(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItemCate = pNMListView->iItem;
	nSubItemCate = pNMListView->iSubItem;

	if (nItemCate >= 0 && nSubItemCate == 0)
	{
		bool bl = !m_list_cate.GetCheck(nItemCate);
		CString szLinhvuc = m_list_cate.GetItemText(nItemCate, colCate);
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if (szLinhvuc == vecItem[i].vSubItem[0])
			{
				m_list_view.SetCheck(i, bl);
			}
		}

		int dem = _GetCountItemsCheck(szLinhvuc, (int)bl);
		_SetStatusKetqua(dem);
	}
}

void DlgListUpdate::OnNMRClickListCate(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListCate(pNMHDR, pResult);
}

void DlgListUpdate::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem == 0)
	{
		bool bl = !m_list_view.GetCheck(nItem);
		int dem = _GetCountItemsCheck(L"", -1);

		if (bl) dem++;
		else dem--;

		_SetStatusKetqua(dem);
	}
}

void DlgListUpdate::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}

int DlgListUpdate::_GetCountItemsCheck(CString szLinhvuc, int iType)
{
	try
	{
		int dem = 0;
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1) dem++;
		}

		if (tslkq > ncount)
		{
			int icheck = 0;
			int nItype = iType;
			for (int i = ncount; i < tslkq; i++)
			{
				nItype = iType;
				if (szLinhvuc != vecItem[i].vSubItem[0]) nItype = -1;
				icheck = _GetCheckItemVisible(vecItem[i].vSubItem[0], nItype);
				if (icheck == 1) dem++;
			}
		}

		return dem;
	}
	catch (...) {}
	return 0;
}

void DlgListUpdate::_SelectAllItems()
{
	int ncount = m_list_view.GetItemCount();
	if (ncount > 0)
	{
		for (int i = 0; i < ncount; i++)
		{
			m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
		}
	}
}

void DlgListUpdate::_GetItemSelect(int icheck)
{
	try
	{
		POSITION pos = m_list_view.GetFirstSelectedItemPosition();
		if (pos != NULL)
		{
			int dem = 0;
			int ncount = m_list_view.GetItemCount();
			for (int i = 0; i < ncount; i++)
			{
				if (m_list_view.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED)
				{
					m_list_view.SetCheck(i, icheck);
					dem++;
				}
			}

			if (dem == ncount)
			{
				ncount = m_list_cate.GetItemCount();
				for (int i = 0; i < ncount; i++)
				{
					m_list_cate.SetCheck(i, icheck);
				}

				if (icheck == 1) _SetStatusKetqua(tslkq);
				else _SetStatusKetqua(0);
			}
			else
			{
				dem = _GetCountItemsCheck(L"", -1);
				_SetStatusKetqua(dem);
			}
		}
	}
	catch (...) {}
}

void DlgListUpdate::OnLsupCheck()
{
	_GetItemSelect(1);
}


void DlgListUpdate::OnLsupUncheck()
{
	_GetItemSelect(0);
}


void DlgListUpdate::OnLsupCheckall()
{
	_SelectAllItems();
}
