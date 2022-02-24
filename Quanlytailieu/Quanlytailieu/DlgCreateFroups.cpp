#include "pch.h"
#include "DlgCreateFroups.h"

IMPLEMENT_DYNAMIC(DlgCreateFroups, CDialogEx)

DlgCreateFroups::DlgCreateFroups(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgCreateFroups::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;

	iStopload = 1, lanshow = 1, ibuocnhay = 50;
	tongkq = 0;

	iGFLTten = 1, iGFLSlg = 2;
	nItem = -1, nSubItem = -1;

	pathFolderDir = L"";

	iwColTen = 200, iwColSL = 80;
	iDlgW = 0, iDlgH = 0;
	bChecked = 1;
}

DlgCreateFroups::~DlgCreateFroups()
{
	delete ff;
	delete bb;
	delete reg;

	vecGroups.clear();
}

void DlgCreateFroups::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, CREGRS_BTN_OK, m_btn_ok);
	DDX_Control(pDX, CREGRS_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, CREGRS_LIST_VIEW, m_list_view);

}

BEGIN_MESSAGE_MAP(DlgCreateFroups, CDialogEx)
	ON_WM_SYSCOMMAND()

	ON_BN_CLICKED(CREGRS_BTN_OK, &DlgCreateFroups::OnBnClickedBtnOk)
	ON_BN_CLICKED(CREGRS_BTN_CANCEL, &DlgCreateFroups::OnBnClickedBtnCancel)
	ON_NOTIFY(LVN_ENDSCROLL, CREGRS_LIST_VIEW, &DlgCreateFroups::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ITEMSTATEICONCLICK, 0, &DlgCreateFroups::OnHdnItemStateIconClickListView)
END_MESSAGE_MAP()

// DlgCreateFroups message handlers
BOOL DlgCreateFroups::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_TREE));
	SetIcon(hIcon, FALSE);

	_LoadRegistry();
	_SetDefault();
	_LoadDanhsachNhom();

	_SetLocationAndSize();

	return TRUE;
}

BOOL DlgCreateFroups::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if(pWndWithFocus==GetDlgItem(RNMF_BTN_OK))
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

void DlgCreateFroups::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}


void DlgCreateFroups::OnBnClickedBtnOk()
{
	try
	{
		int ncheck = 0;
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if ((int)m_list_view.GetCheck(i) == 1)
			{
				ncheck++;
				break;
			}
		}

		if (ncheck > 0)
		{
			vector<int> vecRow;
			int nSelect = _GetAllCheckedItems(vecRow);
			if (nSelect > 0)
			{
				_SaveRegistry();
				CDialogEx::OnOK();

				ff->_xlPutScreenUpdating(false);

				pShActive->Outline->PutAutomaticStyles(false);
				pShActive->Outline->PutSummaryRow(xlSummaryAbove);
				pShActive->Outline->PutSummaryColumn(xlSummaryOnLeft);

				int dem = 0;
				CString szval = L"";
				CString szUpdateStatus = L"Đang tiến hành nhóm dữ liệu. Vui lòng đợi trong giây lát...";
				for (int i = nSelect - 1; i >= 0; i--)
				{
					dem++;
					szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100*dem/nSelect), L"%");
					ff->_xlPutStatus(szval);

					_CreateGroupItems(vecRow[i]);		// <-- Tạo group dữ liệu
				}

				//pShActive->Outline->ShowLevels(1, 1);

				ff->_xlPutStatus(L"Ready");
				ff->_xlPutScreenUpdating(true);
			}
			vecRow.clear();
		}
		else ff->_msgbox(L"Bạn chưa tích chọn nhóm dữ liệu sử dụng.", 2, 0);
	}
	catch (...) {}
}

void DlgCreateFroups::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

/*****************************************************************/

void DlgCreateFroups::_GetDanhsachLinhvuc()
{
	try
	{
		ff->_xlPutScreenUpdating(false);

		_XacdinhSheetLienquan();

		RangePtr PRS;
		MyStrGroups MSFILE;
		MSFILE.szLinhvuc = L"";

		int icheck = 0;
		int total = 0, dem = 0;
		int ncount = irowEnd - irowStart;

		CString szval = L"", szlink = L"", szcheck = L"";
		CString szUpdateStatus = L"Đang tiến hành kiểm tra nhóm. Vui lòng đợi trong giây lát...";

		for (int i = irowStart; i < irowEnd; i++)
		{
			dem++;
			szval.Format(L"%s (%d%s) %s", szUpdateStatus,
				(int)(100 * dem / ncount), L"%", szlink);
			if (szval.GetLength() > 250) szval = szval.Left(250) + L"...";
			ff->_xlPutStatus(szval);

			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) > 0)
			{
				szval = ff->_xlGIOR(pRgActive, i, jcotLV, L"");
				if (szval != L"")
				{
					if (total == 0 || szval != MSFILE.szLinhvuc)
					{
						// Thêm mới
						MSFILE.szLinhvuc = szval;
						MSFILE.szPathFolder = L"";
						if (jcotFolder > 0)
						{
							MSFILE.szPathFolder = ff->_xlGIOR(pRgActive, i, jcotFolder, L"");
							if (MSFILE.szPathFolder == L"")
							{
								MSFILE.szPathFolder = ff->_xlGIOR(pRgActive, i + 1, jcotFolder, L"");
							}
						}

						MSFILE.iRBegin = i;
						MSFILE.iREnd = i;
						MSFILE.iEnable = 1;
						vecGroups.push_back(MSFILE);

						total++;
					}
					else
					{
						// Cộng dồn giá trị
						vecGroups[total - 1].iREnd = i;
					}
				}
			}
		}

		ff->_xlPutStatus(L"Ready");
		ff->_xlPutScreenUpdating(true);

		tongkq = (long)vecGroups.size();
		if (tongkq > 0)
		{
			// Hiển thị hộp thoại
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DoModal();
		}
		else
		{
			ff->_msgbox(L"Không tồn tại dữ liệu để tạo nhóm. Vui lòng kiểm tra lại.", 1, 0);
		}
	}
	catch (...) {}
}

void DlgCreateFroups::_SetDefault()
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

/**********************************************************************************/

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, COLW_FIX);
		m_list_view.InsertColumn(iGFLTten, L" Tên nhóm", LVCFMT_LEFT, iwColTen);
		m_list_view.InsertColumn(iGFLSlg, L" Số lượng", LVCFMT_CENTER, iwColSL);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_CHECKBOXES | LVS_EX_DOUBLEBUFFER);

		HWND header = ListView_GetHeader(m_list_view);
		DWORD dwHeaderStyle = ::GetWindowLong(header, GWL_STYLE);
		dwHeaderStyle |= HDS_CHECKBOXES;
		::SetWindowLong(header, GWL_STYLE, dwHeaderStyle);
		::GetDlgCtrlID(header);
		HDITEM hdi = { 0 };
		hdi.mask = HDI_FORMAT;
		Header_GetItem(header, 0, &hdi);
		hdi.fmt |= HDF_CHECKBOX | HDF_CHECKED/* | HDF_FIXEDWIDTH*/;
		Header_SetItem(header, 0, &hdi);
	}
	catch (...) {}
}

void DlgCreateFroups::OnHdnItemStateIconClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pnmHeader = (LPNMHEADER)pNMHDR;
	if (pnmHeader->pitem->mask & HDI_FORMAT && pnmHeader->pitem->fmt & HDF_CHECKBOX)
	{
		bChecked = !(pnmHeader->pitem->fmt & HDF_CHECKED);
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			m_list_view.SetCheck(i, bChecked);
			//m_list_view.SetItemState(i, -1, LVIS_SELECTED);
		}
	}

	*pResult = 0;
}

void DlgCreateFroups::_XacdinhSheetLienquan()
{
	CString szCopy = shCopy;
	jcotFolder = 0, jcotType = 0, jcotNoidung = 0;

	if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
	{
		ff->_GetThongtinSheetCategory(jcotSTT, jcotLV, icotMH, jcotTen,
			icotFilegoc, icotFileGXD, icotThLuan, icotEnd, irowTieude, irowStart, irowEnd);
	}
	else
	{
		ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
			jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
	}
}

void DlgCreateFroups::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CreateGroups;
		reg->CreateKey(szCreate);

		CString szval = L"";
		szval.Format(L"%d|%d",
			m_list_view.GetColumnWidth(iGFLTten),
			m_list_view.GetColumnWidth(iGFLSlg));

		reg->WriteChar(ff->_ConvertCstringToChars(DlgWidthColumn), ff->_ConvertCstringToChars(szval));

		CRect rec;
		GetWindowRect(&rec);
		szval.Format(L"%dx%d", (int)rec.Width(), (int)rec.Height());
		reg->WriteChar(ff->_ConvertCstringToChars(DlgSize), ff->_ConvertCstringToChars(szval));
	}
	catch (...) {}
}

void DlgCreateFroups::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + CreateGroups;
		reg->CreateKey(szCreate);

		CString szval = reg->ReadString(DlgWidthColumn, L"");
		if (szval != L"")
		{
			vector<CString> vecKey;
			ff->_TackTukhoa(vecKey, szval, L"|", L";", 0, 1);

			int ncount = (int)vecKey.size();
			if (ncount >= 2)
			{
				iwColTen = _wtoi(vecKey[0]);
				iwColSL = _wtoi(vecKey[1]);
			}

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
	}
	catch (...) {}
}

void DlgCreateFroups::_SetLocationAndSize()
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

void DlgCreateFroups::_LoadDanhsachNhom()
{
	try
	{
		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);

		lanshow = 1;
		long iz = ibuocnhay;
		if (tongkq <= iz)
		{
			iz = tongkq;
			iStopload = 0;
		}
		else iStopload = 1;

		// Thêm dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz, bChecked);
	}
	catch (...) {}

}

void DlgCreateFroups::_AddItemsListKetqua(int iBegin, int iEnd, int icheck)
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

		int pos = -1;
		CString szval = L"";
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, iGFLTten, vecGroups[i].szLinhvuc);
			
			szval.Format(L"%d", vecGroups[i].iREnd - vecGroups[i].iRBegin + 1);
			m_list_view.SetItemText(i, iGFLSlg, szval);

			m_list_view.SetCheck(i, icheck);

			if (iBegin > 0 && dem == ncount)
			{
				m_list_view.SetItemState(i, LVIS_SELECTED, LVIS_SELECTED);
			}
		}		
	}
	catch (...) {}
}

void DlgCreateFroups::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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
			long iz = ibuocnhay * lanshow;

			if (tongkq <= iz)
			{
				iz = tongkq;
				iStopload = 0;
			}
			else iStopload = 1;

			// Thêm dữ liệu vào list kết quả
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz, bChecked);

			m_list_view.EnsureVisible(ncount + (int)(b / a) - 5, FALSE);
		}
	}
	catch (...) {}
}


int DlgCreateFroups::_GetAllCheckedItems(vector<int> &vecRow)
{
	try
	{
		vecRow.clear();
		int ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			if (m_list_view.GetCheck(i)) vecRow.push_back(i);
		}

		if (bChecked == 1)
		{
			if (tongkq > ncount)
			{
				for (int i = ncount; i < tongkq; i++) vecRow.push_back(i);
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void DlgCreateFroups::_CreateGroupItems(int nItem)
{
	try
	{
		int Ibg = vecGroups[nItem].iRBegin;
		int Iend = vecGroups[nItem].iREnd;

		RangePtr PRS = pShActive->GetRange(pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Iend, 1));
		/*if ((int)PRS->EntireRow->OutlineLevel >= 2) */PRS->EntireRow->Rows->ClearOutline();

		// Chèn thêm 1 dòng trước nhóm
		CString szID = ff->_xlGIOR(pRgActive, Ibg - 1, jcotSTT, L"");
		if (_wtoi(szID) > 0 || Ibg == irowStart)
		{
			PRS = pRgActive->GetItem(Ibg + 1, 1);
			PRS->EntireRow->Insert(xlShiftDown);

			PRS = pRgActive->GetItem(Ibg, 1);
			PRS->EntireRow->Copy(pRgActive->GetItem(Ibg + 1, 1));
			xl->PutCutCopyMode(XlCutCopyMode(false));

			PRS = pShActive->GetRange(
				pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Ibg, icotEnd));
			PRS->EntireRow->ClearContents();

			_SetStyleGroup(PRS);
			
			pRgActive->PutItem(Ibg, jcotLV, (_bstr_t)vecGroups[nItem].szLinhvuc);			
			if (jcotFolder > 0) pRgActive->PutItem(Ibg, jcotFolder, (_bstr_t)vecGroups[nItem].szPathFolder);

			Ibg++; Iend++;
		}
		else
		{
			szID = ff->_xlGIOR(pRgActive, Ibg - 1, jcotLV, L"");
			if (szID == L"")
			{
				pRgActive->PutItem(Ibg - 1, jcotLV, (_bstr_t)vecGroups[nItem].szLinhvuc);
				if (jcotFolder > 0) pRgActive->PutItem(Ibg - 1, jcotFolder, (_bstr_t)vecGroups[nItem].szPathFolder);
			}

			PRS = pShActive->GetRange(
				pRgActive->GetItem(Ibg - 1, 1), pRgActive->GetItem(Ibg - 1, icotEnd));
			_SetStyleGroup(PRS);
		}

		PRS = pShActive->GetRange(pRgActive->GetItem(Ibg, 1), pRgActive->GetItem(Iend, 1));
		PRS->EntireRow->Rows->Group();
	}
	catch (...) {}
}

void DlgCreateFroups::_SetStyleGroup(RangePtr PRS)
{
	PRS->Font->PutBold(true);
	PRS->Font->PutItalic(false);
	PRS->Font->PutColor(rgbBlack);
	PRS->Interior->PutColor(RGB(153, 255, 204));
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeTop)->PutWeight(xlThin);
	PRS->Borders->GetItem(XlBordersIndex::xlEdgeBottom)->PutWeight(xlThin);
}
