#include "pch.h"
#include "DlgEditProperties.h"
#include "DlgSelectionCells.h"

IMPLEMENT_DYNAMIC(DlgEditProperties, CDialogEx)

#define EDTPRO_LINKERROR	L"Link lỗi"
#define EDTPRO_EDIT			L"Đang sửa"
#define EDTPRO_SUCCESSFUL	L"Thành công"

#define BAMUL_TITLE			L"Nhập mô tả / nội dung file"
#define BAMUL_SUBJECT		L"Đối tượng sử dụng"
#define BAMUL_TAGS			L"Nhập các thẻ Tags (Keywords)"
#define BAMUL_CATEGORIES	L"Nhập các lĩnh vực liên quan"
#define BAMUL_COMMENTS		L"Nhập ghi chú khác"

#define RGBGREEN017680		RGB(0, 176, 80)
#define RGBGRAY929292		RGB(92, 92, 92)

DlgEditProperties::DlgEditProperties(CWnd* pParent /*=NULL*/)
	: CDialogEx(DlgEditProperties::IDD, pParent)
{
	ff = new Function;
	bb = new Base64Ex;
	reg = new CRegistry;

	icotSTT = 1, icotLV = icotSTT + 1, icotMH = icotSTT + 2, icotNoidung = icotSTT + 3;
	icotFilegoc = 11, icotFileGXD = icotFilegoc + 1, icotThLuan = icotFilegoc + 2;
	irowTieude = 3, irowStart = 4, irowEnd = 5;
	icotFolder = 2, icotTen = 3, icotType = 4;

	iStopload = 1, lanshow = 1, ibuocnhay = 50;
	tongkq = 0, tongfilter = 0;

	iGFLType = 1, iGFLTen = 2, iGFLTitle = 3, iGFLSubject = 4, iGFLTags = 5, iGFLCategories = 6;
	iGFLComments = 7, iGFLAuthors = 8, iGFLCompany = 9, iGFLPath = 10;

	nItem = -1, nSubItem = -1;
	nCheckEdit = 0;
	iKeyESC = 0;

	jColumnHand = iGFLPath;
	pathFolderDir = L"";
	szBefore = L"", szAfter = L"", szCopy = L"<null>";
	blIcon = false;

	nTotalCol = getSizeArray(iwCol);
	iwCol[0] = 22, iwCol[1] = 100, iwCol[2] = 200, iwCol[3] = 400;
	iwCol[4] = 120, iwCol[5] = 120, iwCol[6] = 120, iwCol[7] = 120;
	iwCol[8] = 120, iwCol[9] = 120, iwCol[10] = 120, iwCol[11] = 200;
	iDlgW = 0, iDlgH = 0;
}

DlgEditProperties::~DlgEditProperties()
{
	delete ff;
	delete bb;
	delete reg;

	vecItem.clear();
	vecFilter.clear();

	jColumnHand = -1;
}

void DlgEditProperties::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, EDPR_TXT_SEARCH, m_txt_search);
	DDX_Control(pDX, EDPR_BTN_SEARCH, m_btn_search);
	DDX_Control(pDX, EDPR_LIST_VIEW, m_list_view);
	DDX_Control(pDX, EDPR_BTN_UPDATE, m_btn_update);
	DDX_Control(pDX, EDPR_BTN_CANCEL, m_btn_cancel);
	DDX_Control(pDX, EDPR_EDIT_VIEW, m_edit_properties);
	DDX_Control(pDX, EDPR_GROUP_INFO, m_lbl_info);
	DDX_Control(pDX, EDPR_GROUP_LIST, m_lbl_kq);
	DDX_Control(pDX, EDPR_LBL_PROGRESS, m_lbl_progress);

	DDX_Control(pDX, EDPR_TXT_FILENAME, m_txt_filename);
	DDX_Control(pDX, EDPR_TXT_TITLE, m_txt_title);
	DDX_Control(pDX, EDPR_TXT_SUBJECT, m_txt_subject);
	DDX_Control(pDX, EDPR_TXT_TAGS, m_txt_tags);
	DDX_Control(pDX, EDPR_TXT_CATEGORIES, m_txt_categories);
	DDX_Control(pDX, EDPR_TXT_COMMENTS, m_txt_commments);
	DDX_Control(pDX, EDPR_TXT_AUTHORS, m_txt_authors);
	DDX_Control(pDX, EDPR_TXT_COMPANY, m_txt_company);

	DDX_Control(pDX, EDPR_BTN_UP, m_btn_up);
	DDX_Control(pDX, EDPR_BTN_DOWN, m_btn_down);
	DDX_Control(pDX, EDPR_BTN_CLEAR, m_btn_clear);
	DDX_Control(pDX, EDPR_BTN_SAVED, m_btn_saved);
}

BEGIN_MESSAGE_MAP(DlgEditProperties, CDialogEx)
	ON_WM_SIZE()
	ON_WM_SYSCOMMAND()
	ON_WM_CONTEXTMENU()
	ON_WM_GETMINMAXINFO()
	ON_NOTIFY(SPN_MAXMINPOS, EDPR_SPLITER_VER, OnMaxMinInfo)
	ON_NOTIFY(SPN_DELTA, EDPR_SPLITER_VER, OnSplitterDelta)
	ON_BN_CLICKED(EDPR_BTN_SEARCH, &DlgEditProperties::OnBnClickedBtnSearch)
	ON_BN_CLICKED(EDPR_BTN_CANCEL, &DlgEditProperties::OnBnClickedBtnCancel)
	ON_BN_CLICKED(EDPR_BTN_UPDATE, &DlgEditProperties::OnBnClickedBtnUpdate)
	ON_BN_CLICKED(EDPR_STT_2, &DlgEditProperties::OnBnClickedStt2)
	ON_BN_CLICKED(EDPR_STT_3, &DlgEditProperties::OnBnClickedStt3)
	ON_BN_CLICKED(EDPR_STT_4, &DlgEditProperties::OnBnClickedStt4)
	ON_BN_CLICKED(EDPR_STT_5, &DlgEditProperties::OnBnClickedStt5)
	ON_BN_CLICKED(EDPR_STT_6, &DlgEditProperties::OnBnClickedStt6)
	ON_BN_CLICKED(EDPR_STT_7, &DlgEditProperties::OnBnClickedStt7)
	ON_BN_CLICKED(EDPR_STT_8, &DlgEditProperties::OnBnClickedStt8)
	ON_BN_CLICKED(EDPR_BTN_UP, &DlgEditProperties::OnBnClickedBtnUp)
	ON_BN_CLICKED(EDPR_BTN_DOWN, &DlgEditProperties::OnBnClickedBtnDown)
	ON_BN_CLICKED(EDPR_BTN_CLEAR, &DlgEditProperties::OnBnClickedBtnClear)
	ON_BN_CLICKED(EDPR_BTN_SAVED, &DlgEditProperties::OnBnClickedBtnSaved)
	ON_EN_SETFOCUS(EDPR_TXT_TITLE, &DlgEditProperties::OnEnSetfocusTxtTitle)
	ON_EN_KILLFOCUS(EDPR_TXT_TITLE, &DlgEditProperties::OnEnKillfocusTxtTitle)
	ON_EN_SETFOCUS(EDPR_TXT_SUBJECT, &DlgEditProperties::OnEnSetfocusTxtSubject)
	ON_EN_KILLFOCUS(EDPR_TXT_SUBJECT, &DlgEditProperties::OnEnKillfocusTxtSubject)
	ON_EN_SETFOCUS(EDPR_TXT_TAGS, &DlgEditProperties::OnEnSetfocusTxtTags)
	ON_EN_KILLFOCUS(EDPR_TXT_TAGS, &DlgEditProperties::OnEnKillfocusTxtTags)
	ON_EN_SETFOCUS(EDPR_TXT_CATEGORIES, &DlgEditProperties::OnEnSetfocusTxtCategories)
	ON_EN_KILLFOCUS(EDPR_TXT_CATEGORIES, &DlgEditProperties::OnEnKillfocusTxtCategories)
	ON_EN_SETFOCUS(EDPR_TXT_COMMENTS, &DlgEditProperties::OnEnSetfocusTxtComments)
	ON_EN_KILLFOCUS(EDPR_TXT_COMMENTS, &DlgEditProperties::OnEnKillfocusTxtComments)
	ON_EN_SETFOCUS(EDPR_TXT_SEARCH, &DlgEditProperties::OnEnSetfocusTxtSearch)
	ON_EN_CHANGE(EDPR_TXT_TITLE, &DlgEditProperties::OnEnChangeTxtTitle)
	ON_EN_CHANGE(EDPR_TXT_SUBJECT, &DlgEditProperties::OnEnChangeTxtSubject)
	ON_EN_CHANGE(EDPR_TXT_TAGS, &DlgEditProperties::OnEnChangeTxtTags)
	ON_EN_CHANGE(EDPR_TXT_CATEGORIES, &DlgEditProperties::OnEnChangeTxtCategories)
	ON_EN_CHANGE(EDPR_TXT_COMMENTS, &DlgEditProperties::OnEnChangeTxtComments)
	ON_EN_CHANGE(EDPR_TXT_AUTHORS, &DlgEditProperties::OnEnChangeTxtAuthors)
	ON_EN_CHANGE(EDPR_TXT_COMPANY, &DlgEditProperties::OnEnChangeTxtCompany)
	ON_NOTIFY(NM_CLICK, EDPR_LIST_VIEW, &DlgEditProperties::OnNMClickListView)
	ON_NOTIFY(NM_RCLICK, EDPR_LIST_VIEW, &DlgEditProperties::OnNMRClickListView)
	ON_NOTIFY(LVN_KEYDOWN, EDPR_LIST_VIEW, &DlgEditProperties::OnLvnKeydownListView)
	ON_EN_SETFOCUS(EDPR_EDIT_VIEW, &DlgEditProperties::OnEnSetfocusEditView)
	ON_EN_KILLFOCUS(EDPR_EDIT_VIEW, &DlgEditProperties::OnEnKillfocusEditView)
	ON_NOTIFY(LVN_ENDSCROLL, EDPR_LIST_VIEW, &DlgEditProperties::OnLvnEndScrollListView)
	ON_NOTIFY(HDN_ENDTRACK, 0, &DlgEditProperties::OnHdnEndtrackListView)
	ON_COMMAND(MN_PROPERTIES_OPENFILE, &DlgEditProperties::OnPropertiesOpenfile)
	ON_COMMAND(MN_PROPERTIES_OPENFOLDER, &DlgEditProperties::OnPropertiesOpenfolder)
	ON_COMMAND(MN_PROPERTIES_COPY, &DlgEditProperties::OnPropertiesCopy)
	ON_COMMAND(MN_PROPERTIES_PASTE, &DlgEditProperties::OnPropertiesPaste)
END_MESSAGE_MAP()

// DlgEditProperties message handlers
BOOL DlgEditProperties::OnInitDialog()
{
	CDialogEx::OnInitDialog();
	HICON hIcon = LoadIcon(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICON_PROPERTIES));
	SetIcon(hIcon, FALSE);

	pathFolderDir = ff->_GetPathFolder(strNAMEDLL);
	if (pathFolderDir.Right(1) != L"\\") pathFolderDir += L"\\";

	mnIcon.GdiplusStartupInputConfig();

	int nsize = getSizeArray(blRead);
	for (int i = 0; i < nsize; i++) blRead[i] = false;

	_LoadRegistry();
	_SetDefault();
	_Timkiemdulieu();

	_SetLocationAndSize();
	_SetSpliterPane();

	if (nCheckEdit > 0)
	{
		CString szval = L"";
		szval.Format(L"Tồn tại (%d) nội dung file", nCheckEdit);
		if (nCheckEdit > 1) szval += L"s";
		szval += L" được chỉnh sửa trực tiếp trên Excel. Bạn có muốn cập nhật dữ liệu ngay?";
	
		int result = ff->_y(szval, 0, 0, 0);
		if (result == 6)
		{
			OnBnClickedBtnUpdate();
		}	
	}

	return TRUE;
}

BOOL DlgEditProperties::PreTranslateMessage(MSG* pMsg)
{
	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (iWPar == VK_RETURN)
		{
			if (pWndWithFocus == GetDlgItem(EDPR_BTN_UPDATE))
			{
				OnBnClickedBtnUpdate();
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_SEARCH))
			{
				OnBnClickedBtnSearch();
				return TRUE;
			}
		}
		else if (iWPar == VK_TAB)
		{
			if (pWndWithFocus == GetDlgItem(EDPR_TXT_SEARCH))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_LIST_VIEW));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_LIST_VIEW))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_SEARCH));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_TITLE))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_SUBJECT));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_SUBJECT))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_TAGS));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_TAGS))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_CATEGORIES));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_CATEGORIES))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_COMMENTS));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_COMMENTS))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_AUTHORS));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_AUTHORS))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_COMPANY));
				return TRUE;
			}
			else if (pWndWithFocus == GetDlgItem(EDPR_TXT_COMPANY))
			{
				GotoDlgCtrl(GetDlgItem(EDPR_TXT_TITLE));
				return TRUE;
			}
		}
		else if (iWPar == VK_ESCAPE)
		{
			if (iKeyESC == 0) OnBnClickedBtnCancel();
			else
			{
				szBefore = m_list_view.GetItemText(CLRow, CLColumn);
				m_edit_properties.SetWindowTextW(szBefore);
				GotoDlgCtrl(GetDlgItem(EDPR_LIST_VIEW));
			}
			iKeyESC = 0;

			return TRUE;
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(EDPR_TXT_FILENAME))
		{
			if (chr == 0x01)
			{
				m_txt_filename.SetSel(0, -1);
				return TRUE;
			}
		}
	}

	return FALSE;
}

void DlgEditProperties::OnSize(UINT nType, int cx, int cy)
{
	this->_ResizeDialog();
}

void DlgEditProperties::OnSysCommand(UINT nID, LPARAM lParam)
{
	if (nID == SC_CLOSE) OnBnClickedBtnCancel();
	else CDialog::OnSysCommand(nID, lParam);
}

void DlgEditProperties::OnContextMenu(CWnd* /*pWnd*/, CPoint point)
{
	try
	{
		if (nItem == -1 || nSubItem == -1) return;

		HMENU hMMenu = LoadMenu(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MENU_PROPERTIES));
		HMENU pMenu = GetSubMenu(hMMenu, 0);
		if (pMenu)
		{
			CString szval = L"";
			int ilen = 0, ivtri = 30;

			if (!(nSubItem == 0 || nSubItem == iGFLType
				|| nSubItem == iGFLTen || nSubItem == iGFLPath))
			{
				if (nItem >= 0 && nSubItem >= 0)
				{
					CString str = m_list_view.GetItemText(nItem, nSubItem);
					szval.Format(L"Sao chép '%s", str);

					ivtri = 30;
					ilen = szval.GetLength();
					if (ilen > ivtri)
					{
						for (int i = ivtri; i < ilen; i++)
						{
							if (szval.Mid(i, 1) == L" " || szval.Mid(i, 1) == L","
								|| szval.Mid(i, 1) == L";" || szval.Mid(i, 1) == L".")
							{
								ivtri = i;
								break;
							}
						}
						szval = szval.Left(ivtri) + L"...";
					}

					szval += L"'";
					ModifyMenu(pMenu, MN_PROPERTIES_COPY, MF_BYCOMMAND | MF_STRING, MN_PROPERTIES_COPY, szval);
				}

				if (szCopy != L"<null>")
				{
					szval.Format(L"Dán '%s", szCopy);

					ivtri = 30;
					ilen = szval.GetLength();
					if (ilen > ivtri)
					{
						for (int i = ivtri; i < ilen; i++)
						{
							if (szval.Mid(i, 1) == L" " || szval.Mid(i, 1) == L","
								|| szval.Mid(i, 1) == L";" || szval.Mid(i, 1) == L".")
							{
								ivtri = i;
								break;
							}
						}
						szval = szval.Left(ivtri) + L"...";
					}

					szval += L"'";
					ModifyMenu(pMenu, MN_PROPERTIES_PASTE, MF_BYCOMMAND | MF_STRING, MN_PROPERTIES_PASTE, szval);
				}
			}

			HBITMAP hBmp;

			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_PROPERTIES_OPENFILE, IDI_ICON_OPENFILE);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_PROPERTIES_OPENFOLDER, IDI_ICON_OPENFOLDER);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_PROPERTIES_COPY, IDI_ICON_COPYCLIP);
			mnIcon.AddIconMenuItem(pMenu, hBmp, MN_PROPERTIES_PASTE, IDI_ICON_COPY);

			vector<int> vecRow;
			int nSelect = _GetAllSelectedItems(vecRow);
			if (nSelect >= 2)
			{
				EnableMenuItem(pMenu, MN_PROPERTIES_OPENFILE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_PROPERTIES_OPENFOLDER, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_PROPERTIES_COPY, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}
			vecRow.clear();

			if (szCopy == L"<null>")
			{
				EnableMenuItem(pMenu, MN_PROPERTIES_PASTE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			if (nSubItem == 0 || nSubItem == iGFLType
				|| nSubItem == iGFLTen || nSubItem == iGFLPath)
			{
				EnableMenuItem(pMenu, MN_PROPERTIES_COPY, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
				EnableMenuItem(pMenu, MN_PROPERTIES_PASTE, MF_BYCOMMAND | MF_DISABLED | MF_GRAYED);
			}

			CRect rectSubmitButton;
			CListCtrlEx *pListview = reinterpret_cast<CListCtrlEx *>(GetDlgItem(EDPR_LIST_VIEW));
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

void DlgEditProperties::OnBnClickedBtnSearch()
{
	_Timkiemdulieu();
}


void DlgEditProperties::OnBnClickedBtnCancel()
{
	_SaveRegistry();
	CDialogEx::OnCancel();
}

void DlgEditProperties::_SaveRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + EditProperties;
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
	}
	catch (...) {}
}

void DlgEditProperties::_LoadRegistry()
{
	try
	{
		CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + EditProperties;
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
	}
	catch (...) {}
}

void DlgEditProperties::_SetLocationAndSize()
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

void DlgEditProperties::_SetDefault()
{
	try
	{
		m_edit_properties.SubclassDlgItem(EDPR_EDIT_VIEW, this);
		m_edit_properties.SetBkColor(rgbWhite);
		m_edit_properties.SetTextColor(rgbBlue);

		m_txt_filename.SubclassDlgItem(EDPR_TXT_FILENAME, this);
		m_txt_filename.SetBkColor(rgbWhite);
		m_txt_filename.SetTextColor(RGB(192, 0, 0));
		m_txt_filename.SetFont(&m_FontHeaderList);

		m_txt_title.SubclassDlgItem(EDPR_TXT_TITLE, this);
		m_txt_title.SetBkColor(rgbWhite);

		m_txt_subject.SubclassDlgItem(EDPR_TXT_SUBJECT, this);
		m_txt_subject.SetBkColor(rgbWhite);

		m_txt_tags.SubclassDlgItem(EDPR_TXT_TAGS, this);
		m_txt_tags.SetBkColor(rgbWhite);

		m_txt_categories.SubclassDlgItem(EDPR_TXT_CATEGORIES, this);
		m_txt_categories.SetBkColor(rgbWhite);

		m_txt_commments.SubclassDlgItem(EDPR_TXT_COMMENTS, this);
		m_txt_commments.SetBkColor(rgbWhite);

		m_txt_authors.SubclassDlgItem(EDPR_TXT_AUTHORS, this);
		m_txt_authors.SetBkColor(rgbWhite);
		m_txt_authors.SetTextColor(rgbBlue);
		m_txt_authors.SetCueBanner(L"Tác giả");

		m_txt_company.SubclassDlgItem(EDPR_TXT_COMPANY, this);
		m_txt_company.SetBkColor(rgbWhite);
		m_txt_company.SetTextColor(rgbBlue);
		m_txt_company.SetCueBanner(L"Thông tin công ty");

		_SetAllBannerMulti(false);

		m_list_view.InsertColumn(0, L"", LVCFMT_CENTER, 22);
		m_list_view.InsertColumn(iGFLType, L" Trạng thái", LVCFMT_CENTER, iwCol[1]);
		m_list_view.InsertColumn(iGFLTen, L" Tên file", LVCFMT_LEFT, iwCol[2]);
		m_list_view.InsertColumn(iGFLTitle, L" Title (Nội dung)", LVCFMT_LEFT, iwCol[3]);
		m_list_view.InsertColumn(iGFLSubject, L" Subject", LVCFMT_LEFT, iwCol[4]);
		m_list_view.InsertColumn(iGFLTags, L" Tags", LVCFMT_LEFT, iwCol[5]);
		m_list_view.InsertColumn(iGFLCategories, L" Categories", LVCFMT_LEFT, iwCol[6]);
		m_list_view.InsertColumn(iGFLComments, L" Comments", LVCFMT_LEFT, iwCol[7]);
		m_list_view.InsertColumn(iGFLAuthors, L" Authors", LVCFMT_LEFT, iwCol[8]);
		m_list_view.InsertColumn(iGFLCompany, L" Company", LVCFMT_LEFT, iwCol[9]);
		m_list_view.InsertColumn(iGFLPath, L" Đường dẫn", LVCFMT_LEFT, iwCol[10]);

		m_list_view.SetColumnEditor(iGFLTitle, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLSubject, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLTags, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLCategories, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLComments, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLAuthors, &m_edit_properties);
		m_list_view.SetColumnEditor(iGFLCompany, &m_edit_properties);

		m_list_view.SetColumnColors(iGFLPath, rgbWhite, rgbBlue);

		m_list_view.GetHeaderCtrl()->SetFont(&m_FontHeaderList);
		m_list_view.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES |
			LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES);

		/**********************************************************************************/

		m_txt_search.SubclassDlgItem(EDPR_TXT_SEARCH, this);
		m_txt_search.SetBkColor(rgbWhite);
		m_txt_search.SetTextColor(rgbBlue);
		m_txt_search.SetCueBanner(L"Nhập từ khóa tìm kiếm files...");

		m_lbl_progress.SubclassDlgItem(EDPR_LBL_PROGRESS, this);
		m_lbl_progress.SetTextColor(rgbBlue);
		m_lbl_progress.SetWindowText(L"Bạn có thể thay đổi nội dung files ngay tại sheet quản lý.");

		m_btn_search.SetIcon(IDI_ICON_SEARCH, 24, 24);
		m_btn_search.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_search.DrawTransparent(true);
		m_btn_search.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_update.SetThemeHelper(&m_ThemeHelper);
		m_btn_update.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_update.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_update.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_cancel.SetThemeHelper(&m_ThemeHelper);
		m_btn_cancel.SetIcon(IDI_ICON_CANCEL, 24, 24);
		m_btn_cancel.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_cancel.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_up.SetThemeHelper(&m_ThemeHelper);
		m_btn_up.SetIcon(IDI_ICON_UP, 24, 24);
		m_btn_up.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_up.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_down.SetThemeHelper(&m_ThemeHelper);
		m_btn_down.SetIcon(IDI_ICON_DOWN, 24, 24);
		m_btn_down.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_down.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_clear.SetThemeHelper(&m_ThemeHelper);
		m_btn_clear.SetIcon(IDI_ICON_BROOM, 24, 24);
		m_btn_clear.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_clear.SetBtnCursor(IDC_CURSOR_HAND);

		m_btn_saved.SetThemeHelper(&m_ThemeHelper);
		m_btn_saved.SetIcon(IDI_ICON_SAVED, 24, 24);
		m_btn_saved.OffsetColor(CButtonST::BTNST_COLOR_BK_IN, 30);
		m_btn_saved.SetBtnCursor(IDC_CURSOR_HAND);
	}
	catch (...) {}
}

void DlgEditProperties::_XacdinhSheetLienquan()
{
	ff->_GetThongtinSheetFManager(jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType,
		jcotNoidung, icotFilegoc, icotEnd, irowTieude, irowStart, irowEnd);
}

void DlgEditProperties::_Timkiemdulieu()
{
	try
	{
		iKeyESC = 0;
		tongfilter = _FilterData();

		m_list_view.DeleteAllItems();
		ASSERT(m_list_view.GetItemCount() == 0);

		// Delete list image icon
		m_imageList.DeleteImageList();
		ASSERT(m_imageList.GetSafeHandle() == NULL);
		m_imageList.Create(16, 16, ILC_COLOR32, 0, 0);

		lanshow = 1;
		long iz = ibuocnhay;
		if (tongfilter <= iz)
		{
			iz = tongfilter;
			iStopload = 0;
		}
		else iStopload = 1;

		// Thêm dữ liệu vào list kết quả
		_AddItemsListKetqua(0, iz);

		_SetStatusKetqua(tongfilter);

		m_txt_search.SetSel(0, -1);
	}
	catch (...) {}
}

/*****************************************************************/

void DlgEditProperties::_GetDanhsachFiles(int iSelectCell)
{
	try
	{
		if (iSelectCell >= 0)
		{
			ff->_xlPutScreenUpdating(false);

			_XacdinhSheetLienquan();

			if (iSelectCell == 0)
			{
				int ncount = (int)vecSelect.size();
				irowStart = vecSelect[0].row;
				irowEnd = vecSelect[ncount - 1].row + 1;
			}
			else if (iSelectCell == 1)
			{
				_GetGroupRowStartEnd(irowStart, irowEnd);
			}

			vecItem.clear();

			RangePtr PRS;
			MyStrProperties MSPRO;

			int pos = 0, dem = 0;
			int ncount = irowEnd - irowStart;

			vector<CString> vecKey;
			CString szval = L"", szlink = L"";
			CString szUpdateStatus = L"Đang tiến hành đọc thông tin files dữ liệu. Vui lòng đợi trong giây lát...";

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
					PRS = pRgActive->GetItem(i, icotFilegoc);
					szlink = ff->_xlGetHyperLink(PRS);
					if (szlink != L"")
					{
						MSPRO.szType = ff->_GetTypeOfFile(szlink);
						szval = MSPRO.szType.Left(3);
						szval.MakeLower();
						if (szval == L"pdf" || szval == L"doc" || szval == L"dot"
							|| szval == L"xls" || szval == L"xlt" || szval == L"ppt" || szval == L"pot")
						{
							MSPRO.nRow = i;
							MSPRO.iID = (int)vecItem.size();
							MSPRO.szTen = ff->_GetNameOfPath(szlink, pos, -1);

							MSPRO.szPath = szlink;
							MSPRO.szNoidungEx = ff->_xlGIOR(pRgActive, i, jcotNoidung, L"");

							// Lấy tất cả thông tin file
							szval = ff->_GetProperties(szlink, 0);

							MSPRO.szTitle = L"";
							MSPRO.szSubject = L"";
							MSPRO.szTags = L"";
							MSPRO.szCategories = L"";
							MSPRO.szComments = L"";
							MSPRO.szAuthors = L"";
							MSPRO.szCompany = L"";

							if (szval != L"")
							{
								pos = ff->_TackTukhoa(vecKey, szval, L"|", L"|", 0, 1);
								if (pos < 7) for (int j = pos; j < 7; j++) vecKey.push_back(L"");

								MSPRO.szTitle = vecKey[0];
								MSPRO.szSubject = vecKey[1];
								MSPRO.szTags = vecKey[2];
								MSPRO.szCategories = vecKey[3];
								MSPRO.szComments = vecKey[4];
								MSPRO.szAuthors = vecKey[5];
								MSPRO.szCompany = vecKey[6];
							}

							if (MSPRO.szTitle != MSPRO.szNoidungEx) nCheckEdit++;

							vecItem.push_back(MSPRO);
						}
					}
				}
			}

			ff->_xlPutStatus(L"Ready");
			ff->_xlPutScreenUpdating(true);

			tongkq = (int)vecItem.size();
			if (tongkq > 0)
			{
				// Hiển thị hộp thoại
				xl->EnableCancelKey = XlEnableCancelKey(FALSE);
				AFX_MANAGE_STATE(AfxGetStaticModuleState());
				DoModal();
			}
			else
			{
				ff->_msgbox(L"Không tồn tại files dữ liệu có định dạng Document (doc, xls, pdf,...). Vui lòng kiểm tra lại.", 1, 0);
			}
		}
	}
	catch (...) {}
}

void DlgEditProperties::OnBnClickedBtnUpdate()
{
	try
	{
		iKeyESC = 0;

		CString szUpdateStatus = L"Đang cập nhật nội dung. Vui lòng đợi trong giây lát...";
		m_lbl_progress.SetWindowText(szUpdateStatus);

		if (nCheckEdit == 0)
		{
			// Nếu không có sự thay đổi từ Excel --> Câp nhật tiếp từ nút Saved 
			OnBnClickedBtnSaved();
			return;
		}
		nCheckEdit = 0;

		int ncount = (int)vecItem.size();
		for (int i = 0; i < ncount; i++)
		{
			if (vecItem[i].szTitle != vecItem[i].szNoidungEx)
			{
				if (ff->_SetProperties(vecItem[i].szPath, 1, vecItem[i].szNoidungEx))
				{
					vecItem[i].szTitle == vecItem[i].szNoidungEx;
				}
			}
		}

		CString szval = L"";
		ncount = m_list_view.GetItemCount();
		for (int i = 0; i < ncount; i++)
		{
			szval.Format(L"%s (%d%s)", szUpdateStatus, (int)(100 * (i + 1) / ncount), L"%");
			m_lbl_progress.SetWindowText(szval);
			ff->_xlPutStatus(szval);

			szval = m_list_view.GetItemText(i, iGFLType);
			if (szval == EDTPRO_EDIT)
			{
				m_list_view.SetItemText(i, iGFLType, EDTPRO_SUCCESSFUL);
				m_list_view.SetCellColors(i, iGFLType, RGBGREEN017680, rgbWhite);
				m_list_view.SetItemText(i, iGFLTitle, vecItem[vecFilter[i].iID].szNoidungEx);
			}
		}

		m_lbl_progress.SetWindowText(L"");
		ff->_msgbox(L"Cập nhật các tiêu đề thành công.", 0, 0, 5000);
	}
	catch (...) {}
}

int DlgEditProperties::_FilterData()
{
	try
	{
		vecFilter.clear();

		CString szval = L"";
		m_txt_search.GetWindowText(szval);
		szval.Trim();

		if (szval == L"") vecFilter = vecItem;
		else
		{
			szval = ff->_ConvertKhongDau(szval);
			szval.MakeLower();

			vector<CString> vecKey;
			int iKey = ff->_TackTukhoa(vecKey, szval, L" ", L"+", 1, 1);

			int icheck = 0;
			int itotal = (int)vecItem.size();
			for (int i = 0; i < itotal; i++)
			{
				icheck = 0;
				szval = vecItem[i].szTen + L"|" + vecItem[i].szType + L"|"
					+ vecItem[i].szTitle + L"|" + vecItem[i].szSubject + L"|"
					+ vecItem[i].szCategories;
				szval = ff->_ConvertKhongDau(szval);
				szval.MakeLower();

				for (int j = 0; j < iKey; j++)
				{
					if (szval.Find(vecKey[j]) >= 0) icheck++;
					else break;
				}

				if (icheck == iKey)
				{
					vecFilter.push_back(vecItem[i]);
				}
			}

			vecKey.clear();
		}

		return (int)vecFilter.size();
	}
	catch (...) {}
	return 0;
}

void DlgEditProperties::_SetStatusKetqua(int num)
{
	/*CString str = L"";
	if (num > 1) str.Format(L"Bạn đã chọn (%d) files / ", num);
	else str.Format(L"Bạn đã chọn (%d) file / ", num);
	m_lbl_total.SetWindowText(str);*/
}

void DlgEditProperties::_AddItemsListKetqua(int iBegin, int iEnd)
{
	try
	{
		CString szKodau[2] = { L"",L"" };
		for (int i = iBegin; i < iEnd; i++)
		{
			m_list_view.InsertItem(i, L"", 0);
			m_list_view.SetItemText(i, iGFLTen, vecFilter[i].szTen);
			m_list_view.SetItemText(i, iGFLTitle, vecFilter[i].szTitle);
			m_list_view.SetItemText(i, iGFLSubject, vecFilter[i].szSubject);
			m_list_view.SetItemText(i, iGFLTags, vecFilter[i].szTags);
			m_list_view.SetItemText(i, iGFLCategories, vecFilter[i].szCategories);
			m_list_view.SetItemText(i, iGFLComments, vecFilter[i].szComments);
			m_list_view.SetItemText(i, iGFLAuthors, vecFilter[i].szAuthors);
			m_list_view.SetItemText(i, iGFLCompany, vecFilter[i].szCompany);
			m_list_view.SetItemText(i, iGFLPath, vecFilter[i].szPath);

			HICON hi = m_sysImageList.GetFileIcon(L"*." + vecFilter[i].szType);
			m_imageList.Add(hi);

			if (ff->_FileExistsUnicode(vecFilter[i].szPath, 0) != 1)
			{
				m_list_view.SetItemText(i, iGFLType, EDTPRO_LINKERROR);
				m_list_view.SetCellColors(i, iGFLType, rgbRed, rgbWhite);
			}
			else
			{
				szKodau[0] = ff->_ConvertKhongDau(vecFilter[i].szTitle);
				szKodau[1] = ff->_ConvertKhongDau(vecFilter[i].szNoidungEx);

				if (szKodau[0] != szKodau[1])
				{
					m_list_view.SetItemText(i, iGFLType, EDTPRO_EDIT);
					m_list_view.SetCellColors(i, iGFLType, RGB(0, 112, 192), rgbWhite);
				}
			}
		}

		// Load Image vào list dữ liệu
		m_list_view.SetImageList(&m_imageList);

		for (int i = iBegin; i < iEnd; i++)
		{
			// Load Image icon
			m_list_view.SetItem(i, 0, LVIF_IMAGE, NULL, i, 0, 0, 0);
		}
	}
	catch (...) {}
}

void DlgEditProperties::OnEnSetfocusTxtSearch()
{
	iKeyESC = 0;
}


void DlgEditProperties::OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	iKeyESC = 0;
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	nItem = pNMListView->iItem;
	nSubItem = pNMListView->iSubItem;

	if (nItem >= 0 && nSubItem >= 0)
	{
		vector<int> vecRow;
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 1) nItem = vecRow[0];		// Xác định lại vị trí lựa chọn
		vecRow.clear();

		_SetIconButtonSaved(0);
		_LoadProperties(nItem);
	}
}

void DlgEditProperties::_LoadProperties(int nIndex)
{
	try
	{
		CString strTen = m_list_view.GetItemText(nIndex, iGFLTen);
		CString strTitle = m_list_view.GetItemText(nIndex, iGFLTitle);
		CString strSubject = m_list_view.GetItemText(nIndex, iGFLSubject);
		CString strTags = m_list_view.GetItemText(nIndex, iGFLTags);
		CString strCategories = m_list_view.GetItemText(nIndex, iGFLCategories);
		CString strComments = m_list_view.GetItemText(nIndex, iGFLComments);
		CString strAuthors = m_list_view.GetItemText(nIndex, iGFLAuthors);
		CString strCompany = m_list_view.GetItemText(nIndex, iGFLCompany);

		m_txt_filename.SetWindowText(strTen);
		if (!blRead[0]) m_txt_title.SetWindowText(strTitle);
		if (!blRead[1]) m_txt_subject.SetWindowText(strSubject);
		if (!blRead[2]) m_txt_tags.SetWindowText(strTags);
		if (!blRead[3]) m_txt_categories.SetWindowText(strCategories);
		if (!blRead[4]) m_txt_commments.SetWindowText(strComments);
		if (!blRead[5]) m_txt_authors.SetWindowText(strAuthors);
		if (!blRead[6]) m_txt_company.SetWindowText(strCompany);

		_SetAllBannerMulti(false);
	}
	catch (...) {}
}

void DlgEditProperties::OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	OnNMClickListView(pNMHDR, pResult);
}

void DlgEditProperties::OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVKEYDOWN pLVKeyDow = reinterpret_cast<LPNMLVKEYDOWN>(pNMHDR);
	if (pLVKeyDow->wVKey == VK_TAB || pLVKeyDow->wVKey == VK_F2)
	{
		if (nItem != -1 && nSubItem != -1)
		{
			CLRow = nItem, CLColumn = nSubItem;
			m_list_view.DisplayEditor(CLRow, CLColumn);
		}
	}
}


void DlgEditProperties::OnEnSetfocusEditView()
{
	iKeyESC = 1;
}


void DlgEditProperties::OnEnKillfocusEditView()
{
	try
	{
		szBefore = m_list_view.GetItemText(CLRow, CLColumn);
		m_edit_properties.GetWindowTextW(szAfter);

		szAfter.Trim(); szBefore.Trim();

		if (szAfter == szBefore) return;

		CString szval = m_list_view.GetItemText(CLRow, iGFLType);
		if (szval != EDTPRO_LINKERROR)
		{
			int nIndex = vecFilter[CLRow].iID;
			CString szpath = m_list_view.GetItemText(CLRow, iGFLPath);
			CString szType = ff->_GetTypeOfFile(szpath);
			szType.MakeLower();

			if (CLColumn == iGFLTitle)
			{
				vecItem[nIndex].szTitle = szAfter;
				if (ff->_SetProperties(szpath, 1, szAfter))
				{
					m_txt_title.SetWindowText(szAfter);
					pRgActive->PutItem(vecItem[nIndex].nRow, jcotNoidung, (_bstr_t)szAfter);
				}
			}
			else if (CLColumn == iGFLSubject)
			{
				vecItem[nIndex].szSubject = szAfter;
				if (ff->_SetProperties(szpath, 2, szAfter))
				{
					m_txt_subject.SetWindowText(szAfter);
				}
			}
			else if (CLColumn == iGFLTags)
			{
				vecItem[nIndex].szTags = szAfter;
				if (ff->_SetProperties(szpath, 3, szAfter))
				{
					m_txt_tags.SetWindowText(szAfter);
				}
			}
			else if (CLColumn == iGFLCategories)
			{
				vecItem[nIndex].szCategories = szAfter;
				if (ff->_SetProperties(szpath, 4, szAfter))
				{
					m_txt_categories.SetWindowText(szAfter);
				}
			}
			else if (CLColumn == iGFLComments)
			{
				vecItem[nIndex].szComments = szAfter;
				if (ff->_SetProperties(szpath, 5, szAfter))
				{
					m_txt_commments.SetWindowText(szAfter);
				}
			}
			if (CLColumn == iGFLAuthors)
			{
				vecItem[nIndex].szAuthors = szAfter;
				szval = szAfter;

				if (szType == L"pdf")
				{
					szval += (L"|" + m_list_view.GetItemText(CLRow, iGFLCompany));
				}

				if (ff->_SetProperties(szpath, 6, szval))
				{
					m_txt_authors.SetWindowText(szAfter);
				}
			}
			else if (CLColumn == iGFLCompany)
			{
				vecItem[nIndex].szCompany = szAfter;
				szval = szAfter;

				if (szType == L"pdf")
				{
					szval = m_list_view.GetItemText(CLRow, iGFLAuthors) + L"|" + szAfter;
				}
				
				if (ff->_SetProperties(szpath, 7, szval))
				{
					m_txt_company.SetWindowText(szAfter);
				}
			}			

			m_list_view.SetItemText(CLRow, iGFLType, EDTPRO_SUCCESSFUL);
			m_list_view.SetCellColors(CLRow, iGFLType, RGBGREEN017680, rgbWhite);
		}
	}
	catch (...) {}
}

void DlgEditProperties::OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult)
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

			if (tongfilter <= iz)
			{
				iz = tongfilter;
				iStopload = 0;
			}
			else iStopload = 1;

			// Thêm dữ liệu vào list kết quả
			_AddItemsListKetqua(ibuocnhay*(lanshow - 1), iz);

			m_list_view.EnsureVisible(ncount + (int)(b / a) - 5, FALSE);
		}
	}
	catch (...) {}
}

void DlgEditProperties::OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMHEADER pNMListHeader = reinterpret_cast<LPNMHEADER>(pNMHDR);
	int pSubItem = pNMListHeader->iItem;
	if (pSubItem == 0) pNMListHeader->pitem->cxy = 22;
	*pResult = 0;
}

void DlgEditProperties::OnPropertiesOpenfile()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ShellExecute(NULL, L"open", szpath, NULL, NULL, SW_SHOWMAXIMIZED);
	}
}

void DlgEditProperties::OnPropertiesOpenfolder()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
		if (szpath != L"") ff->_ShellExecuteSelectFile(szpath);
	}
}

void DlgEditProperties::OnPropertiesCopy()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		szCopy = m_list_view.GetItemText(nItem, nSubItem);
	}
}

void DlgEditProperties::OnPropertiesPaste()
{
	if (nItem >= 0 && nSubItem >= 0)
	{
		vector<int> vecRow;

		CString szval = L"";
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 0)
		{
			// Đổ dữ liệu vào ô text properties
			_SetTextCopy();

			// Đổ dữ liệu lên sheet
			if (nSubItem == iGFLTitle)
			{
				ff->_xlPutScreenUpdating(false);

				int nIndex = -1;
				for (int i = 0; i < nSelect; i++)
				{
					nIndex = vecFilter[vecRow[i]].iID;
					pRgActive->PutItem(vecItem[nIndex].nRow, jcotNoidung, (_bstr_t)szCopy);
				}

				ff->_xlPutScreenUpdating(true);
			}

			for (int i = 0; i < nSelect; i++)
			{
				// Đổ dữ liệu lên listview
				m_list_view.SetItemText(vecRow[i], nSubItem, szCopy);

				// Lưu dữ liệu vào file document				
				_SetSaveDocument(vecRow[i]);
			}
		}

		vecRow.clear();
	}
}

void DlgEditProperties::_SetSaveDocument(int num)
{
	try
	{
		int nIndex = vecFilter[num].iID;
		CString szpath = m_list_view.GetItemText(num, iGFLPath);
		CString sztype = ff->_GetTypeOfFile(szpath);
		CString szval = L"";

		if (nSubItem == iGFLTitle)
		{
			vecItem[nIndex].szTitle = szCopy;
			ff->_SetProperties(szpath, 1, szCopy);
		}
		else if (nSubItem == iGFLSubject)
		{
			vecItem[nIndex].szSubject = szCopy;
			ff->_SetProperties(szpath, 2, szCopy);
		}
		else if (nSubItem == iGFLTags)
		{
			vecItem[nIndex].szTags = szCopy;
			ff->_SetProperties(szpath, 3, szCopy);
		}
		else if (nSubItem == iGFLCategories)
		{
			vecItem[nIndex].szCategories = szCopy;
			ff->_SetProperties(szpath, 4, szCopy);
		}
		else if (nSubItem == iGFLComments)
		{
			vecItem[nIndex].szComments = szCopy;
			ff->_SetProperties(szpath, 5, szCopy);
		}
		else if (nSubItem == iGFLAuthors)
		{
			vecItem[nIndex].szAuthors = szCopy;
			szval = szCopy;

			if (sztype == L"pdf")
			{
				szval += (L"|" + m_list_view.GetItemText(num, iGFLCompany));
			}

			ff->_SetProperties(szpath, 6, szval);
		}
		else if (nSubItem == iGFLCompany)
		{
			vecItem[nIndex].szCompany = szCopy;
			szval = szCopy;

			if (sztype == L"pdf")
			{
				szval = m_list_view.GetItemText(num, iGFLAuthors) + L"|" + szAfter;
			}

			ff->_SetProperties(szpath, 7, szval);
		}
	}
	catch (...) {}
}

void DlgEditProperties::_SetTextCopy()
{
	if (nSubItem == iGFLTitle) _SetColorCopy(m_txt_title);
	else if (nSubItem == iGFLSubject) _SetColorCopy(m_txt_subject);
	else if (nSubItem == iGFLTags) _SetColorCopy(m_txt_tags);
	else if (nSubItem == iGFLCategories) _SetColorCopy(m_txt_categories);
	else if (nSubItem == iGFLComments) _SetColorCopy(m_txt_commments);
	else if (nSubItem == iGFLAuthors) _SetColorCopy(m_txt_authors);
	else if (nSubItem == iGFLCompany) _SetColorCopy(m_txt_company);
}

void DlgEditProperties::_SetColorCopy(CColorEdit &m_txt_copy)
{
	m_txt_copy.SetTextColor(rgbBlue);
	m_txt_copy.SetWindowText(szCopy);
}

int DlgEditProperties::_GetAllSelectedItems(vector<int> &vecRow)
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

			int nsize = (int)vecRow.size();
			if (nsize == ncount)
			{
				for (int i = ncount; i < tongfilter; i++) vecRow.push_back(i);
			}
		}

		return (int)vecRow.size();
	}
	catch (...) {}
	return 0;
}

void DlgEditProperties::_SetSpliterPane()
{
	try
	{
		CRect rc;
		CWnd* pWnd = GetDlgItem(EDPR_SPLITER_VER);
		pWnd->GetWindowRect(rc);
		ScreenToClient(rc);
		BOOL bRet = m_wndSplitterVer.Create(WS_CHILD | WS_VISIBLE,
			rc, this, EDPR_SPLITER_VER, SPS_VERTICAL | SPS_DELTA_NOTIFY, rgbLightGray);
		if (bRet == TRUE)
		{
			//  register windows for splitter
			this->m_wndSplitterVer.RegisterLinkedWindow(SPLS_LINKED_LEFT, GetDlgItem(EDPR_GROUP_INFO));
			this->m_wndSplitterVer.RegisterLinkedWindow(SPLS_LINKED_RIGHT, GetDlgItem(EDPR_GROUP_LIST));

			//  relayout the splotter to make them good look
			this->m_wndSplitterVer.Relayout();
			this->_ResizeDialog();
		}
	}
	catch (...) {}
}

void DlgEditProperties::_ResizeDialog()
{
	try
	{
		if (!IsWindow(m_wndSplitterVer.GetSafeHwnd())) return;

		CRect rMAIN;
		this->GetClientRect(rMAIN);

		// Xác định cao của mỗi ô text
		int hh = rMAIN.Height();
		int nDCTen = 22, nDCChung = 22;
		int isy = GetSystemMetrics(SM_CYSCREEN);		
		if ((int)(3*hh/2) >= isy) 
		{
			nDCTen = 44;
			nDCChung = (int)(hh/9);
		}
		
		CRect rSPLIT;
		m_wndSplitterVer.GetWindowRect(rSPLIT);
		this->ScreenToClient(rSPLIT);
		rSPLIT.top = rMAIN.top;
		rSPLIT.bottom = rMAIN.bottom;
		this->m_wndSplitterVer.MoveWindow(rSPLIT, TRUE);

		CRect rLEFT;
		rLEFT.left = rMAIN.left + 8;
		rLEFT.top = rMAIN.top + 8;
		rLEFT.right = rSPLIT.left - 2;
		rLEFT.bottom = rMAIN.bottom - 8;
		this->_MoveDlgItem(EDPR_GROUP_INFO, rLEFT, TRUE);

		CRect rBtnUp;
		rBtnUp.bottom = rLEFT.bottom - 10;
		rBtnUp.left = rLEFT.left + 9;
		rBtnUp.top = rBtnUp.bottom - 29;
		rBtnUp.right = rBtnUp.left + 32;
		this->_MoveDlgItem(EDPR_BTN_UP, rBtnUp, TRUE);

		CRect rBtnDown;
		rBtnDown.bottom = rLEFT.bottom - 10;
		rBtnDown.left = rBtnUp.right + 6;
		rBtnDown.top = rBtnDown.bottom - 29;
		rBtnDown.right = rBtnDown.left + 32;
		this->_MoveDlgItem(EDPR_BTN_DOWN, rBtnDown, TRUE);

		CRect rBtnClear;
		rBtnClear.bottom = rLEFT.bottom - 10;
		rBtnClear.left = rBtnDown.right + 6;
		rBtnClear.top = rBtnClear.bottom - 29;
		rBtnClear.right = rBtnClear.left + 32;
		this->_MoveDlgItem(EDPR_BTN_CLEAR, rBtnClear, TRUE);

		CRect rBtnSaved;
		rBtnSaved.bottom = rLEFT.bottom - 10;
		rBtnSaved.left = rBtnClear.right + 6;
		rBtnSaved.top = rBtnSaved.bottom - 29;
		rBtnSaved.right = rLEFT.right - 9;
		this->_MoveDlgItem(EDPR_BTN_SAVED, rBtnSaved, TRUE);

		CRect rLblComp;
		rLblComp.bottom = rBtnSaved.top - 7;
		rLblComp.left = rLEFT.left + 9;
		rLblComp.top = rLblComp.bottom - 22;
		rLblComp.right = rLblComp.left + 70;
		this->_MoveDlgItem(EDPR_STT_8, rLblComp, TRUE);

		CRect rTxtComp;
		rTxtComp.bottom = rBtnSaved.top - 7;
		rTxtComp.left = rLblComp.right + 4;
		rTxtComp.top = rTxtComp.bottom - 22;
		rTxtComp.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_COMPANY, rTxtComp, TRUE);

		CRect rLblAuth;
		rLblAuth.bottom = rTxtComp.top - 7;
		rLblAuth.left = rLEFT.left + 9;
		rLblAuth.top = rLblAuth.bottom - 22;
		rLblAuth.right = rLblAuth.left + 70;
		this->_MoveDlgItem(EDPR_STT_7, rLblAuth, TRUE);

		CRect rTxtAuth;
		rTxtAuth.bottom = rTxtComp.top - 7;
		rTxtAuth.left = rLblAuth.right + 4;
		rTxtAuth.top = rTxtAuth.bottom - 22;
		rTxtAuth.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_AUTHORS, rTxtAuth, TRUE);

		CRect rLineHoz;
		rLineHoz.bottom = rTxtAuth.top - 7;
		rLineHoz.left = rLEFT.left + 5;
		rLineHoz.top = rLineHoz.bottom - 2;
		rLineHoz.right = rLEFT.right - 5;
		this->_MoveDlgItem(EDPR_LINE_HOZ, rLineHoz, TRUE);

		CRect rTxtCmt;
		rTxtCmt.bottom = rLineHoz.top - 7;
		rTxtCmt.left = rLEFT.left + 83;
		rTxtCmt.top = rTxtCmt.bottom - nDCChung;
		rTxtCmt.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_COMMENTS, rTxtCmt, TRUE);

		CRect rLblCmt;
		rLblCmt.top = rTxtCmt.top;
		rLblCmt.bottom = rLblCmt.top + 22;
		rLblCmt.left = rLEFT.left + 9;		
		rLblCmt.right = rLblCmt.left + 70;
		this->_MoveDlgItem(EDPR_STT_6, rLblCmt, TRUE);

		CRect rTxtCate;
		rTxtCate.bottom = rTxtCmt.top - 7;
		rTxtCate.left = rLEFT.left + 83;
		rTxtCate.top = rTxtCate.bottom - nDCChung;
		rTxtCate.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_CATEGORIES, rTxtCate, TRUE);

		CRect rLblCate;
		rLblCate.top = rTxtCate.top;
		rLblCate.bottom = rLblCate.top + 22;
		rLblCate.left = rLEFT.left + 9;		
		rLblCate.right = rLblCate.left + 70;
		this->_MoveDlgItem(EDPR_STT_5, rLblCate, TRUE);

		CRect rTxtTags;
		rTxtTags.bottom = rTxtCate.top - 7;
		rTxtTags.left = rLEFT.left + 83;
		rTxtTags.top = rTxtTags.bottom - nDCChung;
		rTxtTags.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_TAGS, rTxtTags, TRUE);

		CRect rLblTags;
		rLblTags.top = rTxtTags.top;
		rLblTags.bottom = rLblTags.top + 22;
		rLblTags.left = rLEFT.left + 9;		
		rLblTags.right = rLblTags.left + 70;
		this->_MoveDlgItem(EDPR_STT_4, rLblTags, TRUE);

		CRect rTxtSubj;
		rTxtSubj.bottom = rTxtTags.top - 7;
		rTxtSubj.left = rLEFT.left + 83;
		rTxtSubj.top = rTxtSubj.bottom - nDCChung;
		rTxtSubj.right = rLEFT.right - 10;
		this->_MoveDlgItem(EDPR_TXT_SUBJECT, rTxtSubj, TRUE);

		CRect rLblSubj;
		rLblSubj.top = rTxtSubj.top;
		rLblSubj.bottom = rLblSubj.top + 22;
		rLblSubj.left = rLEFT.left + 9;		
		rLblSubj.right = rLblSubj.left + 70;
		this->_MoveDlgItem(EDPR_STT_3, rLblSubj, TRUE);

		CRect rTxtName;
		rTxtName.left = rLEFT.left + 9;
		rTxtName.top = rLEFT.top + 18;
		rTxtName.right = rLEFT.right - 10;
		rTxtName.bottom = rTxtName.top + nDCTen;
		this->_MoveDlgItem(EDPR_TXT_FILENAME, rTxtName, TRUE);

		CRect rTxtTitle;
		rTxtTitle.left = rLEFT.left + 83;
		rTxtTitle.top = rTxtName.bottom + 7;
		rTxtTitle.right = rLEFT.right - 10;
		rTxtTitle.bottom = rTxtSubj.top - 7;
		this->_MoveDlgItem(EDPR_TXT_TITLE, rTxtTitle, TRUE);

		CRect rLblTitle;
		rLblTitle.top = rTxtTitle.top;
		rLblTitle.bottom = rLblTitle.top + 22;
		rLblTitle.left = rLEFT.left + 9;		
		rLblTitle.right = rLblTitle.left + 70;
		
		this->_MoveDlgItem(EDPR_STT_2, rLblTitle, TRUE);

		/*****************************************************************/

		CRect rRIGHT;
		rRIGHT.left = rSPLIT.right + 2;
		rRIGHT.top = rMAIN.top + 8;
		rRIGHT.right = rMAIN.right - 8;
		rRIGHT.bottom = rMAIN.bottom - 8;
		this->_MoveDlgItem(EDPR_GROUP_LIST, rRIGHT, TRUE);

		CRect rLblSearch;
		rLblSearch.left = rRIGHT.left + 9;
		rLblSearch.top = rRIGHT.top + 20;
		rLblSearch.right = rLblSearch.left + 75;
		rLblSearch.bottom = rLblSearch.top + 16;
		this->_MoveDlgItem(EDPR_LBL_SEARCH, rLblSearch, TRUE);

		CRect rBtnSearch;
		rBtnSearch.right = rRIGHT.right - 11;
		rBtnSearch.top = rRIGHT.top + 19;
		rBtnSearch.left = rBtnSearch.right - 21;
		rBtnSearch.bottom = rBtnSearch.top + 21;
		this->_MoveDlgItem(EDPR_BTN_SEARCH, rBtnSearch, TRUE);

		CRect rTxtSearch;
		rTxtSearch.left = rLblSearch.right + 7;
		rTxtSearch.top = rRIGHT.top + 18;
		rTxtSearch.right = rBtnSearch.left - 6;
		rTxtSearch.bottom = rTxtSearch.top + 22;
		this->_MoveDlgItem(EDPR_TXT_SEARCH, rTxtSearch, TRUE);

		CRect rBtnCancel;
		rBtnCancel.bottom = rRIGHT.bottom - 10;
		rBtnCancel.right = rRIGHT.right - 9;
		rBtnCancel.top = rBtnCancel.bottom - 29;
		rBtnCancel.left = rBtnCancel.right - 94;
		this->_MoveDlgItem(EDPR_BTN_CANCEL, rBtnCancel, TRUE);

		CRect rBtnOk;
		rBtnOk.bottom = rRIGHT.bottom - 10;
		rBtnOk.right = rBtnCancel.left - 6;
		rBtnOk.top = rBtnOk.bottom - 29;
		rBtnOk.left = rBtnOk.right - 210;
		this->_MoveDlgItem(EDPR_BTN_UPDATE, rBtnOk, TRUE);

		CRect rLblProperties;
		rLblProperties.left = rRIGHT.left + 9;
		rLblProperties.bottom = rRIGHT.bottom - 10;
		rLblProperties.right = rBtnOk.left - 6;
		rLblProperties.top = rLblProperties.bottom - 29;
		this->_MoveDlgItem(EDPR_LBL_PROGRESS, rLblProperties, TRUE);

		CRect rList;
		rList.left = rRIGHT.left + 10;
		rList.top = rTxtSearch.bottom + 9;
		rList.right = rRIGHT.right - 10;
		rList.bottom = rBtnOk.top - 7;
		this->_MoveDlgItem(EDPR_LIST_VIEW, rList, TRUE);
	}
	catch (...) {}
}

void DlgEditProperties::OnMaxMinInfo(NMHDR* pNMHDR, LRESULT* pResult)
{
	CRect rcTree;
	m_lbl_info.GetWindowRect(rcTree);
	this->ScreenToClient(rcTree);

	CRect rcList;
	m_lbl_kq.GetWindowRect(rcList);
	this->ScreenToClient(rcList);

	SPC_NM_MAXMINPOS* pNewMaxMinPos = (SPC_NM_MAXMINPOS*)pNMHDR;
	if (EDPR_SPLITER_VER == pNMHDR->idFrom)
	{
		pNewMaxMinPos->lMin = rcTree.left + 50;
		pNewMaxMinPos->lMax = rcList.right - 50;
	}
}

void DlgEditProperties::OnSplitterDelta(NMHDR* pNMHDR, LRESULT* pResult)
{
	SPC_NM_DELTA* pDelta = (SPC_NM_DELTA*)pNMHDR;
	if (NULL == pDelta) return;
	this->_ResizeDialog();
}

void DlgEditProperties::_MoveDlgItem(int nID, const CRect& rcPos, BOOL bRepaint)
{
	CWnd* pWnd = this->GetDlgItem(nID);
	if (NULL == pWnd) return;
	pWnd->MoveWindow(rcPos, bRepaint);
}

void DlgEditProperties::OnBnClickedStt2()
{
	_SetBkgColorText(0, m_txt_title);
}

void DlgEditProperties::OnBnClickedStt3()
{
	_SetBkgColorText(1, m_txt_subject);
}

void DlgEditProperties::OnBnClickedStt4()
{
	_SetBkgColorText(2, m_txt_tags);
}

void DlgEditProperties::OnBnClickedStt5()
{
	_SetBkgColorText(3, m_txt_categories);
}

void DlgEditProperties::OnBnClickedStt6()
{
	_SetBkgColorText(4, m_txt_commments);
}

void DlgEditProperties::OnBnClickedStt7()
{
	_SetBkgColorText(5, m_txt_authors);
}

void DlgEditProperties::OnBnClickedStt8()
{
	_SetBkgColorText(6, m_txt_company);
}

void DlgEditProperties::_SetBkgColorText(int nIndex, CColorEdit &m_txt_read)
{
	try
	{
		if (!blRead[nIndex]) m_txt_read.SetBkColor(rgbYellow);
		else m_txt_read.SetBkColor(rgbWhite);
		blRead[nIndex] = !blRead[nIndex];

		// EDPR_TXT_TITLE = 2135
		GotoDlgCtrl(GetDlgItem(2135 + 2*nIndex));
	}
	catch (...) {}
}

void DlgEditProperties::OnBnClickedBtnUp()
{
	if (nItem > 0)
	{
		_SetIconButtonSaved(0);

		m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
		m_list_view.SetItemState(nItem - 1, LVIS_SELECTED, LVIS_SELECTED);

		_LoadProperties(nItem - 1);
		nItem--;

		GotoDlgCtrl(GetDlgItem(EDPR_TXT_TITLE));
	}
}

void DlgEditProperties::OnBnClickedBtnDown()
{
	int ncount = m_list_view.GetItemCount();
	if (nItem < ncount - 1)
	{
		_SetIconButtonSaved(0);

		m_list_view.SetItemState(nItem, 0, LVIS_SELECTED);
		m_list_view.SetItemState(nItem + 1, LVIS_SELECTED, LVIS_SELECTED);

		_LoadProperties(nItem + 1);
		nItem++;

		GotoDlgCtrl(GetDlgItem(EDPR_TXT_TITLE));
	}
}

void DlgEditProperties::OnBnClickedBtnClear()
{
	m_txt_title.SetWindowText(L"");
	m_txt_subject.SetWindowText(L"");
	m_txt_tags.SetWindowText(L"");
	m_txt_categories.SetWindowText(L"");
	m_txt_commments.SetWindowText(L"");
	m_txt_authors.SetWindowText(L"");
	m_txt_company.SetWindowText(L"");

	_SetIconButtonSaved(0);
	_SetAllBannerMulti(false);

	GotoDlgCtrl(GetDlgItem(EDPR_TXT_TITLE));
}

void DlgEditProperties::OnBnClickedBtnSaved()
{
	try
	{
		if (nItem >= 0)
		{
			m_lbl_progress.SetWindowText(L"Đang cập nhật nội dung. Vui lòng đợi trong giây lát...");

			CString szBefore[8] = { L"", L"" ,L"", L"", L"", L"", L"", L"" };
			szBefore[0] = m_list_view.GetItemText(nItem, iGFLTen);
			szBefore[1] = m_list_view.GetItemText(nItem, iGFLTitle);
			szBefore[2] = m_list_view.GetItemText(nItem, iGFLSubject);
			szBefore[3] = m_list_view.GetItemText(nItem, iGFLTags);
			szBefore[4] = m_list_view.GetItemText(nItem, iGFLCategories);
			szBefore[5] = m_list_view.GetItemText(nItem, iGFLComments);
			szBefore[6] = m_list_view.GetItemText(nItem, iGFLAuthors);
			szBefore[7] = m_list_view.GetItemText(nItem, iGFLCompany);

			CString szName = L"";
			CString szAfter[9] = { L"", L"" ,L"", L"", L"", L"", L"", L"", L"" };
			m_txt_filename.GetWindowTextW(szAfter[0]);
			m_txt_title.GetWindowTextW(szAfter[1]);
			m_txt_subject.GetWindowTextW(szAfter[2]);
			m_txt_tags.GetWindowTextW(szAfter[3]);
			m_txt_categories.GetWindowTextW(szAfter[4]);
			m_txt_commments.GetWindowTextW(szAfter[5]);
			m_txt_authors.GetWindowTextW(szAfter[6]);
			m_txt_company.GetWindowTextW(szAfter[7]);

			if (szAfter[1] == BAMUL_TITLE) szAfter[1] = L"";
			if (szAfter[2] == BAMUL_SUBJECT) szAfter[2] = L"";
			if (szAfter[3] == BAMUL_TAGS) szAfter[3] = L"";
			if (szAfter[4] == BAMUL_CATEGORIES) szAfter[4] = L"";
			if (szAfter[5] == BAMUL_COMMENTS) szAfter[5] = L"";

			if (szAfter[0] == szBefore[0])
			{
				bool bl = false;
				int nsize = getSizeArray(szBefore);
				for (int i = 1; i < nsize; i++)
				{
					if (szAfter[i] != szBefore[i])
					{
						bl = true;
						break;
					}
				}

				if (bl)
				{
					CString szpath = m_list_view.GetItemText(nItem, iGFLPath);
					if (ff->_FileExistsUnicode(szpath, 0))
					{
						if (ff->_SetAllProperties(szpath,
							szAfter[1], szAfter[2], szAfter[3], szAfter[4],
							szAfter[5], szAfter[6], szAfter[7]))
						{
							m_list_view.SetItemText(nItem, iGFLTitle, szAfter[1]);
							m_list_view.SetItemText(nItem, iGFLSubject, szAfter[2]);
							m_list_view.SetItemText(nItem, iGFLTags, szAfter[3]);
							m_list_view.SetItemText(nItem, iGFLCategories, szAfter[4]);
							m_list_view.SetItemText(nItem, iGFLComments, szAfter[5]);
							m_list_view.SetItemText(nItem, iGFLAuthors, szAfter[6]);
							m_list_view.SetItemText(nItem, iGFLCompany, szAfter[7]);

							m_list_view.SetItemText(nItem, iGFLType, EDTPRO_SUCCESSFUL);
							m_list_view.SetCellColors(nItem, iGFLType, RGBGREEN017680, rgbWhite);

							int nIndex = vecFilter[nItem].iID;
							pRgActive->PutItem(vecItem[nIndex].nRow, jcotNoidung, (_bstr_t)szAfter[1]);

							_SetIconButtonSaved(1);
						}
					}
				}
			}

			// Tiếp tục kiểm tra xem có chọn nhiều dữ liệu không?
			_SaveAllFieldsFix();

			m_lbl_progress.SetWindowText(L"");
		}
	}
	catch (...) {}
}

void DlgEditProperties::_SaveAllFieldsFix()
{
	try
	{
		vector<int> vecRow;
		int nSelect = _GetAllSelectedItems(vecRow);
		if (nSelect > 1)
		{
			// Tiếp tục kiểm tra xem có trường nào cố định không?
			vector<int> vecFix;
			int nsize = getSizeArray(blRead);
			for (int i = 0; i < nsize; i++)
			{
				if (blRead[i]) vecFix.push_back(i);
			}

			CString szval = L"", szpath = L"";
			int nIndex = 0, isize = (int)vecFix.size();
			if (isize > 0)
			{
				// Bỏ quả 0, bắt đầu từ i = 1
				for (int i = 1; i < nSelect; i++)
				{
					szpath = m_list_view.GetItemText(vecRow[i], iGFLPath);
					for (int j = 0; j < isize; j++)
					{
						switch (vecFix[j])
						{
							case 0:
							{
								m_txt_title.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLTitle, szval);

								nIndex = vecFilter[vecRow[i]].iID;
								pRgActive->PutItem(vecItem[nIndex].nRow, jcotNoidung, (_bstr_t)szval);
							}
							break;

							case 1:
							{
								m_txt_subject.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLSubject, szval);
							}
							break;

							case 2:
							{
								m_txt_tags.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLTags, szval);
							}
							break;

							case 3:
							{
								m_txt_categories.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLCategories, szval);
							}
							break;

							case 4:
							{
								m_txt_commments.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLComments, szval);
							}
							break;

							case 5:
							{
								m_txt_authors.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLAuthors, szval);
							}
							break;

							case 6:
							{
								m_txt_company.GetWindowTextW(szval);
								m_list_view.SetItemText(vecRow[i], iGFLCompany, szval);
							}
							break;

							default:
								break;
						}

						ff->_SetProperties(szpath, vecFix[j] + 1, szval);
					}

					m_list_view.SetItemText(vecRow[i], iGFLType, EDTPRO_SUCCESSFUL);
					m_list_view.SetCellColors(vecRow[i], iGFLType, RGBGREEN017680, rgbWhite);
				}
			}
		}

		vecRow.clear();
	}
	catch (...) {}
}


void DlgEditProperties::_SetIconButtonSaved(int itype)
{
	if (itype == 0)
	{
		CString szval = L"";
		m_btn_saved.GetWindowTextW(szval);
		if (szval != L"Cập nhật thông tin")
		{
			m_btn_saved.SetIcon(IDI_ICON_SAVED, 24, 24);
			m_btn_saved.SetWindowText(L"Cập nhật thông tin");
		}
	}
	else
	{
		m_btn_saved.SetIcon(IDI_ICON_OK, 24, 24);
		m_btn_saved.SetWindowText(L"Cập nhật thành công!");
	}
}

void DlgEditProperties::_SetBannerTextMultiLine(CColorEdit &m_txt_multi, bool blSetfocus, CString szBanner)
{
	try
	{
		CString szval = L"";
		m_txt_multi.GetWindowTextW(szval);
		szval.Trim();

		if (blSetfocus)
		{
			m_txt_multi.SetTextColor(rgbBlue);
			if (szval == szBanner) m_txt_multi.SetWindowTextW(L"");
		}
		else
		{
			if (szval == L"")
			{
				m_txt_multi.SetTextColor(RGBGRAY929292);
				m_txt_multi.SetWindowTextW(szBanner);
			}
			else
			{
				if (szval != szBanner) m_txt_multi.SetTextColor(rgbBlue);
				else m_txt_multi.SetTextColor(RGBGRAY929292);
			}
		}
	}
	catch (...) {}
}

void DlgEditProperties::_SetAllBannerMulti(bool blFocus)
{
	try
	{
		_SetBannerTextMultiLine(m_txt_title, false, BAMUL_TITLE);
		_SetBannerTextMultiLine(m_txt_subject, false, BAMUL_SUBJECT);
		_SetBannerTextMultiLine(m_txt_tags, false, BAMUL_TAGS);
		_SetBannerTextMultiLine(m_txt_categories, false, BAMUL_CATEGORIES);
		_SetBannerTextMultiLine(m_txt_commments, false, BAMUL_COMMENTS);
	}
	catch (...) {}
}

void DlgEditProperties::OnEnSetfocusTxtTitle()
{
	_SetBannerTextMultiLine(m_txt_title, true, BAMUL_TITLE);
}

void DlgEditProperties::OnEnKillfocusTxtTitle()
{
	_SetBannerTextMultiLine(m_txt_title, false, BAMUL_TITLE);
}

void DlgEditProperties::OnEnSetfocusTxtSubject()
{
	_SetBannerTextMultiLine(m_txt_subject, true, BAMUL_SUBJECT);
}


void DlgEditProperties::OnEnKillfocusTxtSubject()
{
	_SetBannerTextMultiLine(m_txt_subject, false, BAMUL_SUBJECT);
}


void DlgEditProperties::OnEnSetfocusTxtTags()
{
	_SetBannerTextMultiLine(m_txt_tags, true, BAMUL_TAGS);
}


void DlgEditProperties::OnEnKillfocusTxtTags()
{
	_SetBannerTextMultiLine(m_txt_tags, false, BAMUL_TAGS);
}


void DlgEditProperties::OnEnSetfocusTxtCategories()
{
	_SetBannerTextMultiLine(m_txt_categories, true, BAMUL_CATEGORIES);
}


void DlgEditProperties::OnEnKillfocusTxtCategories()
{
	_SetBannerTextMultiLine(m_txt_categories, false, BAMUL_CATEGORIES);
}


void DlgEditProperties::OnEnSetfocusTxtComments()
{
	_SetBannerTextMultiLine(m_txt_commments, true, BAMUL_COMMENTS);
}

void DlgEditProperties::OnEnKillfocusTxtComments()
{
	_SetBannerTextMultiLine(m_txt_commments, false, BAMUL_COMMENTS);
}

void DlgEditProperties::OnEnChangeTxtTitle()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtSubject()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtTags()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtCategories()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtComments()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtAuthors()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::OnEnChangeTxtCompany()
{
	_SetIconButtonSaved(0);
}

void DlgEditProperties::_GetGroupRowStartEnd(int &iRowStart, int &iRowEnd)
{
	try
	{
		CString szval = L"";
		for (int i = iRowActive; i >= iRowStart; i--)
		{
			szval = ff->_xlGIOR(pRgActive, i, jcotSTT, L"");
			if (_wtoi(szval) == 0)
			{
				szval = ff->_xlGIOR(pRgActive, i, jcotLV, L"");
				if (szval != L"")
				{
					iRowStart = i;
					for (int j = iRowActive + 1; j <= iRowEnd; j++)
					{
						szval = ff->_xlGIOR(pRgActive, j, jcotSTT, L"");
						if (_wtoi(szval) == 0)
						{
							szval = ff->_xlGIOR(pRgActive, j, jcotLV, L"");
							if (szval != L"")
							{
								iRowEnd = j;
								return;
							}
						}
					}
				}
			}
		}
	}
	catch (...) {}
}

