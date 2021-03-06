#include "pch.h"
#include "ListCtrlEx.h"

CListCtrlEx::EditorInfo::EditorInfo()
	:m_pfnInitEditor(NULL)
	, m_pfnEndEditor(NULL)
	, m_pWnd(NULL)
	, m_bReadOnly(false)
{
}
CListCtrlEx::EditorInfo::EditorInfo(PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd *pWnd)
	:m_pfnInitEditor(pfnInitEditor)
	, m_pfnEndEditor(pfnEndEditor)
	, m_pWnd(pWnd)
	, m_bReadOnly(false)
{
}

CListCtrlEx::CellInfo::CellInfo(int nColumn)
	:m_clrBack(-1)
	, m_clrText(-1)
	, m_dwUserData(NULL)
	, m_nColumn(nColumn)
{

}
CListCtrlEx::CellInfo::CellInfo(int nColumn, COLORREF clrBack, COLORREF clrText, DWORD_PTR dwUserData)
	:m_clrBack(clrBack)
	, m_clrText(clrText)
	, m_dwUserData(dwUserData)
	, m_nColumn(nColumn)
{
}
CListCtrlEx::CellInfo::CellInfo(int nColumn, EditorInfo eiEditor, COLORREF clrBack, COLORREF clrText, DWORD_PTR dwUserData)
	:m_clrBack(clrBack)
	, m_clrText(clrText)
	, m_dwUserData(dwUserData)
	, m_eiEditor(eiEditor)
	, m_nColumn(nColumn)
{
}
CListCtrlEx::CellInfo::CellInfo(int nColumn, EditorInfo eiEditor, DWORD_PTR dwUserData)
	:m_clrBack(-1)
	, m_clrText(-1)
	, m_dwUserData(dwUserData)
	, m_eiEditor(eiEditor)
	, m_nColumn(nColumn)
{
}

CListCtrlEx::ColumnInfo::ColumnInfo(int nColumn)
	:m_eiEditor()
	, m_clrBack(-1)
	, m_clrText(-1)
	, m_nColumn(nColumn)
	, m_eSort(None)
	, m_eCompare(NotSet)
	, m_fnCompare(NULL)
{
}
CListCtrlEx::ColumnInfo::ColumnInfo(int nColumn, PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd *pWnd)
	:m_eiEditor(pfnInitEditor, pfnEndEditor, pWnd)
	, m_nColumn(nColumn)
	, m_clrBack(-1)
	, m_clrText(-1)
	, m_eSort(None)
	, m_eCompare(NotSet)
	, m_fnCompare(NULL)
{
}
CListCtrlEx::ItemData::ItemData(DWORD_PTR dwUserData)
	:m_clrBack(0xFFFFFFFF)
	, m_clrText(0xFFFFFFFF)
	, m_dwUserData(dwUserData)
{
}
CListCtrlEx::ItemData::ItemData(COLORREF clrBack, COLORREF clrText, DWORD_PTR dwUserData)
	: m_clrBack(clrBack)
	, m_clrText(clrText)
	, m_dwUserData(dwUserData)
{
}
CListCtrlEx::ItemData::~ItemData()
{
	while (m_aCellInfo.GetCount() > 0)
	{
		CellInfo *pInfo = (CellInfo*)m_aCellInfo.GetAt(0);
		m_aCellInfo.RemoveAt(0);
		delete pInfo;
	}
}
// CListCtrlEx

IMPLEMENT_DYNAMIC(CListCtrlEx, CListCtrl)

CListCtrlEx::CListCtrlEx()
	:m_pEditor(NULL)
	, m_nEditingRow(-1)
	, m_nEditingColumn(-1)
	, m_msgHook()
	, m_nRow(-1)
	, m_nColumn(-1)
	, m_bHandleDelete(FALSE)
	, m_nSortColumn(-1)
	, m_fnCompare(NULL)
	, m_hAccel(NULL)
	, m_dwSortData(NULL)
	, m_bInvokeAddNew(FALSE)
	, m_nDropIndex(-1)
	, m_pDragImage(NULL)
	, m_nPrevDropIndex(-1)
	, m_uPrevDropState(NULL)
	, m_uScrollTimer(0)
	, m_ScrollDirection(scrollNull)
	, m_dwStyle(NULL)
{
#ifndef _WIN32_WCE
	EnableActiveAccessibility();
#endif

}
#define ID_EDITOR_CTRL		5000
CListCtrlEx::~CListCtrlEx()
{
	DeleteAllItemsData();
	DeleteAllColumnInfo();

	delete m_pDragImage;
	m_pDragImage = NULL;

	KillScrollTimer();
}

BEGIN_MESSAGE_MAP(CListCtrlEx, CListCtrl)
	ON_WM_DESTROY()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONUP()
	ON_WM_TIMER()
	ON_NOTIFY_REFLECT(LVN_BEGINDRAG, OnBeginDrag)
	ON_NOTIFY_REFLECT_EX(NM_DBLCLK, &CListCtrlEx::OnNMDblclk)
	ON_NOTIFY_REFLECT_EX(NM_CUSTOMDRAW, &CListCtrlEx::OnNMCustomdraw)
	ON_NOTIFY_REFLECT(LVN_COLUMNCLICK, &CListCtrlEx::OnHdnItemclick)
END_MESSAGE_MAP()

// CListCtrlEx message handlers

// Retrieves the data (lParam) associated with a particular item.
DWORD_PTR CListCtrlEx::GetItemData(int nItem) const
{
	ItemData* pData = (ItemData*)GetItemDataInternal(nItem);
	if (!pData) return NULL;
	return pData->m_dwUserData;
}
// Retrieves the data (lParam) associated with a particular item.
DWORD_PTR CListCtrlEx::GetItemDataInternal(int nItem) const
{
	return CListCtrl::GetItemData(nItem);
}
int CListCtrlEx::InsertItem(int nItem, LPCTSTR lpszItem)
{
	int ret = CListCtrl::InsertItem(nItem, lpszItem);
	SetItemData(ret, 0);
	return ret;
}
int CListCtrlEx::InsertItem(int nItem, LPCTSTR lpszItem, int nImage)
{
	int ret = CListCtrl::InsertItem(nItem, lpszItem, nImage);
	SetItemData(ret, 0);
	return ret;
}
int CListCtrlEx::InsertItem(const LVITEM* pItem)
{
	int ret = 0;
	LVITEM pI = *pItem;
	if (pItem && (pItem->mask & LVIF_PARAM))
	{
		pI.mask &= (~LVIF_PARAM);
	}
	ret = CListCtrl::InsertItem(&pI);
	SetItemData(ret, pItem->lParam);
	return ret;
}
// Sets the data (lParam) associated with a particular item.
BOOL CListCtrlEx::SetItemData(int nItem, DWORD_PTR dwData)
{
	ItemData* pData = (ItemData*)GetItemDataInternal(nItem);
	if (!pData)
	{
		pData = new ItemData(dwData);
		m_aItemData.Add(pData);
	}
	else
		pData->m_dwUserData = dwData;
	return CListCtrl::SetItemData(nItem, (DWORD_PTR)pData);
}

// Removes a single item from the control.
BOOL CListCtrlEx::DeleteItem(int nItem)
{
	DeleteItemData(nItem);
	return CListCtrl::DeleteItem(nItem);
}
// Removes all items from the control.
BOOL CListCtrlEx::DeleteAllItems()
{
	int nCount = this->GetItemCount();
	DeleteAllItemsData();
	return CListCtrl::DeleteAllItems();
}
BOOL CListCtrlEx::DeleteAllItemsData()
{
	while (m_aItemData.GetCount() > 0)
	{
		ItemData* pData = (ItemData*)m_aItemData.GetAt(0);
		if (pData) delete pData;
		m_aItemData.RemoveAt(0);
	}

	return TRUE;
}
BOOL CListCtrlEx::DeleteItemData(int nItem)
{
	if (nItem < 0 || nItem > GetItemCount()) return FALSE;
	ItemData* pData = (ItemData*)CListCtrl::GetItemData(nItem);

	INT_PTR nCount = m_aItemData.GetCount();
	for (INT_PTR i = 0; i < nCount && pData; i++)
	{
		if (m_aItemData.GetAt(i) == pData)
		{
			m_aItemData.RemoveAt(i);
			break;
		}
	}

	if (pData) delete pData;

	return TRUE;
}
BOOL CListCtrlEx::DeleteAllColumnInfo()
{
	while (m_aColumnInfo.GetCount() > 0)
	{
		ColumnInfo* pData = (ColumnInfo*)m_aColumnInfo.GetAt(0);
		if (pData) delete pData;
		m_aColumnInfo.RemoveAt(0);
	}

	return TRUE;
}
BOOL CListCtrlEx::DeleteColumnInfo(int nColumn)
{
	if (nColumn < 0 || nColumn > GetColumnCount()) return FALSE;
	ColumnInfo* pData = (ColumnInfo*)GetColumnInfo(nColumn);

	INT_PTR nCount = m_aColumnInfo.GetCount();
	for (INT_PTR i = 0; i < nCount && pData; i++)
	{
		if (m_aColumnInfo.GetAt(i) == pData)
		{
			m_aColumnInfo.RemoveAt(i);
			break;
		}
	}

	if (pData) delete pData;

	return TRUE;
}

void CListCtrlEx::SetColumnEditor(int nColumn, PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd* pWnd)
{
	ColumnInfo* pColInfo = GetColumnInfo(nColumn);
	if (!pColInfo)
	{
		pColInfo = new ColumnInfo(nColumn);
		m_aColumnInfo.Add(pColInfo);
	}
	pColInfo->m_eiEditor.m_pfnInitEditor = pfnInitEditor;
	pColInfo->m_eiEditor.m_pfnEndEditor = pfnEndEditor;
	pColInfo->m_eiEditor.m_pWnd = pWnd;

}

void CListCtrlEx::SetColumnEditor(int nColumn, CWnd* pWnd)
{
	SetColumnEditor(nColumn, NULL, NULL, pWnd);
}
void CListCtrlEx::SetRowEditor(int nRow, CWnd* pWnd)
{
	SetRowEditor(nRow, NULL, NULL, pWnd);
}
void CListCtrlEx::SetRowEditor(int nRow, PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd* pWnd)
{
	ItemData* pData = (ItemData*)GetItemDataInternal(nRow);
	if (!pData)
	{
		SetItemData(nRow, 0);
		pData = (ItemData*)GetItemDataInternal(nRow);
	}
	pData->m_eiEditor.m_pfnInitEditor = pfnInitEditor;
	pData->m_eiEditor.m_pfnEndEditor = pfnEndEditor;
	pData->m_eiEditor.m_pWnd = pWnd;

}

CListCtrlEx::ColumnInfo* CListCtrlEx::GetColumnInfo(int nColumn)
{
	INT_PTR nCount = m_aColumnInfo.GetCount();
	for (INT_PTR i = 0; i < nCount; i++)
	{
		ColumnInfo* pColInfo = (ColumnInfo*)m_aColumnInfo.GetAt(i);
		if (pColInfo->m_nColumn == nColumn) return pColInfo;
	}
	return NULL;
}
int CListCtrlEx::FindItem(LVFINDINFO* pFindInfo, int nStart) const
{
	if (pFindInfo->flags & LVIF_PARAM)
	{
		int nCount = GetItemCount();
		for (int i = nStart + 1; i < nCount; i++)
		{
			if (pFindInfo->lParam == GetItemData(i)) return i;
		}
		return -1;
	}
	else
	{
		return CListCtrl::FindItem(pFindInfo, nStart);
	}
}
int CListCtrlEx::GetItemIndexFromData(DWORD_PTR dwData)
{
	LVFINDINFO find;
	find.flags = LVFI_PARAM;
	find.lParam = dwData;
	return CListCtrl::FindItem(&find);
}
int CALLBACK CListCtrlEx::CompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	CListCtrlEx *This = reinterpret_cast<CListCtrlEx*>(lParamSort);
	ColumnInfo *pColInfo;
	int nSort = 0;
	int nCompare = 0;

	if (!This) return 0;
	if (!(This->m_nSortColumn < 0 || This->m_nSortColumn >= This->GetColumnCount()))
	{
		pColInfo = This->GetColumnInfo(This->m_nSortColumn);
		if (pColInfo && (pColInfo->m_eSort & SortBits))
		{
			nSort = pColInfo->m_eSort & Ascending ? 1 : -1;
			if (!This->m_fnCompare && pColInfo->m_fnCompare) This->m_fnCompare = pColInfo->m_fnCompare;
		}
	}
	if (This->m_fnCompare && This->m_fnCompare != &CListCtrlEx::CompareProc)
	{
		ItemData *pD1 = reinterpret_cast<ItemData*>(lParam1);
		ItemData *pD2 = reinterpret_cast<ItemData*>(lParam2);
		if (pD1) lParam1 = pD1->m_dwUserData;
		if (pD2) lParam2 = pD2->m_dwUserData;
		nCompare = This->m_fnCompare(lParam1, lParam2, This->m_dwSortData ? This->m_dwSortData : This->m_nSortColumn);
		if (!This->m_dwSortData && nSort)
		{
			return nCompare * nSort;
		}
	}
	if (!nSort) return 0;

	int nLeft = This->GetItemIndexFromData(lParam1);
	int nRight = This->GetItemIndexFromData(lParam2);

	if (nLeft < 0) nLeft = lParam1;
	if (nRight < 0) nRight = lParam2;
	int nCount = This->GetItemCount();
	if (nLeft < 0 || nRight < 0 || nLeft >= nCount || nRight >= nCount) return 0;
	nCompare = Compare(nLeft, nRight, lParamSort);
	return (nCompare * nSort);
}
void CListCtrlEx::OnHdnItemclick(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW*  phdr = reinterpret_cast<NM_LISTVIEW*>(pNMHDR);
	SortOnColumn(phdr->iSubItem, TRUE);
	*pResult = 0;
	//return FALSE;
}
BOOL CListCtrlEx::SortItems(PFNLVCOMPARE pfnCompare, DWORD_PTR dwData)
{
	int nCount = GetItemCount();
	DWORD_PTR dwEditingItemData = 0;
	if (m_nEditingRow >= 0 && m_nEditingRow < nCount)
		dwEditingItemData = GetItemDataInternal(m_nEditingRow);
	CString dbg;
	dbg.Format(L"\nBefore : %d", m_nEditingRow);
	OutputDebugString(dbg);
	m_fnCompare = pfnCompare;
	m_dwSortData = dwData;
	BOOL ret = CListCtrl::SortItems(CompareProc, (DWORD_PTR)this);
	m_fnCompare = NULL;
	m_dwSortData = NULL;
	if (dwEditingItemData)
		m_nEditingRow = GetItemIndexFromData(dwEditingItemData);

	dbg.Format(L"\nAfter : %d", m_nEditingRow);
	OutputDebugString(dbg);
	return ret;
}

BOOL CListCtrlEx::SortOnColumn(int nColumn, BOOL bChangeOrder)
{
	ColumnInfo *pColInfo;
	if ((pColInfo = GetColumnInfo(nColumn)) && (pColInfo->m_eSort & SortBits))
	{
		if (pColInfo->m_eSort & Auto)
		{
			pColInfo->m_eSort = (Sort)((pColInfo->m_eSort & (Ascending | Descending)) ? pColInfo->m_eSort : pColInfo->m_eSort | Descending);
			if (bChangeOrder) pColInfo->m_eSort = (Sort)(pColInfo->m_eSort ^ (Ascending | Descending));
		}
		CHeaderCtrl *pHdr = this->GetHeaderCtrl();
		HDITEM hd;
		hd.mask = HDI_FORMAT;
		if (pHdr)
		{
			pHdr->GetItem(m_nSortColumn, &hd);
			hd.fmt = hd.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
			pHdr->SetItem(m_nSortColumn, &hd);
		}
		m_nSortColumn = nColumn;
		CListCtrl::SortItems(CompareProc, (DWORD_PTR)this);
		if (pHdr)
		{
			pHdr->GetItem(nColumn, &hd);
			hd.fmt = hd.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
			hd.fmt |= pColInfo->m_eSort & Ascending ? HDF_SORTUP : HDF_SORTDOWN;
			pHdr->SetItem(nColumn, &hd);
		}
		return TRUE;
	}
	return FALSE;
}

BOOL CListCtrlEx::OnNMDblclk(NMHDR *pNMHDR, LRESULT *pResult)
{
	NM_LISTVIEW *pNMListView = (NM_LISTVIEW *)pNMHDR;
	if (!pNMListView) return FALSE;

	CLRow = pNMListView->iItem;
	CLColumn = pNMListView->iSubItem;

	*pResult = DisplayEditor(CLRow, CLColumn);

	return *pResult;
}


BOOL CListCtrlEx::EnsureSubItemVisible(int nItem, int nSubItem, CRect *pRect)
{
	BOOL ret = EnsureVisible(nItem, FALSE);
	CRect rect;
	GetSubItemRect(nItem, nSubItem, LVIR_LABEL, rect);
	CRect rtList;
	GetClientRect(&rtList);
	if (rect.right > rtList.Width()) Scroll(CSize(rect.Width() > rtList.Width() ? rect.left : rect.right - rtList.Width(), 0));
	if (rect.left < 0) Scroll(CSize(rect.left));
	if (pRect)
	{
		GetSubItemRect(nItem, nSubItem, LVIR_LABEL, rect);
		rect.right = min(rect.right, rtList.Width() - 4);
		*pRect = rect;
	}
	return ret;
}
BOOL CListCtrlEx::DisplayEditor(int nItem, int nSubItem)
{
	int nCount = GetItemCount();
	DWORD_PTR dwEditingItemData = 0;
	if (nItem >= 0 && nItem < nCount)
		dwEditingItemData = GetItemDataInternal(nItem);
	HideEditor();
	if (dwEditingItemData)
		nItem = GetItemIndexFromData(dwEditingItemData);
	if (nItem < 0 || nItem > nCount || nSubItem < 0 || nSubItem > this->GetColumnCount()
		|| IsColumnReadOnly(nSubItem) || IsRowReadOnly(nItem) || IsCellReadOnly(nItem, nSubItem)) return FALSE;
	CRect rectSubItem;
	//GetSubItemRect(nItem, nSubItem, LVIR_LABEL, rectSubItem);
	EnsureSubItemVisible(nItem, nSubItem, &rectSubItem);

	CellInfo *pCellInfo = GetCellInfo(nItem, nSubItem);
	ColumnInfo *pColInfo = GetColumnInfo(nSubItem);
	ItemData *pRowInfo = (ItemData *)GetItemDataInternal(nItem);

	m_pEditor = &m_eiDefEditor;
	BOOL bReadOnly = FALSE;
	if (pColInfo && !(bReadOnly |= pColInfo->m_eiEditor.m_bReadOnly) && pColInfo->m_eiEditor.IsSet())
	{
		m_pEditor = &pColInfo->m_eiEditor;
	}
	if (pRowInfo && !(bReadOnly |= pRowInfo->m_eiEditor.m_bReadOnly) && pRowInfo->m_eiEditor.IsSet())
	{
		m_pEditor = &pRowInfo->m_eiEditor;
	}
	if (pCellInfo && !(bReadOnly |= pCellInfo->m_eiEditor.m_bReadOnly) && pCellInfo->m_eiEditor.IsSet())
	{
		m_pEditor = &pCellInfo->m_eiEditor;
	}
	if (bReadOnly || !m_pEditor || !m_pEditor->IsSet() || m_pEditor->m_bReadOnly) { m_pEditor = NULL; return  FALSE; }

	m_nEditingRow = nItem;
	m_nEditingColumn = nSubItem;
	m_nRow = nItem;
	m_nColumn = nSubItem;
	CString text = GetItemText(nItem, nSubItem);
	if (m_pEditor->m_pfnInitEditor)
		m_pEditor->m_pfnInitEditor(&m_pEditor->m_pWnd, nItem, nSubItem, text, GetItemData(nItem), this, TRUE);

	if (!m_pEditor->m_pWnd) return  FALSE;
	SelectItem(-1, FALSE);
	if (!m_pEditor->m_pfnInitEditor) m_pEditor->m_pWnd->SetWindowText(text);

	m_pEditor->m_pWnd->SetParent(this);
	m_pEditor->m_pWnd->SetOwner(this);

	ASSERT(m_pEditor->m_pWnd->GetParent() == this);
	m_pEditor->m_pWnd->SetWindowPos(NULL, rectSubItem.left, rectSubItem.top, rectSubItem.Width(), rectSubItem.Height(), SWP_SHOWWINDOW);
	m_pEditor->m_pWnd->ShowWindow(SW_SHOW);
	m_pEditor->m_pWnd->SetFocus();

	m_msgHook.Attach(m_pEditor->m_pWnd->m_hWnd, this->m_hWnd);

	SetFocusText();

	return TRUE;
}

CListCtrlEx::CellInfo* CListCtrlEx::GetCellInfo(int nItem, int nSubItem)
{
	ItemData* pData = (ItemData*)GetItemDataInternal(nItem);
	if (pData == NULL) return NULL;
	INT_PTR nCount = pData->m_aCellInfo.GetCount();
	for (INT_PTR i = 0; i < nCount; i++)
	{
		CellInfo *pInfo = (CellInfo*)pData->m_aCellInfo.GetAt(i);
		if (pInfo && pInfo->m_nColumn == nSubItem) return pInfo;
	}
	return NULL;
}

void CListCtrlEx::HideEditor(BOOL bUpdate)
{
	CSingleLock lock(&m_oLock, TRUE);
	if (lock.IsLocked() && m_msgHook.Detach())
	{
		if (m_pEditor && m_pEditor->m_pWnd)
		{
			m_pEditor->m_pWnd->ShowWindow(SW_HIDE);
			CString text;
			DWORD_PTR dwData = 0;
			if (GetItemCount() > m_nEditingRow)
			{
				text = GetItemText(m_nEditingRow, m_nEditingColumn);
				dwData = GetItemData(m_nEditingRow);
			}
			else
			{
				bUpdate = FALSE;
			}
			if (m_pEditor->m_pfnEndEditor)
			{
				bUpdate = m_pEditor->m_pfnEndEditor(&m_pEditor->m_pWnd, m_nEditingRow, m_nEditingColumn, text, dwData, this, bUpdate);
			}
			else
			{
				m_pEditor->m_pWnd->GetWindowText(text);
			}
			if (bUpdate)
			{
				SetItemText(m_nEditingRow, m_nEditingColumn, text);
			}
			if (GetItemCount() > m_nEditingRow) Update(m_nEditingRow);
			if (bUpdate == -1) SortOnColumn(m_nEditingColumn);
			m_pEditor = NULL;
		}
	}
	lock.Unlock();
}

int CListCtrlEx::GetColumnCount(void)
{
	CHeaderCtrl *pHdr = GetHeaderCtrl();
	if (pHdr) return pHdr->GetItemCount();
	int i = 0;
	LVCOLUMN col;
	col.mask = LVCF_WIDTH;
	while (GetColumn(i++, &col));
	return i;
}

BOOL CListCtrlEx::PreTranslateMessage(MSG* pMsg)
{
	if ((m_hAccel && GetParent() && GetFocus() == this
		&& ::TranslateAccelerator(GetParent()->m_hWnd, m_hAccel, pMsg))) return TRUE;

	CLRow = m_nRow;
	CLColumn = m_nColumn;

	int iMes = (int)pMsg->message;
	int iWPar = (int)pMsg->wParam;
	CWnd* pWndWithFocus = GetFocus();

	if (iMes == WM_KEYDOWN)
	{
		if (pWndWithFocus == GetDlgItem(RNMF_EDIT_VIEW) 
			|| pWndWithFocus == GetDlgItem(AUTORM_EDIT_LIST))
		{
			if (iWPar == VK_RETURN || iWPar == VK_TAB || iWPar == VK_DOWN)
			{
				if (CLRow + 1 == GetItemCount()) DisplayEditor(0, CLColumn);
				else DisplayEditor(CLRow + 1, CLColumn);

				return TRUE;
			}
			else if (iWPar == VK_UP)
			{
				if (CLRow == 0) DisplayEditor(GetItemCount() - 1, CLColumn);
				else DisplayEditor(CLRow - 1, CLColumn);

				return TRUE;
			}
		}
		else if (pWndWithFocus == GetDlgItem(EDPR_EDIT_VIEW))
		{
			if (iWPar == VK_RETURN || iWPar == VK_DOWN)
			{
				if (CLRow + 1 == GetItemCount()) DisplayEditor(0, CLColumn);
				else DisplayEditor(CLRow + 1, CLColumn);

				return TRUE;
			}
			else if (iWPar == VK_TAB)
			{
				// Trường hợp đặc biệt
				// iGFLTitle = 3, iGFLSubject = 4, iGFLTags = 5, iGFLCategories = 6
				// iGFLComments = 7, iGFLAuthors = 8, iGFLCompany = 9
				if (CLColumn >= 3 && CLColumn < 9) DisplayEditor(CLRow, CLColumn + 1);
				else if (CLColumn == 9) DisplayEditor(CLRow, 3);

				return TRUE;
			}
			else if (iWPar == VK_UP)
			{
				if (CLRow == 0) DisplayEditor(GetItemCount() - 1, CLColumn);
				else DisplayEditor(CLRow - 1, CLColumn);

				return TRUE;
			}
		}
	}
	else if (iMes == WM_CHAR)
	{
		TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
		if (pWndWithFocus == GetDlgItem(RNMF_EDIT_VIEW))
		{
			if (chr == 0x5C || chr == 0x2F || chr == 0x3A || chr == 0x2A ||
				chr == 0x3F || chr == 0x22 || chr == 0x3C || chr == 0x3e || chr == 0x7C)
			{
				m_edit_rename.ShowBalloonTip(L"Hướng dẫn",
					L"Đặt tên không chứa ký tự đặc biệt "
					L"( \\ / : * ? \" < > | ) và tối đa 30 ký tự.", TTI_INFO);
				return TRUE;
			}
		}
	}

	return FALSE;
}


BOOL CListCtrlEx::OnNMCustomdraw(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVCUSTOMDRAW lplvcd = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);

	int iRow = lplvcd->nmcd.dwItemSpec;
	int iCol = lplvcd->iSubItem;
	*pResult = 0;
	switch (lplvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:
		*pResult = CDRF_NOTIFYSUBITEMDRAW;          // ask for subitem notifications.
		break;

	case CDDS_ITEMPREPAINT:
		*pResult = CDRF_NOTIFYSUBITEMDRAW;
		break;
	case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
	{
		COLORREF clrBack = 0xFFFFFFFF;
		COLORREF clrText = 0xFFFFFFFF;

		*pResult = CDRF_DODEFAULT;
		CellInfo *pCell = GetCellInfo(iRow, iCol);
		if (pCell)
		{
			clrBack = pCell->m_clrBack;
			clrText = pCell->m_clrText;
		}
		if (clrBack == 0xFFFFFFFF && clrText == 0xFFFFFFFF)
		{
			ItemData* pData = (ItemData*)GetItemDataInternal(iRow);
			if (pData)
			{
				clrBack = pData->m_clrBack;
				clrText = pData->m_clrText;
			}
		}
		if (clrBack == 0xFFFFFFFF && clrText == 0xFFFFFFFF)
		{
			ColumnInfo *pInfo = GetColumnInfo(iCol);
			if (pInfo)
			{
				clrBack = pInfo->m_clrBack;
				clrText = pInfo->m_clrText;
			}
		}
		if (clrBack != 0xFFFFFFFF)
		{
			lplvcd->clrTextBk = clrBack;
			*pResult = CDRF_NEWFONT;
		}
		else
		{
			if (clrBack != m_clrDefBack)
			{
				lplvcd->clrTextBk = m_clrDefBack;
				*pResult = CDRF_NEWFONT;
			}
		}
		if (clrText != 0xFFFFFFFF)
		{
			lplvcd->clrText = clrText;
			*pResult = CDRF_NEWFONT;
		}
		else
		{
			if (clrText != m_clrDefText)
			{
				lplvcd->clrText = m_clrDefText;
				*pResult = CDRF_NEWFONT;
			}
		}
	}
	break;

	default:
		*pResult = CDRF_DODEFAULT;
	}
	if (*pResult == 0 || *pResult == CDRF_DODEFAULT)
		return FALSE;
	else
		return TRUE;
}

void CListCtrlEx::SetRowColors(int nItem, COLORREF clrBk, COLORREF clrText)
{
	ItemData* pData = (ItemData*)GetItemDataInternal(nItem);
	if (!pData) SetItemData(nItem, 0);
	pData = (ItemData*)GetItemDataInternal(nItem);
	if (!pData) return;

	pData->m_clrText = clrText;
	pData->m_clrBack = clrBk;
	Update(nItem);
}
BOOL CListCtrlEx::AddItem(int ItemIndex, int SubItemIndex, LPCTSTR ItemText, int ImageIndex)
{
	LV_ITEM lvItem;
	lvItem.mask = LVIF_TEXT;
	lvItem.iItem = ItemIndex;
	lvItem.iSubItem = SubItemIndex;
	lvItem.pszText = (LPTSTR)ItemText;
	if (ImageIndex != -1) {
		lvItem.mask |= LVIF_IMAGE;
		lvItem.iImage |= LVIF_IMAGE;
	}
	if (SubItemIndex == 0)
		return InsertItem(&lvItem);
	return SetItem(&lvItem);
}
LRESULT CListCtrlEx::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == UM_HIDEEDITOR)
	{
		HideEditor((BOOL)wParam);
	}
	return CListCtrl::WindowProc(message, wParam, lParam);
}

void CListCtrlEx::SetColumnColors(int nColumn, COLORREF clrBack, COLORREF clrText)
{
	if (nColumn < 0 || nColumn >= GetColumnCount()) return;

	ColumnInfo* pColInfo = GetColumnInfo(nColumn);
	if (!pColInfo)
	{
		pColInfo = new ColumnInfo(nColumn);
		m_aColumnInfo.Add(pColInfo);
	}

	pColInfo->m_clrBack = clrBack;
	pColInfo->m_clrText = clrText;

}
void CListCtrlEx::SetCellColors(int nRow, int nColumn, COLORREF clrBack, COLORREF clrText)
{
	if (nRow < 0 || nRow >= GetItemCount() || nColumn < 0 || nColumn >= GetColumnCount()) return;
	CellInfo* pCellInfo = GetCellInfo(nRow, nColumn);
	if (!pCellInfo)
	{
		SetCellData(nRow, nColumn, 0);
	}
	pCellInfo = GetCellInfo(nRow, nColumn);
	ASSERT(pCellInfo);
	pCellInfo->m_clrBack = clrBack;
	pCellInfo->m_clrText = clrText;

}
void CListCtrlEx::PreSubclassWindow()
{
	m_clrDefBack = GetTextBkColor();
	m_clrDefText = GetTextColor();
	CListCtrl::PreSubclassWindow();
	SetExtendedStyle(LVS_EX_FULLROWSELECT);
	ModifyStyle(0, LVS_REPORT);
}

BOOL CListCtrlEx::SetCellData(int nItem, int nSubItem, DWORD_PTR dwData)
{
	if (nItem < 0 || nItem >= GetItemCount() || nSubItem < 0 || nSubItem >= GetColumnCount()) return FALSE;
	CellInfo* pCellInfo = GetCellInfo(nItem, nSubItem);
	if (!pCellInfo)
	{
		pCellInfo = new CellInfo(nSubItem);
		ItemData *pData = (ItemData*)GetItemDataInternal(nItem);
		if (!pData)
		{
			SetItemData(nItem, 0);
			pData = (ItemData*)GetItemDataInternal(nItem);
		}
		pData->m_aCellInfo.Add(pCellInfo);
	}
	pCellInfo->m_dwUserData = dwData;

	return TRUE;
}

DWORD_PTR CListCtrlEx::GetCellData(int nItem, int nSubItem)
{
	CellInfo* pCellInfo = GetCellInfo(nItem, nSubItem);
	if (pCellInfo) return pCellInfo->m_dwUserData;
	else return 0;
}

void CListCtrlEx::SetCellEditor(int nRow, int nColumn, PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd* pWnd)
{
	if (nRow < 0 || nRow >= GetItemCount() || nColumn < 0 || nColumn >= GetColumnCount()) return;
	CellInfo* pCellInfo = GetCellInfo(nRow, nColumn);
	if (!pCellInfo)
	{
		SetCellData(nRow, nColumn, 0);
		pCellInfo = (CellInfo*)GetCellInfo(nRow, nColumn);
		ASSERT(pCellInfo);
	}

	pCellInfo->m_eiEditor.m_pfnInitEditor = pfnInitEditor;
	pCellInfo->m_eiEditor.m_pfnEndEditor = pfnEndEditor;
	pCellInfo->m_eiEditor.m_pWnd = pWnd;

}

void CListCtrlEx::SetCellEditor(int nRow, int nColumn, CWnd* pWnd)
{
	SetCellEditor(nRow, nColumn, NULL, NULL, pWnd);
}


void CListCtrlEx::SetDefaultEditor(PFNEDITORCALLBACK pfnInitEditor, PFNEDITORCALLBACK pfnEndEditor, CWnd* pWnd)
{
	m_eiDefEditor.m_pfnInitEditor = pfnInitEditor;
	m_eiDefEditor.m_pfnEndEditor = pfnEndEditor;
	m_eiDefEditor.m_pWnd = pWnd;
}

void CListCtrlEx::SetDefaultEditor(CWnd* pWnd)
{
	SetDefaultEditor(NULL, NULL, pWnd);
}

void CListCtrlEx::SetColumnReadOnly(int nColumn, bool bReadOnly)
{
	if (nColumn < 0 || nColumn >= GetColumnCount()) return;
	ColumnInfo *pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	if (!pInfo) SetColumnEditor(nColumn, (CWnd*)NULL);
	pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	ASSERT(pInfo);
	pInfo->m_eiEditor.m_bReadOnly = bReadOnly;
}

void CListCtrlEx::SetColumnSorting(int nColumn, Sort eSort, Comparer eComparer)
{
	if (nColumn < 0 || nColumn >= GetColumnCount() || !(eSort & SortBits)) return;
	ColumnInfo *pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	if (!pInfo) SetColumnEditor(nColumn, (CWnd*)NULL);
	pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	ASSERT(pInfo);
	pInfo->m_eSort = eSort;
	pInfo->m_eCompare = eComparer;
	pInfo->m_fnCompare = NULL;
}
void CListCtrlEx::SetColumnSorting(int nColumn, Sort eSort, PFNLVCOMPARE fnCallBack)
{
	if (nColumn < 0 || nColumn >= GetColumnCount() || !(eSort & SortBits)) return;
	ColumnInfo *pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	if (!pInfo) SetColumnEditor(nColumn, (CWnd*)NULL);
	pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	ASSERT(pInfo);
	pInfo->m_eSort = eSort;
	pInfo->m_eCompare = NotSet;
	pInfo->m_fnCompare = fnCallBack;
}
void CListCtrlEx::SetCellReadOnly(int nRow, int nColumn, bool bReadOnly)
{
	if (nRow < 0 || nRow >= GetItemCount() || nColumn < 0 || nColumn >= GetColumnCount()) return;

	CellInfo *pInfo = (CellInfo*)GetCellInfo(nRow, nColumn);
	if (!pInfo) SetCellEditor(nRow, nColumn, (CWnd*)NULL);
	pInfo = (CellInfo*)GetCellInfo(nRow, nColumn);
	ASSERT(pInfo);
	pInfo->m_eiEditor.m_bReadOnly = bReadOnly;
}
void CListCtrlEx::SetRowReadOnly(int nRow, bool bReadOnly)
{
	if (nRow < 0 || nRow >= GetItemCount()) return;

	ItemData *pInfo = (ItemData*)GetItemDataInternal(nRow);
	if (!pInfo) SetItemData(nRow, 0);
	pInfo = (ItemData*)GetItemDataInternal(nRow);
	ASSERT(pInfo);
	pInfo->m_eiEditor.m_bReadOnly = bReadOnly;
}

BOOL CListCtrlEx::IsColumnReadOnly(int nColumn)
{
	if (nColumn < 0 || nColumn >= GetColumnCount()) return FALSE;
	ColumnInfo *pInfo = (ColumnInfo*)GetColumnInfo(nColumn);
	if (pInfo)
	{
		return pInfo->m_eiEditor.m_bReadOnly;
	}
	return FALSE;
}

BOOL CListCtrlEx::IsRowReadOnly(int nRow)
{
	if (nRow < 0 || nRow >= GetItemCount()) return FALSE;
	ItemData *pInfo = (ItemData*)GetItemDataInternal(nRow);
	if (pInfo)
	{
		return pInfo->m_eiEditor.m_bReadOnly;
	}
	return FALSE;
}

BOOL CListCtrlEx::IsCellReadOnly(int nRow, int nColumn)
{
	if (nRow < 0 || nRow >= GetItemCount() || nColumn < 0 || nColumn >= GetColumnCount()) return FALSE;

	CellInfo *pInfo = (CellInfo*)GetCellInfo(nRow, nColumn);
	if (pInfo)
	{
		return pInfo->m_eiEditor.m_bReadOnly;
	}
	else
		return (IsRowReadOnly(nRow) || IsColumnReadOnly(nColumn));
}

void CListCtrlEx::DeleteSelectedItems(void)
{
	int nItem = -1;
	while ((nItem = GetNextItem(-1, LVNI_SELECTED)) >= 0)
	{
		DeleteItem(nItem);
	}
}
void CListCtrlEx::HandleDeleteKey(BOOL bHandle)
{
	m_bHandleDelete = bHandle;
}
void CListCtrlEx::SelectItem(int nItem, BOOL bSelect)
{
	if (nItem < 0)
	{
		int nIndex = -1;
		while ((nIndex = this->GetNextItem(nIndex, bSelect ? LVNI_ALL : LVNI_SELECTED)) >= 0)
		{
			this->SetItemState(nIndex, (bSelect ? LVIS_SELECTED : 0), LVIS_SELECTED);
		}
	}
	else
	{
		this->SetItemState(nItem, (bSelect ? LVIS_SELECTED : 0), LVIS_SELECTED);
	}
}

BOOL CListCtrlEx::DeleteAllColumns(void)
{
	while (DeleteColumn(0));
	return (GetColumnCount() == 0);
}

BOOL CListCtrlEx::Reset(void)
{
	return (DeleteAllItems() &&
		DeleteAllColumns());
}

int CListCtrlEx::Compare(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	CListCtrlEx *This = (CListCtrlEx *)lParamSort;
	if (!This || This->m_nSortColumn < 0 || This->m_nSortColumn >= This->GetColumnCount()) return 0;
	int nSubItem = This->m_nSortColumn;
	ColumnInfo *pInfo = This->GetColumnInfo(nSubItem);
	if (!pInfo) return 0;

	CString AL = This->GetItemText(lParam1, nSubItem);
	CString AR = This->GetItemText(lParam2, nSubItem);

	LPCSTR strLeft = (LPCSTR)(LPCTSTR)AL;
	LPCSTR strRight = (LPCSTR)(LPCTSTR)AR;

	switch (pInfo->m_eCompare)
	{
	case Int:
		return CompareInt(strLeft, strRight);
	case Double:
		return CompareDouble(strLeft, strRight);
	case StringNoCase:
		return CompareStringNoCase(strLeft, strRight);
	case StringNumber:
		return CompareNumberString(strLeft, strRight);
	case StringNumberNoCase:
		return CompareNumberStringNoCase(strLeft, strRight);
	case String:
		return CompareString(strLeft, strRight);
	case Date:
		return CompareDate(strLeft, strRight);
	case NotSet:
		return 0;
	default:
		return CompareString(strLeft, strRight);
	}
	return CompareString(strLeft, strRight);
}
int CListCtrlEx::CompareInt(LPCSTR pLeftText, LPCSTR pRightText)
{
	return (int)(atol(pLeftText) - atol(pRightText));
}
int CListCtrlEx::CompareDouble(LPCSTR pLeftText, LPCSTR pRightText)
{
	return (int)(atof(pLeftText) - atof(pRightText));
}
int CListCtrlEx::CompareString(LPCSTR pLeftText, LPCSTR pRightText)
{
	CString CRight = (LPCTSTR)pRightText;
	return CString(pLeftText).Compare(CRight);
}
int CListCtrlEx::CompareNumberString(LPCSTR pLeftText, LPCSTR pRightText)
{
	LONGLONG l1 = atol(pLeftText);
	LONGLONG l2 = atol(pRightText);
	if (l1 && l2 && (l1 - l2))
	{
		CString str1, str2;
		str1.Format(L"%ld", l1);
		str2.Format(L"%ld", l2);
		CString left(pLeftText);
		CString right(pRightText);
		if (str1.GetLength() == left.GetLength() && str2.GetLength() == right.GetLength()) return (int)(l1 - l2);
	}
	CString CRight = (LPCTSTR)pRightText;
	return CString(pLeftText).Compare(CRight);
}
int CListCtrlEx::CompareNumberStringNoCase(LPCSTR pLeftText, LPCSTR pRightText)
{
	LONGLONG l1 = atol(pLeftText);
	LONGLONG l2 = atol(pRightText);
	if (l1 && l2 && (l1 - l2))
	{
		CString str1, str2;
		str1.Format(L"%ld", l1);
		str2.Format(L"%ld", l2);
		CString left(pLeftText);
		CString right(pRightText);
		if (str1.GetLength() == left.GetLength() && str2.GetLength() == right.GetLength()) return (int)(l1 - l2);
	}
	CString CRight = (LPCTSTR)pRightText;
	return CString(pLeftText).CompareNoCase(CRight);
}
int CListCtrlEx::CompareStringNoCase(LPCSTR pLeftText, LPCSTR pRightText)
{
	CString CRight = (LPCTSTR)pRightText;
	return CString(pLeftText).CompareNoCase(CRight);
}
int CListCtrlEx::CompareDate(LPCSTR pLeftText, LPCSTR pRightText)
{
	COleDateTime dL, dR;
	CString CLeft = (LPCTSTR)pLeftText;
	CString CRight = (LPCTSTR)pRightText;
	dL.ParseDateTime(CLeft);
	dR.ParseDateTime(CRight);
	return (dL == dR ? 0 : (dL < dR ? -1 : 1));
}


BOOL CListCtrlEx::PreCreateWindow(CREATESTRUCT& cs)
{
	cs.dwExStyle |= LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;

	return CListCtrl::PreCreateWindow(cs);
}

/*************************************************************************************/
// Leo add 08.07.2019

void CListCtrlEx::OnMouseMove(UINT nFlags, CPoint point)
{
	LVHITTESTINFO NowHti;
	NowHti.pt = point;
	int nItem = SubItemHitTest(&NowHti);
	int nSubitem = NowHti.iSubItem;

	if (nItem >= 0)
	{
		CString szval = CListCtrl::GetItemText(nItem, nSubitem);
		if (nSubitem == jColumnHand && szval != L"") ::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_HAND));
		else ::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_ARROW));
	}

	if (m_pDragImage)
	{
		CPoint ptDragImage(point);
		ClientToScreen(&ptDragImage);
		m_pDragImage->DragMove(ptDragImage);
		m_pDragImage->DragLeave(CWnd::GetDesktopWindow());

		static const int nXOffset = 8;
		CRect rect;
		GetWindowRect(rect);
		CWnd* pDropWnd = CWnd::WindowFromPoint(CPoint(rect.left + nXOffset, ptDragImage.y));

		if (pDropWnd == this)
		{
			point.x = nXOffset;
			UpdateSelection(HitTest(point));
		}

		CRect rectClient;
		GetClientRect(rectClient);
		CPoint ptClientDragImage(ptDragImage);
		ScreenToClient(&ptClientDragImage);

		CHeaderCtrl* pHeader = (CHeaderCtrl*)GetDlgItem(0);
		if (pHeader)
		{
			CRect rectHeader;
			pHeader->GetClientRect(rectHeader);
			rectClient.top += rectHeader.Height();
		}

		if (ptClientDragImage.y < rectClient.top)
		{
			SetScrollTimer(scrollUp);
		}
		else if (ptClientDragImage.y > rectClient.bottom)
		{
			SetScrollTimer(scrollDown);
		}
		else
		{
			KillScrollTimer();
		}

		m_pDragImage->DragEnter(CWnd::GetDesktopWindow(), ptDragImage);
	}
	else
	{
		KillScrollTimer();
	}

	CListCtrl::OnMouseMove(nFlags, point);
}

void CListCtrlEx::OnLButtonUp(UINT nFlags, CPoint point)
{
	if (m_pDragImage)
	{
		KillScrollTimer();

		::ReleaseCapture();
		m_pDragImage->DragLeave(CWnd::GetDesktopWindow());
		m_pDragImage->EndDrag();

		delete m_pDragImage;
		m_pDragImage = NULL;

		DropItem();
	}

	blDragging = true;
	CListCtrl::OnLButtonUp(nFlags, point);
}

void CListCtrlEx::OnBeginDrag(NMHDR* pNMHDR, LRESULT* pResult)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	if (pNMListView)
	{
		m_nPrevDropIndex = -1;
		m_uPrevDropState = NULL;

		m_anDragIndexes.RemoveAll();
		POSITION pos = GetFirstSelectedItemPosition();
		while (pos)
		{
			m_anDragIndexes.Add(GetNextSelectedItem(pos));
		}

		DWORD dwStyle = GetStyle();
		if ((dwStyle & LVS_SINGLESEL) == LVS_SINGLESEL)
		{
			m_dwStyle = dwStyle;
			ModifyStyle(LVS_SINGLESEL, NULL);
		}

		if (m_anDragIndexes.GetSize() > 0)
		{
			delete m_pDragImage;
			CPoint ptDragItem;
			m_pDragImage = CreateDragImageEx(&ptDragItem);
			if (m_pDragImage)
			{
				m_pDragImage->BeginDrag(0, ptDragItem);
				m_pDragImage->DragEnter(CWnd::GetDesktopWindow(), pNMListView->ptAction);

				SetCapture();
			}
		}
	}

	*pResult = 0;
}

// Biên dịch 64b --> đổi OnTimer(UINT... = UINT_PTR)
#if defined _M_X64
void CListCtrlEx::OnTimer(UINT_PTR nIDEvent)
#else
void CListCtrlEx::OnTimer(UINT nIDEvent)
#endif
{
	if (nIDEvent == CONST_TIMER_CLISTEX && m_pDragImage)
	{
		WPARAM wParam = NULL;
		int nDropIndex = -1;
		if (m_ScrollDirection == scrollUp)
		{
			wParam = MAKEWPARAM(SB_LINEUP, 0);
			nDropIndex = m_nDropIndex - 1;
		}
		else if (m_ScrollDirection == scrollDown)
		{
			wParam = MAKEWPARAM(SB_LINEDOWN, 0);
			nDropIndex = m_nDropIndex + 1;
		}
		m_pDragImage->DragShowNolock(FALSE);
		SendMessage(WM_VSCROLL, wParam, NULL);
		UpdateSelection(nDropIndex);
		m_pDragImage->DragShowNolock(TRUE);
	}
	else
	{
		CListCtrl::OnTimer(nIDEvent);
	}
}

BOOL CListCtrlEx::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	HD_NOTIFY *pHDN = (HD_NOTIFY*)lParam;
	LPNMHDR pNH = (LPNMHDR)lParam;

	if (wParam == 0 && pNH->code == NM_RCLICK)
	{
		CPoint pt(GetMessagePos());
		CHeaderCtrl* pHeader = (CHeaderCtrl*)GetDlgItem(0);
		pHeader->ScreenToClient(&pt);

		CHeaderRightClick cr;
		cr.m_column = -1;
		cr.m_mousePos = pt;

		CRect rcCol;
		for (int i = 0; Header_GetItemRect(pHeader->m_hWnd, i, &rcCol); i++)
		{
			if (rcCol.PtInRect(pt))
			{
				cr.m_column = i;
				break;
			}
		}

		ClientToScreen(&cr.m_mousePos);

		::SendMessage(GetParent()->GetSafeHwnd(),
			WM_LISTCTRLEX_HEADERRIGHTCLICK, (WPARAM)GetSafeHwnd(), (LPARAM)&cr);
	}

	return CListCtrl::OnNotify(wParam, lParam, pResult);
}


void CListCtrlEx::OnDestroy()
{
	DeleteAllItems();
	m_imgList.DeleteImageList();
	m_headerImgList.DeleteImageList();
	CListCtrl::OnDestroy();
}


void CListCtrlEx::SetFocusText()
{
	CWnd* pWndWithFocus = GetFocus();

	if (pWndWithFocus == GetDlgItem(RNMF_EDIT_VIEW)) m_edit_rename.SetSel(0, -1);
	else if (pWndWithFocus == GetDlgItem(EDPR_EDIT_VIEW)) m_edit_properties.SetSel(0, -1);
	else if (pWndWithFocus == GetDlgItem(AUTORM_EDIT_LIST)) m_edit_autorename.SetSel(0, -1);
}


BOOL CListCtrlEx::SetHeaderImage(int nColumn, int nImageIndex, BOOL bLeftSide)
{
	if (GetHeaderCtrl()->GetImageList() == NULL)
		CListCtrl::GetHeaderCtrl()->SetImageList(GetImageList());

	HDITEM hi;
	::memset(&hi, 0, sizeof(HDITEM));
	hi.mask = HDI_FORMAT;
	if (!GetHeaderCtrl()->GetItem(nColumn, &hi))
		return FALSE;

	hi.mask |= HDI_IMAGE;
	hi.fmt |= HDF_IMAGE;

	if (!bLeftSide)	hi.fmt |= HDF_BITMAP_ON_RIGHT;

	hi.iImage = nImageIndex;
	return CListCtrl::GetHeaderCtrl()->SetItem(nColumn, &hi);
}

int CListCtrlEx::GetHeaderImage(int nColumn) const
{
	HDITEM hi;
	::memset(&hi, 0, sizeof(HDITEM));
	hi.mask = HDI_IMAGE;
	return !GetHeaderCtrl()->GetItem(nColumn, &hi) ? hi.iImage : -1;
}

CImageList* CListCtrlEx::SetHeaderImageList(CImageList *pImageList)
{
	return CListCtrl::GetHeaderCtrl()->SetImageList(pImageList);
}

CImageList* CListCtrlEx::SetHeaderImageList(UINT nBitmapID, COLORREF crMask)
{
	m_headerImgList.Create(nBitmapID, 16, 4, crMask);
	return SetHeaderImageList(&m_headerImgList);
}

BOOL CListCtrlEx::SetItemImage(int nItem, int nSubItem, int nImageIndex)
{
	return CListCtrl::SetItem(nItem, nSubItem, LVIF_IMAGE, NULL, nImageIndex, 0, 0, 0);
}

int CListCtrlEx::GetItemImage(int nItem, int nSubItem) const
{
	LVITEM lvi;
	lvi.iItem = nItem;
	lvi.iSubItem = nSubItem;
	lvi.mask = LVIF_IMAGE;
	return CListCtrl::GetItem(&lvi) ? lvi.iImage : -1;
}

CImageList* CListCtrlEx::SetImageList(CImageList *pImageList)
{
	return CListCtrl::SetImageList(pImageList, LVSIL_SMALL);
}

CImageList* CListCtrlEx::GetImageList() const
{
	return CListCtrl::GetImageList(LVSIL_SMALL);
}

CImageList* CListCtrlEx::SetImageList(UINT nBitmapID, COLORREF crMask)
{
	m_imgList.DeleteImageList();
	ASSERT(m_imgList.GetSafeHandle() == NULL);

	m_imgList.Create(16, 16, ILC_COLOR32, 0, 0);

	CBitmap bm;
	bm.LoadBitmap(nBitmapID);
	m_imgList.Add(&bm, crMask);

	return CListCtrl::SetImageList(&m_imgList, LVSIL_SMALL);
}

void CListCtrlEx::SetScrollTimer(EScrollDirection ScrollDirection)
{
	if (m_uScrollTimer == 0)
	{
		m_uScrollTimer = SetTimer(CONST_TIMER_CLISTEX, 100, NULL);
	}
	m_ScrollDirection = ScrollDirection;
}

void CListCtrlEx::KillScrollTimer()
{
	if (m_uScrollTimer != 0)
	{
		KillTimer(CONST_TIMER_CLISTEX);
		m_uScrollTimer = 0;
		m_ScrollDirection = scrollNull;
	}
}

void CListCtrlEx::UpdateSelection(int nDropIndex)
{
	if (nDropIndex > -1 && nDropIndex < GetItemCount())
	{
		RestorePrevDropItemState();

		m_nPrevDropIndex = nDropIndex;
		m_uPrevDropState = GetItemState(nDropIndex, LVIS_SELECTED);

		SetItemState(nDropIndex, LVIS_SELECTED, LVIS_SELECTED);
		m_nDropIndex = nDropIndex;

		UpdateWindow();
	}
}

void CListCtrlEx::RestorePrevDropItemState()
{
	if (m_nPrevDropIndex > -1 && m_nPrevDropIndex < GetItemCount())
	{
		SetItemState(m_nPrevDropIndex, m_uPrevDropState, LVIS_SELECTED);
	}
}

void CListCtrlEx::DropItem()
{
	RestorePrevDropItemState();

	m_nDropIndex++;
	if (m_nDropIndex < 0 || m_nDropIndex > GetItemCount() - 1)
	{
		m_nDropIndex = GetItemCount();
	}

	int nColumns = 1;
	CHeaderCtrl* pHeader = (CHeaderCtrl*)GetDlgItem(0);
	if (pHeader)
	{
		nColumns = pHeader->GetItemCount();
	}

	for (int nDragItem = 0; nDragItem < m_anDragIndexes.GetSize(); nDragItem++)
	{
		int nDragIndex = m_anDragIndexes[nDragItem];

		if (nDragIndex > -1 && nDragIndex < GetItemCount())
		{
			wchar_t szText[256];
			LV_ITEM lvItem;
			ZeroMemory(&lvItem, sizeof(LV_ITEM));
			lvItem.iItem = nDragIndex;
			lvItem.mask = LVIF_TEXT | LVIF_IMAGE | LVIF_STATE | LVIF_PARAM;
			lvItem.stateMask = LVIS_DROPHILITED | LVIS_FOCUSED | LVIS_SELECTED | LVIS_STATEIMAGEMASK;
			lvItem.pszText = szText;
			lvItem.cchTextMax = sizeof(szText) - 1;
			GetItem(&lvItem);
			BOOL bChecked = GetCheck(nDragIndex);

			SetItemState(nDragIndex, static_cast<UINT>(~LVIS_SELECTED), LVIS_SELECTED);

			lvItem.iItem = m_nDropIndex;
			InsertItem(&lvItem);
			if (bChecked) SetCheck(m_nDropIndex);

			if (nDragIndex > m_nDropIndex) nDragIndex++;

			lvItem.mask = LVIF_TEXT;
			lvItem.iItem = m_nDropIndex;

			for (int nColumn = 1; nColumn < nColumns; nColumn++)
			{
				_tcscpy(lvItem.pszText, (LPCTSTR)(GetItemText(nDragIndex, nColumn)));
				lvItem.iSubItem = nColumn;
				SetItem(&lvItem);
			}

			DeleteItem(nDragIndex);

			for (int nNewDragItem = nDragItem; nNewDragItem < m_anDragIndexes.GetSize(); nNewDragItem++)
			{
				int nNewDragIndex = m_anDragIndexes[nNewDragItem];

				if (nDragIndex < nNewDragIndex && nNewDragIndex < m_nDropIndex)
				{
					m_anDragIndexes[nNewDragItem] = max(nNewDragIndex - 1, 0);
				}
				else if (nDragIndex > nNewDragIndex && nNewDragIndex > m_nDropIndex)
				{
					m_anDragIndexes[nNewDragItem] = nNewDragIndex + 1;
				}
			}
			if (nDragIndex > m_nDropIndex)
			{
				// Item has been added before the drop target, so drop target moves down the list.
				m_nDropIndex++;
			}
		}
	}

	if (m_dwStyle != NULL)
	{
		// Style was modified, so return it back to original style.
		ModifyStyle(NULL, m_dwStyle);
		m_dwStyle = NULL;
	}
}

CImageList* CListCtrlEx::CreateDragImageEx(LPPOINT lpPoint)
{
	CRect rectSingle;
	CRect rectComplete(0, 0, 0, 0);
	int	nIndex = -1;
	BOOL bFirst = TRUE;

	POSITION pos = GetFirstSelectedItemPosition();
	while (pos)
	{
		nIndex = GetNextSelectedItem(pos);
		GetItemRect(nIndex, rectSingle, LVIR_BOUNDS);
		if (bFirst)
		{
			GetItemRect(nIndex, rectComplete, LVIR_BOUNDS);
			bFirst = FALSE;
		}
		rectComplete.UnionRect(rectComplete, rectSingle);
	}

	CClientDC dcClient(this);
	CDC dcMem;
	CBitmap Bitmap;

	if (!dcMem.CreateCompatibleDC(&dcClient))
	{
		return NULL;
	}

	if (!Bitmap.CreateCompatibleBitmap(&dcClient, rectComplete.Width(), rectComplete.Height()))
	{
		return NULL;
	}

	CBitmap* pOldMemDCBitmap = dcMem.SelectObject(&Bitmap);
	dcMem.FillSolidRect(0, 0, rectComplete.Width(), rectComplete.Height(), RGB(0, 255, 0));

	CImageList* pSingleImageList = NULL;
	CPoint pt;

	pos = GetFirstSelectedItemPosition();
	while (pos)
	{
		nIndex = GetNextSelectedItem(pos);
		GetItemRect(nIndex, rectSingle, LVIR_BOUNDS);

		pSingleImageList = CreateDragImage(nIndex, &pt);
		if (pSingleImageList)
		{
			IMAGEINFO ImageInfo;
			pSingleImageList->GetImageInfo(0, &ImageInfo);
			rectSingle.right = rectSingle.left + (ImageInfo.rcImage.right - ImageInfo.rcImage.left);

			pSingleImageList->DrawIndirect(&dcMem, 0,
				CPoint(rectSingle.left - rectComplete.left,
					rectSingle.top - rectComplete.top), rectSingle.Size(),
				CPoint(0, 0));

			delete pSingleImageList;
		}
	}

	dcMem.SelectObject(pOldMemDCBitmap);

	CImageList* pCompleteImageList = new CImageList;
	pCompleteImageList->Create(rectComplete.Width(), rectComplete.Height(), ILC_COLOR | ILC_MASK, 0, 1);
	pCompleteImageList->Add(&Bitmap, RGB(0, 255, 0));

	Bitmap.DeleteObject();

	if (lpPoint)
	{
		CPoint ptCursor;
		GetCursorPos(&ptCursor);
		ScreenToClient(&ptCursor);
		lpPoint->x = ptCursor.x - rectComplete.left;
		lpPoint->y = ptCursor.y - rectComplete.top;
	}

	return pCompleteImageList;
}
