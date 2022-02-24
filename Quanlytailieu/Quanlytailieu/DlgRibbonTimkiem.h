#pragma once
#include "StdAfx.h"

class DlgRibbonTimkiem : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgRibbonTimkiem)

public:
	DlgRibbonTimkiem(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgRibbonTimkiem();

	// Dialog Data
	enum { IDD = DLG_RIBBON_TIMKIEM };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);

public:
	CColorEdit m_txt_search;
	CXPStyleButtonST m_btn_cancel, m_btn_goto, m_btn_search;
	CThemeHelperST m_ThemeHelper;
	CListCtrlEx m_list_view;
	CButton m_check_col, m_check_wb;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	_bstr_t bsConfig;
	_WorksheetPtr psConfig;
	RangePtr prConfig;

	vector<_WorksheetPtr> wpSheet;

	int iCheckCol, iCheckWbook;
	vector<CString> vnameSearch;
	CString szTukhoa;	

	int colSheet, colCell, colNoidung;
	int iwND, iwCEL, iwSHE;
	int iDlgW, iDlgH;

	struct MyStrList
	{
		CString szNoidung;
		CString szCell;

		CString szSheet;
		COLORREF bkgTabColor;

		int iColumn;
		int iRow;

		int iEnable;
	};

	vector<MyStrList> vecFilter;

	int tslkq, lanshow;
	int iStopload, ibuocnhay;

	int nItem, nSubItem;

	MenuIcon mnIcon;

public:
	void _Timkiemdulieu();					// <-- Hàm quan trọng tìm kiếm dữ liệu từ Editbox Ribbon
	
	void _SetDefault();
	void _DeleteAllList();
	void _XacdinhSheetLienquan();

	long _GetCellItems(_WorksheetPtr pSheet);

	void _GotoCell(int nIndex = 0);
	void _AddItemsListKetqua(int iBegin, int iEnd);
	void _SelectAllItems();

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();
	void _ChangeSearch();

	int _FilterData();
	int _LoadListView();
	int _GetNameSearch();
	int _GetAllSelectedItems(vector<int> &vecRow);

	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnSearch();
	afx_msg void OnBnClickedBtnGotocell();
	afx_msg void OnBnClickedCheckCol();
	afx_msg void OnBnClickedCheckWb();

	afx_msg void OnRbsearchGoto();
	afx_msg void OnRbsearchOpenfile();

	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);	
	
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);

};
