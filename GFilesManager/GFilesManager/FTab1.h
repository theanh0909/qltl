#pragma once
#include "pch.h"

typedef struct {
	LVITEM* plvi;
	CString sCol2;
} lvItem, *plvItem;

class FTab1 : public CDialogEx
{
	DECLARE_DYNAMIC(FTab1)

public:
	FTab1(CWnd* pParent = NULL);   // standard constructor
	virtual ~FTab1();

	// Dialog Data
	enum { IDD = FV_TAB1 };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg void OnTimer(UINT_PTR nIDEvent);

public:
	CXPStyleButtonST m_btn_add;
	CThemeHelperST m_ThemeHelper;

	CColorEdit m_txt_search;
	CListCtrlEx m_list_view;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	int iGFLNum, iGFLTen, iGFLPath;
	int iwColNum, iwColTen, iwColPath;

	int iStopload, lanshow, ibuocnhay;
	int tongkq, tongfilter;

	int nItem, nSubItem;

	CString pathFolderDir;
	CString szCheckAddFiles;

	vector<MyStrList> vecItem;
	vector<MyStrList> vecFilter;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	MenuIcon mnIcon;

public:
	void _SetDefault();

	int _LocDulieuNangcao();
	int _GetAllSelectedItems(vector<int> &vecRow);
	bool _IsFilterData();

	void _SaveAllVectorItems();
	void _XacdinhLaidulieu();
	void _Timkiemdulieu();	
	void _ClearAllItems();
	void _ClearAllImages();
	void _LoadListImages();
	void _SelectAllItems();
	void _LoadListMyFiles();
	void _AddItemsListKetqua(int iBegin, int iEnd);	

	afx_msg void OnOpenfile();
	afx_msg void OnOpenfolder();
	afx_msg void OnAddfiles();
	afx_msg void OnDeletefiles();	

	afx_msg void OnMoveup();
	afx_msg void OnMovedown();
	afx_msg void OnReverse();
	afx_msg void OnBgkcolor();

	afx_msg void OnBnClickedBtnAdd();
	afx_msg void OnBnClickedBtnSearch();	

	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);
};
