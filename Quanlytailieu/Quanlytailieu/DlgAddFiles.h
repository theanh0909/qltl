#pragma once
#include "StdAfx.h"

class DlgAddFiles : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgAddFiles)

public:
	DlgAddFiles(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgAddFiles();

	// Dialog Data
	enum { IDD = DLG_ADD_DATA };

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
	CXPStyleButtonST m_btn_cancel, m_btn_search;
	CThemeHelperST m_ThemeHelper;

	CColorEdit m_edtTargetDirectory, m_txt_pathMove, m_lbl_total;
	CColorEdit m_edtWildCard, m_editSearchFiles, m_edtTotalFiles;
	CButton m_chkLookInDir, m_chkLookInSubdirectories, m_check_all;
	CButton m_check_tree_folder, m_check_group;
	CComboBox m_cbb_cell, m_cbb_path;
	CSplitButton m_splbtn_ok;
	CListCtrlEx m_list_view;	

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int iStopload, lanshow, ibuocnhay;
	long tongkq, tongfilter;

	int iGFLType, iGFLTen, iGFLSize, iGFLModified, iGFLPath;
	int nItem, nSubItem;

	int jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType, jcotNoidung, jcotLink;

	int nTotalCol, iwCol[6];
	int iDlgW, iDlgH;

	int iChkDefault;
	bool bChkLookInDir;
	CString szTargetDirectory, szPathCopyMove;
	CString pathFolderDir;

	int iRowPut;
	CString szpathDir;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	vector<CString> vecstrFileList;
	bool bDeletefile;

	struct MyStrFilter
	{
		CString szItem;
		int iEnable;
	};

	vector<MyStrFilter> vecFilter;

	struct MyStrGroups
	{
		CString szLinhvuc;
		CString szPathFolder;

		int iRBegin;
		int iREnd;
		int iEnable;
	};

	vector< MyStrGroups> vecGroups;

	MenuIcon mnIcon;

public:
	void _SetDefault();	
	void _Timkiemdulieu();
	void _LocdulieuNangcao();
	void _XacdinhSheetLienquan();
	void _CreateGroupItems(int nItem);
	void _SetStyleGroup(RangePtr PRS);
	void _SelectAllItems();

	void _AddFilesCategory(int iOnClickBtn, int iMoveCopy = 0);

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();
	
	void _GetAllFileInFolder(CString szDir);
	void _FilterTypeOfFile(vector<MyStrFilter> &vecFt);
	void _FilterSearchText(vector<MyStrFilter> &vecFt);
	void _AddItemsListKetqua(int iBegin, int iEnd, int icheck);	
	void _AutoSothutu(_WorksheetPtr pSheet, int numStart = 1, int nRowStart = 0, int nRowEnd = 0);
	void _GetItemSelect(int icheck);
	void _SetStatusKetqua(int num);

	int _XacdinhDanhsachThumuc(vector<CString> &vecPath, int icolGet);
	int _GetAllSelectedItems(vector<int> &vecRow);
	int _GetAllCheckedItems(vector<int> &vecRow);
	int _GetCountItemsCheck();

	afx_msg void OnBnClickedSplbOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnBrowse();
	afx_msg void OnBnClickedBtnSearch();
	afx_msg void OnBnClickedRadioDir();
	afx_msg void OnBnClickedRadioSubdir();	
	afx_msg void OnBnClickedCheckAll();	
	afx_msg void OnBnClickedBtnPathmove();
	afx_msg void OnBnClickedCheckGroup();

	afx_msg void OnCbnSelchangeCbbPath();
	afx_msg void OnEnSetfocusTxtPathmove();

	afx_msg void OnAddataAddfiles();
	afx_msg void OnAddataCopyfiles();
	afx_msg void OnAddataMovefiles();
	afx_msg void OnAddataOpenfile();
	afx_msg void OnAddataOpenfolder();
	afx_msg void OnAddataDelfiles();

	afx_msg void OnAddataAddfiles2();
	afx_msg void OnAddataCopyfiles2();
	afx_msg void OnAddataMovefiles2();

	afx_msg void OnAddataCheck();
	afx_msg void OnAddataUncheck();
	afx_msg void OnAddataCheckall();

	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);	
};
