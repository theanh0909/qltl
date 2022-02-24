#pragma once
#include "StdAfx.h"

class DlgCheckFiles : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgCheckFiles)

public:
	DlgCheckFiles(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgCheckFiles();

	// Dialog Data
	enum { IDD = DLG_CHECK_FILES };

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
	CXPStyleButtonST m_btn_ok, m_btn_cancel, m_btn_search;
	CThemeHelperST m_ThemeHelper;

	CButton m_check_all, m_check_new, m_check_link;
	CButton m_chkLookInDir, m_chkLookInSubdirectories;
	CColorEdit m_edtTargetDirectory, m_edtTotalFiles, m_lbl_total;
	CListCtrlEx m_list_view;
	CComboBox m_cbb_path;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	CString szTargetDirectory;
	CString pathFolderDir;
	int iLenPathDir;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	vector<CString> vecstrFileList;
	bool bChkLookInDir, bChkLink, bChkNew;
	bool bSetGroup, bDeletefile;	

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int iStopload, lanshow, ibuocnhay;
	long tongkq, tongfilter;

	int iGFLType, iGFLTen, iGFLSize, iGFLModified, iGFLPath;
	int nItem, nSubItem;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

	int nTotalCol, iwCol[6];
	int iDlgW, iDlgH;
	int slnew;

	struct MyStrFilter
	{
		CString szItem;
		int iFileNew;
		int iEnable;
	};

	vector<MyStrFilter> vecFilter;
	vector<CString> vecHyper;

	MenuIcon mnIcon;

public:
	void _GetDanhsachFiles(int iSelectCell = -1);		// <-- Hàm quan trọng xác định số lượng files sử dụng

	void _SetDefault();
	void _Timkiemdulieu();
	void _LocdulieuNangcao();
	void _DanhsachHyperlink();
	void _SelectAllItems();	
	void _XacdinhSheetLienquan();
	void _GetGroupRowStartEnd(int &iRowStart, int &iRowEnd);

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();

	void _GetAllFileInFolder(CString szDir);
	void _AddFilesCategory(int iOnClickBtn);
	void _AddItemsListKetqua(int iBegin, int iEnd, int ichkNew, int ichkLink);
	void _AutoSothutu(_WorksheetPtr pSheet, int numStart = 1, int nRowStart = 0, int nRowEnd = 0);
	void _GetItemSelect(int icheck);
	void _SetStatusKetqua(int num);
	void _CheckFileHyperlink(int nIndex);
	void _PutStyleCell(RangePtr PRS);

	int _CheckFileNew(CString szpathFile);	
	int _GetRowGroups(CString szpathFile);

	int _XacdinhDanhsachThumuc(vector<CString> &vecPath, int icolGet);
	int _GetAllSelectedItems(vector<int> &vecRow);
	int _GetAllCheckedItems(vector<int> &vecRow);
	int _GetCountItemsCheck();

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnBrowse();
	afx_msg void OnBnClickedCheckAll();
	afx_msg void OnBnClickedCheckNews();	
	afx_msg void OnBnClickedCheckLink();
	afx_msg void OnBnClickedRadioDir();
	afx_msg void OnBnClickedRadioSubdir();
	afx_msg void OnCbnSelchangeCbbPath();

	afx_msg void OnChkflAddfiles();
	afx_msg void OnChkflGotosheet();
	afx_msg void OnChkflOpenfile();
	afx_msg void OnChkflOpenfolder();
	afx_msg void OnChkflDelfiles();

	afx_msg void OnChkflCheck();
	afx_msg void OnChkflUncheck();
	afx_msg void OnChkflCheckall();

	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);	
};
