#pragma once
#include "stdafx.h"

class DlgRenameFiles : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgRenameFiles)

public:
	DlgRenameFiles(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgRenameFiles();

	// Dialog Data
	enum { IDD = DLG_RENAME_FILES };

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

	CListCtrlEx m_list_view;
	CButton m_check_all, m_check_new;
	CColorEdit m_editSearchFiles, m_edtWildCard, m_edtTotalFiles;
	CColorEdit m_lbl_total, m_lbl_help;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	CString pathFolderDir;

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	bool bLookInSubdirectories;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int iStopload, lanshow, ibuocnhay;
	long tongkq, tongfilter;

	int iGFLType, iGFLNmNew, iGFLNmOld, iGFLPath;
	int nItem, nSubItem;
	int iKeyESC;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

	int nTotalCol, iwCol[5];
	int iDlgW, iDlgH;

	CString szBefore, szAfter;

	struct MyStrFiles
	{
		CString szHyperlink;
		CString sznmOld;
		CString sznmNew;		

		int iEnable;
		int iRow;
		int iID;
	};

	vector< MyStrFiles> vecItem;
	vector< MyStrFiles> vecFilter;

	MenuIcon mnIcon;

public:
	void _GetDanhsachFiles(int iSelectCell = -1);		// <-- Hàm quan trọng xác định số lượng files sử dụng

	void _SetDefault();
	void _LocdulieuNangcao();
	void _AddItemsListKetqua(int iBegin, int iEnd, int icheck);

	void _XacdinhSheetLienquan();
	void _SelectAllItems();
	void _SetStatusKetqua(int num);

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();
	void _ClickRenameFiles(int iOnClickBtn);

	void _GetItemSelect(int icheck);
	int _GetAllSelectedItems(vector<int> &vecRow);
	int _GetAllCheckedItems(vector<int> &vecRow);
	int _GetCountItemsCheck();

	void _FilterTypeOfFile(vector<MyStrFiles> &vecFt);
	void _FilterSearchText(vector<MyStrFiles> &vecFt);

	void _GetGroupRowStartEnd(int &iRowStart, int &iRowEnd);
	void _UpdatePropertiesFile(int nRow, int itype);
	void _RightClickRenameFiles(int itype);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedCheckAll();
	afx_msg void OnBnClickedCheckNews();
	afx_msg void OnBnClickedBtnSearch();

	afx_msg void OnEnSetfocusTxtSearch();
	afx_msg void OnEnSetfocusTxtTypefile();
	afx_msg void OnEnSetfocusEditView();
	afx_msg void OnEnKillfocusEditView();

	afx_msg void OnRenameRfiles();
	afx_msg void OnRenameOpenfile();
	afx_msg void OnRenameOpenfolder();
	afx_msg void OnRenameDeletefiles();
	afx_msg void OnRenameCheck();
	afx_msg void OnRenameUncheck();
	afx_msg void OnRenameCheckall();
	afx_msg void OnRenameCopy();

	afx_msg void OnNMClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnKeydownListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnHdnEndtrackListView(NMHDR *pNMHDR, LRESULT *pResult);	
};
