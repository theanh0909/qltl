#pragma once
#include "stdafx.h"

class DlgCreateFroups : public CDialogEx
{
	DECLARE_DYNAMIC(DlgCreateFroups)

public:
	DlgCreateFroups(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgCreateFroups();

	// Dialog Data
	enum { IDD = DLG_CREATE_GROUPS };

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);

public:
	CXPStyleButtonST m_btn_ok, m_btn_cancel, m_btn_search;
	CThemeHelperST m_ThemeHelper;

	CListCtrlEx m_list_view;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	CString pathFolderDir;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int iStopload, lanshow, ibuocnhay;
	long tongkq;

	int iGFLTten, iGFLSlg;
	int nItem, nSubItem;

	int jcotSTT, jcotFolder, jcotLV, jcotType, jcotTen, jcotNoidung, jcotLink;

	int iwColTen, iwColSL;
	int iDlgW, iDlgH;

	int bChecked;

	struct MyStrGroups
	{
		CString szLinhvuc;
		CString szPathFolder;

		int iRBegin;
		int iREnd;
		int iEnable;
	};

	vector< MyStrGroups> vecGroups;

public:
	void _GetDanhsachLinhvuc();		// <-- Hàm quan trọng xác định số lượng files sử dụng
	void _XacdinhSheetLienquan();
	void _SetDefault();
	
	void _SaveRegistry();
	void _LoadRegistry();	
	void _LoadDanhsachNhom();
	void _SetLocationAndSize();
	void _CreateGroupItems(int nItem);
	void _SetStyleGroup(RangePtr PRS);
	void _AddItemsListKetqua(int iBegin, int iEnd, int icheck);	
	

	int _GetAllCheckedItems(vector<int> &vecRow);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnLvnEndScrollListView(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnHdnItemStateIconClickListView(NMHDR *pNMHDR, LRESULT *pResult);
};
