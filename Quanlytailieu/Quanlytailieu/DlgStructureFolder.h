#pragma once
#include "stdafx.h"

class DlgStructureFolder : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(DlgStructureFolder)

public:
	DlgStructureFolder(CWnd* pParent = NULL);   // standard constructor
	virtual ~DlgStructureFolder();

	// Dialog Data
	enum { IDD = DLG_DIRSTRUCTURE };

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
	CXPStyleButtonST m_btn_ok, m_btn_cancel, m_btn_temp;
	CThemeHelperST m_ThemeHelper;
	CColumnTreeWnd m_TreeWnd;
	CColorEdit m_txt_name, m_txt_output;
	CComboBoxExt m_cbb_file;
	CColorEdit m_lbl_hdan;

	Function *ff;
	Base64Ex *bb;
	CRegistry *reg;

	_bstr_t bsFManager;
	_WorksheetPtr psFManager;
	RangePtr prFManager;

	CString pathFolderDir;
	CString szTargetDirectory, szTargetOutput;

	int icotSTT, icotLV, icotMH, icotNoidung, icotFilegoc, icotFileGXD, icotThLuan, icotEnd;
	int irowTieude, irowStart, irowEnd;
	int icotFolder, icotTen, icotType;

	int tclName, tclDescribe, tclPath, tclNew;
	int nItem, nSubItem;

	int jcotSTT, jcotLV, jcotFolder, jcotTen, jcotType, jcotNoidung, jcotLink;	
	int nTotalCol, iwCol[4];
	int iDlgW, iDlgH;

	MenuIcon mnIcon;
	CImageList m_imageTree;

	struct MyStrDirStructure
	{
		int iLevel;
		CString szFolder;
		CString szFileName;
		CString szNdung;
		CString szLink;
	};

	vector< MyStrDirStructure> vecDir;
	vector<CString> vecCreate;

	struct MyStrSheet
	{
		CString szName;
		CString szCode;
		CString szKodau;
		CString szLinhvuc;
	};

public:
	void _SetDefault();
	void _XacdinhSheetLienquan();

	void _SetLocationAndSize();
	void _SaveRegistry();
	void _LoadRegistry();	
	void _LoadTreeView(CString szFile);	
	void _TaoSheetDulieuMoi(CString szFileName);	
	void _CopyStyleCategory(_WorksheetPtr pSheet);
	bool _CreateNewSheet(CString szNameSheet);
	void _LoadCombobox(CString szPath);
	bool _LoadFileStructure(CString szFile);

	afx_msg void OnBnClickedBtnOk();
	afx_msg void OnBnClickedBtnCancel();
	afx_msg void OnBnClickedBtnBrowse();
	afx_msg void OnBnClickedBtnTemp();
	afx_msg void OnBnClickedBtnOutput();
	afx_msg void OnCbnSelchangeTxtPath();
};
