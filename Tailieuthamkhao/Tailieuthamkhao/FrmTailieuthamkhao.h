#pragma once
#include "Base64.h"
#include "EasySize.h"
#include "Function.h"
#include "MenuIcon.h"

class FrmTailieuthamkhao : public CDialogEx
{
	DECLARE_EASYSIZE
	DECLARE_DYNAMIC(FrmTailieuthamkhao)

public:
	FrmTailieuthamkhao(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FrmTailieuthamkhao();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DLG_TAILIEUTHAMKHAO };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()

private:
	afx_msg BOOL PreTranslateMessage(MSG* pMsg);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSizing(UINT fwSide, LPRECT pRect);
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnContextMenu(CWnd* /*pWnd*/, CPoint /*point*/);
	afx_msg LRESULT OnHeaderRightClick(WPARAM wParam, LPARAM lParam);	

public:	
	Function ff;
	Base64Ex bb;
	MenuIcon mnIcon;

	CButton m_checkFolder, m_checkCopy, m_btn_browse;

	CThemeHelperST m_ThemeHelper;
	CXPStyleButtonST m_btn_open, m_btn_download, m_btn_close;

	CColorEdit m_txtTimkiem, m_txtPathBrowse;	
	CComboBoxExt m_cbbDulieu, m_cbbPhanloai;
	CListCtrlEx m_listdsdownload;	

	CImageList m_imageList;
	CSysImageList m_sysImageList;

	int demTimer;
	bool blChangeKey, blLoadDefault;

	int tslkq, lanshow;
	int iStopload, ibuocnhay;

	int jUserID;
	int nItem, nSubItem;
	int iTopdownload, iMaxdownload;	
	int clmKey, clmIcon, clmTen, clmPloai, clmMota, clmTrangthai, clmDuongdan, clmID;

	int jcolDownload[9];

	CString szCopyMulti;
	CString spathSaveDownload;
	CString spathLocal, szpathDownloadTL;
	CString szFileCategory, szFolderCategory, szContentCaterory;
	CString szpathListdown;

	CString szBSever, szBFolderdown, szBFilecategory, szBFilesave;
	CString szCreateKey, szFileDll, szFunction;

	// Mảng lưu trữ dữ liệu category
	struct MyStrCategory
	{
		CString filename;
		CString szfolder;
		CString szcontent;
	};
	vector<MyStrCategory> vecCategory;
	
	// Mảng lưu trữ danh sách đã download
	struct MyStructListDown
	{
		CString szPath;
		CString szBase;
	};	
	vector<MyStructListDown> vecListDownload;	

	// Mảng lưu trữ thông tin danh sách download
	struct MyStructDownload
	{
		CString sNoidung;
		CString sPhanloai;
		CString sLinkdown;
		CString sMota;
		CString sLocal;
		CString sTypeFile;

		int iPhanquyen;
		int iNumID;
	};	

	vector<MyStructDownload> vecDSDulieu;
	vector<MyStructDownload> vecDSKetqua;

public:

	// Function
	afx_msg void SetDefault();	
	afx_msg void GetBaseDecodeEx();
	afx_msg void StyleListdulieu();
	afx_msg void DeleteAllCombobox();
	afx_msg void DeleteAllListdulieu();	
	afx_msg void LoadComboboxBodulieu();	
	afx_msg void LocPhanloaidulieu();	
	afx_msg void SetLocationAndSize();
	afx_msg void SetColorType(int numRow);
	afx_msg void AddItemsListKetqua(int iBegin, int iEnd);	
	afx_msg void RightClickDownload(int itype = 0);
	afx_msg void SetDownloadColor(int numRow, CString szpth);
	afx_msg void Savepathdownload(CString szname, CString szpth, CString szURL);
	afx_msg void ShowColumnWidth(int nColumn, int nWidth, int itype);

	afx_msg int GetAllDulieu();
	afx_msg int CheckLicense();
	afx_msg int AutoCheckRobot();
	afx_msg int GetAllListdownload();	
	afx_msg int DownloadFileCategory();
	afx_msg int CheckDownload(int numRow);
	afx_msg int DownloadFile(int numRow, int itype, int iMess);
	afx_msg int CreateFileListDownload(CString spathFile);
	afx_msg int CheckTotalFiles(vector<CString> &vecTotal);	
	afx_msg int LoadDulieusudung(CString szTk = L"", CString szPl = L"");
	
	afx_msg	CString Taothumuccon(CString szpath);
	afx_msg CString _NameColumnHeading(CListCtrl &clist, int column, int itype = 1, CString szName = L"");
	
	afx_msg void SetTimerSearch();
	afx_msg void Timkiemdulieu();

	// Button
	afx_msg void OnBnClickedBtnDownload();	
	afx_msg void OnBnClickedBtnOpenfd();	
	afx_msg void OnBnClickedBtnClose();
	afx_msg void OnBnClickedCheckFolder();
	afx_msg void OnBnClickedCheckCopy();
	afx_msg void OnBnClickedBtnSave();
	afx_msg void OnCbnSelchangeCbbPloai();
	afx_msg void OnCbnSelchangeCbbDulieu();

	// Menu Content
	afx_msg void OnDwDownload();
	afx_msg void OnDwDwandview();
	afx_msg void OnDwOpenfile();
	afx_msg void OnDwOpenfolder();
	afx_msg void OnDwCopyNoidung();

	// Menu List Header
	afx_msg void OnDwShowAll();
	afx_msg void OnDwShowKey();
	afx_msg void OnDwShowNoidung();
	afx_msg void OnDwShowPhanloai();
	afx_msg void OnDwShowMota();
	afx_msg void OnDwShowTrangthai();
	afx_msg void OnDwShowPath();
	
	// List Event
	afx_msg void OnEnChangeTxtKey();
	afx_msg void OnNMClickListKqua(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMDblclkListKqua(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickListKqua(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnLvnEndScrollListKqua(NMHDR *pNMHDR, LRESULT *pResult);	
	afx_msg void OnHdnEndtrackListKqua(NMHDR *pNMHDR, LRESULT *pResult);	
};
