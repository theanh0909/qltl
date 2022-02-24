#if _MSC_VER > 1000
#pragma once
#endif 

#include <afxtempl.h>

class CSysImageList
{
	BOOL m_bLargeIcons;
	HIMAGELIST m_hImageList;
	int m_nFolderImage, m_nHtmlImage, m_nRemoteFolderImage, m_nUnknownTypeImage;
	static CMap<CString, LPCTSTR, int, int&> s_mapIndexCache;
public:
	CSysImageList(BOOL bLargeIcons = FALSE);
	virtual ~CSysImageList();

	CImageList* GetImageList(); // temporary. should not be stored
	HIMAGELIST GetHImageList();

	BOOL Initialize();
	int GetFileIndex(CString&  szFilePath); // will call Initialize if nec.
	int GetFolderIndex(); // will call Initialize if nec.

	HICON GetAppIcon();
	HICON GetFileIcon(CString& szFilePath);
	HICON GetFolderIcon();
	HICON GetUnknownIcon();
	int GetImageIndex(LPCTSTR szFile); // will call Initialize if nec.
private:
	// raw version that resolves directly to SHGetFileInfo
	int GetLocalFolderImage();
	int GetRemoteFolderImage();
	int GetUnknownTypeImage();
};

