#include "stdafx.h"
#include "FileHelper.h"
#include <shlwapi.h>   
#include "SysImageList.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif

CMap<CString, LPCTSTR, int, int&> CSysImageList::s_mapIndexCache;

CSysImageList::CSysImageList(BOOL bLargeIcons) :
	m_bLargeIcons(bLargeIcons),
	m_nFolderImage(-1),
	m_nHtmlImage(-1),
	m_hImageList(NULL),
	m_nRemoteFolderImage(-1),
	m_nUnknownTypeImage(-1)
{

}

CSysImageList::~CSysImageList()
{

}

BOOL CSysImageList::Initialize()
{
	if (m_hImageList)
		return TRUE;

	// set up system image list
	TCHAR szWindows[MAX_PATH];
	GetWindowsDirectory(szWindows, MAX_PATH);

	SHFILEINFO sfi;

	UINT nFlags = SHGFI_SYSICONINDEX | (m_bLargeIcons ? SHGFI_ICON : SHGFI_SMALLICON);
	HIMAGELIST hIL = (HIMAGELIST)SHGetFileInfo(szWindows, 0, &sfi, sizeof(sfi), nFlags);

	if (hIL)
	{
		m_hImageList = hIL;

		// intialize the stock icons
		m_nFolderImage = sfi.iIcon;

		// intialize html and remote folder images on demand
	}

	// else
	return (m_hImageList != NULL);
}

int CSysImageList::GetImageIndex(LPCTSTR szFile)
{
	if (!m_hImageList && !Initialize() || !szFile || !(*szFile))
		return -1;

	SHFILEINFO sfi;

	UINT nFlags = SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES | (m_bLargeIcons ? SHGFI_ICON : SHGFI_SMALLICON);
	HIMAGELIST hIL = (HIMAGELIST)SHGetFileInfo(szFile, FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), nFlags);

	if (hIL)
		return sfi.iIcon;

	return -1;
}

int CSysImageList::GetFileIndex(CString& szFilePath)
{
	if (!m_hImageList && !Initialize() || !szFilePath || !(*szFilePath))
		return -1;

	// check index cache
	int nIndex = -1;

	if (s_mapIndexCache.Lookup(szFilePath, nIndex))
		return nIndex;


	CSplitPath path(szFilePath);
	CString sExt = path.GetExt();
	CString sFName = path.GetFName();
	// check if its a folder first if it has no extension
	BOOL bRemotePath = path.IsRemotePath();
	if (bRemotePath)
	{
		// do not access the file for any reason. 
		// ie. we need to make an educated guess
		// if no extension assume a folder
		if (sExt.IsEmpty() || sExt == _T("."))
			nIndex = GetRemoteFolderImage();
		else
		{
			nIndex = GetImageIndex(sExt);

			// if the icon index is invalid or there's no filename
			// then assume it's a folder
			if (nIndex < 0 || sFName.IsEmpty())
				nIndex = GetRemoteFolderImage();
		}
	}
	else // local file so we can do whatever we like ;)
	{
		if (CFileHelper::IsDirExist(szFilePath))
		{
			nIndex = GetFolderIndex();
		}
		else if (sExt.CompareNoCase(_T(".lnk")) == 0)
		{
			// get icon for item pointed to
			CString sReferencedFile = CFileHelper::ResolveShortcut(szFilePath);
			// RECURSIVE call
			nIndex = GetFileIndex(sReferencedFile);
		}
		else
		{
			nIndex = GetImageIndex(szFilePath);
			// can fail with full path, so we then revert to extension only
			if (nIndex < 0)
				nIndex = GetImageIndex(sExt);
		}
		if (nIndex < 0 || nIndex == GetUnknownTypeImage())
			if (sFName.IsEmpty() || sExt.IsEmpty()) // assume it's a folder unless it looks like a file
				nIndex = GetFolderIndex();
	}
	// record for posterity unless icon was not retrieved
	if (nIndex != -1 && nIndex != GetUnknownTypeImage())
	{
		s_mapIndexCache[szFilePath] = nIndex;
	}

	return nIndex;
}

int CSysImageList::GetLocalFolderImage()
{
	ASSERT(m_nFolderImage != -1);
	return m_nFolderImage;
}

int CSysImageList::GetRemoteFolderImage()
{
	if (m_nRemoteFolderImage == -1)
		m_nRemoteFolderImage = GetImageIndex(_T("\\\\dummy\\."));

	return m_nRemoteFolderImage;
}

int CSysImageList::GetUnknownTypeImage()
{
	if (m_nUnknownTypeImage == -1)
		m_nUnknownTypeImage = GetImageIndex(_T(".6553BB15-9369-4227-BCA0-F523A35F1DAB"));

	return m_nUnknownTypeImage;
}

int CSysImageList::GetFolderIndex()
{
	if (!m_hImageList && !Initialize())
		return -1;

	return m_nFolderImage;
}

HICON CSysImageList::GetAppIcon()
{
	return GetFileIcon(CFileHelper::GetAppFileName());
}

HICON CSysImageList::GetFileIcon(CString& szFilePath)
{
	if (!m_hImageList && !Initialize())
		return NULL;

	int nIndex = GetFileIndex(szFilePath);

	if (nIndex != -1)
		return ImageList_GetIcon(m_hImageList, nIndex, 0);

	return NULL;
}

HICON CSysImageList::GetUnknownIcon()
{
	if (!m_hImageList && !Initialize())
		return NULL;

	int nIndex = GetUnknownTypeImage();

	if (nIndex != -1)
		return ImageList_GetIcon(m_hImageList, nIndex, 0);
	return NULL;
}

HICON CSysImageList::GetFolderIcon()
{
	if (!m_hImageList && !Initialize())
		return NULL;

	return ImageList_GetIcon(m_hImageList, m_nFolderImage, 0);
}

CImageList* CSysImageList::GetImageList()
{
	if (!m_hImageList)
		Initialize();

	CImageList* pIL = CImageList::FromHandle(m_hImageList);
	return pIL;
}

HIMAGELIST CSysImageList::GetHImageList()
{
	if (!m_hImageList)
		Initialize();
	return m_hImageList;
}

