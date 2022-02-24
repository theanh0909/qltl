#include "pch.h"
#include "StdAfx.h"
#include "FileRegister.h"
#include "FileHelper.h"
#include "AtlBase.h"

CFileRegister::CFileRegister(LPCTSTR szExt, LPCTSTR szFileType)
{
	m_sExt = szExt;
	m_sFileType = szFileType;

	m_sExt.TrimRight();
	m_sExt.TrimLeft();
	m_sFileType.TrimRight();
	m_sFileType.TrimLeft();

	if (!m_sExt.IsEmpty() && m_sExt[0] != '.')
		m_sExt = _T(".") + m_sExt;

	// get the app path
	m_sAppPath = CFileHelper::GetAppFileName();

	ASSERT(!m_sAppPath.IsEmpty());
}

CFileRegister::~CFileRegister()
{
}

BOOL CFileRegister::RegisterFileType(LPCTSTR szFileDesc, int nIcon, BOOL bAllowShellOpen, LPCTSTR szParams, BOOL bAskUser)
{
	CRegKey reg;
	CString sKey;
	CString sEntry;
	int nRet = IDYES;
	CString sMessage;
	BOOL bSuccess = TRUE, bChange = FALSE;

	ASSERT(!m_sExt.IsEmpty());
	ASSERT(!m_sFileType.IsEmpty());

	if (m_sExt.IsEmpty() || m_sFileType.IsEmpty())
		return FALSE;

	if (reg.Open(HKEY_CLASSES_ROOT, m_sExt) == ERROR_SUCCESS)
	{
		ULONG ln = 0;
		if (reg.QueryStringValue(_T(""), sEntry.GetBuffer(), &ln) == ERROR_MORE_DATA)
		{
			reg.QueryStringValue(_T(""), sEntry.GetBuffer(ln + 1), &ln);
			sEntry.ReleaseBuffer();
		}
		// if the current filetype is not correct query the use to reset it
		if (!sEntry.IsEmpty() && sEntry.CompareNoCase(m_sFileType) != 0)
		{
			if (bAskUser)
			{
				sMessage.Format(_T("The file extension %s is used by %s for its %s.\n\nWould you like %s to be the default application for this file type."),
					m_sExt, AfxGetAppName(), szFileDesc, AfxGetAppName());

				nRet = AfxMessageBox(sMessage, MB_YESNO | MB_ICONQUESTION);
			}

			bChange = TRUE;
		}
		else
			bChange = sEntry.IsEmpty();

		// if not no then set
		if (nRet != IDNO)
		{
			bSuccess = (reg.SetKeyValue(_T(""), m_sFileType) == ERROR_SUCCESS);

			reg.Close();

			if (reg.Open(HKEY_CLASSES_ROOT, m_sFileType) == ERROR_SUCCESS)
			{
				bSuccess &= (reg.SetKeyValue(_T(""), szFileDesc) == ERROR_SUCCESS);
				reg.Close();

				// app to open file
				if (reg.Open(HKEY_CLASSES_ROOT, m_sFileType + "\\shell\\open\\command") == ERROR_SUCCESS)
				{
					CString sShellOpen;

					if (bAllowShellOpen)
					{
						if (szParams)
							sShellOpen = _T("\"") + m_sAppPath + _T("\" \"%1\" ") + CString(szParams);
						else
							sShellOpen = _T("\"") + m_sAppPath + _T("\" \"%1\"");
					}

					bSuccess &= (reg.SetKeyValue(_T(""), sShellOpen) == ERROR_SUCCESS);
				}
				else
					bSuccess = FALSE;

				// icons
				reg.Close();

				if (reg.Open(HKEY_CLASSES_ROOT, m_sFileType + _T("\\DefaultIcon")) == ERROR_SUCCESS)
				{
					CString sIconPath;
					sIconPath.Format(_T("%s,%d"), m_sAppPath, nIcon);
					bSuccess &= (reg.SetStringValue(_T(""), sIconPath) == ERROR_SUCCESS);
				}
				else
					bSuccess = FALSE;
			}
			else
				bSuccess = FALSE;
		}
		else
			bSuccess = FALSE;
	}
	else
		bSuccess = FALSE;

	if (bSuccess && bChange)
		SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0);

	return bSuccess;
}

BOOL CFileRegister::UnRegisterFileType()
{
	CRegKey reg;
	CString sKey;
	CString sEntry;
	BOOL bSuccess = FALSE;

	ASSERT(!m_sExt.IsEmpty());

	if (m_sExt.IsEmpty())
		return FALSE;

	if (reg.Open(HKEY_CLASSES_ROOT, m_sExt) == ERROR_SUCCESS)
	{
		ULONG ln = 0;
		if (reg.QueryStringValue(_T(""), sEntry.GetBuffer(), &ln) == ERROR_MORE_DATA)
		{
			reg.QueryStringValue(_T(""), sEntry.GetBuffer(ln + 1), &ln);
			sEntry.ReleaseBuffer();
		}


		if (sEntry.IsEmpty())
			return TRUE; // we werent the register app so all's well

		ASSERT(!m_sFileType.IsEmpty());

		if (m_sFileType.IsEmpty())
			return FALSE;

		if (sEntry.CompareNoCase(m_sFileType) != 0)
			return TRUE; // we werent the register app so all's well

		// else delete the keys
		reg.Close();
		reg.Open(HKEY_CLASSES_ROOT, _T(""));
		reg.DeleteSubKey(m_sExt);
		reg.DeleteSubKey(m_sFileType);
		reg.Close();
		SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0);

		bSuccess = TRUE;
	}

	return bSuccess;
}

BOOL CFileRegister::IsRegisteredAppInstalled()
{
	ASSERT(!m_sExt.IsEmpty());

	if (m_sExt.IsEmpty())
		return FALSE;

	CString sRegAppPath = GetRegisteredAppPath();

	if (!sRegAppPath.IsEmpty())
	{
		CFileStatus fs;

		return CFile::GetStatus(sRegAppPath, fs);
	}

	return FALSE;
}

BOOL CFileRegister::IsRegisteredApp()
{
	CRegKey reg;
	CString sEntry;

	ASSERT(!m_sExt.IsEmpty());

	if (m_sExt.IsEmpty())
		return FALSE;

	// if the file type is not empty we check the file type first
	if (!m_sFileType.IsEmpty())
	{
		if (reg.Open(HKEY_CLASSES_ROOT, m_sExt) == ERROR_SUCCESS)
		{
			ULONG ln = 0;
			if (reg.QueryStringValue(_T(""), sEntry.GetBuffer(), &ln) == ERROR_MORE_DATA)
			{
				reg.QueryStringValue(_T(""), sEntry.GetBuffer(ln + 1), &ln);
				sEntry.ReleaseBuffer();
			}
			if (sEntry.IsEmpty() || sEntry.CompareNoCase(m_sFileType) != 0)
				return FALSE;
		}
	}

	// then we check the app itself
	CString sRegAppFileName = GetRegisteredAppPath(TRUE);

	if (!sRegAppFileName.IsEmpty())
	{
		CSplitPath path(m_sAppPath);
		CString sFName = path.GetFName() + path.GetExt();
		if (sRegAppFileName.CompareNoCase(sFName) == 0)
			return TRUE;
	}

	return FALSE;
}

CString CFileRegister::GetRegisteredAppPath(BOOL bFilenameOnly)
{
	CRegKey reg;
	CString sEntry, sAppPath;

	ASSERT(!m_sExt.IsEmpty());

	if (reg.Open(HKEY_CLASSES_ROOT, m_sExt) == ERROR_SUCCESS)
	{
		ULONG ln = 0;
		if (reg.QueryStringValue(_T(""), sEntry.GetBuffer(), &ln) == ERROR_MORE_DATA)
		{
			reg.QueryStringValue(_T(""), sEntry.GetBuffer(ln + 1), &ln);
			sEntry.ReleaseBuffer();
		}

		if (!sEntry.IsEmpty())
		{
			reg.Close();

			if (reg.Open(HKEY_CLASSES_ROOT, sEntry) == ERROR_SUCCESS)
			{
				reg.Close();

				// app to open file
				if (reg.Open(HKEY_CLASSES_ROOT, sEntry + CString("\\shell\\open\\command")) == ERROR_SUCCESS)
				{
					ULONG ln = 0;
					if (reg.QueryStringValue(_T(""), sAppPath.GetBuffer(), &ln) == ERROR_MORE_DATA)
					{
						reg.QueryStringValue(_T(""), sAppPath.GetBuffer(ln + 1), &ln);
						sAppPath.ReleaseBuffer();
					}
				}
			}
		}
	}

	if (!sAppPath.IsEmpty())
	{
		CString sFName;
		CSplitPath path(sAppPath);
		sFName = path.GetFName();
		sFName += ".exe";

		if (bFilenameOnly)
			sAppPath = sFName;
		else
			sAppPath = path.GetFullPath() + sFName;
	}
	return sAppPath;
}

// static versions
BOOL CFileRegister::IsRegisteredAppInstalled(LPCTSTR szExt)
{
	return CFileRegister(szExt).IsRegisteredAppInstalled();
}

CString CFileRegister::GetRegisteredAppPath(LPCTSTR szExt, BOOL bFilenameOnly)
{
	return CFileRegister(szExt).GetRegisteredAppPath(bFilenameOnly);
}

BOOL CFileRegister::RegisterFileType(LPCTSTR szExt, LPCTSTR szFileType, LPCTSTR szFileDesc, int nIcon, BOOL bAllowShellOpen, LPCTSTR szParams, BOOL bAskUser)
{
	return CFileRegister(szExt, szFileType).RegisterFileType(szFileDesc, nIcon, bAllowShellOpen, szParams, bAskUser);
}

BOOL CFileRegister::UnRegisterFileType(LPCTSTR szExt, LPCTSTR szFileType)
{
	return CFileRegister(szExt, szFileType).UnRegisterFileType();
}

BOOL CFileRegister::IsRegisteredApp(LPCTSTR szExt, LPCTSTR szFileType)
{
	return CFileRegister(szExt, szFileType).IsRegisteredApp();
}
