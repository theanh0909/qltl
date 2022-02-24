class CFileHelper
{
private:
	CFileHelper() {};
	~CFileHelper() {};
public:
	static unsigned long GetFileSize(CString&  file);
	static BOOL IsFileExist(CString& file);
	static BOOL IsFileOpen(CString& filePath);
	static BOOL IsShortcut(CString& str);
	static BOOL IsDirExist(CString& dir);
	static BOOL CreateDirIfNeeded(CString& path);
	static CString ResolveShortcut(CString& szShortcut);

	static CString GetModuleFileName(HMODULE hMod);
	static CString GetAppFileName();
	static CString GetAppFolder();
	static CString GetCurDir();
	static CString GetRelativePath(const CString& sFilePath, CString& sRelativeToFolder, BOOL bFolder);
	static DWORD Run(HWND hwnd, LPCTSTR lpFile, LPCTSTR lpDirectory, int nShowCmd);
	// add slash \ to end of path 
	static CString& TerminatePath(CString& sPath);
	// remove last \ of path if any 
	static CString& UnterminatePath(CString& sPath);
	static CString GetFullPath(const CString& sFilePath, const CString& sRelativeToFolder);
};
/*******************************************************************/
class CSplitPath
{
	CString m_path;
public:
	CSplitPath(LPCTSTR lpszPathBuffer = NULL);
	virtual ~CSplitPath();

protected:
	TCHAR m_szDrive[_MAX_DRIVE];
	TCHAR m_szDir[_MAX_DIR];
	TCHAR m_szFName[_MAX_FNAME];
	TCHAR m_szExt[_MAX_EXT];

public:
	void SplitPath(LPCTSTR lpszPathBuffer);
	CString GetDrive() const;
	CString GetDir() const;
	CString GetFName() const;
	CString GetExt() const;
	CString GetFullPath() const;
	CString GetFullName() const;
	static void SuggestFileName(CString& fileFullPath);
	BOOL IsRemotePath(BOOL bAllowFileCheck = FALSE);
	BOOL IsUNCPath();
	BOOL IsMappedPath();
	BOOL IsFixedPath();
	BOOL IsReadonlyPath();
	//used for database file system
	BOOL IsDBPath() const;
	CString MakeDBPath() const;
};