#pragma once
#include <afxdb.h>



#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <list>
using namespace std;
class _declspec(dllexport)CConnection :public CDatabase
{
private:
	CDatabase m_db;
	list<CStringArray*> m_lStrArr;

public:
	int m_nRow;
	int m_nCol;
	bool m_bFind;
	list<CDBVariant>ExecuteRLV(CString szSql);
	CStringArray *ExecuteRSA(CString szSql);
	CStringArray* ExecuteT(CString szSql);
	CConnection();
	virtual ~CConnection();

	BOOL OpenMySqlDB(CString szServer, CString szDatabase, CString szUser, CString szPass, CString szPort = _T("3306"));
	BOOL OpenAccessDB(CString szDatabase, CString szUser, CString szPass);
	BOOL OpenSQLServerDB(CString szServer, CString szDatabase, CString szUser, CString szPass);
	//Connect to CSV file
	BOOL OpenCSVDB(CString szDbq);
	//Connect to Excel file
	BOOL OpenExcelDB(CString szFile);
	//Close database
	BOOL CloseDatabase();

	CString ExecuteRS(CString szSql);
	BOOL ExecuteRB(CString szSql);
	long ExecuteRL(CString szSql);
	int ExecuteRI(CString szSql);
	double ExecuteRD(CString szSql);
	CRecordset* Execute(const CString &str);
	int ExecuteInt(const CString &str);
	//Execute SQL - CSV Command 
	CRecordset* ExecuteCSV(const CString &str);
	list<int> ExecuteRLV1(CString szSql, CString szFind);

	BOOL ImportFile(CString szTableName, CString szColName, int nID, CString szFileName);
	BOOL ExportFile(CString szTableName, CString szColName, int nID, CString szFileName);
};

