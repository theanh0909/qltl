// Quanlytailieu.cpp : Defines the initialization routines for the DLL.
//

#include "pch.h"
#include "framework.h"
#include "StdAfx.h"

#include "Quanlytailieu.h"
#include "DlgListUpdate.h"
#include "DlgOptionMain.h"
#include "DlgCopyData.h"
#include "DlgAddFiles.h"
#include "DlgAutoRename.h"
#include "DlgCheckFiles.h"
#include "DlgRenameFiles.h"
#include "DlgRibbonTimkiem.h"
#include "DlgInsertRows.h"
#include "DlgCreateNewTemp.h"
#include "DlgCreateFroups.h"
#include "DlgSelectionCells.h"
#include "DlgEditProperties.h"
#include "DlgChangeHyperlink.h"
#include "DlgStructureFolder.h"
#include "CCheckHyperlinkError.h"
#include "CRenameFolders.h"
#include "CCheckVersion.h"
#include "CHelpRibbon.h"
#include "CQSortData.h"
#include "CMoveFiles.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CQuanlytailieuApp

BEGIN_MESSAGE_MAP(CQuanlytailieuApp, CWinApp)

END_MESSAGE_MAP()


// CQuanlytailieuApp construction

CQuanlytailieuApp::CQuanlytailieuApp()
{
	
}

CQuanlytailieuApp::~CQuanlytailieuApp()
{
	CoUninitialize();
}

// The one and only CQuanlytailieuApp object

CQuanlytailieuApp theApp;


// CQuanlytailieuApp initialization

BOOL CQuanlytailieuApp::InitInstance()
{
	CWinApp::InitInstance();

	return TRUE;
}

/**********************************************************/

void __stdcall CallOptionMain()
{
	if (!bCheckVersion) return;
	if (!blCheckFileTemp()) return;

	xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	DlgOptionMain p;
	p.DoModal();
}

void __stdcall CallCreateNewSheet()
{
	if (!bCheckVersion) return;
	if (!blCheckFileTemp()) return;

	Function *ff = new Function;
	ff->_xlGetInfoSheetActive();

	xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	DlgCreateNewTemp p;
	p.DoModal();

	delete ff;
}

void __stdcall CallCreateDirStructure()
{
	if (!bCheckVersion) return;
	if (!blCheckFileTemp()) return;

	Function *ff = new Function;
	ff->_xlGetInfoSheetActive();

	xl->EnableCancelKey = XlEnableCancelKey(FALSE);
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	DlgStructureFolder p;
	p.DoModal();

	delete ff;	
}

void __stdcall CallFontUpper()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFontLower()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFontProper()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFontUppro()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFontUnsigned()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFontReverse()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQManaSheet()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQAutoSheet()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQAutoFilter()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQFreezePanes()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQViewFormula()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallQFocus()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallFuncUpdateFiles(bool blUpdate)
{
	if (!bCheckVersion) return;
	if (!blCheckFileTemp()) return;

	Function *ff = new Function;

	sCodeActive = shCategory;
	sNameActive = (LPCTSTR)ff->_xlGetNameSheet((_bstr_t)sCodeActive, 1);
	pShActive = xl->Sheets->GetItem((_bstr_t)sNameActive);
	pRgActive = pShActive->Cells;

	iRowActive = pShActive->Application->ActiveCell->GetRow();
	iColumnActive = pShActive->Application->ActiveCell->GetColumn();

	DlgListUpdate p;
	if (p._CheckUpdateData(blUpdate) == 1)
	{
		if ((int)pShActive->GetVisible() != -1) pShActive->PutVisible(XlSheetVisibility::xlSheetVisible);
		pShActive->Activate();
	}
	/*else
	{
		ff->_xlPutScreenUpdating(false);
		pShActive = xl->Sheets->GetItem(ff->_xlGetNameSheet((_bstr_t)shFManager, 1));
		if ((int)pShActive->GetVisible() != -1) pShActive->PutVisible(XlSheetVisibility::xlSheetVisible);
		pShActive->Activate();
		ff->_xlPutScreenUpdating(true);
	}*/

	delete ff;
}

void __stdcall CallFuncAddFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_RegistryCOMGpro();
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		xl->EnableCancelKey = XlEnableCancelKey(FALSE);
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		DlgAddFiles p;
		p.iRowPut = iRowActive;
		p.DoModal();

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncCopyData()
{
	if (!bCheckVersion) return;
	if (!blCheckFileTemp()) return;

	Function *ff = new Function;
	ff->_xlGetInfoSheetActive();
	ff->_DefaultSheetActive(shCategory);

	DlgCopyData p;
	p._LoadAllCategory();

	delete ff;
}

void __stdcall CallFuncCheckFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_RegistryCOMGpro();
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		int nResult = ff->_xlGetAllCellSelection(0, 1);
		if (nResult > 0)
		{
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DlgSelectionCells ps;
			ps.blEnableSelection = true;
			ps.DoModal();

			nResult = ps.iSelectCell;
			if (nResult >= 0)
			{
				DlgCheckFiles pc;
				pc._GetDanhsachFiles(nResult);
			}
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncQuickSort()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szFile = shFile;
		if (sCodeActive == shFManager || sCodeActive.Left(szFile.GetLength()) == szFile)
		{
			CQSortData *p = new CQSortData;
			p->_SapxepTatcaDulieu();
			delete p;
		}
		else ff->_msgbox(MS_QLFILES, 0, 0);

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncRenameFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		int nResult = ff->_xlGetAllCellSelection(0, 1);
		if (nResult > 0)
		{
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DlgSelectionCells ps;
			ps.blEnableSelection = false;
			ps.DoModal();

			int iSelect = ps.iSelectCell;
			if (iSelect >= 0)
			{
				DlgRenameFiles pr;
				if(iSelect > 0 || nResult > 1) pr._GetDanhsachFiles(iSelect);
				else pr._RightClickRenameFiles(0);
			}
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncRenameFolders()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		CRenameFolders *p = new CRenameFolders;
		p->_GetDanhsachFolders();
		delete p;

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncAutoRename()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;

		xl->EnableCancelKey = XlEnableCancelKey(FALSE);
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		DlgAutoRename p;
		p.DoModal();

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncCheckHyperlink()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szCopy = shCopy, szFile = shFile;
		if (sCodeActive == shCategory || sCodeActive == shFManager
			|| sCodeActive.Left(szCopy.GetLength()) == szCopy
			|| sCodeActive.Left(szFile.GetLength()) == szFile)
		{
			CCheckHyperlinkError *p = new CCheckHyperlinkError;
			p->_KiemtraLienketError();
			delete p;
		}
		else ff->_msgbox(MS_QLVBFILES, 0, 0);

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallFuncChangeHyperlink()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szCopy = shCopy, szFile = shFile;
		CString szCode = (LPCTSTR)pShActive->CodeName;
		if (szCode == shCategory || szCode == shFManager
			|| szCode.Left(szCopy.GetLength()) == szCopy
			|| szCode.Left(szFile.GetLength()) == szFile
			|| szCode == shTieuchuan)
		{
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DlgChangeHyperlink p;
			p.DoModal();
		}	

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallOnChangeFuncTimkiem()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;
		
		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		DlgRibbonTimkiem p;
		p._Timkiemdulieu();

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallClickBtnFuncTimkiem()
{
	CallOnChangeFuncTimkiem();
}

void __stdcall CallUtilCreateGroups()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		DlgCreateFroups p;
		p._GetDanhsachLinhvuc();

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallUtilEditProperties()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_RegistryCOMGpro();
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		int nResult = ff->_xlGetAllCellSelection(0, 1);
		if (nResult > 0)
		{
			xl->EnableCancelKey = XlEnableCancelKey(FALSE);
			AFX_MANAGE_STATE(AfxGetStaticModuleState());
			DlgSelectionCells ps;
			ps.blEnableSelection = false;
			ps.DoModal();

			nResult = ps.iSelectCell;
			if (nResult >= 0)
			{
				DlgEditProperties pe;
				pe._GetDanhsachFiles(nResult);
			}
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallHelpUM()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpUM();
	delete p;
}

void __stdcall CallHelpFAQ()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpFAQ();
	delete p;
}

void __stdcall CallHelpVideo()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpVideo();
	delete p;
}

void __stdcall CallHelpForum()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpForum();
	delete p;
}

void __stdcall CallHelpSoft()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpSoft();
	delete p;
}

void __stdcall CallHelpSupport()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpSupport();
	delete p;
}

void __stdcall CallHelpFeedback()
{
	if (!blCheckFileTemp()) return;
	CHelpRibbon *p = new CHelpRibbon;
	p->_HelpFeedback();
	delete p;
}

void __stdcall CallNewCheckVersion()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		CCheckVersion *ck = new CCheckVersion;
		int iCheckUp = ck->_CompareVersion(1);
		if (iCheckUp != 1)
		{
			int result = ff->_y(L"Bạn đang sử dụng phiên bản mới nhất. Bạn vẫn muốn cập nhật lại?", 0, 0, 0);
			if (result == 6) iCheckUp = 1;
		}

		if (iCheckUp == 1)
		{

			Base64Ex *bb = new Base64Ex;
			CRegistry *reg = new CRegistry;

			CString szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
			reg->CreateKey(szCreate);

			// Luôn hiển thị checkbox bỏ qua update (exe)
			reg->WriteChar(ff->_ConvertCstringToChars(DlgVisibleCheckbox),
				ff->_ConvertCstringToChars(L"1"));

			ck->_OpenFileUpdateGXD();

			delete reg;
			delete bb;
			delete ck;
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallNewInfoVersion()
{
	CallNewCheckVersion();
}


void __stdcall CallNewSupportGXD()
{
	CallHelpSupport();
}

/*******************************************************************************/

void __stdcall CallInsertRow()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		xl->EnableCancelKey = XlEnableCancelKey(FALSE);
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		DlgInsertRows p;
		p.DoModal();

		int num = p.numInsert;
		if (num > 0)
		{
			ff->_xlPutScreenUpdating(false);
			RangePtr PRS = pShActive->GetRange(
				pRgActive->GetItem(iRowActive, 1),
				pRgActive->GetItem(iRowActive + num - 1, 1));
			PRS->EntireRow->Insert(xlShiftDown);
			ff->_xlPutScreenUpdating(true);
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallAutoNumber()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szCopy = shCopy, szFile = shFile;
		if (sCodeActive == shCategory || sCodeActive == shFManager
			|| sCodeActive.Left(szCopy.GetLength()) == szCopy
			|| sCodeActive.Left(szFile.GetLength()) == szFile)
		{
			CString szSheet = shFManager;
			CString szName = L"MNF_STT";

			if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
			{
				szSheet = shCategory;
				szName = L"CT_STT";
			}

			CString bsGet = ff->_xlGetNameSheet(shFManager, 0);
			_WorksheetPtr psGet = xl->Sheets->GetItem((_bstr_t)bsGet);

			int irowStart = 0;
			int icotSTT = ff->_xlGetColumn(psGet, szName, 1);
			int irowTieude = ff->_xlGetRow(psGet, szName, 1);
			if (irowTieude > 0) irowStart = irowTieude + 1;

			if (icotSTT > 0 && irowStart > 0)
			{
				int dem = 0;
				CString szval = L"";

				ff->_xlPutScreenUpdating(false);

				int iRowLast = pShActive->Cells->SpecialCells(xlCellTypeLastCell)->GetRow() + 1;
				for (int i = irowStart; i <= iRowLast; i++)
				{
					szval = ff->_xlGIOR(pRgActive, i, icotSTT, L"");
					if (_wtoi(szval) > 0)
					{
						dem++;
						pRgActive->PutItem(i, icotSTT, dem);
					}
				}

				ff->_xlPutScreenUpdating(true);
			}
		}

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallRenameFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_RegistryCOMGpro();
		ff->_xlGetInfoSheetActive();
		ff->_DefaultSheetActive(shFManager);

		DlgRenameFiles p;
		p._RightClickRenameFiles(1);

		delete ff;

	}
	catch (...) {}
}

void __stdcall CallCopyFiles()
{
	Function *ff = new Function;
	ff->_msgupdate();
	delete ff;
}

void __stdcall CallMoveFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szFile = shFile;
		if (sCodeActive == shFManager || sCodeActive.Left(szFile.GetLength()) == szFile)
		{
			CMoveFiles *p = new CMoveFiles;
			p->_DichuyenDulieu();
			delete p;
		}
		else ff->_msgbox(MS_QLFILES, 0, 0);

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallOpenFile()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		// Hàm hỗ trợ mở file đặc biệt
		ff->_xlOpenFileEx(pShActive, iRowActive, iColumnActive);

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallOpenFolder()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		// Hàm hỗ trợ mở thư mục đặc biệt
		ff->_xlOpenFolderEx(pShActive, iRowActive, iColumnActive);

		delete ff;
	}
	catch (...) {}
}

void __stdcall CallDeleteFiles()
{
	try
	{
		if (!bCheckVersion) return;
		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlGetInfoSheetActive();

		CString szCopy = shCopy, szFile = shFile;
		if (sCodeActive == shCategory || sCodeActive == shFManager
			|| sCodeActive.Left(szCopy.GetLength()) == szCopy
			|| sCodeActive.Left(szFile.GetLength()) == szFile)
		{
			_WorksheetPtr pSheet;
			CString bs = L"", szCopy = shCopy;
			int icotFilegoc = 0, icotFileGXD = 0;
			if (sCodeActive == shCategory || sCodeActive.Left(szCopy.GetLength()) == szCopy)
			{
				bs = (LPCTSTR)ff->_xlGetNameSheet(shCategory, 1);
				pSheet = xl->Sheets->GetItem((_bstr_t)bs);

				icotFilegoc = ff->_xlGetColumn(pSheet, L"CT_FILEGOC", 0);
				icotFileGXD = ff->_xlGetColumn(pSheet, L"CT_FILEGXD", 0);
			}
			else
			{
				bs = (LPCTSTR)ff->_xlGetNameSheet(shFManager, 1);
				pSheet = xl->Sheets->GetItem((_bstr_t)bs);

				icotFilegoc = ff->_xlGetColumn(pSheet, L"MNF_LINKGOC", 0);
			}

			// Xác định số lượng dòng sẽ xóa
			vector<int> vecRow;
			int isize = 0, icheck = 0;
			int iResult = ff->_xlGetAllCellSelection(0, 0);
			for (int i = 0; i < iResult; i++)
			{
				icheck = 0;
				isize = (int)vecRow.size();
				if (isize > 0)
				{
					for (int j = 0; j < isize; j++)
					{
						if (vecSelect[i].row == vecRow[j])
						{
							icheck = 1;
							break;
						}
					}
				}

				if (icheck == 0) vecRow.push_back(vecSelect[i].row);
			}

			RangePtr PRS;
			vector<CString> vecHyper;
			int ncoutHyp = 0;

			isize = (int)vecRow.size();
			if (isize > 1)
			{
				CString szval = L"";
				szval.Format(L"Bạn đang lựa chọn (%d) files để xóa. "
					L"Fils sẽ bị xóa khỏi máy tính của bạn. Bạn có chắc chắn với điều này?", isize);
				iResult = ff->_y(szval, 0, 0, 0);
				if (iResult == 6)
				{
					ff->_xlPutScreenUpdating(false);
					ff->_QuickSortArr(vecRow, 0, isize - 1);
					for (int i = isize - 1; i >= 0; i--)
					{
						if (iColumnActive == icotFileGXD) PRS = pRgActive->GetItem(vecRow[i], icotFileGXD);
						else
						{
							if (icotFilegoc > 0) PRS = pRgActive->GetItem(vecRow[i], icotFilegoc);
						}

						ncoutHyp = ff->_xlGetAllHyperLink(PRS, vecHyper, 1, 0);
						if (ncoutHyp > 0)
						{
							PRS = pRgActive->GetItem(vecRow[i], 1);
							PRS->EntireRow->Delete(xlShiftUp);
							DeleteFile(vecHyper[0]);
						}
						vecHyper.clear();						
					}

					PRS = pRgActive->GetItem(iRowActive, iColumnActive);
					PRS->Select();

					ff->_xlPutScreenUpdating(true);

					CallAutoNumber();
				}								
			}
			else
			{
				if (iColumnActive == icotFileGXD) PRS = pRgActive->GetItem(iRowActive, icotFileGXD);
				else
				{
					if (icotFilegoc > 0) PRS = pRgActive->GetItem(iRowActive, icotFilegoc);
				}

				ncoutHyp = ff->_xlGetAllHyperLink(PRS, vecHyper, 1, 0);
				if (ncoutHyp > 0)
				{
					iResult = ff->_y(
						L"File được chọn sẽ bị xóa khỏi máy tính của bạn. "
						L"\nBạn có chắc chắn xóa file có đường dẫn: \n" + vecHyper[0], 0, 0, 0);
					if (iResult == 6)
					{
						ff->_xlPutScreenUpdating(false);

						PRS = pRgActive->GetItem(iRowActive, 1);
						PRS->EntireRow->Delete(xlShiftUp);
						DeleteFile(vecHyper[0]);

						PRS = pRgActive->GetItem(iRowActive, iColumnActive);
						PRS->Select();

						ff->_xlPutScreenUpdating(true);

						CallAutoNumber();
					}
				}
				vecHyper.clear();
			}
			vecRow.clear();
		}

		delete ff;
	}
	catch (...) {}
}

/**********************************************************/

void __stdcall EventEnter()
{

}

void __stdcall EventRightClick()
{

}

void __stdcall DLLStart()
{
	try
	{
		CoInitialize(NULL);
		xl.GetActiveObject(ExcelApp);
		xl->PutVisible(VARIANT_TRUE);
		pWb = xl->GetActiveWorkbook();

		if (!blCheckFileTemp()) return;

		Function *ff = new Function;
		ff->_xlDeveloperMacroSettings();

		//Tạm thời bỏ qua bCheckVersion --> Không check phiên bản
		bCheckVersion = true;
		if (bCheckVersion)
		{
			m_FontHeaderList.DeleteObject();
			m_FontHeaderList.CreateFont(16, 0, 0, 0, FW_BOLD, false, false, 0,
				ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
				DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, L"Arial");

			m_FontItalic.DeleteObject();
			m_FontItalic.CreateFont(16, 0, 0, 0, FW_NORMAL, true, true, 0,
				ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
				DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, L"Arial");

			Base64Ex *bb = new Base64Ex;
			CRegistry *reg = new CRegistry;
			CCheckVersion *ck = new CCheckVersion;

			// Gọi và luôn sử dụng Dll GHook keyboard + mouse trong suốt quá trình làm việc (bl = false)
			ff->_xlCallFunctionDLL(L"GHook.dll", L"DLLStartHook", false);

			// Ghi phiên bản Excel sử dụng
			CString szCreate = bb->BaseDecodeEx(CreateKeySettings);
			reg->CreateKey(szCreate);

			if (ff->_xlIsOperatingSystem64())
			{
				szOffice = L"64";
				reg->WriteChar("OperatingSystem", "64");
			}
			else
			{
				szOffice = L"32";
				reg->WriteChar("OperatingSystem", "32");
			}
			
			CString szval = ff->_xlGetVersionExcel();
			reg->WriteChar("OfficeVersion", ff->_ConvertCstringToChars(szval));

			szCreate = bb->BaseDecodeEx(CreateKeySettings) + UpldateSoft;
			reg->CreateKey(szCreate);

			int iCheckUpSoft = 1;
			szval = reg->ReadString(DlgCheckUpdateSoft, L"");
			if (szval != L"") iCheckUpSoft = _wtoi(szval.Trim());
			if (iCheckUpSoft != 1) iCheckUpSoft = 0;

			int nResult = 0;
			if (iCheckUpSoft == 1) nResult = ck->_CompareVersion(0);
			if (nResult == 1)
			{
				// Luôn ẩn checkbox bỏ qua update (exe)
				reg->WriteChar(ff->_ConvertCstringToChars(DlgVisibleCheckbox),
					ff->_ConvertCstringToChars(L"0"));

				ck->_OpenFileUpdateGXD();
			}
			else
			{
				// false = Khi khởi động phụ thuộc vào Registry/DlgAutoUpdate
				CallFuncUpdateFiles(false);
			}

			// Khởi động ứng dụng Smart Start
			ck->_OpenFileManagerApp();

			delete reg;
			delete bb;
			delete ck;
		}
		else
		{
			ff->_msgbox(L"Phiên bản dùng thử đã hết hạn sử dụng. "
				L"Liên hệ 1900.0147 để được hỗ trợ giải đáp.", 0, 0);
		}

		delete ff;
	}
	catch (...) {
		MessageBox(NULL, L"Có thể ứng dụng Excel chưa được tắt hoàn toàn. Bạn chuột phải vào thanh TaskBar "
			L"/ chọn Task Manager và tìm đến ứng dụng Microsoft Excel để tắt hoàn toàn (End task).", 
			L"Lỗi ứng dụng", MB_OK | MB_ICONERROR);
	}
}