#include "pch.h"
#include "MultiFileOpenDialog.h"

// CMultiFileOpenDialog
CMultiFileOpenDialog::CMultiFileOpenDialog(BOOL bVistaStyle, 
                                           LPCTSTR lpszDefExt, 
                                           LPCTSTR lpszFileName,
                                           DWORD dwFlags, 
                                           LPCTSTR lpszFilter, 
                                           CWnd* pParentWnd) 

   : CFileDialog(TRUE, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd, 0, bVistaStyle),
     m_pFileBuff(NULL),
     m_bUseExtBuffer(TRUE)
{

}

CMultiFileOpenDialog::~CMultiFileOpenDialog()
{
   delete []m_pFileBuff;
}

DWORD CMultiFileOpenDialog::_CalcRequiredBuffSizeOldStyle()
{
   ASSERT(GetOFN().Flags & OFN_EXPLORER);
   CWnd* pParent = GetParent();
   ASSERT(NULL != pParent);

   // get required sizes for path and file names buffers
   const UINT nPathSize = (UINT)pParent->SendMessage(CDM_GETFOLDERPATH);
   const UINT nFileSize = (UINT)pParent->SendMessage(CDM_GETSPEC);
   return nPathSize + nFileSize + 1;
}

DWORD CMultiFileOpenDialog::_CalcRequiredBuffSize()
{
   DWORD dwRet = 0;

   if(m_bVistaStyle)
      dwRet = _CalcRequiredBuffSizeVistaStyle();
   else
      dwRet = _CalcRequiredBuffSizeOldStyle();
  
   return dwRet;
}

DWORD CMultiFileOpenDialog::_CalcRequiredBuffSizeVistaStyle()
{
   DWORD dwRet = 0;
   IFileOpenDialog* pFileOpen = GetIFileOpenDialog(); 
   ASSERT(NULL != pFileOpen);

   CComPtr<IShellItemArray> pIItemArray;
   HRESULT hr = pFileOpen->GetSelectedItems(&pIItemArray);
   DWORD dwItemCount = 0;
   if(SUCCEEDED(hr))
   {
      hr = pIItemArray->GetCount(&dwItemCount);
      if(SUCCEEDED(hr))
      {
         for(DWORD dwItem = 0; dwItem < dwItemCount; dwItem++)
         {
            CComPtr<IShellItem> pItem;
            hr = pIItemArray->GetItemAt(dwItem, &pItem);
            if(SUCCEEDED(hr))
            {
               LPWSTR pszName = NULL;
               if(dwItem == 0)
               {
                  // get full path and file name
                  hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszName);
               }
               else
               {
                  // get file name
                  hr = pItem->GetDisplayName(SIGDN_NORMALDISPLAY, &pszName);
               }
               if(SUCCEEDED(hr))
               {
                  dwRet += wcslen(pszName) + 1;
                  ::CoTaskMemFree(pszName);
               }
            }
         }
      }
   }
   dwRet += 1; // add an extra character
   // release the IFileOpenDialog pointer
   pFileOpen->Release();
   return dwRet;
}

void CMultiFileOpenDialog::_SetExtBuffer(DWORD dwReqBuffSize)
{
   try
   {
      GetOFN().nMaxFile = dwReqBuffSize;
      delete []m_pFileBuff;
      m_pFileBuff = new TCHAR[dwReqBuffSize]; 
      GetOFN().lpstrFile = m_pFileBuff;
   }
   catch(CException* e)
   {
      e->ReportError();
      e->Delete();
   }
}


BEGIN_MESSAGE_MAP(CMultiFileOpenDialog, CFileDialog)
END_MESSAGE_MAP()

// CMultiFileOpenDialog message handlers

void CMultiFileOpenDialog::OnFileNameChange()
{
   // NOTE: m_bUseExtBuffer is used only for demo/testing purpose;
   // it may be removed in a final release.
   if(m_bUseExtBuffer)
   {
      DWORD dwReqBuffSize = _CalcRequiredBuffSize();
      if(dwReqBuffSize > GetOFN().nMaxFile)
      {
         _SetExtBuffer(dwReqBuffSize);
      }
   }

   CFileDialog::OnFileNameChange();
}
