/*

                       Copyright (c) 1999,2000,2001,2002 Nathan T. Clark

*/

#include <windows.h>

#include "resource.h"

//#include "List.cpp"

#include "Properties.h"
#include "utils.h"

   // IProperties File Properties

   HRESULT Properties::_IProperties::put_FileName(BSTR newFileName) {

   if ( pParent -> temporaryFileInUse )
      pParent -> deleteTemporaryStorage();

   if ( pParent -> fileName ) 
      SysFreeString(pParent -> fileName);

   pParent -> fileName = SysAllocString(newFileName);

   return S_OK;
   }


   HRESULT Properties::_IProperties::get_FileName(BSTR* pGetFileName) {
   if ( ! pGetFileName ) return E_POINTER;
   *pGetFileName = SysAllocString(pParent -> fileName);
   return S_OK;
   }

   HRESULT Properties::_IProperties::put_FileAllowedExtensions(BSTR newFileExtensions) {
   if ( pParent -> fileExtensions ) SysFreeString(pParent -> fileExtensions);
   pParent -> fileExtensions = SysAllocString(newFileExtensions);
   return S_OK;
   }

   HRESULT Properties::_IProperties::get_FileAllowedExtensions(BSTR* pGetFileExtensions) {
   if ( ! pGetFileExtensions ) return E_POINTER;
   *pGetFileExtensions = SysAllocString(pParent -> fileExtensions);
   return S_OK;
   }


   HRESULT Properties::_IProperties::put_FileType(BSTR newFileType) {
   if ( pParent -> fileType ) SysFreeString(pParent -> fileType);
   pParent -> fileType = SysAllocString(newFileType);
   return S_OK;
   }

   HRESULT Properties::_IProperties::get_FileType(BSTR* pGetFileType) {
   if ( ! pGetFileType ) return E_POINTER;
   *pGetFileType = SysAllocString(pParent -> fileType);
   return S_OK;
   }


   HRESULT Properties::_IProperties::put_FileSaveOpenText(BSTR newFileSaveOpenText) {
   if ( pParent -> fileSaveOpenText ) SysFreeString(pParent -> fileSaveOpenText);
   pParent -> fileSaveOpenText = SysAllocString(newFileSaveOpenText);
   return S_OK;
   }

   HRESULT Properties::_IProperties::get_FileSaveOpenText(BSTR* pGetFileSaveOpenText) {
   if ( ! pGetFileSaveOpenText ) return E_POINTER;
   *pGetFileSaveOpenText = SysAllocString(pParent -> fileSaveOpenText);
   return S_OK;
   }


   // IProperties File Actions

   HRESULT Properties::_IProperties::New() {

   HRESULT hr;
   IStorage* pIStorage;
   IStream* pIStream;

   if ( pParent -> temporaryFileInUse )
      pParent -> deleteTemporaryStorage();

   pParent -> createTemporaryStorageFileName();

   if ( S_OK != (hr = pParent -> createStorageAndStream(&pIStorage,&pIStream,L"G-Properties")) )
      return hr;

   InitNew(pIStorage);

   pIStream -> Release();
   pIStorage -> Release();
   pParent -> pCurrent_IO_Object = NULL;
   
   pParent -> isDirty = true;

   pParent -> pCurrent_IO_Object = NULL;

   return S_OK;
   }


   HRESULT Properties::_IProperties::Open(BSTR* pOpenedFileName) {

   if ( ! pOpenedFileName ) 
      return E_POINTER;

   pParent -> prepOpenFileName();

   pParent -> openFileName.Flags |= OFN_FILEMUSTEXIST;
   pParent -> openFileName.Flags |= OFN_PATHMUSTEXIST;

   if ( ! GetOpenFileName(&pParent -> openFileName) ) {
      *pOpenedFileName = SysAllocStringLen(L"",1);
      return S_OK;
   }

   if ( pParent -> parseOpenFileName(true) ) {
      *pOpenedFileName = SysAllocString(pParent -> fileName);
      return OpenFile(*pOpenedFileName);
   }
   else
      *pOpenedFileName = SysAllocStringLen(L"",1);

   return S_OK;
   }


   HRESULT Properties::_IProperties::LoadFile(VARIANT_BOOL* wasSuccessful) {

   if ( ! wasSuccessful ) return E_POINTER;

   *wasSuccessful = FALSE;

   pParent -> prepOpenFileName();

   if ( * pParent -> openFileName.lpstrFile ) {
      WIN32_FIND_DATA findData;
      memset(&findData,0,sizeof(WIN32_FIND_DATA));
      HANDLE h = FindFirstFile(pParent -> openFileName.lpstrFile,&findData);
      if ( INVALID_HANDLE_VALUE != h ) {
         FindClose(h);
         BSTR bstrTemp = SysAllocStringLen(NULL,(UINT)strlen(pParent -> openFileName.lpstrFile) + 1);
         MultiByteToWideChar(CP_ACP,0,pParent -> openFileName.lpstrFile,-1,bstrTemp,(int)strlen(pParent -> openFileName.lpstrFile) + 1);
         HRESULT hr = OpenFile(bstrTemp);
         SysFreeString(bstrTemp);
         *wasSuccessful = (hr == S_OK);
         return S_OK;
      }         
   }

   return S_OK;
   }

   
   HRESULT Properties::_IProperties::OpenFile(BSTR openFileName) {

   if ( ! pParent -> setExistingFile(openFileName) ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PATH];
         char szFileName[MAX_PATH];
         memset(szFileName,0,sizeof(szFileName));
         WideCharToMultiByte(CP_ACP,0,openFileName,-1,szFileName,MAX_PATH,0,0);
         sprintf(szError,"An attempt was made to open a file that does not exist (Properties::OpenFile()).\n\nFile '%s' does not exist.",szFileName);
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return E_FAIL;
   }

   if ( pParent -> isDirty ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PATH];
         sprintf(szError,"Program logic allowed the opening of a properties file without having saved the prior set of properties.\n\n(IProperties::OpenFile(<fileName>))");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
   }

   HRESULT hr;
   IStorage* pIStorage;
   IStream* pIStream;

   if ( S_OK != (hr = pParent -> openStorageAndStream(&pIStorage,&pIStream,L"G-Properties")) ) {
      if ( hr == STG_E_FILENOTFOUND ) {
         if ( pParent -> debuggingEnabled ) {
            char szError[MAX_PATH];
            char szFileName[MAX_PATH];
            memset(szFileName,0,sizeof(szFileName));
            WideCharToMultiByte(CP_ACP,0,openFileName,-1,szFileName,MAX_PATH,0,0);
            sprintf(szError,"An attempt was made to open a file that either does not exist or is not a file saved by the GProperties Component\n\n(Properties::OpenFile(%s)).",szFileName);
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
      }
      return E_FAIL;
   }

   HRESULT rc = LoadFromStorage(NULL,pIStream,pIStorage);

   pIStream -> Release();
   pIStorage -> Release();
   pParent -> pCurrent_IO_Object = NULL;

   return rc;
   }


   HRESULT Properties::_IProperties::Save() {

   char szFileName[MAX_PATH];

   memset(szFileName,0,sizeof(szFileName));

   if ( pParent -> fileName )
      WideCharToMultiByte(CP_ACP,0,pParent -> fileName,-1,szFileName,MAX_PATH,0,0);

   if ( ! szFileName[0] || ! pParent -> fileName || pParent -> temporaryFileInUse ) {

      if ( pParent -> fileName && 0 == wcslen(pParent -> fileName) )
         return S_OK;

      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PATH];
         sprintf(szError,"An attempt was made to save properties but no file name has been supplied yet (IProperties::Save())");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      if ( pParent -> temporaryFileInUse ) {
         SysFreeString(pParent -> fileName);
         pParent -> fileName = NULL;
      }
      BSTR bstrDummy;
      HRESULT hr = SaveAs(&bstrDummy);
      SysFreeString(bstrDummy);
      return hr;
   }

   HRESULT hr;
   IStorage* pIStorage;
   IStream* pIStream;

   if ( S_OK != (hr = pParent -> createStorageAndStream(&pIStorage,&pIStream,L"G-Properties")) )
      return hr;

   SaveToStorage(pIStream,pIStorage);

   pIStream -> Release();
   pIStorage -> Release();
   pParent -> pCurrent_IO_Object = NULL;
        
   return S_OK;
   }


   HRESULT Properties::_IProperties::SaveTo(BSTR fileName) {   
   put_FileName(fileName);
   return Save();
   }


   HRESULT Properties::_IProperties::SaveAs(BSTR* pSavedFileName) {

   if ( ! pSavedFileName ) return E_POINTER;

   pParent -> prepOpenFileName();

   if ( GetSaveFileName(&pParent -> openFileName) ) {
      pParent -> parseOpenFileName(true);
      *pSavedFileName = SysAllocString(pParent -> fileName);
      return Save();
   }
   else
      *pSavedFileName = SysAllocStringLen(NULL,0);

   return S_OK;
   }
