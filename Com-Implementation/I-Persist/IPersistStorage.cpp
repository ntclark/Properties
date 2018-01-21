// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.


#include "Properties.h"

#include "utils.h"

   Properties::_IPersistStorage::_IPersistStorage(Properties *pp) : 
     pParent(pp),
     noScribble(false) { };
  
   Properties::_IPersistStorage::~_IPersistStorage() {};
  
   long __stdcall Properties::_IPersistStorage::QueryInterface(REFIID riid,void **ppv) {
   return pParent -> InternalQueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::_IPersistStorage::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IPersistStorage::Release() {
   return pParent -> InternalRelease(); 
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::GetClassID(CLSID *pcid) {

   if ( pParent -> pIPropertiesClient ) {
      GUID clsid;
      memset(&clsid,0,sizeof(GUID));
      pParent -> pIPropertiesClient -> GetClassID((BYTE*)&clsid);
      if ( memcmp(&clsid,"00000000",8) ) 
         memcpy(pcid,&clsid,sizeof(GUID));
      else {
         if ( pParent -> debuggingEnabled ) {
            char szError[1024];
            sprintf(szError,"In method IPersistStorage::GetClassID(). Your object implements IGPropertiesClient. However, the GetClassID() method did not specify a suitable GUID for the object.\n\nNor did the object call IGProperties::SetClassID().\n\nOne or the other is required for persistence.");
            MessageBox(NULL,szError,"GSystem Properties Component: debugging message",MB_OK);
         }
         return E_FAIL;
      }
      return S_OK;
   }

   CLSID clsidTest;
   memset(&clsidTest,0,sizeof(CLSID));
   if ( ! memcmp(&pParent -> objectCLSID,&clsidTest,sizeof(CLSID)) ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[1024];
         sprintf(szError,"In method IPersistStorage::GetClassID(). Your object does not implement IPropertyPageClient, nor did it call SetClassID on the IProperties Interface.\n\nThe Properties Component has no information to give the caller of IPersistStorage::GetClassID\n\nPersistence will probably fail.");
         MessageBox(NULL,szError,"GSystem Properties Component: debugging message",MB_OK);
      }
      return E_NOTIMPL;
   }
   *pcid = pParent -> objectCLSID;
   return S_OK;
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::IsDirty() {
   return pParent -> pIProperties -> IsDirty();
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::InitNew(IStorage *is) {

   if ( ! is ) return E_POINTER;

   STATSTG statSTG;
   HRESULT hr;

   pParent -> pIStorage_current = is;
   pParent -> pIStorage_current -> Stat(&statSTG,STATFLAG_DEFAULT);

   if ( S_OK != (hr = pParent -> pIStorage_current -> OpenStream(statSTG.pwcsName,0,STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,&pParent -> pIStream_current)) ) { 
      if ( S_OK != (hr = pParent -> pIStorage_current -> CreateStream(statSTG.pwcsName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,0,&pParent -> pIStream_current)) ) {
         if ( STG_E_INVALIDNAME == hr ) {
            char szName[] = "GSystemPropertiesStorage";
            BSTR bstrName = SysAllocStringLen(NULL,(DWORD)strlen(szName));
            MultiByteToWideChar(CP_ACP,0,szName,-1,bstrName,(DWORD)strlen(szName));
            if ( S_OK != (hr = pParent -> pIStorage_current -> CreateStream(bstrName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,0,&pParent -> pIStream_current)) ) {
               CoTaskMemFree(statSTG.pwcsName);
               SysFreeString(bstrName);
               pParent -> pIStorage_current = NULL;
               pParent -> pIStream_current = NULL;
               return hr;
            }
            SysFreeString(bstrName);
         }
      }
   }

   CoTaskMemFree(statSTG.pwcsName);

   pParent -> currentPersistenceMechanism = MECHANISM_STORAGE;

   hr = pParent -> pIProperties -> InitNew(pParent -> pIStorage_current);

   pParent -> pIStorage_current = NULL;
   pParent -> pIStream_current -> Release();
   pParent -> pIStream_current = NULL;

   return hr;
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::Load(IStorage *is) {

   if ( ! is ) return E_POINTER;

   pParent -> pIStorage_current = is;

   char szStreamName[128];
   if ( pParent -> pCurrent_IO_Object ) {
      sprintf(szStreamName,"GProperties%ld",pParent -> pCurrent_IO_Object -> currentItemIndex++);
   } else {
      sprintf(szStreamName,"GProperties");
   }

   BSTR streamName = SysAllocStringLen(NULL,(DWORD)strlen(szStreamName) + 1);
   MultiByteToWideChar(CP_ACP,0,szStreamName,-1,streamName,(DWORD)strlen(szStreamName) + 1);
   HRESULT hr = pParent -> pIStorage_current -> OpenStream(streamName,0,STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,&pParent -> pIStream_current);
   if ( ! SUCCEEDED(hr) ) 
      hr = pParent -> pIStorage_current -> OpenStream(streamName,0,STGM_READ | STGM_SHARE_EXCLUSIVE,0,&pParent -> pIStream_current);
   SysFreeString(streamName);

   if ( S_OK != hr ) {
      pParent -> pIStorage_current = NULL;
      pParent -> pIStream_current = NULL;
      return hr;
   }

   pParent -> currentPersistenceMechanism = MECHANISM_STORAGE;

   hr = pParent -> internalLoad(pParent -> pIStream_current,pParent -> pIStorage_current);

   pParent -> pIStorage_current = NULL;
   pParent -> pIStream_current -> Release();
   pParent -> pIStream_current = NULL;

   return hr;
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::Save(IStorage *is,BOOL sameAsLoad) {

   if ( noScribble ) return E_UNEXPECTED;

   if ( ! is ) return E_POINTER;

   pParent -> pIStorage_current = is;

   char szStreamName[128];
   if ( pParent -> pCurrent_IO_Object ) {
      sprintf(szStreamName,"GProperties%ld",pParent -> pCurrent_IO_Object -> currentItemIndex++);
   } else {
      sprintf(szStreamName,"GProperties");
   }

   BSTR streamName = SysAllocStringLen(NULL,(UINT)strlen(szStreamName) + 1);

   MultiByteToWideChar(CP_ACP,0,szStreamName,-1,streamName,(int)strlen(szStreamName) + 1);

   HRESULT hr = pParent -> pIStorage_current -> CreateStream(streamName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | 0*STGM_CREATE,0,0,&pParent -> pIStream_current);

   SysFreeString(streamName);

   noScribble = true;
   
   pParent -> currentPersistenceMechanism = MECHANISM_STORAGE;

   hr = pParent -> internalSave(pParent -> pIStream_current,pParent -> pIStorage_current);

   pParent -> pIStorage_current = NULL;

   if ( pParent -> pIStream_current )
      pParent -> pIStream_current -> Release();

   pParent -> pIStream_current = NULL;

   return hr;
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::SaveCompleted(IStorage*) {
   noScribble = false;
   return S_OK;
   }
 
 
   STDMETHODIMP Properties::_IPersistStorage::HandsOffStorage() {
   return S_OK;
   }