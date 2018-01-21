// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

#include "List.cpp"
#include "utils.h"


   Properties::_IProperties::_IProperties(Properties* pp) :
     pParent(pp) {}


   Properties::_IProperties::~_IProperties() {
   IGProperty *pProp;
   while ( pProp = pParent -> GetLast() ) {
      pProp -> Release();
      pParent -> Remove(pProp);
   }
   IGProperty *pIProperty = (IGProperty *)NULL;
   while ( pIProperty = pParent -> proxyList.GetNext(pIProperty) )
      pIProperty -> Release();
   return;
   }


   // IUnknown

   long __stdcall Properties::_IProperties::QueryInterface(REFIID riid,void **ppv) {
   if ( riid == IID_IUnknown )
      *ppv = static_cast<IUnknown*>(this);
   else
      if ( riid == IID_IGProperties )
         *ppv = static_cast<IGProperties*>(this);
      else
         return pParent -> InternalQueryInterface(riid,ppv);
   AddRef();
   return S_OK;
   }
   unsigned long __stdcall Properties::_IProperties::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IProperties::Release() { 
   return pParent -> InternalRelease(); 
   }
 

   // IDispatch

   STDMETHODIMP Properties::_IProperties::GetTypeInfoCount(UINT * pctinfo) { 
   *pctinfo = 1;
   return S_OK;
   } 


   long __stdcall Properties::_IProperties::GetTypeInfo(UINT itinfo,LCID lcid,ITypeInfo **pptinfo) { 
   *pptinfo = NULL; 
   if ( itinfo != 0 ) 
      return DISP_E_BADINDEX; 
   *pptinfo = pITypeInfo_IProperties;
   pITypeInfo_IProperties -> AddRef();
   return S_OK; 
   } 
 

   STDMETHODIMP Properties::_IProperties::GetIDsOfNames(REFIID riid,OLECHAR** rgszNames,UINT cNames,LCID lcid, DISPID* rgdispid) { 
   return DispGetIDsOfNames(pITypeInfo_IProperties,rgszNames,cNames,rgdispid);
   }


   STDMETHODIMP Properties::_IProperties::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, 
                                           WORD wFlags,DISPPARAMS FAR* pdispparams, VARIANT FAR* pvarResult,
                                           EXCEPINFO FAR* pexcepinfo, UINT FAR* puArgErr) { 
   return DispInvoke(this,pITypeInfo_IProperties,dispidMember,wFlags,pdispparams,pvarResult,pexcepinfo,puArgErr); 
   }


   // IProperties

   HRESULT Properties::_IProperties::put_DebuggingEnabled(VARIANT_BOOL setEnabled) {
   IGProperty* p = NULL;
   while ( p = pParent -> GetNext(p) ) p -> put_debuggingEnabled(setEnabled);
   pParent -> debuggingEnabled = setEnabled;
   return S_OK;
   }
   HRESULT Properties::_IProperties::get_DebuggingEnabled(VARIANT_BOOL *pGetEnabled) {
   *pGetEnabled = pParent -> debuggingEnabled;
   return S_OK;
   }

 
   HRESULT Properties::_IProperties::get_Property(BSTR theName,IGProperty** ppIProperty) {
   if ( ! ppIProperty ) return E_POINTER;
   *ppIProperty = pParent -> getProperty(theName,"IProperties::get_Property");
   if ( ! *ppIProperty ) return E_FAIL;
   (*ppIProperty) -> AddRef();
   return S_OK;
   }


   HRESULT Properties::_IProperties::get_Count(long* theCount) {
   *theCount = pParent -> Count() + pParent -> proxyList.Count();
   return S_OK;
   }


   HRESULT Properties::_IProperties::get_Size(long *getSize) {
   long deltaGetSize;
   *getSize = 0;
   IGProperty *pProp = pParent -> GetFirst();
   while ( pProp ) {
      pProp -> get_size(&deltaGetSize);
      *getSize += deltaGetSize;
      pProp = pParent -> GetNext(pProp);
   }
   IGProperty* pIProperty = pParent -> proxyList.GetFirst();
   while ( pIProperty ) {
      pIProperty -> get_size(&deltaGetSize);
      *getSize += deltaGetSize;
      pIProperty = pParent -> proxyList.GetNext(pIProperty);
   }
   return S_OK;
   }


   HRESULT Properties::_IProperties::get_IStorage(IStorage** pGetIStorage) {
   if ( ! pGetIStorage ) return E_POINTER;
   *pGetIStorage = pParent -> pIStorage_current;
   return S_OK;
   }


   HRESULT Properties::_IProperties::get_IStream(IStream** pGetIStream) {
   if ( ! pGetIStream ) return E_POINTER;
   *pGetIStream = pParent -> pIStream_current;
   return S_OK;
   }


   HRESULT Properties::_IProperties::Add(BSTR name,IGProperty** ppIProperty) {
   IGProperty* p = pParent -> getProperty(name,NULL);
   if ( (name && wcslen(name) != 0) && p ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         char szName[64];
         memset(szName,0,64);
         WideCharToMultiByte(CP_ACP,0,name,-1,szName,64,0,0);
         sprintf(szError,"An attempt to Add() a property was made when a property of the given name (%s) already exists",szName);
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   char szName[MAX_PROPERTY_SIZE];
   pParent -> pProperty_IClassFactory -> CreateInstance(NULL,IID_IGProperty,reinterpret_cast<void**>(&p));
   if ( name ) {
      WideCharToMultiByte(CP_ACP,0,name,-1,szName,MAX_PROPERTY_SIZE,0,0);
      p -> put_name(name);
   } else {
      memset(szName,0,sizeof(szName));
      sprintf(szName,"_Property_%ld",pParent -> Count() + 1);
      BSTR tempName = SysAllocStringLen(NULL,(DWORD)strlen(szName) + 1);
      MultiByteToWideChar(CP_ACP,0,szName,-1,tempName,(DWORD)strlen(szName) + 1);   
      p -> put_name(tempName);
      SysFreeString(tempName);
   }
   p -> put_type(TYPE_UNSPECIFIED);
   p -> put_size(SIZE_UNSPECIFIED);
   pParent -> Add(p,szName,pParent -> Count() + 1);
   p -> put_debuggingEnabled(pParent -> debuggingEnabled);
   if ( ppIProperty ) {
      p -> QueryInterface(IID_IDispatch,reinterpret_cast<void**>(ppIProperty));
      pParent -> propertiesToRelease.Add(reinterpret_cast<IGProperty*>(*ppIProperty));
   }
   return S_OK;
   }
 

   HRESULT Properties::_IProperties::Include(IGProperty *pIProperty) {
   if ( ! pIProperty ) return E_POINTER;
   pParent -> Add(pIProperty,NULL,pParent -> Count() + 1);
   pIProperty -> AddRef();
   pIProperty -> put_debuggingEnabled(pParent -> debuggingEnabled);
   return S_OK;
   }
 
 
   HRESULT Properties::_IProperties::Remove(BSTR name) {
   if ( ! name ) return E_POINTER;
   if ( ! *name ) return E_POINTER;
   IGProperty *p,*pFound = pParent -> getProperty(name,"IProperties::Remove");
   if ( ! pFound ) return E_FAIL;
   pParent -> List<IGProperty>::Remove(pFound);
   p = (IGProperty *)NULL;
   while ( p = pParent -> propertiesToRelease.GetNext(p) ) {
      if ( p == pFound ) {
         p -> Release();
         pParent -> propertiesToRelease.Remove(p);
         p = NULL;
      }
   }
   return S_OK;
   }
 

   HRESULT Properties::_IProperties::AddObject(IUnknown* pObj) {

   if ( ! pObj ) return E_POINTER;

   IPersist* pIPersist;
   long oldCount = pParent -> persistableObjectInterfaces.Count();

   if ( SUCCEEDED(pObj -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersist))) ) {
      pParent -> persistableObjectInterfaces.Add(new persistableObjectInterface(pObj,MECHANISM_STREAMINIT));
      pIPersist -> Release();
   }

   if ( SUCCEEDED(pObj -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersist))) ) {
      pParent -> persistableObjectInterfaces.Add(new persistableObjectInterface(pObj,MECHANISM_STREAM));
      pIPersist -> Release();
   }

   if ( SUCCEEDED(pObj -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersist))) ) {
      pParent -> persistableObjectInterfaces.Add(new persistableObjectInterface(pObj,MECHANISM_STORAGE));
      pIPersist -> Release();
   }

   if ( SUCCEEDED(pObj -> QueryInterface(IID_IPersistPropertyBag,reinterpret_cast<void**>(&pIPersist))) ) { 
      pParent -> persistableObjectInterfaces.Add(new persistableObjectInterface(pObj,MECHANISM_PROPERTYBAG));
      pIPersist -> Release();
   }

   if ( SUCCEEDED(pObj -> QueryInterface(IID_IPersistPropertyBag2,reinterpret_cast<void**>(&pIPersist))) ) {
      pParent -> persistableObjectInterfaces.Add(new persistableObjectInterface(pObj,MECHANISM_PROPERTYBAG2));
      pIPersist -> Release();
   }

   if ( oldCount == pParent -> persistableObjectInterfaces.Count() ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to persist on an object that doesn't support any of the required persistence interfaces:\n\nIPersistStorage, IPersistStream, IPersistStreamInit, or IPersistPropertyBag.");
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
      }
      return E_FAIL;
   }

   pParent -> persistableObjects.Add(new IUnknown*(pObj));

   return S_OK;
   }

 
   HRESULT Properties::_IProperties::RemoveObject(IUnknown* pObj) {

   if ( ! pObj ) return E_POINTER;

   persistableObjectInterface* p = NULL;
   while ( p = pParent -> persistableObjectInterfaces.GetNext(p) ) {
      if ( p -> pIUnknown == pObj ) {
         pParent -> persistableObjectInterfaces.Remove(p);
         delete p;
         p = NULL;
      }
   }

   IUnknown** ppIUnknown = NULL;
   while ( ppIUnknown = pParent -> persistableObjects.GetNext(ppIUnknown) ) {
      if ( *ppIUnknown == pObj ) {
         pParent -> persistableObjects.Remove(ppIUnknown);
         delete ppIUnknown;
         ppIUnknown = NULL;
      }
   }

   return S_OK;
   }

 
   HRESULT Properties::_IProperties::Advise(IGPropertiesClient *pir) {
   if ( pParent -> pIPropertiesClient ) 
      pParent -> pIPropertiesClient -> Release();
   pParent -> pIPropertiesClient = pir;
   if ( pParent -> pIPropertiesClient )
      pParent -> pIPropertiesClient -> AddRef();
   return S_OK;
   }
 
 
   HRESULT Properties::_IProperties::DirectAccess(BSTR bstrName,enum PropertyType propertyType,void* da,long das) {
   char szName[MAX_PROPERTY_SIZE];
   memset(szName,0,sizeof(szName));
   WideCharToMultiByte(CP_ACP,0,bstrName,-1,szName,MAX_PROPERTY_SIZE,0,0);
   IGProperty *pProp = pParent -> List<IGProperty>::Get(szName);
   if ( ! pProp ) return E_FAIL;
   IGProperty* pIProperty;
   pProp -> QueryInterface(IID_IGProperty,reinterpret_cast<void **>(&pIProperty));
   pIProperty -> directAccess(propertyType,da,das);
   pIProperty -> Release();
   return S_OK;
   }


   HRESULT Properties::_IProperties::SetClassID(BYTE * pclsid) {
   if ( ! pclsid ) return E_POINTER;
/*
   if ( pclsid -> vt != (VT_BYREF | VT_UI1) ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"A call was made to provide a ClassID to the GSystem Properties Component but the type of the variable passed was not correct.\n\nPass a pointer to 128 bytes (VT_BYREF | VT_UI1).");
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
      }
      return E_FAIL;
   }
*/
   memcpy(&pParent -> objectCLSID,pclsid,sizeof(GUID));
   return S_OK;
   }


   HRESULT Properties::_IProperties::CopyTo(IGProperties* pTheDestination) {
   
   char szName[MAX_PROPERTY_SIZE];
   BSTR bstrName = SysAllocStringLen(NULL,MAX_PROPERTY_SIZE);

   IGProperty* pProp;
   IGProperty* pIProperty_This;
   IGProperty** pIProperty_Destination;
   long cntOthers;

   pTheDestination -> GetPropertyInterfaces(&cntOthers,&pIProperty_Destination);
   for ( int k = 0; k < cntOthers; k++ ) {
      pIProperty_Destination[k] -> get_name(&bstrName);
      memset(szName,0,sizeof(szName));
      WideCharToMultiByte(CP_ACP,0,bstrName,-1,szName,MAX_PROPERTY_SIZE,0,0);
      if ( pProp = pParent -> Get(szName) ) {
         pProp -> QueryInterface(IID_IGProperty,reinterpret_cast<void**>(&pIProperty_This));
         pIProperty_This -> copyTo(pIProperty_Destination[k]);
         pIProperty_This -> Release();
      }
      pIProperty_Destination[k] -> Release();
   }

   CoTaskMemFree(pIProperty_Destination);

#if 0
   IGPropertiesClient *pIGPropertiesClient;
   pTheDestination -> QueryInterface(IID_IGPropertiesClient,reinterpret_cast<void **>(&pIGPropertiesClient));
   if ( pIGPropertiesClient ) {
      pIGPropertiesClient -> Loaded();
      pIGPropertiesClient -> Release();
   }
#endif

   return S_OK;
   }


   HRESULT Properties::_IProperties::GetPropertyInterfaces(long* pCntInterfaces,IGProperty*** pTheArray) {
   *pCntInterfaces = pParent -> Count();
   IGProperty** pp = *pTheArray = reinterpret_cast<IGProperty**>(CoTaskMemAlloc(*pCntInterfaces * sizeof(IGProperty*)));
   IGProperty* p = (IGProperty*)NULL;
   while ( p = pParent -> GetNext(p) )
      p -> QueryInterface(IID_IGProperty,reinterpret_cast<void**>(pp++));
   return S_OK;
   }


// Persistence support

   HRESULT Properties::_IProperties::PutHWNDPersistence(HWND hwndPers,HWND hwndInit,HWND hwndLoad,HWND hwndSavePrep) {
   HWND hwndPersistence = hwndPers;
   if ( ! hwndPersistence ) return E_POINTER;
   simplePersistenceWindow* pspw = new simplePersistenceWindow();
   pParent -> simplePersistenceWindows.Add(pspw);
   pspw -> hwnd = hwndPersistence;
   pspw -> hwndInit = hwndInit;
   pspw -> idInit = pspw -> hwndInit ? (short)GetWindowLongPtr(pspw -> hwndInit,GWLP_ID) : 0;
   pspw -> hwndLoad = hwndLoad;
   pspw -> idLoad = pspw -> hwndLoad ? (short)GetWindowLongPtr(pspw -> hwndLoad,GWLP_ID) : 0;
   pspw -> hwndSavePrep = hwndSavePrep;
   pspw -> idSavePrep = pspw -> hwndSavePrep ? (short)GetWindowLongPtr(pspw -> hwndSavePrep,GWLP_ID) : 0;
   return S_OK;
   }


   HRESULT Properties::_IProperties::RemoveHWNDPersistence(HWND hwndPers) {
   HWND hwndPersistence = hwndPers;
   if ( ! hwndPersistence ) return E_POINTER;
   simplePersistenceWindow* pspw = NULL;
   while ( pspw = pParent -> simplePersistenceWindows.GetNext(pspw) ) {
      if ( pspw -> hwnd == hwndPersistence ) {
         pParent -> simplePersistenceWindows.Remove(pspw);
         delete pspw;
         return S_OK;
      }
   }
   if ( pParent -> debuggingEnabled ) {
      char szErrorMessage[1024];
      sprintf(szErrorMessage,"IProperties::RemoveHWNDPersistence. The passed in window handle was not recognized as a window handle that was previously passed to PutHWNDPersitence.");
      MessageBox(NULL,szErrorMessage,"Non-Fatal warning",MB_OK);
   }
   return S_OK;
   }


   HRESULT Properties::_IProperties::IsDirty() {

   if ( pParent -> isDirty ) return S_OK;

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = pParent -> simplePersistenceWindows.GetNext(pspw) ) 
      if ( pspw -> hwnd ) 
         if ( pspw -> idIsDirty )
            if ( 0L == SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idIsDirty,0),(LPARAM)pspw -> hwndIsDirty) ) 
               return S_OK;

   if ( pParent -> pIPropertiesClient )
      if ( S_OK == pParent -> pIPropertiesClient -> IsDirty() ) 
         return S_OK;

   IGProperty *pProp = pParent -> GetFirst();
   while ( pProp ) {
      pProp -> get_isDirty(&pParent -> isDirty);
      if ( pParent -> isDirty ) return S_OK;
      pProp = pParent -> GetNext(pProp);
   }

   return S_FALSE;
   }
 
 
   HRESULT Properties::_IProperties::InitNew(IStorage* pIStorage) {
   return pParent -> internalInitNew(pIStorage);
   }


   HRESULT Properties::_IProperties::SaveToStorage(IStream* pIStream,IStorage* pIStorage) {
   pParent -> currentPersistenceMechanism = MECHANISM_STORAGE;
   return pParent -> internalSave(pIStream,pIStorage);
   }
 
 
   HRESULT Properties::_IProperties::LoadFromStorage(IErrorLog* pIErrorLog,IStream* pIStream,IStorage* pIStorage) {
   pParent -> currentPersistenceMechanism = MECHANISM_STORAGE;
   return pParent -> internalLoad(pIStream,pIStorage);
   }
 
 
   HRESULT Properties::_IProperties::SaveObjectToFile(IUnknown* pObj,BSTR bstrFileName) {

   if ( ! pObj ) 
      return E_POINTER;

   IPersistStorage* pIPersistStorage;

   HRESULT hr = pObj -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
   if ( ! SUCCEEDED(hr) ) {
      if ( pParent -> debuggingEnabled ) {
         char szErrorMessage[256];
         sprintf(szErrorMessage,"IProperties::SaveObjectToFile. The passed object did not support the IPersistStorage interface");
         MessageBox(NULL,szErrorMessage,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return hr;
   }
   
   IStorage* pIStorage;

   hr = StgCreateDocfile(bstrFileName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_FAILIFTHERE,0,&pIStorage);
   if ( STG_E_FILEALREADYEXISTS == hr ) {
      DeleteFileW(bstrFileName);
      hr = StgCreateDocfile(bstrFileName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_FAILIFTHERE,0,&pIStorage);
   }
   if ( S_OK != hr ) {
      if ( pParent -> debuggingEnabled ) {
         char szErrorMessage[256];
         sprintf(szErrorMessage,"IProperties::SaveObjectToFile. The system was not able to create the storage.");
         MessageBox(NULL,szErrorMessage,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return hr;
   }

   pIPersistStorage -> Save(pIStorage,FALSE);
   pIPersistStorage -> SaveCompleted(NULL);

   pIStorage -> Release();

   pIPersistStorage -> Release();

   return S_OK;
   }


   HRESULT Properties::_IProperties::LoadObjectFromFile(IUnknown* pObj,BSTR bstrFileName) {

   if ( ! pObj ) return E_POINTER;

   IPersistStorage* pIPersistStorage;

   HRESULT hr = pObj -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
   if ( ! SUCCEEDED(hr) ) {
      if ( pParent -> debuggingEnabled ) {
         char szErrorMessage[256];
         sprintf(szErrorMessage,"IProperties::LoadObjectFromFile. The passed object did not support the IPersistStorage interface");
         MessageBox(NULL,szErrorMessage,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return hr;
   }

   IStorage* pIStorage;

   hr = StgOpenStorage(bstrFileName,NULL,STGM_READ | STGM_SHARE_DENY_WRITE,NULL,0,&pIStorage);
   if ( ! SUCCEEDED(hr) ) {
      pIPersistStorage -> Release();
      if ( pParent -> debuggingEnabled ) {
         char szErrorMessage[256];
         sprintf(szErrorMessage,"IProperties::LoadObjectFromFile. The system was not able to open an existing storage with the given file name.");
         MessageBox(NULL,szErrorMessage,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return hr;
   }

   pIPersistStorage -> Load(pIStorage);

   pIStorage -> Release();
   pIPersistStorage -> Release();

   return S_OK;
   }


   // Window Contents Support

   HRESULT Properties::_IProperties::GetWindowValue(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowValue");
   if ( pProperty )
      return pProperty -> getWindowValue(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemValue(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowValue(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::GetWindowText(BSTR bstrPropertyName,HWND hwndControl) {
   return GetWindowValue(bstrPropertyName,hwndControl);
   }


   HRESULT Properties::_IProperties::GetWindowItemText(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowItemValue(bstrPropertyName,hwndDialog,idControl);
   }


   HRESULT Properties::_IProperties::SetWindowText(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowText");
   if ( pProperty )
      return pProperty -> setWindowText(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemText(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowText(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   // Combo-boxes

   HRESULT Properties::_IProperties::GetWindowComboBoxSelection(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowComboBoxSelection");
   if ( pProperty)
      return pProperty -> getWindowComboBoxSelection(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemComboBoxSelection(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowComboBoxSelection(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::SetWindowComboBoxSelection(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowComboBoxSelection");
   if ( pProperty )
      return pProperty -> setWindowComboBoxSelection(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemComboBoxSelection(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowComboBoxSelection(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::SetWindowComboBoxList(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowComboBoxList");
   if ( pProperty )
      return pProperty -> setWindowComboBoxList(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemComboBoxList(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowComboBoxList(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::GetWindowComboBoxList(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowComboBoxList");
   if ( pProperty )
      return pProperty -> getWindowComboBoxList(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemComboBoxList(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowComboBoxList(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   // List-boxes

   HRESULT Properties::_IProperties::GetWindowListBoxSelection(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowListBoxSelection");
   if ( pProperty)
      return pProperty -> getWindowListBoxSelection(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemListBoxSelection(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowListBoxSelection(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::SetWindowListBoxSelection(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowListBoxSelection");
   if ( pProperty )
      return pProperty -> setWindowListBoxSelection(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemListBoxSelection(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowListBoxSelection(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::SetWindowListBoxList(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowListBoxList");
   if ( pProperty )
      return pProperty -> setWindowListBoxList(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemListBoxList(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowListBoxList(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::GetWindowListBoxList(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowListBoxList");
   if ( pProperty )
      return pProperty -> getWindowListBoxList(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemListBoxList(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowListBoxList(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }




   HRESULT Properties::_IProperties::GetWindowArrayValues(BSTR bstrPropertyName,SAFEARRAY** ppsaHwndDialogs) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowArrayValue");
   if ( pProperty )
      return pProperty -> getWindowArrayValues(ppsaHwndDialogs);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemArrayValues(BSTR bstrPropertyName,SAFEARRAY** ppsaHwndDialogs,SAFEARRAY** ppsaIdControls) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowItemArrayValue");
   if ( pProperty )
      return pProperty -> getWindowItemArrayValues(ppsaHwndDialogs,ppsaIdControls);
   return S_OK;
   }


   HRESULT Properties::_IProperties::SetWindowArrayValues(BSTR bstrPropertyName,SAFEARRAY** ppsaHwndControls) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowArrayValues");
   if ( pProperty )
      return pProperty -> setWindowArrayValues(ppsaHwndControls);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemArrayValues(BSTR bstrPropertyName,SAFEARRAY** ppsaHwndDialogs,SAFEARRAY** ppsaIdControls) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowItemArrayValues");
   if ( pProperty )
      return pProperty -> setWindowItemArrayValues(ppsaHwndDialogs,ppsaIdControls);
   return E_FAIL;
   }


   // Check boxes

   HRESULT Properties::_IProperties::GetWindowChecked(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::GetWindowChecked");
   if ( pProperty) 
      return pProperty -> getWindowChecked(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::GetWindowItemChecked(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return GetWindowChecked(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Properties::_IProperties::SetWindowChecked(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowChecked");
   if ( pProperty )
      return pProperty -> setWindowChecked(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemChecked(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowChecked(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   // Enabled/Disabled

   HRESULT Properties::_IProperties::SetWindowEnabled(BSTR bstrPropertyName,HWND hwndControl) {
   IGProperty* pProperty = pParent -> getProperty(bstrPropertyName,"IProperties::SetWindowChecked");
   if ( pProperty )
      return pProperty -> setWindowEnabled(hwndControl);
   return E_FAIL;
   }


   HRESULT Properties::_IProperties::SetWindowItemEnabled(BSTR bstrPropertyName,HWND hwndDialog,long idControl) {
   return SetWindowEnabled(bstrPropertyName,GetDlgItem(hwndDialog,idControl));
   }


   // Other Windows API Support

   HRESULT Properties::_IProperties::GetWindowID(HWND hwndControl,long* pID) {
   if ( ! pID ) return E_POINTER;
   *pID = (long)GetWindowLongPtr(hwndControl,GWLP_ID);
   return S_OK;
   }