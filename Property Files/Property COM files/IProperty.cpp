/*

                       Copyright (c) 1996,1997,1998,1999,2000,2001,2008 Nathan T. Clark

*/

#include <windows.h>
#include <ctype.h>
#include <stdio.h>

#include "Property.h"
#include "utils.h"
#include "list.cpp"

   static BYTE storageObjectSignature[] = {0xFA,0xFB,0xFC,0xFD};

   long __stdcall Property::QueryInterface(REFIID riid,void **ppv) {
   *ppv = NULL; 
   Property *p = static_cast<Property *>(this);
   if ( riid == IID_IUnknown || riid == IID_IGProperty || riid == IID_IDispatch ) 
      *ppv = p;
   else
      return E_NOINTERFACE;
   p -> AddRef(); 
   return S_OK; 
   }
   unsigned long __stdcall Property::AddRef() { return 1; }
   unsigned long __stdcall Property::Release() { return 1; }
 
 
   // IDispatch

   STDMETHODIMP Property::GetTypeInfoCount(UINT * pctinfo) { 
   *pctinfo = 1;
   return S_OK;
   } 


   long __stdcall Property::GetTypeInfo(UINT itinfo,LCID lcid,ITypeInfo **pptinfo) { 
   *pptinfo = NULL; 
   if ( itinfo != 0 ) 
      return DISP_E_BADINDEX; 
   *pptinfo = pITypeInfo_IProperty;
   pITypeInfo_IProperty -> AddRef();
   return S_OK; 
   } 
 

   STDMETHODIMP Property::GetIDsOfNames(REFIID riid,OLECHAR** rgszNames,UINT cNames,LCID lcid, DISPID* rgdispid) { 
   return DispGetIDsOfNames(pITypeInfo_IProperty,rgszNames,cNames,rgdispid);
   }


   STDMETHODIMP Property::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, 
                                           WORD wFlags,DISPPARAMS FAR* pdispparams, VARIANT FAR* pvarResult,
                                           EXCEPINFO FAR* pexcepinfo, UINT FAR* puArgErr) { 
   return DispInvoke(this,pITypeInfo_IProperty,dispidMember,wFlags,pdispparams,pvarResult,pexcepinfo,puArgErr); 
   }


   // IGProperty

   // Miscellanous Attributes

   HRESULT Property::put_name(BSTR bstrName) {
   SZSTRING_FROM_STRING(name,sizeof(name),bstrName)
   return S_OK;
   }

   HRESULT Property::get_name(BSTR* bstrName) {
   if ( ! bstrName ) return E_POINTER;
   *bstrName = SysAllocStringLen(NULL,(UINT)strlen(name) + 1);
   MultiByteToWideChar(CP_ACP,0,name,-1,*bstrName,(int)strlen(name) + 1);
   return S_OK;
   }


   HRESULT Property::put_type(PropertyType newType) {
   type = newType;
   if ( TYPE_OBJECT_STORAGE_ARRAY == type ) 
      everAssigned = true;
   return S_OK;
   }

   HRESULT Property::get_type(PropertyType* getType) {
   *getType = type;
   return S_OK;
   }


   HRESULT Property::put_size(long size) {
   if ( v.binarySize )
      delete [] v.binaryValue;
   v.binarySize = size;
   if ( v.binarySize ) {
      v.binaryValue = new BYTE[v.binarySize];
      memset(v.binaryValue,0,v.binarySize);
   }
   updateFromDirectAccess(); 
   return S_OK;
   }


   HRESULT Property::get_size(long *getSize) {
   updateFromDirectAccess();
   switch ( type ) {
   case TYPE_BOOL:
      *getSize = SIZE_BOOL;
      break;
   case TYPE_LONG:
      *getSize = SIZE_LONG;
      break;
   case TYPE_DOUBLE:
      *getSize = SIZE_DOUBLE;
      break;
   case TYPE_SZSTRING: {
      char *p = reinterpret_cast<char*>(propertyBinaryValue());
      if ( p )
         *getSize = (DWORD)strlen(p) + 1;
      else
         *getSize = 0;
      }
      break;
   case TYPE_STRING:
// 10-05-2002: Would trap if bstrValue was never set, which would be the case if the property was never given a value.
//             Create an empty string if never assigned.
      if ( ! bstrValue ) 
         bstrValue = SysAllocStringLen(L"",1);
      *getSize = 2 * ((long)wcslen(bstrValue) + 1);
      break;

   case TYPE_OBJECT_STORAGE_ARRAY: // <--- 09/21/07 ??
   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
      *getSize = v.binarySize;
      break;

   case TYPE_ARRAY: {

      BSTR bstrSize;
      toEncodedText(&bstrSize);
      *getSize = 2 * (DWORD)wcslen(bstrSize);
      SysFreeString(bstrSize);

      }
      break;

   default:
      *getSize = 0;
      break;
   }
   return S_OK;
   }
 


   HRESULT Property::put_isDirty(short setDirty) {
   isDirty = setDirty ? 1 : 0;
   return S_OK;
   }
   HRESULT Property::get_isDirty(short *getDirty) {
   *getDirty = isDirty;
   if ( directAccessAddress ) {
      if ( memcmp(v.binaryValue,directAccessAddress,min(v.binarySize,directAccessSize)) )
         *getDirty = true;
   }
   return S_OK;
   }


   HRESULT Property::get_everAssigned(short *getEverAssigned) {
   if ( directAccessAddress ) 
      *getEverAssigned = true;
   else
      *getEverAssigned = everAssigned;
   return S_OK;
   }

 
   HRESULT Property::put_ignoreSetAction(VARIANT_BOOL setIgnore) {
   ignoreSetAction = setIgnore ? TRUE : FALSE;
   return S_OK;
   }
   HRESULT Property::get_ignoreSetAction(VARIANT_BOOL *getIgnoreSetAction) {
   *getIgnoreSetAction = ignoreSetAction;
   return S_OK;
   }
 

   HRESULT Property::put_debuggingEnabled(VARIANT_BOOL setEnabled) {
   debuggingEnabled = setEnabled;
   return S_OK;
   }


   HRESULT Property::get_debuggingEnabled(VARIANT_BOOL *pGetEnabled) {
   *pGetEnabled = debuggingEnabled;
   return S_OK;
   }

 
   HRESULT Property::get_binaryData(BYTE **getAddress) {
   updateFromDirectAccess();
   switch ( type ) {
   case TYPE_STRING:

      if ( v.binaryValue ) 
         delete [] v.binaryValue;

// 10/05/2002: wcslen traps if bstrValue has never been set.
//             Therefore, set value to "" string.
      if ( ! bstrValue ) bstrValue = SysAllocStringLen(L"",1);

      v.binarySize = 2 * (DWORD)(wcslen(bstrValue) + 1);
      
      v.binaryValue = new BYTE[v.binarySize];
      memset(v.binaryValue,0,v.binarySize);
      WideCharToMultiByte(CP_ACP,0,bstrValue,v.binarySize,(char*)v.binaryValue,v.binarySize,0,0);

      *getAddress = reinterpret_cast<BYTE*>(v.binaryValue);

      break;

   case TYPE_OBJECT_STORAGE_ARRAY: // <-- 09/21/07
   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
      *getAddress = v.binaryValue;
      break;

   case TYPE_SZSTRING:
      *getAddress = v.binaryValue;
      break;

   case TYPE_ARRAY: {

      if ( v.binaryValue ) 
         delete [] v.binaryValue;

      BSTR bstrEncodedText;
      toEncodedText(&bstrEncodedText);
      v.binarySize = (long)wcslen(bstrEncodedText);
      v.binaryValue = new BYTE[v.binarySize];
      WideCharToMultiByte(CP_ACP,0,bstrEncodedText,v.binarySize,(char*)v.binaryValue,v.binarySize,0,0);
      SysFreeString(bstrEncodedText);

      *getAddress = v.binaryValue;

      }
      break;

   default:
      *getAddress = reinterpret_cast<BYTE *>(&v.scalar);
      break;
   }
   return S_OK;
   }
 


   HRESULT Property::get_compressedName(BSTR* theCompressedName) {
   if ( ! theCompressedName ) return E_POINTER;
   long n;
   bool nextUpCase = false;
   char *c,*psz,*result = new char[n = (DWORD)strlen(name) + 1];
   psz = result;
   memset(psz,0,n);
   c = name;
   while ( *c ) {
      if ( (*c == ' ' || ! isalpha(*c)) && *c != '_' && *c != ':' && ! ( *c >= '0' && *c <= '9') ) {
         c++;
         nextUpCase = true;
         continue;
      }
      if ( nextUpCase ) 
         *psz = toupper(*c);
      else
         *psz = *c;
      nextUpCase = false;
      c++;
      psz++;
   }
   *theCompressedName = SysAllocStringLen(NULL,(DWORD)strlen(result) + 1);
   MultiByteToWideChar(CP_ACP,0,result,-1,*theCompressedName,(DWORD)strlen(result) + 1);
   return S_OK;
   }

 
   HRESULT Property::put_encodedText(BSTR theText) {
   if ( ! theText ) return E_POINTER;
   return fromEncodedText(theText);
   }

 
   HRESULT Property::get_encodedText(BSTR* pText) {
   if ( ! pText ) return E_POINTER;
   return toEncodedText(pText);
   }

 

   HRESULT Property::get_variantType(VARTYPE* pvt) {

   if ( ! pvt ) return E_POINTER;

   *pvt = VT_EMPTY;

   switch ( type ) {
   case TYPE_BOOL:
      *pvt = VT_BOOL;
      break;

   case TYPE_LONG:
      *pvt = VT_I4;
      break;

   case TYPE_DOUBLE:
      *pvt = VT_R8;
      break;

   case TYPE_SZSTRING:
   case TYPE_STRING:
      *pvt = VT_BSTR;
      break;

   case TYPE_RAW_BINARY:
   case TYPE_BINARY: 
      *pvt = VT_BYREF | VT_UI1;
      break;

   case TYPE_ARRAY:
      *pvt = VT_ARRAY | VT_VARIANT;
      break;

   }

   return S_OK;
   }

 

   HRESULT Property::addStorageObject(IUnknown* pObject) {

   if ( ! pObject ) return E_POINTER;

   IPersistStream* pIPersistStream;

   HRESULT hr = pObject -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));

   if ( ! SUCCEEDED(hr) ) {
      IPersistStorage *pIPersistStorage;
      hr = pObject -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
      if ( ! SUCCEEDED(hr) ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"An attempt was made to add an object to an array of storage objects but "
                              "the object does not support the IPersistStream interface.");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
         return hr;
      }

      type = TYPE_OBJECT_STORAGE_ARRAY;
      
      storageObjects.Add(new persistenceInterfaces(NULL,pIPersistStorage));

      return S_OK;
   }

   type = TYPE_OBJECT_STORAGE_ARRAY;

   storageObjects.Add(new persistenceInterfaces(pIPersistStream,NULL));

   return S_OK;
   }


   HRESULT Property::removeStorageObject(IUnknown* pObject) {
   
   if ( ! pObject ) return E_POINTER;

   if ( storageObjects.Count() < 1 ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to remove an object from an array of "
                           "storage objects but no objects have ever been added to the property.");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
   }   

   IPersistStream *pIPersistStream = NULL;
   IPersistStorage *pIPersistStorage = NULL;

   HRESULT hr = pObject -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));

   if ( ! SUCCEEDED(hr) ) {
      hr = pObject -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to remove an object from the array of "
                           "storage objects but the object does not support the IPersistStream interface.");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return hr;
   }

   pObject -> Release();

   persistenceInterfaces *p = NULL;
   while ( p = storageObjects.GetNext(p) ) {
      if ( NULL != p -> pIPersistStream && p -> pIPersistStream == pIPersistStream ) {
         p -> Release();
         storageObjects.Remove(p);
         delete p;
         p = NULL;
      }
      if ( NULL != pIPersistStorage && p -> pIPersistStorage == pIPersistStorage ) {
         p -> Release();
         storageObjects.Remove(p);
         delete p;
         p = NULL;
      }
   }

   return S_OK;
   }


   HRESULT Property::clearStorageObjects() {
   persistenceInterfaces *p;
   while ( p = storageObjects.GetFirst() ) {
      p -> Release();
      storageObjects.Remove(p);
      delete p;
   }
   return S_OK;
   }


   HRESULT Property::get_storedObjectCount(long* cntObjects) {
   if ( ! cntObjects ) return E_POINTER;
   if ( ! v.binarySize ) {
      *cntObjects = 0;
      return S_OK;
   }
   if ( type != TYPE_OBJECT_STORAGE_ARRAY ) {
      *cntObjects = 0;
      return S_OK;
   }
   memcpy(cntObjects,v.binaryValue + sizeof(storageObjectSignature),sizeof(long));
   return S_OK;
   }


   HRESULT Property::writeStorageObjects() {

   long cntObjects = storageObjects.Count();

   type = TYPE_OBJECT_STORAGE_ARRAY;

   HGLOBAL hGlobal;

   IStream* pIStream;
   IStorage** ppIStorage = new IStorage*[cntObjects];
   ILockBytes** ppILockBytes = new ILockBytes*[cntObjects];
   long* sizes = new long[cntObjects + 1];

   memset(sizes,0,(cntObjects + 1) * sizeof(long));

   BSTR streamName = SysAllocString(L"GProperty");
  
   persistenceInterfaces *pInterfaces = NULL;

   for ( long k = 0; k < cntObjects; k++ ) {

      pInterfaces = storageObjects.GetNext(pInterfaces);

      CreateILockBytesOnHGlobal(NULL,TRUE,&ppILockBytes[k]);

      HRESULT hr = StgCreateDocfileOnILockBytes(ppILockBytes[k],STGM_READWRITE | STGM_DIRECT | STGM_CREATE | STGM_SHARE_EXCLUSIVE,0,&ppIStorage[k]);

      if ( pInterfaces -> pIPersistStream ) {

         ppIStorage[k] -> CreateStream(streamName,STGM_WRITE | STGM_DIRECT | STGM_CREATE | STGM_SHARE_EXCLUSIVE,0,0,&pIStream);
         pInterfaces -> pIPersistStream -> Save(pIStream,TRUE);
         pIStream -> Commit(0);
         pIStream -> Release();

      } else if ( pInterfaces -> pIPersistStorage ) {

         pInterfaces -> pIPersistStorage -> Save(ppIStorage[k],FALSE);

         pInterfaces -> pIPersistStorage -> SaveCompleted(ppIStorage[k]);

         ppIStorage[k] -> Commit(STGC_DEFAULT);

      }

      GetHGlobalFromILockBytes(ppILockBytes[k],&hGlobal);

      sizes[k] = (long)GlobalSize(hGlobal);

   }

   SysFreeString(streamName);

   if ( v.binarySize ) 
      delete [] v.binaryValue;

   v.binarySize = sizeof(storageObjectSignature) + sizeof(long);
   for ( int k = 0; k < cntObjects; k++ )
      v.binarySize += sizeof(long) + sizes[k];

   v.binaryValue = new BYTE[v.binarySize];

   BYTE *b = v.binaryValue;
   void *pv;

   memcpy(b,storageObjectSignature,sizeof(storageObjectSignature));

   b += sizeof(storageObjectSignature);

   memcpy(b,&cntObjects,sizeof(long));

   b += sizeof(long);

   for ( int k = 0; k < cntObjects; k++ ) {

      memcpy(b,&sizes[k],sizeof(long));

      b += sizeof(long);

      GetHGlobalFromILockBytes(ppILockBytes[k],&hGlobal);

      pv = GlobalLock(hGlobal);
      memcpy(b,pv,sizes[k]);
      GlobalUnlock(hGlobal);

      b += sizes[k];

      ppIStorage[k] -> Release();
      ppILockBytes[k] -> Release();

   }

   delete [] ppIStorage;
   delete [] ppILockBytes;
   delete [] sizes;

   everAssigned = true;

   return S_OK;
   }


   HRESULT Property::readStorageObjects() {

   if ( ! v.binarySize ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to read the objects from sub-storage but the sub-storage has \nnot yet been retrieved from the master storage area.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   long cntObjects = storageObjects.Count();

   BYTE* pbTest = new BYTE[sizeof(storageObjectSignature)];

   memcpy(pbTest,v.binaryValue,sizeof(storageObjectSignature));

   if ( memcmp(pbTest,storageObjectSignature,sizeof(storageObjectSignature)) ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to read a TYPE_OBJECT_STORAGE_ARRAY from it's storage(s) into it's objects,\nhowever, the property's stored information is not valid.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      delete [] pbTest;
      return S_OK;
   }

   delete [] pbTest;

   long cntStoredObjects;

   BYTE *b = v.binaryValue + sizeof(storageObjectSignature);

   memcpy(&cntStoredObjects,b,sizeof(long));

   b += sizeof(long);

   if ( cntStoredObjects != cntObjects ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         if ( cntStoredObjects < cntObjects ) 
            sprintf(szError,"An attempt was made to read a TYPE_OBJECT_STORAGE_ARRAY from it's storage(s) into it's objects.\n\n"
                                 "However, there were fewer objects saved in the storage than there are now expecting to be restored from it.\n\n"
                                 "The property name is '%s'",name[0] ? name : "<unnamed>");
         else
            sprintf(szError,"An attempt was made to read a TYPE_OBJECT_STORAGE_ARRAY from it's storage(s) into it's objects.\n\n"
                                 "However, there were more objects saved in the storage than there are now expecting to be restored from it.\n\n"
                                 "The property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      cntObjects = min(cntStoredObjects,cntObjects);
   }

   IStream* pIStream;
   IStorage* pIStorage;
   ILockBytes* pILockBytes;

   HGLOBAL hGlobal;
   long cntBytes;
   void *pv;

   BSTR streamName = SysAllocString(L"GProperty");
  
   persistenceInterfaces *pInterfaces = NULL;

   for ( long k = 0; k < cntObjects; k++ ) {

      pInterfaces = storageObjects.GetNext(pInterfaces);

      memcpy(&cntBytes,b,sizeof(long));

      b += sizeof(long);

      hGlobal = GlobalAlloc(GMEM_MOVEABLE,cntBytes);

      pv = GlobalLock(hGlobal);

      memcpy(pv,b,cntBytes);

      GlobalUnlock(hGlobal);

      b += cntBytes;

      CreateILockBytesOnHGlobal(hGlobal,TRUE,&pILockBytes);

      StgOpenStorageOnILockBytes(pILockBytes,NULL,STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,NULL,0,&pIStorage);

      if ( pInterfaces -> pIPersistStream ) {

         pIStorage -> OpenStream(streamName,0,STGM_DIRECT | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,&pIStream);

         pInterfaces -> pIPersistStream -> Load(pIStream);

         pIStream -> Release();

      } else if ( pInterfaces -> pIPersistStorage ) {

         pInterfaces -> pIPersistStorage -> Load(pIStorage);

      }

      pIStorage -> Release();
      pILockBytes -> Release();

   }

   return S_OK;
   }


   // General Methods

   HRESULT Property::advise(IGPropertyClient *pir) {
   if ( pIPropertyClient ) 
      pIPropertyClient -> Release();
   pIPropertyClient = pir;
   if ( pIPropertyClient )
      pIPropertyClient -> AddRef();
   return S_OK;
   }
 
 
   HRESULT Property::assign(void *v,long pLength) {
   switch ( type ) {
   case TYPE_BOOL:
      put_boolValue(*reinterpret_cast<short*>(v));
      break;
 
   case TYPE_LONG:
      put_longValue((*reinterpret_cast<long *>(v)));
      break;
 
   case TYPE_DOUBLE:
      put_doubleValue(*reinterpret_cast<double *>(v));
      break;
 
   case TYPE_SZSTRING:
      put_szValue(reinterpret_cast<char*>(v));
      break;
 
   case TYPE_STRING: {
      everAssigned = true;
      if ( bstrValue ) SysFreeString(bstrValue);
      bstrValue = NULL;
      BSTR bstrTemp;
      if ( ! *((char *)v + pLength - 1) ) {
         char *t = new char[pLength + 1];
         bstrTemp = SysAllocStringLen(NULL,pLength + 1);
         memset(t,0,pLength + 1);
         memcpy(t,v,pLength);
         MultiByteToWideChar(CP_ACP,0,t,-1,bstrTemp,pLength + 1);
         delete [] t;
      } else {
         bstrTemp = SysAllocStringLen(NULL,pLength);
         MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v),-1,bstrTemp,pLength);
      }
      put_stringValue(bstrTemp);
      SysFreeString(bstrTemp);
      }
      break;

   case TYPE_OBJECT_STORAGE_ARRAY: // <---- 09/21/07
   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
      put_binaryValue(pLength,reinterpret_cast<BYTE*>(v));
      break;
 
   case TYPE_ARRAY: {

      everAssigned = true;

      BSTR bstrData = SysAllocStringLen(NULL,pLength);
      MultiByteToWideChar(CP_ACP,0,(char*)v,pLength,bstrData,pLength);
      fromEncodedText(bstrData);
      SysFreeString(bstrData);

      }
      break;

   default:
      break;
   }
   return S_OK;
   }
 
 
   HRESULT Property::directAccess(enum PropertyType propertyType,void* da,long das) {
   type = propertyType;
   if ( TYPE_ARRAY == type ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"It is not possible to use the direct memory access feature for an array properties (IGProperty::directAccess())");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return E_FAIL;
   }
   directAccessAddress = da;
   directAccessSize = das;
   if ( ! directAccessSize ) get_size(&directAccessSize);
   put_size(directAccessSize);
   return S_OK;
   }


   HRESULT Property::copyTo(IGProperty* pIProperty_destination) {

   long cbToWrite;
   BYTE *dataToWrite;
 
   get_size(&cbToWrite);
   BYTE* toSpace = new BYTE[cbToWrite];
   get_binaryData(reinterpret_cast<BYTE **>(&dataToWrite));
   memcpy(toSpace,dataToWrite,cbToWrite);

   pIProperty_destination -> assign(toSpace,cbToWrite);

   delete [] toSpace;

   return S_OK;
   }

