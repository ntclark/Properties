// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

#include "resource.h"
#include "utils.h"

#include "List.cpp"

//#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' " \
//                        "version='6.0.0.0' processorArchitecture='x64' publicKeyToken='6595b64144ccf1df' language='*'\"")

   WNDPROC Properties::nativePropertySheetHandler = NULL;
   static BYTE registrationMark[] = {0xFF,0xEF,0xEE,0xEE};
   static BYTE notSavedIndicator[] = {0xEE,0xFE,0xEF,0xFF};

   char szPropertyType[][32] = {
           "TYPE_LONG",
           "TYPE_DOUBLE",
           "TYPE_SZSTRING",
           "TYPE_STRING",
           "TYPE_RAW_BINARY",
           "TYPE_BINARY",
           "TYPE_BOOL",
           "TYPE_DATE",
           "TYPE_VARIANT",
           "TYPE_ARRAY",
           "TYPE_UNSPECIFIED"};

   Properties* Properties::pCurrent_IO_Object = NULL;

   Properties::Properties(IUnknown* pIUnkOuter) :

     List<interface IGProperty>(),

     simplePersistenceWindows(),
     persistableObjectInterfaces(),
     persistableObjects(),

     refCount(100),
     currentItemIndex(0),
     bPreviouslyGotStream(false),

     hwndPropertySheet(NULL),
     pPropertySheetPages(NULL),
#if 0
     cxClientIdeal(0L),
     cyClientIdeal(0L),
     cxSheetIdeal(0L),
     cySheetIdeal(0L),
#endif
     propertyFrameInstanceCount(0L),

     supportIPersistStorage(false),
     supportIPersistStream(true),
     temporaryFileInUse(false),
     debuggingEnabled(false),
     isDirty(false),

     currentPersistenceMechanism(MECHANISM_DEFAULT),

     fileName(NULL),
     fileExtensions(NULL),
     fileType(NULL),
     fileSaveOpenText(NULL),

     propertyFileName(NULL),
     propertyFileSaveText(NULL),
     propertyFileExtensions(NULL),
     propertyFileType(NULL),
     propertyDebuggingEnabled(NULL),

     pIPropertyPageClient_ThisObject(NULL),

     pIPropertiesClient(NULL),
     pIStorage_current(NULL),
     pIStream_current(NULL),  

     pMasterPropertyPageClient(NULL),

     enumConnectionPoints(0),
     enumConnections(0),
     propertyNotifySinkConnectionPoint(this,IID_IPropertyNotifySink),
     connectionPointContainer(this)

   {
 
   if ( pIUnkOuter ) 
      pIUnknownOuter = pIUnkOuter;
   else
      pIUnknownOuter = &innerUnknown;

   memset(&openFileName,0,sizeof(OPENFILENAME));
   memset(&objectCLSID,0,sizeof(CLSID));
   memset(&propertySheetHeader,0,sizeof(PROPSHEETHEADER));

   memset(cxClientIdeal,0,sizeof(cxClientIdeal));
   memset(cyClientIdeal,0,sizeof(cyClientIdeal));
   memset(cxSheetIdeal,0,sizeof(cxSheetIdeal));
   memset(cySheetIdeal,0,sizeof(cySheetIdeal));

   pIProperties = new _IProperties(this);
   pIPersistStorage = new _IPersistStorage(this);
   pIPersistStream = new _IPersistStream(this);
   pIPersistStreamInit = new _IPersistStreamInit(this);
   pIPersistPropertyBag = new _IPersistPropertyBag(this);
   pIPersistPropertyBag2 = new _IPersistPropertyBag2(this);

   CoGetClassObject(CLSID_InnoVisioNateProperty,CLSCTX_INPROC_SERVER,NULL,IID_IClassFactory,reinterpret_cast<void**>(&pProperty_IClassFactory));

   refCount = 0;

   pIProperties -> Add(NULL,&propertyDebuggingEnabled);

   propertyDebuggingEnabled -> directAccess(TYPE_BOOL,&debuggingEnabled,sizeof(debuggingEnabled));

   return;
   }
 
 
   Properties::~Properties() {

   if ( pIPropertiesClient )
      pIPropertiesClient -> Release();

   if ( propertyFileName ) propertyFileName -> Release();
   if ( propertyFileSaveText ) propertyFileSaveText -> Release();
   if ( propertyFileExtensions ) propertyFileExtensions -> Release();
   if ( propertyFileType ) propertyFileType -> Release();
   if ( propertyDebuggingEnabled ) propertyDebuggingEnabled -> Release();

   IGProperty* pIProperty;
   while ( pIProperty = propertiesToRelease.GetFirst() ) {
      pIProperty -> Release();
      propertiesToRelease.Remove(pIProperty);
   }

   simplePersistenceWindow* pspw2 = (simplePersistenceWindow*)NULL;
   while ( pspw2 = simplePersistenceWindows.GetFirst() ) {
      simplePersistenceWindows.Remove(pspw2);
      delete pspw2;
   }

   delete pIProperties;
   delete pIPersistStorage;
   delete pIPersistStream;
   delete pIPersistStreamInit;
   delete pIPersistPropertyBag;
   delete pIPersistPropertyBag2;

   pIProperties = NULL;
   pIPersistStorage = NULL;
   pIPersistStream = NULL;
   pIPersistStreamInit = NULL;
   pIPersistPropertyBag = NULL;
   pIPersistPropertyBag2 = NULL;

   BYTE** pb;
   while ( pb = pushStackList.GetFirst() ) {
      delete [] *pb;
      delete pb;
      pushStackList.Remove(pb);
   }

   if ( fileName ) SysFreeString(fileName);
   if ( fileExtensions ) SysFreeString(fileExtensions);
   if ( fileType ) SysFreeString(fileType);
   if ( fileSaveOpenText ) SysFreeString(fileSaveOpenText);

   fileName = NULL;
   fileExtensions = NULL;
   fileType = NULL;
   fileSaveOpenText = NULL;

   persistableObjectInterface* pObj = (persistableObjectInterface*)NULL;
   while ( pObj = persistableObjectInterfaces.GetFirst() ) {
      persistableObjectInterfaces.Remove(pObj);
      delete pObj;
   }

   IUnknown** ppIUnknown = (IUnknown **)NULL;
   while ( ppIUnknown = persistableObjects.GetFirst() ) {
      persistableObjects.Remove(ppIUnknown);
      delete ppIUnknown;
   }

   for ( std::list<IGPropertyPageClient *>::iterator it = propertyPageClients.begin(); it != propertyPageClients.end(); it++ )
      (*it) -> Release();

   propertyPageClients.clear();

   return;
   }


   IGProperty* Properties::getProperty(BSTR bstrPropertyName,char* szMethodName) {
   char szName[MAX_PROPERTY_SIZE];
   memset(szName,0,sizeof(szName));
   WideCharToMultiByte(CP_ACP,0,bstrPropertyName,-1,szName,MAX_PROPERTY_SIZE,0,0);
   IGProperty* pProperty = Get(szName);
   if ( ! pProperty ) {
      if ( szMethodName && debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An invalid property name (%s) was passed to %s",szName,szMethodName);
         MessageBox(NULL,szError,"Properties Component usage note.",MB_OK);
      }
   }
   return pProperty;
   }


   HRESULT Properties::internalSave(IStream* pIStream,IStorage* pIStorage) {

   if ( ! pIStream ) {
      if ( debuggingEnabled )
         MessageBox(NULL,"IProperties::Save(). The IStream argument to Save() is NULL\n\nThe properties are not saved","GSystem Properties Component; error message",MB_OK);
      return E_UNEXPECTED;
   }
 
   HRESULT hr = S_OK;

   if ( pIPropertiesClient )
      pIPropertiesClient -> SavePrep();

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = simplePersistenceWindows.GetNext(pspw) )
      if ( pspw -> idSavePrep ) SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idSavePrep,0),(LPARAM)pspw -> hwndSavePrep);

   unsigned long cbWritten;
   short everAssigned;
   BYTE *dataToWrite;
   storageInfo si;
 
   memcpy(&si.registrationMark,registrationMark,8);

   IGProperty *pProperty = 0;

   while ( pProperty = GetNext(pProperty) ) {

      pProperty -> get_type(&si.propertyType);

      pProperty -> get_everAssigned(&everAssigned);

      if ( ! everAssigned )
         si.cBytes = sizeof(notSavedIndicator);
      else {
         if ( TYPE_ARRAY == si.propertyType || TYPE_BINARY == si.propertyType ) {
            BSTR bstrData;
            pProperty -> get_encodedText(&bstrData);
            si.cBytes = (long)wcslen(bstrData);
            SysFreeString(bstrData);
         } else 
            pProperty -> get_size(&si.cBytes);
      }

      hr = pIStream -> Write(&si,sizeof(storageInfo),&cbWritten);
      if ( (sizeof(storageInfo) != cbWritten) || FAILED(hr) )
         return hr;

      if ( ! everAssigned ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            char szName[MAX_PROPERTY_SIZE];
            BSTR bstrName;
            pProperty -> get_name(&bstrName);
            WideCharToMultiByte(CP_ACP,0,bstrName,-1,szName,MAX_PROPERTY_SIZE,0,0);
            sprintf(szError,"While attempting to save the properties, IProperties::Save has encountered the property \'%s\' that has never been given a value or doesn't use direct access storage.\n\n"
                                    "This property is not saved.\n\nTurn debugging off to avoid this message.",szName);
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
            SysFreeString(bstrName);
         }
         dataToWrite = reinterpret_cast<BYTE *>(&notSavedIndicator);
      } else {
         if ( TYPE_ARRAY == si.propertyType || TYPE_BINARY == si.propertyType ) {
            BSTR bstrData;
            pProperty -> get_encodedText(&bstrData);
            dataToWrite = new BYTE[wcslen(bstrData) + 1];
            WideCharToMultiByte(CP_ACP,0,bstrData,-1,(char*)dataToWrite,(int)wcslen(bstrData) + 1,0,0);
            SysFreeString(bstrData);
         } else
            pProperty -> get_binaryData(reinterpret_cast<BYTE **>(&dataToWrite));
      }

      if ( dataToWrite ) {
         hr = pIStream -> Write(dataToWrite,si.cBytes,&cbWritten);
         if ( (si.cBytes != (long)cbWritten) || FAILED(hr) ) {
            if ( TYPE_BINARY == si.propertyType || TYPE_ARRAY == si.propertyType ) 
               delete [] dataToWrite;
            return hr;
         }
      } else {
         //
         //NTC: 01-14-2018: I am going to prevent this diagnostic for TYPE_BINARY data types because I think it is acceptable to have a 0-sized binary property
         //
         if ( ! ( TYPE_BINARY == si.propertyType ) ) {
            char szError[MAX_PROPERTY_SIZE];
            char szName[MAX_PROPERTY_SIZE];
            BSTR bstrName;
            pProperty -> get_name(&bstrName);
            WideCharToMultiByte(CP_ACP,0,bstrName,-1,szName,MAX_PROPERTY_SIZE,0,0);
            sprintf(szError,"While attempting to save the properties, IProperties::Save has encountered the property \'%s\' that does not have storage.\n\nThis property is not saved.\n\nTurn debugging off to avoid this message.",szName);
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
            SysFreeString(bstrName);
         }
      }

      if ( everAssigned && ( TYPE_BINARY == si.propertyType || TYPE_ARRAY == si.propertyType) ) 
         delete [] dataToWrite;

      pProperty -> put_isDirty(false);

   }
 
   switch ( currentPersistenceMechanism ) {
   case MECHANISM_STREAM: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAM: {
                  IPersistStream* pIPersistStream;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));
                  pIPersistStream -> Save(pIStream,true);
                  pIPersistStream -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_STREAMINIT: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAMINIT: {
                  IPersistStreamInit* pIPersistStreamInit;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersistStreamInit));
                  pIPersistStreamInit -> Save(pIStream,true);
                  pIPersistStreamInit -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_STORAGE: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAMINIT: {
                  IPersistStreamInit* pIPersistStreamInit;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersistStreamInit));
                  pIPersistStreamInit -> Save(pIStream,true);
                  pIPersistStreamInit -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
/*
              case MECHANISM_STREAM: {
                  IPersistStream* pIPersistStream;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));
                  pIPersistStream -> Save(pIStream,true);
                  pIPersistStream -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               case MECHANISM_STORAGE: {
                  IPersistStorage* pIPersistStorage;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
                  pIPersistStorage -> Save(pIStorage,0);
                  pIPersistStorage -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
*/
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   }

   return S_OK;
   }
 
 

   HRESULT Properties::savePropertyBag(IPropertyBag* pIPropertyBag,IPropertyBag2* pIPropertyBag2,BOOL fClearDirty,BOOL fSaveAllProperties) {

   if ( ! pIPropertyBag ) return E_POINTER;

   HRESULT hr = S_OK;

   if ( pIPropertiesClient )
      pIPropertiesClient -> SavePrep();

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = simplePersistenceWindows.GetNext(pspw) )
      if ( pspw -> idSavePrep ) SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idSavePrep,0),(LPARAM)pspw -> hwndSavePrep);

   short everAssigned;
   VARIANT v;
   long n;
   BSTR bstrName;
   PropertyType type;
   PROPBAG2 propBag2;

   if ( pIPropertyBag2 ) {
      memset(&propBag2,0,sizeof(PROPBAG2));
      propBag2.dwType = PROPBAG2_TYPE_DATA;
   }

   IGProperty *pProperty = 0;
   while ( pProperty = GetNext(pProperty) ) {

      pProperty -> get_compressedName(&bstrName);

      pProperty -> get_everAssigned(&everAssigned);

      pProperty -> get_type(&type);

      char *pszName = new char[n = (long)wcslen(bstrName) + 6];
      memset(pszName,0,n);
      WideCharToMultiByte(CP_ACP,0,bstrName,-1,pszName,n,0,0);
      strcpy(pszName + strlen(pszName),"_type");
      BSTR bstrType = SysAllocStringLen(NULL,n);
      MultiByteToWideChar(CP_ACP,0,pszName,-1,bstrType,n);
      v.vt = VT_I4;
      v.lVal = type;

      if ( pIPropertyBag2 ) {
         propBag2.vt = v.vt;
         propBag2.pstrName = bstrType;
         pIPropertyBag2 -> Write(1,&propBag2,&v);
      }

      if ( pIPropertyBag )
         pIPropertyBag -> Write(bstrType,&v);

      SysFreeString(bstrType);
      delete [] pszName;

      if ( ! everAssigned ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            char szName[MAX_PROPERTY_SIZE];
            WideCharToMultiByte(CP_ACP,0,bstrName,-1,szName,MAX_PROPERTY_SIZE,0,0);
            sprintf(szError,"While attempting to save the properties, IProperties::Save has encountered the property \'%s\' that has never been given a value or doesn't use direct access storage.\n\n"
                              "This property is not saved.\n\nTurn debugging off to avoid this message.",szName);
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
            SysFreeString(bstrName);
         }
      }

      switch ( type ) {
//CHECKME      case TYPE_OBJECT_STORAGE_ARRAY:
      case TYPE_BINARY:
      case TYPE_ARRAY:
         v.vt = VT_BSTR;
         pProperty -> get_encodedText(&v.bstrVal);
         break;

      default:
         v.vt = VT_EMPTY;
         pProperty -> get_variantValue(&v);

      }

      if ( pIPropertyBag2 ) {
         propBag2.vt = v.vt;
         propBag2.pstrName = bstrName;
         pIPropertyBag2 -> Write(1,&propBag2,&v);
      }

      if ( pIPropertyBag )
         hr = pIPropertyBag -> Write(bstrName,&v);
      
      if ( ! SUCCEEDED(hr) ) {
         if ( debuggingEnabled ) {
            BSTR bstrActualName;
            char szError[MAX_PROPERTY_SIZE];
            char szName[MAX_PROPERTY_SIZE];
            pProperty -> get_name(&bstrActualName);
            WideCharToMultiByte(CP_ACP,0,bstrActualName,-1,szName,MAX_PROPERTY_SIZE,0,0);
            sprintf(szError,"While attempting to save the properties, IProperties::Save has encountered the property \'%s\' that is not able to be saved in a IPropertyBag implementation.\n\nThis property is not saved.\n\nTurn debugging off to avoid this message.",szName);
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
            SysFreeString(bstrActualName);
         }
      }

//CHECKME
      if ( /* TYPE_OBJECT_STORAGE_ARRAY == type || */ TYPE_BINARY == type || TYPE_ARRAY == type ) 
         SysFreeString(v.bstrVal);

      SysFreeString(bstrName);

      if ( fClearDirty ) 
         pProperty -> put_isDirty(false);

   }

   switch ( currentPersistenceMechanism ) {
   case MECHANISM_PROPERTYBAG: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG: {
                  IPersistPropertyBag* pIPersistPropertyBag;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag,reinterpret_cast<void**>(&pIPersistPropertyBag));
                  pIPersistPropertyBag -> Save(pIPropertyBag,fClearDirty,fSaveAllProperties);
                  pIPersistPropertyBag -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_PROPERTYBAG2: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG2: {
                  IPersistPropertyBag2* IPersistPropertyBag2;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag2,reinterpret_cast<void**>(&IPersistPropertyBag2));
                  IPersistPropertyBag2 -> Save(pIPropertyBag2,fClearDirty,fSaveAllProperties);
                  IPersistPropertyBag2 -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   }

/*
   if ( pIPropertiesClient )
      pIPropertiesClient -> Saved();
*/
   return hr;

   }
  

   HRESULT Properties::internalLoad(IStream* pIStream,IStorage* pIStorage) {

   HRESULT hr = S_OK;
 
   if ( ! pIStream ) {
      if ( debuggingEnabled )
         MessageBox(NULL,"IProperties::Load(). The IStream argument to Load() is NULL\n\nThe properties are not saved","GSystem Properties Component; error message",MB_OK);
      return E_UNEXPECTED;
   }

   unsigned long cbRead;
   BYTE* dataRead;
   storageInfo si;

   IGProperty *pProperty = 0;

   while ( pProperty = GetNext(pProperty) ) {

      hr = pIStream -> Read(&si,sizeof(storageInfo),&cbRead);

      if ( ( sizeof(storageInfo) != cbRead ) || FAILED(hr) ) {
         if ( debuggingEnabled )
            if ( IDCANCEL == MessageBox(NULL,"IProperties::Load(). An object expected to be read from storage does not appear to have been saved to storage" 
                                 "\nThis may be due to adding properties to the set of properties since the last save."
                                 "\nLoading properties will attempt to continue\n",
                                 "GSystem Properties Component; error message",MB_OKCANCEL) ) {
#if ! defined(_M_X64)
               _asm {
               int 3;
               }
#endif
            }
         continue;
      }

      if ( memcmp(registrationMark,&si.registrationMark,8) ) {
         if ( debuggingEnabled )
            MessageBox(NULL,"IProperties::Load(). Data present on the saved storage does not appear to have been saved by the GSystem Properties Control.\n\nThis can happen if you have added properties and/or objects and are trying to load a storage upon which these properties or objects were not saved!","GSystem Properties Component; error message",MB_OK);
         return E_FAIL;
      }

#ifndef _DEBUG
      debuggingEnabled = false;
#endif

      pProperty -> put_type(si.propertyType);

      if ( 0 == si.cBytes ) {
         pProperty -> put_isDirty(false);
         continue;
      }

      dataRead = new BYTE[si.cBytes];

      hr = pIStream -> Read(dataRead,si.cBytes,&cbRead);

      if ( ( si.cBytes != (long)cbRead ) || FAILED(hr) ) {
         delete [] dataRead;
         if ( ( si.cBytes != (long)cbRead ) && ! FAILED(hr) ) {
            if ( debuggingEnabled )
               MessageBox(NULL,"IProperties::Load(). It appears the saved storage does not contain all the information expected.\n\nThis can happen if you add objects/properties after having saved the storage.\n\nRe-saving the storage will fix this.","GSystem Properties Component; error message",MB_OK);
            return S_OK;
         }
         return hr;
      }

      if ( sizeof(notSavedIndicator) == si.cBytes ) {

         if ( strncmp(reinterpret_cast<char *>(dataRead),reinterpret_cast<char *>(notSavedIndicator),sizeof(notSavedIndicator)) )
            pProperty -> assign(dataRead,si.cBytes);

      } else {

         if ( TYPE_ARRAY == si.propertyType || TYPE_BINARY == si.propertyType ) {
            BSTR bstrData = SysAllocStringLen(NULL,si.cBytes + 1);
            MultiByteToWideChar(CP_ACP,0,(char *)dataRead,si.cBytes,bstrData,si.cBytes + 1);
            pProperty -> put_encodedText(bstrData);
            SysFreeString(bstrData);
         } else
            pProperty -> assign(dataRead,si.cBytes);

      }

      delete [] dataRead;

      pProperty -> put_isDirty(false);

   }

   switch ( currentPersistenceMechanism ) {
   case MECHANISM_STREAM: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAM: {
                  IPersistStream* pIPersistStream;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStream -> Load(pIStream);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistStream -> Load(pIStream);
                  pIPersistStream -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_STREAMINIT: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAMINIT: {
                  IPersistStreamInit* pIPersistStreamInit;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersistStreamInit));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStreamInit -> Load(pIStream);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                      pIPersistStreamInit -> Load(pIStream);
                  pIPersistStreamInit -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_STORAGE: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
/*
               case MECHANISM_STREAM: {
                  IPersistStream* pIPersistStream;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStream,reinterpret_cast<void**>(&pIPersistStream));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStream -> Load(pIStream);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistStream -> Load(pIStream);
                  pIPersistStream -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
*/
               case MECHANISM_STREAMINIT: {
                  IPersistStreamInit* pIPersistStreamInit;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersistStreamInit));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStreamInit -> Load(pIStream);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                      pIPersistStreamInit -> Load(pIStream);
                  pIPersistStreamInit -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
/*
               case MECHANISM_STORAGE: {
                  IPersistStorage* pIPersistStorage;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStorage -> Load(pIStorage);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistStorage -> Load(pIStorage);
                  pIPersistStorage -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
*/
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;
   }

   if ( pIPropertiesClient )
      hr = pIPropertiesClient -> Loaded();

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = simplePersistenceWindows.GetNext(pspw) )
      if ( pspw -> idLoad ) SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idLoad,0),(LPARAM)pspw -> hwndLoad);

   return hr;

   }
 
 
   HRESULT Properties::loadPropertyBag(IPropertyBag* pIPropertyBag,IPropertyBag2* pIPropertyBag2,IErrorLog* pIErrorLog) {

   if ( ! pIPropertyBag ) return E_POINTER;

   HRESULT hr = S_OK;

   BSTR bstrName;
   long n;
   PropertyType type;
   VARIANT v;
   PROPBAG2 propBag2;

   IGProperty *pProperty = 0;
   while ( pProperty = GetNext(pProperty) ) {

      pProperty -> get_compressedName(&bstrName);

      char* pszName = new char[n = (long)wcslen(bstrName) + 6];
      WideCharToMultiByte(CP_ACP,0,bstrName,-1,pszName,n,0,0);
      strcpy(pszName + strlen(pszName),"_type");
      BSTR bstrType = SysAllocStringLen(NULL,n);
      MultiByteToWideChar(CP_ACP,0,pszName,-1,bstrType,n);

      v.vt = VT_I4;

      if ( pIPropertyBag2 ) {
         propBag2.vt = VT_I4;
         propBag2.pstrName = bstrType;
         pIPropertyBag2 -> Read(1,&propBag2,pIErrorLog,&v,&hr);
      }

      if ( pIPropertyBag ) 
         pIPropertyBag -> Read(bstrType,&v,pIErrorLog);
            
      type = (enum PropertyType)v.lVal;

      pProperty -> put_type(type);

      delete [] pszName;
      SysFreeString(bstrType);

      GVariantClear(&v);
      pProperty -> get_variantType(&v.vt);
//CHECKME
      if ( /* TYPE_OBJECT_STORAGE_ARRAY == type || */ TYPE_BINARY == type || TYPE_ARRAY == type ) 
         v.vt = VT_BSTR;

      if ( pIPropertyBag2 ) {
         propBag2.vt = v.vt;
         propBag2.pstrName = bstrName;
         pIPropertyBag2 -> Read(1,&propBag2,pIErrorLog,&v,&hr);
      }

      if ( pIPropertyBag ) 
         hr = pIPropertyBag -> Read(bstrName,&v,pIErrorLog);

      switch ( type ) {
//CHECKME      case TYPE_OBJECT_STORAGE_ARRAY:
      case TYPE_BINARY:
      case TYPE_ARRAY:
         pProperty -> put_encodedText(v.bstrVal);
         break;
      default:
         pProperty -> put_variantValue(v);
         break;
      }

      SysFreeString(bstrName);

      pProperty -> put_isDirty(false);

   }

   switch ( currentPersistenceMechanism ) {
   case MECHANISM_PROPERTYBAG: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG: {
                  IPersistPropertyBag* IPersistPropertyBag;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag,reinterpret_cast<void**>(&IPersistPropertyBag));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     IPersistPropertyBag -> Load(pIPropertyBag,pIErrorLog);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     IPersistPropertyBag -> Load(pIPropertyBag,pIErrorLog);
                  IPersistPropertyBag -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_PROPERTYBAG2: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG2: {
                  IPersistPropertyBag2* IPersistPropertyBag2;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag2,reinterpret_cast<void**>(&IPersistPropertyBag2));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     IPersistPropertyBag2 -> Load(pIPropertyBag2,pIErrorLog);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     IPersistPropertyBag2 -> Load(pIPropertyBag2,pIErrorLog);
                  IPersistPropertyBag2 -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   }

   if ( pIPropertiesClient )
      pIPropertiesClient -> Loaded();

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = simplePersistenceWindows.GetNext(pspw) )
      if ( pspw -> idLoad ) SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idLoad,0),(LPARAM)pspw -> hwndLoad);

   return hr;
   }


   HRESULT Properties::internalInitNew(IStorage* pIStorage) {

   HRESULT hr = S_OK;

   if ( pIPropertiesClient )
      pIPropertiesClient -> InitNew();

   simplePersistenceWindow* pspw = NULL;
   while ( pspw = simplePersistenceWindows.GetNext(pspw) )
      if ( pspw -> idInit ) SendMessage(pspw -> hwnd,WM_COMMAND,MAKEWPARAM(pspw -> idInit,0),(LPARAM)pspw -> hwndInit);

   switch ( currentPersistenceMechanism ) {
   case MECHANISM_STREAMINIT:
   case MECHANISM_STORAGE: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         bool suitableInterfaceFound = false;
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_STREAMINIT: {
                  IPersistStreamInit* pIPersistStreamInit;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStreamInit,reinterpret_cast<void**>(&pIPersistStreamInit));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStreamInit -> InitNew();
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistStreamInit -> InitNew();
                  pIPersistStreamInit -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               case MECHANISM_STORAGE: {
                  IPersistStorage* pIPersistStorage;
                  IOleObject* pIOleObject;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistStorage,reinterpret_cast<void**>(&pIPersistStorage));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     IOleClientSite* pIOleClientSite;
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistStorage -> InitNew(pIStorage);
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistStorage -> InitNew(pIStorage);
                  pIPersistStorage -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_PROPERTYBAG: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG: {
                  IPersistPropertyBag* pIPersistPropertyBag;
                  IOleObject* pIOleObject;
                  IOleClientSite* pIOleClientSite;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag,reinterpret_cast<void**>(&pIPersistPropertyBag));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistPropertyBag -> InitNew();
                     hr = pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else 
                     pIPersistPropertyBag -> InitNew();
                  pIPersistPropertyBag -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   case MECHANISM_PROPERTYBAG2: {
      IUnknown** pExternalObject = NULL;
      while ( pExternalObject = persistableObjects.GetNext(pExternalObject) ) {
         persistableObjectInterface* pObj = (persistableObjectInterface *)NULL;
         bool suitableInterfaceFound = false;
         while ( pObj = persistableObjectInterfaces.GetNext(pObj) ) {
            if ( pObj -> pIUnknown == *pExternalObject ) {
               switch ( pObj -> persistenceMechanism ) {
               case MECHANISM_PROPERTYBAG2: {
                  IPersistPropertyBag2* pIPersistPropertyBag2;
                  IOleObject* pIOleObject;
                  IOleClientSite* pIOleClientSite;
                  pObj -> pIUnknown -> QueryInterface(IID_IPersistPropertyBag2,reinterpret_cast<void**>(&pIPersistPropertyBag2));
                  HRESULT hr = pObj -> pIUnknown -> QueryInterface(IID_IOleObject,reinterpret_cast<void**>(&pIOleObject));
                  if ( SUCCEEDED(hr) ) {
                     pIOleObject -> GetClientSite(&pIOleClientSite);
                     pIOleObject -> Close(OLECLOSE_NOSAVE);
                     pIPersistPropertyBag2 -> InitNew();
                     pIOleObject -> DoVerb(OLEIVERB_SHOW,NULL,pIOleClientSite,0,NULL,NULL);
                     pIOleClientSite -> Release();
                     pIOleObject -> Release();
                  } else
                     pIPersistPropertyBag2 -> InitNew();
                  pIPersistPropertyBag2 -> Release();
                  suitableInterfaceFound = true;
                  }
                  break;
               }
            }
            if ( suitableInterfaceFound ) break;
         }
      }
      }
      break;

   }

   return S_OK;
   }
 
 

   HRESULT Properties::push(BYTE* toSpace) {

   long cbToWrite;
   BYTE *dataToWrite;
 
   IGProperty *pProperty = NULL;
   while ( pProperty = GetNext(pProperty) ) {
      pProperty -> get_size(&cbToWrite);
      memcpy(toSpace,&cbToWrite,4);
      toSpace += 4;
      pProperty -> get_binaryData(reinterpret_cast<BYTE **>(&dataToWrite));
      memcpy(toSpace,dataToWrite,cbToWrite);
      toSpace += cbToWrite;
   }
 
   IGProperty *pIProp = 0;
   while ( pProperty = proxyList.GetNext(pProperty) ) {
      pProperty -> get_size(&cbToWrite);
      memcpy(toSpace,&cbToWrite,4);
      toSpace += 4;
      pProperty -> get_binaryData(reinterpret_cast<BYTE **>(&dataToWrite));
      memcpy(toSpace,dataToWrite,cbToWrite);
      toSpace += cbToWrite;
   }

   return S_OK;
   }
 
 

   HRESULT Properties::pop(BYTE* fromSpace) {

   long cbShouldRead;
   BYTE* dataRead;
 
   IGProperty *pProperty = NULL;
   while ( pProperty = GetNext(pProperty) ) {
      memcpy(&cbShouldRead,fromSpace,4);
      fromSpace += 4;
      dataRead = new BYTE[cbShouldRead];
      memcpy(dataRead,fromSpace,cbShouldRead);
      fromSpace += cbShouldRead;
      pProperty -> assign(dataRead,cbShouldRead);
      pProperty -> put_isDirty(0);
      delete [] dataRead;
   }
 
   while ( pProperty = proxyList.GetNext(pProperty) ) {
      memcpy(&cbShouldRead,fromSpace,4);
      fromSpace += 4;
      dataRead = new BYTE[cbShouldRead];
      memcpy(dataRead,fromSpace,cbShouldRead);
      fromSpace += cbShouldRead;
      pProperty -> put_isDirty(0);
      delete [] dataRead;
   }

   return S_OK;
   }


   int Properties::prepOpenFileName() {

   char szFileFilters[1024];
   char szFileType[128];
   char szFileExtension[256];
   char szFilePath[MAX_PATH];
   char szFileName[MAX_PATH];
   char szFileAndPath[MAX_PATH];
   char szTemp[MAX_PATH];
   char szTitleText[MAX_PATH];
   char *c,*p;
   bool addAllFiles = false;
   long n;

   memset(&openFileName,0,sizeof(OPENFILENAME));

   openFileName.lStructSize = sizeof(OPENFILENAME);
   openFileName.hInstance = gsProperties_hModule;

   memset(szFileFilters,0,sizeof(szFileFilters));
   memset(szFileType,0,sizeof(szFileType));
   memset(szFileExtension,0,sizeof(szFileExtension));
   memset(szFilePath,0,sizeof(szFilePath));
   memset(szFileName,0,sizeof(szFileName));
   memset(szFileAndPath,0,sizeof(szFileAndPath));
   memset(szTitleText,0,sizeof(szTitleText));

   memset(szTemp,0,sizeof(szTemp));
   
   if ( fileType ) {
      WideCharToMultiByte(CP_ACP,0,fileType,-1,szFileType,128,0,0);
      strcpy(szTemp,szFileType);
      c = szTemp;
      while ( *c ) {
         *c = tolower(*c);
         c++;
      }
      addAllFiles = (0 != strcmp(szTemp,"all files"));
   }
   else
      strcpy(szFileType,"All Files");

   if ( fileExtensions ) 
      WideCharToMultiByte(CP_ACP,0,fileExtensions,-1,szFileExtension,256,0,0);
   else
      strcpy(szFileExtension,"*.*");

   if ( szFileExtension[0] ) {
      if ( szFileExtension[0] != '*' ) {
         sprintf(szTemp,"*");
         if ( szFileExtension[0] != '.' && szFileExtension[1] != '.' ) strcat(szTemp,".");
         strcat(szTemp,szFileExtension);
         strcpy(szFileExtension,szTemp);
      }
   }
   else
      sprintf(szFileExtension,"*.*");

   c = szFileExtension;
   while ( *c ) {
      if ( *c == ';' ) {
         p = c + 1;
         if ( ! *p ) break;
         if ( ! *(p + 1) ) break;
         if ( *p != '*' ) {
            strcpy(szTemp,szFileExtension);
            n = (long)(p - szFileExtension);
            memset(szTemp + n,0,255 - n);
            if ( *p != '.' )
               strcat(szTemp,"*.");
            else
               strcat(szTemp,"*");
            strcat(szTemp,p);
            strcpy(szFileExtension,szTemp);
         }
      }
      c++;
   }

   n = sprintf(szFileFilters,"%s (%s)",szFileType,szFileExtension);
   n += sprintf(szFileFilters + n + 1,"%s",szFileExtension) + 1;

   if ( addAllFiles ) {
      n += sprintf(szFileFilters + n + 1,"All Files (*.*)") + 1;
      n += sprintf(szFileFilters + n + 1,"*.*");
   }

   openFileName.lpstrFilter = new char[n + 1];
   memset(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFilter)),0,n + 1);
   memcpy(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFilter)),szFileFilters,n + 1);

/*
   if ( filePath ) {
      WideCharToMultiByte(CP_ACP,0,filePath,-1,szFilePath,MAX_PATH,0,0);
      strcpy(szFileAndPath,szFilePath);
   }
   else
      memset(szFilePath,0,sizeof(szFilePath));
*/
   if ( fileName ) {
      WideCharToMultiByte(CP_ACP,0,fileName,-1,szFileName,MAX_PATH,0,0);
      strcat(szFileAndPath,szFileName);
   }
   else
      memset(szFileAndPath,0,sizeof(szFileAndPath));

   n = (long)strlen(szFileAndPath) + 1;
   openFileName.lpstrFile = new char[MAX_PATH];
   memset(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFile)),0,MAX_PATH);
   memcpy(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFile)),szFileAndPath,n);
   openFileName.nMaxFile = MAX_PATH;

   n = (long)strlen(szFileName) + 1;
   openFileName.lpstrFileTitle = new char[MAX_PATH];
   memset(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFileTitle)),0,MAX_PATH);
   memcpy(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFileTitle)),szFileName,n);
   openFileName.nMaxFileTitle = MAX_PATH;

   n = (long)strlen(szFilePath) + 1;
   openFileName.lpstrInitialDir = new char[MAX_PATH];
   memset(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrInitialDir)),0,MAX_PATH);
   memcpy(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrInitialDir)),szFilePath,n);

   if ( fileSaveOpenText ) 
      WideCharToMultiByte(CP_ACP,0,fileSaveOpenText,-1,szTitleText,MAX_PATH,0,0);
   else
      strcpy(szTitleText,"Create File");

   openFileName.lpstrTitle = new char[n = (long)strlen(szTitleText) + 1];
   memset(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrTitle)),0,n);
   memcpy(reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrTitle)),szTitleText,n);

   openFileName.Flags = OFN_ENABLESIZING | OFN_OVERWRITEPROMPT;

   return 0;
   }


   int Properties::parseOpenFileName(bool preserveExisting) {

   int rc = 0;

   long n = (long)strlen(openFileName.lpstrFile);

   if ( n ) {

      char szFileName[MAX_PATH];

      strcpy(szFileName,openFileName.lpstrFile);

      if ( fileName ) SysFreeString(fileName);

      fileName = SysAllocStringLen(NULL,(UINT)strlen(szFileName));

      MultiByteToWideChar(CP_ACP,0,szFileName,-1,fileName,(int)strlen(szFileName));

      temporaryFileInUse = false;

      rc = 1;

   }
   else {

      if ( ! preserveExisting ) {
         if ( fileName ) SysFreeString(fileName);
         fileName = NULL;
      }

      rc = 0;

   }

   delete [] reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFilter));
   delete [] reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFile));
   delete [] reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrFileTitle));
   delete [] reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrInitialDir));
   delete [] reinterpret_cast<void*>(const_cast<char*>(openFileName.lpstrTitle));

   return rc;

   }


   int Properties::createTemporaryStorageFileName() {

   fileName = SysAllocString(_wtempnam(NULL,L"GProp"));
/*
   long n = wcslen(fileName);
   char *c,*pszName = new char[n + 1];
   memset(pszName,0,n + 1);
   WideCharToMultiByte(CP_ACP,0,fileName,-1,pszName,n + 1,0,0);
   c = pszName + n - 1;
   while ( c > pszName && strncmp(c,P_tmpdir,1) ) c--;
   if ( c > pszName ) {
      if ( fileName ) SysFreeString(fileName);
      fileName = SysAllocStringLen(NULL,n + 1 - (c - pszName));
      MultiByteToWideChar(CP_ACP,0,c + 1,-1,fileName,n + 1 - (c - pszName));
   }

   delete [] pszName;
*/
   temporaryFileInUse = true;

   return 0;
   }


   int Properties::deleteTemporaryStorage() {
   char szFile[MAX_PATH];
   memset(szFile,0,sizeof(szFile));
   WideCharToMultiByte(CP_ACP,0,fileName,-1,szFile,MAX_PATH,0,0);
   DeleteFile(szFile);
   SysFreeString(fileName);
   fileName = NULL;
   temporaryFileInUse = false;
   return 0;
   }


   int Properties::setExistingFile(BSTR newFileName) {

   char szFileName[MAX_PATH];

   memset(szFileName,0,sizeof(szFileName));
   WideCharToMultiByte(CP_ACP,0,newFileName,-1,szFileName,MAX_PATH,0,0);

   WIN32_FIND_DATA findData;
   memset(&findData,0,sizeof(WIN32_FIND_DATA));
   HANDLE h = FindFirstFile(szFileName,&findData);
   if ( INVALID_HANDLE_VALUE == h )
      return 0;
   else 
      pIProperties -> put_FileName(newFileName);

   FindClose(h);

   return 1;
   }


   HRESULT Properties::createStorageAndStream(IStorage** ppIStorage,IStream** ppIStream,BSTR streamName) {

   currentItemIndex = -1;
   pCurrent_IO_Object = NULL;

   HRESULT hr = StgCreateDocfile(fileName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_FAILIFTHERE,0,ppIStorage);
   if ( STG_E_FILEALREADYEXISTS == hr ) {
      DeleteFileW(fileName);
      hr = StgCreateDocfile(fileName,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_FAILIFTHERE,0,ppIStorage);
   } else if ( S_OK != hr && debuggingEnabled ) {
      OLECHAR message[1024];
      wsprintfW(message,L"The Properties control is unable to save the file: %s",fileName);
      MessageBoxW(HWND_DESKTOP,message,L"Error!",MB_OK);
   }

   if ( S_OK != hr ) {
      if ( ppIStorage ) 
         if ( *ppIStorage )
            (*ppIStorage) -> Release();
      return hr;
   }
 
   hr = (*ppIStorage) -> CreateStream(streamName,STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,0,0,ppIStream);
   if ( S_OK != hr )  {
      if ( ppIStorage )
         if ( *ppIStorage )
            (*ppIStorage) -> Release();
   } else  {
      currentItemIndex = 0;
      pCurrent_IO_Object = this;
   }

   return hr;

   }


   HRESULT Properties::openStorageAndStream(IStorage** ppIStorage,IStream** ppIStream,BSTR streamName) {

   currentItemIndex = -1;
   pCurrent_IO_Object = NULL;

   HRESULT hr = StgOpenStorage(fileName,NULL,STGM_READ | STGM_SHARE_DENY_WRITE,NULL,0,ppIStorage);
   if ( S_OK != hr ) return hr;
 
   hr = (*ppIStorage) -> OpenStream(streamName,NULL,STGM_READ | STGM_SHARE_EXCLUSIVE,0,ppIStream);
   if ( S_OK != hr ) 
      (*ppIStorage) -> Release();
   else {
      currentItemIndex = 0;
      pCurrent_IO_Object = this;
   }

   return hr;
   }
