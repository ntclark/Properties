/*

                       Copyright (c) 1996,1997,1998,1999,2000,2001,2002 Nathan T. Clark

*/

#pragma once

   struct simplePropertyWindow {
      simplePropertyWindow() : hwnd(NULL),idStart(0),idOK(0),idApply(0),idCancel(0),pageName(NULL) {};
      HWND hwnd,hwndStart,hwndOK,hwndApply,hwndCancel;
      short idStart,idOK,idApply,idCancel;
      BSTR pageName;
      ~simplePropertyWindow() {
         if ( pageName ) SysFreeString(pageName);
      };
   };


   struct simplePersistenceWindow {
      simplePersistenceWindow() : 
         hwnd(NULL),
         hwndInit(0),hwndLoad(0),hwndSavePrep(0),hwndIsDirty(0),
         idSavePrep(0),idInit(0),idLoad(0),idIsDirty(0) {};
      HWND hwnd,hwndInit,hwndLoad,hwndSavePrep,hwndIsDirty;
      short idSavePrep,idInit,idLoad,idIsDirty;
   };


   struct propertyPageProvider {

      enum ppProviderType {
         undefined = 0,
         simpleWindow = 1,
         propertyPageClient = 2,
         specifyPropertyPages = 3};

      propertyPageProvider(ppProviderType t,IUnknown* pobj,long i,unsigned long sig) : 
         type(t), pIUnknownObject(pobj), index(i), signature(sig), pageCount(0) { };

      ~propertyPageProvider() { };

      IUnknown* pIUnknownObject;
      ppProviderType type;
      unsigned long signature;
      long index;
      long pageCount;
   };

#define PROPERTY_PAGE_COUNT  32

   STDAPI GSystemsPropertyPagesDllGetClassObject(HMODULE hModule,REFCLSID rclsid, REFIID riid, void** ppObject);
   STDAPI GSystemsPropertiesDllGetClassObject(HMODULE hModule,REFCLSID,REFIID riid,void** ppObject);
