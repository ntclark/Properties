// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#pragma once

#include <olectl.h>
#include <stdio.h>
#include <list>
#include <commctrl.h>

#include "GSystemsProperties.h"

#include "Properties_i.h"
#include "ObjectFactory.h"

#include "list.h"
#include "resource.h"

#ifdef CURSIVISION_BUILD
#define SYSTEM_NAME "CursiVision"
#else
#define SYSTEM_NAME "GSystem"
#endif

   extern ITypeInfo *pITypeInfo_IProperties;
   extern long gsProperties_oleMisc;
   extern ITypeInfo *pITypeInfo_IProperty;
   extern HMODULE gsProperties_hModule;

   extern char szPropertyType[][32];

   struct persistableObjectInterface {
      persistableObjectInterface(IUnknown* pObj,PersistenceMechanism pm) : 
         pIUnknown(pObj),
         persistenceMechanism(pm) { pIUnknown -> AddRef(); };
      ~persistableObjectInterface() { pIUnknown -> Release(); };
      IUnknown* pIUnknown;
      enum PersistenceMechanism persistenceMechanism;
   };


   struct storageInfo {
      BYTE registrationMark[8];
      PropertyType propertyType;
      long cBytes;
   };

      
   class Properties : IUnknown, List<interface IGProperty>  {
 
   public:
 
      Properties(IUnknown *pUnknownOuter);
      ~Properties();
 
//   IUnknown
 
      STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
      STDMETHOD_ (ULONG, AddRef)();
      STDMETHOD_ (ULONG, Release)();
 
      STDMETHOD(InternalQueryInterface)(REFIID,void**);
      STDMETHOD_ (ULONG,InternalAddRef)();
      STDMETHOD_ (ULONG,InternalRelease)();

      void ModernPropertySheet(PROPSHEETHEADER *);

      class XNDUnknown : public IUnknown {
         Properties* This() { return (Properties*)((BYTE*)this - offsetof(Properties,innerUnknown)); }
         STDMETHODIMP QueryInterface(REFIID r,void** p) { return This() -> InternalQueryInterface(r,p); }
         STDMETHODIMP_(ULONG) AddRef() { return This() -> InternalAddRef(); }
         STDMETHODIMP_(ULONG) Release() { return This() -> InternalRelease(); }
      } innerUnknown;

   private:

      class _IProperties : public IGProperties {
      public:
 
         _IProperties(Properties* pp);
         ~_IProperties();
 
      // IDispatch

         STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
         STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
         STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgdispid);
         STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);

      // IProperties
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD(put_DebuggingEnabled)(VARIANT_BOOL setEnabled);
         STDMETHOD(get_DebuggingEnabled)(VARIANT_BOOL* getEnabled);
         STDMETHOD(get_Count)(long *);
         STDMETHOD(get_Size)(long *);
         STDMETHOD(get_IStorage)(IStorage**);
         STDMETHOD(get_IStream)(IStream**);

         STDMETHOD(put_FileName)(BSTR);
         STDMETHOD(get_FileName)(BSTR*);
         STDMETHOD(put_FileAllowedExtensions)(BSTR);
         STDMETHOD(get_FileAllowedExtensions)(BSTR*);
         STDMETHOD(put_FileType)(BSTR);
         STDMETHOD(get_FileType)(BSTR*);
         STDMETHOD(put_FileSaveOpenText)(BSTR);
         STDMETHOD(get_FileSaveOpenText)(BSTR*);

         STDMETHOD(New)();
         STDMETHOD(Open)(BSTR* bstrFileName);
         STDMETHOD(LoadFile)(VARIANT_BOOL* wasSuccessful);
         STDMETHOD(OpenFile)(BSTR bstrFileName);
         STDMETHOD(Save)();
         STDMETHOD(SaveTo)(BSTR bstrFileName);
         STDMETHOD(SaveAs)(BSTR* bstrFileName);

         STDMETHOD(get_Property)(BSTR bstrPropertyname,IGProperty** ppIProperty);

         STDMETHOD(Add)(BSTR name,IGProperty** ppIProperty = NULL);
         STDMETHOD(Include)(IGProperty *);
         STDMETHOD(Remove)(BSTR name);
         STDMETHOD(AddObject)(IUnknown*);
         STDMETHOD(RemoveObject)(IUnknown*);
         STDMETHOD(Advise)(IGPropertiesClient *);
         STDMETHOD(DirectAccess)(BSTR propertyName,enum PropertyType,void* directAccess,long directAccessSize);
         STDMETHOD(SetClassID)(BYTE *);
         STDMETHOD(CopyTo)(IGProperties* pTheDestination);
         STDMETHOD(GetPropertyInterfaces)(long* pCntInterfaces,IGProperty*** pTheArray);

      // Persistence support

         STDMETHOD(PutHWNDPersistence)(HWND hwndPersistence,HWND vhwndInit,HWND vhwndLoad,HWND vhwndSavePrep);
         STDMETHOD(RemoveHWNDPersistence)(HWND hwndPersistence);
         STDMETHOD(IsDirty)();
         STDMETHOD(InitNew)(IStorage *);
         STDMETHOD(SaveToStorage)(IStream *,IStorage *);
         STDMETHOD(LoadFromStorage)(IErrorLog *,IStream *,IStorage *);
         STDMETHOD(SaveObjectToFile)(IUnknown*,BSTR);
         STDMETHOD(LoadObjectFromFile)(IUnknown*,BSTR);

      // Property Page support

         STDMETHOD(AdvisePropertyPageClient)(IGPropertyPageClient *,boolean usePropertySheets = FALSE);
         STDMETHOD(AddPropertyPage)(IUnknown *,boolean usePropertySheets = FALSE);
         STDMETHOD(RemovePropertyPage)(IUnknown *,boolean usePropertySheets = FALSE);
         STDMETHOD(PutHWNDPropertyPage)(BSTR displayName,HWND hwndProperties,HWND hwndStart,HWND hwndOk,HWND hwndApply,HWND hwndCancel);
         STDMETHOD(RemoveHWNDPropertyPage)(HWND hwndPropertyPage);
         STDMETHOD(ShowProperties)(HWND hwndOwner,IUnknown* pIUnknown);
         STDMETHOD(EditProperties)(HWND hwndOwner,BSTR strText,IUnknown* pIUnknown);
         STDMETHOD(Push)();
         STDMETHOD(Pop)();
         STDMETHOD(Discard)();
         STDMETHOD(Compare)(VARIANT_BOOL*);
         STDMETHOD(ConnectPropertyNotifySink)(IPropertyNotifySink* pContainerPropertyNotifySink,DWORD * pdwCookie);
         STDMETHOD(FindConnectionPoint)(REFIID riid,IConnectionPoint **);

      // Window contents support

         STDMETHOD(GetWindowValue)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(GetWindowItemValue)(BSTR propertyName,HWND hwndDialog,long idControl);
         STDMETHOD(GetWindowText)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(GetWindowItemText)(BSTR propertyName,HWND hwndDialog,long idControl);
         STDMETHOD(SetWindowText)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(SetWindowItemText)(BSTR propertyName,HWND hwndDialog,long idControl);

         // Combo boxes
         STDMETHOD(SetWindowComboBoxSelection)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(SetWindowItemComboBoxSelection)(BSTR propertyName,HWND hwndControl,long idControl);
         STDMETHOD(GetWindowComboBoxSelection)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(GetWindowItemComboBoxSelection)(BSTR propertyName,HWND hwndControl,long idControl);
         STDMETHOD(SetWindowComboBoxList)(BSTR propertyName,HWND hwndDialog);
         STDMETHOD(SetWindowItemComboBoxList)(BSTR propertyName,HWND hwndDialog,long idControl);
         STDMETHOD(GetWindowComboBoxList)(BSTR propertyName,HWND hwndDialog);
         STDMETHOD(GetWindowItemComboBoxList)(BSTR propertyName,HWND hwndDialog,long idControl);

         // List boxes
         STDMETHOD(SetWindowListBoxSelection)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(SetWindowItemListBoxSelection)(BSTR propertyName,HWND hwndControl,long idControl);
         STDMETHOD(GetWindowListBoxSelection)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(GetWindowItemListBoxSelection)(BSTR propertyName,HWND hwndControl,long idControl);
         STDMETHOD(SetWindowListBoxList)(BSTR propertyName,HWND hwndDialog);
         STDMETHOD(SetWindowItemListBoxList)(BSTR propertyName,HWND hwndDialog,long idControl);
         STDMETHOD(GetWindowListBoxList)(BSTR propertyName,HWND hwndDialog);
         STDMETHOD(GetWindowItemListBoxList)(BSTR propertyName,HWND hwndDialog,long idControl);

         // Vector and matrix properties
         STDMETHOD(GetWindowArrayValues)(BSTR propertyName,SAFEARRAY** hwndControl);
         STDMETHOD(GetWindowItemArrayValues)(BSTR propertyName,SAFEARRAY** hwndControl,SAFEARRAY** idControl);
         STDMETHOD(SetWindowArrayValues)(BSTR propertyName,SAFEARRAY** hwndControl);
         STDMETHOD(SetWindowItemArrayValues)(BSTR propertyName,SAFEARRAY** hwndControl,SAFEARRAY** idControl);

         // Check boxes
         STDMETHOD(SetWindowChecked)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(SetWindowItemChecked)(BSTR propertyName,HWND hwndDialog,long idControl);
         STDMETHOD(GetWindowChecked)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(GetWindowItemChecked)(BSTR propertyName,HWND hwndDialog,long idControl);

         // Enabled/Disabled
         STDMETHOD(SetWindowEnabled)(BSTR propertyName,HWND hwndControl);
         STDMETHOD(SetWindowItemEnabled)(BSTR propertyName,HWND hwndDialog,long idControl);

         // Other windows API support
         STDMETHOD(GetWindowID)(HWND hwndControl,long* theID);
 
       private:

         Properties* pParent;
 
         IUnknown *pParentUnknown;
 
      } *pIProperties;
 
//   IPersistStorage
 
      class _IPersistStorage : public IPersistStorage {
      public:
 
         _IPersistStorage(Properties *pp);
         ~_IPersistStorage();
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD(GetClassID)(CLSID *);
         STDMETHOD(IsDirty)();
         STDMETHOD(InitNew)(IStorage *);
         STDMETHOD(Load)(IStorage *);
         STDMETHOD(Save)(IStorage *,BOOL);
         STDMETHOD(SaveCompleted)(IStorage *);
         STDMETHOD(HandsOffStorage)();
 
      private:
 
         Properties *pParent;
 
         boolean noScribble;
 
      } *pIPersistStorage;
 
//   IPersist
 
      class _IPersistStream : public IPersistStream {
      public:
 
         _IPersistStream(Properties *pp);
         ~_IPersistStream();
 
//      IPersist
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD(GetClassID)(CLSID *);
 
//      IPersistStream
 
         STDMETHOD(IsDirty)();
         STDMETHOD(Load)(IStream *);
         STDMETHOD(Save)(IStream *,int);
         STDMETHOD(GetSizeMax)(ULARGE_INTEGER *);
 
      private:
 
         Properties *pParent;
 
      } * pIPersistStream;
 
      class _IPersistStreamInit : public IPersistStreamInit {
      public:
 
         _IPersistStreamInit(Properties *pp);
         ~_IPersistStreamInit();
 
//      IPersist
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD(GetClassID)(CLSID *);
 
//      IPersistStream
 
         STDMETHOD(IsDirty)();
         STDMETHOD(Load)(IStream *);
         STDMETHOD(Save)(IStream *,int);
         STDMETHOD(GetSizeMax)(ULARGE_INTEGER *);
 
//      IPersistStreamInit
 
         STDMETHOD(InitNew)();
 
      private:
 
         Properties *pParent;
 
      } * pIPersistStreamInit;
 
//   IPersistPropertyBag2
 
      class _IPersistPropertyBag : public IPersistPropertyBag {
      public:
 
         _IPersistPropertyBag(Properties *pp);
         ~_IPersistPropertyBag();
 
//      IPersist
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();

         STDMETHOD(GetClassID)(CLSID *);
 
//      IPersistPropertyBag
 
         STDMETHOD(InitNew)();
         STDMETHOD(Load)(IPropertyBag *is,IErrorLog* pErrLog);
         STDMETHOD(Save)(IPropertyBag *,int,int);
 
      private:
 
         Properties *pParent;
 
      } * pIPersistPropertyBag;


      class _IPersistPropertyBag2 : public IPersistPropertyBag2 {
      public:

         _IPersistPropertyBag2(Properties *pp);
         ~_IPersistPropertyBag2();
 
//      IPersist
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();

         STDMETHOD(GetClassID)(CLSID *);
 
//      IPersistPropertyBag2
 
         STDMETHOD(InitNew)();
         STDMETHOD(Load)(IPropertyBag2 *is,IErrorLog* pErrLog);
         STDMETHOD(Save)(IPropertyBag2 *,int,int);
         STDMETHOD(IsDirty)();

      private:

         Properties *pParent;
 
      } * pIPersistPropertyBag2;

      class _IGPropertyPageClient : public IGPropertyPageClient {

      public:
         
         _IGPropertyPageClient(Properties *pp) : pParent(pp) {};
         ~_IGPropertyPageClient();

//      IPropertyPageClient
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();

         STDMETHOD(BeforeAllPropertyPages)();
         STDMETHOD(GetPropertyPagesInfo)(long* countPages,SAFEARRAY** stringDescriptions,SAFEARRAY** pHelpDirs,SAFEARRAY** pSizes);
         STDMETHOD(CreatePropertyPage)(long indexNumber,long,RECT*,BOOL,long* hwndPropertyPage);
         STDMETHOD(IsPageDirty)(long,BOOL*);
         STDMETHOD(Help)(BSTR);
         STDMETHOD(TranslateAccelerator)(long,long*);
         STDMETHOD(Apply)();
         STDMETHOD(AfterAllPropertyPages)(BOOL);
         STDMETHOD(DestroyPropertyPage)(long indexNumber);

         STDMETHOD(GetPropertySheetHeader)(void *pHeader);
         STDMETHOD(get_PropertyPageCount)(long *pCount);
         STDMETHOD(GetPropertySheets)(void *pSheets);

      private:
   
         Properties* pParent;

      } * pIPropertyPageClient_ThisObject;

// Connection Points

      class _IConnectionPointContainer : IConnectionPointContainer {
      public:
 
         _IConnectionPointContainer(Properties * pp);
         ~_IConnectionPointContainer();

         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD(FindConnectionPoint)(REFIID riid,IConnectionPoint **);
         STDMETHOD(EnumConnectionPoints)(IEnumConnectionPoints **);

      private:

         Properties* pParent;
 
      } connectionPointContainer;


      struct _IConnectionPoint : IConnectionPoint {
      public:
 
         _IConnectionPoint(Properties * pp,REFIID outGoingInterfaceType);
         ~_IConnectionPoint();

         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD (GetConnectionInterface)(IID *);
         STDMETHOD (GetConnectionPointContainer)(IConnectionPointContainer **ppCPC);
         STDMETHOD (Advise)(IUnknown *pUnk,DWORD *pdwCookie);
         STDMETHOD (Unadvise)(DWORD);
         STDMETHOD (EnumConnections)(IEnumConnections **ppEnum);
 
         IUnknown *AdviseSink() { return adviseSink; };
 
      private:
 
         int getSlot();
         int findSlot(DWORD dwCookie);
 
         IUnknown *adviseSink;
         Properties * pParent;
         DWORD nextCookie;
         int countConnections,countLiveConnections;
 
         REFIID outGoingInterfaceType;
 
         CONNECTDATA *connections;
 
      } propertyNotifySinkConnectionPoint;
 
      struct _IEnumConnectionPoints : IEnumConnectionPoints {
      public:
 
         _IEnumConnectionPoints(Properties * pp,_IConnectionPoint **cp,int connectionPointCount);
        ~_IEnumConnectionPoints();
 
         STDMETHOD (QueryInterface)(REFIID riid,void **ppv);
         STDMETHOD_ (ULONG, AddRef)();
         STDMETHOD_ (ULONG, Release)();
 
         STDMETHOD (Next)(ULONG cConnections,IConnectionPoint **rgpcn,ULONG *pcFetched);
         STDMETHOD (Skip)(ULONG cConnections);
         STDMETHOD (Reset)();
         STDMETHOD (Clone)(IEnumConnectionPoints **);
 
      private:
 
         int cpCount,enumeratorIndex;
         Properties * pParent;
         _IConnectionPoint **connectionPoints;
 
      } *enumConnectionPoints;
 
      struct _IEnumConnections : public IEnumConnections {
      public:
 
         _IEnumConnections(IUnknown* pParentUnknown,ULONG cConnections,CONNECTDATA* paConnections,ULONG initialIndex);
         ~_IEnumConnections();
 
          STDMETHOD (QueryInterface)(REFIID, void **);
          STDMETHOD_ (ULONG,AddRef)();
          STDMETHOD_ (ULONG,Release)();
          STDMETHOD (Next)(ULONG, CONNECTDATA*, ULONG*);
          STDMETHOD (Skip)(ULONG);
          STDMETHOD (Reset)();
          STDMETHOD (Clone)(IEnumConnections**);
 
      private:
 
         ULONG       refCount;
         IUnknown    *pParentUnknown;
         ULONG       enumeratorIndex;
         ULONG       countConnections;
         CONNECTDATA *connections;
 
      } *enumConnections;


      STDMETHOD(internalLoad)(IStream* pIStream,IStorage* pIStorage);
      STDMETHOD(internalSave)(IStream* pIStream,IStorage* pIStorage);
      STDMETHOD(internalInitNew)(IStorage* pIStorage);
      STDMETHOD(loadPropertyBag)(IPropertyBag* pIPropertyBag,IPropertyBag2*,IErrorLog* pIErrorLog);
      STDMETHOD(savePropertyBag)(IPropertyBag* pIPropertyBag,IPropertyBag2* pIPropertyBag2,BOOL fClearDirty,BOOL fSaveAllProperties);

      STDMETHOD(push)(BYTE*);
      STDMETHOD(pop)(BYTE*);

      int prepOpenFileName();
      int parseOpenFileName(bool preserveExisting = false);

      int createTemporaryStorageFileName();
      int deleteTemporaryStorage();
      int setExistingFile(BSTR newFileName);

      HRESULT createStorageAndStream(IStorage** ppIStorage,IStream** ppIStream,BSTR streamName);
      HRESULT openStorageAndStream(IStorage** ppIStorage,IStream** ppIStream,BSTR streamName);

      IUnknown *pIUnknownOuter;

      IGProperty* getProperty(BSTR propertyName,char *szMethodName);
      IClassFactory *pProperty_IClassFactory;

      int refCount;
      int currentItemIndex;

      List<simplePersistenceWindow> simplePersistenceWindows;
      List<persistableObjectInterface> persistableObjectInterfaces;
      List<IUnknown*> persistableObjects;

      List<IGProperty> propertiesToRelease;

      short bPreviouslyGotStream;
      short supportIPersistStorage;
      short supportIPersistStream;
      VARIANT_BOOL debuggingEnabled;
      short temporaryFileInUse;
      short isDirty;
      enum PersistenceMechanism currentPersistenceMechanism;

      BSTR fileName,fileExtensions,fileType,fileSaveOpenText;
      OPENFILENAME openFileName;

      IGProperty* propertyFileName, *propertyFileSaveText, *propertyFileExtensions, *propertyFileType;
      IGProperty* propertyDebuggingEnabled;
      IStream* pIStream_current;
      IStorage* pIStorage_current;

      /*
      09/21/2009: This list of IGPropertyPageClient interfaces are the instances of that
      interface to use the new PROPSHEET capabilities of IGPropertyPageClient - in place of the
      unstable PropertyPages.ocx implementation
      */

      IGPropertyPageClient *pMasterPropertyPageClient;
      std::list<IGPropertyPageClient *> propertyPageClients;

      List<BYTE*> pushStackList;

      List<IGProperty> proxyList;

      IGPropertiesClient *pIPropertiesClient;

      CLSID objectCLSID;

      HWND hwndPropertySheet;
      PROPSHEETHEADER propertySheetHeader;
      PROPSHEETPAGE *pPropertySheetPages;
      long cxClientIdeal[16],cyClientIdeal[16];
      long cxSheetIdeal[16],cySheetIdeal[16];
      long propertyFrameInstanceCount;

      static Properties* pCurrent_IO_Object;

      static LRESULT EXPENTRY propertySheetHandler(HWND hwnd,UINT msg,WPARAM mp1,LPARAM mp2);

      static int CALLBACK preparePropertySheet(HWND hwnd,UINT uMsg,LPARAM lParam);
      static WNDPROC nativePropertySheetHandler;

      friend class _IProperties;
      friend class _IPersistStorage;
      friend class _IPersistStream;
      friend class _IPersistStreamInit;
      friend class _IPersistPropertyBag;
      friend class _IPersistPropertyBag2;
      friend class _ISpecifyPropertyPages;
      friend class _IGPropertyPageClient;
      friend class _IConnectionPointContainer;

   };

#include <pshpack1.h>

typedef struct DLGTEMPLATEEX {
    WORD dlgVer;
    WORD signature;
    DWORD helpID;
    DWORD exStyle;
    DWORD style;
    WORD cDlgItems;
    short x;
    short y;
    short cx;
    short cy;
} DLGTEMPLATEEX, *LPDLGTEMPLATEEX;

#include <poppack.h>
