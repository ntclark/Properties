
#include "Properties.h"

#include "List.cpp"
#include "utils.h"

// Property page support

   HRESULT Properties::_IProperties::AdvisePropertyPageClient(IGPropertyPageClient *pir,boolean usePropertySheets) {

   if ( usePropertySheets ) {
      pParent -> pMasterPropertyPageClient = pir;
      pir -> AddRef();
      return S_OK;
   }

   return S_OK;
   }
   

   HRESULT Properties::_IProperties::AddPropertyPage(IUnknown* pIUnknown,boolean usePropertySheets) {

   if ( ! pIUnknown ) return E_POINTER;
   
   if ( usePropertySheets ) {
      IGPropertyPageClient* pIPropertyPageClient;
      HRESULT hr = pIUnknown -> QueryInterface(IID_IGPropertyPageClient,reinterpret_cast<void**>(&pIPropertyPageClient));
      if ( S_OK == hr )
         pParent -> propertyPageClients.insert(pParent -> propertyPageClients.end(),pIPropertyPageClient);
      return hr;
   }

   return S_OK;
   }


   HRESULT Properties::_IProperties::RemovePropertyPage(IUnknown* pIUnknown,boolean usePropertySheets) {

   if ( ! pIUnknown ) return E_POINTER;
   
   if ( usePropertySheets ) {
      IGPropertyPageClient* pIPropertyPageClient;
      HRESULT hr = pIUnknown -> QueryInterface(IID_IGPropertyPageClient,reinterpret_cast<void**>(&pIPropertyPageClient));
      pParent->propertyPageClients.remove(pIPropertyPageClient);
      pIPropertyPageClient -> Release();
      return S_OK;
   }

   return S_OK;
   }


   HRESULT Properties::_IProperties::ShowProperties(HWND hwndOwner,IUnknown *pIUnknownObject) {

   IGPropertyPageClient *pIPropertyPageClient = NULL;

   if ( pIUnknownObject ) {

      pIUnknownObject -> QueryInterface(IID_IGPropertyPageClient,reinterpret_cast<void **>(&pIPropertyPageClient));

      PROPSHEETPAGE *pPropSheetPages = NULL;

      long countPages = 1;

      pIPropertyPageClient -> get_PropertyPageCount(&countPages);

      pPropSheetPages = new PROPSHEETPAGE[countPages];

      memset(pPropSheetPages,0,(countPages * sizeof(PROPSHEETPAGE)));

      pIPropertyPageClient -> GetPropertySheets(reinterpret_cast<void *>(pPropSheetPages));

      PROPSHEETHEADER propSheetHeader;

      memset(&propSheetHeader,0,sizeof(PROPSHEETHEADER));

      pIPropertyPageClient -> GetPropertySheetHeader(reinterpret_cast<void *>(&propSheetHeader));

      propSheetHeader.dwSize = sizeof(PROPSHEETHEADER);
      propSheetHeader.hwndParent = (HWND)hwndOwner;
      propSheetHeader.ppsp = pPropSheetPages;
      propSheetHeader.nPages = countPages;
      propSheetHeader.nStartPage = 0;

      pParent -> ModernPropertySheet(&propSheetHeader);

      delete [] pPropSheetPages;

      pIPropertyPageClient -> Release();

      return S_OK;

   }

   long countPages = 0;

   PROPSHEETPAGE *pPropSheetPages = NULL;

   if ( pParent -> pMasterPropertyPageClient ) 
      pParent -> pMasterPropertyPageClient -> get_PropertyPageCount(&countPages);

   for ( std::list<IGPropertyPageClient *>::iterator it = pParent -> propertyPageClients.begin(); it != pParent -> propertyPageClients.end(); it++ ) {

      IGPropertyPageClient *pIPropertyPageClient = *(it);

      long deltaPages = 0;

      pIPropertyPageClient -> get_PropertyPageCount(&deltaPages);

      countPages += deltaPages;

   }

   long currentPageIndex = 0;

   if ( pParent -> pMasterPropertyPageClient )
      pParent -> pMasterPropertyPageClient -> get_PropertyPageCount(&currentPageIndex);

   pPropSheetPages = new PROPSHEETPAGE[countPages];
   memset(pPropSheetPages,0,(countPages * sizeof(PROPSHEETPAGE)));

   if ( pParent -> pMasterPropertyPageClient )
      pParent -> pMasterPropertyPageClient -> GetPropertySheets(reinterpret_cast<void *>(pPropSheetPages));

   for ( std::list<IGPropertyPageClient *>::iterator it = pParent -> propertyPageClients.begin(); it != pParent -> propertyPageClients.end(); it++ ) {

      IGPropertyPageClient *pIPropertyPageClient = *(it);

      pIPropertyPageClient -> GetPropertySheets(reinterpret_cast<void *>(&pPropSheetPages[currentPageIndex]));

      long deltaPages = 0;

      pIPropertyPageClient -> get_PropertyPageCount(&deltaPages);

      currentPageIndex += deltaPages;

   }

   PROPSHEETHEADER propSheetHeader;

   memset(&propSheetHeader,0,sizeof(PROPSHEETHEADER));

   if ( pParent -> pMasterPropertyPageClient )
      pParent -> pMasterPropertyPageClient -> GetPropertySheetHeader(reinterpret_cast<void *>(&propSheetHeader));
   else {
      propSheetHeader.dwFlags = PSH_PROPSHEETPAGE | PSH_NOCONTEXTHELP;
      propSheetHeader.hInstance = gsProperties_hModule;
      propSheetHeader.pszIcon = 0L;
      propSheetHeader.pszCaption = "Settings";
      propSheetHeader.pfnCallback = NULL;
   }

   propSheetHeader.dwSize = sizeof(PROPSHEETHEADER);

   propSheetHeader.hwndParent = (HWND)hwndOwner;

#ifdef USE_MODERN_PROPERTYSHEETS
   propSheetHeader.ppsp = pPropSheetPages;
#endif

   propSheetHeader.nPages = countPages;

   propSheetHeader.nStartPage = 0;

   pParent -> ModernPropertySheet(&propSheetHeader);

   delete [] pPropSheetPages;

   return S_OK;
   }


   HRESULT Properties::_IProperties::EditProperties(HWND hwndO,BSTR szwName,IUnknown *pIUnknownObject) {

   HWND hwndOwner = hwndO;
   CAUUID pages;
   memset(&pages,0,sizeof(CAUUID));

   if ( pIUnknownObject ) {
      
      ISpecifyPropertyPages* pISpecifyPropertyPages;

      HRESULT hr = pIUnknownObject -> QueryInterface(IID_ISpecifyPropertyPages,reinterpret_cast<void**>(&pISpecifyPropertyPages));

      if ( FAILED(hr) ) {
         if ( pParent -> debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"An attempt was made to Edit the properties of an object that does not support the IID_ISpecifyPropertyPages interface.\n\nMethod IProperties::EditProperties");
            MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
         }
         return hr;
      }

      pISpecifyPropertyPages -> GetPages(&pages);

      IUnknown *pIUnknowns[32];// = NULL;
      memset(pIUnknowns,0,32 * sizeof(IUnknown *));

//      pParent -> pMasterPropertyPageClient -> GetPropertyPageObjectIUnknowns(pIUnknowns);

      OleCreatePropertyFrame(hwndOwner,0,0,szwName,pages.cElems,pIUnknowns,pages.cElems,pages.pElems,NULL,0,NULL);

      pISpecifyPropertyPages -> Release();

   } else {

#if 0 // do something here
      if ( pParent -> pIPropertyPageSupport ) {

         pParent -> pIPropertyPageSupport -> GetPages(&pages);
         IUnknown** ppPropertyPageUnknowns = pParent -> pIPropertyPageSupport -> GetPropertyPageUnknowns();
         OleCreatePropertyFrame(hwndOwner,0,0,szwName,pages.cElems,ppPropertyPageUnknowns,pages.cElems,pages.pElems,NULL,0,NULL);
         pParent -> pIPropertyPageSupport -> ClearPropertyPageUnknowns();

      }
#endif

   }

   CoTaskMemFree(pages.pElems);

   return S_OK;
   }


   HRESULT Properties::_IProperties::Push() {
   long cb;
   get_Size(&cb);
   cb += 4 * (pParent -> Count() + pParent -> proxyList.Count());
   BYTE *b = new BYTE[cb];
   pParent -> pushStackList.Add(new BYTE*(b));
   pParent -> push(b);
   return S_OK;
   }


   HRESULT Properties::_IProperties::Pop() {
   if ( pParent -> pushStackList.Count() < 1 ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt to Pop beyond the end of the stack was made. i.e., more Pops or Discards than Pushes. Method IProperties::Pop");
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   BYTE** pb = pParent -> pushStackList.GetLast();
   pParent -> pop(*pb);
   pParent -> pushStackList.Remove(pb);
   delete [] *pb;
   delete pb;
   return S_OK;
   }


   HRESULT Properties::_IProperties::Discard() {
   if ( pParent -> pushStackList.Count() < 1 ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt to Discard beyond the end of the stack was made. i.e., more Pops or Discards than Pushes. Method IProperties::Discard");
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   BYTE** pb = pParent -> pushStackList.GetLast();
   if ( pb ) {
      pParent -> pushStackList.Remove(pb);
      delete [] *pb;
      delete pb;
   }
   return S_OK;
   }

   
   HRESULT Properties::_IProperties::Compare(VARIANT_BOOL* bIsDifferent) {
   if ( ! bIsDifferent ) return E_POINTER;
   *bIsDifferent = FALSE;
   if ( pParent -> pushStackList.Count() < 1 ) {
      if ( pParent -> debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt to Compare current properties when there is no set of properties on the stack. The \"Compare\" is intended to be between the \"current\" properties and the set of properties on a stack. Method IProperties::Compare");
         MessageBox(NULL,szError,"Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   int hr;
   long cb;
   get_Size(&cb);
   cb += 4 * (pParent -> Count() + pParent -> proxyList.Count());
   BYTE** pbOld = pParent -> pushStackList.GetLast();
   BYTE* pbNew = new BYTE[cb];
   pParent -> push(pbNew);
   hr = memcmp(*pbOld,pbNew,cb) == 0 ? S_OK : S_FALSE;
   delete [] pbNew;
   *bIsDifferent = hr ? FALSE : TRUE;
   return S_OK;
   }


   HRESULT Properties::_IProperties::ConnectPropertyNotifySink(IPropertyNotifySink* p,DWORD* pdwCookie) {
   IConnectionPoint* pIConnectionPoint;
   pParent -> connectionPointContainer.FindConnectionPoint(IID_IPropertyNotifySink,&pIConnectionPoint);
   pIConnectionPoint -> Advise(p,pdwCookie);
   return S_OK;
   }


   STDMETHODIMP Properties::_IProperties::FindConnectionPoint(REFIID riid,IConnectionPoint **ppCP) {
   return pParent -> connectionPointContainer.FindConnectionPoint(riid,ppCP);
   }


   STDMETHODIMP Properties::_IProperties::put_AllowSysMenu(boolean doAllow) {
   pParent -> allowSysMenu = doAllow;
   return S_OK;
   }


   STDMETHODIMP Properties::_IProperties::put_FrameSize(SIZEL frameSize) {
   pParent -> cxSheetIdeal[pParent -> propertyFrameInstanceCount] = max(pParent -> cxSheetIdeal[pParent -> propertyFrameInstanceCount],frameSize.cx);
   pParent -> cySheetIdeal[pParent -> propertyFrameInstanceCount] = max(pParent -> cySheetIdeal[pParent -> propertyFrameInstanceCount],frameSize.cy);
   SetWindowPos(pParent -> hwndPropertySheet,HWND_TOP,0,0,pParent -> cxSheetIdeal[pParent -> propertyFrameInstanceCount],pParent -> cySheetIdeal[pParent -> propertyFrameInstanceCount],SWP_NOMOVE);
   return S_OK;
   }


   STDMETHODIMP Properties::_IProperties::put_PageSize(SIZEL frameSize) {
   pParent -> cxClientIdeal[pParent -> propertyFrameInstanceCount] = max(pParent -> cxClientIdeal[pParent -> propertyFrameInstanceCount],frameSize.cx);
   pParent -> cyClientIdeal[pParent -> propertyFrameInstanceCount] = max(pParent -> cyClientIdeal[pParent -> propertyFrameInstanceCount],frameSize.cy);
   for ( unsigned long k = 0; k < pParent -> propertySheetHeader.nPages; k++ ) {
      HWND hwndPage = (HWND)SendMessage(pParent -> hwndPropertySheet,PSM_INDEXTOHWND,(LPARAM)k,0L);
      SetWindowPos(hwndPage,HWND_TOP,TREEVIEW_WIDTH + 24,16,
                            pParent -> cxClientIdeal[pParent -> propertyFrameInstanceCount],pParent -> cyClientIdeal[pParent -> propertyFrameInstanceCount],SWP_NOACTIVATE);
   }
   return S_OK;
   }


   HRESULT Properties::_IProperties::PutHWNDPropertyPage(BSTR displayName,HWND lhwndProp,HWND lhwndStart,HWND lhwndOK,HWND lhwndApply,HWND lhwndCancel) {
   return E_NOTIMPL;
   }


   HRESULT Properties::_IProperties::RemoveHWNDPropertyPage(HWND lhwndProp) {
   return E_NOTIMPL;
   }