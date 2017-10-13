/*

                       Copyright (c) 1996,1997,1998,1999,2000,2001 Nathan T. Clark

*/

#include <windows.h>

#include "utils.h"

#include "Properties.h"

   Properties::_IConnectionPoint::_IConnectionPoint(Properties *pp,REFIID outgoingIFace) : 
      pParent(pp), 
      adviseSink(0),
      nextCookie(400),
      outGoingInterfaceType(outgoingIFace),
      countLiveConnections(0),
      countConnections(ALLOC_CONNECTIONS)
   { 
   connections = new CONNECTDATA[countConnections];
   memset(connections, 0, countConnections * sizeof(CONNECTDATA));
   return;
   };
 
 
   Properties::_IConnectionPoint::~_IConnectionPoint() {
   for ( int k = 0; k < countConnections; k++ ) 
      if ( connections[k].pUnk ) connections[k].pUnk -> Release();
   delete [] connections;
   return;
   }
 
 
   HRESULT Properties::_IConnectionPoint::QueryInterface(REFIID riid,void **ppv) {
   *ppv = NULL; 
  
    if ( riid == IID_IUnknown )
       *ppv = static_cast<void*>(this); 
    else
 
       if ( riid == IID_IConnectionPoint ) 
          *ppv = static_cast<IConnectionPoint *>(this);
       else
 
          return pParent -> QueryInterface(riid,ppv);
    
   AddRef();
   return S_OK;
   }
 
   STDMETHODIMP_(ULONG) Properties::_IConnectionPoint::AddRef() { return 1; }
 
   STDMETHODIMP_(ULONG) Properties::_IConnectionPoint::Release() { return 1; }
 
 
   STDMETHODIMP Properties::_IConnectionPoint::GetConnectionInterface(IID *pIID) {
   if ( pIID == 0 ) return E_POINTER;
   *pIID = outGoingInterfaceType;
   return S_OK;
   }
 
 
   STDMETHODIMP Properties::_IConnectionPoint::GetConnectionPointContainer(IConnectionPointContainer **ppCPC) {
   return pParent -> QueryInterface(IID_IConnectionPointContainer,(void **)ppCPC);
   }
 
 
   STDMETHODIMP Properties::_IConnectionPoint::Advise(IUnknown *pUnkSink,DWORD *pdwCookie) {
   HRESULT hr;
   IUnknown* pISink = 0;
 
   hr = pUnkSink -> QueryInterface(outGoingInterfaceType,(void **)&pISink);
 
   if ( hr == E_NOINTERFACE ) return CONNECT_E_NOCONNECTION;
   if ( ! SUCCEEDED(hr) ) return hr; 
   if ( ! pISink ) return CONNECT_E_CANNOTCONNECT;
 
   int freeSlot;
   *pdwCookie = 0;
 
   freeSlot = getSlot();
 
   pISink -> AddRef();
 
   connections[freeSlot].pUnk = pISink;
   connections[freeSlot].dwCookie = nextCookie;
 
   *pdwCookie = nextCookie++;
 
   countLiveConnections++;
 
   return S_OK;
   }
 
 
   STDMETHODIMP Properties::_IConnectionPoint::Unadvise(DWORD dwCookie) {
 
   if ( 0 == dwCookie ) return CONNECT_E_NOCONNECTION;
 
   int slot = findSlot(dwCookie);
 
   if ( slot == -1 ) return CONNECT_E_NOCONNECTION;
 
   if ( connections[slot].pUnk ) connections[slot].pUnk -> Release();
 
   connections[slot].dwCookie = 0;
 
   countLiveConnections--;
 
   return S_OK;
   }
 
   STDMETHODIMP Properties::_IConnectionPoint::EnumConnections(IEnumConnections **ppEnum) {
   CONNECTDATA *tempConnections;
   int i,j;
 
   *ppEnum = NULL;
 
   if ( countLiveConnections == 0 ) return OLE_E_NOCONNECTION;
 
   tempConnections = new CONNECTDATA[countLiveConnections];
 
   for ( i = 0, j = 0; i < countConnections && j < countLiveConnections; i++) {
 
     if ( 0 != connections[i].dwCookie ) {
       tempConnections[j].pUnk = (IUnknown *)connections[i].pUnk;
       tempConnections[j].dwCookie = connections[i].dwCookie;
       j++;
     }
   }
 
   _IEnumConnections *p = new _IEnumConnections(this,countLiveConnections,tempConnections,0);
   p -> QueryInterface(IID_IEnumConnections,(void **)ppEnum);
 
   delete [] tempConnections;
 
   return S_OK;
   }
 
 
   int Properties::_IConnectionPoint::getSlot() {
   CONNECTDATA* moreConnections;
   int i;
   i = findSlot(0);
   if ( i > -1 ) return i;
   moreConnections = new CONNECTDATA[countConnections + ALLOC_CONNECTIONS];
   memset( moreConnections, 0, sizeof(CONNECTDATA) * (countConnections + ALLOC_CONNECTIONS));
   memcpy( moreConnections, connections, sizeof(CONNECTDATA) * countConnections);
   delete [] connections;
   connections = moreConnections;
   countConnections += ALLOC_CONNECTIONS;
   return countConnections - ALLOC_CONNECTIONS;
   }
 
 
   int Properties::_IConnectionPoint::findSlot(DWORD dwCookie) {
   for ( int i = 0; i < countConnections; i++ )
      if ( dwCookie == connections[i].dwCookie ) return i;
   return -1;
   }