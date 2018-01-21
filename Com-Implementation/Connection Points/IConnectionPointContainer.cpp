// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"


   Properties::_IConnectionPointContainer::_IConnectionPointContainer(Properties *pp) : 
      pParent(pp) 
   { 
   };
 
 
   Properties::_IConnectionPointContainer::~_IConnectionPointContainer() {}
 
 
   HRESULT Properties::_IConnectionPointContainer::QueryInterface(REFIID riid,void **ppv) {
   *ppv = NULL; 
  
    if ( riid == IID_IUnknown )
       *ppv = static_cast<void*>(this); 
    else
 
       if ( riid == IID_IConnectionPointContainer ) 
          *ppv = static_cast<IConnectionPointContainer *>(this);
       else
 
          return pParent -> QueryInterface(riid,ppv);
    
   AddRef();
   return S_OK;
   }
 
   STDMETHODIMP_(ULONG) Properties::_IConnectionPointContainer::AddRef() { return 1; }
 
   STDMETHODIMP_(ULONG) Properties::_IConnectionPointContainer::Release() { return 1; }
 
 
   STDMETHODIMP Properties::_IConnectionPointContainer::EnumConnectionPoints(IEnumConnectionPoints **ppEnum) {
   _IConnectionPoint* connectionPoints[1];
   *ppEnum = NULL;
 
   if ( pParent -> enumConnectionPoints ) delete pParent -> enumConnectionPoints;
 
   connectionPoints[0] = &pParent -> propertyNotifySinkConnectionPoint;
   pParent -> enumConnectionPoints = new _IEnumConnectionPoints(pParent,connectionPoints,1);

   return pParent -> enumConnectionPoints -> QueryInterface(IID_IEnumConnectionPoints,(void **)ppEnum);
   }
 
 
   STDMETHODIMP Properties::_IConnectionPointContainer::FindConnectionPoint(REFIID riid,IConnectionPoint **ppCP) {
   *ppCP = NULL;
 
   if ( riid == IID_IPropertyNotifySink ) 
      return pParent -> propertyNotifySinkConnectionPoint.QueryInterface(IID_IConnectionPoint,(void**)ppCP);
  
   return CONNECT_E_NOCONNECTION;
   }
 
