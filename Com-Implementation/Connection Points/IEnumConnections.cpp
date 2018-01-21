// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

   Properties::_IEnumConnections::_IEnumConnections(
             IUnknown* pHostObj,
             ULONG cConnections,
             CONNECTDATA* paConnections,
             ULONG initialIndex) :
     refCount(0), 
     pParentUnknown(pHostObj),
     enumeratorIndex(initialIndex),
     countConnections(cConnections)
   {
 
   connections = new CONNECTDATA[countConnections];
 
   for ( UINT i = 0; i < countConnections; i++ ) {
      connections[i] = paConnections[i];
   }
 
   return;
   }
 
 
 
   Properties::_IEnumConnections::~_IEnumConnections() {
   delete [] connections;
   return;
   }
 
 
   STDMETHODIMP Properties::_IEnumConnections::QueryInterface(REFIID riid,void **ppv) {
   *ppv = NULL;
   if ( IID_IUnknown != riid && IID_IEnumConnections != riid) return E_NOINTERFACE;
   *ppv = (LPVOID)this;
   AddRef();
   return S_OK;
   }
 
 
   STDMETHODIMP_(ULONG) Properties::_IEnumConnections::AddRef() {
   pParentUnknown -> AddRef();
   return ++refCount;
   }
 
 
 
   STDMETHODIMP_(ULONG) Properties::_IEnumConnections::Release() {
   pParentUnknown -> Release();
   if ( 0 == --refCount ) {
     refCount++;
     delete this;
     return 0;
   }
   return refCount;
   }
 

 
   STDMETHODIMP Properties::_IEnumConnections::Next(ULONG countRequested,CONNECTDATA* pConnections,ULONG* countEnumerated) {
   ULONG countReturned;

   if ( NULL == pConnections ) return E_POINTER;

   for ( countReturned = 0; 
         enumeratorIndex < countConnections && countRequested > 0; 
         pConnections++, enumeratorIndex++, countReturned++, countRequested-- ) {
      *pConnections = connections[enumeratorIndex];
      pConnections -> pUnk -> AddRef();
   }

   if ( NULL != countEnumerated )
      *countEnumerated = countReturned;

   return countReturned > 0 ? S_OK : S_FALSE;
   }


   STDMETHODIMP Properties::_IEnumConnections::Skip(ULONG cSkip) {
   if ( (enumeratorIndex + cSkip) < countConnections ) return S_FALSE;
   enumeratorIndex += cSkip;
   return S_OK;
   }
 
 
 
   STDMETHODIMP Properties::_IEnumConnections::Reset() {
   enumeratorIndex = 0;
   return S_OK;
   }
 
 
 
   STDMETHODIMP Properties::_IEnumConnections::Clone(IEnumConnections** ppIEnum) {
   _IEnumConnections* p = new _IEnumConnections(pParentUnknown,countConnections,connections,enumeratorIndex);
   return p -> QueryInterface(IID_IEnumConnections,(void **)ppIEnum);
   }
