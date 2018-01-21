// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"
#include "Properties.h"

#ifdef PROPERTIES_SAMPLE

   extern bool trialExpired;

#endif

   ObjectFactory::ObjectFactory(const CLSID &clsid) : theClass(clsid) { };


   HRESULT STDMETHODCALLTYPE ObjectFactory::QueryInterface(const struct _GUID &iid,void **ppv) {
   *ppv = NULL; 
   if ( iid != IID_IUnknown && iid != IID_IClassFactory ) 
      return E_NOINTERFACE; 
   *ppv = this; 
   AddRef(); 
   return S_OK; 
   }
   unsigned long __stdcall ObjectFactory::AddRef() {return 1;}
   unsigned long __stdcall ObjectFactory::Release() {return 1;}
 
 
   HRESULT STDMETHODCALLTYPE ObjectFactory::CreateInstance(IUnknown *pUnkOuter,REFIID riid,void **ppv) {

   HRESULT hres;
   *ppv = NULL;

   if ( theClass == CLSID_InnoVisioNateProperty ) {
      Property *p = new Property();
      hres = p -> QueryInterface(riid,ppv);
      if ( ! ppv ) 
         delete p;
      return hres;
   }

   if ( theClass == CLSID_InnoVisioNateProperties ) {
      if ( pUnkOuter && riid != IID_IUnknown ) 
         return E_INVALIDARG;
      Properties *p = new Properties(pUnkOuter);
      p -> InternalAddRef();
      hres = p -> InternalQueryInterface(riid,ppv);
      p -> InternalRelease();
      if ( ! ppv ) 
         delete p;
      return hres;
   }

   return E_FAIL;
   }
 
 
   HRESULT STDMETHODCALLTYPE ObjectFactory::LockServer(int) {
   return 0;
   }


