// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

   Properties::_IPersistPropertyBag2::_IPersistPropertyBag2(Properties *pp) : 
    pParent(pp) {};

   Properties::_IPersistPropertyBag2::~_IPersistPropertyBag2() {};

   long __stdcall Properties::_IPersistPropertyBag2::QueryInterface(REFIID riid,void **ppv) {
   return pParent -> InternalQueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::_IPersistPropertyBag2::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IPersistPropertyBag2::Release() {
   return pParent -> InternalRelease(); 
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag2::GetClassID(CLSID *pcid) {
   return pParent -> pIPersistStorage -> GetClassID(pcid);
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag2::InitNew() {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG2;
   return pParent -> pIProperties -> InitNew(NULL);
   }


   STDMETHODIMP Properties::_IPersistPropertyBag2::IsDirty() {
   return pParent -> pIProperties -> IsDirty();
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag2::Load(IPropertyBag2 *is,IErrorLog* pErrLog) {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG2;
   return pParent -> loadPropertyBag(NULL,is,pErrLog);
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag2::Save(IPropertyBag2 *is,int fClearDirty,int fSaveAllProperties) {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG2;
   return pParent -> savePropertyBag(NULL,is,fClearDirty,fSaveAllProperties);
   }
