// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"


   Properties::_IPersistStreamInit::_IPersistStreamInit(Properties *pp) : 
    pParent(pp) {};

   Properties::_IPersistStreamInit::~_IPersistStreamInit() {};

   long __stdcall Properties::_IPersistStreamInit::QueryInterface(REFIID riid,void **ppv) {
   return pParent -> InternalQueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::_IPersistStreamInit::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IPersistStreamInit::Release() {
   return pParent -> InternalRelease(); 
   }
 
 
   STDMETHODIMP Properties::_IPersistStreamInit::GetClassID(CLSID *pcid) {
   return pParent -> pIPersistStorage -> GetClassID(pcid);
   }
 
 
   STDMETHODIMP Properties::_IPersistStreamInit::IsDirty() {
   return pParent -> pIProperties -> IsDirty();
   }
 
 
   STDMETHODIMP Properties::_IPersistStreamInit::Load(IStream *is) {
   pParent -> currentPersistenceMechanism = MECHANISM_STREAMINIT;
   pParent -> pIStream_current = is;
   HRESULT rv = pParent -> internalLoad(pParent -> pIStream_current,NULL);
   pParent -> pIStream_current = NULL;
   return rv;
   }
 
 
   STDMETHODIMP Properties::_IPersistStreamInit::Save(IStream *is,int fClearDirty) {
   pParent -> currentPersistenceMechanism = MECHANISM_STREAMINIT;
   pParent -> pIStream_current = is;
   HRESULT rv = pParent -> internalSave(pParent -> pIStream_current,NULL);
   pParent -> pIStream_current = NULL;
   return rv;
   }
 
 
   STDMETHODIMP Properties::_IPersistStreamInit::GetSizeMax(ULARGE_INTEGER *pb) {
   long tempLong;
   pParent -> pIProperties -> get_Size(&tempLong);
   pb -> QuadPart = tempLong;
   return S_OK;
   }
 
   STDMETHODIMP Properties::_IPersistStreamInit::InitNew() {
   pParent -> currentPersistenceMechanism = MECHANISM_STREAMINIT;
   return pParent -> pIProperties -> InitNew(NULL);
   }