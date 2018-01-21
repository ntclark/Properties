// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"


   Properties::_IPersistStream::_IPersistStream(Properties *pp) : 
    pParent(pp) {};

   Properties::_IPersistStream::~_IPersistStream() {};

   long __stdcall Properties::_IPersistStream::QueryInterface(REFIID riid,void **ppv) {
   return pParent -> InternalQueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::_IPersistStream::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IPersistStream::Release() {
   return pParent -> InternalRelease(); 
   }
 
 
   STDMETHODIMP Properties::_IPersistStream::GetClassID(CLSID *pcid) {
   return pParent -> pIPersistStorage -> GetClassID(pcid);
   }
 
 
   STDMETHODIMP Properties::_IPersistStream::IsDirty() {
   return pParent -> pIProperties -> IsDirty();
   }
 
 
   STDMETHODIMP Properties::_IPersistStream::Load(IStream *is) {
   pParent -> currentPersistenceMechanism = MECHANISM_STREAM;
   pParent -> pIStream_current = is;
   HRESULT rv = pParent -> internalLoad(pParent -> pIStream_current,NULL);
   pParent -> pIStream_current = NULL;
   return rv;
   }
 
 
   STDMETHODIMP Properties::_IPersistStream::Save(IStream *is,int fClearDirty) {
   pParent -> currentPersistenceMechanism = MECHANISM_STREAM;
   pParent -> pIStream_current = is;
   HRESULT rv = pParent -> internalSave(pParent -> pIStream_current,NULL);
   pParent -> pIStream_current = NULL;
   return rv;
   }
 
 
   STDMETHODIMP Properties::_IPersistStream::GetSizeMax(ULARGE_INTEGER *pb) {
   long tempLong;
   pParent -> pIProperties -> get_Size(&tempLong);
   pb -> QuadPart = tempLong;
   return S_OK;
   }
