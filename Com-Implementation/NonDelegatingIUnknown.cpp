// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

   long __stdcall Properties::InternalQueryInterface(REFIID riid,void **ppv) {

   *ppv = NULL;
  
   if ( riid == IID_IUnknown ) 
      *ppv = static_cast<void *>(&innerUnknown);
   else

      if ( riid == IID_IDispatch ) 
         *ppv = static_cast<void *>(pIProperties);
      else

      if ( riid == IID_IConnectionPointContainer ) 
         *ppv = static_cast<void *>(&connectionPointContainer);
      else

      if ( riid == IID_IGProperties )
         *ppv = static_cast<void *>(pIProperties);
      else

      if ( riid == IID_IPersistStorage  )
         *ppv = static_cast<void *>(pIPersistStorage);
      else

      if ( riid == IID_IPersistStream )
         *ppv = static_cast<void *>(pIPersistStream);
      else
 
      if ( riid == IID_IPersistStreamInit )
         *ppv = static_cast<void *>(pIPersistStreamInit);
      else

      if ( riid == IID_IPersistPropertyBag )
         *ppv = static_cast<void *>(pIPersistPropertyBag);
      else

      if ( riid == IID_IPersistPropertyBag2 )
         *ppv = static_cast<void *>(pIPersistPropertyBag2);
      else

      if ( riid == IID_IGPropertyPageClient ) 
         *ppv = static_cast<void *>(pIPropertyPageClient_ThisObject);
      else

         if ( riid == IID_IGPropertiesClient ) {
            if ( pIPropertiesClient ) 
               *ppv = static_cast<void *>(pIPropertiesClient);
            else
               return E_NOINTERFACE;
         }
         else
            return E_NOINTERFACE;
 
   static_cast<IUnknown*>(*ppv) -> AddRef(); 
   return S_OK; 
   }
   unsigned long __stdcall Properties::InternalAddRef() {
   return ++refCount;
   }
   unsigned long __stdcall Properties::InternalRelease() {
   if ( 0 == --refCount  ) {
      delete this;
      return 0;
   }
   return refCount;
   }
 
 
