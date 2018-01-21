// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

   long __stdcall Properties::QueryInterface(REFIID riid,void **ppv) {
   return pIUnknownOuter -> QueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::AddRef() {
   return pIUnknownOuter -> AddRef();
   }
   unsigned long __stdcall Properties::Release() {
   return pIUnknownOuter -> Release();
   }
 
 
