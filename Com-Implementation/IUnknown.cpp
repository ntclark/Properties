/*

                       Copyright (c) 1999,2000,2001,2009 Nathan T. Clark

*/
#include <windows.h>

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
 
 
