
#pragma once

#include <olectl.h>

#include "Properties.h"

   class PropertyPage;

   class ObjectFactory : public IClassFactory {
   public:
 
      ObjectFactory(const CLSID &clsid);

      HRESULT STDMETHODCALLTYPE QueryInterface(const struct _GUID &,void **);
      unsigned long __stdcall AddRef();
      unsigned long __stdcall Release();
      HRESULT STDMETHODCALLTYPE CreateInstance(struct IUnknown *,const struct _GUID &,void **);
      HRESULT STDMETHODCALLTYPE LockServer(int);

   private:
 
      CLSID theClass;

   };

