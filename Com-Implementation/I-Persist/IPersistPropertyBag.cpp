/*

                       Copyright (c) 1999,2000,2001 Nathan T. Clark

*/

#include <windows.h>

#include "Properties.h"


   Properties::_IPersistPropertyBag::_IPersistPropertyBag(Properties *pp) : 
    pParent(pp) {};

   Properties::_IPersistPropertyBag::~_IPersistPropertyBag() {};

   long __stdcall Properties::_IPersistPropertyBag::QueryInterface(REFIID riid,void **ppv) {
   return pParent -> InternalQueryInterface(riid,ppv);
   }
   unsigned long __stdcall Properties::_IPersistPropertyBag::AddRef() {
   return pParent -> InternalAddRef(); 
   }
   unsigned long __stdcall Properties::_IPersistPropertyBag::Release() {
   return pParent -> InternalRelease(); 
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag::GetClassID(CLSID *pcid) {
   return pParent -> pIPersistStorage -> GetClassID(pcid);
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag::InitNew() {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG;
   return pParent -> pIProperties -> InitNew(NULL);
   }


 
   STDMETHODIMP Properties::_IPersistPropertyBag::Load(IPropertyBag *is,IErrorLog* pErrLog) {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG;
   return pParent -> loadPropertyBag(is,NULL,pErrLog);
   }
 
 
   STDMETHODIMP Properties::_IPersistPropertyBag::Save(IPropertyBag *is,int fClearDirty,int fSaveAllProperties) {
   pParent -> currentPersistenceMechanism = MECHANISM_PROPERTYBAG;
   return pParent -> savePropertyBag(is,NULL,fClearDirty,fSaveAllProperties);
   }
