/*

                       Copyright (c) 1996,1997,1998,1999,2000,2001 Nathan T. Clark

*/

#include <windows.h>
#include <stdio.h>

#include "Property.h"
#include "utils.h"


   HRESULT Property::put_variantValue(VARIANT newValue) {

   everAssigned = true;

   if ( newValue.vt & (VARTYPE)VT_ARRAY )
      return put_arrayValue(newValue.parray);

   switch ( newValue.vt ) {
   case VT_UI1:
      return put_longValue(newValue.bVal);
   case VT_UI2:
      return put_longValue(newValue.uiVal);
   case VT_UI4:
      return put_longValue(newValue.ulVal);
   case VT_I1:
      return put_longValue(newValue.cVal);
   case VT_I2:
      return put_longValue(newValue.iVal);
   case VT_I4:
      return put_longValue(newValue.lVal);
   case VT_INT:
      return put_longValue(newValue.intVal);
   case VT_UINT:
      return put_longValue(newValue.uintVal);
   case VT_R4:
      return put_doubleValue(newValue.fltVal);
   case VT_R8:
      return put_doubleValue(newValue.dblVal);
   case VT_BOOL:
      return put_boolValue(newValue.boolVal);
   case VT_BSTR:
      return put_stringValue(newValue.bstrVal);
   case VT_BYREF | VT_I1:
      return put_szValue(newValue.pcVal);
   case VT_BYREF | VT_DECIMAL:
   case VT_UNKNOWN:
   case VT_DISPATCH:
      break;
   case VT_BYREF | VT_UI1:
      return put_binaryValue(v.binarySize,newValue.pbVal);
   }
   return S_OK;
   }


   HRESULT Property::get_variantValue(VARIANT* pValue) {

   if ( ! pValue ) return E_POINTER;

   VARIANT vThis = {VT_EMPTY};
   GVariantClear(&vThis);

   updateFromDirectAccess();

   switch ( type ) {
   case TYPE_BOOL:
      vThis.vt = VT_BOOL;
      get_boolValue(&vThis.boolVal);
      break;

   case TYPE_LONG:
      vThis.vt = VT_I4;
      get_longValue(&vThis.lVal);
      break;

   case TYPE_DOUBLE:
      vThis.vt = VT_R8;
      get_doubleValue(&vThis.dblVal);
      break;

   case TYPE_SZSTRING:
   case TYPE_STRING:
      vThis.vt = VT_BSTR;
      get_stringValue(&vThis.bstrVal);
      break;

   case TYPE_RAW_BINARY:
   case TYPE_BINARY: {
      pValue -> vt = VT_BYREF | VT_UI1;
      pValue -> pbVal = (BYTE*)v.binaryValue;
      }
      break;

   case TYPE_ARRAY:
      pValue -> vt = VT_ARRAY | VT_VARIANT;
      get_arrayValue(&pValue -> parray);
      break;

   case TYPE_UNSPECIFIED:
      pValue -> vt = VT_BSTR;
      pValue -> bstrVal = SysAllocString(L"");
      break;

   }

   if ( vThis.vt != VT_EMPTY ) {
      GVariantClear(pValue);
      if ( VT_EMPTY == pValue -> vt ) 
         pValue -> vt = vThis.vt;
      VariantChangeType(pValue,&vThis,0,pValue -> vt);
   }

   return S_OK;
   }


   HRESULT Property::put_arrayValue(SAFEARRAY *psa) {

   if ( ! psa ) return E_POINTER;

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_ARRAY;

   everAssigned = true;

   if ( pVariantArray ) {
      SafeArrayDestroy(pVariantArray);
      pVariantArray = NULL;
   }

   HRESULT rc = S_OK;

   if ( FADF_BSTR & psa -> fFeatures ) {

      long countDims = SafeArrayGetDim(psa);
      long totalCount = 0;

      SAFEARRAYBOUND *pBounds = new SAFEARRAYBOUND[countDims];

      for ( long k = 0; k < countDims; k++ ) {
         SafeArrayGetLBound(psa,k + 1,&pBounds[k].lLbound);
         SafeArrayGetUBound(psa,k + 1,(LONG *)&pBounds[k].cElements);
         pBounds[k].cElements = pBounds[k].cElements - pBounds[k].lLbound + 1;
         totalCount += pBounds[k].cElements;
      }

      pVariantArray = SafeArrayCreate(VT_BSTR,countDims,pBounds);

      BSTR *pValues = NULL;

      BSTR *pSourceValues = NULL;

      SafeArrayAccessData(psa,(void **)&pSourceValues);

      SafeArrayAccessData(pVariantArray,(void **)&pValues);

      for ( long k = 0; k < totalCount; k++ ) {

         *pValues = SysAllocString(*pSourceValues);

         pValues++;
         pSourceValues++;

      }

      delete [] pBounds;

      SafeArrayUnaccessData(psa);

      SafeArrayUnaccessData(pVariantArray);

   } else
      rc = SafeArrayCopy(psa,&pVariantArray);

   return S_OK;
   }


   HRESULT Property::get_arrayValue(SAFEARRAY** ppsa) {

   if ( ! ppsa ) 
      return E_POINTER;

   //if ( ! *ppsa ) 
   //   return E_POINTER;

   if ( pVariantArray )
      SafeArrayCopy(pVariantArray,ppsa);

   //
   //NTC: 01-13-2018: I am taking this diagnostic out - I believe that this situation is okay, i.e., the array was not given a value
   // when the properties were saved.
   //
   // But I am not sure ...
   //

   //else {
   //   if ( debuggingEnabled ) {
   //      char szError[MAX_PATH];
   //      sprintf(szError,"An attempt was made to retrieve the array value of a property but the array value has never been assigned to the property");
   //      MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
   //   }
   //   return E_FAIL;
   //}

   return S_OK;
   }