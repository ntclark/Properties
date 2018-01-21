// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"

#include "utils.h"


   HRESULT Property::put_arrayElement(long index,VARIANT v) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_ARRAY;

   if ( type != TYPE_ARRAY ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::put_arrayElement was called, however, the property is not a TYPE_ARRAY property.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   if ( ! pVariantArray ) {
      SAFEARRAYBOUND rgsaBound;
      rgsaBound.cElements = 0;
      rgsaBound.lLbound = 0;
      pVariantArray = SafeArrayCreateEx(VT_VARIANT,1,&rgsaBound,NULL);
      if ( index != 0 && index != ARRAY_INDEX_FIRST && index != ARRAY_INDEX_LAST && index != ARRAY_INDEX_ADD ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::put_arrayElement was called with an index value out of range.\n\nSpecify within the range of elements in the array, or, ARRAY_INDEX_FIRST or ARRAY_INDEX_LAST or ARRAY_INDEX_ADD \n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
         }
      }
   } 

   everAssigned = true;

   variantStorage* pVariants;
   SAFEARRAYBOUND rgsaBound;

   long cntDims = SafeArrayGetDim(pVariantArray);

   rgsaBound.lLbound = 0;
   rgsaBound.cElements = getSafeArrayVariants(pVariantArray,&pVariants);

   if ( cntDims > 1 ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::put_arrayElement was called for an array property. However, the array is a multi-dimensional array. This array will be converted in place to a 1 dimensional array.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
      SAFEARRAY* pTemp;
      pTemp = SafeArrayCreate(VT_VARIANT,1,&rgsaBound);
      for ( long k = 0; k < (long)rgsaBound.cElements; k++ )
         replaceSafeArrayElement(pVariants[k].vt,pVariants[k].cntBytes,(char *)pVariants[k].pData,pTemp,&k);
      delete [] pVariants;
   }

   long k = index;

   if ( ARRAY_INDEX_FIRST == k ) k = rgsaBound.lLbound;

   if ( ARRAY_INDEX_ADD == k ) {
      rgsaBound.cElements++;
      SafeArrayRedim(pVariantArray,&rgsaBound);
      k = rgsaBound.lLbound + rgsaBound.cElements - 1;
   }

   if ( ARRAY_INDEX_LAST == k ) k = rgsaBound.lLbound + rgsaBound.cElements - 1;

   if ( k < rgsaBound.lLbound || k > rgsaBound.lLbound + (long)rgsaBound.cElements - 1 ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::put_arrayElement was called with an invalid index.\n\nSpecify ARRAY_INDEX_ADD, ARRAY_INDEX_FIRST, ARRAY_INDEX_LAST, or an integer within the range of the array.\n\nThe element is Added to the array.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
      rgsaBound.cElements++;
      SafeArrayRedim(pVariantArray,&rgsaBound);
      k = rgsaBound.lLbound + rgsaBound.cElements - 1;
   }

   if ( ! SUCCEEDED(SafeArrayPutElement(pVariantArray,&k,&v)) ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::put_arrayElement was called with a VARIANT to add to an array property.\n\nHowever, the operation failed.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
      delete [] pVariants;
      return E_FAIL;
   }

   delete [] pVariants;

   return S_OK;
   }


   HRESULT Property::get_arrayElement(long index,VARIANT* pv) {
   if ( ! pv ) return E_POINTER;
   if ( ! pVariantArray ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::get_arrayElement was called before the array property has been provided any values.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   return S_OK;
   }


   HRESULT Property::get_arrayElementCount(long *pCount) {
   if ( ! pCount ) return E_POINTER;
   *pCount = countSafeArrayItems(pVariantArray);
   return S_OK;
   }


   HRESULT Property::clearArray() {
   if ( pVariantArray )
      SafeArrayDestroy(pVariantArray);
   SAFEARRAYBOUND rgsaBound = {0,0};
   pVariantArray = SafeArrayCreate(VT_VARIANT,1,&rgsaBound);
   return S_OK;
   }

