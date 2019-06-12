// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"
#include "utils.h"
#include "list.cpp"

   static char errorBuffer[] = {"An error occurred in the GSystem Property component"};

   Property::Property() :

   type(TYPE_UNSPECIFIED),
   ignoreSetAction(false),
   isDirty(false),
   everAssigned(false),
   debuggingEnabled(false),
   hwndClientWindow(NULL),
   oldClientWindowHandler(NULL),

   pIPropertyClient(NULL),

   bstrValue(NULL),
   pVariantArray(NULL),
   variantIndex(-1),

   directAccessAddress(NULL),
   directAccessSize(0)
 
   {
 
   memset(name,0,sizeof(name));
   strcpy(name,"<unnamed>");
 
   memset(&v,0,sizeof(v));
 
   return;
   }
 
 
   Property::~Property() {
   if ( pIPropertyClient )
      pIPropertyClient -> Release();
   if ( v.binarySize ) 
      delete [] v.binaryValue;
   if ( pVariantArray )
      SafeArrayDestroy(pVariantArray);
   return;
   }


   int Property::updateFromDirectAccess() {
   if ( directAccessAddress && directAccessSize ) {
      memset(&v.scalar,0,sizeof(v.scalar));
      memcpy(&v.scalar,directAccessAddress,min(sizeof(v.scalar),directAccessSize));
      if ( TYPE_STRING == type ) {
         BSTR* pBstr = (BSTR*)directAccessAddress;
         if ( bstrValue ) SysFreeString(bstrValue);
         if ( *pBstr ) 
            bstrValue = SysAllocString(*pBstr);
         else
            bstrValue = SysAllocString(L"");
         delete [] v.binaryValue;
         v.binarySize = (long)wcslen(bstrValue);
         v.binaryValue = (unsigned char *)(new char[v.binarySize + 1]);
         WideCharToMultiByte(CP_ACP,0,bstrValue,-1,reinterpret_cast<char*>(v.binaryValue),v.binarySize + 1,0,0);
      }
      else {
         if ( v.binaryValue ) {
            delete [] v.binaryValue;
            v.binarySize = directAccessSize;
            v.binaryValue = (unsigned char *)(new char[v.binarySize]);
            memset(v.binaryValue,0,v.binarySize);
            memcpy(v.binaryValue,directAccessAddress,min(v.binarySize,directAccessSize));
         }
      }
   }
   return 0;
   }
   

   BYTE* Property::propertyBinaryValue() {

   updateFromDirectAccess();

   if ( TYPE_ARRAY == type ) {

      if ( ! pVariantArray ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"An attempt was made to access the current value of an array property, but the array has never been assigned a value.");
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
         }
         return (BYTE*)errorBuffer;
      }

      variantStorage *pVariants;
      long cntItems = getSafeArrayVariants(pVariantArray,&pVariants);

      if ( variantIndex < 0 || variantIndex > cntItems - 1) {
         if ( debuggingEnabled ) {
            char szError[256];
            sprintf(szError,"An attempt was made to set the value of a member of a VARIANT array, however, the index of the array was either not set or is out of range.\n\nSet the variantIndex property of the property.");
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
         }
         delete [] pVariants;
         return (BYTE *)errorBuffer;
      }

      if ( v.binaryValue ) 
         delete [] v.binaryValue;

      v.binarySize = pVariants[variantIndex].cntBytes;
      v.binaryValue = new BYTE[v.binarySize];
      memcpy(v.binaryValue,pVariants[variantIndex].pData,pVariants[variantIndex].cntBytes);

      delete [] pVariants;

   }

   return v.binaryValue;

   }


   BYTE* Property::propertyScalarValue() {
   updateFromDirectAccess();
   return reinterpret_cast<BYTE*>(&v.scalar);
   }


   int Property::setArrayValue() {

   if ( TYPE_ARRAY != type ) return 0;

   variantStorage* pVariants;
   long cntItems = getSafeArrayVariants(pVariantArray,&pVariants);

   for ( int k = 0; k < cntItems; k++ ) {
      if ( ! strcmp((char*)v.binaryValue,pVariants[k].pszData) ) {
         variantIndex = k;
         delete [] pVariants;
         return S_OK;
      }
   }

   long cntDim = SafeArrayGetDim(pVariantArray);

   if ( variantIndex < 0 || variantIndex > cntItems - 1 ) {

      if ( 1 != cntDim ) {
         delete [] pVariants;
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"An attempt was made to set the value of a member of an array property, however, the index of the array was either not set or is out of range.\nIf the array were a 1 dimensional array, the Properties Component would have re-dimensioned it.\nSet the variantIndex property of the property.\n\nThe property name is \"%s\"",*name ? name : "");
            MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
            return S_OK;
         }
         return E_FAIL;
      }

      SAFEARRAYBOUND rgsaBounds;
      SafeArrayGetLBound(pVariantArray,1,&rgsaBounds.lLbound);
      rgsaBounds.cElements = cntItems + 1;
      SafeArrayRedim(pVariantArray,&rgsaBounds);
      cntItems = cntItems + 1;
      variantIndex = cntItems - 1;

      long vArrayLocator = rgsaBounds.lLbound + cntItems - 1;

      replaceSafeArrayElement(pVariants[variantIndex - 1].vt,(unsigned int)strlen((char*)v.binaryValue),(char*)v.binaryValue,pVariantArray,&vArrayLocator);

   } else {

      replaceSafeArrayElement(pVariants[variantIndex].vt,(unsigned int)strlen((char*)v.binaryValue),(char*)v.binaryValue,pVariantArray,pVariants[variantIndex].vArrayLocator);

   }
   
   delete [] pVariants;

   return 0;

   }


   HRESULT Property::toEncodedText(BSTR* pResult) {
   updateFromDirectAccess();
   switch ( type ) {
   case TYPE_BINARY: {

      int n = 2 * v.binarySize;
      char *result = new char[n + 1];
      memset(result,0,n + 1);

      for ( int k = 0; k < n; k += 2 ) 
         sprintf(result + k,"%2x",v.binaryValue[k/2]);

      char *c = result;
      while ( *c ) {
         if ( *c == ' ' ) *c = '0';
         c++;
      }

      *pResult = SysAllocStringLen(NULL,n + 2);
      MultiByteToWideChar(CP_ACP,0,result,-1,*pResult,n + 2);
      delete [] result;
      }

      break;
   
   case TYPE_ARRAY: {

      VARTYPE vt;
      SafeArrayGetVartype(pVariantArray,&vt);

      variantStorage* pVariantStorage;

      long cntTotalItems = getSafeArrayVariants(pVariantArray,&pVariantStorage);
      if ( ! cntTotalItems ) {
         *pResult = SysAllocStringLen(L"",0);
         return S_OK;
      }

      long cntDims = SafeArrayGetDim(pVariantArray);
      long cntBytes = 40 + cntDims * 2 * sizeof(SAFEARRAYBOUND);

      variantStorage* p = pVariantStorage;
      for ( int k = 0; k < cntTotalItems; k++ ) {
         cntBytes += 2 * p -> cntBytes + 16 + cntDims * 8;
         *p++;
      }

      char* theResults = new char[cntBytes + 1];
      memset(theResults,0,cntBytes + 1);

      char *psz = theResults;

      psz += sprintf(psz,"%08x",cntBytes + 1);
      psz += sprintf(psz,"%08x",cntTotalItems);
      psz += sprintf(psz,"%08x",vt);
      psz += sprintf(psz,"%08x",cntDims);
      psz += sprintf(psz,"%08x",variantIndex);
      
      for ( int k = 0; k < cntDims; k++ ) {
         SAFEARRAYBOUND rgsaBound;
         long uBound;
         BYTE* b = (BYTE*)&rgsaBound;
         SafeArrayGetLBound(pVariantArray,k + 1,&rgsaBound.lLbound);
         SafeArrayGetUBound(pVariantArray,k + 1,&uBound);
         rgsaBound.cElements = uBound - rgsaBound.lLbound + 1;
         psz += sprintf(psz,"%08x",rgsaBound.cElements);
         psz += sprintf(psz,"%08x",rgsaBound.lLbound);
      }

      p = pVariantStorage;
      for ( int k = 0 ; k < cntTotalItems; k++ ) {
         psz += sprintf(psz,"%08x",p -> cntBytes);
         psz += sprintf(psz,"%08x",p -> vt);
         for ( long dim = 0; dim < cntDims; dim++ ) {
            psz += sprintf(psz,"%08x",p -> vArrayLocator[dim]);
         }
         for ( long j = 0; j < 2 * p -> cntBytes; j += 2 )
            sprintf(psz + j,"%02x",p -> pData[j/2]);

         psz += 2 * p -> cntBytes;

         p++;

      }

      delete [] pVariantStorage;

      *pResult = SysAllocStringLen(NULL,cntBytes + 1);
      MultiByteToWideChar(CP_ACP,0,theResults,-1,*pResult,cntBytes + 1);

      }
      return S_OK;

   default: {
      }
   
   }

   return S_OK;
   }



   HRESULT Property::fromEncodedText(BSTR theText) {
   switch ( type ) {
   case TYPE_BINARY: {

      int n = (int)wcslen(theText);
      char *result = new char[n + 1];
      memset(result,0,n + 1);
      WideCharToMultiByte(CP_ACP,0,theText,-1,result,n + 1,0,0);

      if ( v.binarySize ) 
         delete [] v.binaryValue;

      v.binarySize = n / 2;
      v.binaryValue = new BYTE[v.binarySize + 4];
      memset(v.binaryValue,0,v.binarySize);

      for ( int k = 0; k < n; k += 2 ) 
         sscanf(result + k,"%2x",(unsigned int *)(v.binaryValue + k/2));
   
      delete [] result;

      setBinaryValueDirectAccess();

      }
      return S_OK;

   case TYPE_ARRAY: {

      long n = (long)wcslen(theText);
      char* theResults = new char[n + 1];
      char *psz = theResults;

      long cntBytes;
      long cntTotalItems;
      VARTYPE arrayVt;
      long cntDims;

      memset(theResults,0,n + 1);
      WideCharToMultiByte(CP_ACP,0,theText,-1,theResults,n + 1,0,0);

      sscanf(psz,"%8x",&cntBytes);
      psz += 8;      
      sscanf(psz,"%8x",&cntTotalItems);
      psz += 8;      
      sscanf(psz,"%8x",(unsigned int *)&arrayVt);
      psz += 8;      
      sscanf(psz,"%8x",&cntDims);
      psz += 8;
      sscanf(psz,"%8x",&variantIndex);
      psz += 8;

      SAFEARRAYBOUND* prgsaBounds = new SAFEARRAYBOUND[cntDims];

      for ( long k = 0; k < cntDims; k++ ) {
         sscanf(psz,"%8x",&prgsaBounds[k].cElements);
         psz += 8;
         sscanf(psz,"%8x",&prgsaBounds[k].lLbound);
         psz += 8;
      }

      if ( pVariantArray )
         SafeArrayDestroy(pVariantArray);

      pVariantArray = SafeArrayCreate(arrayVt,cntDims,prgsaBounds);

      delete [] prgsaBounds;

      variantStorage vStorage;

      for ( int k = 0 ; k < cntTotalItems; k++ ) {

         sscanf(psz,"%8x",&vStorage.cntBytes);
         psz += 8;

         sscanf(psz,"%8x",(unsigned int *)&vStorage.vt);
         psz += 8;

         vStorage.pData = new BYTE[vStorage.cntBytes + 4];
         vStorage.vArrayLocator = new long[cntDims];

         for ( long dim = 0; dim < cntDims; dim++ ) {
            sscanf(psz,"%8x",&vStorage.vArrayLocator[dim]);
            psz += 8;
         }

         for ( long j = 0; j < 2 * vStorage.cntBytes; j += 2 ) {
            sscanf(psz + j,"%2hx",(unsigned short *)&vStorage.pData[j/2]);
         }

         VARIANT v = {VT_EMPTY};
         GVariantClear(&v);
         if ( VT_BSTR == vStorage.vt ) {
            v.vt = VT_BSTR;
            v.bstrVal = SysAllocStringLen(NULL,(unsigned int)vStorage.cntBytes);
            MultiByteToWideChar(CP_ACP,0,(char*)vStorage.pData,-1,v.bstrVal,vStorage.cntBytes);
            if ( VT_VARIANT == arrayVt ) 
               HRESULT hr = SafeArrayPutElement(pVariantArray,vStorage.vArrayLocator,&v);
            else
               HRESULT hr = SafeArrayPutElement(pVariantArray,vStorage.vArrayLocator,v.bstrVal);
            SysFreeString(v.bstrVal);
         } else {
            if ( vStorage.vt & (VARTYPE)VT_ARRAY ) {

               VARTYPE subArrayType = vStorage.vt & VT_TYPEMASK;

               v.vt = vStorage.vt;

               long cntSubElements = *(long*)(vStorage.pData);
               long cntSubDim = *(long*)(vStorage.pData + sizeof(long));

               SAFEARRAYBOUND* pBound = (SAFEARRAYBOUND*)(vStorage.pData + 2 * sizeof(long));

               v.parray = SafeArrayCreate(subArrayType,cntSubDim,pBound);

               BYTE *b = vStorage.pData + 2 * sizeof(long) + cntSubDim * sizeof(SAFEARRAYBOUND);

               vStorage.pSubElements = new variantStorage[cntSubElements];

               for ( long k = 0; k < cntSubElements; k++ ) {

                  memcpy(vStorage.pSubElements + k,b,sizeof(variantStorage));
                  b += sizeof(variantStorage);

                  vStorage.pSubElements[k].vArrayLocator = NULL;
                  vStorage.pSubElements[k].pData = NULL;
                  vStorage.pSubElements[k].pszData = NULL;
                  vStorage.pSubElements[k].pSubElements = NULL;
                  memset(&vStorage.pSubElements[k].theVariant,0,sizeof(VARIANT));

                  vStorage.pSubElements[k].vArrayLocator = new long[cntSubDim];
                  memcpy(vStorage.pSubElements[k].vArrayLocator,b,cntSubDim * sizeof(long));
                  b += cntSubDim * sizeof(long);

                  vStorage.pSubElements[k].pData = new BYTE[vStorage.pSubElements[k].cntBytes];
                  memcpy(vStorage.pSubElements[k].pData,b,vStorage.pSubElements[k].cntBytes);
                  b += vStorage.pSubElements[k].cntBytes;

                  if ( vStorage.pSubElements[k].vt == VT_BSTR ) {
                     VARIANT vString;
                     VariantInit(&vString);
                     vString.vt = VT_BSTR;
                     vString.bstrVal = SysAllocStringLen(NULL,(unsigned int)vStorage.pSubElements[k].cntBytes);
                     MultiByteToWideChar(CP_ACP,0,(char*)vStorage.pSubElements[k].pData,vStorage.pSubElements[k].cntBytes,
                           vString.bstrVal,vStorage.pSubElements[k].cntBytes);
                     if ( VT_VARIANT == subArrayType ) 
                        HRESULT hr = SafeArrayPutElement(v.parray,vStorage.pSubElements[k].vArrayLocator,&vString);
                     else
                        HRESULT hr = SafeArrayPutElement(v.parray,vStorage.pSubElements[k].vArrayLocator,vString.bstrVal);
                     SysFreeString(vString.bstrVal);
                  } else { 
                     getVariantFromBinaryRepresentation(&v,vStorage.pSubElements[k].vt,vStorage.pSubElements[k].pData,vStorage.pSubElements[k].cntBytes);
                     SafeArrayPutElement(v.parray,vStorage.pSubElements[k].vArrayLocator,&v);
                  }

               }
               SafeArrayPutElement(pVariantArray,vStorage.vArrayLocator,&v);

               delete [] vStorage.pSubElements;

               vStorage.pSubElements = NULL;

            } else {
               getVariantFromBinaryRepresentation(&v,vStorage.vt,vStorage.pData,vStorage.cntBytes);
               SafeArrayPutElement(pVariantArray,vStorage.vArrayLocator,&v);
            }
         }

         delete [] vStorage.pData;
         delete [] vStorage.vArrayLocator;

         vStorage.pData = NULL;
         vStorage.vArrayLocator = NULL;

         psz += 2 * vStorage.cntBytes;

      }

      delete [] theResults;

      }
      return S_OK;

   default: {
   }
   
   }

   return S_OK;
   }