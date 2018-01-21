// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"

#include "utils.h"

#include "list.cpp"


   long Property::getVariantFromWindow(HWND hwnd,variantStorage* ps,char *szMethodName) {

   if ( ! IsWindow(hwnd) ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"%s was called with a window handle that is either not a window or is a window that cannot be accessed.\n\nThe property name is \"%s\"",szMethodName,*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return 0;
   }

   char szClassName[128];
   memset(szClassName,0,sizeof(szClassName));

   GetClassName(hwnd,szClassName,128);

   if ( ! szClassName[0] ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"%s was called with a window handle that is either not a window or is a window that cannot be accessed.\n\nThe property name is \"%s\"",szMethodName,*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return 0;
   }

   char *c = szClassName;
   while ( *c ) {
      *c = tolower(*c);
      c++;
   }

   if ( 0 == strcmp(szClassName,"edit") || 0 == strcmp(szClassName,"thundertextbox") ) {
      getVariantFromEditControl(hwnd,ps);
   } else { 
      if ( 0 == strcmp(szClassName,"button") ) {
         getVariantFromButtonControl(hwnd,ps);
      } else {
         if ( 0 == strcmp(szClassName,"combobox") ) {
            getVariantFromComboBoxControl(hwnd,ps);
         } else {
            if ( 0 == strcmp(szClassName,"listbox") ) {
               getVariantFromListBoxControl(hwnd,ps);
            } else {
               if ( 0 == strcmp(szClassName,"scrollbar") ) {
                  getVariantFromScrollBarControl(hwnd,ps);
               } else {
                  if ( 0 == strcmp(szClassName,"static") ) {
                     getVariantFromStaticControl(hwnd,ps);
                  } else {
                     getStringFromWindow(hwnd,ps);
                  }
               }
            }
         }
      }
   }

   return 0;
   }


   long Property::setWindowFromVariant(HWND hwnd,variantStorage* ps,char *szMethodName) {

   if ( ! IsWindow(hwnd) ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"%s was called with a window handle that is either not a window or is a window that cannot be accessed.\n\nThe property name is \"%s\"",szMethodName,*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return 0;
   }

   char szClassName[128];
   memset(szClassName,0,sizeof(szClassName));

   GetClassName(hwnd,szClassName,128);

   if ( ! szClassName[0] ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"%s was called with a window handle that is either not a window or is a window that cannot be accessed.\n\nThe property name is \"%s\"",szMethodName,*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return 0;
   }

   char *c = szClassName;
   while ( *c ) {
      *c = tolower(*c);
      c++;
   }

   if ( 0 == strcmp(szClassName,"edit") ) {
      setEditControlFromVariant(hwnd,ps);
   } else { 
      if ( 0 == strcmp(szClassName,"button") ) {
         setButtonControlFromVariant(hwnd,ps);
      } else {
         if ( 0 == strcmp(szClassName,"combobox") ) {
            setComboBoxControlFromVariant(hwnd,ps);
         } else {
            if ( 0 == strcmp(szClassName,"listbox") ) {
               setListBoxControlFromVariant(hwnd,ps);
            } else {
               if ( 0 == strcmp(szClassName,"scrollbar") ) {
                  setScrollBarControlFromVariant(hwnd,ps);
               } else {
                  if ( 0 == strcmp(szClassName,"static") ) {
                     setStaticControlFromVariant(hwnd,ps);
                  } else {
                     SetWindowText(hwnd,ps -> pszData);
                  }
               }
            }
         }
      }
   }

   return 0;
   }


   long Property::getSafeArrayVariants(SAFEARRAY* psa,variantStorage** ppVariants) {

   *ppVariants = NULL;

   if ( ! psa ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to access the values of an array, but the array has never been assigned any values.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note",MB_OK);
      }
      return 0;
   }

   *ppVariants = NULL;

   long cntTotalItems = countSafeArrayItems(psa);

   if ( ! cntTotalItems )
      return 0;

   *ppVariants = new variantStorage[cntTotalItems];

   VARTYPE arrayVt;
   SafeArrayGetVartype(psa,&arrayVt);

   long cntDims = SafeArrayGetDim(psa);
   long* vIndexes = new long[cntDims];
   List<long> vLocations;

   getVariantLocations(psa,vLocations,vIndexes,1,cntDims);

   long *pv = NULL;
   variantStorage* p = *ppVariants;

   if ( VT_BSTR == arrayVt ) {

      BSTR bstr;

      while ( pv = vLocations.GetNext(pv) ) {

         SafeArrayGetElement(psa,pv,&bstr);

         p -> vt = VT_BSTR;
         p -> vArrayLocator = pv;
         p -> cntBytes = (long)wcslen(bstr) + 1;
         p -> pData = (BYTE *)new char[p -> cntBytes];
         p -> pszData = new char[p -> cntBytes];

         memset(p -> pData,0,p -> cntBytes);
         memset(p -> pszData,0,p -> cntBytes);

         WideCharToMultiByte(CP_ACP,0,bstr,-1,p -> pszData,p -> cntBytes,0,0);
         memcpy(p -> pData,p -> pszData,p -> cntBytes);

         p -> theVariant.vt = VT_BSTR;
         p -> theVariant.bstrVal = SysAllocString(bstr);

         SysFreeString(bstr);

         *p++;

      }

   } else {

      VARIANT v;

      while ( pv = vLocations.GetNext(pv) ) {

         HRESULT hr = SafeArrayGetElement(psa,pv,&v);
         p -> vt = v.vt;
         p -> vArrayLocator = pv;
         if ( VT_BSTR == v.vt )  {
            p -> cntBytes = (long)wcslen(v.bstrVal) + 1;
            p -> cntszBytes = p -> cntBytes;
            p -> pData = new BYTE[p -> cntBytes];
            p -> pszData = new char[p -> cntBytes];
            memset(p -> pData,0,p -> cntBytes);
            memset(p -> pszData,0,p -> cntszBytes);
            WideCharToMultiByte(CP_ACP,0,v.bstrVal,-1,p -> pszData,p -> cntBytes,0,0);
            memcpy(p -> pData,p -> pszData,p -> cntBytes);
         } else {
            if ( VT_ARRAY & (VARTYPE)v.vt ) {

               long cntElements = getSafeArrayVariants(v.parray,&p -> pSubElements);

               long cntSubDim = SafeArrayGetDim(v.parray);

               p -> cntszBytes = 0;
               p -> pszData = NULL;

               p -> cntBytes = cntSubDim * sizeof(SAFEARRAYBOUND) + 2 * sizeof(long);
               for ( long k = 0; k < cntElements; k++ ) {
                  p -> cntBytes += sizeof(variantStorage);
                  p -> cntBytes += cntSubDim * sizeof(long);
                  p -> cntBytes += p -> pSubElements[k].cntBytes;
               }

               BYTE* b;
               p -> pData = new BYTE[p -> cntBytes];
               b = p -> pData;

               memset(b,0,p -> cntBytes);
               *(long*)(b) = cntElements;
               *(long*)(b + sizeof(long)) = cntSubDim;

               b += 2 * sizeof(long);

               SAFEARRAYBOUND* pBound = (SAFEARRAYBOUND*)(b);

               for ( int k = 0; k < cntSubDim; k++ ) {
                  SafeArrayGetLBound(v.parray,k + 1,&pBound -> lLbound);
                  SafeArrayGetUBound(v.parray,k + 1,(long*)&pBound -> cElements);
                  pBound -> cElements = pBound -> cElements - pBound -> lLbound + 1;
                  pBound++;
               }

               b += cntSubDim * sizeof(SAFEARRAYBOUND);

               for ( int k = 0; k < cntElements; k++ ) {
                  memcpy(b,p -> pSubElements + k,sizeof(variantStorage));
                  b += sizeof(variantStorage);
                  memcpy(b,p -> pSubElements[k].vArrayLocator,cntSubDim * sizeof(long));
                  b += cntSubDim * sizeof(long);
                  memcpy(b,p -> pSubElements[k].pData,p -> pSubElements[k].cntBytes);
                  b += p -> pSubElements[k].cntBytes;
               }

            } else {
               getVariantBinaryRepresentation(arrayVt,&p -> vt,&v,&p -> pData,&p -> cntBytes);
               convertVariantToSzString(&v,&p -> pszData);
            }
         }

         VariantCopy(&p -> theVariant,&v);

         *p++;

      }

   }

   while ( pv = vLocations.GetFirst() ) 
      vLocations.Remove(pv);

   delete [] vIndexes;

   return cntTotalItems;

   }


   int Property::getVariantLocations(SAFEARRAY* pVariantArray,List<long>& vLocations,long *vIndexes,long dim,long cntDims) {

   if ( dim != cntDims ) {

      long lBound,uBound;
      SafeArrayGetLBound(pVariantArray,dim,&lBound);
      SafeArrayGetUBound(pVariantArray,dim,&uBound);

      for ( long j = lBound; j <= uBound; j++ ) {

         vIndexes[dim - 1] = j;

         getVariantLocations(pVariantArray,vLocations,vIndexes,dim + 1,cntDims);

      }         

      return 0;

   }
         
   long lBound,uBound;
   SafeArrayGetLBound(pVariantArray,dim,&lBound);
   SafeArrayGetUBound(pVariantArray,dim,&uBound);

   for ( long j = lBound; j <= uBound; j++ ) {

      vIndexes[cntDims - 1] = j;

      long* v = new long[cntDims];

      for ( long k = 0; k < cntDims; k++ ) 
         v[k] = vIndexes[k];

      vLocations.Add(v);

   }

   return 0;
   }


   int Property::replaceSafeArrayElement(VARTYPE vt,unsigned int cntBytes,char *pszData,SAFEARRAY* psa,long *vIndexes) {

   VARTYPE arrayVt;

   SafeArrayGetVartype(psa,&arrayVt);

   VARIANT v = {VT_EMPTY};
   GVariantClear(&v);

   v.vt = vt;
   if ( VT_BSTR == vt ) {
      v.bstrVal = SysAllocStringLen(NULL,(unsigned int)cntBytes);
      MultiByteToWideChar(CP_ACP,0,pszData,-1,v.bstrVal,cntBytes);
      if ( VT_VARIANT == arrayVt ) 
         SafeArrayPutElement(psa,vIndexes,&v);
      else
         SafeArrayPutElement(psa,vIndexes,v.bstrVal);
      SysFreeString(v.bstrVal);
   } else {
      convertSzStringToVariant(pszData,&v);
      if ( VT_VARIANT != arrayVt ) {
         void* pvData = new BYTE[sizeof(VARIANT)];
         convertVariantToScalarValue(&v,pvData);
         SafeArrayPutElement(psa,vIndexes,pvData);
         delete [] pvData;
      } else
         SafeArrayPutElement(psa,vIndexes,&v);
   }

   return 0;
   }


   int Property::replaceSafeArrayVariant(SAFEARRAY* psa,long *vIndexes,VARIANT* pv) {
   VARTYPE arrayVt;
   SafeArrayGetVartype(psa,&arrayVt);
   if ( VT_BSTR == arrayVt ) {
      if ( VT_BSTR != pv -> vt ) 
         VariantChangeType(pv,pv,0,VT_BSTR);
      SafeArrayPutElement(psa,vIndexes,pv -> bstrVal);
   } else {
      SafeArrayPutElement(psa,vIndexes,pv);
   }
   return 0;
   }