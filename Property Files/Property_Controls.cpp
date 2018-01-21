// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"
#include "utils.h"

   long Property::setButtonControlFromVariant(HWND hwnd,variantStorage* ps) {

   VARIANT vtBool;
   VariantInit(&vtBool);
   VariantChangeType(&vtBool,&ps -> theVariant,0,VT_BOOL);

   ULONG_PTR style = GetWindowLongPtr(hwnd,GWL_STYLE);

   style &= 0x0000000FL;

   if ( style == BS_PUSHBUTTON && ps -> pszData ) {
      SetWindowText(hwnd,ps -> pszData);
   }
   if ( style == BS_DEFPUSHBUTTON && ps -> pszData ) {
      SetWindowText(hwnd,ps -> pszData);
   }
   if ( style == BS_CHECKBOX ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }
   if ( style == BS_AUTOCHECKBOX ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }
   if ( style == BS_3STATE ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }
   if ( style == BS_AUTO3STATE ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }
   if ( style == BS_RADIOBUTTON ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }
   if ( style == BS_AUTORADIOBUTTON ) {
      SendMessage(hwnd,BM_SETCHECK,vtBool.boolVal ? 1L : 0L,0L);
   }

   return 0;
   }

   
   long Property::setEditControlFromVariant(HWND hwnd,variantStorage* ps) {
   SetWindowText(hwnd,ps -> pszData);
   return 0;
   }


   long Property::setComboBoxControlFromVariant(HWND hwnd,variantStorage* ps) {

   if ( ps -> theVariant.vt & (VARTYPE)VT_ARRAY )
      return setComboBoxControlFromArray(hwnd,ps -> theVariant.parray);

   LRESULT cntItems = SendMessage(hwnd,CB_GETCOUNT,0L,0L);

   for ( long k = 0; k < cntItems; k++ ) {
      LONG_PTR n = SendMessage(hwnd,CB_GETLBTEXTLEN,(WPARAM)k,0L);
      char *pszData = new char[n + 1];
      memset(pszData,0,n + 1);
      SendMessage(hwnd,CB_GETLBTEXT,(WPARAM)k,(LPARAM)pszData);
      if ( 0 == strcmp(pszData,ps -> pszData) ) {
         SendMessage(hwnd,CB_SETCURSEL,(WPARAM)k,0L);
         delete [] pszData;
         return 0;
      }
      delete [] pszData;
   }
   return 0;
   }


   long Property::setListBoxControlFromVariant(HWND hwnd,variantStorage* ps) {

   if ( ps -> theVariant.vt & (VARTYPE)VT_ARRAY )
      return setListBoxControlFromArray(hwnd,ps -> theVariant.parray);

   LONG_PTR cntItems = SendMessage(hwnd,LB_GETCOUNT,0L,0L);

   for ( long k = 0; k < cntItems; k++ ) {
      LONG_PTR n = SendMessage(hwnd,LB_GETTEXTLEN,(WPARAM)k,0L);
      char *pszData = new char[n + 1];
      memset(pszData,0,n + 1);
      SendMessage(hwnd,LB_GETTEXT,(WPARAM)k,(LPARAM)pszData);
      if ( 0 == strcmp(pszData,ps -> pszData) ) {
         SendMessage(hwnd,LB_SETCURSEL,(WPARAM)k,0L);
         delete [] pszData;
         return 0;
      }
      delete [] pszData;
   }
   return 0;
   }


   long Property::setScrollBarControlFromVariant(HWND,variantStorage* ps) {
   return 0;
   }


   long Property::setStaticControlFromVariant(HWND hwnd,variantStorage* ps) {
   SetWindowText(hwnd,ps -> pszData);
   return 0;
   }


   long Property::getVariantFromButtonControl(HWND hwnd,variantStorage* ps) {

   VARIANT vtBool;
   VariantInit(&vtBool);
   vtBool.vt = VT_BOOL;
   vtBool.boolVal = FALSE;

   LRESULT bmCheck = 0;

   LONG_PTR style = GetWindowLongPtr(hwnd,GWL_STYLE);

   style &= 0x0000000FL;

   if ( style == BS_PUSHBUTTON ) {
      return getStringFromWindow(hwnd,ps);
   }
   if ( style == BS_DEFPUSHBUTTON ) {
      return getStringFromWindow(hwnd,ps);
   }
   if ( style == BS_CHECKBOX ) {
      bmCheck = SendMessage(hwnd,BM_GETCHECK,0L,0L);
   }
   if ( style == BS_AUTOCHECKBOX ) {
      bmCheck = SendMessage(hwnd,BM_GETCHECK,0L,0L);
   }
   if ( style == BS_3STATE || style == BS_AUTO3STATE ) {
      bmCheck = SendMessage(hwnd,BM_GETCHECK,0L,0L);
      if ( BST_INDETERMINATE == bmCheck ) {
      }
   }
   if ( style == BS_RADIOBUTTON ) {
      bmCheck = SendMessage(hwnd,BM_GETCHECK,0L,0L);
   }
   if ( style == BS_AUTORADIOBUTTON ) {
      bmCheck = SendMessage(hwnd,BM_GETCHECK,0L,0L);
   }

   if ( BST_CHECKED == bmCheck ) {
      vtBool.boolVal = TRUE;
   } else {
      if ( BST_UNCHECKED == bmCheck ) {
         vtBool.boolVal = FALSE;
      } else {
      }
   }

   if ( ps -> vt == VT_BSTR ) 
      VariantChangeType(&ps -> theVariant,&vtBool,VARIANT_ALPHABOOL,VT_BOOL);
   else
      VariantChangeType(&ps -> theVariant,&vtBool,0,VT_BOOL);

   replaceSafeArrayVariant(pVariantArray,ps -> vArrayLocator,&ps -> theVariant);

   return 0;
   }

   
   long Property::getVariantFromEditControl(HWND hwnd,variantStorage* ps) {
   getStringFromWindow(hwnd,ps);
   return 0;
   }


   long Property::getVariantFromComboBoxControl(HWND hwnd,variantStorage* ps) {
   if ( ps -> theVariant.vt & (VARTYPE)VT_ARRAY ) 
      return getArrayFromComboBoxControl(hwnd,ps -> theVariant.parray);
   return 0;
   }


   long Property::getVariantFromListBoxControl(HWND hwnd,variantStorage* ps) {
   if ( ps -> theVariant.vt & (VARTYPE)VT_ARRAY ) 
      return getArrayFromListBoxControl(hwnd,ps -> theVariant.parray);
   return 0;
   }


   long Property::getVariantFromScrollBarControl(HWND,variantStorage* ps) {
   return 0;
   }


   long Property::getVariantFromStaticControl(HWND hwnd,variantStorage* ps) {
   getStringFromWindow(hwnd,ps);
   return 0;
   }

   long Property::getStringFromWindow(HWND hwnd,variantStorage* ps) {
   LONG_PTR n = SendMessage(hwnd,WM_GETTEXTLENGTH,0L,0L);
   if ( ps -> pszData ) delete [] ps -> pszData;
   ps -> pszData = NULL;
   ps -> cntszBytes = 0;
   ps -> pszData = new char[n + 1];
   ps -> cntszBytes = (long)n + 1;
   SendMessage(hwnd,WM_GETTEXT,(WPARAM)(n + 1),(LPARAM)ps -> pszData);
   replaceSafeArrayElement(ps -> vt,(unsigned int)n + 1,ps -> pszData,pVariantArray,ps -> vArrayLocator);
   return 0;
   }


   long Property::setComboBoxControlFromArray(HWND hwnd,SAFEARRAY* psa) {

   SendMessage(hwnd,CB_RESETCONTENT,0L,0L);

   variantStorage* pVariants;
   long cntItems = getSafeArrayVariants(psa,&pVariants);

   for ( long k = 0; k < cntItems; k++ )
      SendMessage(hwnd,CB_ADDSTRING,0L,(LPARAM)pVariants[k].pszData);

   SendMessage(hwnd,CB_SETCURSEL,0L,0L);

   if ( cntItems )
      delete [] pVariants;

   return S_OK;
   }


   long Property::setListBoxControlFromArray(HWND hwnd,SAFEARRAY* psa) {

   SendMessage(hwnd,LB_RESETCONTENT,0L,0L);

   variantStorage* pVariants;
   long cntItems = getSafeArrayVariants(psa,&pVariants);

   for ( long k = 0; k < cntItems; k++ )
      SendMessage(hwnd,LB_ADDSTRING,0L,(LPARAM)pVariants[k].pszData);

   SendMessage(hwnd,CB_SETCURSEL,0L,0L);

   if ( cntItems )
      delete [] pVariants;

   return S_OK;
   }


   long Property::getArrayFromComboBoxControl(HWND hwnd,SAFEARRAY* psa) {

   variantStorage* pVariants;
   long cntArrayItems = getSafeArrayVariants(psa,&pVariants);
   long cntListItems = (long)SendMessage(hwnd,CB_GETCOUNT,0L,0L);

   for ( long k = 0; k < cntArrayItems ; k++ ) {
      
      if ( pVariants[k].pszData ) {
         delete [] pVariants[k].pszData;
         pVariants[k].pszData = NULL;
         pVariants[k].cntszBytes = 0;
      }
      long n = (long)SendMessage(hwnd,CB_GETLBTEXTLEN,k,0L);
      if ( n > 0 ) {
         pVariants[k].cntszBytes = n + 1;
         pVariants[k].pszData = new char[n + 1];
         memset(pVariants[k].pszData,0,n + 1);
         SendMessage(hwnd,CB_GETLBTEXT,k,(LPARAM)pVariants[k].pszData);
         replaceSafeArrayElement(pVariants[k].vt,pVariants[k].cntszBytes,pVariants[k].pszData,psa,pVariants[k].vArrayLocator);
      }

   }

   if ( cntArrayItems ) 
      delete [] pVariants;

   return 0;
   }


   long Property::getArrayFromListBoxControl(HWND hwnd,SAFEARRAY* psa) {

   variantStorage* pVariants;
   long cntArrayItems = getSafeArrayVariants(psa,&pVariants);
   long cntListItems = (long)SendMessage(hwnd,LB_GETCOUNT,0L,0L);

   for ( long k = 0; k < cntArrayItems ; k++ ) {
      
      if ( pVariants[k].pszData ) {
         delete [] pVariants[k].pszData;
         pVariants[k].pszData = NULL;
         pVariants[k].cntszBytes = 0;
      }
      long n = (long)SendMessage(hwnd,LB_GETTEXTLEN,k,0L);
      if ( n > 0 ) {
         pVariants[k].cntszBytes = n + 1;
         pVariants[k].pszData = new char[n + 1];
         memset(pVariants[k].pszData,0,n + 1);
         SendMessage(hwnd,LB_GETTEXT,k,(LPARAM)pVariants[k].pszData);
         replaceSafeArrayElement(pVariants[k].vt,pVariants[k].cntszBytes,pVariants[k].pszData,psa,pVariants[k].vArrayLocator);
      }

   }

   if ( cntArrayItems ) 
      delete [] pVariants;

   return 0;
   }