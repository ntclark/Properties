/*           

                       Copyright (c) 1996,1997,1998,1999,2000,2001,2002 Nathan T. Clark

*/

#include <windows.h>
#include <stdio.h>                           

#include "Property.h"
#include "utils.h"

   // Window Contents support

   HRESULT Property::getWindowValue(HWND hwndControl) {
   char szValue[4096];
   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to Property::getWindowValue");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   GetWindowText(hwndControl,szValue,4096);
   VARIANT_BOOL oldIgnoreSetAction = ignoreSetAction;
   ignoreSetAction = true;
   put_szValue(szValue);
   ignoreSetAction = oldIgnoreSetAction;
   return S_OK;
   }


   HRESULT Property::getWindowItemValue(HWND hwndDialog,long idControl) {
   return getWindowValue(GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Property::getWindowText(HWND hwndControl) {
   return getWindowValue(hwndControl);
   }


   HRESULT Property::getWindowItemText(HWND hwndDialog,long idControl) {
   return getWindowItemValue(hwndDialog,idControl);
   }


   HRESULT Property::setWindowText(HWND hwndControl) {
   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"A NULL window handle was passed to IGProperty::setWindowText");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
/* 07/06/02 Added this switch statement because properties that had a scalar value
   were not getting a return value from propertyBinaryValue() that would display as text.
   Old code was just SetWindowText with propertyBinaryValue().
*/
   char szNumber[64] = {'\0'};
   switch ( type ) {
   case TYPE_LONG:
      sprintf(szNumber,"%ld",v.scalar.longValue);
      break;
   case TYPE_DOUBLE:
      sprintf(szNumber,"%lf",v.scalar.doubleValue);
      break;
   case TYPE_SZSTRING:
   case TYPE_STRING:
   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
   case TYPE_ARRAY:
      SetWindowText(hwndControl,reinterpret_cast<char*>(propertyBinaryValue()));
      break;
   case TYPE_BOOL:
      sprintf(szNumber,"%d",v.scalar.boolValue);
      break;
   case TYPE_VARIANT:
   case TYPE_OBJECT_STORAGE_ARRAY:
      break;
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to set the contents of a window using a type of property that is not valid for this operation. The property name is %s.",name);
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   default:
      break;
   }
   if ( szNumber[0] ) SetWindowText(hwndControl,szNumber);
   return S_OK;
   }


   HRESULT Property::setWindowItemText(HWND hwndDialog,long idControl) {
   return setWindowText(GetDlgItem(hwndDialog,idControl));
   }


   // Combo box selection

   HRESULT Property::getWindowComboBoxSelection(HWND hwndControl) {
   char szValue[MAX_PROPERTY_SIZE];

   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to Property::getWindowComboBoxSelection");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   long rv = (long)SendMessage(hwndControl,CB_GETCURSEL,0L,0L);
   if ( CB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"Property::getWindowComboBoxSelection was called but the window provided is not a Combo-box control or there is no current selection\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   rv = (long)SendMessage(hwndControl,CB_GETLBTEXT,(WPARAM)rv,(LPARAM)szValue);
   if ( CB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"Property::getWindowComboBoxSelection was called but the window provided is not a Combo-box control or there is no current selection\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   VARIANT_BOOL oldIgnoreSetAction = ignoreSetAction;
   ignoreSetAction = true;
   put_szValue(szValue);
   ignoreSetAction = oldIgnoreSetAction;

   return S_OK;
   }


   HRESULT Property::getWindowItemComboBoxSelection(HWND hwndDialog,long idControl) {
   return getWindowComboBoxSelection(GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Property::setWindowComboBoxSelection(HWND hwndControl) {
   char szValue[MAX_PROPERTY_SIZE];

   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to IGProperty::setWindowComboBoxSelection");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   HRESULT hr = get_szValue(szValue);
   if ( S_OK != hr ) return hr;
   LRESULT rv = SendMessage(hwndControl,CB_FINDSTRINGEXACT,-1L,(LPARAM)szValue);

   if ( CB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"The string %s was intended as the selection in a ComboBox control, however, that value is not already in the ComboBox control. Method IGProperty::setWindowComboBoxSelection",propertyBinaryValue());
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   SendMessage(hwndControl,CB_SETCURSEL,rv,0);

   return S_OK;
   }


   HRESULT Property::setWindowItemComboBoxSelection(HWND hwndDialog,long idControl) {
   return setWindowComboBoxSelection(GetDlgItem(hwndDialog,idControl));
   }


   // Combo box list

   HRESULT Property::setWindowComboBoxList(HWND vHwndControl) {

   HWND hwndControl = vHwndControl;

   if ( ! pVariantArray ) {
      SendMessage(hwndControl,CB_RESETCONTENT,0L,0L);
      return S_OK;
   }

   setComboBoxControlFromArray(hwndControl,pVariantArray);

   if ( variantIndex > -1 ) 
      SendMessage(hwndControl,CB_SETCURSEL,variantIndex,0L);
   else
      SendMessage(hwndControl,CB_SETCURSEL,0L,0L);

   return S_OK;

   }


   HRESULT Property::setWindowItemComboBoxList(HWND vHwndDialog,long idControl) {
   return setWindowComboBoxList(GetDlgItem(vHwndDialog,idControl));
   }



   HRESULT Property::getWindowComboBoxList(HWND vHwndControl) {

   if ( ! TYPE_ARRAY == type && ! TYPE_UNSPECIFIED == type ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to retrieve the list contents of a combo box into a property (%s), however, the property is either not an array property.\n\nMethod IGProperty::getWindowComboBoxList",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   HWND hwndControl = vHwndControl;

   char szClassName[64];
   GetClassName(hwndControl,szClassName,64);
   if ( strcmp(szClassName,"ComboBox") ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to retrieve the list contents of a combo box into a property (%s), however, the window is not a ComboBox control..\n\nMethod IGProperty::getWindowComboBoxList",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   long cntListItems = (long)SendMessage(hwndControl,CB_GETCOUNT,0L,0L);

   if ( ! pVariantArray ) {
      SAFEARRAYBOUND rgsaBound;
      rgsaBound.cElements = cntListItems;
      rgsaBound.lLbound = 1;
      pVariantArray = SafeArrayCreate(VT_BSTR,1,&rgsaBound);
   }

   variantStorage* pVariants;
   long cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);

   if ( cntListItems != cntArrayItems ) {
      long cntDim = SafeArrayGetDim(pVariantArray);
      if ( 1 != cntDim ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            if ( cntListItems < cntArrayItems ) 
               sprintf(szError,"There are fewer items in a combo box list than there are members of the array property.\nAnd the arry is not 1-dimensional.\n\nThe Properties Component cannot redimension the array property.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
            else
               sprintf(szError,"While attempting to retrieve the contents of a combo box, there are not as many elements in the array property than there are items in the combo box.\nIf the array property was a 1 dimensional array the Properties Component would have redimensioned it accordingly.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
      } else {
         SAFEARRAYBOUND rgsaBound;
         rgsaBound.cElements = cntListItems;
         SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
         SafeArrayRedim(pVariantArray,&rgsaBound);
         if ( cntArrayItems ) 
            delete [] pVariants;
         cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      }
   }

   delete [] pVariants;

   type = TYPE_ARRAY;
   everAssigned = true;

   variantIndex = (long)SendMessage(hwndControl,CB_GETCURSEL,0L,0L);

   getArrayFromComboBoxControl(hwndControl,pVariantArray);

   return S_OK;
   }


   HRESULT Property::getWindowItemComboBoxList(HWND vHwndDialog,long idControl) {
   return getWindowComboBoxList(GetDlgItem(vHwndDialog,idControl));
   }


   // List box selection

   HRESULT Property::getWindowListBoxSelection(HWND hwndControl) {

   char szValue[MAX_PROPERTY_SIZE];

   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to Property::getWindowListBoxSelection");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   long rv = (long)SendMessage(hwndControl,LB_GETCURSEL,0L,0L);
   if ( CB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"Property::getWindowListBoxSelection was called but the window provided is not a List-box control or there is no current selection\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   rv = (long)SendMessage(hwndControl,LB_GETTEXT,(WPARAM)rv,(LPARAM)szValue);
   if ( CB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"Property::getWindowListBoxSelection was called but the window provided is not a List-box control or there is no current selection\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   VARIANT_BOOL oldIgnoreSetAction = ignoreSetAction;
   ignoreSetAction = true;
   put_szValue(szValue);
   ignoreSetAction = oldIgnoreSetAction;

   return S_OK;
   }


   HRESULT Property::getWindowItemListBoxSelection(HWND hwndDialog,long idControl) {
   return getWindowListBoxSelection(GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Property::setWindowListBoxSelection(HWND hwndControl) {
   char szValue[MAX_PROPERTY_SIZE];
   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to IGProperty::setWindowListBoxSelection");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   HRESULT hr = get_szValue(szValue);
   if ( S_OK != hr ) return hr;
   LRESULT rv = SendMessage(hwndControl,LB_FINDSTRINGEXACT,-1L,(LPARAM)szValue);
   if ( LB_ERR == rv ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"The string %s was intended as the selection in a List box control, however, that value is not already in the List box control. Method IGProperty::setWindowListBoxSelection",propertyBinaryValue());
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   SendMessage(hwndControl,LB_SETCURSEL,rv,0);
   return S_OK;
   }


   HRESULT Property::setWindowItemListBoxSelection(HWND hwndDialog,long idControl) {
   return setWindowListBoxSelection(GetDlgItem(hwndDialog,idControl));
   }


   // List box list

   HRESULT Property::setWindowListBoxList(HWND vHwndControl) {

   HWND hwndControl = vHwndControl;

   if ( ! pVariantArray ) {
      SendMessage(hwndControl,LB_RESETCONTENT,0L,0L);
      return S_OK;
   }

   setListBoxControlFromArray(hwndControl,pVariantArray);

   if ( variantIndex > -1 ) 
      SendMessage(hwndControl,LB_SETCURSEL,variantIndex,0L);
   else
      SendMessage(hwndControl,LB_SETCURSEL,0L,0L);

   return S_OK;
   }



   HRESULT Property::setWindowItemListBoxList(HWND vHwndDialog,long idControl) {
   return setWindowListBoxList(GetDlgItem(vHwndDialog,idControl));
   }



   HRESULT Property::getWindowListBoxList(HWND vHwndControl) {

   HWND hwndControl = vHwndControl;

   if ( ! TYPE_ARRAY == type && ! TYPE_UNSPECIFIED == type ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt was made to retrieve the list contents of a list box into a property (%s), however, the property is either not an array property.\n\nMethod IGProperty::getWindowListBoxList",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   long cntListItems = (long)SendMessage(hwndControl,LB_GETCOUNT,0L,0L);
   if ( LB_ERR == cntListItems ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"An attempt to retrieve the list contents of a list box failed. Perhaps because the control is not a list box. Method IGProperty::setWindowListBoxList");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   type = TYPE_ARRAY;
   everAssigned = true;

   if ( ! pVariantArray ) {
      SAFEARRAYBOUND rgsaBound;
      rgsaBound.cElements = cntListItems;
      rgsaBound.lLbound = 1;
      pVariantArray = SafeArrayCreate(VT_BSTR,1,&rgsaBound);
   }

   variantStorage* pVariants;
   long cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);

   if ( cntListItems != cntArrayItems ) {
      long cntDim = SafeArrayGetDim(pVariantArray);
      if ( 1 != cntDim ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            if ( cntListItems < cntArrayItems ) 
               sprintf(szError,"There are fewer items in a list box than there are members of the array property.\nAnd the arry is not 1-dimensional.\n\nThe Properties Component cannot redimension the array property.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
            else
               sprintf(szError,"While attempting to retrieve the contents of a list box, there are not as many elements in the array property than there are items in the list box.\nIf the array property was a 1 dimensional array the Properties Component would have redimensioned it accordingly.\n\nThe property name is '%s'",name[0] ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
      } else {
         SAFEARRAYBOUND rgsaBound;
         rgsaBound.cElements = cntListItems;
         SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
         SafeArrayRedim(pVariantArray,&rgsaBound);
         if ( cntArrayItems ) 
            delete [] pVariants;
         cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      }
   }
   
   variantIndex = (long)SendMessage(hwndControl,LB_GETCURSEL,0L,0L);

   for ( long k = 0; k < cntArrayItems ; k++ ) {
      
      if ( pVariants[k].pszData ) 
         delete [] pVariants[k].pszData;
      pVariants[k].pszData = NULL;
      pVariants[k].cntszBytes = 0;
      long n = (long)SendMessage(hwndControl,LB_GETTEXTLEN,k,0L);
      if ( n > 0 ) {
         pVariants[k].cntszBytes = n + 1;
         pVariants[k].pszData = new char[n + 1];
         memset(pVariants[k].pszData,0,n + 1);
         SendMessage(hwndControl,LB_GETTEXT,k,(LPARAM)pVariants[k].pszData);
         replaceSafeArrayElement(pVariants[k].vt,pVariants[k].cntszBytes,pVariants[k].pszData,pVariantArray,pVariants[k].vArrayLocator);
      }

   }

   if ( cntArrayItems ) 
      delete [] pVariants;

   return S_OK;
   }


   HRESULT Property::getWindowItemListBoxList(HWND vHwndDialog,long idControl) {
   return getWindowListBoxList(GetDlgItem(vHwndDialog,idControl));
   }


   // Window contents - Vector values

   HRESULT Property::getWindowArrayValues(SAFEARRAY** ppsaHwndControls) {

   if ( ! ppsaHwndControls ) return E_POINTER;
   if ( ! *ppsaHwndControls ) return E_POINTER;

   if ( TYPE_UNSPECIFIED != type && TYPE_ARRAY != type ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::getWindowArrayValues was called with a type of property that is not a TYPE_ARRAY property.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   variantStorage* pVariantHwnds;
   long cntWindowItems = getSafeArrayVariants(*ppsaHwndControls,&pVariantHwnds);

   if ( ! cntWindowItems ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::getWindowArrayValues was called with an empty array of window control handles.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   variantStorage* pVariants;
   long cntArrayItems;
   if ( pVariantArray ) {
      cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      if ( ! cntArrayItems ) {
         long cntDim = SafeArrayGetDim(pVariantArray);
         if ( 1 != cntDim ) {
            delete [] pVariantHwnds;
            if ( debuggingEnabled ) {
               char szError[MAX_PROPERTY_SIZE];
               sprintf(szError,"IGProperty::getWindowArrayValues was called with an empty array property.\nIf the array were only 1 dimension then the Properties Component could have redimensioned it as needed.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
               MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
               return S_OK;
            }
            return E_FAIL;
         }
         SAFEARRAYBOUND rgsaBound;
         SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
         rgsaBound.cElements = cntWindowItems;
         SafeArrayRedim(pVariantArray,&rgsaBound);
         cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      }
   } else {
      SAFEARRAYBOUND rgsaBound;
      rgsaBound.cElements = cntWindowItems;
      rgsaBound.lLbound = 1;
      pVariantArray = SafeArrayCreate(VT_BSTR,1,&rgsaBound);
   }

   type = TYPE_ARRAY;

   everAssigned = true;

   if ( cntWindowItems > cntArrayItems ) {
      long cntDim = SafeArrayGetDim(pVariantArray);
      if ( 1 != cntDim ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::getWindowArrayValues was called with an array property that has fewer elements (%d) than provided windows (%d).\nIf the array were only 1 dimension then the Properties Component could have redimensioned it as needed.\n\nThe property name is \"%s\"",cntArrayItems,cntWindowItems,*name ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
      } else {
         SAFEARRAYBOUND rgsaBound;
         SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
         rgsaBound.cElements = cntWindowItems;
         SafeArrayRedim(pVariantArray,&rgsaBound);
         delete [] pVariants;
         cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      }
   } else {
      if ( cntArrayItems > cntWindowItems ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::getWindowArrayValues was called with an array property that has more elements (%d) than provided windows (%d).\nOnly the windows provided will be queried for their values, the property's array entries over this number will not be modified.\n\nThe property name is \"%s\"",cntArrayItems,cntWindowItems,*name ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
         cntArrayItems = cntWindowItems;
      }
   }

   long k = 0;
   for ( ; k < cntWindowItems && k < cntArrayItems; k++ ) {
      HWND hwnd = *(reinterpret_cast<HWND*>(pVariantHwnds[k].pData));
      getVariantFromWindow(hwnd,pVariants + k,"IGProperty::getWindowArrayValues");
   }

   delete [] pVariantHwnds;
   delete [] pVariants;

   return S_OK;
   }


   HRESULT Property::getWindowItemArrayValues(SAFEARRAY** ppsaHwndDialogs,SAFEARRAY** ppsaIdControls) {

   if ( ! ppsaHwndDialogs ) return E_POINTER;
   if ( ! *ppsaHwndDialogs ) return E_POINTER;
   if ( ! ppsaIdControls) return E_POINTER;
   if ( ! *ppsaIdControls ) return E_POINTER;

   long cntHWNDs = countSafeArrayItems(*ppsaHwndDialogs);
   long cntIDs = countSafeArrayItems(*ppsaIdControls);
   
   if ( cntHWNDs != cntIDs ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::getWindowItemArrayValues was called with an array of dialog windows that is not the same size as the array of control-ids.\nThe 2 passed arrays must be the same size.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return E_FAIL;
   }

   SAFEARRAY* psaHwndControls;

   SafeArrayCopy(*ppsaHwndDialogs,&psaHwndControls);

   HWND *pHWNDs;
   long *pIDs;
   HWND *pTargetHWNDs;

   SafeArrayAccessData(*ppsaHwndDialogs,reinterpret_cast<void**>(&pHWNDs));
   SafeArrayAccessData(*ppsaIdControls,reinterpret_cast<void**>(&pIDs));

   SafeArrayAccessData(psaHwndControls,reinterpret_cast<void**>(&pTargetHWNDs));

   for ( int k = 0; k < cntHWNDs; k++ ) {
      *pTargetHWNDs = GetDlgItem((HWND)*pHWNDs,*pIDs);
      pTargetHWNDs++;
      pHWNDs++;
      pIDs++;
   }
   
   HRESULT hr = getWindowArrayValues(&psaHwndControls);

   SafeArrayDestroy(psaHwndControls);

   return hr;
   }


   HRESULT Property::setWindowArrayValues(SAFEARRAY** ppsaHwndControls) {
   
   if ( ! ppsaHwndControls ) return E_POINTER;
   if ( ! *ppsaHwndControls ) return E_POINTER;

   if ( TYPE_ARRAY != type ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowArrayValues was called with a property that is not a TYPE_ARRAY property.\n\nThe property name is \"%s\"",*name ? name : "");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   
   if ( ! pVariantArray ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowArrayValues was called with an array property that has never been given a value.\n\nThe property name is \"%s\"",*name ? name : "");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   variantStorage* pVariantHwnds;
   long cntWindowItems = getSafeArrayVariants(*ppsaHwndControls,&pVariantHwnds);

   if ( ! cntWindowItems ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowArrayValues was called with an empty array of window control handles.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   variantStorage* pVariants;
   long cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
   if ( ! cntArrayItems ) {
      long cntDim = SafeArrayGetDim(pVariantArray);
      if ( 1 != cntDim ) {
         delete [] pVariantHwnds;
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowArrayValues was called with an empty array property.\nIf the array were only 1 dimension then the Properties Component could have redimensioned it as needed.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
            return S_OK;
         }
         return E_FAIL;
      }
      SAFEARRAYBOUND rgsaBound;
      SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
      rgsaBound.cElements = cntWindowItems;
      SafeArrayRedim(pVariantArray,&rgsaBound);
      cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
   }

   if ( cntWindowItems > cntArrayItems ) {
      long cntDim = SafeArrayGetDim(pVariantArray);
      if ( 1 != cntDim ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowArrayValues was called with an array property that has fewer elements (%d) than provided windows (%d).\nIf the array were only 1 dimension then the Properties Component would have redimensioned it as needed.\n\nThe property name is \"%s\"",cntArrayItems,cntWindowItems,*name ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
      } else {
         SAFEARRAYBOUND rgsaBound;
         SafeArrayGetLBound(pVariantArray,1,&rgsaBound.lLbound);
         rgsaBound.cElements = cntWindowItems;
         SafeArrayRedim(pVariantArray,&rgsaBound);
         delete [] pVariants;
         cntArrayItems = getSafeArrayVariants(pVariantArray,&pVariants);
      }
   } else {
      if ( cntArrayItems > cntWindowItems ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowArrayValues was called with an array property that has more elements (%d) than provided windows (%d).\nOnly the windows provided will be set, the property's array entries over this number will not be used.\n\nThe property name is \"%s\"",cntArrayItems,cntWindowItems,*name ? name : "<unnamed>");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         }
         cntArrayItems = cntWindowItems;
      }
   }

   for ( long k = 0; k < cntArrayItems; k++ ) {
      HWND hwnd = *(reinterpret_cast<HWND*>(pVariantHwnds[k].pData));
      setWindowFromVariant(hwnd,pVariants + k,"IGProperty::setWindowArrayValues");
   }

   delete [] pVariantHwnds;
   delete [] pVariants;

   return S_OK;
   }


   
   HRESULT Property::setWindowItemArrayValues(SAFEARRAY** ppsaHwndDialogs,SAFEARRAY** ppsaIdControls) {

   if ( ! ppsaHwndDialogs ) return E_POINTER;
   if ( ! *ppsaHwndDialogs ) return E_POINTER;
   if ( ! ppsaIdControls) return E_POINTER;
   if ( ! *ppsaIdControls ) return E_POINTER;

   long cntHWNDs = countSafeArrayItems(*ppsaHwndDialogs);
   long cntIDs = countSafeArrayItems(*ppsaIdControls);
   
   if ( cntHWNDs != cntIDs ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowItemArrayTexts was called with an array of dialog windows that is not the same size as the array of control-ids.\nThe 2 passed arrays must be the same size.\n\nThe property name is \"%s\"",*name ? name : "<unnamed>");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   SAFEARRAY* psaHwndControls;

   SafeArrayCopy(*ppsaHwndDialogs,&psaHwndControls);

   HWND *pHWNDs;
   long *pIDs;
   HWND *pTargetHWNDs;

   SafeArrayAccessData(*ppsaHwndDialogs,reinterpret_cast<void**>(&pHWNDs));
   SafeArrayAccessData(*ppsaIdControls,reinterpret_cast<void**>(&pIDs));

   SafeArrayAccessData(psaHwndControls,reinterpret_cast<void**>(&pTargetHWNDs));

   for ( int k = 0; k < cntHWNDs; k++ ) {
      *pTargetHWNDs = GetDlgItem(*pHWNDs,*pIDs);
      pTargetHWNDs++;
      pHWNDs++;
      pIDs++;
   }
   
   setWindowArrayValues(&psaHwndControls);

   SafeArrayDestroy(psaHwndControls);

   return S_OK;
   }


   // Check boxes

   HRESULT Property::getWindowChecked(HWND hwndControl) {
   char szError[MAX_PROPERTY_SIZE];
   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szError,"A NULL window handle was passed to IGProperty::getWindowChecked");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   VARIANT_BOOL oldIgnoreSetAction = ignoreSetAction;
   ignoreSetAction = true;
   long rv = (long)SendMessage(hwndControl,BM_GETCHECK,0L,0L);
   switch ( rv ) {
   case BST_CHECKED:
      put_boolValue(1);
      break;
   case BST_UNCHECKED:
      put_boolValue(0);
      break;
   case BST_INDETERMINATE:
      put_boolValue(0);
      if ( debuggingEnabled ) {
         sprintf(szError,"Property::getWindowChecked was called with a Check box whose value is in the inderminate state. The Properties component handles true or false and does not know other than to use 0 (false) for this case.\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      break;
   default:
      if ( debuggingEnabled ) {
         sprintf(szError,"Property::getWindowChecked was called but the window provided is not a Check-Box/radioButton control or there is no current selection\n\nDisable debugging (debuggingEnabled(false)) to prohibit this message");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
      }
      return E_FAIL;
   }
   ignoreSetAction = oldIgnoreSetAction;
   return S_OK;
   }


   HRESULT Property::getWindowItemChecked(HWND hwndDialog,long idControl) {
   return getWindowChecked(GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Property::setWindowChecked(HWND hwndControl) {

   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"A NULL window handle was passed to IGProperty::setWindowChecked");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }

   updateFromDirectAccess();

   short isChecked = 0;
   char szYesNo[32];

   switch ( type ) {
   case TYPE_BOOL:
      isChecked = (v.scalar.boolValue != 0);
      break;
   case TYPE_LONG:
      isChecked = (v.scalar.longValue != 0);
      break;
   case TYPE_DOUBLE:
      isChecked = (v.scalar.doubleValue != 0.0);
      break;
      break;

   case TYPE_SZSTRING:
   case TYPE_STRING: {
      if ( type == TYPE_STRING ) {
         if ( wcslen(bstrValue) < 1 ) {
            if ( debuggingEnabled ) {
               char szError[MAX_PROPERTY_SIZE];
               sprintf(szError,"IGProperty::setWindowChecked was called for a property (%s) of TYPE_STRING, however the current string is empty.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message",name);
               MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
               return S_OK;
            }
            return E_FAIL;
         }
         memset(szYesNo,0,sizeof(szYesNo));
         WideCharToMultiByte(CP_ACP,0,bstrValue,-1,szYesNo,32,0,0);
      }
      else {
         if ( v.binarySize < 1 ) {
            if ( debuggingEnabled ) {
               char szError[MAX_PROPERTY_SIZE];
               sprintf(szError,"IGProperty::setWindowChecked was called for a property (%s)of TYPE_SZSTRING, however the current string is empty.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message",name);
               MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
               return S_OK;
            }
            return E_FAIL;
         }
         strcpy(szYesNo,reinterpret_cast<char*>(v.binaryValue));
      }
      strlwr(szYesNo);
      if ( szYesNo[0] == 'y' ) { 
         isChecked = 1;
      } else {
         if ( szYesNo[0] == '1' ) 
            isChecked = 1;
      }
      }
      break;

   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
      if ( v.binarySize < 0 ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowChecked was called for a property (%s) of TYPE_BINARY, however the current data is empty.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message",name);
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
            return S_OK;
         }
         return E_FAIL;
      }
      break;

   default:
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowChecked was called for a property (%s) of a non-scalar type, i.e., the binary nature of checked/unchecked cannot be determined from the data in the property.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message",name);
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   SendMessage(hwndControl,BM_SETCHECK,(WPARAM)( isChecked ? BST_CHECKED : BST_UNCHECKED),0L);
   return S_OK;
   }


   HRESULT Property::setWindowItemChecked(HWND hwndDialog,long idControl) {
   return setWindowChecked(GetDlgItem(hwndDialog,idControl));
   }


   HRESULT Property::setWindowEnabled(HWND hwndControl) {
   char szValue[MAX_PROPERTY_SIZE];
   if ( ! hwndControl ) {
      if ( debuggingEnabled ) {
         sprintf(szValue,"A NULL window handle was passed to IGProperty::setWindowEnabled");
         MessageBox(NULL,szValue,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   long longValue = atol(reinterpret_cast<char*>(propertyScalarValue()));
   switch ( type ) {
   case TYPE_BOOL:
   case TYPE_LONG:
   case TYPE_DOUBLE:
      break;

   case TYPE_SZSTRING:
      if ( v.binarySize < 1 ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowEnabled was called for a property of TYPE_SZSTRING, however the current string is empty.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
            return S_OK;
         }
         return E_FAIL;
      }
      break;

   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
      if ( v.binarySize < 0 ) {
         if ( debuggingEnabled ) {
            char szError[MAX_PROPERTY_SIZE];
            sprintf(szError,"IGProperty::setWindowEnabled was called for a property of TYPE_BINARY, however the current data is empty.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message");
            MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
            return S_OK;
         }
         return E_FAIL;
      }
      break;

   default:
      if ( debuggingEnabled ) {
         char szError[MAX_PROPERTY_SIZE];
         sprintf(szError,"IGProperty::setWindowEnabled was called for a property of a non-scalar type, i.e., the binary nature of checked/unchecked cannot be determined from the data in the property.\n\nCall IGProperty::debuggingEnabled(false) to prevent this message");
         MessageBox(NULL,szError,"GSystem Properties Component usage note.",MB_OK);
         return S_OK;
      }
      return E_FAIL;
   }
   EnableWindow(hwndControl,longValue == 0 ? FALSE : TRUE);
   return S_OK;
   }


   HRESULT Property::setWindowItemEnabled(HWND hwndDialog,long idControl) {
   return setWindowEnabled(GetDlgItem(hwndDialog,idControl));
   }