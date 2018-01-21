// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Property.h"

#include "utils.h"


   HRESULT Property::put_value(VARIANT newValue) {
   return put_variantValue(newValue);
   }


   HRESULT Property::get_value(VARIANT* pValue) {
   return get_variantValue(pValue);
   }


   // Values
 
   HRESULT Property::put_szValue(char *sIn) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_SZSTRING;

   everAssigned = true;

   if ( v.binaryValue )
      delete [] v.binaryValue;

   v.binarySize = (long)strlen(sIn) + 1;
   v.binaryValue = new BYTE[v.binarySize];
   memcpy(v.binaryValue,sIn,v.binarySize);

   if ( bstrValue )
      SysFreeString(bstrValue);
   bstrValue = SysAllocStringLen(NULL,v.binarySize);
   MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v.binaryValue),-1,bstrValue,v.binarySize);
   
   setBinaryValueDirectAccess();
   setScalarValueAndDirectAccess();

   if ( ! ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   setArrayValue();

   return S_OK;

   }
 
 
   HRESULT Property::get_szValue(char *sOut) {
   if ( ! propertyBinaryValue() ) {
      sOut[0] = '\0';
      return S_OK;
   }
   strcpy(sOut,reinterpret_cast<char*>(propertyBinaryValue()));
   return S_OK;
   }
 

   HRESULT Property::put_stringValue(BSTR bstrNewValue) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_STRING;

   everAssigned = true;

   if ( bstrNewValue ) {
      if ( bstrValue ) 
         SysReAllocString(&bstrValue,bstrNewValue);
      else
         bstrValue = SysAllocString(bstrNewValue);
   } else {
      bstrValue = SysAllocStringLen(L"",0);
   }

   if ( v.binarySize )
      delete [] v.binaryValue;

   v.binarySize = (long)wcslen(bstrValue) + 1;
   v.binaryValue = new BYTE[v.binarySize];
   WideCharToMultiByte(CP_ACP,0,bstrValue,-1,reinterpret_cast<char *>(v.binaryValue),v.binarySize,0,0);

   setBinaryValueDirectAccess();
   setScalarValueAndDirectAccess();

   if ( ! ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   setArrayValue();

   return S_OK;
   }

   
   HRESULT Property::get_stringValue(BSTR* pbstrValue) {
   updateFromDirectAccess();
   if ( TYPE_STRING == type ) 
      *pbstrValue = SysAllocString(bstrValue);
   else if ( TYPE_SZSTRING == type ) {
      *pbstrValue = SysAllocStringLen(NULL,(UINT)strlen((char *)v.binaryValue));
      MultiByteToWideChar(CP_ACP,0,(char *)v.binaryValue,-1,*pbstrValue,(int)strlen((char *)v.binaryValue));
   } else {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::get_stringValue was called for a type of property (%ld) that is not compatible with this operation\n\n(Only types TYPE_STRING or TYPE_SZSTRING are compatible with this operation)",type);
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
   }
   return S_OK;
   }


   HRESULT Property::put_binaryValue(long cbSize,BYTE* bIn) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_BINARY;

   if ( type != TYPE_BINARY && type != TYPE_RAW_BINARY && type != TYPE_OBJECT_STORAGE_ARRAY ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::put_binaryValue was called for a type of property (%ld) that is not compatible with this operation\n\n(Only type TYPE_BINARY is compatible with this operation)",type);
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
      return E_FAIL;
   }

   everAssigned = true;

   if ( v.binarySize ) 
      delete [] v.binaryValue;

   if ( ! bIn ) {
      v.binarySize = 0;
      v.binaryValue = NULL;
      return S_OK;
   }

   v.binarySize = cbSize;
   v.binaryValue = new BYTE[v.binarySize];

   memcpy(v.binaryValue,bIn,v.binarySize);

   if ( bstrValue )
      SysFreeString(bstrValue);

   bstrValue = SysAllocStringLen(NULL,v.binarySize);
   MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v.binaryValue),v.binarySize,bstrValue,v.binarySize);

   setBinaryValueDirectAccess();
   setScalarValueAndDirectAccess();

   if ( ! ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   return S_OK;
   }
 

   HRESULT Property::get_binaryValue(long cbSize,BYTE** bOut) {

   if ( type != TYPE_BINARY && type != TYPE_RAW_BINARY ) {
      if ( debuggingEnabled ) {
         char szWarning[MAX_PROPERTY_SIZE];
         sprintf(szWarning,"IGProperty::get_binaryValue was called for a type of property (%ld) "
                              "that is not compatible with this operation\n\n(Only type TYPE_BINARY is compatible with this operation)",type);
         MessageBox(NULL,szWarning,"GSystem Properties Component usage note",MB_OK);
      }
      return E_FAIL;
   }

   updateFromDirectAccess();

   memcpy(*bOut,v.binaryValue,min(v.binarySize,cbSize));

   return S_OK;
   }
 
 
 
   HRESULT Property::put_longValue(long inValue) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_LONG;

   everAssigned = true;

   if ( v.binarySize )
      delete [] v.binaryValue;

   v.binarySize = 32;
   v.binaryValue = new BYTE[v.binarySize];
   memset(v.binaryValue,0,32);
   sprintf(reinterpret_cast<char *>(v.binaryValue),"%ld",inValue);

   if ( bstrValue )
      SysFreeString(bstrValue);
   bstrValue = SysAllocStringLen(NULL,v.binarySize);
   MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v.binaryValue),-1,bstrValue,v.binarySize);

   setBinaryValueDirectAccess();
   if ( ! setScalarValueAndDirectAccess() ) {            // 09/21/2002: set scalar value if setScalarValue... returns false
      v.scalar.longValue = inValue;
   }

   if ( ! ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   setArrayValue();

   return S_OK;
   }
 

   HRESULT Property::get_longValue(long *getValue) {
   updateFromDirectAccess();
   switch ( type ) {
   case TYPE_BOOL:
      *getValue = v.scalar.boolValue ? 1L : 0L;
      return S_OK;
   case TYPE_DOUBLE:
      *getValue = (long)v.scalar.doubleValue;
      return S_OK;
   default:
      *getValue = v.scalar.longValue;
      break;
   }
   return S_OK;
   }
 
 
 
   HRESULT Property::put_doubleValue(double inValue) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_DOUBLE;

   everAssigned = true;

   if ( v.binarySize )
      delete [] v.binaryValue;

   v.binarySize = MAX_PATH;
   v.binaryValue = new BYTE[v.binarySize];
   memset(v.binaryValue,0,v.binarySize);
   sprintf(reinterpret_cast<char *>(v.binaryValue),"%lf",inValue);

   if ( bstrValue )
      SysFreeString(bstrValue);
   bstrValue = SysAllocStringLen(NULL,v.binarySize);
   MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v.binaryValue),-1,bstrValue,v.binarySize);

   setBinaryValueDirectAccess();
   if ( ! setScalarValueAndDirectAccess() ) {         // 9/21/2002: set scalar value if setScalarValue... returns false
      v.scalar.doubleValue = inValue;
   }

   if ( !ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   setArrayValue();

   return S_OK;
   }
 

   HRESULT Property::get_doubleValue(double *getValue) {
   updateFromDirectAccess();
   switch ( type ) {
   case TYPE_LONG:
      *getValue = (double)v.scalar.longValue;
      return S_OK;
   case TYPE_BOOL:
      *getValue = v.scalar.boolValue ? 1.0 : 0.0;
      return S_OK;
   default:
      *getValue = v.scalar.doubleValue;
      break;
   }
   return S_OK;
   }
 
 
   HRESULT Property::put_boolValue(short inValue) {

   if ( type == TYPE_UNSPECIFIED ) type = TYPE_BOOL;

   everAssigned = true;

   if ( v.binarySize )
      delete [] v.binaryValue;

   v.binarySize = 32;
   v.binaryValue = new BYTE[v.binarySize];
   memset(v.binaryValue,0,32);
   sprintf(reinterpret_cast<char *>(v.binaryValue),"%d",inValue ? 1 : 0);

   if ( bstrValue )
      SysFreeString(bstrValue);
   bstrValue = SysAllocStringLen(NULL,v.binarySize);
   MultiByteToWideChar(CP_ACP,0,reinterpret_cast<char*>(v.binaryValue),-1,bstrValue,v.binarySize);

   setBinaryValueDirectAccess();
   if ( ! setScalarValueAndDirectAccess() ) {      // 9/21/2002: Set scalar value if setScalarValue... returns false
      v.scalar.boolValue = inValue;
   }

   if ( ! ignoreSetAction ) 
      if ( pIPropertyClient ) 
         pIPropertyClient -> Changed(this);

   isDirty = true;

   setArrayValue();

   return S_OK;
   }


   HRESULT Property::get_boolValue(short *getValue) {
   updateFromDirectAccess();
   *getValue = v.scalar.boolValue;
   return S_OK;
   }
 

   HRESULT Property::put_arrayIndex(long newIndex) {
   variantIndex = newIndex;
   return S_OK;
   }


   HRESULT Property::get_arrayIndex(long* pIndex) {
   if ( ! pIndex ) return E_POINTER;
   *pIndex = variantIndex;
   return S_OK;
   }


   int Property::setBinaryValueDirectAccess() {
   switch ( type ) {
   case TYPE_OBJECT_STORAGE_ARRAY: // <-- 09/21/07
   case TYPE_RAW_BINARY:
   case TYPE_BINARY:
   case TYPE_SZSTRING:
      if ( directAccessAddress ) {
         memset(directAccessAddress,0,directAccessSize);
         memcpy(directAccessAddress,v.binaryValue,min(v.binarySize,directAccessSize));
      }
      break;

   case TYPE_STRING: 
      if ( directAccessAddress )
         SysReAllocString(reinterpret_cast<BSTR*>(directAccessAddress),bstrValue);
      break;

   }
   return 0;
   }


   int Property::setScalarValueAndDirectAccess() {
   memset(&v.scalar,0,sizeof(v.scalar));
   int rc = 0;
   switch ( type ) {
   case TYPE_BOOL:
      v.scalar.boolValue = static_cast<short>(atol(reinterpret_cast<char *>(v.binaryValue)));
      if ( directAccessAddress ) 
         memcpy(directAccessAddress,&v.scalar.boolValue,min(sizeof(v.scalar.boolValue),directAccessSize));
      rc = 1;
      break;

   case TYPE_LONG:
      v.scalar.longValue = atol(reinterpret_cast<char *>(v.binaryValue));
      if ( directAccessAddress ) 
         memcpy(directAccessAddress,&v.scalar.longValue,min(sizeof(v.scalar.longValue),directAccessSize));
      rc = 1;
      break;

   case TYPE_DOUBLE:
      v.scalar.doubleValue = atof(reinterpret_cast<char *>(v.binaryValue));
      if ( directAccessAddress ) 
         memcpy(directAccessAddress,&v.scalar.doubleValue,min(sizeof(v.scalar.doubleValue),directAccessSize));
      rc = 1;
      break;

   } 
   return rc;
   }


   void *Property::pointer() {
   return v.binaryValue;
   }
