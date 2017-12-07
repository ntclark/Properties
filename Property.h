
#pragma once

#include <limits.h>
#include <float.h>

#include "utils.h"

#include "Properties.h"

   struct variantStorage {
      VARTYPE vt;
      long cntBytes;
      long cntszBytes;
      long* vArrayLocator;
      variantStorage* pSubElements;
      BYTE* pData;
      char* pszData;
      VARIANT theVariant;
      variantStorage() : vt(0), cntBytes(0), cntszBytes(0), pData(NULL), vArrayLocator(NULL), pszData(NULL), pSubElements(NULL) { VariantInit(&theVariant); };
      ~variantStorage() { if ( pData ) delete [] pData; 
                          if ( vArrayLocator ) delete [] vArrayLocator; 
                          if ( pszData ) delete [] pszData;
                          if ( pSubElements ) delete [] pSubElements; 
                          GVariantClear(&theVariant); };
   };


   class Property : IGProperty {
   public:
   
      Property();
      ~Property();
   
     // IUnknown methods
   
      STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
      STDMETHOD_(ULONG, AddRef)();
      STDMETHOD_(ULONG, Release)();
   
    // IDispatch

      STDMETHOD(GetTypeInfoCount)(UINT *pctinfo);
      STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
      STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgdispid);
      STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);

      STDMETHOD(put_value)(VARIANT newValue);
      STDMETHOD(get_value)(VARIANT* theValue);

      STDMETHOD(put_name)(BSTR);
      STDMETHOD(get_name)(BSTR*);
      STDMETHOD(put_type)(PropertyType);
      STDMETHOD(get_type)(PropertyType*);
      STDMETHOD(put_size)(long setSize);
      STDMETHOD(get_size)(long *getSize);

      STDMETHOD(put_isDirty)(short setDirtyState);
      STDMETHOD(get_isDirty)(short *getDirtyState);
      STDMETHOD(get_everAssigned)(short* getEverAssigned);
      STDMETHOD(put_ignoreSetAction)(VARIANT_BOOL );
      STDMETHOD(get_ignoreSetAction)(VARIANT_BOOL *);
      STDMETHOD(put_debuggingEnabled)(VARIANT_BOOL);
      STDMETHOD(get_debuggingEnabled)(VARIANT_BOOL*);
      STDMETHOD(get_binaryData)(BYTE **getAddress);
      STDMETHOD(get_compressedName)(BSTR*);
      STDMETHOD(put_encodedText)(BSTR);
      STDMETHOD(get_encodedText)(BSTR*);
      STDMETHOD(get_variantType)(VARTYPE*);

      STDMETHOD(put_longValue)(long);
      STDMETHOD(get_longValue)(long *);
      STDMETHOD(put_doubleValue)(double);
      STDMETHOD(get_doubleValue)(double *);
      STDMETHOD(put_binaryValue)(long cbSize,BYTE *);
      STDMETHOD(get_binaryValue)(long cbSize,BYTE **);
      STDMETHOD(put_szValue)(char *);
      STDMETHOD(get_szValue)(char *);
      STDMETHOD(put_stringValue)(BSTR);
      STDMETHOD(get_stringValue)(BSTR*);
      STDMETHOD(put_arrayValue)(SAFEARRAY*);
      STDMETHOD(get_arrayValue)(SAFEARRAY**);
      STDMETHOD(put_boolValue)(VARIANT_BOOL);
      STDMETHOD(get_boolValue)(VARIANT_BOOL*);

      STDMETHOD(put_variantValue)(VARIANT);
      STDMETHOD(get_variantValue)(VARIANT*);

      STDMETHOD(put_arrayIndex)(long);
      STDMETHOD(get_arrayIndex)(long*);

      STDMETHOD(put_arrayElement)(long,VARIANT);
      STDMETHOD(get_arrayElement)(long,VARIANT*);
      STDMETHOD(get_arrayElementCount)(long*);
      STDMETHOD(clearArray)();

      STDMETHOD(advise)(IGPropertyClient *);
      STDMETHOD(directAccess)(enum PropertyType,void* directAccess,long directAccessSize);
      STDMETHOD(copyTo)(IGProperty* pTheDestination);
      STDMETHOD(addStorageObject)(IUnknown*);
      STDMETHOD(removeStorageObject)(IUnknown*);
      STDMETHOD(clearStorageObjects)();
      STDMETHOD(get_storedObjectCount)(long*);
      STDMETHOD(writeStorageObjects)();
      STDMETHOD(readStorageObjects)();

      // Window Contents Support

      STDMETHOD(getWindowValue)(HWND hwndControl);
      STDMETHOD(getWindowItemValue)(HWND hwndDialog,long idControl);
      STDMETHOD(getWindowText)(HWND hwndControl);
      STDMETHOD(getWindowItemText)(HWND hwndDialog,long idControl);
      STDMETHOD(setWindowText)(HWND hwndControl);
      STDMETHOD(setWindowItemText)(HWND hwndDialog,long idControl);

      // Combo box selection
      STDMETHOD(setWindowComboBoxSelection)(HWND hwndControl);
      STDMETHOD(setWindowItemComboBoxSelection)(HWND hwndControl,long idControl);
      STDMETHOD(getWindowComboBoxSelection)(HWND hwndControl);
      STDMETHOD(getWindowItemComboBoxSelection)(HWND hwndControl,long idControl);

      // Combo box list
      STDMETHOD(setWindowComboBoxList)(HWND hwndControl);
      STDMETHOD(setWindowItemComboBoxList)(HWND hwndDialog,long idControl);
      STDMETHOD(getWindowComboBoxList)(HWND hwndControl);
      STDMETHOD(getWindowItemComboBoxList)(HWND hwndDialog,long idControl);

      // List box selection
      STDMETHOD(setWindowListBoxSelection)(HWND hwndControl);
      STDMETHOD(setWindowItemListBoxSelection)(HWND hwndControl,long idControl);
      STDMETHOD(getWindowListBoxSelection)(HWND hwndControl);
      STDMETHOD(getWindowItemListBoxSelection)(HWND hwndControl,long idControl);

      // List box list
      STDMETHOD(setWindowListBoxList)(HWND hwndControl);
      STDMETHOD(setWindowItemListBoxList)(HWND hwndDialog,long idControl);
      STDMETHOD(getWindowListBoxList)(HWND hwndControl);
      STDMETHOD(getWindowItemListBoxList)(HWND hwndDialog,long idControl);
   
      // array properties
      STDMETHOD(getWindowArrayValues)(SAFEARRAY** hwndControl);
      STDMETHOD(getWindowItemArrayValues)(SAFEARRAY** hwndDialog,SAFEARRAY** idControl);
      STDMETHOD(setWindowArrayValues)(SAFEARRAY** hwndControl);
      STDMETHOD(setWindowItemArrayValues)(SAFEARRAY** hwndDialog,SAFEARRAY** idControl);

      // Check boxes
      STDMETHOD(setWindowChecked)(HWND hwndControl);
      STDMETHOD(setWindowItemChecked)(HWND hwndDialog,long idControl);
      STDMETHOD(getWindowChecked)(HWND hwndControl);
      STDMETHOD(getWindowItemChecked)(HWND hwndDialog,long idControl);

      // Enabled/Disabled
      STDMETHOD(setWindowEnabled)(HWND hwndControl);
      STDMETHOD(setWindowItemEnabled)(HWND hwndDialog,long idControl);

      STDMETHOD(assign)(void* anyValue,long valueLength);

      void * STDMETHODCALLTYPE pointer();

   private:

      int updateFromDirectAccess();
      int setScalarValueAndDirectAccess();
      int setBinaryValueDirectAccess();
      HRESULT toEncodedText(BSTR*);
      HRESULT fromEncodedText(BSTR);

      BYTE* propertyBinaryValue();
      BYTE* propertyScalarValue();
      int setArrayValue();

      long getSafeArrayVariants(SAFEARRAY* psa,variantStorage** ppVariants);
      int getVariantLocations(SAFEARRAY* pVariantArray,List<long>& vLocations,long *vIndexes,long dim,long cntDims);

      int replaceSafeArrayElement(VARTYPE vt,unsigned int cntBytes,char *pData,SAFEARRAY* psa,long *vIndexes);
      int replaceSafeArrayVariant(SAFEARRAY* psa,long *vIndexes,VARIANT* pv);

      long setWindowFromVariant(HWND hwnd,variantStorage* ps,char* szMethodName);
      long getVariantFromWindow(HWND hwnd,variantStorage* ps,char* szMethodName);

      long setEditControlFromVariant(HWND hwnd,variantStorage* ps);
      long setButtonControlFromVariant(HWND hwnd,variantStorage* ps);
      long setComboBoxControlFromVariant(HWND hwnd,variantStorage* ps);
      long setListBoxControlFromVariant(HWND hwnd,variantStorage* ps);
      long setScrollBarControlFromVariant(HWND hwnd,variantStorage* ps);
      long setStaticControlFromVariant(HWND hwnd,variantStorage* ps);

      long getVariantFromEditControl(HWND hwnd,variantStorage* ps);
      long getVariantFromButtonControl(HWND hwnd,variantStorage* ps);
      long getVariantFromComboBoxControl(HWND hwnd,variantStorage* ps);
      long getVariantFromListBoxControl(HWND hwnd,variantStorage* ps);
      long getVariantFromScrollBarControl(HWND hwnd,variantStorage* ps);
      long getVariantFromStaticControl(HWND hwnd,variantStorage* ps);

      long setComboBoxControlFromArray(HWND hwnd,SAFEARRAY* psa);
      long setListBoxControlFromArray(HWND hwnd,SAFEARRAY* psa);
      long getArrayFromComboBoxControl(HWND hwnd,SAFEARRAY* psa);
      long getArrayFromListBoxControl(HWND hwnd,SAFEARRAY* psa);

      long getStringFromWindow(HWND hwnd,variantStorage* ps);

      char name[64];
      PropertyType type;
      VARIANT_BOOL ignoreSetAction;
      VARIANT_BOOL isDirty;
      VARIANT_BOOL everAssigned;
      VARIANT_BOOL debuggingEnabled;
      HWND hwndClientWindow;
      LRESULT (EXPENTRY *oldClientWindowHandler)(HWND,UINT,WPARAM,LPARAM);
      
      struct persistenceInterfaces {
         persistenceInterfaces(IPersistStream *pstream,IPersistStorage *pstorage) : pIPersistStream(pstream), pIPersistStorage(pstorage) {};
         IPersistStream *pIPersistStream;
         IPersistStorage *pIPersistStorage;
         long Release() { if ( pIPersistStream ) pIPersistStream -> Release(); if ( pIPersistStorage ) pIPersistStorage -> Release(); return 0; };
      };

      List<persistenceInterfaces> storageObjects;

      IGPropertyClient *pIPropertyClient;
   
      struct {
         long binarySize;
         BYTE* binaryValue;
         union {
            short boolValue;
            long longValue;
            double doubleValue;
         } scalar;
      } v;

      BSTR bstrValue;

      SAFEARRAY* pVariantArray;
      int variantIndex;

      void* directAccessAddress;
      long directAccessSize;
   
   };
