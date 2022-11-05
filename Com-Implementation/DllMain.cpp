
#include "Properties.h"

#include <stdio.h>
#include <shlwapi.h>

#include "utils.h"
#include "resource.h"
#include "GSystemsProperties.h"

#include "Properties_i.c"

   HMODULE gsProperties_hModule;

   char szModuleName[MAX_PATH];
   BSTR wstrModuleName;

   ITypeInfo *pITypeInfo_IProperties = NULL;
   ITypeInfo *pITypeInfo_IProperty = NULL;

   ObjectFactory *propertyFactory = NULL;
   ObjectFactory *propertiesFactory = NULL;

   BOOL WINAPI DllMain(HANDLE module,ULONG flag,void *) {

   switch ( flag ) {

   case DLL_PROCESS_ATTACH: {

      gsProperties_hModule = reinterpret_cast<HMODULE>(module);

      memset(szModuleName,0,sizeof(szModuleName));

      GetModuleFileName(gsProperties_hModule,szModuleName,MAX_PATH);

      wstrModuleName = SysAllocStringLen(NULL,MAX_PATH);

      char szLibraryName[MAX_PATH];
      memset(szLibraryName,0,MAX_PATH);
      sprintf(szLibraryName,"%s\\1",szModuleName);

      MultiByteToWideChar(CP_ACP,0,szLibraryName,-1,wstrModuleName,MAX_PATH);  

      ITypeLib *ptLib;

      LoadTypeLib(wstrModuleName,&ptLib);

      MultiByteToWideChar(CP_ACP,0,szModuleName,-1,wstrModuleName,MAX_PATH);

      ptLib -> GetTypeInfoOfGuid(IID_IGProperties,&pITypeInfo_IProperties);

      ptLib -> GetTypeInfoOfGuid(IID_IGProperty,&pITypeInfo_IProperty);

      ptLib -> Release();

      propertyFactory = new ObjectFactory(CLSID_InnoVisioNateProperty);
      propertiesFactory = new ObjectFactory(CLSID_InnoVisioNateProperties);
      }

      break;
 
   case DLL_PROCESS_DETACH:
      delete propertyFactory;
      delete propertiesFactory;
      pITypeInfo_IProperties -> Release();
      pITypeInfo_IProperty -> Release();
      SysFreeString(wstrModuleName);
      break;
               
   default:
      break;
   }
   return TRUE;
   }


   STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppObject) {

   *ppObject = NULL;

   if ( rclsid == CLSID_InnoVisioNateProperties )
      return propertiesFactory -> QueryInterface(riid,ppObject);

   if ( rclsid == CLSID_InnoVisioNateProperty )
      return propertyFactory -> QueryInterface(riid,ppObject);

   return CLASS_E_CLASSNOTAVAILABLE;

   }
 

   char *OBJECT_NAME[] = {"InnoVisioNate.CVGProperties","InnoVisioNate.CVGProperty"};
   char *OBJECT_NAME_V[] = {"InnoVisioNate.CVGProperties.1","InnoVisioNate.CVGProperty.1"};
   char *OBJECT_VERSION[] = {"1.0","1.0"};
   char *OBJECT_DESCRIPTION[] = {"InnoVisioNate CursiVision Properties Object","InnoVisioNate CursiVision Property Object"};
   GUID OBJECT_CLSID[] = {CLSID_InnoVisioNateProperties,CLSID_InnoVisioNateProperty};
   GUID OBJECT_LIBID[] = {LIBID_InnoVisioNateProperties,GUID_NULL};

   STDAPI DllRegisterServer() {

   HRESULT rc = S_OK;
   ITypeLib *ptLib;
   HKEY keyHandle,clsidHandle;
   DWORD disposition;
   char szTemp[256],szCLSID[256];
   LPOLESTR oleString;

   if ( S_OK != LoadTypeLib(wstrModuleName,&ptLib) )
      rc = ResultFromScode(SELFREG_E_TYPELIB);
   else
      if ( S_OK != RegisterTypeLib(ptLib,wstrModuleName,NULL) )
         rc = ResultFromScode(SELFREG_E_TYPELIB);

   for ( long cycle = 0; cycle < 2; cycle++ ) {
  
      StringFromCLSID(OBJECT_CLSID[cycle],&oleString);

      WideCharToMultiByte(CP_ACP,0,oleString,-1,szCLSID,256,0,0);

      RegOpenKeyEx(HKEY_CLASSES_ROOT,"CLSID",0,KEY_CREATE_SUB_KEY,&keyHandle);
    
         RegCreateKeyEx(keyHandle,szCLSID,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&clsidHandle,&disposition);
         sprintf(szTemp,OBJECT_DESCRIPTION[cycle]);
         RegSetValueEx(clsidHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
         sprintf(szTemp,"Control");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,"");
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
         sprintf(szTemp,"ProgID");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,OBJECT_NAME_V[cycle]);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
         sprintf(szTemp,"InprocServer");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szModuleName,(DWORD)strlen(szModuleName));
    
         sprintf(szTemp,"InprocServer32");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szModuleName,(DWORD)strlen(szModuleName));
//         RegSetValueEx(keyHandle,"ThreadingModel",0,REG_SZ,(BYTE *)"Free",5);
         RegSetValueEx(keyHandle,"ThreadingModel",0,REG_SZ,(BYTE *)"Both",5);
//         RegSetValueEx(keyHandle,"ThreadingModel",0,REG_SZ,(BYTE *)"Apartment",9);
    
         sprintf(szTemp,"LocalServer");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szModuleName,(DWORD)strlen(szModuleName));
       
         if ( ! ( OBJECT_LIBID[cycle] == GUID_NULL ) ) {
            sprintf(szTemp,"TypeLib");
            RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
       
            StringFromCLSID(OBJECT_LIBID[cycle],&oleString);
            WideCharToMultiByte(CP_ACP,0,oleString,-1,szTemp,256,0,0);
            RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
         }

         sprintf(szTemp,"ToolboxBitmap32");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
// //      sprintf(szTemp,"%s, 1",szModuleName);
// //      RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szModuleName));
    
         sprintf(szTemp,"Version");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,OBJECT_VERSION[cycle]);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
         sprintf(szTemp,"MiscStatus");
         RegCreateKeyEx(clsidHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,"0");
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
         sprintf(szTemp,"1");
         RegCreateKeyEx(keyHandle,szTemp,0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,"%ld",
                    OLEMISC_ALWAYSRUN |
                    OLEMISC_ACTIVATEWHENVISIBLE | 
                    OLEMISC_RECOMPOSEONRESIZE | 
                    OLEMISC_INSIDEOUT |
                    OLEMISC_SETCLIENTSITEFIRST |
                    OLEMISC_CANTLINKINSIDE );
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
      RegCreateKeyEx(HKEY_CLASSES_ROOT,OBJECT_NAME[cycle],0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegCreateKeyEx(keyHandle,"CurVer",0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         sprintf(szTemp,OBJECT_NAME_V[cycle]);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szTemp,(DWORD)strlen(szTemp));
    
      RegCreateKeyEx(HKEY_CLASSES_ROOT,OBJECT_NAME_V[cycle],0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegCreateKeyEx(keyHandle,"CLSID",0,NULL,REG_OPTION_NON_VOLATILE,KEY_ALL_ACCESS,NULL,&keyHandle,&disposition);
         RegSetValueEx(keyHandle,NULL,0,REG_SZ,(BYTE *)szCLSID,(DWORD)strlen(szCLSID));
    
   }

   return S_OK;
   }
  
  
   STDAPI DllUnregisterServer() {

   HRESULT rc = S_OK;
   HKEY keyHandle;
   char szCLSID[256];
   LPOLESTR oleString;
  
   for ( long cycle = 0; cycle < 2; cycle ++ ) {

      StringFromCLSID(OBJECT_CLSID[cycle],&oleString);
      WideCharToMultiByte(CP_ACP,0,oleString,-1,szCLSID,256,0,0);

      RegOpenKeyEx(HKEY_CLASSES_ROOT,"CLSID",0,KEY_CREATE_SUB_KEY,&keyHandle);

      rc = SHDeleteKey(keyHandle,szCLSID);

      rc = SHDeleteKey(HKEY_CLASSES_ROOT,OBJECT_NAME[cycle]);

      rc = SHDeleteKey(HKEY_CLASSES_ROOT,OBJECT_NAME_V[cycle]);

   }

   return S_OK;
   }

   STDAPI DllCanUnloadNow(void) {
   return S_OK;
   }
  
