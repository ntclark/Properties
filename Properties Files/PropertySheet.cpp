// Copyright 2018 InnoVisioNate Inc. All rights reserved.
// Use of this source code is governed by a BSD-style license that can be
// found in the LICENSE file.

#include "Properties.h"

   Properties *pThis = NULL;

   void Properties::ModernPropertySheet(PROPSHEETHEADER *pSH) {

#ifndef USE_MODERN_PROPERTYSHEETS
   PropertySheet(pSH);
   return;
#endif

   memcpy(&propertySheetHeader,pSH,sizeof(PROPSHEETHEADER));

   propertySheetHeader.dwFlags |= PSH_USECALLBACK;
   propertySheetHeader.dwFlags |= PSH_PROPSHEETPAGE;
   propertySheetHeader.dwFlags |= PSH_NOAPPLYNOW;

   propertySheetHeader.pfnCallback = preparePropertySheet;

   pPropertySheetPages = new PROPSHEETPAGE[propertySheetHeader.nPages];

   memcpy(pPropertySheetPages,propertySheetHeader.ppsp,propertySheetHeader.nPages * sizeof(PROPSHEETPAGE));

   propertySheetHeader.ppsp = pPropertySheetPages;

   for ( unsigned long k = 0; k < propertySheetHeader.nPages; k++ ) {
      PROPSHEETPAGE *pPage = &pPropertySheetPages[k];
      if ( pPage -> dwFlags & PSP_DLGINDIRECT )
         continue;
      pPage -> pResource = (DLGTEMPLATE *)LoadResource(pPage -> hInstance,FindResource(pPage -> hInstance,pPage -> pszTemplate,RT_DIALOG));
      pPage -> dwFlags |= PSP_DLGINDIRECT;
   }

   pThis = this;

   propertyFrameInstanceCount++;

   hwndPropertySheet = (HWND)PropertySheet(&propertySheetHeader);

   propertyFrameInstanceCount--;

   delete [] pPropertySheetPages;

   pPropertySheetPages = NULL;

   return;
   }


   int CALLBACK Properties::preparePropertySheet(HWND hwnd,UINT uMsg,LPARAM lParam) {

   if ( ! ( PSCB_INITIALIZED == uMsg ) )
      return 0;

   SetWindowLongPtr(hwnd,GWLP_USERDATA,(LONG_PTR)(void *)pThis);

   HWND hwndTreeView = CreateWindowEx(WS_EX_CLIENTEDGE,WC_TREEVIEWA,"",
                                          WS_CHILD | WS_VISIBLE | TVS_SHOWSELALWAYS,16,16,TREEVIEW_WIDTH,32,hwnd,(HMENU)IDDI_PROPERTY_SHEET_TREEVIEW,gsProperties_hModule,0L);

   CreateWindowEx(0L,"BUTTON","Ok",WS_CHILD | WS_VISIBLE,16,16,TREEVIEW_WIDTH,32,hwnd,(HMENU)IDDI_PROPERTY_SHEET_OK,gsProperties_hModule,0L);

   CreateWindowEx(0L,"BUTTON","Cancel",WS_CHILD | WS_VISIBLE,16,16,TREEVIEW_WIDTH,32,hwnd,(HMENU)IDDI_PROPERTY_SHEET_CANCEL,gsProperties_hModule,0L);

   HFONT hGUIFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

   SendMessage(GetDlgItem(hwnd,IDDI_PROPERTY_SHEET_OK),WM_SETFONT,(WPARAM)hGUIFont,(LPARAM)TRUE);

   SendMessage(GetDlgItem(hwnd,IDDI_PROPERTY_SHEET_CANCEL),WM_SETFONT,(WPARAM)hGUIFont,(LPARAM)TRUE);

   pThis -> cxClientIdeal[pThis -> propertyFrameInstanceCount] = 0L;
   pThis -> cyClientIdeal[pThis -> propertyFrameInstanceCount] = 0L;

   HWND hwndTabControl = (HWND)SendMessage(hwnd,PSM_GETTABCONTROL,0L,0L);

   for ( UINT k = 0; k < pThis -> propertySheetHeader.nPages; k++ ) {

      TVINSERTSTRUCT tvItemInsert = {0};
      tvItemInsert.hParent = TVI_ROOT;
      tvItemInsert.hInsertAfter = TVI_LAST;
      tvItemInsert.item.mask = TVIF_TEXT | TVIF_PARAM;
      tvItemInsert.item.pszText = (LPSTR)pThis -> pPropertySheetPages[k].pszTitle;
      tvItemInsert.item.lParam = (LONG_PTR)k;

      SendMessage(hwndTreeView,TVM_INSERTITEM,0L,(LPARAM)&tvItemInsert);

      DLGTEMPLATE *pTemplate = (DLGTEMPLATE *)pThis -> propertySheetHeader.ppsp[k].pResource;
      DLGTEMPLATEEX *pTemplateEx = (DLGTEMPLATEEX *)pTemplate;

      long cxNativeClient = pTemplate -> cx;
      long cyNativeClient = pTemplate -> cy;
      long cxBaseUnits = LOWORD(GetDialogBaseUnits());
      long cyBaseUnits = HIWORD(GetDialogBaseUnits());

      WORD fontSize = 0;
      LOGFONTW logFont = {0};

      if ( 0xFFFF == pTemplateEx -> signature ) {

         cxNativeClient = pTemplateEx -> cx;
         cyNativeClient = pTemplateEx -> cy;

         WORD *pw = (WORD *)(pTemplateEx + 1);

         if ( (WORD)-1 == *pw ) 
            pw += 2;
         else
            while ( *pw++ );

         if ( (WORD)-1 == *pw )
            pw += 2;
         else
            while ( *pw++ );

         while ( *pw++ );

         fontSize = *pw;

         BYTE *pb = (BYTE *)pw;

         if ( DS_SETFONT & pTemplateEx -> style )
            pb += 3 * sizeof(WORD);

         HDC hDC = ::GetDC(NULL);

         logFont.lfHeight = -MulDiv(fontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
         logFont.lfWeight = FW_NORMAL;
         logFont.lfCharSet = DEFAULT_CHARSET;

         wcscpy(logFont.lfFaceName,(WCHAR *)pb);

         HFONT hNewFont = CreateFontIndirectW(&logFont);

         if ( hNewFont ) {
            HFONT hFontOld = (HFONT)SelectObject(hDC, hNewFont);
            TEXTMETRIC tm;
            GetTextMetrics(hDC, &tm);
            cyBaseUnits = tm.tmHeight + tm.tmExternalLeading;
            SIZE size;
            GetTextExtentPoint32(hDC,"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz",52,&size);
            cxBaseUnits = (size.cx + 26) / 52;
            SelectObject(hDC, hFontOld);
            DeleteObject(hNewFont);
         }

      }

      cxNativeClient = MulDiv(cxNativeClient, cxBaseUnits, 4);
      cyNativeClient = MulDiv(cyNativeClient, cyBaseUnits, 8);

      LONG_PTR rc = SendMessage(hwnd,PSM_SETCURSEL,(LPARAM)k,(LPARAM)0);

      HWND hwndPage = (HWND)SendMessage(hwnd,PSM_INDEXTOHWND,(LPARAM)k,0L);

      SetWindowLongPtr(hwndPage,GWL_EXSTYLE,GetWindowLongPtr(hwndPage,GWL_EXSTYLE) | WS_EX_CLIENTEDGE);

      pThis -> cxClientIdeal[pThis -> propertyFrameInstanceCount] = max(pThis -> cxClientIdeal[pThis -> propertyFrameInstanceCount],cxNativeClient);

      pThis -> cyClientIdeal[pThis -> propertyFrameInstanceCount] = max(pThis -> cyClientIdeal[pThis -> propertyFrameInstanceCount],cyNativeClient);

   }

   if ( NULL == nativePropertySheetHandler ) 
      nativePropertySheetHandler = (WNDPROC)SetWindowLongPtr(hwnd,GWLP_WNDPROC,(ULONG_PTR)propertySheetHandler);
   else
      SetWindowLongPtr(hwnd,GWLP_WNDPROC,(ULONG_PTR)propertySheetHandler);

   pThis -> cxSheetIdeal[pThis -> propertyFrameInstanceCount] = pThis -> cxClientIdeal[pThis -> propertyFrameInstanceCount] + 16 + TREEVIEW_WIDTH + 8;
   pThis -> cySheetIdeal[pThis -> propertyFrameInstanceCount] = pThis -> cyClientIdeal[pThis -> propertyFrameInstanceCount] + 16 + 16 + 16 + 8;

   pThis -> cxSheetIdeal[pThis -> propertyFrameInstanceCount] += 16;

   SetWindowPos(hwnd,HWND_TOP,0,0,pThis -> cxSheetIdeal[pThis -> propertyFrameInstanceCount],pThis -> cySheetIdeal[pThis -> propertyFrameInstanceCount] + 12,SWP_NOMOVE);

   long xIdeal = 16 + TREEVIEW_WIDTH + 8;
   long yIdeal = 16;

   LONG_PTR style = GetWindowLongPtr(hwndTabControl,GWL_STYLE);

   style &= ~TCS_MULTILINE;
   style |= TCS_SINGLELINE;

   SetWindowLongPtr(hwndTabControl,GWL_STYLE,style);

   RECT rcTabs = {0};

   SendMessage(hwndTabControl,TCM_ADJUSTRECT,(WPARAM)TRUE,(LPARAM)&rcTabs);

   SetWindowPos(hwndTabControl,HWND_TOP,xIdeal + rcTabs.left,yIdeal + rcTabs.top,
                        pThis -> cxClientIdeal[pThis -> propertyFrameInstanceCount],pThis -> cyClientIdeal[pThis -> propertyFrameInstanceCount],0L);

   ShowWindow(hwndTabControl,SW_HIDE);
   ShowWindow(GetDlgItem(hwnd,IDOK),SW_HIDE);
   ShowWindow(GetDlgItem(hwnd,IDCANCEL),SW_HIDE);

   HTREEITEM hTreeItem = (HTREEITEM)SendMessage(hwndTreeView,TVM_GETNEXTITEM,(WPARAM)TVGN_ROOT,(LPARAM)NULL);

   PostMessage(hwndTreeView,TVM_SELECTITEM,(WPARAM)TVGN_CARET,(LPARAM)hTreeItem);

   return 0L;
   }


   LRESULT EXPENTRY Properties::propertySheetHandler(HWND hwnd,UINT msg,WPARAM wParam,LPARAM lParam) {

   Properties *p = (Properties *)GetWindowLongPtr(hwnd,GWLP_USERDATA);

   switch ( msg ) {

   case WM_SIZE: {

      SetWindowPos(GetDlgItem(hwnd,IDDI_PROPERTY_SHEET_TREEVIEW),HWND_TOP,16,16,TREEVIEW_WIDTH,p -> cyClientIdeal[p -> propertyFrameInstanceCount],0L);

      for ( unsigned long k = 0; k < p -> propertySheetHeader.nPages; k++ ) {
         HWND hwndPage = (HWND)SendMessage(hwnd,PSM_INDEXTOHWND,(LPARAM)k,0L);
         SetWindowPos(hwndPage,HWND_TOP,16 + TREEVIEW_WIDTH + 8,16,
                              p -> cxClientIdeal[p -> propertyFrameInstanceCount],p -> cyClientIdeal[p -> propertyFrameInstanceCount],SWP_NOACTIVATE);
      }

      RECT rcNative,rcAdjustment = {0,0,0,0};

      AdjustWindowRect(&rcAdjustment,(DWORD)GetWindowLongPtr(hwnd,GWL_STYLE),FALSE);

      GetWindowRect(GetDlgItem(hwnd,IDOK),&rcNative);

      SetWindowPos(GetDlgItem(hwnd,IDDI_PROPERTY_SHEET_OK),HWND_TOP,
                                             LOWORD(lParam) - 2 * ((rcNative.right - rcNative.left)) - 16,
                                                HIWORD(lParam) - 2 * (rcNative.bottom - rcNative.top) - 24 + (rcAdjustment.bottom - rcAdjustment.top),
                                             rcNative.right - rcNative.left,rcNative.bottom - rcNative.top,0L);

      GetWindowRect(GetDlgItem(hwnd,IDCANCEL),&rcNative);

      SetWindowPos(GetDlgItem(hwnd,IDDI_PROPERTY_SHEET_CANCEL),HWND_TOP,
                                             LOWORD(lParam) - ((rcNative.right - rcNative.left) + 8),
                                                HIWORD(lParam) - 2 * (rcNative.bottom - rcNative.top) - 24 + (rcAdjustment.bottom - rcAdjustment.top),
                                                rcNative.right - rcNative.left,rcNative.bottom - rcNative.top,0L);

      }
      break;

   case WM_COMMAND: {
      switch ( wParam ) {
      case IDDI_PROPERTY_SHEET_OK:
         SendMessage(hwnd,PSM_PRESSBUTTON,(WPARAM)PSBTN_OK,0L);
         break;
      case IDDI_PROPERTY_SHEET_CANCEL:
         SendMessage(hwnd,PSM_PRESSBUTTON,(WPARAM)PSBTN_CANCEL,0L);
         break;
      }
      }
      break;
      
   case WM_NOTIFY: {
   
      if ( IDDI_PROPERTY_SHEET_TREEVIEW != wParam )
         break;

      NMHDR *pNotifyHeader = (NMHDR *)lParam;

      switch ( pNotifyHeader -> code ) {

      case TVN_SELCHANGEDW: {

         NMTREEVIEW *pTreeView = (NMTREEVIEW *)lParam;

         HWND hwndPage = (HWND)SendMessage(hwnd,PSM_INDEXTOHWND,pTreeView -> itemOld.lParam,0L);

         ShowWindow(hwndPage,SW_HIDE);

         hwndPage = (HWND)SendMessage(hwnd,PSM_INDEXTOHWND,pTreeView -> itemNew.lParam,0L);

         //NTC: 12-06-2022
         // I am not sure why this was here but it was messing up my efforts to change the size
         // of a property sheet dynamically.
         // Consider a way to adjust "..ClientIdeal" if this is really necessary
         //
#if 0
         SetWindowPos(hwndPage,HWND_TOP,16 + TREEVIEW_WIDTH + 8,16,
                                    p -> cxClientIdeal[p -> propertyFrameInstanceCount],p -> cyClientIdeal[p -> propertyFrameInstanceCount],0L);
#endif
         ShowWindow(hwndPage,SW_SHOW);

         }
         break;

      default:
         break;
      }

      }
      break;

   }
 
   return CallWindowProc(nativePropertySheetHandler,hwnd,msg,wParam,lParam);
   }
