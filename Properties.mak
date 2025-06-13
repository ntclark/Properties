# Microsoft Developer Studio Generated NMAKE File, Based on Properties.dsp
!IF "$(CFG)" == ""
CFG=Properties - Win32 Debug
!MESSAGE No configuration specified. Defaulting to Properties - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "Properties - Win32 Release" && "$(CFG)" != "Properties - Win32 Debug" && "$(CFG)" != "Properties - Win32 Evaluation Release"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Properties.mak" CFG="Properties - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Properties - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Properties - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "Properties - Win32 Evaluation Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

!IF  "$(CFG)" == "Properties - Win32 Release"

OUTDIR=.
INTDIR=.
# Begin Custom Macros
OutDir=.
# End Custom Macros

ALL : "..\lib\Properties.ocx" "$(OUTDIR)\Properties.tlb" ".\regsvr32.trg"


CLEAN :
	-@erase "$(INTDIR)\COMImplementation.obj"
	-@erase "$(INTDIR)\COMUtils.obj"
	-@erase "$(INTDIR)\DllMain.obj"
	-@erase "$(INTDIR)\IConnectionPoint.obj"
	-@erase "$(INTDIR)\IConnectionPointContainer.obj"
	-@erase "$(INTDIR)\IEnumConnectionPoints.obj"
	-@erase "$(INTDIR)\IEnumConnections.obj"
	-@erase "$(INTDIR)\IOleControl.obj"
	-@erase "$(INTDIR)\IOleInPlaceObject.obj"
	-@erase "$(INTDIR)\IOleObject.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag2.obj"
	-@erase "$(INTDIR)\IPersistStorage.obj"
	-@erase "$(INTDIR)\IPersistStream.obj"
	-@erase "$(INTDIR)\IPersistStreamInit.obj"
	-@erase "$(INTDIR)\IProperties.obj"
	-@erase "$(INTDIR)\IProperties_Files.obj"
	-@erase "$(INTDIR)\IProperties_propertyPages.obj"
	-@erase "$(INTDIR)\IProperty.obj"
	-@erase "$(INTDIR)\IProperty_arrays.obj"
	-@erase "$(INTDIR)\IProperty_value.obj"
	-@erase "$(INTDIR)\IProperty_variantValue.obj"
	-@erase "$(INTDIR)\IProperty_windows.obj"
	-@erase "$(INTDIR)\IPropertyPageClient.obj"
	-@erase "$(INTDIR)\IQuickActivate.obj"
	-@erase "$(INTDIR)\IRunnableObject.obj"
	-@erase "$(INTDIR)\IUnknown.obj"
	-@erase "$(INTDIR)\IViewObject.obj"
	-@erase "$(INTDIR)\mathUtils.obj"
	-@erase "$(INTDIR)\NonDelegatingIUnknown.obj"
	-@erase "$(INTDIR)\ObjectFactory.obj"
	-@erase "$(INTDIR)\Properties.obj"
	-@erase "$(INTDIR)\Properties.res"
	-@erase "$(INTDIR)\Properties.tlb"
	-@erase "$(INTDIR)\Property.obj"
	-@erase "$(INTDIR)\Property_Controls.obj"
	-@erase "$(INTDIR)\Property_Variants.obj"
	-@erase "$(INTDIR)\registryUtils.obj"
	-@erase "$(INTDIR)\Utils.obj"
	-@erase "$(INTDIR)\VariantUtils.obj"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(OUTDIR)\Properties.exp"
	-@erase "..\lib\Properties.ocx"
	-@erase ".\regsvr32.trg"

LIB32=link.exe -lib
F90=df.exe
F90_PROJ=
F90_OBJS=

.SUFFIXES: .fpp

.for{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f90{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.fpp{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /GX /O2 /I "..\\" /I "..\h" /D "EMBEDDED_PROPERTIES" /D "_WIN32_DCOM" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /FD /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/D "NDEBUG" /char ascii7 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Properties.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib /nologo /dll /incremental:no /pdb:"$(OUTDIR)\Properties.pdb" /machine:I386 /def:".\Properties.def" /out:"..\lib\Properties.ocx" /implib:"$(OUTDIR)\Properties.lib" 
DEF_FILE= \
	".\Properties.def"
LINK32_OBJS= \
	"$(INTDIR)\IOleControl.obj" \
	"$(INTDIR)\IOleInPlaceObject.obj" \
	"$(INTDIR)\IQuickActivate.obj" \
	"$(INTDIR)\IRunnableObject.obj" \
	"$(INTDIR)\IViewObject.obj" \
	"$(INTDIR)\IOleObject.obj" \
	"$(INTDIR)\IPersistPropertyBag.obj" \
	"$(INTDIR)\IPersistPropertyBag2.obj" \
	"$(INTDIR)\IPersistStorage.obj" \
	"$(INTDIR)\IPersistStream.obj" \
	"$(INTDIR)\IPersistStreamInit.obj" \
	"$(INTDIR)\IConnectionPoint.obj" \
	"$(INTDIR)\IConnectionPointContainer.obj" \
	"$(INTDIR)\IEnumConnectionPoints.obj" \
	"$(INTDIR)\IEnumConnections.obj" \
	"$(INTDIR)\COMImplementation.obj" \
	"$(INTDIR)\DllMain.obj" \
	"$(INTDIR)\IPropertyPageClient.obj" \
	"$(INTDIR)\IUnknown.obj" \
	"$(INTDIR)\NonDelegatingIUnknown.obj" \
	"$(INTDIR)\IProperty.obj" \
	"$(INTDIR)\IProperty_arrays.obj" \
	"$(INTDIR)\IProperty_value.obj" \
	"$(INTDIR)\IProperty_variantValue.obj" \
	"$(INTDIR)\IProperty_windows.obj" \
	"$(INTDIR)\Property.obj" \
	"$(INTDIR)\Property_Controls.obj" \
	"$(INTDIR)\Property_Variants.obj" \
	"$(INTDIR)\IProperties.obj" \
	"$(INTDIR)\IProperties_Files.obj" \
	"$(INTDIR)\IProperties_propertyPages.obj" \
	"$(INTDIR)\Properties.obj" \
	"$(INTDIR)\COMUtils.obj" \
	"$(INTDIR)\mathUtils.obj" \
	"$(INTDIR)\registryUtils.obj" \
	"$(INTDIR)\Utils.obj" \
	"$(INTDIR)\VariantUtils.obj" \
	"$(INTDIR)\ObjectFactory.obj" \
	"$(INTDIR)\Properties.res"

"..\lib\Properties.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

OutDir=.
TargetPath=\prj\lib\Properties.ocx
InputPath=\prj\lib\Properties.ocx
SOURCE="$(InputPath)"

"$(OUTDIR)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	<<tempfile.bat 
	@echo off 
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
<< 
	
SOURCE="$(InputPath)"
DS_POSTBUILD_DEP=$(INTDIR)\postbld.dep

ALL : $(DS_POSTBUILD_DEP)

# Begin Custom Macros
OutDir=.
# End Custom Macros

$(DS_POSTBUILD_DEP) : "..\lib\Properties.ocx" "$(OUTDIR)\Properties.tlb" ".\regsvr32.trg"
   del ..\lib\Properties.lib
	lib /NOLOGO /OUT:..\lib\Properties.lib *.obj
	echo Helper for Post-build step > "$(DS_POSTBUILD_DEP)"

!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"

OUTDIR=.
INTDIR=.
# Begin Custom Macros
OutDir=.
# End Custom Macros

ALL : ".\Properties" "..\Properties.ocx" "..\Properties.tlb" "..\Properties_i.h" "..\Properties_i.c" "$(OUTDIR)\Properties.bsc" ".\regsvr32.trg"


CLEAN :
	-@erase "$(INTDIR)\COMImplementation.obj"
	-@erase "$(INTDIR)\COMImplementation.sbr"
	-@erase "$(INTDIR)\COMUtils.obj"
	-@erase "$(INTDIR)\COMUtils.sbr"
	-@erase "$(INTDIR)\DllMain.obj"
	-@erase "$(INTDIR)\DllMain.sbr"
	-@erase "$(INTDIR)\IConnectionPoint.obj"
	-@erase "$(INTDIR)\IConnectionPoint.sbr"
	-@erase "$(INTDIR)\IConnectionPointContainer.obj"
	-@erase "$(INTDIR)\IConnectionPointContainer.sbr"
	-@erase "$(INTDIR)\IEnumConnectionPoints.obj"
	-@erase "$(INTDIR)\IEnumConnectionPoints.sbr"
	-@erase "$(INTDIR)\IEnumConnections.obj"
	-@erase "$(INTDIR)\IEnumConnections.sbr"
	-@erase "$(INTDIR)\IOleControl.obj"
	-@erase "$(INTDIR)\IOleControl.sbr"
	-@erase "$(INTDIR)\IOleInPlaceObject.obj"
	-@erase "$(INTDIR)\IOleInPlaceObject.sbr"
	-@erase "$(INTDIR)\IOleObject.obj"
	-@erase "$(INTDIR)\IOleObject.sbr"
	-@erase "$(INTDIR)\IPersistPropertyBag.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag.sbr"
	-@erase "$(INTDIR)\IPersistPropertyBag2.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag2.sbr"
	-@erase "$(INTDIR)\IPersistStorage.obj"
	-@erase "$(INTDIR)\IPersistStorage.sbr"
	-@erase "$(INTDIR)\IPersistStream.obj"
	-@erase "$(INTDIR)\IPersistStream.sbr"
	-@erase "$(INTDIR)\IPersistStreamInit.obj"
	-@erase "$(INTDIR)\IPersistStreamInit.sbr"
	-@erase "$(INTDIR)\IProperties.obj"
	-@erase "$(INTDIR)\IProperties.sbr"
	-@erase "$(INTDIR)\IProperties_Files.obj"
	-@erase "$(INTDIR)\IProperties_Files.sbr"
	-@erase "$(INTDIR)\IProperties_propertyPages.obj"
	-@erase "$(INTDIR)\IProperties_propertyPages.sbr"
	-@erase "$(INTDIR)\IProperty.obj"
	-@erase "$(INTDIR)\IProperty.sbr"
	-@erase "$(INTDIR)\IProperty_arrays.obj"
	-@erase "$(INTDIR)\IProperty_arrays.sbr"
	-@erase "$(INTDIR)\IProperty_value.obj"
	-@erase "$(INTDIR)\IProperty_value.sbr"
	-@erase "$(INTDIR)\IProperty_variantValue.obj"
	-@erase "$(INTDIR)\IProperty_variantValue.sbr"
	-@erase "$(INTDIR)\IProperty_windows.obj"
	-@erase "$(INTDIR)\IProperty_windows.sbr"
	-@erase "$(INTDIR)\IPropertyPageClient.obj"
	-@erase "$(INTDIR)\IPropertyPageClient.sbr"
	-@erase "$(INTDIR)\IQuickActivate.obj"
	-@erase "$(INTDIR)\IQuickActivate.sbr"
	-@erase "$(INTDIR)\IRunnableObject.obj"
	-@erase "$(INTDIR)\IRunnableObject.sbr"
	-@erase "$(INTDIR)\IUnknown.obj"
	-@erase "$(INTDIR)\IUnknown.sbr"
	-@erase "$(INTDIR)\IViewObject.obj"
	-@erase "$(INTDIR)\IViewObject.sbr"
	-@erase "$(INTDIR)\mathUtils.obj"
	-@erase "$(INTDIR)\mathUtils.sbr"
	-@erase "$(INTDIR)\NonDelegatingIUnknown.obj"
	-@erase "$(INTDIR)\NonDelegatingIUnknown.sbr"
	-@erase "$(INTDIR)\ObjectFactory.obj"
	-@erase "$(INTDIR)\ObjectFactory.sbr"
	-@erase "$(INTDIR)\Properties.obj"
	-@erase "$(INTDIR)\Properties.res"
	-@erase "$(INTDIR)\Properties.sbr"
	-@erase "$(INTDIR)\Property.obj"
	-@erase "$(INTDIR)\Property.sbr"
	-@erase "$(INTDIR)\Property_Controls.obj"
	-@erase "$(INTDIR)\Property_Controls.sbr"
	-@erase "$(INTDIR)\Property_Variants.obj"
	-@erase "$(INTDIR)\Property_Variants.sbr"
	-@erase "$(INTDIR)\registryUtils.obj"
	-@erase "$(INTDIR)\registryUtils.sbr"
	-@erase "$(INTDIR)\Utils.obj"
	-@erase "$(INTDIR)\Utils.sbr"
	-@erase "$(INTDIR)\VariantUtils.obj"
	-@erase "$(INTDIR)\VariantUtils.sbr"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(INTDIR)\vc60.pdb"
	-@erase "$(OUTDIR)\Properties.bsc"
	-@erase "$(OUTDIR)\Properties.exp"
	-@erase "..\Properties.ocx"
	-@erase "..\Properties.pdb"
	-@erase "..\Properties.tlb"
	-@erase "..\Properties_i.c"
	-@erase "..\Properties_i.h"
	-@erase ".\regsvr32.trg"
	-@erase "Properties"

LIB32=link.exe -lib
F90=df.exe
F90_PROJ=/browser 
F90_OBJS=

.SUFFIXES: .fpp

.for{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f90{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.fpp{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

CPP=cl.exe
CPP_PROJ=/nologo /MD /W3 /GR /GX /Zi /Od /I "..\\" /I "..\h" /D "N_EMBEDDED_PROPERTIES" /D "_WIN32_DCOM" /D "N_PROPERTIES_SAMPLE" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D _WIN32_WINNT=0x501 /Fr /J /FD /GZ /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/D "_DEBUG" /char ascii7 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "_DEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Properties.bsc" 
BSC32_SBRS= \
	"$(INTDIR)\IOleControl.sbr" \
	"$(INTDIR)\IOleInPlaceObject.sbr" \
	"$(INTDIR)\IQuickActivate.sbr" \
	"$(INTDIR)\IRunnableObject.sbr" \
	"$(INTDIR)\IViewObject.sbr" \
	"$(INTDIR)\IOleObject.sbr" \
	"$(INTDIR)\IPersistPropertyBag.sbr" \
	"$(INTDIR)\IPersistPropertyBag2.sbr" \
	"$(INTDIR)\IPersistStorage.sbr" \
	"$(INTDIR)\IPersistStream.sbr" \
	"$(INTDIR)\IPersistStreamInit.sbr" \
	"$(INTDIR)\IConnectionPoint.sbr" \
	"$(INTDIR)\IConnectionPointContainer.sbr" \
	"$(INTDIR)\IEnumConnectionPoints.sbr" \
	"$(INTDIR)\IEnumConnections.sbr" \
	"$(INTDIR)\COMImplementation.sbr" \
	"$(INTDIR)\DllMain.sbr" \
	"$(INTDIR)\IPropertyPageClient.sbr" \
	"$(INTDIR)\IUnknown.sbr" \
	"$(INTDIR)\NonDelegatingIUnknown.sbr" \
	"$(INTDIR)\IProperty.sbr" \
	"$(INTDIR)\IProperty_arrays.sbr" \
	"$(INTDIR)\IProperty_value.sbr" \
	"$(INTDIR)\IProperty_variantValue.sbr" \
	"$(INTDIR)\IProperty_windows.sbr" \
	"$(INTDIR)\Property.sbr" \
	"$(INTDIR)\Property_Controls.sbr" \
	"$(INTDIR)\Property_Variants.sbr" \
	"$(INTDIR)\IProperties.sbr" \
	"$(INTDIR)\IProperties_Files.sbr" \
	"$(INTDIR)\IProperties_propertyPages.sbr" \
	"$(INTDIR)\Properties.sbr" \
	"$(INTDIR)\COMUtils.sbr" \
	"$(INTDIR)\mathUtils.sbr" \
	"$(INTDIR)\registryUtils.sbr" \
	"$(INTDIR)\Utils.sbr" \
	"$(INTDIR)\VariantUtils.sbr" \
	"$(INTDIR)\ObjectFactory.sbr"

"$(OUTDIR)\Properties.bsc" : "$(OUTDIR)" $(BSC32_SBRS)
    $(BSC32) @<<
  $(BSC32_FLAGS) $(BSC32_SBRS)
<<

LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib /nologo /dll /incremental:no /pdb:"..\Properties.pdb" /debug /machine:I386 /def:".\Properties.def" /out:"..\Properties.ocx" /implib:"$(OUTDIR)\Properties.lib" /libpath:"..\\" /libpath:"..\lib" 
LINK32_OBJS= \
	"$(INTDIR)\IOleControl.obj" \
	"$(INTDIR)\IOleInPlaceObject.obj" \
	"$(INTDIR)\IQuickActivate.obj" \
	"$(INTDIR)\IRunnableObject.obj" \
	"$(INTDIR)\IViewObject.obj" \
	"$(INTDIR)\IOleObject.obj" \
	"$(INTDIR)\IPersistPropertyBag.obj" \
	"$(INTDIR)\IPersistPropertyBag2.obj" \
	"$(INTDIR)\IPersistStorage.obj" \
	"$(INTDIR)\IPersistStream.obj" \
	"$(INTDIR)\IPersistStreamInit.obj" \
	"$(INTDIR)\IConnectionPoint.obj" \
	"$(INTDIR)\IConnectionPointContainer.obj" \
	"$(INTDIR)\IEnumConnectionPoints.obj" \
	"$(INTDIR)\IEnumConnections.obj" \
	"$(INTDIR)\COMImplementation.obj" \
	"$(INTDIR)\DllMain.obj" \
	"$(INTDIR)\IPropertyPageClient.obj" \
	"$(INTDIR)\IUnknown.obj" \
	"$(INTDIR)\NonDelegatingIUnknown.obj" \
	"$(INTDIR)\IProperty.obj" \
	"$(INTDIR)\IProperty_arrays.obj" \
	"$(INTDIR)\IProperty_value.obj" \
	"$(INTDIR)\IProperty_variantValue.obj" \
	"$(INTDIR)\IProperty_windows.obj" \
	"$(INTDIR)\Property.obj" \
	"$(INTDIR)\Property_Controls.obj" \
	"$(INTDIR)\Property_Variants.obj" \
	"$(INTDIR)\IProperties.obj" \
	"$(INTDIR)\IProperties_Files.obj" \
	"$(INTDIR)\IProperties_propertyPages.obj" \
	"$(INTDIR)\Properties.obj" \
	"$(INTDIR)\COMUtils.obj" \
	"$(INTDIR)\mathUtils.obj" \
	"$(INTDIR)\registryUtils.obj" \
	"$(INTDIR)\Utils.obj" \
	"$(INTDIR)\VariantUtils.obj" \
	"$(INTDIR)\ObjectFactory.obj" \
	"$(INTDIR)\Properties.res"

"..\Properties.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

OutDir=.
TargetPath=\prj\Properties.ocx
InputPath=\prj\Properties.ocx
SOURCE="$(InputPath)"

"$(OUTDIR)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	<<tempfile.bat 
	@echo off 
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
<< 
	
SOURCE="$(InputPath)"
PostBuild_Desc=Delete sample dependencies
DS_POSTBUILD_DEP=$(INTDIR)\postbld.dep

ALL : $(DS_POSTBUILD_DEP)

# Begin Custom Macros
OutDir=.
# End Custom Macros

$(DS_POSTBUILD_DEP) : ".\Properties" "..\Properties.ocx" "..\Properties.tlb" "..\Properties_i.h" "..\Properties_i.c" "$(OUTDIR)\Properties.bsc" ".\regsvr32.trg"
   del ..\lib\Properties.lib
	lib /NOLOGO /OUT:..\lib\Properties.lib *.obj
	echo Helper for Post-build step > "$(DS_POSTBUILD_DEP)"

!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"

OUTDIR=.
INTDIR=.
# Begin Custom Macros
OutDir=.
# End Custom Macros

ALL : "..\lib\Properties.ocx" "$(OUTDIR)\Properties.tlb" ".\regsvr32.trg"


CLEAN :
	-@erase "$(INTDIR)\COMImplementation.obj"
	-@erase "$(INTDIR)\COMUtils.obj"
	-@erase "$(INTDIR)\DllMain.obj"
	-@erase "$(INTDIR)\IConnectionPoint.obj"
	-@erase "$(INTDIR)\IConnectionPointContainer.obj"
	-@erase "$(INTDIR)\IEnumConnectionPoints.obj"
	-@erase "$(INTDIR)\IEnumConnections.obj"
	-@erase "$(INTDIR)\IOleControl.obj"
	-@erase "$(INTDIR)\IOleInPlaceObject.obj"
	-@erase "$(INTDIR)\IOleObject.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag.obj"
	-@erase "$(INTDIR)\IPersistPropertyBag2.obj"
	-@erase "$(INTDIR)\IPersistStorage.obj"
	-@erase "$(INTDIR)\IPersistStream.obj"
	-@erase "$(INTDIR)\IPersistStreamInit.obj"
	-@erase "$(INTDIR)\IProperties.obj"
	-@erase "$(INTDIR)\IProperties_Files.obj"
	-@erase "$(INTDIR)\IProperties_propertyPages.obj"
	-@erase "$(INTDIR)\IProperty.obj"
	-@erase "$(INTDIR)\IProperty_arrays.obj"
	-@erase "$(INTDIR)\IProperty_value.obj"
	-@erase "$(INTDIR)\IProperty_variantValue.obj"
	-@erase "$(INTDIR)\IProperty_windows.obj"
	-@erase "$(INTDIR)\IPropertyPageClient.obj"
	-@erase "$(INTDIR)\IQuickActivate.obj"
	-@erase "$(INTDIR)\IRunnableObject.obj"
	-@erase "$(INTDIR)\IUnknown.obj"
	-@erase "$(INTDIR)\IViewObject.obj"
	-@erase "$(INTDIR)\mathUtils.obj"
	-@erase "$(INTDIR)\NonDelegatingIUnknown.obj"
	-@erase "$(INTDIR)\ObjectFactory.obj"
	-@erase "$(INTDIR)\Properties.obj"
	-@erase "$(INTDIR)\Properties.res"
	-@erase "$(INTDIR)\Properties.tlb"
	-@erase "$(INTDIR)\Property.obj"
	-@erase "$(INTDIR)\Property_Controls.obj"
	-@erase "$(INTDIR)\Property_Variants.obj"
	-@erase "$(INTDIR)\registryUtils.obj"
	-@erase "$(INTDIR)\Utils.obj"
	-@erase "$(INTDIR)\VariantUtils.obj"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(OUTDIR)\Properties.exp"
	-@erase "..\lib\Properties.ocx"
	-@erase ".\regsvr32.trg"

LIB32=link.exe -lib
F90=df.exe
F90_PROJ=
F90_OBJS=

.SUFFIXES: .fpp

.for{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.f90{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

.fpp{$(F90_OBJS)}.obj:
   $(F90) $(F90_PROJ) $<  

CPP=cl.exe
CPP_PROJ=/nologo /MT /W3 /GX /O2 /I "..\\" /I "..\h" /D "PROPERTIES_SAMPLE" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /FD /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/D "NDEBUG" /char ascii7 
RSC=rc.exe
RSC_PROJ=/l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\Properties.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib /nologo /dll /incremental:no /pdb:"$(OUTDIR)\Properties.pdb" /machine:I386 /def:".\Properties.def" /out:"..\lib\Properties.ocx" /implib:"$(OUTDIR)\Properties.lib" 
DEF_FILE= \
	".\Properties.def"
LINK32_OBJS= \
	"$(INTDIR)\IOleControl.obj" \
	"$(INTDIR)\IOleInPlaceObject.obj" \
	"$(INTDIR)\IQuickActivate.obj" \
	"$(INTDIR)\IRunnableObject.obj" \
	"$(INTDIR)\IViewObject.obj" \
	"$(INTDIR)\IOleObject.obj" \
	"$(INTDIR)\IPersistPropertyBag.obj" \
	"$(INTDIR)\IPersistPropertyBag2.obj" \
	"$(INTDIR)\IPersistStorage.obj" \
	"$(INTDIR)\IPersistStream.obj" \
	"$(INTDIR)\IPersistStreamInit.obj" \
	"$(INTDIR)\IConnectionPoint.obj" \
	"$(INTDIR)\IConnectionPointContainer.obj" \
	"$(INTDIR)\IEnumConnectionPoints.obj" \
	"$(INTDIR)\IEnumConnections.obj" \
	"$(INTDIR)\COMImplementation.obj" \
	"$(INTDIR)\DllMain.obj" \
	"$(INTDIR)\IPropertyPageClient.obj" \
	"$(INTDIR)\IUnknown.obj" \
	"$(INTDIR)\NonDelegatingIUnknown.obj" \
	"$(INTDIR)\IProperty.obj" \
	"$(INTDIR)\IProperty_arrays.obj" \
	"$(INTDIR)\IProperty_value.obj" \
	"$(INTDIR)\IProperty_variantValue.obj" \
	"$(INTDIR)\IProperty_windows.obj" \
	"$(INTDIR)\Property.obj" \
	"$(INTDIR)\Property_Controls.obj" \
	"$(INTDIR)\Property_Variants.obj" \
	"$(INTDIR)\IProperties.obj" \
	"$(INTDIR)\IProperties_Files.obj" \
	"$(INTDIR)\IProperties_propertyPages.obj" \
	"$(INTDIR)\Properties.obj" \
	"$(INTDIR)\COMUtils.obj" \
	"$(INTDIR)\mathUtils.obj" \
	"$(INTDIR)\registryUtils.obj" \
	"$(INTDIR)\Utils.obj" \
	"$(INTDIR)\VariantUtils.obj" \
	"$(INTDIR)\ObjectFactory.obj" \
	"$(INTDIR)\Properties.res"

"..\lib\Properties.ocx" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

OutDir=.
TargetPath=\prj\lib\Properties.ocx
InputPath=\prj\lib\Properties.ocx
SOURCE="$(InputPath)"

"$(OUTDIR)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	<<tempfile.bat 
	@echo off 
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
<< 
	

!ENDIF 


!IF "$(NO_EXTERNAL_DEPS)" != "1"
!IF EXISTS("Properties.dep")
!INCLUDE "Properties.dep"
!ELSE 
!MESSAGE Warning: cannot find "Properties.dep"
!ENDIF 
!ENDIF 


!IF "$(CFG)" == "Properties - Win32 Release" || "$(CFG)" == "Properties - Win32 Debug" || "$(CFG)" == "Properties - Win32 Evaluation Release"
SOURCE=.\IOleControl.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IOleControl.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IOleControl.obj"	"$(INTDIR)\IOleControl.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IOleControl.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IOleInPlaceObject.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IOleInPlaceObject.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IOleInPlaceObject.obj"	"$(INTDIR)\IOleInPlaceObject.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IOleInPlaceObject.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IQuickActivate.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IQuickActivate.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IQuickActivate.obj"	"$(INTDIR)\IQuickActivate.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IQuickActivate.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IRunnableObject.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IRunnableObject.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IRunnableObject.obj"	"$(INTDIR)\IRunnableObject.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IRunnableObject.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IViewObject.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IViewObject.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IViewObject.obj"	"$(INTDIR)\IViewObject.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IViewObject.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IOleObject.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IOleObject.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IOleObject.obj"	"$(INTDIR)\IOleObject.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IOleObject.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPersistPropertyBag.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPersistPropertyBag.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPersistPropertyBag.obj"	"$(INTDIR)\IPersistPropertyBag.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPersistPropertyBag.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPersistPropertyBag2.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPersistPropertyBag2.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPersistPropertyBag2.obj"	"$(INTDIR)\IPersistPropertyBag2.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPersistPropertyBag2.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPersistStorage.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPersistStorage.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPersistStorage.obj"	"$(INTDIR)\IPersistStorage.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPersistStorage.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPersistStream.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPersistStream.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPersistStream.obj"	"$(INTDIR)\IPersistStream.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPersistStream.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPersistStreamInit.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPersistStreamInit.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPersistStreamInit.obj"	"$(INTDIR)\IPersistStreamInit.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPersistStreamInit.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IConnectionPoint.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IConnectionPoint.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IConnectionPoint.obj"	"$(INTDIR)\IConnectionPoint.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IConnectionPoint.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IConnectionPointContainer.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IConnectionPointContainer.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IConnectionPointContainer.obj"	"$(INTDIR)\IConnectionPointContainer.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IConnectionPointContainer.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IEnumConnectionPoints.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IEnumConnectionPoints.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IEnumConnectionPoints.obj"	"$(INTDIR)\IEnumConnectionPoints.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IEnumConnectionPoints.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IEnumConnections.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IEnumConnections.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IEnumConnections.obj"	"$(INTDIR)\IEnumConnections.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IEnumConnections.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\COMImplementation.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\COMImplementation.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\COMImplementation.obj"	"$(INTDIR)\COMImplementation.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\COMImplementation.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\DllMain.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\DllMain.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\DllMain.obj"	"$(INTDIR)\DllMain.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\DllMain.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IPropertyPageClient.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IPropertyPageClient.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IPropertyPageClient.obj"	"$(INTDIR)\IPropertyPageClient.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IPropertyPageClient.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IUnknown.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IUnknown.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IUnknown.obj"	"$(INTDIR)\IUnknown.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IUnknown.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\NonDelegatingIUnknown.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\NonDelegatingIUnknown.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\NonDelegatingIUnknown.obj"	"$(INTDIR)\NonDelegatingIUnknown.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\NonDelegatingIUnknown.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\Properties.odl

!IF  "$(CFG)" == "Properties - Win32 Release"

MTL_SWITCHES=/D "NDEBUG" /tlb "$(OUTDIR)\Properties.tlb" /char ascii7 

"$(OUTDIR)\Properties.tlb" : $(SOURCE) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"

MTL_SWITCHES=/D "_DEBUG" /tlb "..\Properties.tlb" /h "..\Properties_i.h" /iid "..\Properties_i.c" /char ascii7 

"..\Properties.tlb"	"..\Properties_i.h"	"..\Properties_i.c" : $(SOURCE) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"

MTL_SWITCHES=/D "NDEBUG" /tlb "$(OUTDIR)\Properties.tlb" /char ascii7 

"$(OUTDIR)\Properties.tlb" : $(SOURCE) "$(OUTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\IProperty.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperty.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperty.obj"	"$(INTDIR)\IProperty.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperty.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperty_arrays.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperty_arrays.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperty_arrays.obj"	"$(INTDIR)\IProperty_arrays.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperty_arrays.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperty_value.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperty_value.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperty_value.obj"	"$(INTDIR)\IProperty_value.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperty_value.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperty_variantValue.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperty_variantValue.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperty_variantValue.obj"	"$(INTDIR)\IProperty_variantValue.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperty_variantValue.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperty_windows.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperty_windows.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperty_windows.obj"	"$(INTDIR)\IProperty_windows.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperty_windows.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\Property.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Property.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Property.obj"	"$(INTDIR)\Property.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Property.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\Property_Controls.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Property_Controls.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Property_Controls.obj"	"$(INTDIR)\Property_Controls.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Property_Controls.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\Property_Variants.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Property_Variants.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Property_Variants.obj"	"$(INTDIR)\Property_Variants.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Property_Variants.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperties.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperties.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperties.obj"	"$(INTDIR)\IProperties.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperties.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperties_Files.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperties_Files.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperties_Files.obj"	"$(INTDIR)\IProperties_Files.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperties_Files.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\IProperties_propertyPages.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\IProperties_propertyPages.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\IProperties_propertyPages.obj"	"$(INTDIR)\IProperties_propertyPages.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\IProperties_propertyPages.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\Properties.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Properties.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Properties.obj"	"$(INTDIR)\Properties.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Properties.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=..\Utils\COMUtils.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\COMUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\COMUtils.obj"	"$(INTDIR)\COMUtils.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\COMUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=..\Utils\mathUtils.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\mathUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\mathUtils.obj"	"$(INTDIR)\mathUtils.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\mathUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=..\Utils\registryUtils.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\registryUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\registryUtils.obj"	"$(INTDIR)\registryUtils.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\registryUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=..\Utils\Utils.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Utils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Utils.obj"	"$(INTDIR)\Utils.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Utils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=..\Utils\VariantUtils.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\VariantUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\VariantUtils.obj"	"$(INTDIR)\VariantUtils.sbr" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\VariantUtils.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) $(CPP_PROJ) $(SOURCE)


!ENDIF 

SOURCE=.\ObjectFactory.cpp

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\ObjectFactory.obj" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\ObjectFactory.obj"	"$(INTDIR)\ObjectFactory.sbr" : $(SOURCE) "$(INTDIR)"


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\ObjectFactory.obj" : $(SOURCE) "$(INTDIR)"


!ENDIF 

SOURCE=.\builddate

!IF  "$(CFG)" == "Properties - Win32 Release"

!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"

TargetName=Properties
InputPath=.\builddate
USERDEP__BUILD="Properties.rc"	

"$(INTDIR)\Properties" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)" $(USERDEP__BUILD)
	<<tempfile.bat 
	@echo off 
	..\GetDate > $(TargetName)
<< 
	

!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"

!ENDIF 

SOURCE=.\Properties.rc

!IF  "$(CFG)" == "Properties - Win32 Release"


"$(INTDIR)\Properties.res" : $(SOURCE) "$(INTDIR)"
	$(RSC) /l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "NDEBUG" $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Debug"


"$(INTDIR)\Properties.res" : $(SOURCE) "$(INTDIR)"
	$(RSC) /l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "_DEBUG" $(SOURCE)


!ELSEIF  "$(CFG)" == "Properties - Win32 Evaluation Release"


"$(INTDIR)\Properties.res" : $(SOURCE) "$(INTDIR)"
	$(RSC) /l 0x409 /fo"$(INTDIR)\Properties.res" /i "..\\" /i "..\h" /d "NDEBUG" $(SOURCE)


!ENDIF 


!ENDIF 

