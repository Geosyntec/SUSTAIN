# Microsoft Developer Studio Project File - Name="SUSTAIN" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=SUSTAIN - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "SUSTAIN.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "SUSTAIN.mak" CFG="SUSTAIN - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "SUSTAIN - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "SUSTAIN - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
F90=df.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "SUSTAIN - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD F90 /browser
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "_WINDLL" /D "_AFXDLL" /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /fo"Release/SUSTAIN.res" /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo /o"Release/SUSTAIN.bsc"
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 /nologo /subsystem:windows /dll /machine:I386 /out:"Release/SUSTAINOPT.dll"

!ELSEIF  "$(CFG)" == "SUSTAIN - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# SUBTRACT F90 /browser
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /Gm /GX /ZI /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "_WINDLL" /D "_AFXDLL" /FR /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /fo"Debug/SUSTAIN.res" /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo /o"Debug/SUSTAIN.bsc"
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /out:"Debug/SUSTAINOPT.dll" /pdbtype:sept

!ENDIF 

# Begin Target

# Name "SUSTAIN - Win32 Release"
# Name "SUSTAIN - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\BMPData.cpp
# End Source File
# Begin Source File

SOURCE=.\SUSTAIN.rc
# End Source File
# Begin Source File

SOURCE=.\BMPOptimizer.cpp
# End Source File
# Begin Source File

SOURCE=.\BMPOptimizerGA.cpp
# End Source File
# Begin Source File

SOURCE=.\BMPRunner.cpp
# End Source File
# Begin Source File

SOURCE=.\BMPSite.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\climate.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\controls.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\datetime.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\dynwave.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\error.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\findroot.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\flowrout.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\gage.cpp
# End Source File
# Begin Source File

SOURCE=.\Global.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\gwater.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\hash.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\iface.cpp
# End Source File
# Begin Source File

SOURCE=.\Individual.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\infil.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\inflow.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\input.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\keywords.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\kinwave.cpp
# End Source File
# Begin Source File

SOURCE=.\LandUse.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\landuseswmm.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\link.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\massbal.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\mathexpr.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\mempool.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\node.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\odesolve.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\output.cpp
# End Source File
# Begin Source File

SOURCE=.\Population.cpp
# End Source File
# Begin Source File

SOURCE=.\ProgressWnd.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\project.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\qualrout.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\rain.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\rdii.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\report.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\routing.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\runoff.cpp
# End Source File
# Begin Source File

SOURCE=.\Sediment.cpp
# End Source File
# Begin Source File

SOURCE=.\SiteLandUse.cpp
# End Source File
# Begin Source File

SOURCE=.\SitePointSource.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\snow.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\stats.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\StringToken.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\subcatch.cpp
# End Source File
# Begin Source File

SOURCE=.\SUSTAIN.cpp
# End Source File
# Begin Source File

SOURCE=.\SUSTAIN.def
# End Source File
# Begin Source File

SOURCE=.\SWMM5\swmm5.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\table.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\toposort.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\transect.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\treatmnt.cpp
# End Source File
# Begin Source File

SOURCE=.\SWMM5\xsect.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\BMPData.h
# End Source File
# Begin Source File

SOURCE=.\BMPOptimizer.h
# End Source File
# Begin Source File

SOURCE=.\BMPOptimizerGA.h
# End Source File
# Begin Source File

SOURCE=.\BMPRunner.h
# End Source File
# Begin Source File

SOURCE=.\BMPSite.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\consts.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\datetime.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\enums.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\error.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\findroot.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\funcs.h
# End Source File
# Begin Source File

SOURCE=.\Global.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\globals.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\hash.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\headers.h
# End Source File
# Begin Source File

SOURCE=.\Individual.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\keywords.h
# End Source File
# Begin Source File

SOURCE=.\LandUse.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\macros.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\mathexpr.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\mempool.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\objects.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\odesolve.h
# End Source File
# Begin Source File

SOURCE=.\Population.h
# End Source File
# Begin Source File

SOURCE=.\ProgressWnd.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\Sediment.h
# End Source File
# Begin Source File

SOURCE=.\SiteLandUse.h
# End Source File
# Begin Source File

SOURCE=.\SitePointSource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\StringToken.h
# End Source File
# Begin Source File

SOURCE=.\SUSTAIN.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\swmm5.h
# End Source File
# Begin Source File

SOURCE=.\SWMM5\text.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\SUSTAIN.rc2
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
