========================================================================
    ACTIVE TEMPLATE LIBRARY : WatcherAddinCppAtl Project Overview
========================================================================

AppWizard has created this WatcherAddinCppAtl project for you to use as the starting point for
writing your Dynamic Link Library (DLL).

This file contains a summary of what you will find in each of the files that
make up your project.

WatcherAddinCppAtl.vcxproj
    This is the main project file for VC++ projects generated using an Application Wizard.
    It contains information about the version of Visual C++ that generated the file, and
    information about the platforms, configurations, and project features selected with the
    Application Wizard.

WatcherAddinCppAtl.vcxproj.filters
    This is the filters file for VC++ projects generated using an Application Wizard. 
    It contains information about the association between the files in your project 
    and the filters. This association is used in the IDE to show grouping of files with
    similar extensions under a specific node (for e.g. ".cpp" files are associated with the
    "Source Files" filter).

WatcherAddinCppAtl.idl
    This file contains the IDL definitions of the type library, the interfaces
    and co-classes defined in your project.
    This file will be processed by the MIDL compiler to generate:
        C++ interface definitions and GUID declarations (WatcherAddinCppAtl.h)
        GUID definitions                                (WatcherAddinCppAtl_i.c)
        A type library                                  (WatcherAddinCppAtl.tlb)
        Marshaling code                                 (WatcherAddinCppAtl_p.c and dlldata.c)

WatcherAddinCppAtl.h
    This file contains the C++ interface definitions and GUID declarations of the
    items defined in WatcherAddinCppAtl.idl. It will be regenerated by MIDL during compilation.

WatcherAddinCppAtl.cpp
    This file contains the object map and the implementation of your DLL's exports.

WatcherAddinCppAtl.rc
    This is a listing of all of the Microsoft Windows resources that the
    program uses.

WatcherAddinCppAtl.def
    This module-definition file provides the linker with information about the exports
    required by your DLL. It contains exports for:
        DllGetClassObject
        DllCanUnloadNow
        DllRegisterServer
        DllUnregisterServer
        DllInstall

/////////////////////////////////////////////////////////////////////////////
Other standard files:

StdAfx.h, StdAfx.cpp
    These files are used to build a precompiled header (PCH) file
    named WatcherAddinCppAtl.pch and a precompiled types file named StdAfx.obj.

Resource.h
    This is the standard header file that defines resource IDs.

/////////////////////////////////////////////////////////////////////////////
Proxy/stub DLL project and module definition file:

WatcherAddinCppAtlps.vcxproj
    This file is the project file for building a proxy/stub DLL if necessary.
	The IDL file in the main project must contain at least one interface and you must
	first compile the IDL file before building the proxy/stub DLL.	This process generates
	dlldata.c, WatcherAddinCppAtl_i.c and WatcherAddinCppAtl_p.c which are required
	to build the proxy/stub DLL.

WatcherAddinCppAtlps.vcxproj.filters
    This is the filters file for the proxy/stub project. It contains information about the 
    association between the files in your project and the filters. This association is 
    used in the IDE to show grouping of files with similar extensions under a specific
    node (for e.g. ".cpp" files are associated with the "Source Files" filter).

WatcherAddinCppAtlps.def
    This module definition file provides the linker with information about the exports
    required by the proxy/stub.

/////////////////////////////////////////////////////////////////////////////
