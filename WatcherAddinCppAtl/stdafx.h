// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#pragma warning(push)
#pragma warning(disable: 4298)
#pragma warning(disable: 4337)
#pragma warning(disable: 4192)

// C:\Program Files (x86)\Common Files\Designer\MSADDNDR.DLL
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	rename_namespace("AddinDesign")					\
	auto_rename auto_search raw_interfaces_only 

// C:\Program Files (x86)\Microsoft Office\Office12\MSWORD.OLB
#import "libid:00020905-0000-0000-C000-000000000046"\
	rename_namespace("Word")						\
	rename("ExitWindows","WordExitWindows")			\
	rename("FindText", "WordFindText")				\
	auto_rename auto_search raw_interfaces_only

// C:\Program Files (x86)\Microsoft Office\Office12\MSPPT.OLB
#import "libid:91493440-5A91-11CF-8700-00AA0060263B"\
	rename_namespace("PowerPoint")					\
	auto_rename auto_search raw_interfaces_only

// C:\Program Files (x86)\Microsoft Office\Office12\XL5EN32.OLB
#import "libid:00020813-0000-0000-C000-000000000046"\
	rename_namespace("Excel")					\
	auto_rename auto_search raw_interfaces_only

#pragma warning(pop)


#include "stl.h"
#include "Utils.h"
