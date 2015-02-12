// ExcelEvents.h : Declaration of the CExcelEvents

#pragma once
#include "resource.h"       // main symbols
#include "WatcherAddinCppAtl_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

using WorkbookBeforeCloseHandler = std::function < void(Excel::_Workbook*) > ;
using WorkbookActivateHandler = std::function < void(Excel::_Workbook*) > ;
using WorkbookDeactivateHandler = std::function < void(Excel::_Workbook*) > ;

class ATL_NO_VTABLE ExcelEvents :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<ExcelEvents, &CLSID_ExcelEvents>,
	public IDispatchImpl<IExcelEvents, &IID_IExcelEvents, &LIBID_WatcherAddinCppAtlLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispEventSimpleImpl<0, ExcelEvents, &__uuidof(Excel::AppEvents) >
{
	NO_COPY_ASSIGN(ExcelEvents);

	static _ATL_FUNC_INFO WorkbookBeforeCloseInfo;
	static _ATL_FUNC_INFO WindowActivateInfo;
	static _ATL_FUNC_INFO WindowDeactivateInfo;

public:
	DECLARE_REGISTRY_RESOURCEID(IDR_WORDEVENTS)
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	ExcelEvents() {}
	virtual ~ExcelEvents() {}

private:
	static const int DISPID_WORKBOOK_BEFORE_CLOSE = 0x00000622;
	static const int DISPID_WINDOW_ACTIVATE = 0x00000614;
	static const int DISPID_WINDOW_DEACTIVATE = 0x00000615;

public:
	void SetWorkbookBeforeCloseHandler(WorkbookBeforeCloseHandler handler);
	void SetWorkbookActivateHandler(WorkbookActivateHandler handler);
	void SetWorkbookDeactivateHandler(WorkbookDeactivateHandler handler);

	HRESULT FinalConstruct();
	void FinalRelease() {}

private:
	void STDMETHODCALLTYPE OnWorkbookBeforeClose(Excel::_Workbook* workbook, VARIANT_BOOL* Cancel);
	void STDMETHODCALLTYPE OnWindowActivate(Excel::_Workbook* workbook, Excel::Window* window);
	void STDMETHODCALLTYPE OnWindowDeactivate(Excel::_Workbook* workbook, Excel::Window* window);

BEGIN_COM_MAP(ExcelEvents)
END_COM_MAP()

BEGIN_SINK_MAP(ExcelEvents)
	// Occurs immediately before any open workbook closes.
	//
	SINK_ENTRY_INFO(0, __uuidof(Excel::AppEvents), DISPID_WORKBOOK_BEFORE_CLOSE, OnWorkbookBeforeClose, &WorkbookBeforeCloseInfo)

	// Occurs when any workbook window is activated.
	//
	SINK_ENTRY_INFO(0, __uuidof(Excel::AppEvents), DISPID_WINDOW_ACTIVATE, OnWindowActivate, &WindowActivateInfo)

	// Occurs when any workbook window is deactivated.
	//
	SINK_ENTRY_INFO(0, __uuidof(Excel::AppEvents), DISPID_WINDOW_DEACTIVATE, OnWindowDeactivate, &WindowDeactivateInfo)
END_SINK_MAP()

private:
	WorkbookBeforeCloseHandler m_workbookBeforeCloseHandler;
	WorkbookActivateHandler m_workbookActivateHandler;
	WorkbookDeactivateHandler m_workbookDeactivateHandler;
};

OBJECT_ENTRY_AUTO(__uuidof(ExcelEvents), ExcelEvents)
