// PowerPointEvents.h : Declaration of the CPowerPointEvents

#pragma once
#include "resource.h"       // main symbols
#include "WatcherAddinCppAtl_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

using PresentationCloseHandler = std::function<void(PowerPoint::_Presentation*)>;
using PresentationSaveHandler = std::function<void(PowerPoint::_Presentation*)>;

class ATL_NO_VTABLE PowerPointEvents :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<PowerPointEvents, &CLSID_PowerPointEvents>,
	public IDispatchImpl<IPowerPointEvents, &IID_IPowerPointEvents, &LIBID_WatcherAddinCppAtlLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispEventSimpleImpl<0, PowerPointEvents, &__uuidof(PowerPoint::EApplication) >
{
	NO_COPY_ASSIGN(PowerPointEvents);

	static _ATL_FUNC_INFO PresentationCloseInfo;
	static _ATL_FUNC_INFO PresentationSaveInfo;

public:
	DECLARE_REGISTRY_RESOURCEID(IDR_POWERPOINTEVENTS)
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	PowerPointEvents() {}
	virtual ~PowerPointEvents() {}

	HRESULT FinalConstruct();
	void FinalRelease()	{}

	void SetPresentationCloseHandler(PresentationCloseHandler handler);
	void SetPresentationSaveHandler(PresentationSaveHandler handler);
	
private:
	static const int DISPID_PRESENTATION_CLOSE = 2004;
	static const int DISPID_PRESENTATION_SAVE = 2005;

BEGIN_COM_MAP(PowerPointEvents)
	COM_INTERFACE_ENTRY(IPowerPointEvents)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

BEGIN_SINK_MAP(PowerPointEvents)
	// Occurs before any open presentation is saved (after save dialog).
	//
	SINK_ENTRY_INFO(0, __uuidof(PowerPoint::EApplication), DISPID_PRESENTATION_SAVE, OnPresentationSave, &PresentationSaveInfo)

	// Occurs immediately before any open presentation closes, as it is removed from the Presentations collection.
	//
	SINK_ENTRY_INFO(0, __uuidof(PowerPoint::EApplication), DISPID_PRESENTATION_CLOSE, OnPresentationClose, &PresentationCloseInfo)
END_SINK_MAP()
	
private:
	void STDMETHODCALLTYPE OnPresentationSave(PowerPoint::_Presentation* presentation);
	void STDMETHODCALLTYPE OnPresentationClose(PowerPoint::_Presentation* presentation);

private:
	PresentationCloseHandler m_presentationCloseHandler;
	PresentationSaveHandler m_presentationSaveHandler;
};

OBJECT_ENTRY_AUTO(__uuidof(PowerPointEvents), PowerPointEvents)
