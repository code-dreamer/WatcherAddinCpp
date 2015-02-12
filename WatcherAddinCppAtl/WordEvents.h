// WordEvents.h : Declaration of the WordEvents

#pragma once
#include "resource.h"       // main symbols
#include "WatcherAddinCppAtl_i.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

using DocumentBeforeCloseHandler = std::function<void(Word::_Document*)>;
using DocumentActivateHandler = std::function<void(Word::_Document*)>;
using DocumentDeactivateHandler = std::function<void(Word::_Document*)>;

class ATL_NO_VTABLE WordEvents :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<WordEvents, &CLSID_WordEvents>,
	public IDispEventSimpleImpl<0, WordEvents, &__uuidof(Word::ApplicationEvents2) >
{
	NO_COPY_ASSIGN(WordEvents);

	static _ATL_FUNC_INFO DocumentBeforeCloseInfo;
	static _ATL_FUNC_INFO WindowActivateInfo;
	static _ATL_FUNC_INFO WindowDeactivateInfo;

public:
	DECLARE_REGISTRY_RESOURCEID(IDR_WORDEVENTS)
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	WordEvents() {}
	virtual ~WordEvents() {}

private:
	static const int DISPID_DOCUMENT_BEFORE_CLOSE = 0x00000006;
	static const int DISPID_WINDOW_ACTIVATE = 0x0000000a;
	static const int DISPID_WINDOW_DEACTIVATE = 0x0000000b;

public:
	void SetDocumentBeforeCloseHandler(DocumentBeforeCloseHandler handler);
	void SetDocumentActivateHandler(DocumentActivateHandler handler);
	void SetDocumentDeactivateHandler(DocumentDeactivateHandler handler);

	HRESULT FinalConstruct();
	void FinalRelease() {}

private:
	void STDMETHODCALLTYPE OnDocumentBeforeClose(Word::_Document* doc, VARIANT_BOOL* Cancel);
	void STDMETHODCALLTYPE OnWindowActivate(Word::_Document* doc, Word::Window* window);
	void STDMETHODCALLTYPE OnWindowDeactivate(Word::_Document* doc, Word::Window* window);

BEGIN_COM_MAP(WordEvents)
END_COM_MAP()

BEGIN_SINK_MAP(WordEvents)
	// Occurs immediately before any open document closes.
	//
	SINK_ENTRY_INFO(0, __uuidof(Word::ApplicationEvents2), DISPID_DOCUMENT_BEFORE_CLOSE, OnDocumentBeforeClose, &DocumentBeforeCloseInfo)

	// Occurs when any document window is activated.
	//
	SINK_ENTRY_INFO(0, __uuidof(Word::ApplicationEvents2), DISPID_WINDOW_ACTIVATE, OnWindowActivate, &WindowActivateInfo)

	// Occurs when any document window is deactivated.
	//
	SINK_ENTRY_INFO(0, __uuidof(Word::ApplicationEvents2), DISPID_WINDOW_DEACTIVATE, OnWindowDeactivate, &WindowDeactivateInfo)
END_SINK_MAP()

private:	
	DocumentBeforeCloseHandler m_documentBeforeCloseHandler;
	DocumentActivateHandler m_documentActivateHandler;
	DocumentDeactivateHandler m_documentDeactivateHandler;
};

OBJECT_ENTRY_AUTO(__uuidof(WordEvents), WordEvents)
