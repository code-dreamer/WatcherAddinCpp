// Connect.h : Declaration of the Connect

#pragma once
#include "resource.h"       // main symbols
#include "WatcherAddinCppAtl_i.h"

#include "ErrorHandling.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

typedef ATL::IDispatchImpl < AddinDesign::_IDTExtensibility2, 
	&__uuidof(AddinDesign::_IDTExtensibility2), &__uuidof(AddinDesign::__AddInDesignerObjects), /* wMajor = */ 1 >
	IDTImpl;

// Main manager class for this Addin
//
class ATL_NO_VTABLE Connect :
	public ATL::CComObjectRootEx<ATL::CComSingleThreadModel>,
	public ATL::CComCoClass<Connect, &CLSID_Connect>,
	public ATL::IDispatchImpl<IConnect, &IID_IConnect, &LIBID_WatcherAddinCppAtlLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDTImpl
{
	NO_COPY_ASSIGN(Connect);

public:
	Connect() {}
	virtual ~Connect() {}

	HRESULT FinalConstruct();
	void FinalRelease() {};

DECLARE_REGISTRY_RESOURCEID(IDR_CONNECT)
DECLARE_PROTECT_FINAL_CONSTRUCT()

// _IDTExtensibility2 Methods
//
public:
	// Implements the OnConnection method of the IDTExtensibility2 interface.
	// Receives notification that the Add-in is being loaded.

	// officeAap
	// Root object of the host application.

	// connectMode
	// Describes how the Add-in is being loaded.
	
	// addInInst
	//	Object representing this Add-in.
	
	// custom
	// Array of parameters that are host application specific.
	//
	STDMETHOD(OnConnection)(LPDISPATCH officeAap, AddinDesign::ext_ConnectMode connectMode, LPDISPATCH addInInst, SAFEARRAY** custom);

	// Implements the OnDisconnection method of the IDTExtensibility2 interface.
	// Receives notification that the Add-in is being unloaded.
	
	// removeMode
	// Describes how the Add-in is being unloaded.
	
	// custom
	// Array of parameters that are host application specific.
	//
	STDMETHOD(OnDisconnection)(AddinDesign::ext_DisconnectMode RemoveMode, SAFEARRAY** custom);

	// Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
	// Receives notification that the collection of Add-ins has changed.
	
	// custom
	// Array of parameters that are host application specific.
	//
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY** custom);

	/// <summary>
	// Implements the OnStartupComplete method of the IDTExtensibility2 interface.
	// Receives notification that the host application has completed loading.
	
	// custom
	// Array of parameters that are host application specific.
	//
	STDMETHOD(OnStartupComplete)(SAFEARRAY** custom);

	// Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
	// Receives notification that the host application is being unloaded.
	
	// custom
	// Array of parameters that are host application specific.
	//
	STDMETHOD(OnBeginShutdown)(SAFEARRAY** custom);
//////////////////////////////////////////////////////////////////////////

private:
	BEGIN_COM_MAP(Connect)
		COM_INTERFACE_ENTRY(IConnect)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()

private:
	void OnOfficeDocumentClosed(const BSTR& path);

private:
	HANDLE m_outfile{};
	class IOfficeAppHandler* m_officeAppHandler{};
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), Connect)
