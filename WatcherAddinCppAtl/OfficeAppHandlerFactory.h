#pragma once

class IOfficeAppHandler;

// Simple factory to encapsulate handler construction, 
// which depends on office application type
//
namespace OfficeAppHandlerFactory {
;
//	Creates specific implementation behind IDocumentHandler
// that can raise event after document close.
// officeAapplication:
// Office host application object.
//
IOfficeAppHandler* CreateHandler(LPDISPATCH officeApp);
};

