#include "stdafx.h"
#include "OfficeAppHandlerFactory.h"
#include "WordHandler.h"
#include "PowerPointHandler.h"
#include "ExcelHandler.h"

using namespace ATL;

namespace OfficeAppHandlerFactory {
;
IOfficeAppHandler* CreateHandler(LPDISPATCH officeApp)
{
	assert(officeApp);

	CComQIPtr<Word::_Application> wordApp{officeApp};
	if (wordApp)
		return new WordHandler{wordApp};

	CComQIPtr<PowerPoint::_Application> powerPointApp{officeApp};
	if (powerPointApp)
		return new PowerPointHandler{powerPointApp};

	CComQIPtr<Excel::_Application> excelApp{officeApp};
	if (excelApp)
		return new ExcelHandler{excelApp};

	assert(false);
	return nullptr;
}

} // DocumentHandlerFactory
