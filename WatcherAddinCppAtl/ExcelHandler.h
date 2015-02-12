#pragma once

#include "IOfficeAppHandler.h"
#include "ExcelEvents.h"

class ExcelHandler : public IOfficeAppHandler
{
	NO_COPY_ASSIGN(ExcelHandler);

public:
	ExcelHandler(ATL::CComQIPtr<Excel::_Application> app);
	virtual ~ExcelHandler();

private:
	void OnWorkbookBeforeClose(Excel::_Workbook* workbook);
	void OnWorkbookActivate(Excel::_Workbook* workbook);
	void OnWorkbookDeactivate(Excel::_Workbook* workbook);

private:
	ATL::CComPtr<ExcelEvents> m_excelEvents;
	std::vector<Excel::_Workbook*> m_knownWorkbooks;
	ATL::CComQIPtr <Excel::_Application> m_app;
};

