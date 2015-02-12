#include "stdafx.h"
#include "ExcelHandler.h"

using namespace std::placeholders;

ExcelHandler::ExcelHandler(CComQIPtr<Excel::_Application> app)
	: m_app{app}
{
	assert(m_app);

	CComObject<ExcelEvents>* excelEventsObj;
	HRESULT hr = CComObject<ExcelEvents>::CreateInstance(&excelEventsObj);
	assert(SUCCEEDED(hr));

	m_excelEvents = excelEventsObj;
	m_excelEvents->SetWorkbookBeforeCloseHandler(std::bind(&ExcelHandler::OnWorkbookBeforeClose, this, _1));
	m_excelEvents->SetWorkbookActivateHandler(std::bind(&ExcelHandler::OnWorkbookActivate, this, _1));
	m_excelEvents->SetWorkbookDeactivateHandler(std::bind(&ExcelHandler::OnWorkbookDeactivate, this, _1));
	hr = m_excelEvents->DispEventAdvise(m_app);
	assert(SUCCEEDED(hr));
}

ExcelHandler::~ExcelHandler()
{
	HRESULT hr = m_excelEvents->DispEventUnadvise(m_app);
	assert(SUCCEEDED(hr));
}

void ExcelHandler::OnWorkbookBeforeClose(Excel::_Workbook* workbook)
{
	assert(workbook);

	auto it = std::find(m_knownWorkbooks.cbegin(), m_knownWorkbooks.cend(), workbook);
	if (it == m_knownWorkbooks.cend())
		m_knownWorkbooks.push_back(workbook); // the document doesn't exist in collection
}

void ExcelHandler::OnWorkbookActivate(Excel::_Workbook* workbook)
{
	assert(workbook);

	auto it = std::find(m_knownWorkbooks.cbegin(), m_knownWorkbooks.cend(), workbook);
	if (it != m_knownWorkbooks.cend())
		m_knownWorkbooks.erase(it); // user canceled the close dlg
}

void ExcelHandler::OnWorkbookDeactivate(Excel::_Workbook* workbook)
{
	assert(workbook);

	auto it = std::find(m_knownWorkbooks.cbegin(), m_knownWorkbooks.cend(), workbook);
	if (it != m_knownWorkbooks.cend()) {
		// Check if the document is saved (ignore unsaved documents without path)
		BSTR saveDir;
		HRESULT hr = workbook->get_Path(0, &saveDir);
		assert(SUCCEEDED(hr));
		if (SysStringLen(saveDir) != 0) {
			// Getting the full path of the document
			BSTR fullPath;
			hr = workbook->get_FullName(0, &fullPath);
			assert(SUCCEEDED(hr));

			FireCloseHandler(fullPath);
		}

		m_knownWorkbooks.erase(it);
	}
}
