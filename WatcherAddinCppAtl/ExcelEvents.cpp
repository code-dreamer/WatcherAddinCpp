// ExcelEvents.cpp : Implementation of CExcelEvents

#include "stdafx.h"
#include "ExcelEvents.h"

_ATL_FUNC_INFO ExcelEvents::WorkbookBeforeCloseInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_BOOL}};
_ATL_FUNC_INFO ExcelEvents::WindowActivateInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO ExcelEvents::WindowDeactivateInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_DISPATCH | VT_BYREF}};

HRESULT ExcelEvents::FinalConstruct()
{
	return S_OK;
}

void ExcelEvents::SetWorkbookBeforeCloseHandler(WorkbookBeforeCloseHandler handler)
{
	m_workbookBeforeCloseHandler = handler;
}

void ExcelEvents::SetWorkbookActivateHandler(WorkbookActivateHandler handler)
{
	m_workbookActivateHandler = handler;
}

void ExcelEvents::SetWorkbookDeactivateHandler(WorkbookDeactivateHandler handler)
{
	m_workbookDeactivateHandler = handler;
}

void STDMETHODCALLTYPE ExcelEvents::OnWorkbookBeforeClose(Excel::_Workbook* workbook, VARIANT_BOOL* UNUSED(Cancel))
{
	assert(workbook);

	if (m_workbookBeforeCloseHandler)
		m_workbookBeforeCloseHandler(workbook);
}

void STDMETHODCALLTYPE ExcelEvents::OnWindowActivate(Excel::_Workbook* workbook, Excel::Window* UNUSED(window))
{
	assert(workbook);

	if (m_workbookActivateHandler)
		m_workbookActivateHandler(workbook);
}

void STDMETHODCALLTYPE ExcelEvents::OnWindowDeactivate(Excel::_Workbook* workbook, Excel::Window* UNUSED(window))
{
	assert(workbook);

	if (m_workbookDeactivateHandler)
		m_workbookDeactivateHandler(workbook);
}