#include "stdafx.h"
#include "WordEvents.h"

_ATL_FUNC_INFO WordEvents::DocumentBeforeCloseInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_BOOL}};
_ATL_FUNC_INFO WordEvents::WindowActivateInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO WordEvents::WindowDeactivateInfo = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH | VT_BYREF, VT_DISPATCH | VT_BYREF}};

HRESULT WordEvents::FinalConstruct()
{
	return S_OK;
}

void WordEvents::SetDocumentBeforeCloseHandler(DocumentBeforeCloseHandler handler)
{
	m_documentBeforeCloseHandler = handler;
}

void WordEvents::SetDocumentActivateHandler(DocumentActivateHandler handler)
{
	m_documentActivateHandler = handler;
}

void WordEvents::SetDocumentDeactivateHandler(DocumentDeactivateHandler handler)
{
	m_documentDeactivateHandler = handler;
}

void STDMETHODCALLTYPE WordEvents:: OnDocumentBeforeClose(Word::_Document* doc, VARIANT_BOOL* UNUSED(Cancel))
{
	assert(doc);
	
	if (m_documentBeforeCloseHandler)
		m_documentBeforeCloseHandler(doc);
}

void STDMETHODCALLTYPE WordEvents::OnWindowActivate(Word::_Document* doc, Word::Window* UNUSED(window))
{
	assert(doc);

	if (m_documentActivateHandler)
		m_documentActivateHandler(doc);
}

void STDMETHODCALLTYPE WordEvents::OnWindowDeactivate(Word::_Document* doc, Word::Window* UNUSED(window))
{
	assert(doc);

	if (m_documentDeactivateHandler)
		m_documentDeactivateHandler(doc);
}
