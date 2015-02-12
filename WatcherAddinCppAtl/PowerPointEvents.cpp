#include "stdafx.h"
#include "PowerPointEvents.h"

_ATL_FUNC_INFO PowerPointEvents::PresentationCloseInfo = {CC_STDCALL, VT_EMPTY, 1, {VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO PowerPointEvents::PresentationSaveInfo = {CC_STDCALL, VT_EMPTY, 1, {VT_DISPATCH | VT_BYREF}};

HRESULT PowerPointEvents::FinalConstruct()
{
	return S_OK;
}

void STDMETHODCALLTYPE PowerPointEvents::OnPresentationSave(PowerPoint::_Presentation* presentation)
{
	assert(presentation);

	if (m_presentationSaveHandler)
		m_presentationSaveHandler(presentation);
}

void STDMETHODCALLTYPE PowerPointEvents::OnPresentationClose(PowerPoint::_Presentation* presentation)
{
	assert(presentation);

	if (m_presentationCloseHandler)
		m_presentationCloseHandler(presentation);
}

void PowerPointEvents::SetPresentationSaveHandler(PresentationSaveHandler handler)
{
	m_presentationSaveHandler = handler;
}

void PowerPointEvents::SetPresentationCloseHandler(PresentationCloseHandler handler)
{
	m_presentationCloseHandler = handler;
}
