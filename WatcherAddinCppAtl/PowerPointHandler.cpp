#include "stdafx.h"
#include "PowerPointHandler.h"
#include "Timer.h"

using namespace ATL;
using namespace std::placeholders;

PowerPointHandler::PowerPointHandler(CComQIPtr<PowerPoint::_Application> app)
	: m_app{app}
{
	assert(m_app);

	CComObject<PowerPointEvents>* powerPointEventsObj;
	HRESULT hr = CComObject<PowerPointEvents>::CreateInstance(&powerPointEventsObj);
	assert(SUCCEEDED(hr));

	m_powerPointEvents = powerPointEventsObj;
	
	m_powerPointEvents->SetPresentationCloseHandler(std::bind(&PowerPointHandler::OnPresentationClose, this, _1));
	m_powerPointEvents->SetPresentationSaveHandler(std::bind(&PowerPointHandler::OnPresentationSave, this, _1));
	hr = m_powerPointEvents->DispEventAdvise(m_app);
	assert(SUCCEEDED(hr));
}

PowerPointHandler::~PowerPointHandler()
{
	HRESULT hr = m_powerPointEvents->DispEventUnadvise(m_app);
	assert(SUCCEEDED(hr));
}

void PowerPointHandler::OnPresentationSave(PowerPoint::_Presentation* presentation)
{
	PowerPoint::DocumentWindows* documentWindows;
	HRESULT hr = presentation->get_Windows(&documentWindows);
	assert(SUCCEEDED(hr));
	assert(documentWindows);
	
	auto it = std::find(m_almostClosedPresentations.cbegin(), m_almostClosedPresentations.cend(), documentWindows);
	if (it != m_almostClosedPresentations.cend()) {
		NotifyAboutPresentationClose(presentation);
		m_almostClosedPresentations.erase(it);
	}
}

void PowerPointHandler::OnPresentationClose(PowerPoint::_Presentation* presentation)
{
	assert(presentation);

	Office::MsoTriState state;
	HRESULT hr = presentation->get_Saved(&state);
	assert(SUCCEEDED(hr));

	if (state == Office::MsoTriState::msoTrue) {
		NotifyAboutPresentationClose(presentation);
	}
	else {
		PowerPoint::DocumentWindows* documentWindows;
		HRESULT hr = presentation->get_Windows(&documentWindows);
		assert(SUCCEEDED(hr));
		assert(documentWindows);

		auto it = std::find(m_almostClosedPresentations.cbegin(), m_almostClosedPresentations.cend(), documentWindows);
		if (it == m_almostClosedPresentations.cend())
			m_almostClosedPresentations.push_back(documentWindows);
	}
}

void PowerPointHandler::NotifyAboutPresentationClose(PowerPoint::_Presentation* presentation)
{
	assert(presentation);

	BSTR saveDir;
	presentation->get_Path(&saveDir);
	if (SysStringLen(saveDir) != 0) {
		BSTR fullPath;
		presentation->get_FullName(&fullPath);
		FireCloseHandler(fullPath);
	}
}
