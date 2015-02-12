#pragma once

#include "IOfficeAppHandler.h"
#include "PowerPointEvents.h"

class PowerPointHandler : public IOfficeAppHandler
{
	NO_COPY_ASSIGN(PowerPointHandler);

public:
	PowerPointHandler(ATL::CComQIPtr<PowerPoint::_Application> app);
	virtual ~PowerPointHandler();

private:
	void OnPresentationSave(PowerPoint::_Presentation* presentation);
	void OnPresentationClose(PowerPoint::_Presentation* presentation);

	void NotifyAboutPresentationClose(PowerPoint::_Presentation* presentation);

private:
	ATL::CComPtr<PowerPointEvents> m_powerPointEvents;
	ATL::CComQIPtr <PowerPoint::_Application> m_app; 
	std::vector<PowerPoint::DocumentWindows*> m_almostClosedPresentations;
};
