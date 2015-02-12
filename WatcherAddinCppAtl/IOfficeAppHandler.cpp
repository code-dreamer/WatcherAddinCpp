#include "stdafx.h"
#include "IOfficeAppHandler.h"

void IOfficeAppHandler::SetDocCloseHandler(DocCloseHandler docCloseHandler)
{
	m_closedHandler = docCloseHandler;
}

void IOfficeAppHandler::FireCloseHandler(const BSTR& path)
{
	if (m_closedHandler)
		m_closedHandler(path);
}
