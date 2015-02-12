#pragma once

#include "IOfficeAppHandler.h"
#include "WordEvents.h"

class WordHandler : public IOfficeAppHandler
{
	NO_COPY_ASSIGN(WordHandler);

public:
	WordHandler(ATL::CComQIPtr<Word::_Application> app);
	virtual ~WordHandler();

	void HandleDisconnection() {};

private:
	void OnDocumentBeforeClose(Word::_Document* doc);
	void OnDocumentActivate(Word::_Document* doc);
	void OnDocumentDeactivate(Word::_Document* doc);

private:
	ATL::CComPtr<WordEvents> m_wordEvents;
	std::vector<Word::_Document*> m_knownDocs;
	ATL::CComQIPtr <Word::_Application> m_app; 
};
