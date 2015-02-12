#include "stdafx.h"
#include "WordHandler.h"

using namespace std::placeholders;

WordHandler::WordHandler(CComQIPtr<Word::_Application> app)
	: m_app{app}
{
	assert(m_app);

	CComObject<WordEvents>* wordEventsObj;
	HRESULT hr = CComObject<WordEvents>::CreateInstance(&wordEventsObj);
	assert(SUCCEEDED(hr));
	
	m_wordEvents = wordEventsObj;
	m_wordEvents->SetDocumentBeforeCloseHandler(std::bind(&WordHandler::OnDocumentBeforeClose, this, _1));
	m_wordEvents->SetDocumentActivateHandler(std::bind(&WordHandler::OnDocumentActivate, this, _1));
	m_wordEvents->SetDocumentDeactivateHandler(std::bind(&WordHandler::OnDocumentDeactivate, this, _1));
	hr = m_wordEvents->DispEventAdvise(m_app);
	assert(SUCCEEDED(hr));
}

WordHandler::~WordHandler()
{
	HRESULT hr = m_wordEvents->DispEventUnadvise(m_app);
	assert(SUCCEEDED(hr));
}

void WordHandler::OnDocumentBeforeClose(Word::_Document* doc)
{
	assert(doc);

	auto it = std::find(m_knownDocs.cbegin(), m_knownDocs.cend(), doc);
	if (it == m_knownDocs.cend())
		m_knownDocs.push_back(doc); // the document doesn't exist in collection
}

void WordHandler::OnDocumentActivate(Word::_Document* doc)
{
	assert(doc);

	auto it = std::find(m_knownDocs.cbegin(), m_knownDocs.cend(), doc);
	if (it != m_knownDocs.cend())
		m_knownDocs.erase(it); // user canceled the close dlg
}

void WordHandler::OnDocumentDeactivate(Word::_Document* doc)
{
	assert(doc);

	auto it = std::find(m_knownDocs.cbegin(), m_knownDocs.cend(), doc);
	if (it != m_knownDocs.cend()) {
		// Check if the document is saved (ignore unsaved documents without path)
		BSTR saveDir;
		HRESULT hr = doc->get_Path(&saveDir);
		assert(SUCCEEDED(hr));
		if (SysStringLen(saveDir) != 0) {
			// Getting the full path of the document
			BSTR fullPath;
			hr = doc->get_FullName(&fullPath);
			assert(SUCCEEDED(hr));
			FireCloseHandler(fullPath);
		}
			
		m_knownDocs.erase(it);
	}
}
