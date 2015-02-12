#include "stdafx.h"
#include "Connect.h"
#include "ErrorHandling.h"
#include "EnvUtils.h"
#include "OfficeAppHandlerFactory.h"
#include "IOfficeAppHandler.h"
#include "FilesysUtils.h"

using namespace std::placeholders;
using namespace AddinDesign;

const std::wstring gOutFilename{L"ClosedDocuments.txt"};

HRESULT Connect::OnConnection(LPDISPATCH officeAap, ext_ConnectMode UNUSED(ConnectMode), LPDISPATCH UNUSED(addInInst), SAFEARRAY** UNUSED(custom))
{
	try {
		// If host app is null, we can't do anything
		if (!officeAap)
			throw new std::exception("officeAapplication is null");

		std::tr2::sys::wpath outFilepath{EnvUtils::GetTempDir()};
		outFilepath.append(gOutFilename.cbegin(), gOutFilename.cend());

		DWORD CreateFlag = FilesysUtils::IsFileExist(outFilepath.string()) ? OPEN_EXISTING : CREATE_ALWAYS;
		m_outfile = CreateFile(outFilepath.string().c_str(),
								   FILE_APPEND_DATA,
								   FILE_SHARE_READ|FILE_SHARE_WRITE,
								   NULL,
								   CreateFlag,
								   FILE_ATTRIBUTE_NORMAL,
								   NULL);

		if (m_outfile ==  INVALID_HANDLE_VALUE)
			throw new std::exception("Can't open log file");

		// Create handler for specific office app host
		m_officeAppHandler = OfficeAppHandlerFactory::CreateHandler(officeAap);
		// Connect event
		m_officeAppHandler->SetDocCloseHandler(std::bind(&Connect::OnOfficeDocumentClosed, this, _1));
	}
	catch (const std::exception& e) {
		ErrorHandling::ReportError(e);
	}
	
	return S_OK;
}

HRESULT Connect::OnDisconnection(ext_DisconnectMode UNUSED(RemoveMode), SAFEARRAY** UNUSED(custom))
{
	CloseHandle(m_outfile);

	return S_OK;
}

HRESULT Connect::OnAddInsUpdate(SAFEARRAY** UNUSED(custom))
{
	return S_OK;
}

HRESULT Connect::OnStartupComplete(SAFEARRAY** UNUSED(custom))
{
	return S_OK;
}

HRESULT Connect::OnBeginShutdown(SAFEARRAY** UNUSED(custom))
{
	return S_OK;
}

void Connect::OnOfficeDocumentClosed(const BSTR& path)
{
	DWORD dwBytes;
	BOOL ok = WriteFile(m_outfile,
			  path,
			  SysStringLen(path) * 2,
			  &dwBytes,
			  NULL);
	if (ok == FALSE) {
		ErrorHandling::ReportError("Log write failed");
		return;
	}

	ok = WriteFile(m_outfile,
			L"\r\n",
			4,
			&dwBytes,
			NULL);
	if (ok == FALSE) {
		ErrorHandling::ReportError("Log write failed");
		return;
	}

	FlushFileBuffers(m_outfile);
}

HRESULT Connect::FinalConstruct()
{
	return S_OK;
}
