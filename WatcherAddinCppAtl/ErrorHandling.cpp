#include "stdafx.h"
#include "ErrorHandling.h"
#include "ResUtils.h"

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

namespace ErrorHandling {
;
std::wstring ReadProductName()
{
	LPWSTR  currDllPath = new TCHAR[_MAX_PATH];
	::GetModuleFileName((HINSTANCE)&__ImageBase, currDllPath, _MAX_PATH);

	using namespace ResUtils;
	// replace "040904e4" with the language ID of your resources
	VersionInfo versionInfo = ReadVersionInfo(L"041904b0", currDllPath);
	return versionInfo.productName.empty() ? L"Office Addin" : versionInfo.productName;
}

const std::wstring gProductName = ReadProductName();

void ReportError(const std::exception& exception)
{
	ReportError(exception.what());
}

void ReportError(const char* message)
{
	assert(message && strlen(message) > 0);

	const std::wstring messageStr = ATL::CA2W{message};
	MessageBoxW(NULL, messageStr.c_str(), gProductName.c_str(), MB_OK | MB_ICONERROR);
}

void ReportError(const std::wstring& message)
{
	assert(!message.empty());
	
	MessageBoxW(NULL, message.c_str(), gProductName.c_str(), MB_OK | MB_ICONERROR);
}

} // namespace ErrorHandling
