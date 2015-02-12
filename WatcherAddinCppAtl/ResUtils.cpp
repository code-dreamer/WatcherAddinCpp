#include "stdafx.h"
#include "ResUtils.h"
#include "FilesysUtils.h"

namespace ResUtils
{
;
VersionInfo ReadVersionInfo(const std::wstring& languageId, const std::wstring& modulePath)
{
	assert(!languageId.empty());
	assert(!modulePath.empty());
	assert(FilesysUtils::IsFileExist(modulePath));

	VersionInfo result;

	// allocate a block of memory for the version info
	DWORD dummy;
	DWORD verSize = GetFileVersionInfoSize(modulePath.c_str(), &dummy);
	if (verSize == 0)
		return result;
	
	std::vector<BYTE> data(verSize);

	// load the version info
	BOOL ok = GetFileVersionInfo(modulePath.c_str(), NULL, verSize, &data[0]);
	if (ok == FALSE)
		return result;
	
	auto ReadVersionValue = [&, languageId](const std::wstring& valueName) -> std::wstring {
		assert(!valueName.empty());
		std::wstring value;
		std::wstring query = L"\\StringFileInfo\\" + languageId + L"\\" + valueName;
		LPVOID rawValue{};
		unsigned int rawValueLen = 0;
		if (VerQueryValue(&data[0], query.c_str(), &rawValue, &rawValueLen) == TRUE)
			value.assign((wchar_t*)rawValue, rawValueLen - 1);

		return value;
	};

	result.productName = ReadVersionValue(L"ProductName");
	result.internalName = ReadVersionValue(L"InternalName");
	result.companyName = ReadVersionValue(L"CompanyName");
	std::wstring productVersion = ReadVersionValue(L"ProductVersion");
	size_t ind = productVersion.find(L".");
	assert(ind != std::string::npos);
	result.productVersionString = productVersion.substr(0, ind);
	result.fullProductName = result.productName + L" " + result.productVersionString;

	return result;
}

} // namespace ResUtils
