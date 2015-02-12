#pragma once

namespace ResUtils {
;
struct VersionInfo
{
	std::wstring fullProductName;
	std::wstring productName;
	std::wstring internalName;
	std::wstring companyName;
	std::wstring productVersionString;
};
VersionInfo ReadVersionInfo(const std::wstring& languageId, const std::wstring& modulePath);

} // namespace ResUtils
