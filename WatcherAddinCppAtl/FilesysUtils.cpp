#include "stdafx.h"
#include "FilesysUtils.h"

namespace FilesysUtils {
;
bool IsFileExist(const std::wstring& pathStr)
{
	assert(!pathStr.empty());

	const std::tr2::sys::wpath path{pathStr.cbegin(), pathStr.cend()};
	return std::tr2::sys::exists(path);
}

} // namespace FilesysUtils
