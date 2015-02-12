#include "stdafx.h"
#include "EnvUtils.h"

namespace EnvUtils {
;
std::wstring GetTempDir()
{
	wchar_t buf[MAX_PATH];

	if (GetTempPath(MAX_PATH, buf) == 0)
		throw std::exception("TEMP dir is undefined");

	return buf;
}

} // namespace EnvUtils
