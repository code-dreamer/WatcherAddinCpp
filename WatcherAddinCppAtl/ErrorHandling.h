#pragma once

// Some utility methods to process errors.
//
namespace ErrorHandling {
;
// Show exception error to user.
//
void ReportError(const std::exception& exception);

// Show string error to user.
//
void ReportError(const std::wstring& message);
void ReportError(const char* message);

} // namespace ErrorHandling
