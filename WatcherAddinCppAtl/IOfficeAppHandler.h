#pragma once

using DocCloseHandler = std::function<void(const BSTR&)>;

class IOfficeAppHandler
{
	NO_COPY_ASSIGN(IOfficeAppHandler);

public:
	IOfficeAppHandler() {};
	virtual ~IOfficeAppHandler() = 0 {};

	void SetDocCloseHandler(DocCloseHandler docCloseHandler);

protected:
	void FireCloseHandler(const BSTR& path);

private:
	DocCloseHandler m_closedHandler{};
};
