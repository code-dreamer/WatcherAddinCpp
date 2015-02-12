// dllmain.h : Declaration of module class.

class WatcherAddinCppAtlModule : public ATL::CAtlDllModuleT< WatcherAddinCppAtlModule >
{
public :
	DECLARE_LIBID(LIBID_WatcherAddinCppAtlLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WATCHERADDINCPPATL, "{30EF3B65-8C21-420B-A439-F4221568C6CD}")
};

extern class WatcherAddinCppAtlModule _AtlModule;
