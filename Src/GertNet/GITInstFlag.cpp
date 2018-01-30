#include "stdafx.h"

//moved from OtlGIT.h
bool OtlGIT::m_sbIsInstantiated = FALSE;

/*//moved from OtlThreads.h
void otlApartmentMain ( COtlFunctor fmain, DWORD dwCoInit)
{
    if (SUCCEEDED(::CoInitializeEx(0, dwCoInit)))
        try
        {
            fmain();
            ::CoUninitialize();
        }
        catch ( ... )
        {
            ::CoUninitialize();
            throw;
        }
}

COtlFunctor otlMakeAptMainFunctor ( COtlFunctor fmain, DWORD dwCoInit)
{
	// make a functor that calls otlApartmentMain( my_worker_functor, dwCoInit )
    return otlMakeFunctor(otlApartmentMain, fmain, dwCoInit);
}
*/