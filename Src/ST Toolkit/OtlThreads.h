// OtlThreads.h
// Copyright (C) 1999 Stingray Software Inc.
// All Rights Reserved

#ifndef __OTLTHREADS_H__
#define __OTLTHREADS_H__

#ifndef __OTLBASE_H__
	#error otlthreads.h requires otlbase.h to be included first
#endif

#include <otlfunctor.h>

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

namespace StingrayOTL
{

/////////////////////////////////////////////////////////////////////////////
// COtlThread
// encapsulates a thread procedure as a functor

//void otlApartmentMain ( COtlFunctor fmain, DWORD dwCoInit);
inline void otlApartmentMain ( COtlFunctor fmain, DWORD dwCoInit)
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

//COtlFunctor otlMakeAptMainFunctor ( COtlFunctor fmain, DWORD dwCoInit);
inline COtlFunctor otlMakeAptMainFunctor ( COtlFunctor fmain, DWORD dwCoInit)
{
	// make a functor that calls otlApartmentMain( my_worker_functor, dwCoInit )
    return otlMakeFunctor(otlApartmentMain, fmain, dwCoInit);
}


class COtlThreadFunction
{
public:
	COtlThreadFunction(COtlFunctor fmain, DWORD dwCoInit = COINIT_APARTMENTTHREADED)
		: m_hthread(0), m_dwTID(0)
	{
		m_functor = otlMakeAptMainFunctor(fmain, dwCoInit);
		m_hthread = ::CreateThread(NULL, 0, ThreadEntry, &m_functor, CREATE_SUSPENDED, &m_dwTID);
	}

	virtual ~COtlThreadFunction()
	{
		WaitForFinish();
	}

	virtual void WaitForFinish(DWORD dwTimeout = INFINITE)
	{
		if(m_hthread)
		 {    
//begin AF: !!!
                    if( ::WaitForSingleObject(m_hthread, 0) == WAIT_TIMEOUT )
                      {
                        DWORD dwCnt = SuspendThread( m_hthread );
     	                ResumeThread( m_hthread );
                        if( dwCnt == 0 )
                          ::WaitForSingleObject(m_hthread, dwTimeout);
                      }              		    
//end AF: !!!
		    ::CloseHandle(m_hthread);
		    m_hthread = 0;
		 }
	}

	virtual BOOL Start()
	{
		OTLASSERT(m_hthread);
		if(0xFFFFFFFF == ::ResumeThread(m_hthread))
			return FALSE;
		return TRUE;

	}
	
	static DWORD WINAPI ThreadEntry(void *pv)
	{
		COtlFunctor* pFunctor = (COtlFunctor*)pv;
		(*pFunctor)();
		return 0;
	}


	COtlFunctor m_functor;
	HANDLE m_hthread;
	DWORD m_dwTID;

};


};	// namespace StingrayOTL

#endif // __OTLTHREADS_H__
