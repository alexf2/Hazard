// OtlGIT.h
// Copyright (C) 1999 Stingray Software Inc.
// All Rights Reserved

#ifndef __OTLGIT_H__
#define __OTLGIT_H__

#ifndef __OTLBASE_H__
	#error otlgit.h requires otlbase.h to be included first
#endif

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#include <vector>

namespace StingrayOTL
{

/////////////////////////////////////////////////////////////////////////////
// class OtlGIT

// This is a general helper class to assist developers needing to marshal
//  pointers from one apartment to another. 

class OtlGIT 
{
	static bool m_sbIsInstantiated;

	IGlobalInterfaceTable* m_pGIT;

	// list of allocated marshaled object references stored in streams...
	std::vector<LPSTREAM> m_streams;

public:
	OtlGIT() 
	{
		// check to see if already registered...
		if(m_sbIsInstantiated)
			return;

		// constructor should get a pointer to the global interface table
		m_pGIT = NULL;
		HRESULT hr;
		hr = CoCreateInstance(CLSID_StdGlobalInterfaceTable,                            
							NULL,
                            CLSCTX_INPROC_SERVER,
                            IID_IGlobalInterfaceTable,
                            (void **)&m_pGIT);

		if(SUCCEEDED(hr)) 
		{

		}
		m_sbIsInstantiated = TRUE;

	}

	~OtlGIT() 
	{
		std::vector<LPSTREAM>::iterator iter;
		for (iter = m_streams.begin(); iter < m_streams.end(); iter++) 
		{
			(*iter)->Release();
			m_streams.erase(iter);
		}

		// Destructor should release the pointer to the global interface table
		if(m_pGIT) 
		{
			m_pGIT->Release();
		} 

		m_sbIsInstantiated = FALSE;
	}

	HRESULT RegisterInterface(IUnknown *pUnk,  
							REFIID  riid,   
							DWORD *pdwCookie) {
		if(m_pGIT)
		{
			return m_pGIT->RegisterInterfaceInGlobal(pUnk, riid, pdwCookie); 
		} else // manage our own table...
		{
			LPSTREAM pStm = 0;

			HRESULT hr;
			hr = CoMarshalInterThreadInterfaceInStream(riid,
														pUnk,
														&pStm);
			if(SUCCEEDED(hr)) {
				m_streams.push_back(pStm);
				*pdwCookie = (DWORD)pStm;
			}
			return hr;
		}
	}
	HRESULT RevokeInterface(DWORD  dwCookie)
	{
		if(m_pGIT)
		{
			return m_pGIT->RevokeInterfaceFromGlobal(dwCookie);
		} else
		{
			std::vector<LPSTREAM>::iterator iter;
			for (iter = m_streams.begin(); iter < m_streams.end(); iter++) 
			{
				if ((DWORD)*iter == (DWORD)dwCookie) 
				{
					(*iter)->Release();
					m_streams.erase(iter);
					break;
				}
			}

			return S_OK;
		}
	}
	HRESULT GetInterface(DWORD dwCookie, REFIID riid, void** ppv)
	{
		if(m_pGIT) 
		{
			return m_pGIT->GetInterfaceFromGlobal(dwCookie, riid, ppv);
		} else
		{
			LPSTREAM pStm = NULL;

			std::vector<LPSTREAM>::iterator iter;
			for (iter = m_streams.begin(); iter < m_streams.end(); iter++) 
			{
				if ((DWORD)*iter == (DWORD)dwCookie) 
				{
					pStm = (*iter);
					*iter = 0;
					break;
				}
			}

			HRESULT hr;
			hr = S_OK;

			if(pStm)
			{
				hr = CoGetInterfaceAndReleaseStream(pStm,
													riid,
													ppv);
			} else
			{
				hr = STG_E_INVALIDPOINTER;
			}

			return hr;
		}
	}

};

//bool OtlGIT::m_sbIsInstantiated = FALSE;

};	// namespace StingrayOTL

#endif // __OTLGIT_H__