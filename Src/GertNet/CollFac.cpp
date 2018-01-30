// CollFac.cpp : Implementation of CCollFac
#include "stdafx.h"
#include "GertNet.h"
#include "CollFac.h"

/////////////////////////////////////////////////////////////////////////////
// CCollFac

STDMETHODIMP CCollFac::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ICollFac,
		&IID_IPersistStorage,
		&IID_IPersistStreamInit,		
		&IID_IClonable,
		&IID_IBSTRKey
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
