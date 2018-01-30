// CollGosts.cpp : Implementation of CCollGosts
#include "stdafx.h"
#include "AWPlugin.h"
#include "GostTable.h"
#include "CollGosts.h"

/////////////////////////////////////////////////////////////////////////////
// CCollGosts

STDMETHODIMP CCollGosts::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ICollGosts
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
