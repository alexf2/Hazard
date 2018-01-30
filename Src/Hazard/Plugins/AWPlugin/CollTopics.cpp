// CollTopics.cpp : Implementation of CCollTopics
#include "stdafx.h"
#include "AWPlugin.h"
#include "GostTable.h"
#include "CollGosts.h"
#include "CollTopics.h"


/////////////////////////////////////////////////////////////////////////////
// CCollTopics

STDMETHODIMP CCollTopics::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ICollTopics
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
