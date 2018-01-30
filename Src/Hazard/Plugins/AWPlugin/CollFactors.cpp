// CollFactors.cpp : Implementation of CCollFactors
#include "stdafx.h"
#include "AWPlugin.h"
#include "FactorTable.h"
#include "CollFTables.h"
#include "CollFactors.h"

/////////////////////////////////////////////////////////////////////////////
// CCollFactors

STDMETHODIMP CCollFactors::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ICollFactors
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
