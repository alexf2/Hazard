// ICollFChange.cpp : Implementation of CICollFChange
#include "stdafx.h"
#include "GertNet.h"
#include "ICollFChange.h"

/////////////////////////////////////////////////////////////////////////////
// CICollFChange

STDMETHODIMP CICollFChange::InterfaceSupportsErrorInfo(REFIID riid)
{	
	   static const IID* arr[] = 
	   {
		   &IID_IICollFChange,		   
		   &IID_IPersistStreamInit,		
		   &IID_IClonable,
		   &IID_ILongKey
	   };
	   for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	   {
		   if (InlineIsEqualGUID(*arr[i],riid))
			   return S_OK;
	   }
	   return S_FALSE;
}
